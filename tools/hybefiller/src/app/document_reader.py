"""Local PDF/OCR reader. Page positions and confidence remain available for audit."""
from __future__ import annotations

import hashlib
import json
import os
import re
import tempfile
from contextlib import contextmanager
from pathlib import Path

CACHE_VERSION = 'hybe-1.2.0-ocr-2200-v2'
_ENGINE = None

class Cancelled(Exception):
    pass

def check_cancel(cancel):
    if cancel is not None and cancel.is_set():
        raise Cancelled('사용자가 중지했습니다.')

def sha256_file(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest()

@contextmanager
def pdf_snapshot(path,cancel=None):
    """Hash and parse the same immutable bytes, even if the source is edited."""
    descriptor,temporary=tempfile.mkstemp(prefix='hybe-source-',suffix='.pdf')
    digest=hashlib.sha256()
    try:
        with os.fdopen(descriptor,'wb') as target,Path(path).open('rb') as source:
            for block in iter(lambda:source.read(1024*1024),b''):
                check_cancel(cancel)
                target.write(block)
                digest.update(block)
        yield Path(temporary),digest.hexdigest()
        if sha256_file(path)!=digest.hexdigest():
            raise ValueError('분석 중 원본 PDF가 변경되어 결과를 보류했습니다.')
    finally:
        Path(temporary).unlink(missing_ok=True)

def ocr_engine():
    global _ENGINE
    if _ENGINE is None:
        from rapidocr_onnxruntime import RapidOCR
        import rapidocr_onnxruntime.utils as utils
        # Older RapidOCR releases ignore thread kwargs; set ONNX session options.
        original = utils.SessionOptions
        def options():
            value = original()
            value.intra_op_num_threads = 1
            value.inter_op_num_threads = 1
            return value
        utils.SessionOptions = options
        try:
            _ENGINE = RapidOCR(det_model_path=None, det_limit_type='max', det_limit_side_len=2200)
        finally:
            utils.SessionOptions = original
    return _ENGINE

def compact(value):
    return re.sub(r'[^A-Z0-9]', '', value.upper())

def commercial_header(text):
    value = compact(text)
    return any(word in value for word in ('SHIPPEREXPORTER', 'SHIPPEREXPECTER',
        'COMMERCIALINVOICE', 'PACKINGLIST', 'FORACCOUNTRISK','MATERIALCODE','UPCCODE')) or (
        'HYBE' in value and any(word in value for word in ('INVOICE','IMVOICE')))

def commercial_continuation(header,previous):
    value=compact(header)
    return bool(previous and previous['kind'].startswith('commercial') and not any(
        word in value for word in ('IATANOTICE','IATANOTTCE','CERTIFICATEOFORIGIN','AIRWAYBILL','CARRIERSLIMITATION')))

def group_rows(boxes):
    entries = []
    for points, text, confidence in boxes:
        x1, x2 = min(p[0] for p in points), max(p[0] for p in points)
        y1, y2 = min(p[1] for p in points), max(p[1] for p in points)
        entries.append({'x':(x1+x2)/2, 'y':(y1+y2)/2, 'left':x1,'right':x2,'height':y2-y1,
                        'text':str(text), 'confidence':float(confidence)})
    rows = []
    for entry in sorted(entries, key=lambda e:(e['y'],e['x'])):
        if rows and abs(rows[-1]['y']-entry['y']) <= max(5, entry['height']*.55):
            rows[-1]['cells'].append(entry)
        else:
            rows.append({'y':entry['y'], 'cells':[entry]})
    for row in rows:
        row['cells'].sort(key=lambda e:e['x'])
        row['text'] = ' '.join(c['text'] for c in row['cells'])
    return rows

def page_text(page):
    if page.get('boxes'):
        return '\n'.join(row['text'] for row in group_rows(page['boxes']))
    return page.get('native', '')

def _save_cache(cache_file,data):
    descriptor,temporary=tempfile.mkstemp(prefix='ocr-',suffix='.tmp',dir=cache_file.parent)
    try:
        with os.fdopen(descriptor,'w',encoding='utf-8') as stream:
            json.dump(data,stream,ensure_ascii=False)
        os.replace(temporary,cache_file)
    finally:
        if Path(temporary).exists():
            Path(temporary).unlink()

def read_pdf(path, *, cache_dir=None, progress=None, cancel=None):
    import numpy as np
    import pdfplumber
    import pypdfium2 as pdfium
    path = Path(path)
    cache = Path(cache_dir) if cache_dir else Path(os.getenv('LOCALAPPDATA') or Path.home()) / 'HybeFiller/cache'
    cache.mkdir(parents=True, exist_ok=True)
    with pdf_snapshot(path,cancel) as (snapshot,digest):
        cache_file=cache/(digest+'.json')
        cached_pages=[]
        if cache_file.exists():
            try:
                cached=json.loads(cache_file.read_text(encoding='utf-8'))
                if isinstance(cached,dict) and cached.get('version')==CACHE_VERSION and cached.get('sha256')==digest:
                    candidate=cached.get('pages')
                    if isinstance(candidate,list) and all(isinstance(p,dict) and type(p.get('page')) is int and
                            p['page']==i for i,p in enumerate(candidate,1)):
                        cached_pages=candidate
            except (ValueError,OSError):
                if progress: progress(f'{path.name}: 캐시 손상으로 다시 분석합니다.')
        pages=[]
        with pdfium.PdfDocument(snapshot) as doc,pdfplumber.open(snapshot) as textdoc:
            if len(cached_pages)!=len(doc): cached_pages=[]
            for index in range(len(doc)):
                check_cancel(cancel)
                if progress: progress(f'{path.name}: {index+1}/{len(doc)}쪽 원문 재검증')
                page=doc[index]
                try:
                    textpage=page.get_textpage()
                    try: native=textpage.get_text_range()
                    finally: textpage.close()
                    item={'page':index+1,'kind':'native','native':native,'boxes':[]}
                    # Cached native text and geometry must never approve a value.
                    if len(native.strip())>=100:
                        item['native']=textdoc.pages[index].extract_text() or native
                        if 'AIR WAYBILL' in item['native'].upper():
                            item['words']=textdoc.pages[index].extract_words()
                            item['size']=[textdoc.pages[index].width,textdoc.pages[index].height]
                    else:
                        old=cached_pages[index] if cached_pages else {}
                        if (old.get('kind') in ('commercial','commercial_continuation','other_image')
                            and isinstance(old.get('boxes'),list) and isinstance(old.get('header'),str)
                            and isinstance(old.get('size'),list) and len(old['size'])==2):
                            item.update({k:old[k] for k in ('kind','boxes','header','size')})
                        else:
                            bitmap=page.render(scale=1.5)
                            try: image=bitmap.to_pil().convert('RGB')
                            finally: bitmap.close()
                            image.thumbnail((2200,2200)); item['size']=list(image.size)
                            header=image.crop((0,0,image.width,int(image.height*.36)))
                            header.thumbnail((1200,1200))
                            boxes,_=ocr_engine()(np.asarray(header))
                            item['header']=' '.join(box[1] for box in boxes or [])
                            if commercial_header(item['header']): item['kind']='commercial'
                            elif commercial_continuation(item['header'],pages[-1] if pages else None):
                                item['kind']='commercial_continuation'
                            else: item['kind']='other_image'
                            if item['kind'].startswith('commercial'):
                                check_cancel(cancel)
                                boxes,_=ocr_engine()(np.asarray(image)); item['boxes']=boxes or []
                                if not item['boxes']: raise ValueError(f'{index+1}쪽 상업서류 OCR 결과가 비어 있습니다.')
                    pages.append(item)
                finally: page.close()
        _save_cache(cache_file,{'version':CACHE_VERSION,'sha256':digest,'pages':pages})
    return pages, digest
