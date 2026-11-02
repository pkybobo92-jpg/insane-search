"""Batch coordinator. Errors remain visible and source files are never modified."""
from __future__ import annotations
import csv
import json
import traceback
from datetime import datetime
from pathlib import Path

from .accuracy import number
from .document_reader import Cancelled, check_cancel, sha256_file
from .excel_writer import write_results,check_template
from .pdf_parser import parse_pdf_record
from .reference import ReferenceIndex, blank
from .reviewed import ReviewedIndex
from .provenance import enforce_field_proofs,reject_conflicting_houses
from .field_mapping import FIELD_HEADER_ALIASES

VERSION='1.3.1'
REQUIRED=('house_bl_no','invoice_no','invoice_value_usd','item_quantity','sku_count')
STATUS_LABELS={'verified':'필수값 원문 확정','review':'확인 필요','error':'실패',
               'skipped_duplicate':'중복 제외','cancelled':'중지'}

def serialize(value):
    if hasattr(value,'to_serializable_dict'):
        return value.to_serializable_dict()
    if isinstance(value,(datetime,)):
        return value.isoformat()
    return str(value)

def validate(record):
    problems=[]
    for field in REQUIRED:
        if blank(getattr(record,field)):
            problems.append('필수값 미확인: '+FIELD_HEADER_ALIASES[field][0])
    for field in ('invoice_value_usd','item_quantity','sku_count','weight_quantity','pallet_count'):
        value=getattr(record,field)
        if not blank(value) and (number(value) is None or number(value)<=0):
            problems.append('양수 숫자 검증 실패: '+FIELD_HEADER_ALIASES[field][0])
    return problems

def run_batch(pdf_paths,template_path,output_path,*,reference_path=None,reviewed_path=None,allow_duplicates=False,
              dry_run=False,progress=None,cancel=None,cache_dir=None):
    paths=[Path(path).resolve() for path in pdf_paths]
    template,output=Path(template_path).resolve(),Path(output_path).resolve()
    if not paths:
        raise ValueError('PDF를 한 개 이상 선택해 주세요.')
    if len(set(paths))!=len(paths):
        raise ValueError('같은 PDF 파일이 두 번 선택되었습니다.')
    for path in [template]+paths+([Path(reference_path).resolve()] if reference_path else [])+([Path(reviewed_path).resolve()] if reviewed_path else []):
        if not path.is_file():
            raise FileNotFoundError('파일을 찾을 수 없습니다: '+str(path))
        if path==output:
            raise ValueError('출력 경로는 원본 파일과 달라야 합니다.')
    for path in paths:
        if path.suffix.lower()!='.pdf':
            raise ValueError('PDF 확장자가 아닙니다: '+str(path))
    if output.exists():
        raise FileExistsError('출력 파일이 이미 있습니다. 새 파일명을 사용해 주세요.')
    output.parent.mkdir(parents=True,exist_ok=True)
    log_path=output.with_suffix('.audit.jsonl')
    review_path=output.with_suffix('.review.csv')
    if log_path.exists() or review_path.exists():
        raise FileExistsError('같은 이름의 실행 로그가 있습니다. 새 파일명을 사용해 주세요.')
    template_hash=check_template(template)
    reference=ReferenceIndex(reference_path) if reference_path else None
    reviewed=ReviewedIndex(reviewed_path) if reviewed_path else None
    result={'version':VERSION,'output':str(output),'audit':str(log_path),'review':str(review_path),
            'dry_run':dry_run,'items':[],'written':0,'cancelled':False}
    def notify(message):
        if progress:
            progress(message)
    with log_path.open('x',encoding='utf-8') as log:
        def audit(event,**data):
            log.write(json.dumps({'time':datetime.now().astimezone().isoformat(),'event':event,**data},
                                 ensure_ascii=False,default=serialize)+'\n')
            log.flush()
        audit('started',version=VERSION,pdf_count=len(paths),template=str(template),
              template_sha256=template_hash,reference=str(reference_path or ''),
              reference_sha256=reference.sha256 if reference else None,
              reviewed_sha256=reviewed.sha256 if reviewed else None,dry_run=dry_run)
        seen=set()
        for index,path in enumerate(paths,1):
            item={'pdf':str(path),'record':None,'status':'error','evidence':{},'error':None}
            try:
                check_cancel(cancel)
                notify(f'[{index}/{len(paths)}] {path.name}')
                record,evidence=parse_pdf_record(path,progress=notify,cancel=cancel,cache_dir=cache_dir)
                item.update(record=record,evidence=evidence)
                if reviewed:
                    reviewed.apply(record,evidence)
                enforce_field_proofs(record,evidence)
                if reference:
                    reference.reconcile(record,evidence)
                evidence['issues']+=validate(record)
                record.source_notes=list(dict.fromkeys(evidence['issues']))
                if record.house_bl_no in seen and not allow_duplicates:
                    item['status']='skipped_duplicate'
                    evidence['issues'].append('배치 내 중복 House BL: 결과 행을 추가하지 않았습니다.')
                else:
                    seen.add(record.house_bl_no)
                    item['status']='review' if evidence['issues'] else 'verified'
                notify(f'{path.name}: '+STATUS_LABELS[item['status']])
            except Cancelled as exc:
                item.update(status='cancelled',error=str(exc))
                result['cancelled']=True
            except Exception as exc:
                item.update(status='error',error=str(exc),traceback=traceback.format_exc())
                notify(f'{path.name}: 실패 — {exc}')
            result['items'].append(item)
            audit('parsed',**item)
            if result['cancelled']:
                break
        try:
            guards={template:template_hash}
            if reference: guards[reference.path.resolve()]=reference.sha256
            if reviewed: guards[reviewed.path.resolve()]=reviewed.sha256
            for item in result['items']:
                if item['status'] in ('error','cancelled'): continue
                path=Path(item['pdf'])
                expected=item['evidence'].get('pdf_sha256')
                if sha256_file(path)!=expected:
                    reason='처리 중 원본 PDF가 변경되어 이 문서는 출력하지 않았습니다.'
                    item.update(status='error',error=reason)
                    item['evidence']['issues'].append(reason)
                    audit('source_changed',pdf=str(path),decision='record withheld')
                else:
                    guards[path]=expected
            def assert_inputs_unchanged():
                for path,expected in guards.items():
                    if sha256_file(path)!=expected:
                        raise ValueError('저장 전 원본 변경 감지: '+path.name)
            if conflicts:=reject_conflicting_houses(result['items'],allow_duplicates):
                audit('conflicting_house_documents',houses=conflicts,decision='none of the conflicting records written')
            assert_inputs_unchanged()
            if not dry_run:
                result['written']=write_results(template,output,result['items'],allow_duplicates=allow_duplicates,
                    template_sha256=template_hash,pre_publish=assert_inputs_unchanged)
            with review_path.open('x',encoding='utf-8-sig',newline='') as stream:
                writer=csv.writer(stream)
                writer.writerow(['PDF','House BL','상태','출력 행','금액','수량','SKU','검토 사유','미확정 빈칸 항목'])
                for item in result['items']:
                    source_record=item['record']
                    rec=item['record'] if item['status'] not in ('error','cancelled') else None
                    values=[Path(item['pdf']).name,source_record.house_bl_no if source_record else '',STATUS_LABELS[item['status']],item.get('target_row',''),
                        rec.invoice_value_usd if rec else '',rec.item_quantity if rec else '',rec.sku_count if rec else '',
                        ' / '.join(item.get('evidence',{}).get('issues',[])+([item['error']] if item['error'] else [])),
                        ', '.join(FIELD_HEADER_ALIASES[key][0] for key in item.get('evidence',{}).get('unverified',{}) if rec is None or getattr(rec,key,None) is None)]
                    # CSV is intended for Excel; quote formula-like external strings explicitly.
                    writer.writerow(["'"+v if isinstance(v,str) and v.startswith(('=','+','-','@')) else v for v in values])
        except Exception:
            audit('save_failed',traceback=traceback.format_exc())
            raise
        result['counts']={status:sum(item['status']==status for item in result['items'])
                          for status in ('verified','review','error','skipped_duplicate','cancelled')}
        audit('completed',written=result['written'],counts=result['counts'],cancelled=result['cancelled'],
              output_sha256=sha256_file(output) if output.exists() else None)
    return result
