"""Conservative primary-document extraction. No legacy guesses or Excel backfill."""
from __future__ import annotations
import re
from datetime import date
from pathlib import Path
from .accuracy import identifier,number,plain_number,native_invoice_values,collect_invoice_metrics
from .document_reader import group_rows,page_text,read_pdf
from .models import ShipmentRecord

AIRPORTS={
    'KRICN':('INCHEON','ICN'),'DEFRA':('FRANKFURT','FRA'),
    'AUSYD':('SYDNEY','SYD'),'USEWR':('NEWARK','EWR'),
    'USLAX':('LOS ANGELES','LAX'),'USORD':('CHICAGO','ORD'),'USBOS':('BOSTON','BOS'),
}

def _rows(page,left=None):
    boxes=[]
    for word in page.get('words',[]):
        if left is not None and word['x1']>left:
            continue
        boxes.append([[[word['x0'],word['top']],[word['x1'],word['top']],
                       [word['x1'],word['bottom']],[word['x0'],word['bottom']]],word['text'],1.0])
    rows=group_rows(boxes)
    if left is not None:
        for row in rows:
            row['clipped']=any(w['x0']<left<w['x1'] and
                abs((w['top']+w['bottom'])/2-row['y'])<=max(5,(w['bottom']-w['top'])*.55)
                for w in page.get('words',[]))
    return rows

def _fold(value):
    return re.sub(r'[^A-Z0-9가-힣]','',str(value).upper())

def _awb_house(page):
    # The header identifies the HAWB independently of invoice cross references.
    header='\n'.join(page.get('native','').splitlines()[:3])
    header='\n'.join(line for line in header.splitlines() if 'INVOICE' not in line.upper())
    found={identifier(value) for value in re.findall(
        r'(?<![A-Z0-9_])PLIHQ\s*2\s*[A-Z]\s*\d{5}(?![A-Z0-9])',header.upper())}
    return next(iter(found)) if len(found)==1 else None

def _chargeable_weights(page):
    words=page.get('words',[])
    width,height=page.get('size',[0,0])
    values=[]
    for anchor in [w for w in words if w['text'].upper()=='CHARGEABLE']:
        for word in words:
            if (word['x0']>=anchor['x0']-width*.014 and word['x1']<=anchor['x1']+width*.026
                and word['top']>anchor['bottom'] and word['top']-anchor['bottom']<height*.08):
                value=number(word['text'])
                if value is not None and value>0:
                    values.append((plain_number(value),page['page'],'Chargeable Weight: '+word['text']))
    return values

def _party(page,label):
    rows=_rows(page,page.get('size',[0,0])[0]*.48)
    start=next((i for i,row in enumerate(rows) if _fold(label) in _fold(row['text'])),None)
    if start is None:
        return None,None
    for row in rows[start+1:]:
        if row['y']-rows[start]['y']>page['size'][1]*.08:
            break
        value=row['text'].strip()
        if row.get('clipped'):
            return None,None
        if value and re.match(r'[A-Za-z가-힣]',value) and not any(
            text in _fold(value) for text in ('ACCOUNTNUMBER','AIRWAYBILL','NOTNEGOTIABLE')):
            return value,row
    return None,None

def _airport(page,label):
    rows=_rows(page,page.get('size',[0,0])[0]*.40)
    start=next((i for i,row in enumerate(rows) if _fold(label) in _fold(row['text'])),None)
    if start is None:
        return None,None
    for row in rows[start+1:start+2]:
        if row['y']-rows[start]['y']>page['size'][1]*.05:
            continue
        text=row['text'].upper()
        codes={code for code,aliases in AIRPORTS.items() if any(re.search(r'\b'+re.escape(alias)+r'\b',text) for alias in aliases)}
        if len(codes)==1:
            return next(iter(codes)),row
    return None,None

def parse_pages(pages,fallback_house_bl):
    texts=[page_text(page) for page in pages]
    found={identifier(value) for value in re.findall(r'(?<![A-Z0-9])PLIHQ\s*2\s*[A-Z]\s*\d{5}(?!\d)','\n'.join(texts).upper())}
    fallback=identifier(fallback_house_bl)
    if not found:
        raise ValueError('PDF 본문에서 House BL을 확인하지 못했습니다. 파일명만으로 확정하지 않습니다.')
    if re.fullmatch(r'PLIHQ2[A-Z]\d{5}',fallback) and fallback not in found:
        raise ValueError('파일명 House BL과 PDF 원문이 다릅니다.')
    if len(found)>1 and fallback not in found:
        raise ValueError('서로 다른 House BL이 포함되어 대상 운송장을 확정하지 못했습니다.')
    house=fallback if fallback in found else next(iter(found))
    record=ShipmentRecord()
    evidence={'fields':{},'proofs':{},'issues':[],'candidates':{},'unverified':{},
              'invoice_tables':collect_invoice_metrics(pages),'policy':'primary-evidence-only-v1.3.0'}

    def accept(field,observations,reason):
        observations=[item for item in observations if item[0] not in (None,'')]
        values={str(item[0]) for item in observations}
        if len(values)!=1:
            evidence['unverified'][field]='원문 근거 없음' if not values else '원문 간 값 충돌'
            if len(values)>1:
                evidence['issues'].append(field+': 원문 값이 서로 달라 빈칸으로 보류했습니다.')
                evidence.setdefault('withheld',{})[field]={'method':'native_pdf_conflict','value':None,
                    'pages':sorted({item[1] for item in observations}),'excerpts':[item[2] for item in observations],
                    'reason':'서로 다른 원문 값: '+' / '.join(sorted(values))}
            return
        value=observations[0][0]
        setattr(record,field,value)
        source_pages=sorted({item[1] for item in observations})
        evidence['fields'][field]=reason+' / '+','.join(map(str,source_pages))+'쪽'
        evidence['proofs'][field]={'method':'native_pdf_label','pages':source_pages,
            'excerpts':[item[2] for item in observations],'value':value}

    header_houses={_awb_house(page) for page in pages if 'AIR WAYBILL' in page.get('native','').upper()}
    header_houses.discard(None)
    if len(header_houses)>1:
        raise ValueError('서로 다른 House BL의 HAWB가 한 PDF에 있습니다. 운송장별 원본으로 분리해 주세요.')
    awbs=[page for page in pages if 'AIR WAYBILL' in page.get('native','').upper()
          and _awb_house(page)==house]
    manifests=[page for page in pages if 'AIR CARGO MANIFEST' in page.get('native','').upper()
               and house in identifier(page.get('native',''))]
    accept('house_bl_no',[(house,p['page'],house) for p in awbs or manifests],
           'PDF HAWB/Manifest House BL 명시')
    if record.house_bl_no is None:
        raise ValueError('HAWB/Manifest 원문에서 대상 House BL을 확정하지 못했습니다.')

    invoice_values=[]
    foreign_invoices=[]
    for page in awbs:
        for value in native_invoice_values([page],house):
            observation=(value,page['page'],'INVOICE NO: '+value)
            suffix=re.search(r'PLIHQ2[A-Z]\d{5}$',value or '')
            if suffix and identifier(suffix.group())!=house:
                evidence['issues'].append('HAWB Invoice No.에 다른 House BL이 적혀 있어 인보이스 번호를 보류했습니다.')
                foreign_invoices.append(observation)
                continue
            if value and not re.search(r'_PLIHQ2[A-Z]\d{0,4}$',value) and not re.fullmatch(r'\d{2,4}[-.]\d{3,4}[-.]\d{4}',value):
                invoice_values.append(observation)
    accept('invoice_no',invoice_values+foreign_invoices if invoice_values else [],'PDF HAWB Invoice No. 명시, 추정/부분복원 없음')

    for field,label in [('shipper',"Shipper's Name and Address"),('consignee',"Consignee's Name and Address")]:
        observations=[]
        for page in awbs:
            value,row=_party(page,label)
            if value:
                observations.append((value,page['page'],row['text']))
        accept(field,observations,'PDF HAWB '+label+'의 첫 회사명 행')

    people=[]
    for page in awbs:
        rows=_rows(page,page.get('size',[0,0])[0]*.48)
        start=next((i for i,r in enumerate(rows) if 'SHIPPERSNAMEANDADDRESS' in _fold(r['text'])),None)
        end=next((i for i,r in enumerate(rows) if 'CONSIGNEESNAMEANDADDRESS' in _fold(r['text'])),len(rows))
        if start is not None:
            for row in rows[start+1:end]:
                match=re.match(r"([A-Za-z][A-Za-z .'-]{2,60})\s*/\s*\+?\d",row['text'])
                if match:
                    people.append((match.group(1).strip(),page['page'],row['text']))
    evidence['candidates']['shipper_contacts']=[{'name':name,'page':page,'excerpt':excerpt} for name,page,excerpt in people]
    evidence['unverified']['requester']='송하인 연락담당자와 업무 요청자가 동일하다는 근거가 없어 빈칸으로 남겼습니다.'

    weights=[]
    for page in awbs:
        weights.extend(_chargeable_weights(page))
    accept('weight_quantity',weights,'PDF HAWB Chargeable Weight 명시')
    if record.weight_quantity is not None:
        accept('unit',[('CWT',page['page'],'Chargeable Weight') for page in awbs],
               '중량 기준: HAWB Chargeable Weight')
    accept('transport_mode',[('AIR',page['page'],'Air Waybill') for page in awbs],'PDF Air Waybill')

    for field,label in [('pol','Airport of Departure'),('pod','Airport of Destination')]:
        observations=[]
        for page in awbs:
            value,row=_airport(page,label)
            if value:
                observations.append((value,page['page'],row['text']))
        accept(field,observations,'PDF HAWB '+label+'의 공항명/코드')

    flights,departures=[],[]
    for page in manifests:
        text=page['native']
        flight=re.search(r'FLIGHT\s*NO\.?\s*[:：]\s*([A-Z0-9]{2}\d{2,4}[A-Z]?)\b',text,re.I)
        if flight:
            flights.append((flight.group(1).upper(),page['page'],flight.group(0)))
        departed=re.search(r'ONBOARD\s*DATE\s*[:：]\s*(\d{4}-\d{2}-\d{2})',text,re.I)
        if departed:
            departures.append((date.fromisoformat(departed.group(1)),page['page'],departed.group(0)))
    accept('vessel_voy',flights,'PDF Manifest Flight No. 명시')
    accept('atd',departures,'PDF Manifest Onboard Date 명시')

    terms=[]
    for page in pages:
        page_houses={identifier(value) for value in re.findall(
            r'(?<![A-Z0-9])PLIHQ\s*2\s*[A-Z]\s*\d{5}(?!\d)',page.get('native','').upper())}
        if page_houses!={house}:
            continue
        for match in re.finditer(r'\bINCOTERMS\s*[:：]?\s*(EXW|FCA|CPT|CIP|DAP|DPU|DDP|FAS|FOB|CFR|CIF)\b',page.get('native',''),re.I):
            terms.append((match.group(1).upper(),page['page'],match.group(0)))
    accept('incoterms',terms,'PDF Incoterms 명시 필드 (통관 FOB 금액 제외)')

    for field in ('invoice_value_usd','item_quantity','sku_count'):
        evidence['unverified'][field]='이미지 숫자는 직접 원문 대조한 검토확정 파일이 있어야 입력합니다.'
    for field in record.__dataclass_fields__:
        if field!='source_notes' and getattr(record,field) is None:
            evidence['unverified'].setdefault(field,'원문에서 해당 항목의 명시적 근거가 확인되지 않았습니다.')
    evidence['candidates']['invoice_tables']=evidence['invoice_tables']
    return record,evidence

def parse_pdf_record(pdf_path,*,progress=None,cancel=None,cache_dir=None):
    pages,digest=read_pdf(pdf_path,cache_dir=cache_dir,progress=progress,cancel=cancel)
    record,evidence=parse_pages(pages,Path(pdf_path).stem)
    evidence.update(pdf_sha256=digest,pages=len(pages),
                    commercial_pages=[p['page'] for p in pages if p['kind'].startswith('commercial')])
    for proof in evidence['proofs'].values():
        proof['pdf_sha256']=digest
    return record,evidence
