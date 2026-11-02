"""One output rule for previews, workbooks and CSV: unproved values stay blank."""
from .field_mapping import FIELD_HEADER_ALIASES
import re
from datetime import date,datetime
from .accuracy import identifier,number,plain_number
from .reference import NUMERIC,DATES

def enforce_field_proofs(record,evidence):
    evidence.setdefault('proofs',{})
    evidence.setdefault('unverified',{})
    evidence.setdefault('issues',[])
    for field in record.__dataclass_fields__:
        if field=='source_notes' or getattr(record,field) is None:
            continue
        value=getattr(record,field)
        proof=evidence['proofs'].get(field,{})
        pages=proof.get('pages')
        valid_pages=isinstance(pages,list) and bool(pages) and all(type(p) is int and 1<=p<=evidence.get('pages',0) for p in pages)
        valid_hash=bool(re.fullmatch(r'[0-9a-f]{64}',str(evidence.get('pdf_sha256',''))))
        bound_hash=proof.get('pdf_sha256')==evidence.get('pdf_sha256')
        valid_value=True
        if field in NUMERIC:
            parsed=number(value,integer=field in {'sku_count','item_quantity','pallet_count'})
            try:
                valid_value=parsed is not None and parsed>0 and plain_number(parsed)==value
            except ValueError:
                valid_value=False
        elif field in DATES:
            valid_value=isinstance(value,date) and not isinstance(value,datetime)
        else:
            valid_value=isinstance(value,str) and bool(value.strip())
        if (proof.get('method') not in ('native_pdf_label','visual_pdf') or proof.get('value')!=value
            or not valid_pages or not valid_hash or not bound_hash or not valid_value):
            evidence.setdefault('rejected_candidates',{})[field]=value
            setattr(record,field,None)
            evidence.get('fields',{}).pop(field,None)
            evidence['proofs'].pop(field,None)
            evidence['unverified'][field]='원문 확정 근거가 없어 자동 후보 값을 출력하지 않았습니다.'
            evidence['issues'].append('빈칸 처리: '+FIELD_HEADER_ALIASES[field][0])

def reject_conflicting_houses(items,allow_duplicates=False):
    # Allowing identical duplicates never chooses a winner among revisions.
    groups={}
    for item in items:
        record=item.get('record')
        if record is not None and record.house_bl_no and item['status']!='cancelled':
            groups.setdefault(identifier(record.house_bl_no),[]).append(item)
    conflicts=[]
    for house,group in groups.items():
        if len(group)>1 and (any(item['status']=='error' for item in group) or
                            len({item.get('evidence',{}).get('pdf_sha256') for item in group})>1):
            reason='동일 House BL에 서로 다른 원본 또는 실패한 수정본이 있어 우선순위를 정하지 않고 모두 보류했습니다.'
            for item in group:
                item.update(status='error',error=reason)
                item['evidence']['issues'].append(reason)
            conflicts.append(house)
    return conflicts
