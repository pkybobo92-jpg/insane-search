"""Apply explicit visual review decisions only to the exact reviewed PDF bytes."""
import json
import re
import hashlib
from datetime import date
from pathlib import Path
from .accuracy import number,plain_number,identifier
from .models import ShipmentRecord
from .reference import NUMERIC,DATES

def _unique_object(pairs):
    result={}
    for key,value in pairs:
        if key in result:
            raise ValueError('검토확정 JSON에 중복 키가 있습니다: '+key)
        result[key]=value
    return result

class ReviewedIndex:
    def __init__(self,path):
        self.path=Path(path)
        raw=self.path.read_bytes()
        self.sha256=hashlib.sha256(raw).hexdigest()
        data=json.loads(raw.decode('utf-8-sig'),object_pairs_hook=_unique_object)
        if (not isinstance(data,dict) or type(data.get('schema')) is not int or data['schema']!=1
            or not isinstance(data.get('reviewer'),str) or not data['reviewer'].strip()
            or data.get('review_basis')!='visual_pdf'):
            raise ValueError('직접 PDF 대조 검토확정 파일 형식이 아닙니다.')
        self.reviewer=data['reviewer']
        self.documents=data.get('documents',{})
        if not isinstance(self.documents,dict):
            raise ValueError('검토확정 documents는 해시별 객체여야 합니다.')
        for digest,entry in self.documents.items():
            if not re.fullmatch(r'[0-9a-f]{64}',digest) or not isinstance(entry,dict) or not isinstance(entry.get('fields'),dict):
                raise ValueError('검토확정 PDF 해시 또는 항목 형식이 잘못되었습니다.')

    def apply(self,record,evidence):
        entry=self.documents.get(evidence['pdf_sha256'])
        if entry is None:
            evidence['issues'].append('이 PDF의 현재 SHA256에 대응하는 직접 검토확정 값이 없습니다.')
            return
        if identifier(entry.get('house_bl_no'))!=identifier(record.house_bl_no):
            raise ValueError('검토확정 파일의 House BL과 PDF 원문이 다릅니다.')
        allowed=set(ShipmentRecord.__dataclass_fields__)-{'source_notes','house_bl_no'}
        checked=[]
        for field,decision in entry['fields'].items():
            if (field not in allowed or not isinstance(decision,dict) or 'value' not in decision
                or not isinstance(decision.get('reason'),str) or not decision['reason'].strip()):
                raise ValueError('검토확정 필드 또는 검토 근거가 잘못되었습니다: '+str(field))
            pages=decision.get('pages')
            if not isinstance(pages,list) or not pages or any(type(p) is not int or not 1<=p<=evidence['pages'] for p in pages):
                raise ValueError('검토확정 원문 페이지 범위가 잘못되었습니다: '+field)
            value=decision.get('value')
            if value is not None and field in NUMERIC:
                parsed=number(value,integer=field in {'sku_count','item_quantity','pallet_count'})
                if parsed is None or parsed<=0:
                    raise ValueError('검토확정 숫자가 유효하지 않습니다: '+field)
                value=plain_number(parsed)
            elif value is not None and field in DATES:
                if not isinstance(value,str) or not re.fullmatch(r'\d{4}-\d{2}-\d{2}',value):
                    raise ValueError('검토확정 날짜는 YYYY-MM-DD 형식이어야 합니다: '+field)
                value=date.fromisoformat(value)
            elif value is not None and (not isinstance(value,str) or not value.strip()):
                raise ValueError('검토확정 텍스트가 유효하지 않습니다: '+field)
            if field=='invoice_no' and value is not None:
                suffix=re.search(r'PLIHQ2[A-Z]\d+$',identifier(value))
                if suffix and suffix.group()!=identifier(record.house_bl_no):
                    raise ValueError('검토확정 Invoice No.의 House BL이 원문과 다릅니다.')
            checked.append((field,value,decision,pages))
        # Validate the whole decision set before changing either shared object.
        for field,value,decision,pages in checked:
            if value is None:
                setattr(record,field,None)
                evidence['proofs'].pop(field,None)
                evidence['fields'].pop(field,None)
                evidence['unverified'][field]=decision['reason']
                evidence.setdefault('withheld',{})[field]={'method':'visual_pdf_withheld','pages':pages,
                    'value':None,'reason':decision['reason'],'reviewer':self.reviewer,'pdf_sha256':evidence['pdf_sha256']}
                continue
            setattr(record,field,value)
            evidence['unverified'].pop(field,None)
            evidence.get('withheld',{}).pop(field,None)
            evidence['fields'][field]='PDF 원문 이미지 직접 대조 확정 / '+','.join(map(str,pages))+'쪽'
            evidence['proofs'][field]={'method':'visual_pdf','reviewer':self.reviewer,'pages':pages,
                                      'value':value,'reason':decision['reason'],'pdf_sha256':evidence['pdf_sha256']}
        evidence['reviewed_file']=str(self.path)
