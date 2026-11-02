"""The reference workbook is a comparison source, never a source of output values."""
from collections import defaultdict
from datetime import date,datetime
from pathlib import Path
import re
import unicodedata
import openpyxl
import hashlib
from io import BytesIO
from .accuracy import identifier,number
from .field_mapping import find_header_row,FIELD_HEADER_ALIASES

NUMERIC={'invoice_value_usd','weight_quantity','sku_count','item_quantity','pallet_count'}
DATES={'order_received_date','booking_notice_date','outbound_date','atd','ata',
       'requested_delivery_date','arrival_destination_date','billing_completed_date'}

def blank(value):
    return value is None or not str(value).strip() or str(value).strip()=='확인필요'

def invoice_key(value,house):
    value=re.sub(r'\s+','',unicodedata.normalize('NFKC',str(value or ''))).upper()
    suffix=identifier(house)
    return value[:-len(suffix)-1] if suffix and value.endswith('_'+suffix) else value

def invoice_matches(value,key,house):
    if invoice_key(value,house)==key:
        return True
    is_integer=isinstance(value,int) and not isinstance(value,bool) or isinstance(value,float) and value.is_integer()
    match=re.fullmatch(r'(\d+)(?:_\d{4}_(?:AIR|SEA)\d+)?',key)
    return bool(is_integer and match and int(value)==int(match.group(1)))

def canonical(value,field):
    if blank(value):
        return None
    if field in NUMERIC:
        return number(value)
    if field in DATES:
        if isinstance(value,(date,datetime)):
            return value.strftime('%Y-%m-%d')
        return str(value).strip().replace('T00:00:00','').replace(' 00:00:00','')
    return identifier(value)

class ReferenceIndex:
    def __init__(self,path):
        self.path=Path(path)
        self.rows=defaultdict(list)
        raw=self.path.read_bytes()
        self.sha256=hashlib.sha256(raw).hexdigest()
        book=openpyxl.load_workbook(BytesIO(raw),data_only=True,read_only=True)
        try:
            for sheet in book:
                header=find_header_row(sheet)
                if not header:
                    continue
                row_index,mapping=header
                for row_no,values in enumerate(sheet.iter_rows(min_row=row_index+1,values_only=True),row_index+1):
                    data={key:values[col-1] if col<=len(values) else None for key,col in mapping.items()}
                    if key:=identifier(data.get('house_bl_no')):
                        self.rows[key].append({'values':data,'sheet':sheet.title,'row':row_no})
        finally:
            book.close()
        if not self.rows:
            raise ValueError('대조 엑셀에서 House BL / Invoice 헤더와 데이터를 찾지 못했습니다.')

    def reconcile(self,record,evidence):
        key=invoice_key(record.invoice_no,record.house_bl_no)
        matched=[row for row in self.rows.get(identifier(record.house_bl_no),[])
                 if key and invoice_matches(row['values'].get('invoice_no'),key,record.house_bl_no)]
        info={'comparison_only':True,'path':str(self.path),'matched_rows':[
              {'sheet':row['sheet'],'row':row['row']} for row in matched],
              'filled':[],'not_used':[],'conflicts':[]}
        evidence['reference']=info
        notes=evidence.setdefault('comparison_notes',[])
        if not matched:
            notes.append('참고 엑셀 대조 행 없음: 출력 값은 원문 근거만 사용합니다.')
            return
        for field in FIELD_HEADER_ALIASES:
            values=[row['values'].get(field) for row in matched]
            distinct={canonical(value,field) for value in values}
            distinct.discard(None)
            actual=getattr(record,field)
            if not distinct:
                continue
            if blank(actual):
                info['not_used'].append({'field':field,'reason':'원문 확정 근거가 없어 엑셀 값으로 채우지 않음'})
                continue
            if len(distinct)>1 or canonical(actual,field) not in distinct:
                info['conflicts'].append({'field':field,'pdf':str(actual),'reference':[str(v) for v in values],
                    'decision':'원문 확정값 유지; 엑셀 값 미반영'})
                notes.append('참고 엑셀과 다름 (원문 확정값 유지): '+FIELD_HEADER_ALIASES[field][0])
