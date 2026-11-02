"""Exact numeric/identity parsing. OCR text is evidence for review, never approval."""
from __future__ import annotations
import re
import unicodedata
import math
from decimal import Decimal,InvalidOperation
from .document_reader import compact,group_rows,page_text

INCOTERMS=('EXW','FCA','CPT','CIP','DAP','DPU','DDP','FAS','FOB','CFR','CIF')

def identifier(value):
    if value is None:
        return ''
    if isinstance(value,float) and value.is_integer():
        value=int(value)
    return re.sub(r'[\s\-–—_]','',unicodedata.normalize('NFKC',str(value))).upper()

def number(value,*,integer=False):
    text=str(value).strip()
    # Commas are accepted only as explicit three-digit separators. No O/0, l/1,
    # missing decimal, locale or arithmetic-based character repairs.
    if not re.fullmatch(r'(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?',text):
        return None
    try:
        result=Decimal(text.replace(',',''))
    except InvalidOperation:
        return None
    if not result.is_finite() or result<0 or (integer and result!=result.to_integral_value()):
        return None
    return result

def plain_number(value):
    # Excel numeric cells have at most 15 significant decimal digits. Check the
    # original Decimal before converting to float, so lost digits cannot pass.
    digits=list(value.as_tuple().digits)
    while len(digits)>1 and digits[-1]==0:
        digits.pop()
    if len(digits)>15 or (value and value.adjusted() < -307):
        raise ValueError('Excel 숫자 정밀도로 원문 값을 보존할 수 없습니다.')
    try:
        converted=float(value)
    except (OverflowError,ValueError):
        raise ValueError('Excel 숫자 범위를 벗어났습니다.') from None
    if not math.isfinite(converted) or Decimal(str(converted))!=value:
        raise ValueError('숫자 변환 시 원문 값이 달라집니다.')
    return int(value) if value==value.to_integral_value() else float(value)

def collect_invoice_metrics(pages):
    """Retain raw table rows to locate review evidence; never repair or total OCR."""
    result=[]
    for page in pages:
        text=compact(page_text(page))
        if page.get('boxes') and all(word in text for word in ('UNITPRICE','AMOUNT','USD')):
            result.append({'page':page['page'],'auto_approved':False,'source':'OCR; direct visual review required',
                'raw_rows':[row['text'] for row in group_rows(page['boxes'])]})
    return result

def native_invoice_values(pages,house_bl):
    values=[]
    for page in pages:
        text=page.get('native','')
        if 'AIR WAYBILL' not in text.upper() or identifier(house_bl) not in identifier(text):
            continue
        for match in re.finditer(r'(?mi)INVOICE\s*NO\.?[ \t]*[:：]?[ \t]*([^\r\n]*)',text):
            if not re.search(r'\d',match.group(1)):
                continue
            value=re.sub(r'[ \t]+','',match.group(1)).upper().rstrip('.,')
            if not re.fullmatch(r'[A-Z0-9][A-Z0-9_,./\-]{3,}',value) or value.endswith('_'):
                continue
            if not re.fullmatch(r'\d{3}[.\-]\d{3}[.\-]\d{4}',value):
                values.append(value)
    return values

def native_invoice(pages,house_bl):
    values=set(native_invoice_values(pages,house_bl))
    return next(iter(values)) if len(values)==1 else None
