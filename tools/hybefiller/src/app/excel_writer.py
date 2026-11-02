"""Append to a preserved template with explicit headers and atomic, no-overwrite save."""
from __future__ import annotations
import os
import tempfile
import hashlib
from io import BytesIO
from copy import copy
from datetime import date,datetime
from pathlib import Path
import openpyxl
from .field_mapping import FIELD_ORDER, FIELD_HEADER_ALIASES, find_header_row
from .accuracy import identifier,number
from .provenance import enforce_field_proofs,reject_conflicting_houses

def _load_template(template,expected_sha256=None):
    raw=template.read_bytes()
    digest=hashlib.sha256(raw).hexdigest()
    if expected_sha256 is not None and digest!=expected_sha256:
        raise ValueError('처리 중 양식이 변경되어 저장을 중단했습니다.')
    book=openpyxl.load_workbook(BytesIO(raw),keep_vba=template.suffix.lower()=='.xlsm')
    try:
        matches=[(sheet,result) for sheet in book if (result:=find_header_row(sheet))]
        if len(matches)!=1:
            raise ValueError(f'양식에서 입력 시트를 하나로 확정할 수 없습니다({len(matches)}개).')
        sheet,(header,mapping)=matches[0]
        missing=set(FIELD_ORDER)-set(mapping)
        if missing:
            raise ValueError('양식에 필요한 헤더가 없습니다: '+', '.join(sorted(missing)))
        for row in sheet.iter_rows(min_row=header+1):
            if any(row[col-1].value is not None for col in mapping.values()):
                raise ValueError('업무 값이 들어 있는 양식은 사용할 수 없습니다. 빈 양식을 선택하고 기존 실적은 비교할 엑셀로 지정해 주세요.')
        return book,sheet,header,mapping,digest
    except Exception:
        book.close()
        raise


def check_template(path):
    template=Path(path)
    if template.suffix.lower() not in ('.xlsx','.xlsm'):
        raise ValueError('양식은 .xlsx 또는 .xlsm 파일이어야 합니다.')
    book,_,_,_,digest=_load_template(template)
    book.close()
    return digest


def write_results(template_path,output_path,items,*,allow_duplicates=False,template_sha256=None,pre_publish=None):
    template,output=Path(template_path).resolve(),Path(output_path).resolve()
    if template==output or output.exists():
        raise ValueError('기존 파일을 덮어쓰지 않습니다. 새로운 출력 파일명을 선택해 주세요.')
    if template.suffix.lower() not in ('.xlsx','.xlsm'):
        raise ValueError('양식은 .xlsx 또는 .xlsm 파일이어야 합니다.')
    if output.suffix.lower()!=template.suffix.lower():
        raise ValueError('양식과 출력 파일 확장자가 같아야 합니다. 매크로 파일은 .xlsm으로 저장해 주세요.')
    book,sheet,header,mapping,_=_load_template(template,template_sha256)
    try:
        existing=set()
        last=header
        style_row=max(header+1,last)
        written=0
        reject_conflicting_houses(items,allow_duplicates)
        for item in items:
            if item['status'] in ('error','cancelled'):
                continue
            record=item['record']
            enforce_field_proofs(record,item.setdefault('evidence',{}))
            key=identifier(record.house_bl_no)
            if not key:
                item.update(status='error',error='원문 확정 House BL이 없어 결과 행을 만들지 않았습니다.')
                continue
            if item['status']=='verified' and item['evidence'].get('issues'):
                item['status']='review'
            if key and key in existing and not allow_duplicates:
                item['status']='skipped_duplicate'
                continue
            last+=1
            if style_row<last:
                for col in range(1,sheet.max_column+1):
                    source,target=sheet.cell(style_row,col),sheet.cell(last,col)
                    target._style=copy(source._style)
                    target.alignment=copy(source.alignment)
                    target.protection=copy(source.protection)
                sheet.row_dimensions[last].height=sheet.row_dimensions[style_row].height
            for field,column in mapping.items():
                value=getattr(record,field)
                cell=sheet.cell(last,column)
                if isinstance(value,str) and value.startswith(('=','+','-','@')):
                    cell.value=value
                    cell.data_type='s'
                else:
                    cell.value=value
                if isinstance(value,str):
                    alignment=copy(cell.alignment)
                    alignment.wrap_text=True
                    cell.alignment=alignment
                if field in ('house_bl_no','invoice_no','order_no'):
                    if value is not None:
                        cell.value=str(value)
                        cell.data_type='s'
                    cell.number_format='@'
                elif isinstance(value,(date,datetime)):
                    cell.number_format='yyyy-mm-dd'
                elif field in ('invoice_value_usd','weight_quantity','item_quantity','sku_count','pallet_count') and value is not None:
                    places=max(0,-number(value).as_tuple().exponent)
                    if field=='invoice_value_usd': places=max(2,places)
                    cell.number_format=('#,##0'+('.'+'0'*places if places else '')) if places<=15 else '0.##############E+00'
            item['target_row']=last
            if key:
                existing.add(key)
            written+=1
        # Keep diagnostic data outside the user's reporting columns.
        report=book.create_sheet('검토내역' if '검토내역' not in book.sheetnames else '검토내역_추가')
        report.append(['PDF','House BL','상태','항목','내용'])
        for item in items:
            house=item['record'].house_bl_no if item.get('record') else ''
            issues=item.get('evidence',{}).get('issues',[])
            if item.get('error'):
                issues=issues+[item['error']]
            status={'verified':'필수값 원문 확정','review':'확인 필요','error':'실패','skipped_duplicate':'중복 제외','cancelled':'중지'}[item['status']]
            report.append([Path(item['pdf']).name,house,status,'요약',' / '.join(issues)])
            for field,origin in item.get('evidence',{}).get('fields',{}).items():
                report.append([Path(item['pdf']).name,house,status,FIELD_HEADER_ALIASES.get(field,[field])[0],origin])
            for field,reason in item.get('evidence',{}).get('unverified',{}).items():
                if getattr(item['record'],field,None) is None:
                    report.append([Path(item['pdf']).name,house,'미확정·빈칸',FIELD_HEADER_ALIASES.get(field,[field])[0],reason])
            for note in item.get('evidence',{}).get('comparison_notes',[]):
                report.append([Path(item['pdf']).name,house,'참고 대조·원문 유지','엑셀 비교',note])
            for difference in item.get('evidence',{}).get('reference',{}).get('conflicts',[]):
                report.append([Path(item['pdf']).name,house,'참고 대조·원문 유지',FIELD_HEADER_ALIASES[difference['field']][0],
                    '원문 확정값: '+difference['pdf']+' / 참고 엑셀: '+', '.join(difference['reference'])])
        report.freeze_panes='A2'
        report.auto_filter.ref=report.dimensions
        for col,width in {'A':34,'B':22,'C':22,'D':28,'E':100}.items():
            report.column_dimensions[col].width=width
        # Force text for any external value to prevent Excel formula injection.
        for row in report:
            for cell in row:
                if isinstance(cell.value,str):
                    cell.data_type='s'
        proof_sheet=book.create_sheet('원문근거')
        proof_sheet.append(['PDF','House BL','항목','출력값','원문 페이지','확정 방식','확정·보류 사유 / 발췌','PDF SHA256'])
        for item in items:
            record=item.get('record')
            if record is None or item['status'] in ('error','cancelled'):
                continue
            evidence=item.get('evidence',{})
            for field in FIELD_ORDER:
                proof=evidence.get('proofs',{}).get(field) or evidence.get('withheld',{}).get(field,{})
                proof_sheet.append([Path(item['pdf']).name,record.house_bl_no,FIELD_HEADER_ALIASES[field][0],
                    getattr(record,field),','.join(map(str,proof.get('pages',[]))),
                    proof.get('method','미확정·빈칸'),proof.get('reason') or ' / '.join(proof.get('excerpts',[])) or evidence.get('unverified',{}).get(field,''),
                    evidence.get('pdf_sha256','')])
        for diagnostic in (report,proof_sheet):
            for row in diagnostic:
                for cell in row:
                    if isinstance(cell.value,str):
                        cell.data_type='s'
                    alignment=copy(cell.alignment)
                    alignment.wrap_text=True
                    cell.alignment=alignment
                    if isinstance(cell.value,(date,datetime)):
                        cell.number_format='yyyy-mm-dd'
        proof_sheet.freeze_panes='D2'
        proof_sheet.auto_filter.ref=proof_sheet.dimensions
        for col,width in {'A':30,'B':22,'C':28,'D':48,'E':16,'F':22,'G':95,'H':68}.items():
            proof_sheet.column_dimensions[col].width=width
        output.parent.mkdir(parents=True,exist_ok=True)
        descriptor,temporary=tempfile.mkstemp(prefix='hybe-',suffix=output.suffix,dir=output.parent)
        os.close(descriptor)
        try:
            book.save(temporary)
            check=openpyxl.load_workbook(temporary,read_only=True,data_only=False,keep_vba=output.suffix=='.xlsm')
            rows=()
            try:
                expected_rows={item['target_row']:item['record'] for item in items if 'target_row' in item}
                rows=check[sheet.title].iter_rows(min_row=min(expected_rows),max_row=max(expected_rows),values_only=True) if expected_rows else ()
                for row_no,values in enumerate(rows,min(expected_rows) if expected_rows else 1):
                    if row_no not in expected_rows: continue
                    for field,column in mapping.items():
                        actual=values[column-1]
                        expected=getattr(expected_rows[row_no],field)
                        if isinstance(actual,datetime) and isinstance(expected,date): actual=actual.date()
                        if actual!=expected:
                            raise ValueError('저장 후 값 대조 실패: '+FIELD_HEADER_ALIASES[field][0])
            finally:
                if hasattr(rows,'close'): rows.close()
                check.close()
            if pre_publish: pre_publish()
            os.rename(temporary,output)  # Windows refuses to overwrite a newly created destination.
        finally:
            if Path(temporary).exists():
                Path(temporary).unlink()
        return written
    finally:
        book.close()
