import copy
import hashlib
import json
import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch
import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'src'))
from app.accuracy import number,identifier,native_invoice,collect_invoice_metrics
from app.models import ShipmentRecord
from app.pdf_parser import parse_pages
from app.provenance import enforce_field_proofs
from app.reviewed import ReviewedIndex
from app.reference import ReferenceIndex
from app.excel_writer import write_results
from app.service import run_batch
from app.field_mapping import FIELD_ORDER,FIELD_HEADER_ALIASES

HOUSE='PLIHQ2H00001'
DIGEST='a'*64

def word(text,x,y,right=None):
    return {'text':text,'x0':x,'x1':right if right is not None else x+35,'top':y,'bottom':y+10}

def awb(weight='88.5',page=1,invoice='00000123_'+HOUSE):
    return {'page':page,'kind':'native','native':f'{HOUSE}\nAir Waybill\nINVOICE NO : {invoice}\n1 65.0 K\n',
        'boxes':[],'size':[600,800],'words':[word('Chargeable',200,400,245),word('Weight',205,412),
            word(weight,212,435,248),word('65.0',80,435),word('999.0',330,435)]}

def proved(record):
    values={k:v for k,v in vars(record).items() if k!='source_notes' and v is not None}
    return {'proofs':{k:{'method':'native_pdf_label','pages':[1],'value':v,'pdf_sha256':DIGEST,'excerpts':['Synthetic explicit label']} for k,v in values.items()},
        'fields':{k:'Synthetic explicit label' for k in values},'issues':[],'unverified':{},'pdf_sha256':DIGEST,'pages':2}

def template(path,rows=()):
    book=openpyxl.Workbook()
    sheet=book.active
    sheet.title='Report'
    sheet.append([FIELD_HEADER_ALIASES[k][0] for k in FIELD_ORDER])
    sheet.cell(2,1).fill=PatternFill('solid',fgColor='ABCDEF')
    sheet['AF2']='=1+2'
    for n,record in enumerate(rows,2):
        for col,key in enumerate(FIELD_ORDER,1):
            sheet.cell(n,col).value=getattr(record,key)
    book.save(path)
    book.close()

class ExactParsing(unittest.TestCase):
    def test_exact_decimal_and_comma(self):
        self.assertEqual(number('84,759.40'),Decimal('84759.40'))
        self.assertEqual(number('980.5'),Decimal('980.5'))
        self.assertEqual(identifier(' 001-002 003 '),'001002003')

    def test_ambiguous_ocr_numbers_rejected(self):
        for value in ('1O0','l23','12,34','1,23.45','NaN','Infinity','-1','1e3'):
            with self.subTest(value=value):
                self.assertIsNone(number(value))
        self.assertIsNone(number('3.5',integer=True))

    def test_leading_zero_invoice_preserved(self):
        self.assertEqual(native_invoice([awb()],HOUSE),'00000123_'+HOUSE)
        self.assertEqual(native_invoice([awb(invoice='00000123_ '+HOUSE)],HOUSE),'00000123_'+HOUSE)

    def test_partial_invoice_not_restored(self):
        record,_=parse_pages([awb(invoice='00000123_PLIHQ2H0000')],HOUSE)
        self.assertIsNone(record.invoice_no)
        self.assertIsNone(native_invoice([awb(invoice='00000123_')],HOUSE))

    def test_foreign_invoice_house_is_blank(self):
        record,evidence=parse_pages([awb(invoice='00000123_PLIHQ2H00002')],HOUSE)
        self.assertIsNone(record.invoice_no)
        self.assertTrue(evidence['issues'])

    def test_filename_does_not_override_body(self):
        with self.assertRaises(ValueError):
            parse_pages([awb()],'PLIHQ2H99999')
        with self.assertRaises(ValueError):
            parse_pages([{'page':1,'kind':'native','native':'No transport identifier','boxes':[]}],HOUSE)
    def test_longer_house_serial_is_not_truncated(self):
        page=awb()
        page['native']=page['native'].replace(HOUSE,'PLIHQ2H000012')
        with self.assertRaises(ValueError):
            parse_pages([page],'PLIHQ2H000012')

    def test_chargeable_column_with_blank_rate_class(self):
        record,evidence=parse_pages([awb()],HOUSE)
        self.assertEqual(record.weight_quantity,88.5)
        self.assertEqual(evidence['proofs']['weight_quantity']['pages'],[1])

    def test_conflicting_copies_blank_weight(self):
        record,evidence=parse_pages([awb('88.5'),awb('89.5',2)],HOUSE)
        self.assertIsNone(record.weight_quantity)
        self.assertIn('원문 간 값 충돌',evidence['unverified']['weight_quantity'])

    def test_ocr_arithmetic_is_never_approval(self):
        page={'page':2,'kind':'commercial','native':'','boxes':[
            [[[0,0],[500,0],[500,20],[0,20]],'UNIT PRICE USD AMOUNT TOTAL 100 9.99 999',1.0]]}
        record,evidence=parse_pages([awb(),page],HOUSE)
        self.assertIsNone(record.invoice_value_usd)
        self.assertIsNone(record.item_quantity)
        self.assertIsNone(record.sku_count)
        self.assertFalse(collect_invoice_metrics([page])[0]['auto_approved'])

    def test_unstated_business_fields_blank(self):
        record,_=parse_pages([awb()],HOUSE)
        for key in ('requester','pickup_location','transport_type','pallet_count','ata','requested_delivery_date','incoterms'):
            self.assertIsNone(getattr(record,key),key)

class ProofPolicy(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.folder=Path(self.temp.name)
    def tearDown(self):
        self.temp.cleanup()
    def review(self,decision,house=HOUSE,digest=DIGEST):
        path=self.folder/'review.json'
        path.write_text(json.dumps({'schema':1,'reviewer':'Synthetic reviewer','review_basis':'visual_pdf',
            'documents':{digest:{'house_bl_no':house,'fields':{'invoice_value_usd':decision}}}}),encoding='utf-8')
        return ReviewedIndex(path)
    def test_unproved_values_removed_at_output_gate(self):
        record=ShipmentRecord(house_bl_no=HOUSE,invoice_value_usd=999)
        evidence=proved(record)
        evidence['proofs'].pop('invoice_value_usd')
        enforce_field_proofs(record,evidence)
        self.assertIsNone(record.invoice_value_usd)
        self.assertEqual(record.house_bl_no,HOUSE)
    def test_proof_for_other_value_cannot_be_reused(self):
        record=ShipmentRecord(invoice_value_usd=100)
        evidence=proved(record)
        record.invoice_value_usd=200
        enforce_field_proofs(record,evidence)
        self.assertIsNone(record.invoice_value_usd)
    def test_review_requires_exact_pdf_hash(self):
        index=self.review({'value':'123.45','pages':[2],'reason':'Directly inspected synthetic invoice'})
        record=ShipmentRecord(house_bl_no=HOUSE)
        evidence=proved(record)
        evidence['pdf_sha256']='b'*64
        index.apply(record,evidence)
        self.assertIsNone(record.invoice_value_usd)
        self.assertTrue(evidence['issues'])
    def test_review_applies_exact_cents_and_reason(self):
        index=self.review({'value':'123.45','pages':[2],'reason':'Directly inspected synthetic invoice'})
        record=ShipmentRecord(house_bl_no=HOUSE)
        evidence=proved(record)
        index.apply(record,evidence)
        enforce_field_proofs(record,evidence)
        self.assertEqual(record.invoice_value_usd,123.45)
        self.assertEqual(evidence['proofs']['invoice_value_usd']['method'],'visual_pdf')
    def test_explicit_blank_overrides_a_native_candidate(self):
        index=self.review({'value':None,'pages':[1,2],'reason':'Contradicting source copies'})
        record=ShipmentRecord(house_bl_no=HOUSE,invoice_value_usd=100)
        evidence=proved(record)
        index.apply(record,evidence)
        self.assertIsNone(record.invoice_value_usd)
        self.assertNotIn('invoice_value_usd',evidence['proofs'])
        self.assertEqual(evidence['withheld']['invoice_value_usd']['pages'],[1,2])
    def test_review_rejects_foreign_house(self):
        index=self.review({'value':100,'pages':[1],'reason':'Synthetic'},house='PLIHQ2H00002')
        record=ShipmentRecord(house_bl_no=HOUSE)
        with self.assertRaises(ValueError):
            index.apply(record,proved(record))
    def test_review_rejects_invalid_pages_and_numbers(self):
        for decision in ({'value':100,'pages':[3],'reason':'Synthetic'},
                         {'value':'1O0','pages':[1],'reason':'Synthetic'},
                         {'value':100,'pages':[1]}):
            with self.subTest(decision=decision):
                record=ShipmentRecord(house_bl_no=HOUSE)
                with self.assertRaises(ValueError):
                    self.review(decision).apply(record,proved(record))
    def test_excel_never_fills_or_overwrites(self):
        path=self.folder/'comparison.xlsx'
        template(path,[ShipmentRecord(house_bl_no=HOUSE,invoice_no='00000123_'+HOUSE,invoice_value_usd=999,item_quantity=100,pallet_count=7,ata=date(2026,1,2))])
        index=ReferenceIndex(path)
        record=ShipmentRecord(house_bl_no=HOUSE,invoice_no='00000123_'+HOUSE,invoice_value_usd=123.45)
        original=copy.deepcopy(record)
        evidence=proved(record)
        index.reconcile(record,evidence)
        self.assertEqual(record,original)
        self.assertEqual(evidence['reference']['filled'],[])
        self.assertTrue(evidence['reference']['conflicts'])
        self.assertFalse(evidence['issues'])

class WorkbookAndBatch(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.folder=Path(self.temp.name)
        self.template=self.folder/'template.xlsx'
        template(self.template)
    def tearDown(self):
        self.temp.cleanup()
    def item(self,record):
        return {'pdf':'synthetic.pdf','record':record,'evidence':proved(record),'status':'review'}
    def test_workbook_preserves_format_identifiers_cents_and_blanks(self):
        before=hashlib.sha256(self.template.read_bytes()).hexdigest()
        out=self.folder/'result.xlsx'
        rec=ShipmentRecord(house_bl_no=HOUSE,invoice_no='00000123',invoice_value_usd=123.45,weight_quantity=88.5,atd=date(2026,1,2))
        write_results(self.template,out,[self.item(rec)])
        book=openpyxl.load_workbook(out)
        sheet=book['Report']
        self.assertEqual(sheet.cell(2,FIELD_ORDER.index('invoice_no')+1).value,'00000123')
        self.assertEqual(sheet.cell(2,FIELD_ORDER.index('invoice_value_usd')+1).value,123.45)
        self.assertEqual(sheet.cell(2,FIELD_ORDER.index('weight_quantity')+1).value,88.5)
        self.assertIsNone(sheet.cell(2,FIELD_ORDER.index('ata')+1).value)
        self.assertEqual(sheet['AF2'].value,'=1+2')
        self.assertEqual(sheet.cell(2,1).fill.fgColor.rgb,'00ABCDEF')
        self.assertIn('원문근거',book.sheetnames)
        book.close()
        self.assertEqual(hashlib.sha256(self.template.read_bytes()).hexdigest(),before)
    def test_writer_itself_blocks_unproved_values(self):
        rec=ShipmentRecord(house_bl_no=HOUSE,invoice_value_usd=999)
        item=self.item(rec)
        item['evidence']['proofs'].pop('invoice_value_usd')
        out=self.folder/'result.xlsx'
        write_results(self.template,out,[item])
        book=openpyxl.load_workbook(out)
        self.assertIsNone(book['Report'].cell(2,FIELD_ORDER.index('invoice_value_usd')+1).value)
        book.close()
    def test_duplicate_house_not_appended_twice(self):
        rec=ShipmentRecord(house_bl_no=HOUSE)
        self.assertEqual(write_results(self.template,self.folder/'result.xlsx',[self.item(rec),self.item(copy.deepcopy(rec))]),1)
    def test_existing_file_never_overwritten(self):
        with self.assertRaises(ValueError):
            write_results(self.template,self.template,[])
    def test_different_pdf_bytes_same_house_neither_written(self):
        left=self.item(ShipmentRecord(house_bl_no=HOUSE,invoice_value_usd=100))
        right=self.item(ShipmentRecord(house_bl_no=HOUSE,invoice_value_usd=200))
        right['evidence']['pdf_sha256']='b'*64
        self.assertEqual(write_results(self.template,self.folder/'result.xlsx',[left,right]),0)
        self.assertEqual([left['status'],right['status']],['error','error'])
    def test_formula_like_pdf_text_is_text(self):
        out=self.folder/'result.xlsx'
        write_results(self.template,out,[self.item(ShipmentRecord(house_bl_no=HOUSE,shipper='=1+2'))])
        book=openpyxl.load_workbook(out)
        self.assertEqual(book['Report'].cell(2,FIELD_ORDER.index('shipper')+1).data_type,'s')
        book.close()
    def test_bad_pdf_logged_and_other_document_saved(self):
        pdfs=[self.folder/'bad.pdf',self.folder/'good.pdf']
        for pdf in pdfs:
            pdf.write_bytes(b'%PDF synthetic fixture')
        record=ShipmentRecord(house_bl_no=HOUSE,invoice_no='00000123')
        evidence=proved(record)
        digest=hashlib.sha256(pdfs[1].read_bytes()).hexdigest()
        evidence['pdf_sha256']=digest
        for proof in evidence['proofs'].values(): proof['pdf_sha256']=digest
        output=self.folder/'batch.xlsx'
        with patch('app.service.parse_pdf_record',side_effect=[ValueError('Corrupt PDF fixture'),(record,evidence)]):
            result=run_batch(pdfs,self.template,output)
        self.assertEqual(result['written'],1)
        self.assertEqual(result['counts']['error'],1)
        self.assertIn('Corrupt PDF fixture',Path(result['audit']).read_text(encoding='utf-8'))
        self.assertTrue(Path(result['review']).exists())

if __name__=='__main__':
    unittest.main()
