"""Regression cases for data corruption boundaries, using synthetic documents only."""
import copy
import hashlib
import json
import tempfile
import unittest
from pathlib import Path
from decimal import Decimal
from unittest.mock import patch
from test_accuracy import HOUSE, DIGEST, awb, proved, template, ShipmentRecord
from app.document_reader import read_pdf, CACHE_VERSION, pdf_snapshot
from app.pdf_parser import _party, _airport
from app.pdf_parser import parse_pages
from app.provenance import reject_conflicting_houses
from app.provenance import enforce_field_proofs
from app.reviewed import ReviewedIndex
from app.service import run_batch
from app.excel_writer import write_results


def pdf_bytes(text):
    # A genuine one-page, native-text PDF; no OCR or external document required.
    stream=('BT /F1 12 Tf 40 750 Td ('+text+') Tj ET').encode('ascii')
    objects=[b'<< /Type /Catalog /Pages 2 0 R >>',
             b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
             b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 800] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
             b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
             b'<< /Length '+str(len(stream)).encode()+b' >>\nstream\n'+stream+b'\nendstream']
    data=b'%PDF-1.4\n'; offsets=[0]
    for n,obj in enumerate(objects,1):
        offsets.append(len(data)); data+=f'{n} 0 obj\n'.encode()+obj+b'\nendobj\n'
    start=len(data)
    data+=b'xref\n0 6\n0000000000 65535 f \n'
    data+=b''.join(f'{offset:010d} 00000 n \n'.encode() for offset in offsets[1:])
    return data+f'trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n{start}\n%%EOF\n'.encode()


class ReviewBoundaries(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.folder=Path(self.temp.name)
        self.addCleanup(self.temp.cleanup)
    def index(self,fields):
        path=self.folder/'decisions.json'
        path.write_text(json.dumps({'schema':1,'reviewer':'Synthetic reviewer','review_basis':'visual_pdf',
            'documents':{DIGEST:{'house_bl_no':HOUSE,'fields':fields}}}),encoding='utf-8')
        return ReviewedIndex(path)

    def test_house_cannot_be_overridden_by_review(self):
        index=self.index({'house_bl_no':{'value':'PLIHQ2H99999','pages':[1],'reason':'Synthetic'}})
        record=ShipmentRecord(house_bl_no=HOUSE)
        with self.assertRaises(ValueError): index.apply(record,proved(record))
        self.assertEqual(record.house_bl_no,HOUSE)

    def test_review_for_foreign_invoice_is_rejected(self):
        index=self.index({'invoice_no':{'value':'00000123_PLIHQ2H99999','pages':[1],'reason':'Synthetic'}})
        record=ShipmentRecord(house_bl_no=HOUSE)
        with self.assertRaises(ValueError): index.apply(record,proved(record))

    def test_invalid_later_decision_does_not_partially_apply(self):
        index=self.index({'invoice_value_usd':{'value':'999','pages':[1],'reason':'Synthetic'},
                          'sku_count':{'value':'1O','pages':[1],'reason':'Synthetic'}})
        record=ShipmentRecord(house_bl_no=HOUSE); evidence=proved(record)
        before=copy.deepcopy(evidence)
        with self.assertRaises(ValueError): index.apply(record,evidence)
        self.assertIsNone(record.invoice_value_usd)
        self.assertEqual(evidence,before)

    def test_missing_value_is_not_an_explicit_blank(self):
        index=self.index({'invoice_value_usd':{'pages':[1],'reason':'Synthetic'}})
        record=ShipmentRecord(house_bl_no=HOUSE)
        with self.assertRaises(ValueError): index.apply(record,proved(record))

    def test_duplicate_json_keys_are_rejected(self):
        path=self.folder/'duplicate.json'
        path.write_text('{"schema":1,"schema":1,"reviewer":"x","review_basis":"visual_pdf","documents":{}}',encoding='utf-8')
        with self.assertRaises(ValueError): ReviewedIndex(path)

    def test_review_precision_is_not_silently_rounded(self):
        index=self.index({'invoice_value_usd':{'value':'1234567890123.456','pages':[1],'reason':'Synthetic'}})
        record=ShipmentRecord(house_bl_no=HOUSE)
        with self.assertRaises(ValueError): index.apply(record,proved(record))

    def test_proof_from_different_pdf_is_removed(self):
        record=ShipmentRecord(house_bl_no=HOUSE,invoice_value_usd=100)
        evidence=proved(record)
        evidence['proofs']['invoice_value_usd'].update(method='visual_pdf',pdf_sha256='b'*64)
        enforce_field_proofs(record,evidence)
        self.assertIsNone(record.invoice_value_usd)

    def test_proved_invalid_numeric_values_are_still_blank(self):
        for value in (True,-1,0,float('inf'),1.5,'1234567890123456'):
            with self.subTest(value=value):
                record=ShipmentRecord(house_bl_no=HOUSE,item_quantity=value)
                evidence=proved(record)
                enforce_field_proofs(record,evidence)
                self.assertIsNone(record.item_quantity)

    def test_party_crossing_column_boundary_is_not_truncated(self):
        page=awb(); page['words']=[
            {'text':"Shipper's Name and Address",'x0':10,'x1':180,'top':10,'bottom':20},
            {'text':'SYNTHETIC','x0':10,'x1':100,'top':30,'bottom':40},
            {'text':'COMPANY','x0':270,'x1':310,'top':30,'bottom':40}]
        self.assertIsNone(_party(page,"Shipper's Name and Address")[0])

    def test_invoice_mention_cannot_select_another_house_awb(self):
        page=awb(invoice='00000123_'+HOUSE)
        page['native']=page['native'].replace(HOUSE,'PLIHQ2H99999',1)
        with self.assertRaises(ValueError): parse_pages([page],HOUSE)

    def test_two_house_awbs_in_one_pdf_are_not_selected_by_filename(self):
        other=awb(page=2)
        other['native']=other['native'].replace(HOUSE,'PLIHQ2H99999')
        with self.assertRaises(ValueError): parse_pages([awb(),other],HOUSE)

    def test_other_document_incoterms_not_attached_to_selected_house(self):
        other={'page':2,'kind':'native','native':'PLIHQ2H99999\nCOMMERCIAL INVOICE\nINCOTERMS: FOB','boxes':[]}
        record,_=parse_pages([awb(),other],HOUSE)
        self.assertIsNone(record.incoterms)

    def test_incoterms_from_explicitly_linked_native_document_is_kept(self):
        linked={'page':2,'kind':'native','native':'B/L No. '+HOUSE+'\nINCOTERMS: FCA','boxes':[]}
        record,_=parse_pages([awb(),linked],HOUSE)
        self.assertEqual(record.incoterms,'FCA')

    def test_conflicting_invoice_labels_on_same_page_are_blank(self):
        page=awb(); page['native']+='INVOICE NO: 00000999_'+HOUSE+'\n'
        record,_=parse_pages([page],HOUSE)
        self.assertIsNone(record.invoice_no)

    def test_adjacent_column_date_does_not_erase_explicit_airport(self):
        page=awb(); page['words']=[
            {'text':'Airport of Destination','x0':40,'x1':140,'top':310,'bottom':320},
            {'text':'FRANKFURT','x0':40,'x1':130,'top':330,'bottom':340},
            {'text':'OCT/01/2026','x0':235,'x1':280,'top':330,'bottom':340}]
        self.assertEqual(_airport(page,'Airport of Destination')[0],'DEFRA')


class FileBoundaries(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.folder=Path(self.temp.name)
        self.addCleanup(self.temp.cleanup)
        self.template=self.folder/'template.xlsx'; template(self.template)

    def item(self,record):
        return {'pdf':'synthetic.pdf','record':record,'evidence':proved(record),'status':'review','error':None}
    def test_populated_template_is_rejected(self):
        template(self.template,[ShipmentRecord(house_bl_no='PLIHQ2H99999',invoice_value_usd=999)])
        with self.assertRaises(ValueError):
            write_results(self.template,self.folder/'old_data.xlsx',[self.item(ShipmentRecord(house_bl_no=HOUSE))])

    def test_duplicate_option_cannot_allow_conflicting_sources(self):
        left=self.item(ShipmentRecord(house_bl_no=HOUSE))
        right=copy.deepcopy(left); right['evidence']['pdf_sha256']='b'*64
        reject_conflicting_houses([left,right],allow_duplicates=True)
        self.assertEqual([left['status'],right['status']],['error','error'])

    def test_known_failed_revision_blocks_other_same_house_document(self):
        left=self.item(ShipmentRecord(house_bl_no=HOUSE))
        right=copy.deepcopy(left); right['evidence']['pdf_sha256']='b'*64
        right.update(status='error',error='Synthetic failed revision')
        reject_conflicting_houses([left,right])
        self.assertEqual([left['status'],right['status']],['error','error'])

    def test_numeric_format_preserves_extra_decimal_places(self):
        import openpyxl
        from app.field_mapping import FIELD_ORDER
        output=self.folder/'precise.xlsx'
        write_results(self.template,output,[self.item(ShipmentRecord(house_bl_no=HOUSE,
                      invoice_value_usd=123.456,weight_quantity=88.125))])
        book=openpyxl.load_workbook(output)
        try:
            for field,places in [('invoice_value_usd',3),('weight_quantity',3)]:
                cell=book['Report'].cell(2,FIELD_ORDER.index(field)+1)
                self.assertGreaterEqual(len(cell.number_format.split('.')[1]),places)
        finally: book.close()

    def test_changed_input_is_not_written(self):
        path=self.folder/'changing.pdf'; path.write_bytes(b'%PDF original synthetic')
        record=ShipmentRecord(house_bl_no=HOUSE); evidence=proved(record)
        evidence['pdf_sha256']=hashlib.sha256(path.read_bytes()).hexdigest()
        for proof in evidence['proofs'].values(): proof['pdf_sha256']=evidence['pdf_sha256']
        def parse(*args,**kwargs):
            path.write_bytes(b'%PDF changed synthetic')
            return record,evidence
        with patch('app.service.parse_pdf_record',side_effect=parse):
            result=run_batch([path],self.template,self.folder/'changed.xlsx')
        self.assertEqual(result['written'],0)
        self.assertEqual(result['counts']['error'],1)

    def test_cache_cannot_supply_native_values(self):
        path=self.folder/'native.pdf'
        text='SYNTHETIC ORIGINAL NATIVE DOCUMENT '+HOUSE+' Air Waybill '+('Explicit native source text '*5)
        path.write_bytes(pdf_bytes(text)); digest=hashlib.sha256(path.read_bytes()).hexdigest()
        cache=self.folder/'cache'; cache.mkdir()
        (cache/(digest+'.json')).write_text(json.dumps({'version':CACHE_VERSION,'sha256':digest,
            'pages':[{'page':1,'kind':'native','native':'FORGED WRONG NATIVE VALUE '+HOUSE,
                      'boxes':[],'words':[],'size':[600,800]}]}),encoding='utf-8')
        pages,actual=read_pdf(path,cache_dir=cache)
        self.assertEqual(actual,digest)
        self.assertIn('SYNTHETIC ORIGINAL NATIVE DOCUMENT',pages[0]['native'])
        self.assertNotIn('FORGED',pages[0]['native'])

    def test_snapshot_stays_exact_and_is_removed_on_changed_source(self):
        path=self.folder/'snapshot.pdf'; path.write_bytes(b'%PDF first version')
        snapshot=None
        with self.assertRaises(ValueError):
            with pdf_snapshot(path) as (snapshot,digest):
                path.write_bytes(b'%PDF second version')
                self.assertEqual(snapshot.read_bytes(),b'%PDF first version')
                self.assertEqual(digest,hashlib.sha256(snapshot.read_bytes()).hexdigest())
        self.assertFalse(snapshot.exists())

    def test_saved_value_corruption_blocks_publish_and_removes_temporary(self):
        import openpyxl
        from app.field_mapping import FIELD_ORDER
        original_save=openpyxl.Workbook.save
        def altered_save(book,path):
            original_save(book,path)
            corrupted=openpyxl.load_workbook(path)
            corrupted['Report'].cell(2,FIELD_ORDER.index('invoice_value_usd')+1).value=999
            original_save(corrupted,path); corrupted.close()
        output=self.folder/'corrupt.xlsx'
        with patch.object(openpyxl.Workbook,'save',altered_save):
            with self.assertRaises(ValueError):
                write_results(self.template,output,[self.item(ShipmentRecord(house_bl_no=HOUSE,invoice_value_usd=123))])
        self.assertFalse(output.exists())
        self.assertEqual(list(self.folder.glob('hybe-*.xlsx')),[])

    def test_output_created_during_save_is_never_overwritten(self):
        output=self.folder/'raced.xlsx'
        with self.assertRaises(FileExistsError):
            write_results(self.template,output,[self.item(ShipmentRecord(house_bl_no=HOUSE))],
                pre_publish=lambda:output.write_bytes(b'external existing result'))
        self.assertEqual(output.read_bytes(),b'external existing result')
        self.assertEqual(list(self.folder.glob('hybe-*.xlsx')),[])

    def test_template_change_during_batch_aborts_publication(self):
        import openpyxl
        path=self.folder/'synthetic.pdf'; path.write_bytes(b'%PDF synthetic unchanged')
        record=ShipmentRecord(house_bl_no=HOUSE); evidence=proved(record)
        evidence['pdf_sha256']=hashlib.sha256(path.read_bytes()).hexdigest()
        for proof in evidence['proofs'].values(): proof['pdf_sha256']=evidence['pdf_sha256']
        def parse(*args,**kwargs):
            book=openpyxl.load_workbook(self.template); book['Report']['AF2']='Changed context'
            book.save(self.template); book.close()
            return record,evidence
        output=self.folder/'template_changed.xlsx'
        with patch('app.service.parse_pdf_record',side_effect=parse):
            with self.assertRaises(ValueError): run_batch([path],self.template,output)
        self.assertFalse(output.exists())
        self.assertIn('save_failed',output.with_suffix('.audit.jsonl').read_text(encoding='utf-8'))
