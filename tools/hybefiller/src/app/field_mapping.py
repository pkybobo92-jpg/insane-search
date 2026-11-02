"""Report field names and aliases retained from the supplied release."""
import re

def normalize_text(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()

def normalized_header(value):
    return normalize_text(value).replace(" ", "").lower()

FIELD_ORDER = ['requester', 'order_no', 'album_name', 'shipper', 'consignee', 'incoterms', 'invoice_no', 'house_bl_no', 'vessel_voy', 'transport_mode', 'transport_type', 'unit', 'invoice_value_usd', 'weight_quantity', 'sku_count', 'item_quantity', 'pallet_count', 'pickup_location', 'pol', 'pod', 'final_destination', 'order_received_date', 'booking_notice_date', 'outbound_date', 'atd', 'ata', 'requested_delivery_date', 'arrival_destination_date', 'billing_completed_date']

FIELD_HEADER_ALIASES = {'requester': ['오더요청자', 'Requester'], 'order_no': ['Order No.', 'Order No', 'PO#', 'PO No.', 'PO Number'], 'album_name': ['앨범명', 'Album Name', 'Goods Description'], 'shipper': ['Shipper'], 'consignee': ['Consignee'], 'incoterms': ['Incoterms'], 'invoice_no': ['Invoice No.', 'Invoice No', 'Inv. No.'], 'house_bl_no': ['House BL No.', 'House BL No', 'HAWB No.', 'HAWB No', 'B/L No.', 'B/L No'], 'vessel_voy': ['VESSEL/VOY', 'VESSEL/FLIGHT', 'Flight No.', 'Flight No'], 'transport_mode': ['운송수단', 'Transport Mode', 'Ship Mode'], 'transport_type': ['운송 형태', 'Transport Type'], 'unit': ['진행단위', 'Unit'], 'invoice_value_usd': ['Invoice Value(USD)', 'Invoice Value USD'], 'weight_quantity': ['물량 (kg 또는 대수)', '물량\n(kg 또는 대수)', '물량 \n(kg 또는 대수)', 'Weight', 'GWT / CWT / CBM / QTY'], 'sku_count': ['SKU 개수', 'SKU Count'], 'item_quantity': ['개별 수량', 'Quantity'], 'pallet_count': ['Pallet 수량', 'PKG', 'Package Count'], 'pickup_location': ['픽업지', 'Pickup Location'], 'pol': ['POL', 'DPTR'], 'pod': ['POD', 'DEST'], 'final_destination': ['최종도착지', 'Final Destination', 'F.DEST'], 'order_received_date': ['Order 접수일', 'Order\n접수일'], 'booking_notice_date': ['운송 예약 안내일', '운송 예약\n안내일'], 'outbound_date': ['출고일'], 'atd': ['ATD', 'Onboard Date'], 'ata': ['ATA', 'Arrival Date'], 'requested_delivery_date': ['최종 도착지 납기요청일', '최종 도착지\n납기요청일'], 'arrival_destination_date': ['최종 도착지 도착일', '최종 도착지\n도착일'], 'billing_completed_date': ['비용 청구 완료일', '비용 청구\n완료일']}

NORMALIZED_FIELD_HEADER_ALIASES = {
    key:{normalized_header(value) for value in values}
    for key,values in FIELD_HEADER_ALIASES.items()
}

def find_header_row(worksheet,max_scan_rows=20):
    for row_no,row in enumerate(worksheet.iter_rows(min_row=1,max_row=min(max_scan_rows,worksheet.max_row)),1):
        mapping={}
        for cell in row:
            text=normalized_header(cell.value)
            for field,aliases in NORMALIZED_FIELD_HEADER_ALIASES.items():
                if text and text in aliases:
                    if field in mapping:
                        raise ValueError('중복 헤더: '+field)
                    mapping[field]=cell.column
        if 'house_bl_no' in mapping and 'invoice_no' in mapping:
            return row_no,mapping
    return None
