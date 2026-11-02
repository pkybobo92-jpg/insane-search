"""Shipment data contract shared by document review and workbook output."""
from __future__ import annotations
from dataclasses import asdict,dataclass,field
from datetime import date,datetime

@dataclass
class ShipmentRecord:
    requester: str|None=None
    order_no: str|None=None
    album_name: str|None=None
    shipper: str|None=None
    consignee: str|None=None
    incoterms: str|None=None
    invoice_no: str|None=None
    house_bl_no: str|None=None
    vessel_voy: str|None=None
    transport_mode: str|None=None
    transport_type: str|None=None
    unit: str|None=None
    invoice_value_usd: float|int|str|None=None
    weight_quantity: float|int|str|None=None
    sku_count: int|str|None=None
    item_quantity: int|str|None=None
    pallet_count: int|str|None=None
    pickup_location: str|None=None
    pol: str|None=None
    pod: str|None=None
    final_destination: str|None=None
    order_received_date: date|datetime|str|None=None
    booking_notice_date: date|datetime|str|None=None
    outbound_date: date|datetime|str|None=None
    atd: date|datetime|str|None=None
    ata: date|datetime|str|None=None
    requested_delivery_date: date|datetime|str|None=None
    arrival_destination_date: date|datetime|str|None=None
    billing_completed_date: date|datetime|str|None=None
    source_notes: list[str]=field(default_factory=list)

    def to_serializable_dict(self):
        return {key:value.isoformat() if isinstance(value,(date,datetime)) else value
                for key,value in asdict(self).items()}

    def to_display_pairs(self):
        from .field_mapping import FIELD_ORDER,FIELD_HEADER_ALIASES
        return [(FIELD_HEADER_ALIASES[key][0],getattr(self,key)) for key in FIELD_ORDER]
