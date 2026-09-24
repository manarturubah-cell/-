import io
from openpyxl import Workbook
import app

def test_template_signature_is_required():
    wb=Workbook(); ws=wb.active
    ws.append(['Code','Description','Qty','Unit','Price'])
    ws.append(['1','Test item',2,'CTN',10])
    data=io.BytesIO(); wb.save(data)
    lines=app.parse_xlsx_bytes(data.getvalue(), template={'header_signature':'wrong','column_map':{'raw_item_name':5,'qty':1}})
    assert lines[0]['raw_item_name']=='Test item'
    assert lines[0]['qty']==2

def test_number_rules_are_stable():
    assert app._excel_number('1,234') == 1234
    assert app._excel_number('1,23') == 1.23
    assert app._excel_number('1.234,56') == 1234.56

def test_header_label_does_not_take_next_label_as_value():
    rows=[['اسم المورد','رقم الفاتورة','12345'],['اسم الصنف','الكمية','الوحدة','السعر'],['x',1,'CTN',10]]
    meta=app._excel_extract_invoice_meta(rows,2)
    assert meta['supplier_name']==''
    assert meta['invoice_number']=='12345'
