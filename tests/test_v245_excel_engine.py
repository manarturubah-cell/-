import sqlite3
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app

def test_number_parser_consistent():
    assert app._excel_number("1,234") == 1234.0
    assert app._excel_number("1,23") == 1.23
    assert app._excel_number("1.234,56") == 1234.56

def test_row_kind_uses_quantity():
    m={"raw_item_name":2,"qty":3}
    assert app._excel_row_kind(["","",5],m)=="SUMMARY"

def test_content_mapping_requires_enough_values():
    assert app._analyze_column_content([[""],[1],[2],[3]],0,1) is None

def test_corrections_latest_wins(tmp_path, monkeypatch):
    db=tmp_path/'t.sqlite'
    conn=sqlite3.connect(db)
    conn.row_factory=sqlite3.Row
    conn.execute("CREATE TABLE extraction_corrections(id INTEGER PRIMARY KEY AUTOINCREMENT,supplier_id INTEGER,header_signature TEXT,field TEXT,original_column INTEGER,corrected_column INTEGER,learned_at TEXT)")
    conn.execute("INSERT INTO extraction_corrections VALUES(1,5,'sig','qty',3,4,'2026-01-01')")
    conn.execute("INSERT INTO extraction_corrections VALUES(2,5,'sig','qty',4,6,'2026-01-02')")
    conn.commit();conn.close()
    monkeypatch.setattr(app,'dbconn',lambda: (lambda c:(setattr(c,'row_factory',sqlite3.Row) or c))(sqlite3.connect(db)))
    out=app._apply_corrections_hint({'qty':3},5,'sig')
    assert out['qty']==6

def test_field_confidence_for_computed_values():
    rows=[["name","qty","price","total"],["A",2,5,None]]
    mapping={"raw_item_name":1,"qty":2,"unit_price":3,"gross_amount":4}
    out=app._excel_parse_with_mapping(rows,'S',1,mapping,{},{"raw_item_name":100,"qty":100,"unit_price":100,"gross_amount":100})
    assert out and out[0]['field_confidence']['qty']==100

def test_multiline_header_detection():
    rows=[['اسم الصنف','الكمية','سعر','الإجمالي'],['الوصف','المطلوبة','الوحدة','شامل الضريبة'],['A',2,10,23.0]]
    merged=app._merge_multiline_header(rows,0)
    assert merged and 'الكمية المطلوبة' in merged[1]

def test_csv_uses_excel_number_parser():
    data='اسم الصنف,الكمية,السعر\nA,"1,5","12,50"\n'.encode('utf-8')
    rows=app.parse_csv_bytes(data)
    assert rows[0]['qty']==1.5
    assert rows[0]['unit_price']==12.5
