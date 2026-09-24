from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app


def test_mapping_supplier_name_is_used_for_suggestion(tmp_path, monkeypatch):
    db = tmp_path / 'v247_match.sqlite'
    monkeypatch.setattr(app, 'DB', db)
    app.init_item_management_schema()
    app.migrate_item_export_and_warehouse_fields()
    app.migrate_schema_v38()
    c = app.dbconn()
    c.execute("INSERT INTO suppliers(supplier_name,normalized_name) VALUES(?,?)", ('شركة ثمار الاختبار','شركه ثمار الاختبار'))
    sid = c.execute('SELECT last_insert_rowid()').fetchone()[0]
    c.execute("INSERT INTO local_items(item_code,item_name,normalized_name,main_unit,active,source,export_status) VALUES(?,?,?,?,1,'IMPORTED','IMPORTED')", ('100023','بطاطس هيفارم 2.5 كجم * 4','بطاطس هيفارم 2.5 كجم x 4','كرتون'))
    iid = c.execute('SELECT last_insert_rowid()').fetchone()[0]
    c.execute("INSERT INTO supplier_item_mappings(supplier_id,local_item_id,supplier_item_code,supplier_item_name,normalized_name,status,match_method,usage_count) VALUES(?,?,?,?,?,'APPROVED','SUPPLIER_CODE',1)", (sid,iid,None,'POTATO HYFARM 2.5KG X 4PKT','potato hyfarm 2.5kg x 4pkt'))
    c.commit(); c.close()
    out = app.match_lines_bulk([{'supplier_item_code':'1131','raw_item_name':'POTATO HYFARM 2.5KG X 4PKT (9/9)','unit':'كرتون'}], sid)[0]['match']
    assert out['status'] == 'SUGGESTED'
    assert int(out['item']['id']) == iid
    assert out['method'] in ('SUPPLIER_MAPPING_NAME','SUPPLIER_CODE_UNKNOWN_REVIEW')
    assert out['score'] >= 80


def test_rematch_saved_invoice_is_read_only_and_uses_latest_mapping(tmp_path, monkeypatch):
    db = tmp_path / 'v247_invoice.sqlite'
    monkeypatch.setattr(app, 'DB', db)
    app.init_item_management_schema()
    app.migrate_item_export_and_warehouse_fields()
    app.migrate_schema_v38()
    c = app.dbconn()
    c.execute("INSERT INTO suppliers(supplier_name,normalized_name) VALUES(?,?)", ('شركة ثمار الاختبار','شركه ثمار الاختبار'))
    sid = c.execute('SELECT last_insert_rowid()').fetchone()[0]
    c.execute("INSERT INTO local_items(item_code,item_name,normalized_name,main_unit,active,source,export_status) VALUES(?,?,?,?,1,'IMPORTED','IMPORTED')", ('100023','بطاطس هيفارم 2.5 كجم * 4','بطاطس هيفارم 2.5 كجم x 4','كرتون'))
    iid = c.execute('SELECT last_insert_rowid()').fetchone()[0]
    c.execute("INSERT INTO invoices(supplier_id,invoice_number,invoice_date,warehouse_number,processing_status) VALUES(?,?,?,?, 'REVIEW_REQUIRED')", (sid,'TH-1','2026-09-21','1'))
    inv = c.execute('SELECT last_insert_rowid()').fetchone()[0]
    c.execute("INSERT INTO invoice_lines(invoice_id,line_number,supplier_item_code,raw_item_name,normalized_item_name,supplier_unit,quantity,unit_price_halalah,match_status) VALUES(?,?,?,?,?,?,?,?,'UNMATCHED')", (inv,1,'1131','POTATO HYFARM 2.5KG X 4PKT','potato hyfarm 2.5kg x 4pkt','كرتون',10,5000))
    c.commit(); c.close()
    # Learn the mapping after the invoice was originally saved.
    c = app.dbconn()
    c.execute("INSERT INTO supplier_item_mappings(supplier_id,local_item_id,supplier_item_code,supplier_item_name,normalized_name,status,match_method,usage_count) VALUES(?,?,?,?,?,'APPROVED','SUPPLIER_CODE',1)", (sid,iid,'1131','POTATO HYFARM 2.5KG X 4PKT','potato hyfarm 2.5kg x 4pkt'))
    c.commit(); c.close()
    before = app.dbconn().execute("SELECT match_status,matched_local_item_id FROM invoice_lines WHERE invoice_id=?", (inv,)).fetchone()
    app.dbconn().close()
    result = app.rematch_invoice(inv)
    m = result['lines'][0]['match']
    assert m['status'] == 'AUTO_MATCHED'
    assert int(m['item']['id']) == iid
    c = app.dbconn()
    after = c.execute("SELECT match_status,matched_local_item_id FROM invoice_lines WHERE invoice_id=?", (inv,)).fetchone()
    c.close()
    assert tuple(after) == tuple(before)
