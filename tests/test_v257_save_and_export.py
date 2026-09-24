from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app


def setup_db(path, monkeypatch):
    monkeypatch.setattr(app, 'DB', path)
    app.init_item_management_schema()
    app.migrate_invoice_financial_fields()
    app.migrate_item_export_and_warehouse_fields()
    app.migrate_schema_v38()
    app.migrate_schema_v39()
    app.migrate_schema_v40()


def test_save_ignores_stale_suggestion_item_id_in_match_history(tmp_path, monkeypatch):
    setup_db(tmp_path / 'save.sqlite', monkeypatch)
    c = app.dbconn()
    c.execute("INSERT INTO suppliers(supplier_name,normalized_name,active) VALUES(?,?,1)", ('مورد اختبار', 'مورد اختبار'))
    sid = c.execute('SELECT last_insert_rowid()').fetchone()[0]
    c.commit(); c.close()
    item = app.create_local_item({'item_name':'صنف اختبار','group_code':'100','main_unit':'حبة','default_purchase_unit_name':'حبة','default_sale_unit_name':'حبة'})
    line = {
        'supplier_item_code': None, 'raw_item_name':'صنف اختبار المورد', 'qty':1, 'unit':'حبة', 'unit_price':5,
        'match': {'status':'APPROVED','score':90,'method':'USER','item':{'id':item['id']},'mapping_id':None},
        # This id can become stale in a browser session; it must never violate match_history FK.
        '_suggestion': {'item_id':99999999,'mapping_id':99999999,'score':88},
        'user_decision':'APPROVE',
    }
    out=app.save_invoice({'supplier_id':sid,'invoice_number':'FK-TEST-1','invoice_date':'2026-09-23','warehouse_number':'1','lines':[line]})
    assert out['lines_saved']==1
    c=app.dbconn()
    row=c.execute("SELECT suggested_item_id,suggested_mapping_id FROM match_history WHERE invoice_line_id=(SELECT id FROM invoice_lines WHERE invoice_id=?)",(out['invoice_id'],)).fetchone()
    assert row['suggested_item_id'] is None
    assert row['suggested_mapping_id'] is None
    assert c.execute('PRAGMA foreign_key_check').fetchall()==[]
    c.close()


def test_new_item_export_rows_use_sale_and_purchase_default_unit_ids():
    rows=[{
        'group_code':'100','item_code':'100001','item_name':'اختبار','item_type':'سلعي',
        'default_sale_unit_id':2,'default_purchase_unit_id':3,
        'units':[
            {'id':1,'unit_name':'حبة','conversion_factor':1},
            {'id':2,'unit_name':'كرتون','conversion_factor':10},
            {'id':3,'unit_name':'بالة','conversion_factor':20},
        ]
    }]
    out=app._item_export_rows(rows)
    assert out[0][6:8]==['','']
    assert out[1][6:8]==['1','']
    assert out[2][6:8]==['','1']


def test_confirm_selected_items_imported_only_changes_exported_items(tmp_path, monkeypatch):
    setup_db(tmp_path / 'confirm.sqlite', monkeypatch)
    a = app.create_local_item({'item_name':'صنف أ','group_code':'100','main_unit':'حبة','default_purchase_unit_name':'حبة','default_sale_unit_name':'حبة'})
    b = app.create_local_item({'item_name':'صنف ب','group_code':'100','main_unit':'حبة','default_purchase_unit_name':'حبة','default_sale_unit_name':'حبة'})
    c = app.dbconn()
    c.execute("UPDATE local_items SET export_status='EXPORTED' WHERE id=?", (a['id'],))
    c.execute("UPDATE local_items SET export_status='NEW' WHERE id=?", (b['id'],))
    c.commit(); c.close()
    assert app.confirm_selected_items_imported([a['id'], b['id']]) == 1
    c = app.dbconn()
    rows = {int(r['id']): r['export_status'] for r in c.execute('SELECT id,export_status FROM local_items WHERE id IN (?,?)',(a['id'],b['id']))}
    assert rows[a['id']] == 'IMPORTED'
    assert rows[b['id']] == 'NEW'
    assert c.execute('SELECT imported_confirmed_at FROM local_items WHERE id=?',(a['id'],)).fetchone()[0] is not None
    c.close()


def test_confirm_all_imported_only_changes_exported_items(tmp_path, monkeypatch):
    setup_db(tmp_path / 'confirm_all.sqlite', monkeypatch)
    a = app.create_local_item({'item_name':'صنف أ','group_code':'100','main_unit':'حبة','default_purchase_unit_name':'حبة','default_sale_unit_name':'حبة'})
    b = app.create_local_item({'item_name':'صنف ب','group_code':'100','main_unit':'حبة','default_purchase_unit_name':'حبة','default_sale_unit_name':'حبة'})
    c = app.dbconn()
    c.execute("UPDATE local_items SET export_status='EXPORTED' WHERE id IN (?,?)", (a['id'],b['id']))
    c.commit(); c.close()
    assert app.confirm_items_imported() == 2
    c = app.dbconn()
    assert c.execute("SELECT COUNT(*) FROM local_items WHERE export_status='EXPORTED'").fetchone()[0] == 0
    assert c.execute("SELECT COUNT(*) FROM local_items WHERE id IN (?,?) AND export_status='IMPORTED'",(a['id'],b['id'])).fetchone()[0] == 2
    c.close()


def test_export_invoice_xls_uses_frozen_purchase_unit(tmp_path, monkeypatch):
    setup_db(tmp_path / 'invoice_export.sqlite', monkeypatch)
    c = app.dbconn()
    c.execute("INSERT INTO suppliers(supplier_name,normalized_name,active) VALUES(?,?,1)", ('مورد تصدير', 'مورد تصدير'))
    sid = c.execute('SELECT last_insert_rowid()').fetchone()[0]
    c.commit(); c.close()
    item = app.create_local_item({'item_name':'صنف تصدير','group_code':'100','main_unit':'حبة','default_purchase_unit_name':'كرتون','default_sale_unit_name':'حبة','extra_units':[{'unit_name':'كرتون','conversion_factor':10}]})
    line = {
        'supplier_item_code':'EXP-1', 'raw_item_name':'صنف تصدير المورد', 'qty':2, 'unit':'CTN', 'unit_price':10,
        'match': {'status':'APPROVED','score':100,'method':'USER','item':{'id':item['id']},'mapping_id':None},
        'user_decision':'APPROVE',
    }
    out = app.save_invoice({'supplier_id':sid,'invoice_number':'EXP-TEST-1','invoice_date':'2026-09-23','warehouse_number':'1','lines':[line]})
    data, number = app.export_invoice_xls(out['invoice_id'])
    assert data[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'  # native BIFF8/OLE workbook
    assert number == 'EXP-TEST-1'
    c = app.dbconn()
    row = c.execute('SELECT export_unit_name FROM invoice_lines WHERE invoice_id=?',(out['invoice_id'],)).fetchone()
    assert row['export_unit_name'] == 'كرتون'
    c.close()


def test_invoice_detail_exposes_persisted_local_item_identity_and_display_fields(tmp_path, monkeypatch):
    setup_db(tmp_path / 'invoice_detail.sqlite', monkeypatch)
    c = app.dbconn()
    c.execute("INSERT INTO suppliers(supplier_name,normalized_name,active) VALUES(?,?,1)", ('مورد تفاصيل', 'مورد تفاصيل'))
    sid = c.execute('SELECT last_insert_rowid()').fetchone()[0]
    c.commit(); c.close()
    item = app.create_local_item({'item_name':'صنف محفوظ','group_code':'100','main_unit':'حبة','default_purchase_unit_name':'حبة','default_sale_unit_name':'حبة'})
    out = app.save_invoice({'supplier_id':sid,'invoice_number':'DETAIL-TEST-1','invoice_date':'2026-09-23','warehouse_number':'1','lines':[{
        'supplier_item_code':'D-1','raw_item_name':'صنف محفوظ عند المورد','qty':1,'unit':'حبة','unit_price':5,
        'match': {'status':'APPROVED','score':100,'method':'USER','item':{'id':item['id']},'mapping_id':None},
        'user_decision':'APPROVE'
    }]})
    detail=app.invoice_detail(out['invoice_id'])
    row=detail['lines'][0]
    assert int(row['local_item_id']) == int(item['id'])
    assert row['local_item_code'] == item['item_code']
    assert row['local_item_name'] == item['item_name']


def test_edit_saved_invoice_frontend_uses_persisted_item_id_as_source_of_truth():
    html=Path(__file__).resolve().parents[1].joinpath('index.html').read_text(encoding='utf-8')
    assert 'const savedItemId=Number(x.local_item_id||x.matched_local_item_id||0)' in html
    assert 'const fromCache=savedItemId?(allItems.find(z=>Number(z.id)===savedItemId)||null):null' in html
    assert "method:'SAVED'" in html


def test_backup_restore_has_explicit_pre_read_size_limit():
    src=(Path(app.__file__).resolve()).read_text(encoding='utf-8')
    marker='if p=="/api/backup/restore":'
    block=src.split(marker,1)[1].split('if p=="/api/items/update":',1)[0]
    assert 'max_restore=500*1024*1024' in block
    assert 'if length > max_restore:' in block
    assert 'body=self.rfile.read(length)' in block
    assert block.index('if length > max_restore:') < block.index('body=self.rfile.read(length)')


def test_top_navigation_tracks_active_tab():
    html=(Path(app.__file__).resolve().parent/'index.html').read_text(encoding='utf-8')
    assert 'data-tab="home" class="active"' in html
    assert "document.querySelectorAll('nav button[data-tab]').forEach(b=>b.classList.toggle('active',b.dataset.tab===id))" in html
