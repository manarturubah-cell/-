from pathlib import Path
import sqlite3
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app

HTML = Path(__file__).resolve().parents[1] / 'index.html'


def test_local_item_picker_has_search_and_safe_limit():
    s = HTML.read_text(encoding='utf-8')
    assert '🔎 بحث برقم أو اسم الصنف...' in s
    assert 'function filterLocalItems(i,q)' in s
    assert "arr=(String(q||'').trim()?arr.slice(0,150):arr.slice(0,80));" in s


def test_new_item_can_be_edited_from_invoice_screen():
    s = HTML.read_text(encoding='utf-8')
    assert 'async function openEditItemFromInvoice(id,i)' in s
    assert "fetch('/api/items/update'" in s
    assert '✏️ تعديل الصنف' in s
    assert 'يمكن تحديد الوحدة الأساسية ووحدة البيع ووحدة الشراء' in s


def test_update_item_records_history_and_barcode(tmp_path, monkeypatch):
    db = tmp_path / 'v246.sqlite'
    monkeypatch.setattr(app, 'DB', db)
    app.init_item_management_schema()
    app.migrate_item_export_and_warehouse_fields()
    app.migrate_schema_v39()
    item = app.create_local_item({
        'item_name':'صنف اختبار 246', 'group_code':'900', 'main_unit':'حبة',
        'conversion_factor':12, 'item_type':'سلعي', 'barcode':'111222'
    })
    out = app.update_item({
        'id':item['id'], 'item_name':'صنف اختبار 246 معدل', 'group_code':'901',
        'item_code':'901001', 'main_unit':'حبة', 'item_type':'سلعي',
        'pack_size':24, 'barcode':'333444'
    })
    assert out['item']['item_code'] == '901001'
    assert out['item']['group_code'] == '901'
    assert out['item']['item_name'] == 'صنف اختبار 246 معدل'
    assert out['history']
    assert out['history'][0]['action'] == 'UPDATE'
    assert 'item_name' in (out['history'][0]['changed_fields'] or '')
    c = app.dbconn()
    row = c.execute("SELECT barcode,conversion_factor FROM item_units WHERE local_item_id=? AND is_default=1", (item['id'],)).fetchone()
    carton = c.execute("SELECT conversion_factor FROM item_units WHERE local_item_id=? AND normalized_unit=?", (item['id'], app.norm('كرتون'))).fetchone()
    c.close()
    assert row['barcode'] == '333444'
    assert row['conversion_factor'] == 1
    assert carton is None
