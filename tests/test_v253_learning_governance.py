from pathlib import Path
import sqlite3
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import app


def setup_db(path, monkeypatch):
    monkeypatch.setattr(app, 'DB', path)
    app.init_item_management_schema()
    app.migrate_invoice_financial_fields()
    app.migrate_item_export_and_warehouse_fields()
    app.migrate_schema_v38()
    app.migrate_schema_v39()
    app.migrate_schema_v40()


def seed_supplier_and_item():
    c = app.dbconn()
    c.execute("INSERT INTO suppliers(supplier_name,normalized_name,active) VALUES(?,?,1)", ('شركة اختبار', 'شركه اختبار'))
    sid = c.execute('SELECT last_insert_rowid()').fetchone()[0]
    c.commit()
    c.close()
    item = app.create_local_item({
        'item_name': 'برجر امريكانا 4x24',
        'group_code': '100',
        'main_unit': 'كرتون',
        'default_purchase_unit_name': 'كرتون',
        'default_sale_unit_name': 'كرتون',
    })
    c.close()
    return int(sid), int(item['id'])


def test_pending_learning_is_not_persisted_until_invoice_save(tmp_path, monkeypatch):
    setup_db(tmp_path / 'learning.sqlite', monkeypatch)
    sid, iid = seed_supplier_and_item()
    line = {
        'supplier_id': sid,
        'supplier_item_code': '024097(1)',
        'raw_item_name': 'امريكانا برجر 4**24',
        'qty': 2,
        'unit': 'كرتون',
        'unit_price': 10,
        'match': {'status': 'APPROVED', 'score': 96, 'method': 'USER_CONFIRMED_SUGGESTION',
                  'item': {'id': iid, 'item_code': '100001', 'item_name': 'برجر امريكانا 4x24'}, 'mapping_id': None},
        '_learning': {'local_item_id': iid, 'learn_code': True, 'learn_alias': True, 'source': 'USER'},
    }
    c = app.dbconn()
    assert c.execute("SELECT COUNT(*) FROM supplier_item_mappings").fetchone()[0] == 0
    assert c.execute("SELECT COUNT(*) FROM supplier_item_aliases").fetchone()[0] == 0
    c.close()

    out = app.save_invoice({
        'supplier_id': sid,
        'invoice_number': 'T-1',
        'invoice_date': '2026-09-23',
        'warehouse_number': '1',
        'vat_rate': 15,
        'discount_halalah': 0,
        'lines': [line],
    })
    assert out['learned_code_rows'] == 1
    assert out['learned_alias_rows'] == 1
    c = app.dbconn()
    assert c.execute("SELECT COUNT(*) FROM supplier_item_mappings WHERE supplier_id=? AND local_item_id=?", (sid, iid)).fetchone()[0] == 1
    assert c.execute("SELECT COUNT(*) FROM supplier_item_aliases WHERE supplier_id=? AND local_item_id=? AND active=1", (sid, iid)).fetchone()[0] == 1
    c.close()


def test_failed_invoice_save_does_not_commit_learning(tmp_path, monkeypatch):
    setup_db(tmp_path / 'failed.sqlite', monkeypatch)
    sid, iid = seed_supplier_and_item()
    line = {
        'supplier_item_code': '9911', 'raw_item_name': 'اسم لن يتعلم', 'qty': 1, 'unit': 'كرتون', 'unit_price': 5,
        'match': {'status': 'APPROVED', 'item': {'id': iid}, 'mapping_id': None},
        '_learning': {'local_item_id': iid, 'learn_code': True, 'learn_alias': True, 'source': 'USER'},
    }
    original_learn_alias = app._learn_alias
    def fail_after_code(*args, **kwargs):
        raise RuntimeError('اختبار فشل بعد إنشاء كود المورد')
    monkeypatch.setattr(app, '_learn_alias', fail_after_code)
    try:
        app.save_invoice({'supplier_id': sid, 'invoice_number': 'ROLLBACK-1', 'invoice_date': '2026-09-23', 'warehouse_number': '1', 'lines': [line]})
    except RuntimeError:
        pass
    else:
        raise AssertionError('expected transactional failure')
    finally:
        monkeypatch.setattr(app, '_learn_alias', original_learn_alias)
    c = app.dbconn()
    assert c.execute("SELECT COUNT(*) FROM invoices WHERE invoice_number=?", ('ROLLBACK-1',)).fetchone()[0] == 0
    assert c.execute("SELECT COUNT(*) FROM supplier_item_mappings WHERE supplier_id=?", (sid,)).fetchone()[0] == 0
    assert c.execute("SELECT COUNT(*) FROM supplier_item_aliases WHERE supplier_id=?", (sid,)).fetchone()[0] == 0
    c.close()

def test_saved_invoice_edit_does_not_relearn_without_fresh_user_action(tmp_path, monkeypatch):
    setup_db(tmp_path / 'edit.sqlite', monkeypatch)
    sid, iid = seed_supplier_and_item()
    line = {
        'supplier_item_code': '7001', 'raw_item_name': 'برجر امريكانا 4x24', 'qty': 1, 'unit': 'كرتون', 'unit_price': 10,
        'match': {'status': 'APPROVED', 'method': 'MANUAL', 'item': {'id': iid}, 'mapping_id': None},
        '_learning': {'local_item_id': iid, 'learn_code': True, 'learn_alias': True, 'source': 'USER'},
    }
    app.save_invoice({'supplier_id': sid, 'invoice_number': 'EDIT-1', 'invoice_date': '2026-09-23', 'warehouse_number': '1', 'lines': [line]})
    c = app.dbconn()
    before = c.execute("SELECT usage_count FROM supplier_item_aliases WHERE supplier_id=? AND local_item_id=?", (sid, iid)).fetchone()[0]
    # The current product only permits edits of review invoices. Simulate a saved review invoice
    # here so the test isolates the learning-governance rule rather than invoice approval rules.
    c.execute("UPDATE invoices SET processing_status='REVIEW_REQUIRED' WHERE id=1")
    c.commit(); c.close()
    saved = app.invoice_detail(1)
    loaded = []
    for x in saved['lines']:
        loaded.append({
            'id': x['id'], 'supplier_item_code': x['supplier_item_code'], 'raw_item_name': x['raw_item_name'],
            'qty': x['quantity'], 'unit': x['supplier_unit'], 'unit_price': x['unit_price_halalah'] / 100,
            'match': {'status': x['match_status'], 'item': {'id': x['matched_local_item_id']}, 'mapping_id': x['matched_mapping_id']},
            'user_decision': x['user_decision'],
        })
    app.update_invoice(1, {'supplier_id': sid, 'invoice_number': 'EDIT-1', 'invoice_date': '2026-09-23', 'warehouse_number': '1', 'lines': loaded})
    c = app.dbconn()
    after = c.execute("SELECT usage_count FROM supplier_item_aliases WHERE supplier_id=? AND local_item_id=?", (sid, iid)).fetchone()[0]
    c.close()
    assert after == before


def test_management_can_disable_and_reenable_alias(tmp_path, monkeypatch):
    setup_db(tmp_path / 'manage.sqlite', monkeypatch)
    sid, iid = seed_supplier_and_item()
    c = app.dbconn()
    aid = app._learn_alias(c, sid, iid, 'برجر امريكانا', source='USER')[0]
    c.commit(); c.close()
    app.toggle_learning_alias(aid, False)
    c = app.dbconn()
    assert c.execute('SELECT active FROM supplier_item_aliases WHERE id=?', (aid,)).fetchone()[0] == 0
    c.close()
    app.toggle_learning_alias(aid, True)
    c = app.dbconn()
    assert c.execute('SELECT active FROM supplier_item_aliases WHERE id=?', (aid,)).fetchone()[0] == 1
    c.close()


def test_learning_history_is_written_for_final_user_learning(tmp_path, monkeypatch):
    setup_db(tmp_path / 'history.sqlite', monkeypatch)
    sid, iid = seed_supplier_and_item()
    line = {
        'supplier_item_code': '8008', 'raw_item_name': 'برجر امريكانا', 'qty': 1, 'unit': 'كرتون', 'unit_price': 9,
        'match': {'status': 'APPROVED', 'method': 'USER_CONFIRMED_SUGGESTION', 'score': 94, 'item': {'id': iid}, 'mapping_id': None},
        '_learning': {'local_item_id': iid, 'learn_code': True, 'learn_alias': True, 'source': 'USER', 'reason': 'SUGGESTION_CONFIRMED'},
    }
    app.save_invoice({'supplier_id': sid, 'invoice_number': 'H-1', 'invoice_date': '2026-09-23', 'warehouse_number': '1', 'lines': [line]})
    c = app.dbconn()
    row = c.execute("SELECT action,source,invoice_id,invoice_line_id FROM supplier_learning_history ORDER BY id DESC LIMIT 1").fetchone()
    c.close()
    assert row['action'] == 'LEARN_FROM_INVOICE'
    assert row['source'] == 'USER'
    assert row['invoice_id'] is not None and row['invoice_line_id'] is not None
