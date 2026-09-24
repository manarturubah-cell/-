from pathlib import Path


def test_items_api_exposes_export_status():
    app = Path(__file__).resolve().parents[1] / 'app.py'
    text = app.read_text(encoding='utf-8')
    assert 'li.export_status' in text


def test_items_ui_uses_export_status_for_export_badge():
    html = (Path(__file__).resolve().parents[1] / 'index.html').read_text(encoding='utf-8')
    assert 'function itemExportBadge(status)' in html
    assert "NEW:['جديد للتصدير','new']" in html
    assert "EXPORTED:['مصدّر — بانتظار تأكيد الاستيراد','review']" in html
    assert "IMPORTED:['مستورد','ok']" in html
    assert "UPDATE_PENDING:['تعديل للتصدير','review']" in html
    assert "x.source==='USER_CREATED'?' — <span class=\"badge new\">جديد للتصدير</span>':''" not in html
