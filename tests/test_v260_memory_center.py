from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / "app.py").read_text(encoding="utf-8")
HTML = (ROOT / "index.html").read_text(encoding="utf-8")

def test_learning_history_route_parses_its_own_query_string():
    block = APP[APP.index('if p=="/api/learning/history"'):APP.index('if p=="/api/learning/conflicts"')]
    assert "q=urlparse(self.path).query" in block
    assert "learning_history(limit,sid or None,iid or None,entity_type,action,source)" in block

def test_learning_history_supports_auditable_filters_and_invoice_number():
    assert "entity_type=None, action=None, source=None" in APP
    assert "i.invoice_number" in APP
    assert "entity_label" in APP
    assert "before_json" in HTML and "after_json" in HTML

def test_learning_management_exposes_usage_and_last_use():
    assert "m.usage_count" in APP
    assert "m.last_seen_at" in APP
    assert "x.usage_count" in HTML
    assert "x.last_seen_at" in HTML

def test_learning_management_uses_correction_language_and_soft_disable():
    assert "✏️ تصحيح" in HTML
    assert "تعطيل" in HTML
    assert "إعادة تفعيل" in HTML

def test_learning_history_has_filters():
    assert "lhSupplier" in HTML
    assert "lhType" in HTML
    assert "lhAction" in HTML
    assert "lhSource" in HTML
