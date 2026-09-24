from pathlib import Path

HTML = Path(__file__).resolve().parents[1] / 'index.html'

def test_new_item_button_for_any_unresolved_line_without_local_item():
    s = HTML.read_text(encoding='utf-8')
    assert "const needsNewItem=!selected && !['AUTO_MATCHED','APPROVED'].includes(upperStatus);" in s
    assert '➕ إنشاء صنف جديد' in s


def test_v245_version():
    v = (HTML.parent / 'VERSION.txt').read_text(encoding='utf-8')
    assert v.startswith('v2.60')
