from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
HTML = (ROOT / 'index.html').read_text(encoding='utf-8')


def test_busy_and_toast_helpers_exist_before_match_use():
    assert 'function busy(' in HTML
    assert 'function toast(' in HTML
    assert 'async function runMatching(' in HTML


def test_match_routes_through_run_matching():
    assert 'async function match(){return runMatching(false)}' in HTML


def test_auto_matching_after_upload_is_enabled():
    assert 'await runMatching(true)' in HTML
    assert 'detected&&$(\'supplierSelect\').value' in HTML


def test_version_is_consistent_in_ui():
    version = (ROOT / 'VERSION.txt').read_text(encoding='utf-8').strip()
    assert version == 'v2.60'
    assert '>v2.60</div>' in HTML
