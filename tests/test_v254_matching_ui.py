from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
HTML = (ROOT / "index.html").read_text(encoding="utf-8")


def test_suggestion_confirmation_uses_suggestion_item_and_updates_local_item():
    assert "const suggestion=(x.match?.suggestions||[]).find" in HTML
    assert "const it=(suggestion?.item)||allItems.find" in HTML
    assert "status:'APPROVED'" in HTML
    assert "item:it" in HTML
    assert "سيظهر الصنف المختار الآن في خانة «الصنف لدينا»" in HTML


def test_change_item_restores_review_state_without_deleting_line():
    assert "function rememberReviewState(i)" in HTML
    assert "function resetLineToReview(i)" in HTML
    assert "x.match=cloneLineState(saved.match)" in HTML
    assert "delete x._learning" in HTML
    assert "🔄 تغيير الصنف" in HTML
    assert "deleteInvoiceLine(i)" in HTML  # deletion remains a separate action


def test_learning_is_replaced_when_choice_is_changed():
    assert "delete x._suggestion" in HTML
    assert "delete x._reviewState" in HTML
    assert "renderLearningBanner();" in HTML


def test_suggestion_button_does_not_embed_unescaped_json_string_in_double_quoted_onclick():
    # The method argument previously produced nested double quotes inside onclick,
    # so the browser rendered a dead button. The method is optional and defaults safely.
    assert "chooseSuggestion(${i},${Number(it.id)},${z.mapping_id==null?'null':Number(z.mapping_id)},${Number(z.score||0)})" in HTML
    assert "function chooseSuggestion(i,itemId,mappingId,score,method='SUGGESTION')" in HTML
