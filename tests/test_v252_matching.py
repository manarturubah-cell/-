from pathlib import Path
import sys

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
import app


def test_supplier_code_sequence_suffix_is_only_cleaned_when_it_matches_row_sequence():
    assert app._code_keys("024097(1)", 1) == [("FULL", "240971"), ("SEQUENCE_SUFFIX", "24097")]
    assert app._code_keys("024097(1)", 2) == [("FULL", "240971")]
    assert app._code_keys("AB(12)", 1) == [("FULL", "ab12")]


def test_name_features_make_word_order_and_pack_format_comparable():
    a=app._name_features("برجر امريكانا 4*24")
    b=app._name_features("امريكانا برجر 4**24")
    assert a[1] == b[1]
    assert a[2] == b[2] == {"4","24"}


def test_candidate_score_rewards_numeric_and_word_agreement():
    line={"raw_item_name":"امريكانا برجر 4**24","unit":"كرتون","unit_price":100}
    item={"item_name":"برجر امريكانا 4x24","main_unit":"كرتون"}
    score,evidence=app._candidate_score(line,item,'LOCAL_NAME',[])
    assert score >= 80
    assert evidence["numbers"] == 100.0
    assert evidence["unit"] is True


def test_excel_header_detection_includes_line_headers_and_small_samples():
    assert app._looks_like_header("سعر الوحدة") is True
    assert app._looks_like_header("كود الصنف") is True
    rows=[["اسم الصنف","الكمية","سعر الوحدة"],["أ","1","10"],["ب","2","11"],["ج","3","12"],["د","4","13"],["هـ","5","14"]]
    assert app._analyze_column_content(rows,1,2) == "qty_candidate"

def test_header_alternatives_keep_unselected_candidates():
    rows=[["كود الصنف","اسم الصنف","الكمية","Qty","سعر الوحدة"]]
    result=app._excel_classify_header_row(rows)
    assert result is not None
    alternatives=result[6]
    assert "qty" in alternatives and len(alternatives["qty"]) >= 1


def test_match_uses_supplier_code_after_sequence_suffix_cleanup():
    import sqlite3
    db=sqlite3.connect(':memory:')
    db.row_factory=sqlite3.Row
    db.execute('CREATE TABLE dummy(x INTEGER)')
    item={'id':1,'item_code':'900001','item_name':'برجر امريكانا 4x24','main_unit':'كرتون','active':1}
    mapping={'mapping_id':7,'supplier_id':3,'local_item_id':1,'supplier_item_name':'برجر امريكانا 4x24','supplier_item_code':'024097','normalized_name':'برجر امريكانا 4x24','supplier_unit':'كرتون','normalized_unit':'كرتون','usage_count':2,'confidence_score':100,'status':'APPROVED',**item}
    line={'supplier_id':3,'supplier_item_code':'024097(1)','raw_item_name':'امريكانا برجر 4**24','unit':'كرتون','unit_price':100}
    result=app._match_one(db,line,context={'all_items':[item],'mappings':[mapping],'aliases':[],'prices':{}},sequence=1)
    assert result['status']=='AUTO_MATCHED'
    assert result['method']=='SUPPLIER_CODE_SEQUENCE_CLEANED'
    assert result['item']['id']==1
    db.close()


def test_match_returns_multiple_ranked_suggestions_without_forcing_match():
    import sqlite3
    db=sqlite3.connect(':memory:')
    db.row_factory=sqlite3.Row
    items=[
        {'id':1,'item_code':'1001','item_name':'برجر امريكانا 4x24','main_unit':'كرتون','active':1},
        {'id':2,'item_code':'1002','item_name':'برجر امريكانا 2x24','main_unit':'كرتون','active':1},
        {'id':3,'item_code':'1003','item_name':'دجاج امريكانا 4x24','main_unit':'كرتون','active':1},
    ]
    line={'supplier_id':9,'supplier_item_code':'','raw_item_name':'امريكانا برجر 4**24','unit':'كرتون','unit_price':100}
    result=app._match_one(db,line,context={'all_items':items,'mappings':[],'aliases':[],'prices':{}},sequence=1)
    assert result['status']=='SUGGESTED'
    assert len(result['suggestions'])>=2
    assert result['suggestions'][0]['item']['id']==1
    db.close()
