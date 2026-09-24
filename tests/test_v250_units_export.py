from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app

def test_simple_export_prefers_purchase_default_not_largest_factor():
    units=[
        {'id':1,'unit_name':'حبة','conversion_factor':1,'active':1},
        {'id':2,'unit_name':'كرتون','conversion_factor':10,'active':1},
        {'id':3,'unit_name':'كيس','conversion_factor':20,'active':1},
    ]
    chosen=app._preferred_purchase_unit(units,2,'حبة')
    assert chosen['unit_name']=='كرتون'

def test_simple_export_falls_back_to_main_unit_when_purchase_default_missing():
    units=[
        {'id':1,'unit_name':'جرام','conversion_factor':1,'active':1},
        {'id':2,'unit_name':'كرتون','conversion_factor':24,'active':1},
    ]
    chosen=app._preferred_purchase_unit(units,None,'جرام')
    assert chosen['unit_name']=='جرام'

def test_simple_export_does_not_guess_largest_factor():
    units=[
        {'id':1,'unit_name':'حبة','conversion_factor':1,'active':1},
        {'id':2,'unit_name':'كرتون','conversion_factor':24,'active':1},
        {'id':3,'unit_name':'بالة','conversion_factor':48,'active':1},
    ]
    chosen=app._preferred_purchase_unit(units,None,'حبة')
    assert chosen['unit_name']=='حبة'
