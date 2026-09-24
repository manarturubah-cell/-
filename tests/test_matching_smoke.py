"""Isolated smoke tests for supplier-code matching and approval.

These tests create a temporary copy of the application and database. They never
open or modify the user's operational database.
"""
from __future__ import annotations
import json, os, shutil, sqlite3, subprocess, sys, tempfile, time, urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PORT = 18767
BASE_URL = f"http://127.0.0.1:{PORT}"


def request_json(path, payload):
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(BASE_URL + path, data=data, headers={"Content-Type":"application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=10) as r:
        return json.loads(r.read().decode("utf-8"))


def main():
    with tempfile.TemporaryDirectory(prefix="mutakamil_match_test_") as td:
        work = Path(td) / ROOT.name
        shutil.copytree(ROOT, work)
        env = os.environ.copy()
        env["MUTAKAMIL_PORT"] = str(PORT)
        proc = subprocess.Popen([sys.executable, "app.py"], cwd=work, env=env,
                                stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
        try:
            db = work / "data" / "قاعدة_البيانات.sqlite"
            for _ in range(50):
                if db.exists():
                    try:
                        urllib.request.urlopen(BASE_URL + "/", timeout=1).close()
                        break
                    except Exception:
                        pass
                time.sleep(.1)
            else:
                raise AssertionError("application did not start")

            c = sqlite3.connect(db)
            c.row_factory = sqlite3.Row
            c.execute("PRAGMA foreign_keys=ON")
            c.execute("INSERT INTO suppliers(supplier_name,normalized_name) VALUES(?,?)", ("مورد اختبار", "مورد اختبار"))
            supplier_id = c.execute("SELECT last_insert_rowid()").fetchone()[0]
            c.execute("INSERT INTO local_items(item_code,item_name,normalized_name,main_unit,active,source,export_status) VALUES(?,?,?,?,1,'IMPORTED','IMPORTED')", ("900001","جبنة موزاريلا اختبار","جبنه موزاريلا اختبار","حبة"))
            item1 = c.execute("SELECT last_insert_rowid()").fetchone()[0]
            c.execute("INSERT INTO local_items(item_code,item_name,normalized_name,main_unit,active,source,export_status) VALUES(?,?,?,?,1,'IMPORTED','IMPORTED')", ("900002","صنف آخر اختبار","صنف اخر اختبار","حبة"))
            item2 = c.execute("SELECT last_insert_rowid()").fetchone()[0]
            c.execute("INSERT INTO supplier_item_mappings(supplier_id,local_item_id,supplier_item_code,supplier_item_name,normalized_name,status,match_method,usage_count) VALUES(?,?,?,?,?,'APPROVED','SUPPLIER_CODE',1)", (supplier_id,item1,"024097","جبنة موزاريلا اختبار","جبنه موزاريلا اختبار"))
            mapping1 = c.execute("SELECT last_insert_rowid()").fetchone()[0]
            c.commit(); c.close()

            # Leading zeros and Excel .0 conversion must resolve to the same local item.
            for code in ("024097", "24097", "24097.0"):
                r = request_json("/api/match", {"supplier_id": supplier_id, "lines":[{"supplier_id":supplier_id,"supplier_item_code":code,"raw_item_name":"جبنة موزاريلا اختبار","unit":"حبة"}]})
                m = r["lines"][0]["match"]
                assert m["status"] == "AUTO_MATCHED", (code, m)
                assert int(m["item"]["id"]) == item1, (code, m)
                assert int(m["mapping_id"]) == mapping1, (code, m)

            # A normalized-code conflict must not guess between two local items.
            c = sqlite3.connect(db)
            c.execute("INSERT INTO supplier_item_mappings(supplier_id,local_item_id,supplier_item_code,supplier_item_name,normalized_name,status,match_method,usage_count) VALUES(?,?,?,?,?,'APPROVED','SUPPLIER_CODE',1)", (supplier_id,item2,"24097","صنف آخر اختبار","صنف اخر اختبار"))
            c.commit(); c.close()
            r = request_json("/api/match", {"supplier_id": supplier_id, "lines":[{"supplier_id":supplier_id,"supplier_item_code":"24097.0","raw_item_name":"اسم غير حاسم","unit":"حبة"}]})
            m = r["lines"][0]["match"]
            assert m["status"] in ("NEW_ITEM", "SUGGESTED"), m
            assert m.get("item") is None or int(m["item"]["id"]) not in (item1, item2), m

            # Approval is idempotent: repeating it must not raise an FK error.
            line = {"supplier_id":supplier_id,"supplier_item_code":"024097","raw_item_name":"جبنة موزاريلا اختبار","unit":"حبة",
                    "match":{"status":"SUGGESTED","item":{"id":item1,"item_name":"جبنة موزاريلا اختبار"},"mapping_id":mapping1,"score":95,"method":"MANUAL"}}
            for _ in range(2):
                r = request_json("/api/approve", {"supplier_id":supplier_id,"line":line})
                assert r.get("ok") is True, r

            print("PASS: supplier-code normalization, conflict safety, repeated approval")
        finally:
            proc.terminate()
            try:
                proc.wait(timeout=3)
            except subprocess.TimeoutExpired:
                proc.kill()


if __name__ == "__main__":
    main()
