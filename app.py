import os, sqlite3, json, re, webbrowser, threading, time, tempfile, shutil, subprocess
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, unquote
from pathlib import Path

BASE = Path(__file__).resolve().parent
# Keep live data outside the release package. On first run after upgrading an older
# portable release, migrate the legacy root database into the dedicated data folder.
DATA_DIR = BASE / "data"
DATA_DIR.mkdir(exist_ok=True)
DB = DATA_DIR / "قاعدة_البيانات.sqlite"
LEGACY_DB = BASE / "قاعدة_البيانات.sqlite"
APP_DB_SCHEMA_VERSION = 40
if not DB.exists() and LEGACY_DB.exists() and LEGACY_DB.stat().st_size > 0:
    try:
        shutil.copy2(LEGACY_DB, DB)
    except Exception:
        pass

def backup_database(force=False):
    """Create a safety backup before a database-schema upgrade.
    Normal launches do not create endless backups once the DB schema is current.
    """
    import shutil, datetime, sqlite3
    if not DB.exists() or DB.stat().st_size == 0:
        return None
    if not force:
        try:
            c = sqlite3.connect(DB)
            current = int(c.execute('PRAGMA user_version').fetchone()[0] or 0)
            c.close()
            if current >= APP_DB_SCHEMA_VERSION:
                return None
        except Exception:
            pass
    backup_dir = BASE / "نسخ_احتياطية"
    backup_dir.mkdir(exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = backup_dir / f"قاعدة_البيانات_{stamp}.sqlite"
    try:
        shutil.copy2(DB, dest)
        return dest
    except Exception:
        return None
PORT = int(os.environ.get("MUTAKAMIL_PORT", "8767"))

def norm(s):
    s = str(s or "").strip().lower()
    mp = {"أ":"ا","إ":"ا","آ":"ا","ى":"ي","ة":"ه","ؤ":"و","ئ":"ي","ـ":""}
    s = re.sub(r"[أإآىةؤئـ]", lambda m: mp[m.group()], s)
    s = re.sub(r"[٠-٩]", lambda m: str("٠١٢٣٤٥٦٧٨٩".index(m.group())), s)
    s = s.replace("×","x").replace("*","x")
    return re.sub(r"\s+", " ", s)

def normalize_supplier_code(value):
    """Normalize supplier codes only for matching/search, never for stored display.

    PDF/OCR/Excel conversions often turn codes such as 024097 into 24097 or
    24097.0. The original value is always retained; this helper is only a
    comparison key.
    """
    s = norm(value)
    s = re.sub(r"[\s,]", "", s)
    if re.fullmatch(r"[0-9]+\.0+", s):
        s = s.split('.', 1)[0]
    if re.fullmatch(r"[0-9]+", s):
        s = s.lstrip('0') or '0'
    return s

def tokens(s):
    return set(x for x in norm(s).split() if len(x) > 1)

def jaccard(a,b):
    A,B=tokens(a),tokens(b)
    if not A or not B: return 0
    return round(100*len(A&B)/len(A|B),1)

def dbconn():
    c=sqlite3.connect(DB, timeout=10)
    c.row_factory=sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    c.execute("PRAGMA busy_timeout=10000")
    c.execute("PRAGMA journal_mode=WAL")
    return c

def items():
    c=dbconn()
    rows=c.execute("""SELECT li.id,li.item_code,li.group_code,li.item_name,li.main_unit,li.item_type,li.source,li.export_status,li.active,
                      GROUP_CONCAT(DISTINCT su.supplier_item_code) supplier_codes
                      FROM local_items li
                      LEFT JOIN supplier_item_mappings su ON su.local_item_id=li.id
                      WHERE li.active=1 GROUP BY li.id ORDER BY li.item_code""").fetchall()
    c.close()
    return [dict(r) for r in rows]

def suppliers():
    c=dbconn()
    rows=c.execute("SELECT id,supplier_code,supplier_name,vat_number,commercial_register,phone,active FROM suppliers WHERE active=1 ORDER BY supplier_name").fetchall()
    c.close()
    return [dict(r) for r in rows]

def mappings():
    c=dbconn()
    rows=c.execute("""SELECT m.id,m.supplier_id,m.local_item_id,m.supplier_item_code,m.supplier_item_name,
                      m.supplier_unit,m.match_method,m.first_seen_at,m.last_seen_at,m.usage_count,
                      m.confidence_score,m.status,s.supplier_name,li.item_code,li.item_name,
                      (SELECT COUNT(*) FROM supplier_item_aliases a WHERE a.supplier_id=m.supplier_id AND a.local_item_id=m.local_item_id) AS alias_count,
                      (SELECT COUNT(*) FROM supplier_item_aliases a WHERE a.supplier_id=m.supplier_id AND a.local_item_id=m.local_item_id AND COALESCE(a.active,1)=1) AS active_alias_count,
                      (SELECT GROUP_CONCAT(a.alias_name,' | ') FROM supplier_item_aliases a WHERE a.supplier_id=m.supplier_id AND a.local_item_id=m.local_item_id) AS aliases
                      FROM supplier_item_mappings m
                      JOIN suppliers s ON s.id=m.supplier_id
                      JOIN local_items li ON li.id=m.local_item_id
                      ORDER BY s.supplier_name,m.supplier_item_code,m.supplier_item_name""").fetchall()
    c.close()
    return [dict(r) for r in rows]

def aliases():
    c=dbconn()
    rows=c.execute("""SELECT a.id,a.supplier_id,a.local_item_id,a.mapping_id,a.alias_name,a.source,a.usage_count,a.active,
                      a.first_seen_at,a.last_seen_at,a.created_at,a.updated_at,
                      s.supplier_name,li.item_code,li.item_name
                      FROM supplier_item_aliases a
                      JOIN suppliers s ON s.id=a.supplier_id
                      JOIN local_items li ON li.id=a.local_item_id
                      ORDER BY s.supplier_name,li.item_code,a.alias_name""").fetchall()
    c.close()
    return [dict(r) for r in rows]


# ---------------- Excel Extraction Engine v2 ----------------
EXCEL_SEMANTICS = {
    "supplier_name": ["اسم المورد", "المورد", "اسم البائع", "supplier", "supplier name", "supplier_name", "vendor", "vendor name"],
    "invoice_number": ["رقم الفاتورة", "رقم الفاتوره", "رقم المستند", "رقم الوثيقة", "voucher no", "voucher number", "invoice no", "invoice number", "invoice #", "document no"],
    "invoice_date": ["تاريخ الفاتورة", "التاريخ", "تاريخ", "date", "invoice date", "invoice_date", "voucher date"],
    "supplier_vat": ["الرقم الضريبي", "رقم ضريبي", "vat no", "vat number", "tax no", "tax number"],
    "customer_name": ["اسم العميل", "العميل", "customer", "customer name", "buyer"],
    "supplier_item_code": ["كود المورد", "كود الصنف عند المورد", "رقم الصنف المورد", "رقم الصنف", "item no", "item number", "item code", "supplier code", "supplier_item_code", "code", "sku", "product code"],
    "raw_item_name": ["اسم الصنف", "اسم الصنف المورد", "وصف الصنف", "الوصف", "الصنف", "البيان", "item name", "name", "description", "product", "product name", "item description", "product description"],
    "qty": ["الكمية", "كمية", "العدد", "عدد", "qty", "quantity", "ordered qty", "quantity ordered", "order qty", "qty ordered"],
    "unit": ["الوحدة", "الوحده", "unit", "uom", "unit of measure"],
    "unit_price": ["سعر الوحدة", "سعرالوحدة", "سعر الوحدة قبل الضريبة", "السعر", "سعر", "unit price", "unit_price", "price", "unitprice", "rate"],
    "vat_rate": ["نسبة الضريبة", "الضريبة %", "الضريبة%", "ضريبة %", "vat %", "vat%", "vat rate", "tax rate", "tax %", "tax%"],
    "vat_amount": ["قيمة الضريبة", "مبلغ الضريبة", "الضريبة", "قيمة vat", "vat amount", "tax amount", "vat_amount", "tax_amount"],
    "net_amount": ["الصافي", "صافي", "قبل الضريبة", "المبلغ قبل الضريبة", "الإجمالي غير شامل الضريبة", "الإجمالي الخاضع للضريبة", "net", "subtotal", "net amount", "amount before vat", "taxable"],
    "gross_amount": ["الإجمالي", "الاجمالي", "الإجمالي شامل الضريبة", "المجموع", "المجموع الشامل", "المبلغ الصافي شامل الضريبة", "total", "grand total", "gross", "gross amount", "total amount"],
    "discount": ["الخصم", "الخصومات", "مجموع الخصومات", "discount", "discount amount"],
    "expiry": ["تاريخ الانتهاء", "تاريخ الصلاحية", "expiry", "expiry date", "expiration"],
    "batch": ["الدفعة", "رقم الدفعة", "batch", "batch no", "lot", "lot number"],
}

EXCEL_ROW_SUMMARY_WORDS = ["الإجمالي", "الاجمالي", "المجموع", "الصافي", "الضريبة", "الخصم", "الخصومات", "total", "subtotal", "vat", "tax", "discount", "amount due", "net amount"]
EXCEL_SEMANTIC_ALIASES = {}
EXCEL_KEYWORD_FALLBACK = {
 "raw_item_name":["وصف","البيان","الصنف","description","itemname","productname"],"qty":["كمية","الكمية","qty","quantity"],
 "unit":["الوحدة","unit","uom"],"unit_price":["سعر","السعر","price","rate"],"supplier_item_code":["رقم الصنف","itemno","itemnumber","sku","code"],
 "net_amount":["خاضع للضريبة","قبل الضريبة","taxable","subtotal","net"],"vat_amount":["الضريبة","vat","tax"],"gross_amount":["المجموع الشامل","شامل الضريبة","total","grandtotal","gross"]}
EXCEL_SEMANTIC_LABELS = {"supplier_name":"اسم المورد","invoice_number":"رقم الفاتورة","invoice_date":"تاريخ الفاتورة","supplier_vat":"الرقم الضريبي للمورد","customer_name":"اسم العميل","supplier_item_code":"كود المورد","raw_item_name":"اسم الصنف","qty":"الكمية","unit":"الوحدة","unit_price":"سعر الوحدة","vat_rate":"نسبة الضريبة","vat_amount":"قيمة الضريبة","net_amount":"الصافي","gross_amount":"الإجمالي","discount":"الخصم","expiry":"تاريخ الانتهاء","batch":"الدفعة"}
REGEX_VAT_RATE = re.compile(r"(\d+(?:[.,]\d+)?)\s*%")

def _excel_clean(v):
    if v is None: return ""
    return str(v).replace("\xa0", " ").strip()

def _excel_key(v):
    # Keep separators for bilingual header splitting, but normalize each side.
    return re.sub(r"[^a-z0-9\u0600-\u06ff]+", " ", norm(v)).strip()

def _excel_compact(v):
    return re.sub(r"[^a-z0-9\u0600-\u06ff]", "", norm(v))
EXCEL_SEMANTIC_ALIASES.update({k:frozenset(_excel_compact(v) for v in vals) for k,vals in EXCEL_SEMANTICS.items()})

def _excel_number_candidates(v):
    s=_excel_clean(v)
    if not s:return []
    for token in ("ر.س","SAR","$","%","﷼","ريال","EGP","USD"): s=s.replace(token,"")
    s=s.replace("٬","").replace("٫",".").replace(" ","").strip()
    if not s:return []
    out=[]
    if "," in s and "." in s:
        # Last separator normally determines the decimal mark.
        try:
            if s.rfind(",") > s.rfind("."):
                out.append((float(s.replace(".","").replace(",",".")),0.95,"فاصلة عشرية"))
            else:
                out.append((float(s.replace(",","")),0.95,"نقطة عشرية"))
        except Exception: pass
        return out
    if "," in s:
        parts=s.split(",")
        if len(parts)==2:
            a,b=parts
            if len(b)==3 and a.replace("-","").isdigit():
                try: out.append((float(s.replace(",","")),0.90,"فواصل آلاف"))
                except Exception: pass
                try: out.append((float(f"{a}.{b}"),0.30,"فاصلة عشرية محتملة"))
                except Exception: pass
            else:
                try: out.append((float(f"{a}.{b}"),0.90,"فاصلة عشرية"))
                except Exception: pass
                try: out.append((float(s.replace(",","")),0.30,"فواصل آلاف محتملة"))
                except Exception: pass
        else:
            try: out.append((float(s.replace(",","")),0.85,"فواصل آلاف متعددة"))
            except Exception: pass
        return sorted(out,key=lambda x:x[1],reverse=True)
    try: out.append((float(s),1.0,"رقم مباشر"))
    except Exception: pass
    return out

def _excel_number(v):
    candidates=_excel_number_candidates(v)
    return candidates[0][0] if candidates else None

def _excel_header_match(value, semantic):
    raw=_excel_clean(value)
    if not raw:return (0,"")
    k=_excel_compact(raw); aliases=EXCEL_SEMANTIC_ALIASES.get(semantic,frozenset())
    if semantic=="gross_amount" and ("total" in k or "المجموعالشامل" in k or "شامل" in k):return (120,"عنوان إجمالي شامل الضريبة")
    if semantic=="net_amount" and ("taxable" in k or "خاضعللضريبه" in k or "قبلالضريبه" in k):return (120,"عنوان مبلغ خاضع/قبل الضريبة")
    if semantic=="vat_rate" and ("vat" in k or "ضريبه" in k) and ("%" in raw or "rate" in k or "نسبه" in k):return (125,"عنوان نسبة الضريبة")
    if semantic=="vat_amount" and ("vat" in k or "ضريبه" in k) and not ("%" in raw or "rate" in k or "نسبه" in k):return (112,"عنوان قيمة الضريبة")
    if semantic=="unit_price" and ("price" in k or "السعر" in k or "سعر" in k) and not any(x in k for x in ("tax","vat","ضريبه")):return (115,"عنوان سعر الوحدة")
    if k in aliases:return (100,"تطابق مباشر")
    parts=[p for p in re.split(r"[/|\\\n:()]+",raw) if p.strip()]; pkeys=[_excel_compact(x) for x in parts]
    best=0;reason=""
    for pk in pkeys:
        if pk in aliases:best=max(best,98);reason="تطابق جزء من العنوان";continue
        for a in aliases:
            if len(pk)>=4 and (pk in a or a in pk):best=max(best,88);reason="تطابق دلالي قريب"
    for word in EXCEL_KEYWORD_FALLBACK.get(semantic,[]):
        wk=_excel_compact(word)
        if wk and (wk in k or wk in ''.join(pkeys)):best=max(best,78);reason=reason or "تطابق كلمة مفتاحية"
    return best,reason

def _merge_multiline_header(rows, ri):
    if ri + 1 >= len(rows): return None
    curr=[_excel_clean(v) for v in rows[ri]]; nxt=[_excel_clean(v) for v in rows[ri+1]]
    if not any(curr) or not any(nxt): return None
    line_sems=("supplier_item_code","raw_item_name","qty","unit","unit_price","vat_rate","vat_amount","net_amount","gross_amount","discount","expiry","batch")
    meta_sems=("supplier_name","invoice_number","invoice_date","supplier_vat","customer_name","warehouse_number")
    curr_hits=sum(1 for v in curr if v and any(_excel_header_match(v,sem)[0]>=78 for sem in line_sems))
    curr_meta_hits=sum(1 for v in curr if v and any(_excel_header_match(v,sem)[0]>=78 for sem in meta_sems))
    next_hits=sum(1 for v in nxt if v and any(_excel_header_match(v,sem)[0]>=78 for sem in line_sems))
    if curr_hits < 2 or next_hits < 2 or curr_meta_hits: return None
    width=max(len(curr),len(nxt)); merged=[]
    for i in range(width):
        a=curr[i] if i<len(curr) else ""; b=nxt[i] if i<len(nxt) else ""
        merged.append(f"{a} {b}".strip() if a and b else (a or b))
    return merged

def _excel_classify_header_row(rows,max_scan=80):
    best=None; semantics=("supplier_item_code","raw_item_name","qty","unit","unit_price","vat_rate","vat_amount","net_amount","gross_amount","discount","expiry","batch")
    for ri,row in enumerate(rows[:max_scan]):
        candidates=[row]
        merged=_merge_multiline_header(rows,ri)
        if merged: candidates.append(merged)
        for vals0 in candidates:
            vals=[_excel_clean(v) for v in vals0]; by_sem={}
            for ci,v in enumerate(vals):
                if not v:continue
                for sem in semantics:
                    sc,rs=_excel_header_match(v,sem)
                    if sc>=78:by_sem.setdefault(sem,[]).append((sc,ci+1,rs))
            mapping={};scores={};reasons={};used=set(); alternatives={}
            for sem in sorted(by_sem,key=lambda z:max(x[0] for x in by_sem[z]),reverse=True):
                for sc,ci,rs in sorted(by_sem[sem],key=lambda x:(x[0],-x[1]),reverse=True):
                    if ci not in used:
                        mapping[sem]=ci;scores[sem]=sc;reasons[sem]=rs;used.add(ci);break
                alternatives[sem]=[(sc,ci,rs) for sc,ci,rs in sorted(by_sem.get(sem,[]),key=lambda x:(x[0],-x[1]),reverse=True) if mapping.get(sem)!=ci]
            core=sum(x in mapping for x in ("raw_item_name","qty","unit","unit_price")); transactional=sum(x in mapping for x in ("supplier_item_code","raw_item_name","qty","unit","unit_price","net_amount","vat_amount","gross_amount")); bilingual=sum("/" in v or "\n" in v for v in vals)
            score=core*20+transactional*7+bilingual*2+(25 if "raw_item_name" in mapping else 0)+(15 if "qty" in mapping else 0)+(15 if "unit_price" in mapping else 0)+(8 if "unit" in mapping and ("qty" in mapping or "unit_price" in mapping) else 0)
            if best is None or score>best[0]:best=(score,ri+1,mapping,scores,reasons,vals,alternatives, bool(merged))
    return best if best and "raw_item_name" in best[2] and len(best[2])>=2 else None

def _find_header_vat_rate(values):
    for cell in values or []:
        txt=_excel_clean(cell)
        if "vat" in _excel_compact(txt) or "ضريب" in _excel_compact(txt):
            m=REGEX_VAT_RATE.search(txt)
            if m:return float(m.group(1).replace(",","."))
    return None

def _looks_like_header(value):
    txt=_excel_clean(value)
    if not txt:return False
    # Header detection is used while extracting invoice metadata, so it must also
    # recognize line-level headers; otherwise a value such as "سعر الوحدة" can be
    # accidentally accepted as the supplier/invoice value next to it.
    semantics=("supplier_name","invoice_number","invoice_date","supplier_vat","customer_name","warehouse_number",
               "supplier_item_code","raw_item_name","qty","unit","unit_price","vat_rate","vat_amount","net_amount","gross_amount","discount")
    return any(_excel_header_match(txt,sem)[0]>=75 for sem in semantics)

def _excel_extract_invoice_meta(rows,header_row):
    meta={"supplier_name":"","invoice_number":"","invoice_date":"","supplier_vat":"","customer_name":"","warehouse_number":"","vat_rate":None,"notes":[]}
    for row in rows[:max(0,header_row-1)]:
        vals=[_excel_clean(v) for v in row]
        for i,v in enumerate(vals):
            if not v:continue
            for sem in ("supplier_name","invoice_number","invoice_date","supplier_vat","customer_name","warehouse_number"):
                if _excel_header_match(v,sem)[0]>=75 and i+1<len(vals) and vals[i+1] and not _looks_like_header(vals[i+1]) and not meta[sem]:meta[sem]=vals[i+1]
            for sem in ("supplier_name","invoice_number","invoice_date","supplier_vat","customer_name"):
                for lab in EXCEL_SEMANTICS.get(sem,[]):
                    if _excel_compact(lab) in _excel_compact(v) and ":" in v:
                        val=v.split(":",1)[1].strip()
                        if val and not meta[sem]:meta[sem]=val
                        break
    return meta

def _excel_row_kind(row, mapping):
    vals=[_excel_clean(v) for v in row]
    joined=_excel_compact(" ".join(v for v in vals if v))
    if not joined: return "BLANK"
    summary_words=[_excel_compact(x) for x in EXCEL_ROW_SUMMARY_WORDS]
    code=_excel_clean(row[mapping["supplier_item_code"]-1]) if "supplier_item_code" in mapping and mapping["supplier_item_code"]<=len(row) else ""
    name=_excel_clean(row[mapping["raw_item_name"]-1]) if "raw_item_name" in mapping and mapping["raw_item_name"]<=len(row) else ""
    q=_excel_number(row[mapping["qty"]-1]) if "qty" in mapping and mapping["qty"]<=len(row) else None
    if any(w in joined for w in summary_words) and not code and (q is None or q==0): return "SUMMARY"
    if q is not None and q > 0 and not name and not code: return "SUMMARY"
    # Do not treat footer/customer/bank rows as items merely because a text value
    # happens to sit in the code column. A line needs a name, or code + at least
    # one transactional field (quantity/unit/price).
    has_tx=False
    for sem in ("qty","unit","unit_price","net_amount","gross_amount"):
        ci=mapping.get(sem)
        if ci and ci<=len(row) and _excel_clean(row[ci-1]):
            has_tx=True; break
    if name and (has_tx or code): return "ITEM"
    if code and has_tx: return "ITEM"
    return "OTHER"

def _excel_validate_line(q,p,vat,vat_amt,net,gross):
    diagnostics=[]; checks=[]
    if q is None or p is None:
        return 0,"ناقص: الكمية أو سعر الوحدة غير معروف",False
    if q<=0 or p<0:
        return 0,"غير صالح: الكمية/السعر",False
    base=q*p
    if net is not None:
        ok=abs(base-net)<=max(0.05,abs(net)*0.01); checks.append(ok); diagnostics.append("الكمية × السعر ≈ الصافي" if ok else "الكمية × السعر لا يطابق الصافي")
    if vat is not None:
        vr=vat*100 if vat<=1 else vat
        if net is not None and vat_amt is not None:
            ok=abs(net*vr/100-vat_amt)<=max(0.05,abs(vat_amt)*0.02); checks.append(ok); diagnostics.append("الصافي × الضريبة ≈ قيمة الضريبة" if ok else "قيمة الضريبة لا تتطابق")
        if gross is not None:
            expected=(net if net is not None else base)*(1+vr/100)
            ok=abs(expected-gross)<=max(0.05,abs(gross)*0.01); checks.append(ok); diagnostics.append("الإجمالي شامل الضريبة متطابق" if ok else "الإجمالي الشامل لا يتطابق")
    elif gross is not None:
        ok=abs(base-gross)<=max(0.05,abs(gross)*0.01); checks.append(ok); diagnostics.append("الإجمالي = الكمية × السعر" if ok else "الإجمالي لا يطابق الكمية × السعر")
    if not checks: return 72,"استخراج دلالي صحيح دون إجمالي مصدر للتحقق",False
    good=sum(checks); conf=88+int(12*good/len(checks)) if good else 55
    return min(100,conf),"؛ ".join(diagnostics),all(checks)

def _analyze_column_content(rows, header_row, col_idx):
    vals=[]
    for row in rows[header_row:header_row+50]:
        if col_idx-1 < len(row):
            v=_excel_clean(row[col_idx-1])
            if v: vals.append(v)
    if len(vals) < 5:return None
    nums=[_excel_number(v) for v in vals]; ratio=sum(n is not None for n in nums)/len(vals)
    if ratio>=0.8:
        ns=[n for n in nums if n is not None]
        if ns and all(n==int(n) and 0<n<10000 for n in ns):return "qty_candidate"
        if ns and any(n!=int(n) for n in ns):return "price_candidate"
        if ns and all(n>100 for n in ns):return "amount_candidate"
    if sum(1 for v in vals if len(v)>5 and not v.isdigit())/len(vals)>0.6:return "name_candidate"
    return None

def _apply_content_mapping(rows, header_row, mapping, width):
    mapping=dict(mapping); used=set(mapping.values())
    candidates=[]
    for ci in range(1,width+1):
        if ci in used:continue
        candidates.append((ci,_analyze_column_content(rows,header_row,ci)))
    if "qty" not in mapping:
        for ci,k in candidates:
            if k=="qty_candidate":mapping["qty"]=ci;used.add(ci);break
    if "unit_price" not in mapping:
        for ci,k in candidates:
            if ci not in used and k=="price_candidate":mapping["unit_price"]=ci;used.add(ci);break
    return mapping

def _excel_parse_with_mapping(rows, sheet_name, header_row, mapping, meta=None, field_scores=None):
    out=[]; meta=meta or {}
    for row in rows[header_row:]:
        if _excel_row_kind(row,mapping)!="ITEM":
            continue
        def val(sem):
            ci=mapping.get(sem)
            return _excel_clean(row[ci-1]) if ci and ci<=len(row) else ""
        name=val("raw_item_name"); code=val("supplier_item_code")
        q=_excel_number(val("qty")); unit=val("unit"); p=_excel_number(val("unit_price"))
        vr=_excel_number(val("vat_rate")); va=_excel_number(val("vat_amount")); net=_excel_number(val("net_amount")); gross=_excel_number(val("gross_amount"))
        if vr is None and meta.get("vat_rate") not in (None, ""): vr=float(meta.get("vat_rate"))
        # Do not invent source values, but calculate missing financial values for the unified model.
        computed=[]
        if q is not None and p is not None:
            base=q*p
            if net is None:
                net=base; computed.append("الصافي محسوب من الكمية × السعر")
            if vr is not None and va is None:
                vr2=vr*100 if vr<=1 else vr; va=net*vr2/100; computed.append("الضريبة محسوبة")
            if gross is None and va is not None:
                gross=net+va; computed.append("الإجمالي شامل الضريبة محسوب")
        conf,diag,verified=_excel_validate_line(q,p,vr,va,net,gross)
        if computed: diag=("؛ ".join(computed) + ("؛ "+diag if diag else ""))
        status="VERIFIED" if conf>=90 else ("REVIEW" if conf>=60 else "ERROR")
        field_conf={}
        for sem,valx in (("supplier_item_code",code),("raw_item_name",name),("qty",q),("unit",unit),("unit_price",p),("vat_rate",vr),("vat_amount",va),("net_amount",net),("gross_amount",gross)):
            if valx in (None,""): continue
            if sem in computed or (sem=="net_amount" and "الصافي محسوب" in " ".join(computed)) or (sem=="vat_amount" and "الضريبة محسوبة" in " ".join(computed)) or (sem=="gross_amount" and "الإجمالي شامل الضريبة محسوب" in " ".join(computed)):
                field_conf[sem]=70
            else: field_conf[sem]=int(min(100, field_scores.get(sem,90) if field_scores else 90))
        out.append({"supplier_name":meta.get("supplier_name", ""),"supplier_item_code":code,"raw_item_name":name,
                    "qty":q or 0,"unit":unit,"unit_price":p or 0,"vat_rate":vr,"vat_amount":va,"net_amount":net,"gross_amount":gross,
                    "extraction_confidence":conf,"field_confidence":field_conf,"extraction_status":status,"extraction_note":diag,"_sheet":sheet_name})
    return out

def prepare_xlsx_workbook(data):
    from openpyxl import load_workbook
    import io,zipfile
    t=time.perf_counter(); wb=load_workbook(filename=io.BytesIO(data),data_only=True,read_only=True)
    formula_present=False
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z: formula_present=any(b"<f" in z.read(n) for n in z.namelist() if n.startswith("xl/worksheets/") and n.endswith(".xml"))
    except Exception:pass
    sheets=[]; prepared=[]
    for ws in wb.worksheets:
        rows=[list(r) for r in ws.iter_rows(values_only=True)]; h=_excel_classify_header_row(rows); prepared.append({"sheet":ws.title,"rows":rows,"header":h})
        if h:
            _,hr,mapping,scores,reasons,vals,alternatives,merged_header=h; inv={ci:sem for sem,ci in mapping.items()}
            headers=[{"column":ci,"header":_excel_clean(v),"semantic":EXCEL_SEMANTIC_LABELS.get(inv.get(ci,""),inv.get(ci,"")),"score":scores.get(inv.get(ci,""),0),"reason":reasons.get(inv.get(ci,""),"")} for ci,v in enumerate(vals,1) if _excel_clean(v)]
            sig="|".join(_excel_compact(v) for v in vals if _excel_clean(v)); meta=_excel_extract_invoice_meta(rows,hr); meta["vat_rate"]=_find_header_vat_rate(vals)
            ambiguity=[]
            for sem,cands in (alternatives or {}).items():
                if len(cands)>1:
                    ambiguity.append({"semantic":EXCEL_SEMANTIC_LABELS.get(sem,sem),"columns":[x[1] for x in cands],"chosen":mapping.get(sem),"reason":"تم اختيار أعلى تطابق مع الاحتفاظ بالبدائل للمراجعة"})
            sheets.append({"sheet":ws.title,"header_row":hr,"headers":headers,"column_map":{EXCEL_SEMANTIC_LABELS.get(k,k):v for k,v in mapping.items()},"signature":sig,"status":"DETECTED","invoice":meta,"score":h[0],"column_ambiguities":ambiguity,"merged_header":bool(merged_header)})
        else:sheets.append({"sheet":ws.title,"header_row":None,"headers":[],"column_map":{},"signature":"","status":"HEADER_NOT_FOUND","invoice":{},"score":0})
    primary=max((x for x in sheets if x["status"]=="DETECTED"),key=lambda x:x.get("score",0),default=None)
    return {"structure":{"sheets":sheets,"primary":primary,"sheet_count":len(sheets),"engine":"v2.2","formula_diagnostics":{"formula_present":formula_present,"warning":"يوجد خلايا بمعادلات؛ النتائج غير المحفوظة قد تظهر فارغة ولا يتم اختراعها." if formula_present else ""},"timing":{"prepare_ms":round((time.perf_counter()-t)*1000,2)}},"sheets":prepared,"formula_present":formula_present}

def analyze_xlsx_structure(data):
    return prepare_xlsx_workbook(data)["structure"]

def _apply_corrections_hint(mapping, supplier_id, header_signature):
    if not header_signature: return dict(mapping)
    c=dbconn()
    try:
        rows=c.execute("""SELECT field,corrected_column FROM extraction_corrections
                         WHERE supplier_id IS ? AND header_signature=?
                         ORDER BY learned_at DESC,id DESC""",(supplier_id,header_signature)).fetchall()
    finally: c.close()
    out=dict(mapping); seen=set()
    for r in rows:
        field=_semantic_key(r["field"])
        if field in seen or not r["corrected_column"]: continue
        col=int(r["corrected_column"]); out[field]=col; seen.add(field)
    return out

def parse_xlsx_bytes(data,template=None,analyzed=None):
    t=time.perf_counter(); prepared=analyzed if analyzed and analyzed.get("sheets") is not None else prepare_xlsx_workbook(data); all_lines=[]
    for item in prepared["sheets"]:
        rows=item["rows"];h=item.get("header")
        if not h:continue
        _,hr,mapping,scores,reasons,vals=h[:6]; sig="|".join(_excel_compact(v) for v in vals if _excel_clean(v))
        if template and template.get("column_map") and template.get("header_signature")==sig:
            tm={_semantic_key(k):int(v) for k,v in (template.get("column_map") or {}).items() if str(v).isdigit()}
            if tm.get("raw_item_name"):mapping=dict(mapping);mapping.update(tm)
            mapping=_apply_corrections_hint(mapping, template.get("supplier_id"), sig)
        mapping=_apply_content_mapping(rows,hr,mapping,len(vals))
        meta=_excel_extract_invoice_meta(rows,hr); vrate=_find_header_vat_rate(vals)
        if vrate is not None:meta["vat_rate"]=vrate
        all_lines.extend(_excel_parse_with_mapping(rows,item["sheet"],hr,mapping,meta,scores))
    seen=set();final=[]
    for x in all_lines:
        sig=(norm(x.get("supplier_item_code")),norm(x.get("raw_item_name")),str(x.get("qty")),str(x.get("unit_price")),str(x.get("unit")))
        if not x.get("raw_item_name") and not x.get("supplier_item_code"):continue
        if sig in seen:continue
        seen.add(sig);final.append(x)
    prepared["structure"]["timing"]["extract_ms"]=round((time.perf_counter()-t)*1000,2)
    if prepared.get("formula_present"):
        for x in final:
            if x.get("net_amount") is None or x.get("gross_amount") is None:
                x["extraction_status"]="REVIEW";x["extraction_note"]=(x.get("extraction_note") or "")+" || تحذير المعادلات: قد لا تكون بعض النتائج المحفوظة متاحة."
    return final

def _semantic_key(label):
    for k in EXCEL_SEMANTICS:
        if EXCEL_SEMANTIC_LABELS.get(k)==label or k==label: return k
    return label

def extraction_template_get(supplier_id,header_signature):
    c=dbconn();r=None
    if supplier_id:r=c.execute("SELECT * FROM invoice_extraction_templates WHERE supplier_id=? AND header_signature=? AND active=1 LIMIT 1",(supplier_id,header_signature)).fetchone()
    if not r:r=c.execute("SELECT * FROM invoice_extraction_templates WHERE supplier_id IS NULL AND header_signature=? AND active=1 LIMIT 1",(header_signature,)).fetchone()
    c.close()
    if not r:return None
    d=dict(r);d["column_map"]=json.loads(d.pop("column_map_json") or "{}");return d

def extraction_template_save(payload):
    supplier_id=int(payload.get('supplier_id') or 0) or None
    signature=str(payload.get('header_signature') or '').strip()
    mapping=payload.get('column_map') or {}
    name=str(payload.get('template_name') or 'قالب Excel').strip() or 'قالب Excel'
    if not signature or not mapping: raise ValueError('بيانات قالب الاستخراج غير مكتملة.')
    c=dbconn()
    existing=c.execute("SELECT id,column_map_json FROM invoice_extraction_templates WHERE supplier_id IS ? AND header_signature=?",(supplier_id,signature)).fetchone()
    if existing:
        old_map=json.loads(existing['column_map_json'] or "{}")
        for field,new_col in mapping.items():
            old_col=old_map.get(field)
            if old_col is not None and str(old_col)!=str(new_col):
                c.execute("INSERT INTO extraction_corrections(supplier_id,header_signature,field,original_column,corrected_column) VALUES(?,?,?,?,?)",(supplier_id,signature,field,int(old_col),int(new_col)))
        c.execute("UPDATE invoice_extraction_templates SET template_name=?,column_map_json=?,updated_at=CURRENT_TIMESTAMP,active=1 WHERE id=?",(name,json.dumps(mapping,ensure_ascii=False),existing['id']))
        tid=existing['id']
    else:
        sname=None
        if supplier_id:
            r=c.execute('SELECT supplier_name FROM suppliers WHERE id=?',(supplier_id,)).fetchone(); sname=r['supplier_name'] if r else None
        cur=c.execute("INSERT INTO invoice_extraction_templates(supplier_id,supplier_name,template_name,header_signature,column_map_json) VALUES(?,?,?,?,?)",(supplier_id,sname,name,signature,json.dumps(mapping,ensure_ascii=False)))
        tid=cur.lastrowid
    c.commit(); r=c.execute('SELECT * FROM invoice_extraction_templates WHERE id=?',(tid,)).fetchone(); c.close()
    d=dict(r); d['column_map']=json.loads(d.pop('column_map_json') or '{}'); return d

def parse_csv_bytes(data):
    import csv, io
    text=data.decode("utf-8-sig",errors="replace")
    rows=list(csv.DictReader(io.StringIO(text)))
    def get(r,names):
        nm={norm(x) for x in names}
        for k,v in r.items():
            if norm(k) in nm: return v or ""
        return ""
    out=[]
    for r in rows:
        name=get(r,["اسم الصنف","الوصف","الصنف","item_name","name","description"])
        code=get(r,["كود المورد","كود الصنف","supplier_item_code","supplier_code","item_code"])
        if not name and not code: continue
        q=_excel_number(get(r,["الكمية","qty","quantity"])) or 0
        p=_excel_number(get(r,["سعر الوحدة","السعر","unit_price","price"])) or 0
        out.append({"supplier_name":get(r,["اسم المورد","المورد","supplier_name","supplier"]),
                    "supplier_item_code":code,"raw_item_name":name,"qty":q,
                    "unit":get(r,["الوحدة","الوحده","unit"]),"unit_price":p})
    return out

def find_tesseract():
    """Find Tesseract on Windows/Linux without requiring PATH configuration."""
    import shutil, os
    candidates=[]
    w=os.environ.get("ProgramFiles", r"C:\\Program Files")
    w86=os.environ.get("ProgramFiles(x86)", r"C:\\Program Files (x86)")
    local=os.environ.get("LOCALAPPDATA", "")
    for base in [w,w86,local]:
        if base:
            candidates += [os.path.join(base,"Tesseract-OCR","tesseract.exe"),
                           os.path.join(base,"Tesseract-OCR","tesseract")]
    candidates += [shutil.which("tesseract") or ""]
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    return None

def ocr_file(data, filename):
    from PIL import Image, ImageOps, ImageEnhance, ImageFilter
    import pytesseract, io
    suffix=Path(filename).suffix.lower()
    images=[]
    if suffix==".pdf":
        import fitz
        doc=fitz.open(stream=data,filetype="pdf")
        # First try native PDF text. Digital PDFs do not need Tesseract at all.
        native=[]
        for page in doc:
            native.append(page.get_text("text") or "")
        native_text="\n".join(native).strip()
        if len(re.sub(r"\s+","",native_text)) >= 80:
            return native_text
        for page in doc:
            pix=page.get_pixmap(matrix=fitz.Matrix(1.6,1.6), alpha=False)
            images.append(Image.frombytes("RGB",[pix.width,pix.height],pix.samples))
    else:
        images=[Image.open(io.BytesIO(data)).convert("RGB")]

    tess=find_tesseract()
    if not tess:
        raise RuntimeError("لم يتم العثور على Tesseract OCR. شغّل ملف SETUP_OCR.bat مرة واحدة ثم أعد تشغيل التطبيق.")
    pytesseract.pytesseract.tesseract_cmd=tess
    try:
        langs=pytesseract.get_languages(config='')
    except Exception:
        langs=[]
    if 'ara' not in langs:
        raise RuntimeError("Tesseract موجود لكن لغة العربية (ara) غير مثبتة. شغّل SETUP_OCR.bat لإكمال تثبيت العربية.")

    texts=[]
    for im in images:
        gray=ImageOps.grayscale(im)
        gray=ImageEnhance.Contrast(gray).enhance(1.6)
        gray=gray.filter(ImageFilter.SHARPEN)
        txt=pytesseract.image_to_string(gray, lang="ara+eng", config="--oem 3 --psm 6")
        texts.append(txt)
    return "\n".join(texts)

def parse_ocr_invoice_text(text):
    """Convert OCR table rows into reviewable invoice lines using conservative table heuristics."""
    lines=[]
    numpat=r"(?:\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)"
    for raw in str(text or "").splitlines():
        line=re.sub(r"[\u200e\u200f\u202a-\u202e]","",str(raw))
        line=re.sub(r"\s+"," ",line).strip()
        if len(line)<8: continue
        nl=norm(line)
        if any(k in nl for k in ["المجموع","الاجمالي","subtotal","total","tax invoice","فاتوره ضريبيه","tax sales invoice","رقم العميل","الرقم الضريبي","التاريخ","اسم العميل","رقم الفاتوره"]): continue
        # Normalize leading OCR punctuation so a supplier code can be anchored.
        cleanline=re.sub(r"^[^0-9A-Za-z\u0600-\u06FF]+","",line)
        cm=re.match(r"([0-9]{3,8})(?=\s)",cleanline)
        if not cm: continue
        code=cm.group(1)
        rest=cleanline[cm.end():].strip()
        if not re.search(r"[A-Za-z\u0600-\u06FF]",rest): continue
        # Take numeric tokens after the description. The final commercial triple is
        # normally qty, unit price, tax/gross. Ignore extra trailing page numbers.
        nums=list(re.finditer(numpat,rest))
        if len(nums)<3: continue
        # Exclude weight/pack numbers in the description by selecting the first triple
        # that is near the end and has a plausible quantity/price pair.
        best=None
        for j in range(len(nums)-3,-1,-1):
            trio=nums[j:j+3]
            a=[float(x.group().replace(',','')) for x in trio]
            desc=rest[:trio[0].start()].strip(" -|/[]()")
            if not re.search(r"[A-Za-z\u0600-\u06FF]",desc): continue
            qty,price,tax=a
            if qty<=0 or price<0: continue
            # Quantity is generally an integer and less than the invoice monetary amounts.
            if abs(qty-round(qty))<1e-9 and qty<=100000 and price<=100000 and tax>=price:
                best=(desc,qty,price,tax); break
        if not best: continue
        name,qty,price,tax=best
        if any(k in norm(name) for k in ["رقم","التاريخ","العميل","الضريبي","الفاتوره"]): continue
        lines.append({"supplier_name":"","supplier_item_code":code,"raw_item_name":name,
                      "qty":qty,"unit":"","unit_price":price,"_ocr_raw":line})
    out=[]; seen=set()
    for x in lines:
        sig=(norm(x["supplier_item_code"]),norm(x["raw_item_name"]),str(x["qty"]))
        if sig in seen: continue
        seen.add(sig); out.append(x)
    return out

def _name_features(value):
    """Extract order-independent lexical/numeric features from a supplier description."""
    raw=norm(value)
    raw=re.sub(r'x+', 'x', raw)
    # Keep Arabic words, Latin words and numeric fragments as separate tokens. This
    # deliberately breaks forms such as 4x24 / 4*x*24 into comparable components.
    parts=re.findall(r'[\u0600-\u06ff]+|[a-z]+|\d+(?:[\.,]\d+)?', raw)
    words={x for x in parts if not re.fullmatch(r'\d+(?:[\.,]\d+)?',x) and x not in {'x'}}
    nums={x.replace(',','.') for x in parts if re.fullmatch(r'\d+(?:[\.,]\d+)?',x)}
    compact=' '.join(parts)
    return raw, words, nums, compact

def _set_similarity(a,b):
    if not a or not b: return 0.0
    return 100.0*len(a & b)/len(a | b)

def _sequence_similarity(a,b):
    from difflib import SequenceMatcher
    if not a or not b: return 0.0
    return 100.0*SequenceMatcher(None,a,b,autojunk=False).ratio()

def _numeric_similarity(nums_a,nums_b):
    if not nums_a or not nums_b:
        return None
    if nums_a == nums_b: return 100.0
    overlap=len(nums_a & nums_b)
    if not overlap: return 0.0
    return 100.0*overlap/max(len(nums_a),len(nums_b))

def _price_similarity(current_price, historical_prices):
    if current_price in (None,'') or not historical_prices: return None
    try: current=float(current_price)
    except Exception: return None
    if current <= 0: return None
    vals=sorted(float(x) for x in historical_prices if x is not None and float(x)>0)
    if not vals: return None
    mid=len(vals)//2
    median=vals[mid] if len(vals)%2 else (vals[mid-1]+vals[mid])/2
    if median <= 0: return None
    diff=abs(current-median)/median
    return max(0.0,100.0-(diff*100.0*2.0))

def _candidate_score(line, candidate, source_kind='LOCAL_NAME', price_history=None):
    raw, words, nums, compact=_name_features(line.get('raw_item_name') or '')
    cname=candidate.get('supplier_item_name') if source_kind in ('SUPPLIER_MAPPING_NAME','ALIAS_SUGGESTION') else candidate.get('item_name')
    craw,cwords,cnums,ccompact=_name_features(cname or '')
    word_score=_set_similarity(words,cwords)
    char_score=_sequence_similarity(compact,ccompact)
    num_score=_numeric_similarity(nums,cnums)
    # Numeric agreement is especially valuable for pack/size descriptions. If one side
    # has no numeric data, do not penalize the candidate; simply redistribute the weight.
    components=[(word_score,0.55),(char_score,0.25)]
    if num_score is not None: components.append((num_score,0.20))
    total=sum(v*w for v,w in components)/sum(w for _,w in components)
    unit=norm(line.get('unit'))
    cand_unit=norm(candidate.get('supplier_unit') or candidate.get('normalized_unit') or candidate.get('main_unit'))
    if unit and cand_unit:
        total=total*0.92 + (100.0 if unit==cand_unit else 0.0)*0.08
    ps=_price_similarity(line.get('unit_price'), price_history or [])
    if ps is not None:
        total=total*0.90 + ps*0.10
    return round(min(100.0,max(0.0,total)),1), {'name':round(total,1),'word':round(word_score,1),'numbers':None if num_score is None else round(num_score,1),'unit':unit and cand_unit and unit==cand_unit,'price':None if ps is None else round(ps,1)}

def _build_match_context(c, supplier_id, all_items=None):
    if all_items is None:
        all_items=c.execute("SELECT * FROM local_items WHERE active=1").fetchall()
    mappings=[]; aliases=[]; prices={}
    if supplier_id:
        mappings=c.execute("""SELECT m.id AS mapping_id,m.supplier_id,m.local_item_id,m.supplier_item_name,m.supplier_item_code,
                                    m.normalized_name,m.supplier_unit,m.normalized_unit,m.usage_count,m.confidence_score,m.status,li.*
                             FROM supplier_item_mappings m JOIN local_items li ON li.id=m.local_item_id
                             WHERE m.supplier_id=? AND m.status!='DISABLED'""",(int(supplier_id),)).fetchall()
        alias_cols={r['name'] for r in c.execute("PRAGMA table_info(supplier_item_aliases)").fetchall()}
        alias_active_clause=" AND a.active=1" if 'active' in alias_cols else ""
        aliases=c.execute(f"""SELECT a.id,a.mapping_id,a.supplier_id,a.local_item_id,a.alias_name,a.normalized_name,a.usage_count,
                                  li.* FROM supplier_item_aliases a JOIN local_items li ON li.id=a.local_item_id
                                  WHERE a.supplier_id=?{alias_active_clause}""",(int(supplier_id),)).fetchall()
        rows=c.execute("SELECT local_item_id,unit_price_halalah FROM supplier_item_prices WHERE supplier_id=? AND unit_price_halalah>0 ORDER BY invoice_date DESC,id DESC",(int(supplier_id),)).fetchall()
        for r in rows:
            iid=int(r['local_item_id']); prices.setdefault(iid,[])
            if len(prices[iid])<12: prices[iid].append(float(r['unit_price_halalah'])/100.0)
    return {'all_items':all_items,'mappings':mappings,'aliases':aliases,'prices':prices}

def _code_keys(value, sequence=None):
    """Return comparison keys without overwriting the original extracted code."""
    raw=str(value or '').strip(); keys=[]
    full=normalize_supplier_code(re.sub(r'[()]','',raw))
    if full: keys.append(('FULL',full))
    m=re.fullmatch(r'(.+?)\((\d+)\)\s*',raw)
    if m:
        suffix=int(m.group(2)); base=normalize_supplier_code(m.group(1))
        if base and sequence is not None and suffix==int(sequence): keys.append(('SEQUENCE_SUFFIX',base))
    return keys

def _record_supplier_mapping_conflict(c, supplier_id, supplier_item_code, local_item_id=None,
                                       existing_local_item_id=None, invoice_id=None, invoice_line_id=None,
                                       reason=''):
    """Persist code/alias conflicts using either the legacy or v2 conflict schema."""
    try:
        cols={r['name'] for r in c.execute("PRAGMA table_info(supplier_mapping_conflicts)").fetchall()}
        if {'attempted_local_item_id','conflicting_local_item_id','supplier_code'} <= cols:
            c.execute("""INSERT INTO supplier_mapping_conflicts
                        (supplier_id,supplier_code,normalized_code,invoice_id,invoice_line_id,
                         attempted_local_item_id,conflicting_local_item_id,status,notes,created_at)
                        VALUES(?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)""",
                      (supplier_id,str(supplier_item_code or '').strip(),normalize_supplier_code(supplier_item_code),invoice_id,invoice_line_id,
                       local_item_id,existing_local_item_id,'PENDING',reason or 'تعارض في ربط المورد'))
        elif {'local_item_id','existing_local_item_id','supplier_item_code'} <= cols:
            c.execute("""INSERT INTO supplier_mapping_conflicts
                        (supplier_id,supplier_item_code,normalized_code,local_item_id,existing_local_item_id,
                         invoice_id,invoice_line_id,reason,status,created_at)
                        VALUES(?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)""",
                      (supplier_id,str(supplier_item_code or '').strip(),normalize_supplier_code(supplier_item_code),local_item_id,existing_local_item_id,
                       invoice_id,invoice_line_id,reason or 'تعارض في ربط المورد','OPEN'))
    except sqlite3.OperationalError:
        pass

def supplier_mapping_conflicts(limit=300, status=None):
    c=dbconn()
    try:
        cols={r['name'] for r in c.execute("PRAGMA table_info(supplier_mapping_conflicts)").fetchall()}
        where=[]; params=[]
        if status:
            where.append('mc.status=?'); params.append(status)
        clause=(' WHERE '+' AND '.join(where)) if where else ''
        if {'attempted_local_item_id','conflicting_local_item_id','supplier_code'} <= cols:
            rows=c.execute(f"""SELECT mc.*,s.supplier_name,
                              li1.item_code AS attempted_item_code,li1.item_name AS attempted_item_name,
                              li2.item_code AS conflicting_item_code,li2.item_name AS conflicting_item_name
                              FROM supplier_mapping_conflicts mc
                              LEFT JOIN suppliers s ON s.id=mc.supplier_id
                              LEFT JOIN local_items li1 ON li1.id=mc.attempted_local_item_id
                              LEFT JOIN local_items li2 ON li2.id=mc.conflicting_local_item_id
                              {clause} ORDER BY mc.id DESC LIMIT ?""",params+[int(limit or 300)]).fetchall()
        else:
            rows=c.execute(f"""SELECT mc.*,s.supplier_name,
                              li1.item_code AS attempted_item_code,li1.item_name AS attempted_item_name,
                              li2.item_code AS conflicting_item_code,li2.item_name AS conflicting_item_name
                              FROM supplier_mapping_conflicts mc
                              LEFT JOIN suppliers s ON s.id=mc.supplier_id
                              LEFT JOIN local_items li1 ON li1.id=mc.local_item_id
                              LEFT JOIN local_items li2 ON li2.id=mc.existing_local_item_id
                              {clause} ORDER BY mc.id DESC LIMIT ?""",params+[int(limit or 300)]).fetchall()
        return [dict(r) for r in rows]
    finally:
        c.close()

def _active_local_item(c, local_item_id):
    try:
        iid=int(local_item_id or 0)
    except Exception:
        return None
    if not iid:
        return None
    row=c.execute("SELECT id FROM local_items WHERE id=? AND active=1 LIMIT 1",(iid,)).fetchone()
    return int(row['id']) if row else None


def toggle_learning_alias(alias_id, active):
    aid=int(alias_id or 0); target=1 if active else 0
    if not aid: raise ValueError('معرف الاسم البديل مطلوب.')
    c=dbconn()
    try:
        row=c.execute("SELECT * FROM supplier_item_aliases WHERE id=?",(aid,)).fetchone()
        if not row: raise ValueError('الاسم البديل غير موجود.')
        before=dict(row)
        target_item=_active_local_item(c,row['local_item_id'])
        if target and not target_item:
            raise ValueError('لا يمكن تفعيل الاسم البديل لأن الصنف المحلي المرتبط به غير نشط أو غير موجود.')
        if target:
            others=_active_alias_matches(c,row['supplier_id'],row['alias_name'],aid)
            if any(int(r['local_item_id'])!=int(row['local_item_id']) for r in others):
                raise ValueError('لا يمكن تفعيل الاسم البديل لأنه مستخدم حاليًا لصنف محلي آخر لدى نفس المورد.')
        c.execute("UPDATE supplier_item_aliases SET active=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",(target,aid))
        after=dict(c.execute("SELECT * FROM supplier_item_aliases WHERE id=?",(aid,)).fetchone())
        log_learning_history(c,'ALIAS',aid,row['supplier_id'],row['local_item_id'],'ENABLE' if target else 'DISABLE',before,after,source='USER')
        c.commit(); return after
    except: c.rollback(); raise
    finally: c.close()


def update_learning_alias(payload):
    aid=int(payload.get('id') or 0); alias=str(payload.get('alias_name') or '').strip(); local_item_id=int(payload.get('local_item_id') or 0)
    if not aid or not alias or not local_item_id: raise ValueError('معرف الاسم والاسم والصنف المحلي مطلوبة.')
    c=dbconn()
    try:
        row=c.execute("SELECT * FROM supplier_item_aliases WHERE id=?",(aid,)).fetchone()
        if not row: raise ValueError('الاسم البديل غير موجود.')
        target_item=_active_local_item(c,local_item_id)
        if not target_item: raise ValueError('الصنف المحلي المحدد غير موجود أو غير نشط.')
        before=dict(row)
        other=_active_alias_matches(c,row['supplier_id'],alias,aid)
        if any(int(r['local_item_id'])!=local_item_id for r in other): raise ValueError('هذا الاسم متعلم بالفعل لصنف محلي آخر لدى نفس المورد.')
        # A mapping_id pointing to the old item must not remain attached after an alias re-assignment.
        mapping_id=row['mapping_id']
        if mapping_id:
            mr=c.execute("SELECT local_item_id FROM supplier_item_mappings WHERE id=?",(mapping_id,)).fetchone()
            if mr and int(mr['local_item_id'])!=local_item_id:
                mapping_id=None
        active=1 if payload.get('active',True) else 0
        c.execute("UPDATE supplier_item_aliases SET local_item_id=?,mapping_id=?,alias_name=?,normalized_name=?,active=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",
                  (local_item_id,mapping_id,alias,norm(alias),active,aid))
        after=dict(c.execute("SELECT * FROM supplier_item_aliases WHERE id=?",(aid,)).fetchone())
        log_learning_history(c,'ALIAS',aid,row['supplier_id'],local_item_id,'UPDATE',before,after,source='USER')
        c.commit(); return after
    except: c.rollback(); raise
    finally: c.close()


def toggle_learning_mapping(mapping_id, active):
    mid=int(mapping_id or 0); target=1 if active else 0
    if not mid: raise ValueError('معرف الربط مطلوب.')
    c=dbconn()
    try:
        row=c.execute("SELECT * FROM supplier_item_mappings WHERE id=?",(mid,)).fetchone()
        if not row: raise ValueError('الربط غير موجود.')
        before=dict(row)
        target_item=_active_local_item(c,row['local_item_id'])
        if target and not target_item:
            raise ValueError('لا يمكن تفعيل الربط لأن الصنف المحلي غير نشط أو غير موجود.')
        if target and str(row['supplier_item_code'] or '').strip():
            others=_active_code_matches(c,row['supplier_id'],row['supplier_item_code'],mid)
            if any(int(r['local_item_id'])!=int(row['local_item_id']) for r in others):
                _record_supplier_mapping_conflict(c,row['supplier_id'],row['supplier_item_code'],row['local_item_id'],int(others[0]['local_item_id']),reason='محاولة تفعيل ربط متعارض')
                raise ValueError('لا يمكن تفعيل الربط لأن كود المورد مستخدم حاليًا لصنف محلي آخر.')
        status='APPROVED' if target else 'DISABLED'
        c.execute("UPDATE supplier_item_mappings SET status=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",(status,mid))
        after=dict(c.execute("SELECT * FROM supplier_item_mappings WHERE id=?",(mid,)).fetchone())
        log_learning_history(c,'MAPPING',mid,row['supplier_id'],row['local_item_id'],'ENABLE' if target else 'DISABLE',before,after,source='USER')
        c.commit(); return after
    except: c.rollback(); raise
    finally: c.close()


def update_learning_mapping(payload):
    mid=int(payload.get('id') or 0); local_item_id=int(payload.get('local_item_id') or 0)
    code=str(payload.get('supplier_item_code') or '').strip() or None; name=str(payload.get('supplier_item_name') or '').strip(); unit=str(payload.get('supplier_unit') or '').strip() or None
    if not mid or not local_item_id or not name: raise ValueError('معرف الربط والصنف واسم المورد مطلوبة.')
    c=dbconn()
    try:
        row=c.execute("SELECT * FROM supplier_item_mappings WHERE id=?",(mid,)).fetchone()
        if not row: raise ValueError('الربط غير موجود.')
        if not _active_local_item(c,local_item_id): raise ValueError('الصنف المحلي المحدد غير موجود أو غير نشط.')
        before=dict(row)
        if code:
            others=_active_code_matches(c,row['supplier_id'],code,mid)
            if any(int(r['local_item_id'])!=local_item_id for r in others):
                _record_supplier_mapping_conflict(c,row['supplier_id'],code,local_item_id,int(others[0]['local_item_id']),reason='تعديل ربط إلى كود مورد متعارض')
                raise ValueError('كود المورد مستخدم لصنف محلي آخر؛ عالج التعارض أولًا.')
        c.execute("""UPDATE supplier_item_mappings SET local_item_id=?,supplier_item_code=?,supplier_item_name=?,normalized_name=?,supplier_unit=?,normalized_unit=?,updated_at=CURRENT_TIMESTAMP WHERE id=?""",
                  (local_item_id,code,name,norm(name),unit,norm(unit or ''),mid))
        after=dict(c.execute("SELECT * FROM supplier_item_mappings WHERE id=?",(mid,)).fetchone())
        # Move only aliases explicitly attached to this mapping; keep unrelated aliases untouched.
        if int(row['local_item_id'])!=local_item_id:
            aliases_rows=c.execute("SELECT * FROM supplier_item_aliases WHERE mapping_id=?",(mid,)).fetchall()
            for a in aliases_rows:
                olda=dict(a); c.execute("UPDATE supplier_item_aliases SET local_item_id=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",(local_item_id,a['id']))
                newa=dict(c.execute("SELECT * FROM supplier_item_aliases WHERE id=?",(a['id'],)).fetchone())
                log_learning_history(c,'ALIAS',a['id'],row['supplier_id'],local_item_id,'REASSIGN',olda,newa,source='USER',notes=f'نقل بسبب نقل الربط {mid}')
        log_learning_history(c,'MAPPING',mid,row['supplier_id'],local_item_id,'UPDATE',before,after,source='USER')
        c.commit(); return after
    except: c.rollback(); raise
    finally: c.close()


def _resolve_learning_conflict(conflict_id, note=''):
    cid=int(conflict_id or 0)
    if not cid: raise ValueError('معرف التعارض مطلوب.')
    c=dbconn()
    try:
        cols={r['name'] for r in c.execute("PRAGMA table_info(supplier_mapping_conflicts)").fetchall()}
        row=c.execute("SELECT * FROM supplier_mapping_conflicts WHERE id=?",(cid,)).fetchone()
        if not row: raise ValueError('سجل التعارض غير موجود.')
        before=dict(row); note=str(note or '').strip() or 'تمت المراجعة يدويًا.'
        if {'resolved_at','resolution_note'} <= cols:
            c.execute("UPDATE supplier_mapping_conflicts SET status='RESOLVED',resolved_at=CURRENT_TIMESTAMP,resolution_note=? WHERE id=?",(note,cid))
        else:
            c.execute("UPDATE supplier_mapping_conflicts SET status='RESOLVED',resolved_at=CURRENT_TIMESTAMP WHERE id=?",(cid,))
        after=dict(c.execute("SELECT * FROM supplier_mapping_conflicts WHERE id=?",(cid,)).fetchone())
        log_learning_history(c,'CONFLICT',cid,row['supplier_id'],row['attempted_local_item_id'] if 'attempted_local_item_id' in row.keys() else row['local_item_id'],'RESOLVE',before,after,source='USER',notes=note)
        c.commit(); return after
    except: c.rollback(); raise
    finally: c.close()


def _active_code_matches(c, supplier_id, code, exclude_mapping_id=None):
    key=normalize_supplier_code(code)
    if not supplier_id or not key: return []
    rows=c.execute("""SELECT m.id,m.local_item_id,m.supplier_item_code,m.status
                      FROM supplier_item_mappings m
                      WHERE m.supplier_id=? AND m.status!='DISABLED'
                        AND m.supplier_item_code IS NOT NULL AND trim(m.supplier_item_code)<>''""",(int(supplier_id),)).fetchall()
    out=[]
    for r in rows:
        if exclude_mapping_id is not None and int(r['id'])==int(exclude_mapping_id): continue
        if normalize_supplier_code(r['supplier_item_code'])==key: out.append(r)
    return out

def _active_alias_matches(c, supplier_id, alias, exclude_alias_id=None):
    n=norm(alias)
    if not supplier_id or not n: return []
    rows=c.execute("""SELECT id,local_item_id,mapping_id,alias_name,usage_count,active
                      FROM supplier_item_aliases
                      WHERE supplier_id=? AND active=1 AND normalized_name=?""",(int(supplier_id),n)).fetchall()
    if exclude_alias_id is not None:
        rows=[r for r in rows if int(r['id'])!=int(exclude_alias_id)]
    return rows

def _learn_alias(c, supplier_id, local_item_id, alias, mapping_id=None, source='USER',
                 invoice_id=None, invoice_line_id=None, note=''):
    """Create/reactivate/update one supplier-specific alias safely."""
    alias=str(alias or '').strip()
    if not supplier_id or not local_item_id or not alias: return None,None
    n=norm(alias)
    matches=_active_alias_matches(c,supplier_id,alias)
    other=[r for r in matches if int(r['local_item_id'])!=int(local_item_id)]
    if other:
        _record_supplier_mapping_conflict(c,supplier_id,alias,local_item_id,int(other[0]['local_item_id']),invoice_id,invoice_line_id,
                                           'اسم المورد/الاسم البديل مرتبط بصنف محلي آخر')
        return None,'الاسم الوارد مرتبط حاليًا بصنف محلي آخر؛ لم يتم تغيير التعلم لهذا الاسم.'
    row=c.execute("SELECT id,mapping_id,usage_count,active FROM supplier_item_aliases WHERE supplier_id=? AND local_item_id=? AND normalized_name=? LIMIT 1",
                  (supplier_id,local_item_id,n)).fetchone()
    if row:
        aid=int(row['id'])
        c.execute("""UPDATE supplier_item_aliases SET mapping_id=COALESCE(?,mapping_id),active=1,
                   usage_count=usage_count+1,last_seen_at=CURRENT_TIMESTAMP,updated_at=CURRENT_TIMESTAMP,
                   source=? WHERE id=?""",(mapping_id,source,aid))
        return aid,None
    cur=c.execute("""INSERT INTO supplier_item_aliases
                    (supplier_id,local_item_id,mapping_id,alias_name,normalized_name,source,usage_count,first_seen_at,last_seen_at,active)
                    VALUES(?,?,?,?,?,?,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP,1)""",
                  (supplier_id,local_item_id,mapping_id,alias,n,source))
    return cur.lastrowid,None

def _learn_supplier_code(c, supplier_id, local_item_id, code, supplier_name, unit,
                         invoice_id=None, invoice_line_id=None, score=100, method='MANUAL'):
    """Create/update a supplier-code mapping without taking ownership from another item."""
    code=str(code or '').strip()
    if not code: return None,None
    key=normalize_supplier_code(code)
    if not key: return None,None
    rows=c.execute("""SELECT id,local_item_id,supplier_item_code,status
                      FROM supplier_item_mappings
                      WHERE supplier_id=? AND supplier_item_code IS NOT NULL AND trim(supplier_item_code)<>''""",(supplier_id,)).fetchall()
    same=[]; others=[]
    for r in rows:
        if normalize_supplier_code(r['supplier_item_code'])!=key: continue
        if int(r['local_item_id'])==int(local_item_id): same.append(r)
        else: others.append(r)
    if others:
        _record_supplier_mapping_conflict(c,supplier_id,code,local_item_id,int(others[0]['local_item_id']),invoice_id,invoice_line_id,
                                           'كود المورد مستخدم/محفوظ لصنف محلي آخر؛ لم يتم تبديل الملكية')
        return None,'كود المورد مرتبط بصنف محلي آخر؛ تم حفظ الفاتورة دون تغيير ملكية الكود.'
    if same:
        # Prefer an existing exact/raw-code row. Because supplier code is unique in the
        # original schema, reactivating is safer than inserting a duplicate beside a disabled row.
        target=sorted(same,key=lambda r:(0 if str(r['supplier_item_code']).strip()==code else 1,0 if r['status']!='DISABLED' else 1,int(r['id'])))[0]
        mid=int(target['id'])
        c.execute("""UPDATE supplier_item_mappings SET supplier_item_code=?,supplier_item_name=?,normalized_name=?,
                   supplier_unit=?,normalized_unit=?,confidence_score=?,match_method=?,status='APPROVED',
                   approved_by='USER',approved_at=COALESCE(approved_at,CURRENT_TIMESTAMP),usage_count=usage_count+1,
                   last_seen_at=CURRENT_TIMESTAMP,updated_at=CURRENT_TIMESTAMP WHERE id=?""",
                  (code,supplier_name or 'صنف مورد',norm(supplier_name or ''),unit or None,norm(unit or ''),score or 100,method or 'MANUAL',mid))
        return mid,None
    cur=c.execute("""INSERT INTO supplier_item_mappings
                     (supplier_id,local_item_id,supplier_item_code,supplier_item_name,normalized_name,supplier_unit,
                      normalized_unit,conversion_factor,confidence_score,match_method,status,approved_by,approved_at,
                      first_seen_at,last_seen_at,usage_count)
                     VALUES(?,?,?,?,?,?,?,1,?,?,?,?,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP,1)""",
                  (supplier_id,local_item_id,code,supplier_name or 'صنف مورد',norm(supplier_name or ''),unit or None,norm(unit or ''),score or 100,method or 'MANUAL','APPROVED','USER'))
    return cur.lastrowid,None

def _apply_learning_for_line(c, supplier_id, local_item_id, line, invoice_id=None, invoice_line_id=None,
                             mapping_id=None, source='USER', score=100, method='MANUAL',
                             learn_code=True, learn_alias=True):
    """Persist final user-confirmed learning. Called only inside invoice save/update."""
    if not local_item_id: return {'mapping_id':mapping_id,'alias_id':None,'warnings':[],'learned':False}
    if not _active_local_item(c,local_item_id):
        return {'mapping_id':mapping_id,'alias_id':None,'warnings':['الصنف المحلي المحدد غير نشط؛ لم يتم حفظ التعلم.'],'learned':False}
    warnings=[]; learned=False
    code=str(line.get('supplier_item_code') or '').strip()
    name=str(line.get('raw_item_name') or '').strip()
    unit=str(line.get('unit') or '').strip()
    mid=mapping_id
    if learn_code and code:
        mid2,warning=_learn_supplier_code(c,supplier_id,local_item_id,code,name,unit,invoice_id,invoice_line_id,score,method)
        if mid2: mid=mid2; learned=True
        if warning: warnings.append(warning)
    alias_id=None
    if learn_alias and name:
        alias_id,warning=_learn_alias(c,supplier_id,local_item_id,name,mid,source,invoice_id,invoice_line_id)
        if alias_id: learned=True
        if warning: warnings.append(warning)
    return {'mapping_id':mid,'alias_id':alias_id,'warnings':warnings,'learned':learned}


def _valid_local_item_for_context(value):
    try: return int(value) if value not in (None,'') else None
    except Exception: return None

def _valid_mapping(c, supplier_id, mid, expected_item_id=None):
    try: mid=int(mid or 0)
    except Exception: return None
    if not mid: return None
    r=c.execute("SELECT id,local_item_id FROM supplier_item_mappings WHERE id=? AND supplier_id=? AND status!='DISABLED' LIMIT 1",(mid,supplier_id)).fetchone()
    if not r: return None
    if expected_item_id is not None and int(r['local_item_id'])!=int(expected_item_id): return None
    return int(r['id'])

def _match_one(c, line, all_items=None, context=None, sequence=None):
    """Read-only matching. Supplier code is strongest when present; unresolved codes never auto-match."""
    supplier_id=line.get('supplier_id')
    code=str(line.get('supplier_item_code') or '').strip()
    name=str(line.get('raw_item_name') or '').strip()
    unit=str(line.get('unit') or '').strip()
    if context is None: context=_build_match_context(c,supplier_id,all_items)
    all_items=context['all_items']; mappings=context['mappings']; aliases=context['aliases']; prices=context['prices']
    code_present=bool(supplier_id and code)
    code_resolved=False; code_conflict=False; code_evidence=None

    if code_present:
        exact=[m for m in mappings if str(m['supplier_item_code'] or '').strip().lower()==code.lower()]
        exact_local={int(r['local_item_id']) for r in exact}
        if len(exact_local)==1 and exact:
            r=sorted(exact,key=lambda x:(int(x['usage_count'] or 0),int(x['mapping_id'])),reverse=True)[0]
            return {'status':'AUTO_MATCHED','score':100,'method':'SUPPLIER_CODE','item':dict(r),'mapping_id':r['mapping_id'],'suggestions':[],
                    'evidence':{'supplier_code':'exact'}}
        if len(exact_local)>1: code_conflict=True; code_evidence='exact_conflict'
        else:
            for key_kind,key in _code_keys(code,sequence):
                normalized=[m for m in mappings if normalize_supplier_code(m['supplier_item_code'])==key]
                local_ids={int(r['local_item_id']) for r in normalized}
                if len(local_ids)==1 and normalized:
                    r=sorted(normalized,key=lambda x:(int(x['usage_count'] or 0),int(x['mapping_id'])),reverse=True)[0]
                    score=99 if key_kind=='FULL' else (98 if key_kind=='SEQUENCE_SUFFIX' else 97)
                    method='SUPPLIER_CODE_NORMALIZED' if key_kind=='FULL' else ('SUPPLIER_CODE_SEQUENCE_CLEANED' if key_kind=='SEQUENCE_SUFFIX' else 'SUPPLIER_CODE_PARENS_CLEANED')
                    return {'status':'AUTO_MATCHED','score':score,'method':method,'item':dict(r),'mapping_id':r['mapping_id'],'suggestions':[],
                            'evidence':{'supplier_code':key_kind.lower()}}
                if len(local_ids)>1:
                    code_conflict=True; code_evidence=f'{key_kind.lower()}_conflict'; break
            if not code_resolved and not code_conflict: code_evidence='unknown'

    nname=norm(name)
    # Exact learned names/aliases auto-match only when the invoice has no supplier code.
    if supplier_id and nname and not code_present:
        exact=[m for m in mappings if m['normalized_name']==nname or norm(m['supplier_item_name'])==nname]
        exact_items={int(r['local_item_id']) for r in exact}
        if len(exact_items)==1 and exact:
            r=sorted(exact,key=lambda x:(int(x['usage_count'] or 0),int(x['mapping_id'])),reverse=True)[0]
            return {'status':'AUTO_MATCHED','score':99,'method':'EXACT_SUPPLIER_NAME','item':dict(r),'mapping_id':r['mapping_id'],'suggestions':[],
                    'evidence':{'supplier_name':'exact'}}
        exact_alias=[a for a in aliases if a['normalized_name']==nname and int(a['active'] if 'active' in a.keys() else 1)==1]
        alias_items={int(a['local_item_id']) for a in exact_alias}
        if len(alias_items)==1 and exact_alias:
            r=sorted(exact_alias,key=lambda x:(int(x['usage_count'] or 0),int(x['id'])),reverse=True)[0]
            return {'status':'AUTO_MATCHED','score':98,'method':'ALIAS','item':dict(r),'mapping_id':r['mapping_id'],'suggestions':[],
                    'evidence':{'alias':'exact'}}

    candidates={}; base_line={'raw_item_name':name,'unit':unit,'unit_price':line.get('unit_price')}
    for r in mappings:
        iid=int(r['local_item_id']); score,ev=_candidate_score(base_line,dict(r),'SUPPLIER_MAPPING_NAME',prices.get(iid,[]))
        if iid not in candidates or score>candidates[iid]['score']:
            candidates[iid]={'item':dict(r),'score':score,'method':'SUPPLIER_MAPPING_NAME','mapping_id':r['mapping_id'],'evidence':ev}
    for r in aliases:
        if int(r['active'] if 'active' in r.keys() else 1)!=1: continue
        iid=int(r['local_item_id']); score,ev=_candidate_score(base_line,dict(r),'ALIAS_SUGGESTION',prices.get(iid,[]))
        if iid not in candidates or score>candidates[iid]['score']:
            candidates[iid]={'item':dict(r),'score':score,'method':'ALIAS_SUGGESTION','mapping_id':r['mapping_id'],'evidence':ev}
    for r in all_items:
        rd=dict(r); iid=int(r['id']); score,ev=_candidate_score(base_line,rd,'LOCAL_NAME',prices.get(iid,[]))
        if iid not in candidates or score>candidates[iid]['score']:
            candidates[iid]={'item':rd,'score':score,'method':'LOCAL_NAME','mapping_id':None,'evidence':ev}
    ranked=sorted(candidates.values(),key=lambda x:(x['score'],int(x['item'].get('usage_count') or 0)),reverse=True)
    suggestions=[]
    for r in ranked[:5]:
        if r['score']<45: continue
        it=r['item']; ev=dict(r['evidence'])
        if code_present: ev['supplier_code']=code_evidence
        suggestions.append({'item':{k:it[k] for k in ('id','item_code','item_name','main_unit') if k in it},'score':round(r['score']),
                             'method':r['method'],'mapping_id':r['mapping_id'],'evidence':ev})
    if code_present and nname:
        exact_by_item={}
        for r in mappings:
            if r['normalized_name']==nname or norm(r['supplier_item_name'])==nname:
                iid=int(r['local_item_id']); exact_by_item[iid]={'item':{k:r[k] for k in ('id','item_code','item_name','main_unit') if k in r},'score':96,'method':'EXACT_NAME_REVIEW','mapping_id':r.get('mapping_id'),'evidence':{'supplier_code':code_evidence or 'unknown','supplier_name':'exact'}}
        for r in aliases:
            if int(r['active'] if 'active' in r.keys() else 1)==1 and r['normalized_name']==nname:
                iid=int(r['local_item_id']); exact_by_item[iid]={'item':{k:r[k] for k in ('id','item_code','item_name','main_unit') if k in r},'score':96,'method':'EXACT_NAME_REVIEW','mapping_id':r.get('mapping_id'),'evidence':{'supplier_code':code_evidence or 'unknown','supplier_name':'exact'}}
        existing_ids={int(z['item']['id']) for z in suggestions}
        for z in exact_by_item.values():
            if int(z['item']['id']) not in existing_ids: suggestions.insert(0,z)
        suggestions=suggestions[:5]
    best=ranked[0] if ranked else None
    if not best or best['score']<80:
        return {'status':'NEW_ITEM','score':round(best['score'] if best else 0),'method':'NEW_ITEM','item':None,'mapping_id':None,
                'suggestions':suggestions,'evidence':(dict(best['evidence']) if best else {'supplier_code':code_evidence})}
    if code_present and not code_resolved:
        return {'status':'SUGGESTED','score':round(best['score']),'method':('SUPPLIER_CODE_CONFLICT_REVIEW' if code_conflict else 'SUPPLIER_CODE_UNKNOWN_REVIEW'),
                'item':best['item'],'mapping_id':best['mapping_id'],'suggestions':suggestions,'evidence':{**dict(best['evidence']),'supplier_code':code_evidence}}
    return {'status':'SUGGESTED','score':round(best['score']),'method':best['method'],'item':best['item'],'mapping_id':best['mapping_id'],
            'suggestions':suggestions,'evidence':best['evidence']}

def rematch_invoice(invoice_id):
    """Re-evaluate a saved reviewable invoice against the current learning state.

    No invoice row is modified here; the caller can review the returned matches and then
    use the normal save/update action. Approved invoices are also allowed to be re-evaluated
    for review purposes, but the operation remains read-only.
    """
    invoice_id=int(invoice_id or 0)
    c=dbconn()
    try:
        inv=c.execute("SELECT i.*,s.supplier_name FROM invoices i JOIN suppliers s ON s.id=i.supplier_id WHERE i.id=?",(invoice_id,)).fetchone()
        if not inv: raise ValueError('الفاتورة غير موجودة.')
        rows=c.execute("SELECT * FROM invoice_lines WHERE invoice_id=? ORDER BY line_number",(invoice_id,)).fetchall()
        lines=[]
        for r in rows:
            lines.append({
                'id':r['id'],'line_number':r['line_number'],'supplier_id':inv['supplier_id'],'supplier_name':inv['supplier_name'],
                'supplier_item_code':r['supplier_item_code'] or '', 'raw_item_name':r['raw_item_name'] or '',
                'qty':float(r['quantity'] or 0), 'unit':r['supplier_unit'] or '',
                'unit_price':float(r['unit_price_halalah'] or 0)/100,
                'net_amount':float(r['net_amount_halalah'] or 0)/100,
                'vat_amount':float(r['vat_amount_halalah'] or 0)/100,
                'gross_amount':float(r['gross_amount_halalah'] or 0)/100,
                'match':{'status':r['match_status'],'score':r['match_score'] or 0,
                         'item':({'id':r['matched_local_item_id']} if r['matched_local_item_id'] else None),
                         'mapping_id':r['matched_mapping_id']}
            })
        return {'invoice':dict(inv),'lines':match_lines_bulk(lines,inv['supplier_id'])}
    finally: c.close()

def match_lines_bulk(lines, supplier_id=None):
    if not lines: return []
    c=dbconn()
    try:
        context=_build_match_context(c,supplier_id)
        out=[]
        for idx,x in enumerate(lines,1):
            r=dict(x)
            if supplier_id: r['supplier_id']=supplier_id
            # Preserve an explicit source line number when present; otherwise the row order
            # is used only to detect OCR/PDF codes polluted by the line sequence.
            seq=r.get('line_number') or idx
            r['match']=_match_one(c,r,context=context,sequence=seq)
            out.append(r)
        return out
    finally: c.close()

def match_line(line):
    return match_lines_bulk([line])[0]

def suggest_group_for_name(name):
    """Suggest an existing group using token similarity; never creates a group automatically."""
    c=dbconn()
    rows=c.execute("SELECT group_code, COUNT(*) n FROM local_items WHERE active=1 AND group_code IS NOT NULL AND trim(group_code)<>'' GROUP BY group_code ORDER BY group_code").fetchall()
    candidates=[]
    for r in rows:
        gc=str(r['group_code'])
        names=c.execute("SELECT item_name FROM local_items WHERE active=1 AND group_code=?",(gc,)).fetchall()
        scores=[jaccard(name,x['item_name']) for x in names]
        score=max(scores) if scores else 0
        candidates.append((gc,score,r['n']))
    c.close()
    candidates.sort(key=lambda x:(x[1],x[2]),reverse=True)
    if not candidates:
        return {"group_code":"","score":0,"reason":"NO_GROUPS"}
    gc,score,n=candidates[0]
    return {"group_code":gc,"score":round(score),"items_count":n,"reason":"NAME_SIMILARITY"}

def next_item_code(group_code):
    """Generate the next 3-digit sequence inside a numeric group prefix."""
    gc=str(group_code or '').strip()
    if not gc or not re.fullmatch(r'\d+',gc):
        raise ValueError('كود المجموعة يجب أن يكون رقميًا.')
    c=dbconn()
    try:
        rows=c.execute("SELECT item_code FROM local_items WHERE item_code GLOB '[0-9]*'").fetchall()
    finally:
        c.close()
    prefix=gc
    used=set()
    for r in rows:
        code=str(r['item_code'] or '').strip()
        if code.startswith(prefix) and code.isdigit() and len(code)==len(prefix)+3:
            used.add(int(code[len(prefix):]))
    for suffix in range(1,1000):
        if suffix not in used:
            return prefix+f'{suffix:03d}'
    raise ValueError(f'المجموعة {gc} وصلت إلى الحد الأقصى للترقيم (999 صنفًا). أنشئ مجموعة جديدة أو راجع سياسة الترقيم.')

def groups():
    c=dbconn()
    rows=c.execute("SELECT group_code, COUNT(*) item_count, MIN(item_code) first_code, MAX(item_code) last_code FROM local_items WHERE active=1 AND group_code IS NOT NULL AND trim(group_code)<>'' GROUP BY group_code ORDER BY group_code").fetchall()
    c.close()
    return [dict(r) for r in rows]

def create_local_item(payload):
    name=str(payload.get('item_name') or '').strip()
    group_code=str(payload.get('group_code') or '').strip()
    main_unit=str(payload.get('main_unit') or '').strip()
    item_type=str(payload.get('item_type') or '').strip() or 'سلعي'
    barcode=str(payload.get('barcode') or '').strip() or None
    pack=payload.get('pack_size')
    factor=payload.get('conversion_factor', pack if pack not in (None,'') else 1)
    try: factor=float(factor)
    except: factor=1
    if not name: raise ValueError('اسم الصنف مطلوب.')
    if not main_unit: raise ValueError('الوحدة الأساسية مطلوبة.')
    if not group_code: raise ValueError('اختر أو اكتب كود المجموعة.')
    if not re.fullmatch(r'\d+',group_code): raise ValueError('كود المجموعة يجب أن يكون رقميًا.')
    if factor<=0: raise ValueError('العبوة/معامل التحويل يجب أن يكون أكبر من صفر.')
    code=str(payload.get('item_code') or '').strip() or next_item_code(group_code)
    if not code.isdigit(): raise ValueError('رقم الصنف يجب أن يكون رقميًا.')
    c=dbconn()
    try:
        ex=c.execute("SELECT id,item_code,item_name FROM local_items WHERE item_code=? LIMIT 1",(code,)).fetchone()
        if ex: raise ValueError(f"رقم الصنف {code} مستخدم بالفعل للصنف: {ex['item_name']}")
        exn=c.execute("SELECT id,item_code,item_name FROM local_items WHERE normalized_name=? AND active=1 LIMIT 1",(norm(name),)).fetchone()
        if exn: raise ValueError(f"يوجد صنف محلي بنفس الاسم بالفعل: {exn['item_code']} — {exn['item_name']}")
        cur=c.execute("""INSERT INTO local_items(item_code,group_code,item_name,normalized_name,item_type,main_unit,active,source,export_status)
                         VALUES(?,?,?,?,?,?,1,'USER_CREATED','NEW')""",
                      (code,group_code,name,norm(name),item_type,main_unit))
        iid=cur.lastrowid
        curu=c.execute("""INSERT INTO item_units(local_item_id,unit_name,normalized_unit,conversion_factor,barcode,is_default)
                     VALUES(?,?,?,?,?,1)""",(iid,main_unit,norm(main_unit),1,barcode))
        unit_ids={norm(main_unit):curu.lastrowid}
        extras=payload.get('extra_units') or []
        if not isinstance(extras,list): raise ValueError('قائمة الوحدات الإضافية غير صالحة.')
        for u in extras:
            uname=str((u or {}).get('unit_name') or '').strip()
            if not uname: continue
            if norm(uname) in unit_ids: raise ValueError(f'الوحدة مكررة: {uname}')
            try: uf=float((u or {}).get('conversion_factor') or 1)
            except Exception: raise ValueError(f'معامل الوحدة غير صالح: {uname}')
            if uf<=0: raise ValueError(f'معامل الوحدة يجب أن يكون أكبر من صفر: {uname}')
            ub=str((u or {}).get('barcode') or '').strip() or None
            curu=c.execute("INSERT INTO item_units(local_item_id,unit_name,normalized_unit,conversion_factor,barcode,is_default,active) VALUES(?,?,?,?,?,0,1)",(iid,uname,norm(uname),uf,ub))
            unit_ids[norm(uname)]=curu.lastrowid
        def resolve_new_default(value,label):
            if value in (None,''): return None
            v=str(value).strip()
            uid=unit_ids.get(norm(v))
            if not uid: raise ValueError(f'{label} يجب أن تكون إحدى وحدات الصنف.')
            return uid
        sale_value=payload.get('default_sale_unit_name')
        purchase_value=payload.get('default_purchase_unit_name')
        sale_uid=resolve_new_default(sale_value,'وحدة البيع الافتراضية')
        purchase_uid=resolve_new_default(purchase_value,'وحدة الشراء الافتراضية')
        c.execute("UPDATE local_items SET default_sale_unit_id=?,default_purchase_unit_id=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",(sale_uid,purchase_uid,iid))
        c.execute("INSERT OR IGNORE INTO item_groups(group_code,group_name) VALUES(?,?)",(group_code,''))
        c.commit()
        row=c.execute("SELECT id,item_code,group_code,item_name,item_type,main_unit,default_sale_unit_id,default_purchase_unit_id FROM local_items WHERE id=?",(iid,)).fetchone()
        return dict(row)
    except:
        c.rollback(); raise
    finally: c.close()

def save_new_item_from_invoice(payload):
    """Create the local item now; defer supplier mapping/alias learning to invoice save."""
    supplier_id=int(payload.get('supplier_id') or 0); line=payload.get('line') or {}
    if not supplier_id: raise ValueError('المورد مطلوب.')
    item=create_local_item(payload)
    return {"item":item,"mapping_id":None,"pending_learning":{
        'supplier_id':supplier_id,'local_item_id':int(item['id']),'supplier_item_code':str(line.get('supplier_item_code') or '').strip(),
        'alias_name':str(line.get('raw_item_name') or '').strip(),'unit':str(line.get('unit') or '').strip(),'source':'USER'
    }}

def export_mappings_xlsx():
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from io import BytesIO
    c=dbconn()
    mapping_rows=c.execute("""SELECT m.id,m.supplier_id,m.local_item_id,m.supplier_item_code,m.supplier_item_name,
                              m.supplier_unit,m.conversion_factor,m.confidence_score,m.match_method,m.status,m.usage_count,
                              s.supplier_code,s.supplier_name,s.vat_number,s.commercial_register,
                              li.item_code,li.item_name,li.main_unit,
                              (SELECT COUNT(*) FROM supplier_item_aliases a WHERE a.supplier_id=m.supplier_id AND a.local_item_id=m.local_item_id) alias_count,
                              (SELECT GROUP_CONCAT(a.alias_name,' | ') FROM supplier_item_aliases a WHERE a.supplier_id=m.supplier_id AND a.local_item_id=m.local_item_id) aliases
                       FROM supplier_item_mappings m
                       JOIN suppliers s ON s.id=m.supplier_id
                       JOIN local_items li ON li.id=m.local_item_id
                       ORDER BY s.supplier_name,li.item_code,m.supplier_item_code,m.supplier_item_name""").fetchall()
    alias_rows=c.execute("""SELECT a.id,s.supplier_code,s.supplier_name,li.item_code,li.item_name,
                                  a.alias_name,a.source,a.usage_count,a.first_seen_at,a.last_seen_at
                           FROM supplier_item_aliases a
                           JOIN suppliers s ON s.id=a.supplier_id
                           JOIN local_items li ON li.id=a.local_item_id
                           ORDER BY s.supplier_name,li.item_code,a.alias_name""").fetchall()
    supplier_rows=c.execute("SELECT supplier_code,supplier_name,vat_number,commercial_register,phone,active FROM suppliers ORDER BY supplier_name").fetchall()
    item_rows=c.execute("SELECT item_code,group_code,item_name,main_unit,item_type,active FROM local_items ORDER BY item_code").fetchall()
    c.close()

    wb=Workbook()
    ws=wb.active; ws.title='جدول الربط'
    headers=['معرف الربط','كود المورد','المورد','كود الصنف المحلي','الصنف المحلي','كود الصنف عند المورد','اسم الصنف عند المورد','الوحدة عند المورد','معامل التحويل','عدد الأسماء البديلة','الأسماء البديلة','الثقة','طريقة المطابقة','الحالة','عدد الاستخدامات']
    ws.append(headers)
    for cell in ws[1]: cell.font=Font(bold=True)
    for r in mapping_rows:
        ws.append([r['id'],r['supplier_code'] or '',r['supplier_name'],r['item_code'],r['item_name'],r['supplier_item_code'] or '',r['supplier_item_name'] or '',r['supplier_unit'] or '',r['conversion_factor'] or '',r['alias_count'] or 0,r['aliases'] or '',r['confidence_score'] or 0,r['match_method'] or '',r['status'] or '',r['usage_count'] or 0])

    wa=wb.create_sheet('الأسماء البديلة')
    ah=['معرف الاسم','كود المورد','المورد','كود الصنف المحلي','الصنف المحلي','الاسم البديل','المصدر','عدد الاستخدامات','أول ظهور','آخر ظهور']
    wa.append(ah)
    for cell in wa[1]: cell.font=Font(bold=True)
    for r in alias_rows:
        wa.append([r['id'],r['supplier_code'] or '',r['supplier_name'],r['item_code'],r['item_name'],r['alias_name'],r['source'],r['usage_count'] or 0,r['first_seen_at'] or '',r['last_seen_at'] or ''])

    wsu=wb.create_sheet('الموردون')
    sh=['كود المورد','اسم المورد','الرقم الضريبي','السجل التجاري','الهاتف','نشط']
    wsu.append(sh)
    for cell in wsu[1]: cell.font=Font(bold=True)
    for r in supplier_rows: wsu.append([r['supplier_code'] or '',r['supplier_name'],r['vat_number'] or '',r['commercial_register'] or '',r['phone'] or '',r['active']])

    wi=wb.create_sheet('الأصناف المحلية')
    ih=['كود الصنف','كود المجموعة','اسم الصنف','الوحدة الرئيسية','نوع الصنف','نشط']
    wi.append(ih)
    for cell in wi[1]: cell.font=Font(bold=True)
    for r in item_rows: wi.append([r['item_code'],r['group_code'] or '',r['item_name'],r['main_unit'] or '',r['item_type'] or '',r['active']])

    for sheet in wb.worksheets:
        sheet.freeze_panes='A2'
        sheet.auto_filter.ref=sheet.dimensions
        for col in sheet.columns:
            maxlen=min(max(len(str(x.value or '')) for x in col)+2,45)
            sheet.column_dimensions[col[0].column_letter].width=maxlen
        for row in sheet.iter_rows():
            for cell in row: cell.alignment=Alignment(vertical='top',wrap_text=True)
    bio=BytesIO(); wb.save(bio); return bio.getvalue()



def supplier_snapshot(c, sid):
    r=c.execute("SELECT id,supplier_code,supplier_name,normalized_name,vat_number,commercial_register,phone,active,created_at,updated_at FROM suppliers WHERE id=?",(sid,)).fetchone()
    return dict(r) if r else None

def log_supplier_change(c, sid, action, before, after):
    b=json.dumps(before,ensure_ascii=False,sort_keys=True) if before else None
    a=json.dumps(after,ensure_ascii=False,sort_keys=True) if after else None
    changed=[]
    if before and after:
        changed=[k for k in set(before)|set(after) if before.get(k)!=after.get(k)]
    c.execute("INSERT INTO supplier_change_history(entity_id,action,before_json,after_json,changed_fields) VALUES(?,?,?,?,?)",
              (sid,action,b,a,', '.join(sorted(changed))))

def update_supplier(payload):
    sid=int(payload.get('id') or 0)
    name=str(payload.get('supplier_name') or '').strip()
    code=str(payload.get('supplier_code') or '').strip() or None
    vat=str(payload.get('vat_number') or '').strip() or None
    cr=str(payload.get('commercial_register') or payload.get('cr_number') or '').strip() or None
    phone=str(payload.get('phone') or '').strip() or None
    active=1 if bool(payload.get('active', True)) else 0
    if not sid or not name:
        raise ValueError('معرف المورد واسم المورد مطلوبان.')
    c=dbconn()
    try:
        before=supplier_snapshot(c,sid)
        if not before: raise ValueError('المورد غير موجود.')
        sn=norm(name)
        ex=c.execute("SELECT id,supplier_name FROM suppliers WHERE normalized_name=? AND id<>? LIMIT 1",(sn,sid)).fetchone()
        if ex: raise ValueError(f'اسم المورد مستخدم بالفعل للمورد: {ex[1]}.')
        if code:
            ex=c.execute("SELECT id,supplier_name FROM suppliers WHERE supplier_code=? AND id<>? LIMIT 1",(code,sid)).fetchone()
            if ex: raise ValueError(f'كود المورد مستخدم بالفعل للمورد: {ex[1]}.')
        c.execute("UPDATE suppliers SET supplier_code=?,supplier_name=?,normalized_name=?,vat_number=?,commercial_register=?,phone=?,active=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",
                  (code,name,sn,vat,cr,phone,active,sid))
        after=supplier_snapshot(c,sid)
        log_supplier_change(c,sid,'UPDATE',before,after)
        c.commit()
        return after
    except:
        c.rollback(); raise
    finally:
        c.close()


def _preferred_purchase_unit(units, purchase_unit_id=None, main_unit=None):
    """Choose the unit for simple item exports without guessing from conversion size.

    The explicit purchase default wins. If it is not set or no longer points to an
    active unit, fall back to the item's main/basic unit. Never infer the unit from
    the largest conversion factor.
    """
    if purchase_unit_id is not None:
        for u in units:
            if int(u.get('id') or 0) == int(purchase_unit_id) and u.get('active', 1):
                return u
    if main_unit:
        main_norm=norm(main_unit)
        for u in units:
            if u.get('active', 1) and norm(u.get('unit_name')) == main_norm:
                return u
    for u in units:
        if u.get('active', 1):
            return u
    return {'unit_name': main_unit or '', 'conversion_factor': 1}


def export_invoice_xls(invoice_id):
    """Export a saved purchase invoice in the exact Al-Mutakamil Plus import layout as Excel 97-2003 (.xls).
    All cells are written as TEXT.
    """
    import io, tempfile, subprocess, shutil, os
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils import get_column_letter
    invoice_id=int(invoice_id or 0)
    c=dbconn()
    try:
        inv=c.execute("""SELECT i.*, s.supplier_name FROM invoices i JOIN suppliers s ON s.id=i.supplier_id WHERE i.id=? AND i.active=1""",(invoice_id,)).fetchone()
        if not inv: raise ValueError('الفاتورة غير موجودة.')
        rows=c.execute("""SELECT il.*, li.item_code AS local_item_code, li.item_name AS local_item_name,
                                li.default_purchase_unit_id, pu.unit_name AS current_purchase_unit
                         FROM invoice_lines il LEFT JOIN local_items li ON li.id=il.matched_local_item_id
                         LEFT JOIN item_units pu ON pu.id=li.default_purchase_unit_id AND pu.active=1
                         WHERE il.invoice_id=? ORDER BY il.line_number""",(invoice_id,)).fetchall()
        missing=[(r['line_number'],r['local_item_code'] or r['raw_item_name']) for r in rows if not (r['export_unit_name'] or r['current_purchase_unit'])]
        if missing:
            labels=', '.join(f'{n}: {name}' for n,name in missing)
            raise ValueError('لا يمكن تصدير الفاتورة لأن وحدة الشراء الافتراضية غير محددة للبنود التالية: '+labels)
    finally:
        c.close()
    headers=['رقم الصنف لدينا','اسم الصنف','الوحدة','المخزن','الكمية','المجاني','تاريخ الانتهاء','الدفعة','التكلفة']
    wb=Workbook(); ws=wb.active; ws.title='الفاتورة'
    text_fmt='@'
    for col,h in enumerate(headers,1):
        cell=ws.cell(1,col,h); cell.number_format=text_fmt
    for r,row in enumerate(rows,2):
        vals=[
            row['local_item_code'] or '',
            row['local_item_name'] or row['raw_item_name'] or '',
            (row['export_unit_name'] if row['export_unit_name'] else row['current_purchase_unit']) or '',
            inv['warehouse_number'] or '',
            _text_number(row['quantity']),
            '0',
            '0',
            '0',
            _text_money(row['unit_price_halalah']),
        ]
        for col,val in enumerate(vals,1):
            cell=ws.cell(r,col,str(val) if val is not None else '')
            cell.number_format=text_fmt
            cell.data_type='s'
    ws.freeze_panes='A2'
    widths=[20,42,14,14,14,12,18,14,16]
    for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w
    with tempfile.TemporaryDirectory(prefix='mutakamil_invoice_') as td:
        td=Path(td); xlsx=td/'invoice.xlsx'; wb.save(xlsx)
        outdir=td/'out'; outdir.mkdir()
        xls=outdir/'invoice.xls'
        # Prefer Microsoft Excel via COM on Windows. This creates a real BIFF8
        # Excel 97-2003 workbook (FileFormat=56) and does not require LibreOffice.
        excel_cmd = shutil.which('powershell.exe') or shutil.which('powershell')
        converted = False
        excel_error = ''
        if excel_cmd:
            ps = td/'convert_xls.ps1'
            # OpenPyXL has already stored every cell as text. Excel SaveAs 56
            # converts the workbook to the native .xls BIFF8 format.
            script = r'''$ErrorActionPreference = "Stop"
$xlsx = [System.IO.Path]::GetFullPath($args[0])
$xls  = [System.IO.Path]::GetFullPath($args[1])
$excel = $null
$wb = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($xlsx, 0, $true)
  $wb.SaveAs($xls, 56)
  $wb.Close($false)
  $wb = $null
  $excel.Quit()
  $excel = $null
} finally {
  if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
  if ($excel -ne $null) { try { $excel.Quit() } catch {} }
}
'''
            ps.write_text(script, encoding='utf-8')
            try:
                proc=subprocess.run([excel_cmd,'-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',str(ps),str(xlsx),str(xls)],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,timeout=60)
                converted = proc.returncode==0 and xls.exists() and xls.stat().st_size>0
                if not converted:
                    excel_error=(proc.stderr.strip() or proc.stdout.strip())
            except Exception as ex:
                excel_error=str(ex)

        # LibreOffice remains a secondary option for machines without Excel.
        if not converted:
            soffice=shutil.which('soffice') or shutil.which('libreoffice')
            if soffice:
                cmd=[soffice,'--headless','--convert-to','xls:MS Excel 97','--outdir',str(outdir),str(xlsx)]
                proc=subprocess.run(cmd,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,timeout=60)
                converted = proc.returncode==0 and xls.exists() and xls.stat().st_size>0
                if not converted:
                    excel_error=proc.stderr.strip() or proc.stdout.strip() or excel_error

        if not converted:
            raise RuntimeError('تعذر إنشاء ملف Excel 97-2003. يلزم وجود Microsoft Excel مثبتًا على الجهاز (ويمكن استخدام LibreOffice كبديل).'+((' التفاصيل: '+excel_error) if excel_error else ''))
        return xls.read_bytes(), str(inv['invoice_number'])



def _text_number(v):
    try:
        n=float(v or 0)
        if n.is_integer(): return str(int(n))
        return ('%.6f'%n).rstrip('0').rstrip('.')
    except Exception: return str(v or '')

def _text_money(halalah):
    try:
        return ('%.2f'%(float(halalah or 0)/100)).rstrip('0').rstrip('.')
    except Exception: return str(halalah or '')


def _xls_biff8_bytes(workbook_builder, filename_prefix='items'):
    import tempfile, subprocess, shutil
    from pathlib import Path
    with tempfile.TemporaryDirectory(prefix='mutakamil_export_') as td:
        td=Path(td); xlsx=td/(filename_prefix+'.xlsx'); outdir=td/'out'; outdir.mkdir(); xls=outdir/(filename_prefix+'.xls')
        workbook_builder(xlsx)
        excel_cmd=shutil.which('powershell.exe') or shutil.which('powershell'); converted=False; err=''
        if excel_cmd:
            ps=td/'convert.ps1'
            ps.write_text('$ErrorActionPreference = "Stop"\n$xlsx = [System.IO.Path]::GetFullPath($args[0]); $xls = [System.IO.Path]::GetFullPath($args[1])\n$excel=$null; $wb=$null\ntry { $excel=New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false; $wb=$excel.Workbooks.Open($xlsx,0,$true); $wb.SaveAs($xls,56); $wb.Close($false); $wb=$null; $excel.Quit(); $excel=$null }\nfinally { if($wb -ne $null){try{$wb.Close($false)}catch{}}; if($excel -ne $null){try{$excel.Quit()}catch{}} }\n',encoding='utf-8')
            try:
                pr=subprocess.run([excel_cmd,'-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',str(ps),str(xlsx),str(xls)],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,timeout=60)
                converted=pr.returncode==0 and xls.exists() and xls.stat().st_size>0; err=pr.stderr.strip() or pr.stdout.strip()
            except Exception as ex: err=str(ex)
        if not converted:
            soffice=shutil.which('soffice') or shutil.which('libreoffice')
            if soffice:
                try:
                    pr=subprocess.run([soffice,'--headless','--convert-to','xls:MS Excel 97','--outdir',str(outdir),str(xlsx)],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,timeout=60)
                    converted=pr.returncode==0 and xls.exists() and xls.stat().st_size>0; err=pr.stderr.strip() or pr.stdout.strip() or err
                except Exception as ex: err=str(ex)
        if not converted: raise RuntimeError('تعذر إنشاء ملف Excel 97-2003. يلزم Microsoft Excel أو LibreOffice على الجهاز.'+((' التفاصيل: '+err) if err else ''))
        return xls.read_bytes()

def _write_text_xlsx(path, headers, rows, sheet='الأصناف'):
    from openpyxl import Workbook
    wb=Workbook(); ws=wb.active; ws.title=sheet; text_fmt='@'
    for c,h in enumerate(headers,1): ws.cell(1,c,str(h)).number_format=text_fmt
    for r_idx,row in enumerate(rows,2):
        for c_idx,val in enumerate(row,1):
            cell=ws.cell(r_idx,c_idx,'' if val is None else str(val)); cell.number_format=text_fmt; cell.data_type='s'
    ws.freeze_panes='A2'; ws.sheet_view.rightToLeft=True; wb.save(path)

def _item_export_rows(item_rows):
    rows=[]
    for item in item_rows:
        units=item['units'] or []
        sale_id=item.get('default_sale_unit_id')
        purchase_id=item.get('default_purchase_unit_id')
        if not units:
            raise ValueError(f"الصنف {item['item_code']} لا يحتوي على وحدات.")
        for u in units:
            uid=u.get('id')
            rows.append([item['group_code'] or '',item['item_code'],item['item_name'],'سلعي',u['unit_name'],_text_number(u['conversion_factor']),
                         '1' if sale_id and int(uid)==int(sale_id) else '',
                         '1' if purchase_id and int(uid)==int(purchase_id) else ''])
    return rows

def _get_export_items(status):
    c=dbconn()
    try:
        if status=='NEW_OR_EXPORTED':
            items=c.execute("SELECT * FROM local_items WHERE active=1 AND export_status IN ('NEW','EXPORTED') ORDER BY CAST(group_code AS INTEGER),CAST(item_code AS INTEGER),item_code").fetchall()
        else:
            items=c.execute("SELECT * FROM local_items WHERE active=1 AND export_status=? ORDER BY CAST(group_code AS INTEGER),CAST(item_code AS INTEGER),item_code",(status,)).fetchall()
        data=[]
        for it in items:
            units=[dict(u) for u in c.execute("SELECT id,unit_name,conversion_factor,is_default FROM item_units WHERE local_item_id=? AND active=1 ORDER BY id",(it['id'],)).fetchall()]
            data.append({'group_code':it['group_code'],'item_code':it['item_code'],'item_name':it['item_name'],'item_type':it['item_type'],'units':units,'id':it['id'],'export_status':it['export_status'],'default_sale_unit_id':it['default_sale_unit_id'],'default_purchase_unit_id':it['default_purchase_unit_id']})
        return data
    finally: c.close()

def _export_items_status(status, new_status, filename, sheet):
    data=_get_export_items(status); headers=['رقم المجموعة','رقم الصنف','الاسم المحلي','نوع الصنف','الوحدة','العبوة','وحدة البيع الافتراضية','وحدة الشراء الافتراضية']
    if not data:
        raise ValueError('لا توجد أصناف جاهزة للتصدير حاليًا.')
    blob=_xls_biff8_bytes(lambda path:_write_text_xlsx(path,headers,_item_export_rows(data),sheet),filename)
    # Do not remove the items from the export queue just because the file was downloaded.
    # The user must first import the file into Al-Mutakamil, then confirm the import.
    c=dbconn()
    try:
        c.executemany("UPDATE local_items SET last_exported_at=CURRENT_TIMESTAMP, export_status=? WHERE id=?",[(new_status,x['id']) for x in data]); c.commit()
    finally: c.close()
    return blob,len(data)

def export_new_items_xls(): return _export_items_status('NEW_OR_EXPORTED','EXPORTED','اصناف_جديدة_للمتكامل','الأصناف الجديدة')
def export_updated_items_xls(): return _export_items_status('UPDATE_PENDING','EXPORTED','تعديلات_الأصناف_للمتكامل','تعديلات الأصناف')

def export_items_xlsx(simple=False):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from openpyxl.utils import get_column_letter
    from io import BytesIO
    c=dbconn()
    try:
        if simple:
            item_rows=c.execute("SELECT * FROM local_items WHERE active=1 ORDER BY item_code").fetchall()
            rows=[]
            for item in item_rows:
                units=[dict(u) for u in c.execute("SELECT id,unit_name,conversion_factor,active FROM item_units WHERE local_item_id=? AND active=1 ORDER BY id",(item['id'],)).fetchall()]
                preferred=_preferred_purchase_unit(units,item['default_purchase_unit_id'],item['main_unit'])
                rows.append({
                    'group_code':item['group_code'], 'item_code':item['item_code'], 'item_name':item['item_name'],
                    'export_unit':preferred.get('unit_name') or item['main_unit'] or '',
                    'pack':preferred.get('conversion_factor') or 1,
                })
            wb=Workbook(); ws=wb.active; ws.title='الأصناف'
            headers=['رقم المجموعة','كود الصنف','اسم الصنف','الوحدة','التعبئة']
            ws.append(headers)
            for cell in ws[1]: cell.font=Font(bold=True)
            for r in rows: ws.append([r['group_code'] or '',r['item_code'] or '',r['item_name'] or '',r['export_unit'] or '',r['pack'] or 1])
            sheets=[ws]
        else:
            item_rows=c.execute("""SELECT li.id,li.item_code,li.group_code,li.item_name,li.item_type,li.main_unit,li.active,li.source,li.created_at,li.updated_at,
                                         li.default_sale_unit_id,li.default_purchase_unit_id,
                                         su.unit_name AS default_sale_unit,pu.unit_name AS default_purchase_unit,
                                         COALESCE((SELECT iu.conversion_factor FROM item_units iu WHERE iu.local_item_id=li.id AND iu.is_default=1 AND iu.active=1 ORDER BY iu.id LIMIT 1),1) AS pack,
                                         COALESCE((SELECT iu.barcode FROM item_units iu WHERE iu.local_item_id=li.id AND iu.is_default=1 AND iu.active=1 ORDER BY iu.id LIMIT 1),'') AS barcode
                                  FROM local_items li
                                  LEFT JOIN item_units su ON su.id=li.default_sale_unit_id AND su.active=1
                                  LEFT JOIN item_units pu ON pu.id=li.default_purchase_unit_id AND pu.active=1
                                  ORDER BY li.item_code""").fetchall()
            unit_rows=c.execute("""SELECT li.item_code,li.item_name,iu.unit_name,iu.conversion_factor,iu.barcode,iu.is_default,iu.active
                                  FROM item_units iu JOIN local_items li ON li.id=iu.local_item_id ORDER BY li.item_code,iu.is_default DESC,iu.id""").fetchall()
            supplier_rows=c.execute("""SELECT li.item_code,li.item_name,s.supplier_code,s.supplier_name,m.supplier_item_code,m.supplier_item_name,m.supplier_unit,m.conversion_factor,m.confidence_score,m.status,m.usage_count
                                      FROM supplier_item_mappings m JOIN local_items li ON li.id=m.local_item_id JOIN suppliers s ON s.id=m.supplier_id
                                      ORDER BY li.item_code,s.supplier_name,m.supplier_item_code""").fetchall()
            alias_rows=c.execute("""SELECT li.item_code,li.item_name,s.supplier_code,s.supplier_name,a.alias_name,a.usage_count,a.first_seen_at,a.last_seen_at,a.source
                                  FROM supplier_item_aliases a JOIN local_items li ON li.id=a.local_item_id JOIN suppliers s ON s.id=a.supplier_id
                                  ORDER BY li.item_code,s.supplier_name,a.alias_name""").fetchall()
            price_rows=c.execute("""SELECT li.item_code,li.item_name,s.supplier_code,s.supplier_name,p.invoice_date,p.unit_price_halalah,p.quantity,p.supplier_unit AS unit_name
                                  FROM supplier_item_prices p JOIN local_items li ON li.id=p.local_item_id JOIN suppliers s ON s.id=p.supplier_id
                                  ORDER BY li.item_code,p.invoice_date DESC,p.id DESC""").fetchall()
            group_rows=c.execute("""SELECT g.group_code,g.group_name,g.active,COUNT(li.id) item_count,
                                         COALESCE(MAX(CASE WHEN li.active=1 THEN li.item_code END),'') last_code
                                  FROM item_groups g LEFT JOIN local_items li ON li.group_code=g.group_code
                                  GROUP BY g.group_code ORDER BY g.group_code""").fetchall()
            wb=Workbook(); ws=wb.active; ws.title='الأصناف'; sheets=[ws]
            specs=[
                ('الأصناف',['كود الصنف','رقم المجموعة','اسم الصنف','الوحدة الرئيسية','التعبئة','الباركود','نوع الصنف','الحالة','المصدر','تاريخ الإنشاء','آخر تحديث','وحدة البيع الافتراضية','وحدة الشراء الافتراضية'],item_rows),
                ('الوحدات',['كود الصنف','اسم الصنف','الوحدة','معامل التحويل','الباركود','أساسية','الحالة'],unit_rows),
                ('الموردون',['كود الصنف','اسم الصنف','كود المورد','اسم المورد','كود الصنف عند المورد','اسم الصنف عند المورد','الوحدة عند المورد','معامل التحويل','الثقة','الحالة','الاستخدامات'],supplier_rows),
                ('الأسماء البديلة',['كود الصنف','اسم الصنف','كود المورد','اسم المورد','الاسم البديل','عدد الاستخدامات','أول ظهور','آخر ظهور','المصدر'],alias_rows),
                ('الأسعار',['كود الصنف','اسم الصنف','كود المورد','اسم المورد','تاريخ الشراء','سعر الوحدة (هللة)','الكمية','الوحدة'],price_rows),
                ('المجموعات',['رقم المجموعة','اسم المجموعة','الحالة','عدد الأصناف','آخر رقم مستخدم'],group_rows),
            ]
            # Fill first sheet and create the rest.
            for idx,(title,headers,rows) in enumerate(specs):
                sh=ws if idx==0 else wb.create_sheet(title)
                if idx==0: sh.title=title
                sh.append(headers)
                for cell in sh[1]: cell.font=Font(bold=True)
                for r in rows:
                    if title=='الأصناف': vals=[r['item_code'],r['group_code'] or '',r['item_name'],r['main_unit'] or '',r['pack'] or 1,r['barcode'] or '',r['item_type'] or '', 'نشط' if r['active'] else 'في السلة',r['source'] or '',r['created_at'] or '',r['updated_at'] or '',r['default_sale_unit'] or '',r['default_purchase_unit'] or '']
                    elif title=='الوحدات': vals=[r['item_code'],r['item_name'],r['unit_name'],r['conversion_factor'],r['barcode'] or '', 'نعم' if r['is_default'] else 'لا','نشطة' if r['active'] else 'موقوفة']
                    elif title=='الموردون': vals=[r['item_code'],r['item_name'],r['supplier_code'] or '',r['supplier_name'],r['supplier_item_code'] or '',r['supplier_item_name'] or '',r['supplier_unit'] or '',r['conversion_factor'] or '',r['confidence_score'] or 0,r['status'] or '',r['usage_count'] or 0]
                    elif title=='الأسماء البديلة': vals=[r['item_code'],r['item_name'],r['supplier_code'] or '',r['supplier_name'],r['alias_name'],r['usage_count'] or 0,r['first_seen_at'] or '',r['last_seen_at'] or '',r['source'] or '']
                    elif title=='الأسعار': vals=[r['item_code'],r['item_name'],r['supplier_code'] or '',r['supplier_name'],r['invoice_date'] or '',r['unit_price_halalah'] or 0,r['quantity'] or 0,r['unit_name'] or '']
                    else: vals=[r['group_code'],r['group_name'] or '', 'نشطة' if r['active'] else 'موقوفة',r['item_count'] or 0,r['last_code'] or '']
                    sh.append(vals)
                sheets.append(sh) if sh not in sheets else None
        for sh in wb.worksheets:
            sh.freeze_panes='A2'; sh.auto_filter.ref=sh.dimensions; sh.sheet_view.rightToLeft=True
            for col in sh.columns:
                vals=[str(x.value or '') for x in col]
                width=min(max(max((len(v) for v in vals),default=10)+2,10),45)
                sh.column_dimensions[get_column_letter(col[0].column)].width=width
                for cell in col: cell.alignment=Alignment(vertical='top',wrap_text=True)
        bio=BytesIO(); wb.save(bio); return bio.getvalue()
    finally:
        c.close()


def init_item_management_schema():
    """Bootstrap every core table/index so a fresh installation can start from an empty DB."""
    c=sqlite3.connect(DB, timeout=10)
    c.row_factory=sqlite3.Row
    try:
        c.execute("PRAGMA foreign_keys=ON")
        c.execute("PRAGMA busy_timeout=10000")
        c.execute("PRAGMA journal_mode=WAL")
        c.executescript("""
        CREATE TABLE IF NOT EXISTS suppliers (
            id INTEGER PRIMARY KEY,
            supplier_code TEXT,
            supplier_name TEXT NOT NULL,
            normalized_name TEXT NOT NULL,
            vat_number TEXT,
            commercial_register TEXT,
            phone TEXT,
            active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0,1)),
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS local_items (
            id INTEGER PRIMARY KEY,
            item_code TEXT NOT NULL,
            group_code TEXT,
            item_name TEXT NOT NULL,
            normalized_name TEXT NOT NULL,
            item_type TEXT,
            main_unit TEXT,
            active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0,1)),
            source TEXT DEFAULT 'AL_MUTAKAMIL',
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS item_units (
            id INTEGER PRIMARY KEY,
            local_item_id INTEGER NOT NULL,
            unit_name TEXT NOT NULL,
            normalized_unit TEXT NOT NULL,
            conversion_factor NUMERIC NOT NULL DEFAULT 1 CHECK(conversion_factor>0),
            barcode TEXT,
            is_default INTEGER NOT NULL DEFAULT 0 CHECK(is_default IN (0,1)),
            notes TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            active INTEGER NOT NULL DEFAULT 1,
            FOREIGN KEY(local_item_id) REFERENCES local_items(id) ON DELETE CASCADE ON UPDATE CASCADE
        );
        CREATE TABLE IF NOT EXISTS supplier_item_mappings (
            id INTEGER PRIMARY KEY,
            supplier_id INTEGER NOT NULL,
            local_item_id INTEGER NOT NULL,
            supplier_item_code TEXT,
            supplier_item_name TEXT NOT NULL,
            normalized_name TEXT NOT NULL,
            supplier_unit TEXT,
            normalized_unit TEXT,
            conversion_factor NUMERIC DEFAULT 1 CHECK(conversion_factor IS NULL OR conversion_factor>0),
            confidence_score NUMERIC CHECK(confidence_score IS NULL OR (confidence_score>=0 AND confidence_score<=100)),
            match_method TEXT NOT NULL DEFAULT 'MANUAL',
            status TEXT NOT NULL DEFAULT 'PENDING_REVIEW',
            approved_by TEXT,
            approved_at TEXT,
            first_seen_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            last_seen_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            usage_count INTEGER NOT NULL DEFAULT 0 CHECK(usage_count>=0),
            notes TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(supplier_id) REFERENCES suppliers(id) ON DELETE RESTRICT ON UPDATE CASCADE,
            FOREIGN KEY(local_item_id) REFERENCES local_items(id) ON DELETE RESTRICT ON UPDATE CASCADE
        );
        CREATE TABLE IF NOT EXISTS supplier_item_aliases (
            id INTEGER PRIMARY KEY,
            supplier_id INTEGER NOT NULL,
            local_item_id INTEGER NOT NULL,
            mapping_id INTEGER,
            alias_name TEXT NOT NULL,
            normalized_name TEXT NOT NULL,
            source TEXT NOT NULL DEFAULT 'INVOICE',
            usage_count INTEGER NOT NULL DEFAULT 1 CHECK(usage_count>=0),
            first_seen_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            last_seen_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(supplier_id) REFERENCES suppliers(id) ON DELETE RESTRICT ON UPDATE CASCADE,
            FOREIGN KEY(local_item_id) REFERENCES local_items(id) ON DELETE RESTRICT ON UPDATE CASCADE,
            FOREIGN KEY(mapping_id) REFERENCES supplier_item_mappings(id) ON DELETE SET NULL ON UPDATE CASCADE
        );
        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY,
            supplier_id INTEGER NOT NULL,
            invoice_number TEXT NOT NULL,
            invoice_date TEXT NOT NULL,
            source_file TEXT,
            source_type TEXT NOT NULL DEFAULT 'PDF',
            subtotal_halalah INTEGER DEFAULT 0,
            vat_amount_halalah INTEGER DEFAULT 0,
            total_amount_halalah INTEGER DEFAULT 0,
            processing_status TEXT NOT NULL DEFAULT 'NEW',
            extraction_confidence NUMERIC CHECK(extraction_confidence IS NULL OR (extraction_confidence>=0 AND extraction_confidence<=100)),
            notes TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(supplier_id) REFERENCES suppliers(id) ON DELETE RESTRICT ON UPDATE CASCADE
        );
        CREATE TABLE IF NOT EXISTS invoice_lines (
            id INTEGER PRIMARY KEY,
            invoice_id INTEGER NOT NULL,
            line_number INTEGER NOT NULL,
            supplier_item_code TEXT,
            raw_item_name TEXT NOT NULL,
            normalized_item_name TEXT NOT NULL,
            supplier_unit TEXT,
            normalized_unit TEXT,
            quantity NUMERIC NOT NULL DEFAULT 0 CHECK(quantity>=0),
            unit_price_halalah INTEGER DEFAULT 0,
            net_amount_halalah INTEGER DEFAULT 0,
            vat_amount_halalah INTEGER DEFAULT 0,
            gross_amount_halalah INTEGER DEFAULT 0,
            matched_local_item_id INTEGER,
            matched_mapping_id INTEGER,
            match_score NUMERIC CHECK(match_score IS NULL OR (match_score>=0 AND match_score<=100)),
            match_status TEXT NOT NULL DEFAULT 'UNMATCHED',
            user_decision TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(invoice_id) REFERENCES invoices(id) ON DELETE CASCADE ON UPDATE CASCADE,
            FOREIGN KEY(matched_local_item_id) REFERENCES local_items(id) ON DELETE SET NULL ON UPDATE CASCADE,
            FOREIGN KEY(matched_mapping_id) REFERENCES supplier_item_mappings(id) ON DELETE SET NULL ON UPDATE CASCADE,
            UNIQUE(invoice_id,line_number)
        );
        CREATE TABLE IF NOT EXISTS match_history (
            id INTEGER PRIMARY KEY,
            invoice_line_id INTEGER NOT NULL,
            suggested_item_id INTEGER,
            actual_item_id INTEGER,
            suggested_mapping_id INTEGER,
            actual_mapping_id INTEGER,
            algorithm_score NUMERIC,
            decision TEXT NOT NULL,
            decision_source TEXT NOT NULL DEFAULT 'USER',
            notes TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(invoice_line_id) REFERENCES invoice_lines(id) ON DELETE CASCADE ON UPDATE CASCADE,
            FOREIGN KEY(suggested_item_id) REFERENCES local_items(id) ON DELETE SET NULL ON UPDATE CASCADE,
            FOREIGN KEY(actual_item_id) REFERENCES local_items(id) ON DELETE SET NULL ON UPDATE CASCADE,
            FOREIGN KEY(suggested_mapping_id) REFERENCES supplier_item_mappings(id) ON DELETE SET NULL ON UPDATE CASCADE,
            FOREIGN KEY(actual_mapping_id) REFERENCES supplier_item_mappings(id) ON DELETE SET NULL ON UPDATE CASCADE
        );
        CREATE TABLE IF NOT EXISTS supplier_item_prices (
            id INTEGER PRIMARY KEY,
            supplier_id INTEGER NOT NULL,
            local_item_id INTEGER NOT NULL,
            mapping_id INTEGER,
            invoice_id INTEGER,
            invoice_line_id INTEGER,
            invoice_date TEXT NOT NULL,
            supplier_unit TEXT,
            conversion_factor NUMERIC DEFAULT 1 CHECK(conversion_factor IS NULL OR conversion_factor>0),
            unit_price_halalah INTEGER NOT NULL DEFAULT 0,
            quantity NUMERIC NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(supplier_id) REFERENCES suppliers(id) ON DELETE RESTRICT ON UPDATE CASCADE,
            FOREIGN KEY(local_item_id) REFERENCES local_items(id) ON DELETE RESTRICT ON UPDATE CASCADE,
            FOREIGN KEY(mapping_id) REFERENCES supplier_item_mappings(id) ON DELETE SET NULL ON UPDATE CASCADE,
            FOREIGN KEY(invoice_id) REFERENCES invoices(id) ON DELETE SET NULL ON UPDATE CASCADE,
            FOREIGN KEY(invoice_line_id) REFERENCES invoice_lines(id) ON DELETE SET NULL ON UPDATE CASCADE
        );
        CREATE TABLE IF NOT EXISTS item_change_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entity_type TEXT NOT NULL DEFAULT 'ITEM', entity_id INTEGER NOT NULL,
            action TEXT NOT NULL, before_json TEXT, after_json TEXT,
            changed_fields TEXT, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS supplier_change_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entity_id INTEGER NOT NULL, action TEXT NOT NULL, before_json TEXT, after_json TEXT,
            changed_fields TEXT, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS item_groups (
            group_code TEXT PRIMARY KEY, group_name TEXT NOT NULL DEFAULT '', active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS unit_catalog (
            unit_name TEXT PRIMARY KEY, normalized_unit TEXT NOT NULL UNIQUE, active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS invoice_extraction_templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_id INTEGER,
            supplier_name TEXT,
            template_name TEXT NOT NULL,
            header_signature TEXT NOT NULL,
            column_map_json TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            usage_count INTEGER NOT NULL DEFAULT 0,
            active INTEGER NOT NULL DEFAULT 1,
            FOREIGN KEY(supplier_id) REFERENCES suppliers(id) ON DELETE SET NULL ON UPDATE CASCADE
        );
        CREATE TABLE IF NOT EXISTS extraction_corrections (id INTEGER PRIMARY KEY AUTOINCREMENT,supplier_id INTEGER,header_signature TEXT NOT NULL,field TEXT NOT NULL,original_column INTEGER,corrected_column INTEGER,learned_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,FOREIGN KEY(supplier_id) REFERENCES suppliers(id) ON DELETE SET NULL ON UPDATE CASCADE);
        CREATE INDEX IF NOT EXISTS ix_extraction_corrections_signature ON extraction_corrections(supplier_id,header_signature,field,learned_at DESC,id DESC);
        CREATE UNIQUE INDEX IF NOT EXISTS ux_extraction_template_supplier_signature ON invoice_extraction_templates(supplier_id,header_signature) WHERE supplier_id IS NOT NULL;
        CREATE UNIQUE INDEX IF NOT EXISTS ux_extraction_template_general_signature ON invoice_extraction_templates(header_signature) WHERE supplier_id IS NULL;
        CREATE UNIQUE INDEX IF NOT EXISTS ux_suppliers_normalized_name ON suppliers(normalized_name);
        CREATE UNIQUE INDEX IF NOT EXISTS ux_suppliers_code ON suppliers(supplier_code) WHERE supplier_code IS NOT NULL AND supplier_code<>'';
        CREATE UNIQUE INDEX IF NOT EXISTS ux_local_items_item_code ON local_items(item_code);
        CREATE INDEX IF NOT EXISTS ix_local_items_normalized_name ON local_items(normalized_name);
        CREATE UNIQUE INDEX IF NOT EXISTS ux_item_units_item_unit_factor ON item_units(local_item_id,normalized_unit,conversion_factor);
        CREATE UNIQUE INDEX IF NOT EXISTS ux_item_units_barcode ON item_units(barcode) WHERE barcode IS NOT NULL AND barcode<>'';
        CREATE INDEX IF NOT EXISTS ix_item_units_item ON item_units(local_item_id);
        CREATE UNIQUE INDEX IF NOT EXISTS ux_mapping_supplier_code ON supplier_item_mappings(supplier_id,supplier_item_code) WHERE supplier_item_code IS NOT NULL AND supplier_item_code<>'';
        CREATE INDEX IF NOT EXISTS ix_mapping_supplier ON supplier_item_mappings(supplier_id);
        CREATE INDEX IF NOT EXISTS ix_mapping_local_item ON supplier_item_mappings(local_item_id);
        CREATE INDEX IF NOT EXISTS ix_mapping_normalized_name ON supplier_item_mappings(supplier_id,normalized_name);
        CREATE INDEX IF NOT EXISTS ix_mapping_status ON supplier_item_mappings(status);
        CREATE UNIQUE INDEX IF NOT EXISTS ux_alias_supplier_name_item ON supplier_item_aliases(supplier_id,normalized_name,local_item_id);
        CREATE INDEX IF NOT EXISTS ix_alias_supplier_name ON supplier_item_aliases(supplier_id,normalized_name);
        CREATE UNIQUE INDEX IF NOT EXISTS ux_invoice_supplier_number ON invoices(supplier_id,invoice_number);
        CREATE INDEX IF NOT EXISTS ix_invoice_date ON invoices(invoice_date);
        CREATE INDEX IF NOT EXISTS ix_invoice_supplier_date ON invoices(supplier_id,invoice_date);
        CREATE INDEX IF NOT EXISTS ix_invoice_status ON invoices(processing_status);
        CREATE INDEX IF NOT EXISTS ix_invoice_lines_invoice ON invoice_lines(invoice_id);
        CREATE INDEX IF NOT EXISTS ix_invoice_lines_match_status ON invoice_lines(match_status);
        CREATE INDEX IF NOT EXISTS ix_invoice_lines_normalized_name ON invoice_lines(normalized_item_name);
        CREATE INDEX IF NOT EXISTS ix_invoice_lines_supplier_code ON invoice_lines(supplier_item_code);
        CREATE INDEX IF NOT EXISTS ix_match_history_invoice_line ON match_history(invoice_line_id);
        CREATE INDEX IF NOT EXISTS ix_match_history_actual_item ON match_history(actual_item_id);
        CREATE INDEX IF NOT EXISTS ix_prices_supplier_item_date ON supplier_item_prices(supplier_id,local_item_id,invoice_date);
        CREATE INDEX IF NOT EXISTS ix_prices_item_date ON supplier_item_prices(local_item_id,invoice_date);
        """)
        # Migrate older databases that predate the active unit flag.
        cols={r[1] for r in c.execute("PRAGMA table_info(item_units)").fetchall()}
        if 'active' not in cols:
            c.execute("ALTER TABLE item_units ADD COLUMN active INTEGER NOT NULL DEFAULT 1")
        # Backfill groups from existing items.
        for r in c.execute("SELECT DISTINCT group_code FROM local_items WHERE group_code IS NOT NULL AND trim(group_code)<>''").fetchall():
            c.execute("INSERT OR IGNORE INTO item_groups(group_code,group_name) VALUES(?,?)",(r['group_code'],''))
        # Backfill the global unit catalog from all existing item units.
        for r in c.execute("SELECT DISTINCT unit_name,normalized_unit FROM item_units WHERE trim(unit_name)<>''").fetchall():
            c.execute("INSERT OR IGNORE INTO unit_catalog(unit_name,normalized_unit) VALUES(?,?)",(r['unit_name'],r['normalized_unit']))
        c.execute("INSERT OR IGNORE INTO unit_catalog(unit_name,normalized_unit) VALUES(?,?)",('حبة',norm('حبة')))
        c.execute("INSERT OR IGNORE INTO unit_catalog(unit_name,normalized_unit) VALUES(?,?)",('كرتون',norm('كرتون')))
        # A local item may legitimately have multiple supplier-code spellings
        # (e.g. 024097, 24097, 24097.0). Keep uniqueness only on the exact
        # supplier code; do not block additional codes for the same item.
        c.execute("DROP INDEX IF EXISTS ux_mapping_supplier_name_unit_pack_item")
        c.commit()
    except Exception:
        c.rollback()
        raise
    finally:
        c.close()


def item_snapshot(c, iid):
    r=c.execute("SELECT id,item_code,group_code,item_name,normalized_name,item_type,main_unit,default_sale_unit_id,default_purchase_unit_id,active,source,created_at,updated_at FROM local_items WHERE id=?",(iid,)).fetchone()
    return dict(r) if r else None

def log_item_change(c, iid, action, before, after):
    b=json.dumps(before,ensure_ascii=False,sort_keys=True) if before else None
    a=json.dumps(after,ensure_ascii=False,sort_keys=True) if after else None
    changed=[]
    if before and after:
        changed=[k for k in set(before)|set(after) if before.get(k)!=after.get(k)]
    c.execute("INSERT INTO item_change_history(entity_type,entity_id,action,before_json,after_json,changed_fields) VALUES('ITEM',?,?,?,?,?)",
              (iid,action,b,a,', '.join(changed)))

def items_all(include_deleted=False):
    c=dbconn()
    where='' if include_deleted else 'WHERE li.active=1'
    rows=c.execute(f"""SELECT li.id,li.item_code,li.group_code,li.item_name,li.item_type,li.main_unit,
                       li.default_sale_unit_id,li.default_purchase_unit_id,
                       su.unit_name AS default_sale_unit, pu.unit_name AS default_purchase_unit,
                       li.active,li.source,li.export_status,
                       GROUP_CONCAT(DISTINCT iu.unit_name) units,
                       GROUP_CONCAT(DISTINCT iu.barcode) barcodes,
                       GROUP_CONCAT(DISTINCT m.supplier_item_code) supplier_codes,
                       COUNT(DISTINCT m.id) mapping_count,
                       COUNT(DISTINCT a.id) alias_count
                       FROM local_items li
                       LEFT JOIN item_units iu ON iu.local_item_id=li.id AND iu.active=1
                       LEFT JOIN item_units su ON su.id=li.default_sale_unit_id AND su.active=1
                       LEFT JOIN item_units pu ON pu.id=li.default_purchase_unit_id AND pu.active=1
                       LEFT JOIN supplier_item_mappings m ON m.local_item_id=li.id AND m.status!='DISABLED'
                       LEFT JOIN supplier_item_aliases a ON a.local_item_id=li.id
                       {where} GROUP BY li.id ORDER BY li.active DESC,li.item_code""").fetchall(); c.close()
    return [dict(r) for r in rows]

def item_card(iid):
    c=dbconn(); it=c.execute("SELECT * FROM local_items WHERE id=?",(iid,)).fetchone()
    if not it: c.close(); return None
    units=[dict(r) for r in c.execute("SELECT * FROM item_units WHERE local_item_id=? ORDER BY is_default DESC,id",(iid,)).fetchall()]
    maps=[dict(r) for r in c.execute("""SELECT m.*,s.supplier_name,s.supplier_code FROM supplier_item_mappings m JOIN suppliers s ON s.id=m.supplier_id WHERE m.local_item_id=? ORDER BY s.supplier_name,m.id""",(iid,)).fetchall()]
    als=[dict(r) for r in c.execute("""SELECT a.*,s.supplier_name FROM supplier_item_aliases a JOIN suppliers s ON s.id=a.supplier_id WHERE a.local_item_id=? ORDER BY a.last_seen_at DESC,a.id DESC""",(iid,)).fetchall()]
    prices=[dict(r) for r in c.execute("""SELECT p.*,s.supplier_name FROM supplier_item_prices p JOIN suppliers s ON s.id=p.supplier_id WHERE p.local_item_id=? ORDER BY p.invoice_date DESC,p.id DESC LIMIT 100""",(iid,)).fetchall()]
    hist=[dict(r) for r in c.execute("SELECT * FROM item_change_history WHERE entity_type='ITEM' AND entity_id=? ORDER BY id DESC LIMIT 100",(iid,)).fetchall()]
    inv=[dict(r) for r in c.execute("""SELECT i.invoice_number,i.invoice_date,s.supplier_name,il.raw_item_name,il.quantity,il.unit_price_halalah,il.net_amount_halalah,il.vat_amount_halalah,il.gross_amount_halalah
                                      FROM invoice_lines il JOIN invoices i ON i.id=il.invoice_id JOIN suppliers s ON s.id=i.supplier_id
                                      WHERE il.matched_local_item_id=? ORDER BY i.invoice_date DESC,i.id DESC LIMIT 100""",(iid,)).fetchall()]
    c.close(); return {"item":dict(it),"units":units,"mappings":maps,"aliases":als,"prices":prices,"history":hist,"invoices":inv}

def managed_units():
    c=dbconn()
    rows=c.execute("""SELECT u.unit_name,u.normalized_unit,u.active,
                            (SELECT COUNT(DISTINCT iu.local_item_id) FROM item_units iu WHERE iu.normalized_unit=u.normalized_unit) item_count
                     FROM unit_catalog u ORDER BY u.active DESC,u.unit_name""").fetchall()
    c.close(); return [dict(r) for r in rows]

def managed_groups():
    c=dbconn(); rows=c.execute("""SELECT g.group_code,g.group_name,g.active,COUNT(li.id) item_count,MAX(li.item_code) last_code
                                    FROM item_groups g LEFT JOIN local_items li ON li.group_code=g.group_code AND li.active=1
                                    GROUP BY g.group_code ORDER BY g.active DESC,g.group_code""").fetchall(); c.close(); return [dict(r) for r in rows]

def add_managed_unit(payload):
    name=str(payload.get('unit_name') or '').strip()
    if not name: raise ValueError('اسم الوحدة مطلوب.')
    c=dbconn()
    try:
        ex=c.execute('SELECT unit_name FROM unit_catalog WHERE normalized_unit=? LIMIT 1',(norm(name),)).fetchone()
        if ex: raise ValueError(f'الوحدة موجودة بالفعل: {ex[0]}')
        c.execute('INSERT INTO unit_catalog(unit_name,normalized_unit,active) VALUES(?,?,1)',(name,norm(name)))
        c.commit(); return {'unit_name':name}
    except: c.rollback(); raise
    finally: c.close()

def update_managed_unit(payload):
    old_name=str(payload.get('old_name') or '').strip(); name=str(payload.get('unit_name') or '').strip()
    if not old_name or not name: raise ValueError('اسم الوحدة مطلوب.')
    if norm(old_name)==norm(name): return {'unit_name':name}
    c=dbconn()
    try:
        r=c.execute('SELECT * FROM unit_catalog WHERE normalized_unit=? LIMIT 1',(norm(old_name),)).fetchone()
        if not r: raise ValueError('الوحدة غير موجودة.')
        ex=c.execute('SELECT unit_name FROM unit_catalog WHERE normalized_unit=? LIMIT 1',(norm(name),)).fetchone()
        if ex: raise ValueError(f'الوحدة الجديدة موجودة بالفعل: {ex[0]}')
        c.execute('UPDATE unit_catalog SET unit_name=?,normalized_unit=?,updated_at=CURRENT_TIMESTAMP WHERE normalized_unit=?',(name,norm(name),norm(old_name)))
        # Keep existing item data consistent with the renamed global unit.
        c.execute('UPDATE item_units SET unit_name=?,normalized_unit=?,updated_at=CURRENT_TIMESTAMP WHERE normalized_unit=?',(name,norm(name),norm(old_name)))
        c.execute("UPDATE local_items SET main_unit=?,updated_at=CURRENT_TIMESTAMP,export_status=CASE WHEN export_status='NEW' THEN 'NEW' ELSE 'UPDATE_PENDING' END WHERE main_unit=?",(name,old_name))
        c.commit(); return {'unit_name':name}
    except: c.rollback(); raise
    finally: c.close()

def toggle_managed_unit(unit_name, active):
    c=dbconn()
    try:
        r=c.execute('SELECT * FROM unit_catalog WHERE normalized_unit=? LIMIT 1',(norm(unit_name),)).fetchone()
        if not r: raise ValueError('الوحدة غير موجودة.')
        if not active and r['normalized_unit'] in (norm('حبة'),norm('كرتون')):
            raise ValueError('لا يمكن حذف أو إيقاف الوحدة الأساسية حبة أو كرتون.')
        count=c.execute('SELECT COUNT(DISTINCT local_item_id) FROM item_units WHERE normalized_unit=?',(r['normalized_unit'],)).fetchone()[0]
        if not active and count==0:
            c.execute('DELETE FROM unit_catalog WHERE normalized_unit=?',(r['normalized_unit'],)); c.commit(); return {'unit_name':r['unit_name'],'active':0,'deleted':1}
        c.execute('UPDATE unit_catalog SET active=?,updated_at=CURRENT_TIMESTAMP WHERE normalized_unit=?',(1 if active else 0,r['normalized_unit'])); c.commit()
        return {'unit_name':r['unit_name'],'active':1 if active else 0,'deleted':0}
    except: c.rollback(); raise
    finally: c.close()

def add_managed_group(payload):
    code=str(payload.get('group_code') or '').strip(); name=str(payload.get('group_name') or '').strip()
    if not code or not code.isdigit(): raise ValueError('رقم المجموعة مطلوب ويجب أن يكون رقميًا.')
    if not name: raise ValueError('اسم المجموعة مطلوب.')
    c=dbconn()
    try:
        ex=c.execute('SELECT group_code FROM item_groups WHERE group_code=?',(code,)).fetchone()
        if ex: raise ValueError(f'المجموعة {code} موجودة بالفعل.')
        c.execute('INSERT INTO item_groups(group_code,group_name,active) VALUES(?,?,1)',(code,name)); c.commit(); return {'group_code':code,'group_name':name}
    except: c.rollback(); raise
    finally: c.close()

def update_managed_group(payload):
    code=str(payload.get('group_code') or '').strip(); name=str(payload.get('group_name') or '').strip()
    if not code or not name: raise ValueError('رقم المجموعة واسمها مطلوبان.')
    c=dbconn()
    try:
        r=c.execute('SELECT * FROM item_groups WHERE group_code=?',(code,)).fetchone()
        if not r: raise ValueError('المجموعة غير موجودة.')
        c.execute('UPDATE item_groups SET group_name=?,updated_at=CURRENT_TIMESTAMP WHERE group_code=?',(name,code)); c.commit(); return {'group_code':code,'group_name':name}
    except: c.rollback(); raise
    finally: c.close()

def toggle_managed_group(code, active):
    code=str(code or '').strip()
    c=dbconn()
    try:
        r=c.execute('SELECT * FROM item_groups WHERE group_code=?',(code,)).fetchone()
        if not r: raise ValueError('المجموعة غير موجودة.')
        count=c.execute('SELECT COUNT(*) FROM local_items WHERE group_code=?',(code,)).fetchone()[0]
        if not active and count==0:
            c.execute('DELETE FROM item_groups WHERE group_code=?',(code,)); c.commit(); return {'group_code':code,'active':0,'deleted':1}
        c.execute('UPDATE item_groups SET active=?,updated_at=CURRENT_TIMESTAMP WHERE group_code=?',(1 if active else 0,code)); c.commit(); return {'group_code':code,'active':1 if active else 0,'deleted':0}
    except: c.rollback(); raise
    finally: c.close()


def update_item(payload):
    iid=int(payload.get('id') or 0); c=dbconn()
    try:
        before=item_snapshot(c,iid)
        if not before: raise ValueError('الصنف غير موجود.')
        name=str(payload.get('item_name') or '').strip(); group=str(payload.get('group_code') or '').strip(); code=str(payload.get('item_code') or '').strip()
        unit=str(payload.get('main_unit') or '').strip() or before['main_unit']
        typ=str(payload.get('item_type') or '').strip() or 'سلعي'
        if not name or not group or not code or not unit: raise ValueError('الكود والاسم والمجموعة والوحدة الأساسية مطلوبة.')
        if not group.isdigit() or not code.isdigit(): raise ValueError('المجموعة ورقم الصنف يجب أن يكونا رقميين.')
        ex=c.execute("SELECT id,item_name FROM local_items WHERE item_code=? AND id<>? LIMIT 1",(code,iid)).fetchone()
        if ex: raise ValueError(f'رقم الصنف {code} مستخدم بالفعل للصنف {ex[1]}.')
        ex=c.execute("SELECT id,item_code FROM local_items WHERE normalized_name=? AND active=1 AND id<>? LIMIT 1",(norm(name),iid)).fetchone()
        if ex: raise ValueError(f'يوجد صنف بنفس الاسم: {ex[1]}.')
        oldgroup=before['group_code']
        c.execute("UPDATE local_items SET item_code=?,group_code=?,item_name=?,normalized_name=?,item_type=?,main_unit=?,export_status=CASE WHEN export_status='NEW' THEN 'NEW' ELSE 'UPDATE_PENDING' END,updated_at=CURRENT_TIMESTAMP WHERE id=?",
                  (code,group,name,norm(name),typ,unit,iid))
        c.execute("UPDATE item_units SET unit_name=?,normalized_unit=?,updated_at=CURRENT_TIMESTAMP WHERE local_item_id=? AND is_default=1",(unit,norm(unit),iid))
        if payload.get('barcode') is not None:
            barcode=str(payload.get('barcode') or '').strip() or None
            c.execute("UPDATE item_units SET barcode=?,updated_at=CURRENT_TIMESTAMP WHERE local_item_id=? AND is_default=1",(barcode,iid))
        # Explicit sale/purchase defaults are independent from is_default (main unit).
        current_units={int(r['id']):r for r in c.execute("SELECT * FROM item_units WHERE local_item_id=? AND active=1",(iid,)).fetchall()}
        sale_uid=before.get('default_sale_unit_id') if 'default_sale_unit_id' in before else None
        purchase_uid=before.get('default_purchase_unit_id') if 'default_purchase_unit_id' in before else None
        def valid_default(value,label):
            if value in (None,'','null'): return None
            try: uid=int(value)
            except Exception: raise ValueError(f'{label} غير صالح.')
            if uid not in current_units: raise ValueError(f'{label} يجب أن تكون وحدة نشطة لهذا الصنف.')
            return uid
        if 'default_sale_unit_id' in payload: sale_uid=valid_default(payload.get('default_sale_unit_id'),'وحدة البيع الافتراضية')
        if 'default_purchase_unit_id' in payload: purchase_uid=valid_default(payload.get('default_purchase_unit_id'),'وحدة الشراء الافتراضية')
        c.execute("UPDATE local_items SET default_sale_unit_id=?,default_purchase_unit_id=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",(sale_uid,purchase_uid,iid))
        c.execute("INSERT OR IGNORE INTO item_groups(group_code,group_name) VALUES(?,?)",(group,''))
        after=item_snapshot(c,iid); log_item_change(c,iid,'UPDATE',before,after); c.commit()
        return item_card(iid)
    except: c.rollback(); raise
    finally: c.close()

def set_item_active(iid, active):
    c=dbconn()
    try:
        before=item_snapshot(c,iid)
        if not before: raise ValueError('الصنف غير موجود.')
        c.execute("UPDATE local_items SET active=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",(1 if active else 0,iid))
        after=item_snapshot(c,iid); log_item_change(c,iid,'RESTORE' if active else 'DELETE',before,after); c.commit(); return after
    except: c.rollback(); raise
    finally: c.close()

def add_item_unit(payload):
    iid=int(payload.get('local_item_id') or 0); name=str(payload.get('unit_name') or '').strip(); factor=float(payload.get('conversion_factor') or 1); barcode=str(payload.get('barcode') or '').strip() or None; default=1 if payload.get('is_default') else 0
    if not iid or not name or factor<=0: raise ValueError('الصنف والوحدة ومعامل التحويل مطلوبة.')
    c=dbconn()
    try:
        if default: c.execute("UPDATE item_units SET is_default=0 WHERE local_item_id=?",(iid,))
        cur=c.execute("INSERT INTO item_units(local_item_id,unit_name,normalized_unit,conversion_factor,barcode,is_default,active) VALUES(?,?,?,?,?,?,1)",(iid,name,norm(name),factor,barcode,default)); uid=cur.lastrowid
        if default: c.execute("UPDATE local_items SET main_unit=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",(name,iid))
        c.execute("UPDATE local_items SET export_status=CASE WHEN export_status='NEW' THEN 'NEW' ELSE 'UPDATE_PENDING' END,updated_at=CURRENT_TIMESTAMP WHERE id=?",(iid,))
        log_item_change(c,iid,'UNIT_ADD',None,{'unit_id':uid,'unit_name':name,'conversion_factor':factor,'barcode':barcode,'is_default':default}); c.commit(); return item_card(iid)
    except: c.rollback(); raise
    finally: c.close()

def update_item_unit(payload):
    uid=int(payload.get('id') or 0); name=str(payload.get('unit_name') or '').strip(); factor=float(payload.get('conversion_factor') or 1); barcode=str(payload.get('barcode') or '').strip() or None; default=1 if payload.get('is_default') else 0
    c=dbconn()
    try:
        r=c.execute("SELECT * FROM item_units WHERE id=?",(uid,)).fetchone()
        if not r: raise ValueError('الوحدة غير موجودة.')
        iid=r['local_item_id']
        if not name or factor<=0: raise ValueError('اسم الوحدة ومعامل التحويل مطلوبان.')
        if default: c.execute("UPDATE item_units SET is_default=0 WHERE local_item_id=?",(iid,))
        c.execute("UPDATE item_units SET unit_name=?,normalized_unit=?,conversion_factor=?,barcode=?,is_default=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",(name,norm(name),factor,barcode,default,uid))
        if default: c.execute("UPDATE local_items SET main_unit=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",(name,iid))
        c.execute("UPDATE local_items SET export_status=CASE WHEN export_status='NEW' THEN 'NEW' ELSE 'UPDATE_PENDING' END,updated_at=CURRENT_TIMESTAMP WHERE id=?",(iid,))
        log_item_change(c,iid,'UNIT_UPDATE',dict(r),{'id':uid,'unit_name':name,'conversion_factor':factor,'barcode':barcode,'is_default':default}); c.commit(); return item_card(iid)
    except: c.rollback(); raise
    finally: c.close()

def toggle_item_unit(uid, active):
    c=dbconn()
    try:
        r=c.execute("SELECT * FROM item_units WHERE id=?",(uid,)).fetchone()
        if not r: raise ValueError('الوحدة غير موجودة.')
        if not active:
            refs=c.execute("SELECT default_sale_unit_id,default_purchase_unit_id FROM local_items WHERE id=?",(r['local_item_id'],)).fetchone()
            if refs and (refs['default_sale_unit_id']==uid or refs['default_purchase_unit_id']==uid):
                raise ValueError('لا يمكن إيقاف وحدة مرتبطة كوحدة بيع أو شراء افتراضية. اختر وحدة أخرى أولًا.')
            if int(r['is_default'] or 0)==1:
                raise ValueError('لا يمكن إيقاف الوحدة الأساسية للصنف.')
        c.execute("UPDATE item_units SET active=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",(1 if active else 0,uid)); log_item_change(c,r['local_item_id'],'UNIT_RESTORE' if active else 'UNIT_DISABLE',dict(r),{'id':uid,'active':1 if active else 0}); c.commit(); return item_card(r['local_item_id'])
    except: c.rollback(); raise
    finally: c.close()

def undo_last_item(iid):
    c=dbconn()
    try:
        h=c.execute("SELECT * FROM item_change_history WHERE entity_type='ITEM' AND entity_id=? ORDER BY id DESC LIMIT 1",(iid,)).fetchone()
        if not h: raise ValueError('لا توجد عملية يمكن التراجع عنها.')
        if h['action'] in ('UPDATE','DELETE','RESTORE') and h['before_json']:
            b=json.loads(h['before_json']); c.execute("UPDATE local_items SET item_code=?,group_code=?,item_name=?,normalized_name=?,item_type=?,main_unit=?,active=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",
                (b['item_code'],b['group_code'],b['item_name'],b['normalized_name'],b['item_type'],b['main_unit'],b['active'],iid))
        elif h['action']=='UNIT_ADD':
            after=json.loads(h['after_json'] or '{}'); c.execute("DELETE FROM item_units WHERE id=?",(after.get('unit_id'),))
        elif h['action']=='UNIT_UPDATE':
            b=json.loads(h['before_json'] or '{}'); c.execute("UPDATE item_units SET unit_name=?,normalized_unit=?,conversion_factor=?,barcode=?,is_default=?,active=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",
                (b.get('unit_name'),b.get('normalized_unit'),b.get('conversion_factor',1),b.get('barcode'),b.get('is_default',0),b.get('active',1),b['id']))
        else: raise ValueError('هذه العملية لا تدعم التراجع الآمن.')
        c.execute("INSERT INTO item_change_history(entity_type,entity_id,action,before_json,after_json,changed_fields) VALUES('ITEM',?,?,?,?,'UNDO')",(iid,'UNDO',h['after_json'],h['before_json']))
        c.commit(); return item_card(iid)
    except: c.rollback(); raise
    finally: c.close()


def _halalah(value):
    try:
        return int(round(float(value or 0) * 100))
    except Exception:
        return 0

def _calculate_invoice_line_amounts(x, vat_rate):
    qty=float(x.get('qty',x.get('quantity') or 0) or 0)
    unit_price=_halalah(x.get('unit_price'))
    net=_halalah(x.get('net_amount') if x.get('net_amount') not in (None,'') else qty*float(x.get('unit_price') or 0))
    gross_raw=x.get('gross_amount'); vat_raw=x.get('vat_amount')
    if gross_raw not in (None,'') and float(gross_raw or 0)>0:
        gross=_halalah(gross_raw)
        if gross<net:
            vat_line=_halalah(vat_raw) if vat_raw not in (None,'') else _halalah(net*vat_rate/100); gross=net+vat_line
        else:
            vat_line=_halalah(vat_raw) if vat_raw not in (None,'') else max(0,gross-net)
    else:
        vat_line=_halalah(vat_raw) if vat_raw not in (None,'') else _halalah(net*vat_rate/100); gross=net+vat_line
    return qty,unit_price,net,vat_line,gross

def _persist_line_state(c, supplier_id, x):
    m=x.get('match') or {}; status=str(m.get('status') or 'UNMATCHED').upper(); item=m.get('item') or {}
    iid=_valid_local_item_for_context(item.get('id'))
    iid=iid if iid and c.execute('SELECT id FROM local_items WHERE id=? AND active=1',(iid,)).fetchone() else None
    mid=_valid_mapping(c,supplier_id,m.get('mapping_id'),iid)
    raw_decision=str(x.get('user_decision') or '').strip().upper()
    decision_map={'AUTO':'APPROVE','MANUAL':'APPROVE','APPROVED':'APPROVE','AUTO_MATCHED':'APPROVE','PENDING':'CHANGE',
                  'SUGGESTED':'CHANGE','NEW_ITEM':'CREATE_NEW','UNMATCHED':'CHANGE','REJECTED':'REJECT','REJECT':'REJECT',
                  'APPROVE':'APPROVE','CHANGE':'CHANGE','CREATE_NEW':'CREATE_NEW'}
    decision=decision_map.get(raw_decision) or decision_map.get(status) or 'CHANGE'
    accepted=status in ('APPROVED','AUTO_MATCHED') and iid is not None
    # Learning is never inferred from the persisted user_decision alone. Only a fresh
    # in-memory _learning descriptor created by an explicit user action can commit new learning.
    manual=bool(isinstance(x.get('_learning'),dict) and x.get('_learning'))
    return {'status':status,'actual_item_id':iid,'mapping_id':mid,'decision':decision,'accepted':accepted,'manual':manual}

def _save_learning_if_confirmed(c,supplier_id,invoice_id,line_id,x,state,warnings):
    """Commit only the fresh in-memory learning chosen by the user during this invoice session."""
    pending=x.get('_learning') or {}
    if not state['accepted'] or not isinstance(pending,dict) or not pending:
        return state['mapping_id'],0,0,False
    try:
        pending_item=int(pending.get('local_item_id') or 0)
    except Exception:
        pending_item=0
    if pending_item != int(state['actual_item_id']):
        # The current line choice is authoritative. A stale pending descriptor is ignored.
        return state['mapping_id'],0,0,False
    learn_code=bool(pending.get('learn_code',False))
    learn_alias=bool(pending.get('learn_alias',False))
    if not learn_code and not learn_alias:
        return state['mapping_id'],0,0,False
    lr=_apply_learning_for_line(
        c,supplier_id,state['actual_item_id'],x,invoice_id,line_id,state['mapping_id'],'USER',
        (x.get('match') or {}).get('score') or 100,
        (x.get('match') or {}).get('method') or 'MANUAL',
        learn_code=learn_code,learn_alias=learn_alias
    )
    if lr['mapping_id']!=state['mapping_id']:
        state['mapping_id']=lr['mapping_id']
        c.execute('UPDATE invoice_lines SET matched_mapping_id=? WHERE id=?',(state['mapping_id'],line_id))
    warnings.extend(lr['warnings'])
    log_learning_history(
        c,'INVOICE_LINE',line_id,supplier_id,state['actual_item_id'],'LEARN_FROM_INVOICE',
        after={'supplier_item_code':x.get('supplier_item_code') if learn_code else None,
               'alias_name':x.get('raw_item_name') if learn_alias else None,
               'learn_code':learn_code,'learn_alias':learn_alias,
               'local_item_id':state['actual_item_id'],'mapping_id':lr.get('mapping_id'),
               'alias_id':lr.get('alias_id'),'learned':lr.get('learned')},
        invoice_id=invoice_id,invoice_line_id=line_id,source='USER'
    )
    return (state['mapping_id'],
            1 if lr.get('mapping_id') and learn_code and x.get('supplier_item_code') else 0,
            1 if lr.get('alias_id') and learn_alias else 0,
            bool(lr.get('learned')))

def save_invoice(payload):
    """Atomically save the invoice and only then commit final user learning."""
    supplier_id=int(payload.get('supplier_id') or 0); invoice_number=str(payload.get('invoice_number') or '').strip()
    invoice_date=str(payload.get('invoice_date') or '').strip(); source_file=str(payload.get('source_file') or '').strip() or None
    warehouse_number=str(payload.get('warehouse_number') or '').strip() or None; raw_source_type=str(payload.get('source_type') or '').strip().upper()
    source_map={'PDF':'PDF','IMAGE':'IMAGE','IMG':'IMAGE','JPG':'IMAGE','JPEG':'IMAGE','PNG':'IMAGE','WEBP':'IMAGE','EXCEL':'EXCEL','XLS':'EXCEL','XLSX':'EXCEL','CSV':'EXCEL','SCAN':'SCAN','SCANNED':'SCAN','OTHER':'OTHER'}
    source_type=source_map.get(raw_source_type,'OTHER'); lines=payload.get('lines') or []
    if not supplier_id: raise ValueError('المورد مطلوب.')
    if not invoice_number: raise ValueError('رقم الفاتورة مطلوب.')
    if not invoice_date: raise ValueError('تاريخ الفاتورة مطلوب.')
    if not warehouse_number or not warehouse_number.isdigit(): raise ValueError('رقم المخزن مطلوب ويجب أن يكون رقمًا فقط.')
    if not lines: raise ValueError('لا توجد بنود لحفظها.')
    c=dbconn()
    try:
        supplier=c.execute('SELECT id,supplier_name FROM suppliers WHERE id=?',(supplier_id,)).fetchone()
        if not supplier: raise ValueError('المورد غير موجود.')
        if c.execute('SELECT id FROM invoices WHERE supplier_id=? AND invoice_number=?',(supplier_id,invoice_number)).fetchone():
            raise ValueError(f'الفاتورة رقم {invoice_number} محفوظة مسبقًا لهذا المورد. استخدم رقمًا مختلفًا أو افتح سجل الفواتير.')
        subtotal=sum(_halalah(x.get('net_amount') if x.get('net_amount') not in (None,'') else float(x.get('qty') or 0)*float(x.get('unit_price') or 0)) for x in lines)
        discount=max(0,int(payload.get('discount_halalah') or 0)); discount=min(discount,subtotal); taxable=subtotal-discount
        vat_rate=max(0,float(payload.get('vat_rate') if payload.get('vat_rate') not in (None,'') else 15)); vat=int(round(taxable*vat_rate/100)); total=taxable+vat
        states=[]
        for x in lines:
            m=x.get('match') or {}; st=str(m.get('status') or 'UNMATCHED').upper(); iid=_valid_local_item_for_context((m.get('item') or {}).get('id'))
            iid=iid if iid and c.execute('SELECT id FROM local_items WHERE id=? AND active=1',(iid,)).fetchone() else None
            states.append((st,iid))
        approved=all(st in ('APPROVED','AUTO_MATCHED') and iid is not None for st,iid in states)
        if approved:
            missing=[str(i) for i,(st,iid) in enumerate(states,1) if not c.execute("SELECT iu.id FROM local_items li JOIN item_units iu ON iu.id=li.default_purchase_unit_id AND iu.active=1 WHERE li.id=?",(iid,)).fetchone()]
            if missing: raise ValueError('لا يمكن اعتماد الفاتورة قبل تحديد وحدة الشراء الافتراضية للبنود: '+', '.join(missing))
        cur=c.execute('''INSERT INTO invoices(supplier_id,invoice_number,invoice_date,source_file,source_type,warehouse_number,subtotal_halalah,discount_halalah,vat_rate,vat_amount_halalah,total_amount_halalah,processing_status,notes)
                         VALUES(?,?,?,?,?,?,?,?,?,?,?,CASE WHEN ? THEN 'APPROVED' ELSE 'REVIEW_REQUIRED' END,?)''',
                      (supplier_id,invoice_number,invoice_date,source_file,source_type,warehouse_number,subtotal,discount,vat_rate,vat,total,approved,str(payload.get('notes') or '').strip() or None))
        invoice_id=cur.lastrowid; saved=price_rows=history_rows=learned_codes=learned_aliases=0; warnings=[]
        for i,x in enumerate(lines,1):
            state=_persist_line_state(c,supplier_id,x); qty,unit_price,net,vat_line,gross=_calculate_invoice_line_amounts(x,vat_rate); export_unit_name=None
            if approved and state['actual_item_id']:
                ur=c.execute("SELECT iu.unit_name FROM local_items li JOIN item_units iu ON iu.id=li.default_purchase_unit_id AND iu.active=1 WHERE li.id=?",(state['actual_item_id'],)).fetchone(); export_unit_name=ur['unit_name'] if ur else None
            cur2=c.execute('''INSERT INTO invoice_lines(invoice_id,line_number,supplier_item_code,raw_item_name,normalized_item_name,supplier_unit,normalized_unit,quantity,unit_price_halalah,net_amount_halalah,vat_amount_halalah,gross_amount_halalah,matched_local_item_id,matched_mapping_id,match_score,match_status,user_decision,export_unit_name)
                              VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                           (invoice_id,i,str(x.get('supplier_item_code') or '').strip() or None,str(x.get('raw_item_name') or '').strip(),norm(x.get('raw_item_name')),str(x.get('unit') or '').strip() or None,norm(x.get('unit')),qty,unit_price,net,vat_line,gross,state['actual_item_id'],state['mapping_id'],(x.get('match') or {}).get('score'),state['status'],state['decision'],export_unit_name))
            line_id=cur2.lastrowid; saved+=1
            _,lc,la,_did_learn=_save_learning_if_confirmed(c,supplier_id,invoice_id,line_id,x,state,warnings)
            learned_codes+=lc; learned_aliases+=la
            sug=x.get('_suggestion') or {}; sug_iid=_valid_local_item_for_context(sug.get('item_id'));
            # Suggestions live in the browser and may become stale after an item is changed/removed.
            # Never insert a stale suggestion id into match_history because that would violate its FK.
            if sug_iid and not c.execute('SELECT id FROM local_items WHERE id=?',(sug_iid,)).fetchone(): sug_iid=None
            sug_mid=_valid_mapping(c,supplier_id,sug.get('mapping_id'),sug_iid)
            c.execute('''INSERT INTO match_history(invoice_line_id,suggested_item_id,actual_item_id,suggested_mapping_id,actual_mapping_id,algorithm_score,decision,decision_source,notes)
                         VALUES(?,?,?,?,?,?,?,?,?)''',
                      (line_id,sug_iid,state['actual_item_id'],sug_mid,state['mapping_id'],sug.get('score',(x.get('match') or {}).get('score')),state['decision'],'USER' if x.get('_learning') else 'SYSTEM',(x.get('match') or {}).get('method')))
            history_rows+=1
            if state['accepted'] and unit_price>=0:
                factor=1
                if state['mapping_id']:
                    rr=c.execute('SELECT conversion_factor FROM supplier_item_mappings WHERE id=?',(state['mapping_id'],)).fetchone(); factor=float(rr['conversion_factor'] or 1) if rr else 1
                c.execute('''INSERT INTO supplier_item_prices(supplier_id,local_item_id,mapping_id,invoice_id,invoice_line_id,invoice_date,supplier_unit,conversion_factor,unit_price_halalah,quantity)
                             VALUES(?,?,?,?,?,?,?,?,?,?)''',(supplier_id,state['actual_item_id'],state['mapping_id'],invoice_id,line_id,invoice_date,str(x.get('unit') or '').strip() or None,factor,unit_price,qty)); price_rows+=1
        c.commit()
        return {'invoice_id':invoice_id,'invoice_number':invoice_number,'lines_saved':saved,'match_history_rows':history_rows,'price_history_rows':price_rows,
                'learned_code_rows':learned_codes,'learned_alias_rows':learned_aliases,'warnings':warnings,'subtotal_halalah':subtotal,'vat_amount_halalah':vat,'total_amount_halalah':total}
    except Exception:
        c.rollback(); raise
    finally: c.close()

def update_invoice(invoice_id,payload):
    """Update a review invoice; line matching/learning is committed with the edit transaction."""
    invoice_id=int(invoice_id or 0); lines=payload.get('lines') or []
    if not lines: raise ValueError('لا توجد بنود لتحديث الفاتورة.')
    c=dbconn()
    try:
        inv=c.execute('SELECT * FROM invoices WHERE id=?',(invoice_id,)).fetchone()
        if not inv: raise ValueError('الفاتورة غير موجودة.')
        if inv['processing_status']!='REVIEW_REQUIRED': raise ValueError('لا يمكن تعديل الفاتورة بعد اعتمادها.')
        supplier_id=int(payload.get('supplier_id') or inv['supplier_id']); invoice_number=str(payload.get('invoice_number') or inv['invoice_number']).strip(); invoice_date=str(payload.get('invoice_date') or inv['invoice_date']).strip(); warehouse_number=str(payload.get('warehouse_number') or inv['warehouse_number'] or '').strip() or None
        dup=c.execute('SELECT id FROM invoices WHERE supplier_id=? AND invoice_number=? AND id<>?',(supplier_id,invoice_number,invoice_id)).fetchone()
        if dup: raise ValueError(f'الفاتورة رقم {invoice_number} محفوظة مسبقًا لهذا المورد.')
        subtotal=sum(_halalah(x.get('net_amount') if x.get('net_amount') not in (None,'') else float(x.get('quantity',x.get('qty') or 0))*float(x.get('unit_price',0) or 0)) for x in lines)
        discount=max(0,int(payload.get('discount_halalah') or 0)); discount=min(discount,subtotal); taxable=subtotal-discount
        vat_rate=max(0,float(payload.get('vat_rate') if payload.get('vat_rate') not in (None,'') else inv['vat_rate'] or 15)); vat=max(0,int(round(taxable*vat_rate/100))); total=taxable+vat
        states=[]
        for x in lines:
            m=x.get('match') or {}; st=str(m.get('status') or 'UNMATCHED').upper(); iid=_valid_local_item_for_context((m.get('item') or {}).get('id')); iid=iid if iid and c.execute('SELECT id FROM local_items WHERE id=? AND active=1',(iid,)).fetchone() else None; states.append((st,iid))
        fully_approved=all(st in ('APPROVED','AUTO_MATCHED') and iid is not None for st,iid in states)
        if fully_approved:
            missing=[str(i) for i,(st,iid) in enumerate(states,1) if not c.execute("SELECT iu.id FROM local_items li JOIN item_units iu ON iu.id=li.default_purchase_unit_id AND iu.active=1 WHERE li.id=?",(iid,)).fetchone()]
            if missing: raise ValueError('لا يمكن اعتماد الفاتورة قبل تحديد وحدة الشراء الافتراضية للبنود: '+', '.join(missing))
        c.execute('UPDATE invoices SET supplier_id=?,invoice_number=?,invoice_date=?,warehouse_number=?,subtotal_halalah=?,discount_halalah=?,vat_rate=?,vat_amount_halalah=?,total_amount_halalah=?,processing_status=?,updated_at=CURRENT_TIMESTAMP WHERE id=?',
                  (supplier_id,invoice_number,invoice_date,warehouse_number,subtotal,discount,vat_rate,vat,total,'APPROVED' if fully_approved else 'REVIEW_REQUIRED',invoice_id))
        old_lines={int(r['id']):dict(r) for r in c.execute('SELECT * FROM invoice_lines WHERE invoice_id=?',(invoice_id,)).fetchall()}; kept=set(); warnings=[]; price_rows=history_rows=learned_codes=learned_aliases=0
        for i,x in enumerate(lines,1):
            lid=int(x.get('id') or 0); old=old_lines.get(lid)
            if not old: continue
            kept.add(lid); state=_persist_line_state(c,supplier_id,x); qty,unit_price,net,vat_line,gross=_calculate_invoice_line_amounts(x,vat_rate); export_unit_name=old.get('export_unit_name')
            if fully_approved and state['actual_item_id']:
                ur=c.execute("SELECT iu.unit_name FROM local_items li JOIN item_units iu ON iu.id=li.default_purchase_unit_id AND iu.active=1 WHERE li.id=?",(state['actual_item_id'],)).fetchone(); export_unit_name=ur['unit_name'] if ur else None
            elif not state['accepted']: export_unit_name=None
            c.execute('''UPDATE invoice_lines SET line_number=?,supplier_item_code=?,raw_item_name=?,normalized_item_name=?,supplier_unit=?,normalized_unit=?,quantity=?,unit_price_halalah=?,net_amount_halalah=?,vat_amount_halalah=?,gross_amount_halalah=?,matched_local_item_id=?,matched_mapping_id=?,match_score=?,match_status=?,user_decision=?,export_unit_name=?,updated_at=CURRENT_TIMESTAMP WHERE id=? AND invoice_id=?''',
                      (i,str(x.get('supplier_item_code') or '').strip() or None,str(x.get('raw_item_name') or '').strip(),norm(x.get('raw_item_name')),str(x.get('unit',x.get('supplier_unit') or '') or '').strip() or None,norm(x.get('unit',x.get('supplier_unit') or '')),qty,unit_price,net,vat_line,gross,state['actual_item_id'],state['mapping_id'],(x.get('match') or {}).get('score'),state['status'],state['decision'],export_unit_name,lid,invoice_id))
            _,lc,la,_did_learn=_save_learning_if_confirmed(c,supplier_id,invoice_id,lid,x,state,warnings); learned_codes+=lc; learned_aliases+=la
            sug=x.get('_suggestion') or {}; sug_iid=_valid_local_item_for_context(sug.get('item_id'));
            # Suggestions live in the browser and may become stale after an item is changed/removed.
            # Never insert a stale suggestion id into match_history because that would violate its FK.
            if sug_iid and not c.execute('SELECT id FROM local_items WHERE id=?',(sug_iid,)).fetchone(): sug_iid=None
            sug_mid=_valid_mapping(c,supplier_id,sug.get('mapping_id'),sug_iid)
            c.execute('''INSERT INTO match_history(invoice_line_id,suggested_item_id,actual_item_id,suggested_mapping_id,actual_mapping_id,algorithm_score,decision,decision_source,notes)
                         VALUES(?,?,?,?,?,?,?,?,?)''',(lid,sug_iid,state['actual_item_id'],sug_mid,state['mapping_id'],sug.get('score',(x.get('match') or {}).get('score')),state['decision'],'USER' if x.get('_learning') else 'SYSTEM',(x.get('match') or {}).get('method'))); history_rows+=1
            if state['accepted']:
                existing=c.execute('SELECT id FROM supplier_item_prices WHERE invoice_id=? AND invoice_line_id=? LIMIT 1',(invoice_id,lid)).fetchone()
                if existing:
                    c.execute('UPDATE supplier_item_prices SET local_item_id=?,mapping_id=?,invoice_date=?,supplier_unit=?,unit_price_halalah=?,quantity=? WHERE id=?',(state['actual_item_id'],state['mapping_id'],invoice_date,str(x.get('unit',x.get('supplier_unit') or '') or '').strip() or None,unit_price,qty,existing['id']))
                else:
                    c.execute('''INSERT INTO supplier_item_prices(supplier_id,local_item_id,mapping_id,invoice_id,invoice_line_id,invoice_date,supplier_unit,conversion_factor,unit_price_halalah,quantity)
                                 VALUES(?,?,?,?,?,?,?,?,?,?)''',(supplier_id,state['actual_item_id'],state['mapping_id'],invoice_id,lid,invoice_date,str(x.get('unit',x.get('supplier_unit') or '') or '').strip() or None,1,unit_price,qty))
                price_rows+=1
        for lid in set(old_lines)-kept: c.execute('DELETE FROM invoice_lines WHERE id=? AND invoice_id=?',(lid,invoice_id))
        c.commit()
        return {'invoice_id':invoice_id,'message':'تم تعديل الفاتورة بنجاح وإعادة حساب الخصم والضريبة والإجمالي.','price_history_rows':price_rows,'match_history_rows':history_rows,
                'learned_code_rows':learned_codes,'learned_alias_rows':learned_aliases,'warnings':warnings}
    except Exception:
        c.rollback(); raise
    finally: c.close()

def export_selected_items_xls(ids):
    ids=[int(x) for x in (ids or []) if int(x)>0]
    if not ids: raise ValueError('لم يتم تحديد أصناف.')
    c=dbconn()
    try:
        qs=','.join('?' for _ in ids)
        items=c.execute(f"SELECT * FROM local_items WHERE active=1 AND export_status='NEW' AND id IN ({qs}) ORDER BY CAST(group_code AS INTEGER),CAST(item_code AS INTEGER),item_code",ids).fetchall()
        if not items: raise ValueError('لا توجد أصناف جديدة غير مصدرة ضمن الاختيار.')
        data=[]
        for it in items:
            units=[dict(u) for u in c.execute("SELECT id,unit_name,conversion_factor,is_default FROM item_units WHERE local_item_id=? AND active=1 ORDER BY id",(it['id'],)).fetchall()]
            data.append({'group_code':it['group_code'],'item_code':it['item_code'],'item_name':it['item_name'],'item_type':it['item_type'],'units':units,'id':it['id'],'export_status':it['export_status'],'default_sale_unit_id':it['default_sale_unit_id'],'default_purchase_unit_id':it['default_purchase_unit_id']})
        blob=_xls_biff8_bytes(lambda path:_write_text_xlsx(path,['رقم المجموعة','رقم الصنف','الاسم المحلي','نوع الصنف','الوحدة','العبوة','وحدة البيع الافتراضية','وحدة الشراء الافتراضية'],_item_export_rows(data),'الأصناف المحددة'),'اصناف_محددة_للمتكامل')
        c.executemany("UPDATE local_items SET last_exported_at=CURRENT_TIMESTAMP,export_status='EXPORTED' WHERE id=? AND export_status='NEW'",[(x['id'],) for x in data]); c.commit()
        return blob,len(data)
    finally: c.close()

def confirm_selected_items_imported(ids):
    ids=[int(x) for x in (ids or []) if int(x)>0]
    if not ids: return 0
    c=dbconn()
    try:
        qs=','.join('?' for _ in ids)
        c.execute(f"UPDATE local_items SET imported_confirmed_at=CURRENT_TIMESTAMP,export_status='IMPORTED' WHERE export_status='EXPORTED' AND id IN ({qs})",ids)
        n=c.execute('SELECT changes()').fetchone()[0]; c.commit(); return n
    finally: c.close()

def confirm_items_imported():
    c=dbconn()
    try:
        c.execute("UPDATE local_items SET imported_confirmed_at=CURRENT_TIMESTAMP,export_status='IMPORTED' WHERE export_status='EXPORTED'"); n=c.execute("SELECT changes()").fetchone()[0]; c.commit(); return n
    finally: c.close()


def set_invoice_active(invoice_id, active):
    invoice_id=int(invoice_id or 0); c=dbconn()
    try:
        inv=c.execute('SELECT * FROM invoices WHERE id=?',(invoice_id,)).fetchone()
        if not inv: raise ValueError('الفاتورة غير موجودة.')
        if active:
            c.execute("UPDATE invoices SET active=1,deleted_at=NULL,updated_at=CURRENT_TIMESTAMP WHERE id=?",(invoice_id,)); msg='تمت استعادة الفاتورة.'
        else:
            c.execute("UPDATE invoices SET active=0,deleted_at=CURRENT_TIMESTAMP,updated_at=CURRENT_TIMESTAMP WHERE id=?",(invoice_id,)); msg='تم نقل الفاتورة إلى سلة المحذوفات.'
        c.commit(); return {'invoice_id':invoice_id,'message':msg}
    except Exception:
        c.rollback(); raise
    finally: c.close()



def invoice_detail(invoice_id):
    c=dbconn()
    try:
        inv=c.execute("""SELECT i.*,s.supplier_name,s.supplier_code
                         FROM invoices i JOIN suppliers s ON s.id=i.supplier_id WHERE i.id=?""",(int(invoice_id),)).fetchone()
        if not inv:
            raise ValueError('الفاتورة غير موجودة.')
        lines=c.execute("""SELECT il.*,
                                il.matched_local_item_id AS local_item_id,
                                li.item_code AS local_item_code,
                                li.item_name AS local_item_name,
                                li.source AS local_item_source,
                                mh.decision AS history_decision, mh.algorithm_score AS history_score, mh.decision_source
                         FROM invoice_lines il
                         LEFT JOIN local_items li ON li.id=il.matched_local_item_id
                         LEFT JOIN (
                           SELECT invoice_line_id,decision,algorithm_score,decision_source
                           FROM (
                             SELECT invoice_line_id,decision,algorithm_score,decision_source,
                                    ROW_NUMBER() OVER (PARTITION BY invoice_line_id ORDER BY id DESC) rn
                             FROM match_history
                           ) WHERE rn=1
                         ) mh ON mh.invoice_line_id=il.id
                         WHERE il.invoice_id=? ORDER BY il.line_number""",(int(invoice_id),)).fetchall()
        return {'invoice':dict(inv),'lines':[dict(r) for r in lines]}
    finally:
        c.close()

def invoice_history(limit=200, active_only=False):
    c=dbconn(); where='WHERE i.active=1' if active_only else ''
    rows=c.execute(f'''SELECT i.*,s.supplier_name,
                     (SELECT COUNT(*) FROM invoice_lines il WHERE il.invoice_id=i.id) line_count,
                     (SELECT COUNT(*) FROM invoice_lines il WHERE il.invoice_id=i.id AND il.match_status IN ('APPROVED','AUTO_MATCHED')) matched_count
                     FROM invoices i JOIN suppliers s ON s.id=i.supplier_id {where}
                     ORDER BY i.active DESC,i.invoice_date DESC,i.id DESC LIMIT ?''',(limit,)).fetchall()
    c.close(); return [dict(r) for r in rows]

backup_database()
# v2.22 invoice financial fields migration: discount amount + VAT percentage.
def migrate_invoice_financial_fields():
    c=dbconn()
    try:
        cols={r['name'] for r in c.execute("PRAGMA table_info(invoices)").fetchall()}
        if 'discount_halalah' not in cols:
            c.execute("ALTER TABLE invoices ADD COLUMN discount_halalah INTEGER DEFAULT 0")
        if 'vat_rate' not in cols:
            c.execute("ALTER TABLE invoices ADD COLUMN vat_rate NUMERIC DEFAULT 15")
        c.commit()
    finally:
        c.close()
init_item_management_schema()
migrate_invoice_financial_fields()

def migrate_item_export_and_warehouse_fields():
    c=dbconn()
    try:
        cols={r['name'] for r in c.execute("PRAGMA table_info(local_items)").fetchall()}
        if 'export_status' not in cols: c.execute("ALTER TABLE local_items ADD COLUMN export_status TEXT DEFAULT 'IMPORTED'")
        if 'last_exported_at' not in cols: c.execute("ALTER TABLE local_items ADD COLUMN last_exported_at TEXT")
        if 'imported_confirmed_at' not in cols: c.execute("ALTER TABLE local_items ADD COLUMN imported_confirmed_at TEXT")
        # الترحيل الأولي فقط عند إضافة العمود لأول مرة.
        # لا نعيد ضبط حالة الأصناف عند كل تشغيل للتطبيق.
        if 'export_status' not in cols:
            c.execute("UPDATE local_items SET export_status='NEW' WHERE source='USER_CREATED'")
            c.execute("UPDATE local_items SET export_status='IMPORTED' WHERE export_status IS NULL OR trim(export_status)=''")
        else:
            c.execute("UPDATE local_items SET export_status='IMPORTED' WHERE export_status IS NULL OR trim(export_status)=''")
        icols={r['name'] for r in c.execute("PRAGMA table_info(invoices)").fetchall()}
        if 'warehouse_number' not in icols: c.execute("ALTER TABLE invoices ADD COLUMN warehouse_number TEXT")
        if 'active' not in icols: c.execute("ALTER TABLE invoices ADD COLUMN active INTEGER NOT NULL DEFAULT 1")
        if 'deleted_at' not in icols: c.execute("ALTER TABLE invoices ADD COLUMN deleted_at TEXT")
        c.execute("CREATE INDEX IF NOT EXISTS ix_invoices_active ON invoices(active)")
        c.commit()
    finally: c.close()

migrate_item_export_and_warehouse_fields()

def migrate_schema_v38():
    c=dbconn()
    try:
        c.execute("DROP INDEX IF EXISTS ux_extraction_template_supplier_signature"); c.execute("DROP INDEX IF EXISTS ux_extraction_template_general_signature")
        c.execute("CREATE UNIQUE INDEX IF NOT EXISTS ux_extraction_template_supplier_signature ON invoice_extraction_templates(supplier_id,header_signature) WHERE supplier_id IS NOT NULL")
        c.execute("CREATE UNIQUE INDEX IF NOT EXISTS ux_extraction_template_general_signature ON invoice_extraction_templates(header_signature) WHERE supplier_id IS NULL")
        c.execute("CREATE TABLE IF NOT EXISTS extraction_corrections (id INTEGER PRIMARY KEY AUTOINCREMENT,supplier_id INTEGER,header_signature TEXT NOT NULL,field TEXT NOT NULL,original_column INTEGER,corrected_column INTEGER,learned_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,FOREIGN KEY(supplier_id) REFERENCES suppliers(id) ON DELETE SET NULL ON UPDATE CASCADE)")
        c.execute("CREATE INDEX IF NOT EXISTS ix_extraction_corrections_signature ON extraction_corrections(supplier_id,header_signature,field,learned_at DESC,id DESC)"); c.commit()
    finally:c.close()

migrate_schema_v38()

def migrate_schema_v39():
    c=dbconn()
    try:
        lcols={r['name'] for r in c.execute("PRAGMA table_info(local_items)").fetchall()}
        if 'default_sale_unit_id' not in lcols:
            c.execute("ALTER TABLE local_items ADD COLUMN default_sale_unit_id INTEGER")
        if 'default_purchase_unit_id' not in lcols:
            c.execute("ALTER TABLE local_items ADD COLUMN default_purchase_unit_id INTEGER")
        icols={r['name'] for r in c.execute("PRAGMA table_info(invoice_lines)").fetchall()}
        if 'export_unit_name' not in icols:
            c.execute("ALTER TABLE invoice_lines ADD COLUMN export_unit_name TEXT")
        c.execute("CREATE INDEX IF NOT EXISTS ix_local_items_default_sale_unit ON local_items(default_sale_unit_id)")
        c.execute("CREATE INDEX IF NOT EXISTS ix_local_items_default_purchase_unit ON local_items(default_purchase_unit_id)")
        c.execute("CREATE INDEX IF NOT EXISTS ix_invoice_lines_export_unit ON invoice_lines(export_unit_name)")
        c.commit()
    finally:
        c.close()

migrate_schema_v39()

def migrate_schema_v40():
    """Learning governance: soft-disable aliases, conflict log, and auditable learning history."""
    c=dbconn()
    try:
        acols={r['name'] for r in c.execute("PRAGMA table_info(supplier_item_aliases)").fetchall()}
        if 'active' not in acols:
            c.execute("ALTER TABLE supplier_item_aliases ADD COLUMN active INTEGER NOT NULL DEFAULT 1")
        c.execute("UPDATE supplier_item_aliases SET active=1 WHERE active IS NULL")
        c.execute("CREATE INDEX IF NOT EXISTS ix_alias_supplier_active_name ON supplier_item_aliases(supplier_id,active,normalized_name)")
        c.execute("""CREATE TABLE IF NOT EXISTS supplier_mapping_conflicts (
                     id INTEGER PRIMARY KEY AUTOINCREMENT,
                     supplier_id INTEGER NOT NULL,
                     supplier_item_code TEXT,
                     normalized_code TEXT,
                     local_item_id INTEGER,
                     existing_local_item_id INTEGER,
                     invoice_id INTEGER,
                     invoice_line_id INTEGER,
                     reason TEXT,
                     status TEXT NOT NULL DEFAULT 'OPEN',
                     resolved_at TEXT,
                     resolution_note TEXT,
                     created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                     FOREIGN KEY(supplier_id) REFERENCES suppliers(id) ON DELETE RESTRICT ON UPDATE CASCADE,
                     FOREIGN KEY(local_item_id) REFERENCES local_items(id) ON DELETE SET NULL ON UPDATE CASCADE,
                     FOREIGN KEY(existing_local_item_id) REFERENCES local_items(id) ON DELETE SET NULL ON UPDATE CASCADE,
                     FOREIGN KEY(invoice_id) REFERENCES invoices(id) ON DELETE SET NULL ON UPDATE CASCADE,
                     FOREIGN KEY(invoice_line_id) REFERENCES invoice_lines(id) ON DELETE SET NULL ON UPDATE CASCADE
                   )""")
        c.execute("CREATE INDEX IF NOT EXISTS ix_supplier_mapping_conflicts_supplier_code ON supplier_mapping_conflicts(supplier_id,normalized_code,status)")
        c.execute("""CREATE TABLE IF NOT EXISTS supplier_learning_history (
                     id INTEGER PRIMARY KEY AUTOINCREMENT,
                     entity_type TEXT NOT NULL,
                     entity_id INTEGER,
                     supplier_id INTEGER,
                     local_item_id INTEGER,
                     action TEXT NOT NULL,
                     before_json TEXT,
                     after_json TEXT,
                     invoice_id INTEGER,
                     invoice_line_id INTEGER,
                     source TEXT,
                     notes TEXT,
                     created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                     FOREIGN KEY(supplier_id) REFERENCES suppliers(id) ON DELETE SET NULL ON UPDATE CASCADE,
                     FOREIGN KEY(local_item_id) REFERENCES local_items(id) ON DELETE SET NULL ON UPDATE CASCADE,
                     FOREIGN KEY(invoice_id) REFERENCES invoices(id) ON DELETE SET NULL ON UPDATE CASCADE,
                     FOREIGN KEY(invoice_line_id) REFERENCES invoice_lines(id) ON DELETE SET NULL ON UPDATE CASCADE
                   )""")
        c.execute("CREATE INDEX IF NOT EXISTS ix_learning_history_supplier_date ON supplier_learning_history(supplier_id,created_at DESC,id DESC)")
        c.execute("CREATE INDEX IF NOT EXISTS ix_learning_history_item_date ON supplier_learning_history(local_item_id,created_at DESC,id DESC)")
        c.execute("CREATE INDEX IF NOT EXISTS ix_learning_history_invoice_line ON supplier_learning_history(invoice_line_id,created_at DESC,id DESC)")
        c.execute("CREATE TABLE IF NOT EXISTS app_meta (key TEXT PRIMARY KEY,value TEXT)")
        c.execute("INSERT INTO app_meta(key,value) VALUES('schema_version',?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",('2.53',))
        c.commit()
    finally:
        c.close()

def log_learning_history(c, entity_type, entity_id, supplier_id, local_item_id, action,
                         before=None, after=None, invoice_id=None, invoice_line_id=None,
                         source='SYSTEM', notes=''):
    c.execute("""INSERT INTO supplier_learning_history
                 (entity_type,entity_id,supplier_id,local_item_id,action,before_json,after_json,invoice_id,invoice_line_id,source,notes)
                 VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
              (entity_type,entity_id,supplier_id,local_item_id,action,
               json.dumps(before,ensure_ascii=False,default=str) if before is not None else None,
               json.dumps(after,ensure_ascii=False,default=str) if after is not None else None,
               invoice_id,invoice_line_id,source,notes or None))

def learning_history(limit=300, supplier_id=None, local_item_id=None, entity_type=None, action=None, source=None):
    c=dbconn()
    try:
        where=[]; params=[]
        if supplier_id:
            where.append('h.supplier_id=?'); params.append(int(supplier_id))
        if local_item_id:
            where.append('h.local_item_id=?'); params.append(int(local_item_id))
        if entity_type:
            where.append('h.entity_type=?'); params.append(str(entity_type).upper())
        if action:
            where.append('h.action=?'); params.append(str(action).upper())
        if source:
            where.append('h.source=?'); params.append(str(source).upper())
        clause=(' WHERE '+' AND '.join(where)) if where else ''
        params.append(max(1,min(int(limit or 300),1000)))
        rows=c.execute(f"""SELECT h.*,s.supplier_name,li.item_code,li.item_name,
                              i.invoice_number,
                              CASE WHEN h.entity_type='MAPPING' THEN 'كود مورد'
                                   WHEN h.entity_type='ALIAS' THEN 'اسم بديل'
                                   WHEN h.entity_type='CONFLICT' THEN 'تعارض' ELSE h.entity_type END AS entity_label
                            FROM supplier_learning_history h
                            LEFT JOIN suppliers s ON s.id=h.supplier_id
                            LEFT JOIN local_items li ON li.id=h.local_item_id
                            LEFT JOIN invoices i ON i.id=h.invoice_id
                            {clause} ORDER BY h.id DESC LIMIT ?""",params).fetchall()
        return [dict(r) for r in rows]
    finally:
        c.close()

migrate_schema_v40()

def set_db_schema_version():
    c = dbconn()
    try:
        c.execute(f'PRAGMA user_version = {APP_DB_SCHEMA_VERSION}')
        c.execute("CREATE TABLE IF NOT EXISTS app_meta (key TEXT PRIMARY KEY,value TEXT)")
        c.execute("INSERT INTO app_meta(key,value) VALUES('schema_version',?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", ('2.53',))
        c.commit()
    finally:
        c.close()

set_db_schema_version()

def create_backup_zip():
    import zipfile, datetime
    backup_dir=BASE/'نسخ_احتياطية'; backup_dir.mkdir(exist_ok=True)
    stamp=datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    zip_path=backup_dir/f'نسخة_احتياطية_{stamp}.zip'
    c=dbconn()
    try:
        c.execute('PRAGMA wal_checkpoint(TRUNCATE)')
    finally: c.close()
    with zipfile.ZipFile(zip_path,'w',zipfile.ZIP_DEFLATED) as z:
        z.write(DB,'قاعدة_البيانات.sqlite')
    return zip_path

def restore_backup_zip(data):
    import io, zipfile, tempfile, sqlite3, datetime
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        names=[n for n in z.namelist() if n.lower().endswith('.sqlite') and not n.endswith('/')]
        if not names: raise ValueError('ملف النسخة الاحتياطية غير صالح: لم توجد قاعدة بيانات SQLite.')
        raw=z.read(names[0])
    tmp=BASE/'استعادة_مؤقتة.sqlite'
    tmp.write_bytes(raw)
    try:
        tc=sqlite3.connect(tmp)
        ok=tc.execute('PRAGMA integrity_check').fetchone()[0]
        tc.close()
        if str(ok).lower()!='ok': raise ValueError('قاعدة البيانات داخل النسخة الاحتياطية غير سليمة.')
        # Safety backup of current DB before replacement.
        create_backup_zip()
        if DB.exists():
            try:
                c=dbconn(); c.execute('PRAGMA wal_checkpoint(TRUNCATE)'); c.close()
            except Exception: pass
        shutil.copy2(tmp,DB)
        return True
    finally:
        try: tmp.unlink()
        except: pass

class Handler(BaseHTTPRequestHandler):
    def send_json(self,obj,status=200):
        data=json.dumps(obj,ensure_ascii=False).encode("utf-8")
        self.send_response(status); self.send_header("Content-Type","application/json; charset=utf-8")
        self.send_header("Content-Length",str(len(data))); self.end_headers(); self.wfile.write(data)
    def do_GET(self):
        p=urlparse(self.path).path
        if p=="/api/items": return self.send_json({"items":items()})
        if p=="/api/items/all":
            q=urlparse(self.path).query; import urllib.parse; params=urllib.parse.parse_qs(q); inc=(params.get('include_deleted') or ['0'])[0]=='1'
            return self.send_json({"items":items_all(inc)})
        if p=="/api/item/card":
            q=urlparse(self.path).query; import urllib.parse; params=urllib.parse.parse_qs(q); iid=int((params.get('id') or ['0'])[0] or 0)
            card=item_card(iid)
            return self.send_json(card or {"error":"الصنف غير موجود"},404 if not card else 200)
        if p=="/api/item/units": return self.send_json({"units":managed_units()})
        if p=="/api/item/groups": return self.send_json({"groups":managed_groups()})
        if p=="/api/groups": return self.send_json({"groups":groups()})
        if p=="/api/item/next-code":
            q=urlparse(self.path).query
            import urllib.parse
            params=urllib.parse.parse_qs(q); gc=(params.get('group_code') or [''])[0]
            try: return self.send_json({"ok":True,"group_code":gc,"next_code":next_item_code(gc)})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/item/suggest-group":
            q=urlparse(self.path).query
            import urllib.parse
            params=urllib.parse.parse_qs(q); name=(params.get('name') or [''])[0]
            return self.send_json(suggest_group_for_name(name))
        if p=="/api/suppliers": return self.send_json({"suppliers":suppliers()})
        if p=="/api/suppliers/all":
            c=dbconn(); rows=c.execute("SELECT id,supplier_code,supplier_name,vat_number,commercial_register,phone,active FROM suppliers ORDER BY supplier_name").fetchall(); c.close()
            return self.send_json({"suppliers":[dict(r) for r in rows]})
        if p=="/api/export/items/full" or p=="/api/export/items/simple":
            try:
                data=export_items_xlsx(simple=(p.endswith('/simple')))
                filename='قاعدة_الأصناف_متعدد_الأوراق.xlsx' if p.endswith('/full') else 'قاعدة_الأصناف_مبسط.xlsx'
                import urllib.parse
                self.send_response(200); self.send_header("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", "attachment; filename*=UTF-8''"+urllib.parse.quote(filename))
                self.send_header("Content-Length",str(len(data))); self.end_headers(); self.wfile.write(data); return
            except Exception as e: return self.send_json({"error":str(e)},500)
        if p=="/api/items/new":
            try:
                c=dbconn()
                rows=c.execute("SELECT li.id,li.item_code,li.group_code,li.item_name,li.item_type,li.main_unit,li.source,li.export_status,li.last_exported_at,li.created_at,li.updated_at,COUNT(iu.id) unit_count,MAX(CASE WHEN iu.normalized_unit=? THEN iu.conversion_factor END) pack_size FROM local_items li LEFT JOIN item_units iu ON iu.local_item_id=li.id AND iu.active=1 WHERE li.active=1 AND li.export_status IN ('NEW','EXPORTED') GROUP BY li.id ORDER BY CASE WHEN li.export_status='NEW' THEN 0 ELSE 1 END,CAST(li.group_code AS INTEGER),CAST(li.item_code AS INTEGER),li.item_code",(norm('كرتون'),)).fetchall(); c.close()
                return self.send_json({'items':[dict(r) for r in rows]})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},500)
        if p=="/api/items/export/new":
            try:
                data,n=export_new_items_xls(); import urllib.parse; filename='اصناف_جديدة_للمتكامل.xls'; self.send_response(200); self.send_header('Content-Type','application/vnd.ms-excel'); self.send_header('X-Export-Count',str(n)); self.send_header('Content-Disposition',"attachment; filename*=UTF-8''"+urllib.parse.quote(filename)); self.send_header('Content-Length',str(len(data))); self.end_headers(); self.wfile.write(data); return
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},500)
        if p=="/api/items/export/updated":
            try:
                data,n=export_updated_items_xls(); import urllib.parse; filename='تعديلات_الأصناف_للمتكامل.xls'; self.send_response(200); self.send_header('Content-Type','application/vnd.ms-excel'); self.send_header('Content-Disposition',"attachment; filename*=UTF-8''"+urllib.parse.quote(filename)); self.send_header('Content-Length',str(len(data))); self.end_headers(); self.wfile.write(data); return
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},500)
        if p=="/api/items/export/confirm":
            try: return self.send_json({'ok':True,'count':confirm_items_imported()})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},500)
        if p=="/api/invoices": return self.send_json({"invoices":invoice_history()})
        if p.startswith("/api/invoices/export/"):
            try:
                invoice_id=int(p.rsplit('/',1)[-1]); data,num=export_invoice_xls(invoice_id)
                import urllib.parse
                filename=f'فاتورة_{num}.xls'
                self.send_response(200); self.send_header('Content-Type','application/vnd.ms-excel')
                self.send_header('Content-Disposition',"attachment; filename*=UTF-8''"+urllib.parse.quote(filename))
                self.send_header('Content-Length',str(len(data))); self.end_headers(); self.wfile.write(data); return
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},500)
        if p.startswith("/api/invoices/") and p.endswith("/rematch"):
            try:
                invoice_id=int(p.split('/')[3])
                return self.send_json({'ok':True,**rematch_invoice(invoice_id)})
            except Exception as e:
                return self.send_json({"ok":False,"message":str(e)},400)
        if p.startswith("/api/invoices/") and p != "/api/invoices/save":
            try:
                invoice_id=int(p.rsplit('/',1)[-1])
                return self.send_json(invoice_detail(invoice_id))
            except Exception as e:
                return self.send_json({"ok":False,"message":str(e)},404)
        if p=="/api/mappings": return self.send_json({"mappings":mappings()})
        if p=="/api/aliases": return self.send_json({"aliases":aliases()})
        if p=="/api/learning/history":
            try:
                import urllib.parse
                q=urlparse(self.path).query
                params=urllib.parse.parse_qs(q)
                sid=int((params.get('supplier_id') or ['0'])[0] or 0)
                iid=int((params.get('local_item_id') or ['0'])[0] or 0)
                limit=int((params.get('limit') or ['300'])[0] or 300)
                entity_type=(params.get('entity_type') or [''])[0].strip() or None
                action=(params.get('action') or [''])[0].strip() or None
                source=(params.get('source') or [''])[0].strip() or None
                return self.send_json({'history':learning_history(limit,sid or None,iid or None,entity_type,action,source)})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        if p=="/api/learning/conflicts":
            try:
                import urllib.parse
                params=urllib.parse.parse_qs(q); status=(params.get('status') or [''])[0] or None
                return self.send_json({'conflicts':supplier_mapping_conflicts(int((params.get('limit') or ['300'])[0] or 300),status)})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        if p=="/api/export/mappings":
            try:
                data=export_mappings_xlsx()
                self.send_response(200)
                self.send_header("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", "attachment; filename*=UTF-8''%D8%AC%D8%AF%D9%88%D9%84_%D8%B1%D8%A8%D8%B7_%D8%A7%D9%84%D8%A3%D8%B5%D9%86%D8%A7%D9%81.xlsx")
                self.send_header("Content-Length",str(len(data))); self.end_headers(); self.wfile.write(data); return
            except Exception as e:
                return self.send_json({"error":str(e)},500)
        if p=="/":
            self.path="/index.html"
        raw_path=unquote(urlparse(self.path).path)
        f=(BASE/raw_path.lstrip("/")).resolve()
        try:
            f.relative_to(BASE)
        except ValueError:
            return self.send_error(403,"Forbidden")
        if f.exists() and f.is_file():
            data=f.read_bytes(); self.send_response(200)
            self.send_header("Content-Type","text/html; charset=utf-8" if f.suffix==".html" else "application/octet-stream")
            self.send_header("Content-Length",str(len(data))); self.end_headers(); self.wfile.write(data); return
        self.send_error(404)
    def do_POST(self):
        p=urlparse(self.path).path
        if p=="/api/items/export/selected":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); data,count=export_selected_items_xls(payload.get('ids') or [])
                import urllib.parse
                filename='اصناف_محددة_للمتكامل.xls'; self.send_response(200); self.send_header('Content-Type','application/vnd.ms-excel'); self.send_header('X-Export-Count',str(count)); self.send_header('Content-Disposition',"attachment; filename*=UTF-8''"+urllib.parse.quote(filename)); self.send_header('Content-Length',str(len(data))); self.end_headers(); self.wfile.write(data); return
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        if p=="/api/items/export/confirm-selected":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); return self.send_json({'ok':True,'count':confirm_selected_items_imported(payload.get('ids') or [])})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        if p=="/api/invoices/delete" or p=="/api/invoices/restore":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); iid=int(payload.get('id') or 0); active=p.endswith('restore')
                out=set_invoice_active(iid,active); return self.send_json({'ok':True,**out})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        if p=="/api/backup/create":
            try:
                path=create_backup_zip(); import urllib.parse
                data=path.read_bytes(); self.send_response(200); self.send_header('Content-Type','application/zip'); self.send_header('Content-Disposition',"attachment; filename*=UTF-8''"+urllib.parse.quote(path.name)); self.send_header('Content-Length',str(len(data))); self.end_headers(); self.wfile.write(data); return
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},500)
        if p=="/api/backup/restore":
            try:
                ctype=self.headers.get('Content-Type',''); length=int(self.headers.get('Content-Length','0'))
                max_restore=500*1024*1024
                if length > max_restore:
                    return self.send_json({'ok':False,'message':f'حجم ملف الاستعادة أكبر من الحد المسموح ({max_restore//(1024*1024)} MB).'},413)
                body=self.rfile.read(length)
                m=re.search(r'boundary=(?:"([^"]+)"|([^;]+))',ctype)
                if not m: return self.send_json({'ok':False,'message':'ملف الاستعادة غير صالح.'},400)
                boundary=(m.group(1) or m.group(2)).encode(); raw=None
                for part in body.split(b'--'+boundary):
                    if b'Content-Disposition' not in part: continue
                    if b'filename=' not in part: continue
                    if b'\r\n\r\n' not in part: continue
                    raw=part.split(b'\r\n\r\n',1)[1].rsplit(b'\r\n',1)[0]; break
                if raw is None: return self.send_json({'ok':False,'message':'لم يتم استلام ملف النسخة الاحتياطية.'},400)
                restore_backup_zip(raw); return self.send_json({'ok':True,'message':'تمت استعادة النسخة الاحتياطية بنجاح. أغلق التطبيق وأعد تشغيله قبل متابعة العمل.'})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        if p=="/api/items/update":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); return self.send_json({"ok":True,**update_item(payload),"message":"تم تعديل الصنف وتسجيل العملية."})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/items/delete-pending-export":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); iid=int(payload.get('id') or 0); c=dbconn(); r=c.execute("SELECT export_status,item_name FROM local_items WHERE id=?",(iid,)).fetchone(); c.close()
                if not r: raise ValueError('الصنف غير موجود.')
                if r['export_status']!='EXPORTED': raise ValueError('يمكن حذف الصنف من هذه الشاشة فقط عندما يكون مصدّرًا بانتظار تأكيد الاستيراد.')
                out=set_item_active(iid,False); return self.send_json({"ok":True,"item":out,"message":"تم نقل الصنف إلى سلة الأصناف."})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/items/delete" or p=="/api/items/restore":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); iid=int(payload.get('id') or 0); active=p.endswith('restore')
                return self.send_json({"ok":True,"item":set_item_active(iid,active),"message":"تمت استعادة الصنف." if active else "تم نقل الصنف إلى سلة المحذوفات."})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/items/undo":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); return self.send_json({"ok":True,**undo_last_item(int(payload.get('id') or 0)),"message":"تم التراجع عن آخر عملية."})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/item/unit/add":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); return self.send_json({"ok":True,**add_item_unit(payload),"message":"تمت إضافة الوحدة."})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/item/unit/update":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); return self.send_json({"ok":True,**update_item_unit(payload),"message":"تم تعديل الوحدة."})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/item/unit/toggle":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); return self.send_json({"ok":True,**toggle_item_unit(int(payload.get('id') or 0),bool(payload.get('active'))),"message":"تم تحديث حالة الوحدة."})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/item/manager/unit/add":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); return self.send_json({"ok":True,**add_managed_unit(payload),"message":"تمت إضافة الوحدة."})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/item/manager/unit/update":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); return self.send_json({"ok":True,**update_managed_unit(payload),"message":"تم تعديل اسم الوحدة."})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/item/manager/unit/toggle":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); return self.send_json({"ok":True,**toggle_managed_unit(payload.get('unit_name'),bool(payload.get('active'))),"message":"تم تحديث حالة الوحدة."})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/item/manager/group/add":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); return self.send_json({"ok":True,**add_managed_group(payload),"message":"تمت إضافة المجموعة."})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/item/manager/group/update":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); return self.send_json({"ok":True,**update_managed_group(payload),"message":"تم تعديل اسم المجموعة."})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/item/manager/group/toggle":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8")); return self.send_json({"ok":True,**toggle_managed_group(payload.get('group_code'),bool(payload.get('active'))),"message":"تم تحديث حالة المجموعة."})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/items/create":
            c=None
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8"))
                item=create_local_item(payload)
                return self.send_json({"ok":True,"item":item,"message":f"تم إنشاء الصنف {item['item_code']} بنجاح."})
            except Exception as e:
                return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/items/create-from-invoice":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8"))
                out=save_new_item_from_invoice(payload)
                return self.send_json({"ok":True,**out,"message":f"تم إنشاء الصنف {out['item']['item_code']} وربطه بالمورد وحفظ اسم المورد كاسم بديل."})
            except Exception as e:
                return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/supplier/update":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8"))
                supplier=update_supplier(payload)
                return self.send_json({"ok":True,"supplier":supplier,"message":"تم تعديل بيانات المورد وتسجيل العملية."})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/suppliers":
            c=None
            try:
                n=int(self.headers.get("Content-Length","0")); body=self.rfile.read(n)
                payload=json.loads(body.decode("utf-8"))
                name=str(payload.get("supplier_name") or "").strip()
                code=str(payload.get("supplier_code") or "").strip() or None
                vat=str(payload.get("vat_number") or "").strip() or None
                cr=str(payload.get("cr_number") or payload.get("commercial_register") or "").strip() or None
                phone=str(payload.get("phone") or "").strip() or None
                if not name: return self.send_json({"ok":False,"message":"اسم المورد مطلوب."},400)
                c=dbconn(); sn=norm(name)
                ex=c.execute("SELECT id,supplier_code,supplier_name FROM suppliers WHERE normalized_name=? LIMIT 1",(sn,)).fetchone()
                if ex:
                    c.close(); return self.send_json({"ok":False,"message":"المورد موجود بالفعل في قاعدة البيانات."},409)
                if code:
                    ex=c.execute("SELECT id FROM suppliers WHERE supplier_code=? LIMIT 1",(code,)).fetchone()
                    if ex:
                        c.close(); return self.send_json({"ok":False,"message":"كود المورد مستخدم بالفعل لمورد آخر."},409)
                cur=c.execute("INSERT INTO suppliers(supplier_code,supplier_name,normalized_name,vat_number,commercial_register,phone) VALUES(?,?,?,?,?,?)",(code,name,sn,vat,cr,phone))
                sid=cur.lastrowid; c.commit(); c.close()
                return self.send_json({"ok":True,"supplier":{"id":sid,"supplier_code":code or "","supplier_name":name,"phone":phone or ""}})
            except Exception as e:
                try: c.close()
                except: pass
                return self.send_json({"ok":False,"message":str(e)},500)
        if p=="/api/extraction/template":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8"))
                return self.send_json({"ok":True,"template":extraction_template_save(payload)})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/extraction/template/get":
            try:
                n=int(self.headers.get("Content-Length","0")); payload=json.loads(self.rfile.read(n).decode("utf-8"))
                t=extraction_template_get(int(payload.get('supplier_id') or 0),str(payload.get('header_signature') or ''))
                return self.send_json({"ok":True,"template":t})
            except Exception as e: return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/upload":
            try:
                ctype=self.headers.get("Content-Type","")
                length=int(self.headers.get("Content-Length","0"))
                max_upload=100*1024*1024
                if length > max_upload:
                    return self.send_json({"error":f"حجم الملف أكبر من الحد المسموح ({max_upload//(1024*1024)} MB)."},413)
                body=self.rfile.read(length)
                # Minimal multipart/form-data parser (sufficient for browser FormData with one file).
                m=re.search(r'boundary=(?:"([^"]+)"|([^;]+))',ctype)
                if not m: return self.send_json({"error":"Invalid multipart upload"},400)
                boundary=(m.group(1) or m.group(2)).encode()
                parts=body.split(b"--"+boundary)
                data=None; filename=""; 
                for part in parts:
                    if b"Content-Disposition" not in part: continue
                    mm=re.search(br'filename="([^"]*)"',part)
                    if not mm: continue
                    filename=mm.group(1).decode("utf-8","replace")
                    head,content=part.split(b"\r\n\r\n",1)
                    data=content.rsplit(b"\r\n",1)[0]
                    break
                if data is None: return self.send_json({"error":"No file received"},400)
                ext=Path(filename).suffix.lower()
                if ext in (".xlsx",".xlsm"):
                    prepared=prepare_xlsx_workbook(data)
                    structure=prepared["structure"]
                    template=None
                    pmeta=structure.get("primary") or {}
                    invmeta=pmeta.get("invoice") or {}
                    supplier_name=str(invmeta.get("supplier_name") or "").strip()
                    if supplier_name:
                        try:
                            c=dbconn()
                            sr=c.execute("SELECT id,supplier_name FROM suppliers WHERE normalized_name=? LIMIT 1",(norm(supplier_name),)).fetchone()
                            c.close()
                            if sr and pmeta.get("signature"):
                                template=extraction_template_get(int(sr["id"]),str(pmeta["signature"]))
                        except Exception:
                            template=None
                    lines=parse_xlsx_bytes(data,template=template,analyzed=prepared)
                    structure["template_used"]={"id":template.get("id"),"name":template.get("template_name")} if template else None
                    if pmeta.get("invoice"):
                        structure["invoice"] = pmeta.get("invoice")
                    if not lines:
                        return self.send_json({"ok":False,"type":"excel","filename":filename,"lines":[],"extraction":structure,
                                               "message":"تم فتح ملف Excel لكن لم يتم التعرف على صفوف بنود. راجع تحليل الورقة والـHeader.",
                                               "detail":"المحرك لم يجد جدول أصناف بدرجة ثقة كافية؛ لم يتم التخمين حفاظًا على صحة البيانات."},422)
                    return self.send_json({"ok":True,"type":"excel","filename":filename,"lines":lines,"extraction":structure,
                                           "message":f"تمت قراءة Excel: {len(lines)} بند — محرك الاستخراج الدلالي v2.",
                                           "template_used":bool(template)})
                if ext==".csv" or ext==".txt":
                    lines=parse_csv_bytes(data)
                    return self.send_json({"ok":True,"type":"csv","filename":filename,"lines":lines,
                                           "message":f"تمت قراءة CSV فعليًا: {len(lines)} بند."})
                if ext in (".pdf",".png",".jpg",".jpeg",".webp",".bmp",".tif",".tiff"):
                    try:
                        text=ocr_file(data,filename)
                        parsed=parse_ocr_invoice_text(text)
                        return self.send_json({"ok":True,"type":"ocr","filename":filename,"text":text,"lines":parsed,
                                               "message":f"تمت قراءة PDF/الصورة محليًا عبر OCR: {len(parsed)} بند قابل للمراجعة. راجع البنود ثم شغّل المطابقة الذاتية."})
                    except Exception as e:
                        return self.send_json({"ok":False,"type":"ocr","filename":filename,
                                               "message":"تعذر تشغيل OCR. تأكد من تثبيت Tesseract OCR، وأن لغة العربية ara مثبتة.","detail":str(e)},500)
                return self.send_json({"error":"صيغة الملف غير مدعومة"},400)
            except Exception as e:
                return self.send_json({"error":str(e)},500)
        n=int(self.headers.get("Content-Length","0"))
        max_bytes = 500*1024*1024 if p=="/api/backup/restore" else 100*1024*1024
        if n > max_bytes:
            return self.send_json({"error":f"حجم الملف أكبر من الحد المسموح ({max_bytes//(1024*1024)} MB)."},413)
        body=self.rfile.read(n)
        if p=="/api/match":
            try:
                payload=json.loads(body.decode("utf-8"))
                lines=payload.get("lines",[])
                supplier_id=payload.get("supplier_id")
                out=match_lines_bulk(lines, supplier_id)
                return self.send_json({"lines":out})
            except Exception as e:
                return self.send_json({"error":str(e)},500)
        if p.startswith('/api/invoices/update/'):
            try:
                invoice_id=int(p.rsplit('/',1)[-1]); payload=json.loads(body.decode('utf-8')); out=update_invoice(invoice_id,payload); return self.send_json({'ok':True,**out,'message':'تم تعديل الفاتورة بنجاح وإعادة حساب الخصم والضريبة والإجمالي.'})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        if p=="/api/invoices/save":
            try:
                payload=json.loads(body.decode("utf-8"))
                out=save_invoice(payload)
                return self.send_json({"ok":True,**out,"message":f"تم حفظ الفاتورة رقم {out['invoice_number']} وربط سجل المطابقة وتاريخ الأسعار."})
            except Exception as e:
                return self.send_json({"ok":False,"message":str(e)},400)
        if p=="/api/learning/mapping/update":
            try:
                payload=json.loads(body.decode('utf-8')); return self.send_json({'ok':True,'mapping':update_learning_mapping(payload),'message':'تم تعديل رابط المورد وتسجيل العملية في سجل التعلم.'})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        if p=="/api/learning/mapping/toggle":
            try:
                payload=json.loads(body.decode('utf-8')); active=bool(payload.get('active')); mapping=toggle_learning_mapping(int(payload.get('id') or 0),active); return self.send_json({'ok':True,'mapping':mapping,'message':'تم تفعيل الربط.' if active else 'تم تعطيل الربط دون حذف سجله.'})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        if p=="/api/learning/alias/update":
            try:
                payload=json.loads(body.decode('utf-8')); return self.send_json({'ok':True,'alias':update_learning_alias(payload),'message':'تم تعديل الاسم البديل وتسجيل العملية.'})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        if p=="/api/learning/alias/toggle":
            try:
                payload=json.loads(body.decode('utf-8')); active=bool(payload.get('active')); alias=toggle_learning_alias(int(payload.get('id') or 0),active); return self.send_json({'ok':True,'alias':alias,'message':'تم تفعيل الاسم البديل.' if active else 'تم تعطيل الاسم البديل دون حذفه.'})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        if p=="/api/learning/conflict/resolve":
            try:
                payload=json.loads(body.decode('utf-8')); out=_resolve_learning_conflict(int(payload.get('id') or 0),payload.get('note') or ''); return self.send_json({'ok':True,'conflict':out,'message':'تم إغلاق التعارض وتسجيل قرار المراجعة.'})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        if p=="/api/link":
            # Backward-compatible endpoint: invoice linking is now local-only. Learning is
            # committed by /api/invoices/save or /api/invoices/update after final review.
            try:
                payload=json.loads(body.decode('utf-8')); line=payload.get('line',{}); local_item_id=int(payload.get('local_item_id') or 0)
                if not local_item_id: return self.send_json({'ok':False,'message':'اختر الصنف المحلي أولاً.'},400)
                return self.send_json({'ok':True,'deferred':True,'mapping_id':None,'message':'تم تسجيل اختيار الربط مؤقتًا داخل الفاتورة. سيتم حفظ التعلم عند حفظ الفاتورة.'})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        if p=="/api/alias/delete":
            try:
                payload=json.loads(body.decode('utf-8')); aid=int(payload.get('alias_id') or 0)
                if not aid: return self.send_json({'ok':False,'message':'معرف الاسم البديل مطلوب.'},400)
                out=toggle_learning_alias(aid,False)
                return self.send_json({'ok':True,'alias':out,'message':'تم تعطيل الاسم البديل دون حذفه نهائيًا. يمكن إعادة تفعيله من إدارة التعلم.'})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        if p=="/api/supplier/deactivate":
            try:
                payload=json.loads(body.decode("utf-8")); sid=int(payload.get("supplier_id") or 0)
                if not sid: return self.send_json({"ok":False,"message":"معرف المورد مطلوب."},400)
                c=dbconn(); c.execute("UPDATE suppliers SET active=0,updated_at=CURRENT_TIMESTAMP WHERE id=?",(sid,)); c.commit(); c.close()
                return self.send_json({"ok":True,"message":"تم إيقاف المورد من القوائم. الروابط التاريخية لم تُحذف."})
            except Exception as e:
                try: c.close()
                except: pass
                return self.send_json({"ok":False,"message":str(e)},500)
        if p=="/api/alias":
            try:
                payload=json.loads(body.decode('utf-8')); alias=str(payload.get('alias_name') or '').strip(); local_item_id=int(payload.get('local_item_id') or 0)
                if not alias or not local_item_id: return self.send_json({'ok':False,'message':'الصنف والاسم البديل مطلوبان.'},400)
                return self.send_json({'ok':True,'deferred':True,'message':'تم تسجيل الاسم البديل مؤقتًا داخل الفاتورة. سيتم حفظه عند حفظ الفاتورة.'})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        if p=="/api/approve":
            try:
                payload=json.loads(body.decode('utf-8')); line=payload.get('line',{}); item=(line.get('match') or {}).get('item')
                if not item: return self.send_json({'ok':False,'message':'لا يوجد صنف مرشح للاعتماد.'},400)
                return self.send_json({'ok':True,'deferred':True,'message':'تم اعتماد الاختيار داخل الفاتورة فقط. سيتم تثبيت التعلم عند حفظ الفاتورة.'})
            except Exception as e: return self.send_json({'ok':False,'message':str(e)},400)
        self.send_json({"error":"Not found"},404)


def run_app():
    print("Mutakamil Plus local app:", f"http://127.0.0.1:{PORT}")
    os.chdir(BASE)
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    threading.Timer(0.8, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}/")).start()
    server.serve_forever()


if __name__ == "__main__":
    run_app()
