# -*- coding: utf-8 -*-
"""calc.py — حسابات وقرارات المخزون والمبيعات (بدون Streamlit وبدون شيتات).

كل دالة هنا بتاخد أرقام/قوائم وترجّع نتيجة. نفس الكود اللي كان جوه stock_requests.py
بدون أي تعديل في السلوك. بيتستدعى من stock_requests.py، وبعدين من أي سيرفر/موقع.
"""
from datetime import datetime
from datetime import timedelta
import io
import re
import pandas as pd
import calendar as _ct_cal

# ══════════════════════════════════════════════
# قرار فجوة القناة (FBN / FBB / FBA / FSAB) — منطق واحد لكل القنوات
# ══════════════════════════════════════════════
def gap_cover(stock, avg):
    """يكفي كام يوم: المخزون ÷ متوسط البيع اليومي (9999 = مفيش مبيعات)."""
    return (stock / avg) if avg > 0 else 9999

def gap_need_qty(avg, stock, cov_days):
    """الكمية المطلوبة لتغطية cov_days يوم (تقريب لأعلى، ولا تقل عن صفر)."""
    return max(int(-(-(avg * cov_days - stock) // 1)), 0)

def gap_action(need, whq, avg, cov_days, sched="none"):
    """القرار: ok | pending | scheduled | send | low_wh | send_nowh.
    need: المخزون أقل من أيام التغطية. whq: المتوفر بالمستودع. sched: none|pending|scheduled|expired."""
    if not need:
        return "ok"
    if sched == "pending":
        return "pending"
    if sched == "scheduled":
        return "scheduled"
    if whq > 0 and whq >= avg * cov_days:
        return "send"
    if whq > 0:
        return "low_wh"
    return "send_nowh"

def _build_tacweed_bundle(raw):
    data_map, code_map, cost_map, code_cost_map = {}, {}, {}, {}
    if raw and len(raw) >= 2:
        header = [h.strip() for h in raw[0]]
        idx = {h: i for i, h in enumerate(header)}
        def col(row, *names):
            for name in names:
                i = idx.get(name)
                if i is not None and i < len(row):
                    return row[i]
            return ""
        for row in raw[1:]:
            sku_raw = col(row, "SKU", "Sku", "sku")
            if not str(sku_raw).strip():
                continue
            sku_up = str(sku_raw).strip().upper()
            code = str(col(row, "Code01", "code01", "Code 01")).strip()
            cost = _f2_or_none(col(row, "Cost", "cost", "التكلفة", "تكلفة"))
            data_map[sku_up] = {"code": code, "cost": cost}
            if code:
                code_map[sku_up] = code
            if cost is not None:
                cost_map[sku_up] = cost
            if code and cost is not None and code not in code_cost_map:
                code_cost_map[code] = cost
    return {"data": data_map, "code_map": code_map, "cost_map": cost_map, "code_cost_map": code_cost_map}

def _build_tacweed_amazon_bundle(raw):
    data_map, code_map, cost_map, code_cost_map = {}, {}, {}, {}
    if raw and len(raw) >= 2:
        header = [h.strip() for h in raw[0]]
        idx = {h: i for i, h in enumerate(header)}
        def col(row, *names):
            for name in names:
                i = idx.get(name)
                if i is not None and i < len(row):
                    return row[i]
            return ""
        for row in raw[1:]:
            asin_raw = col(row, "ASIN", "Asin", "asin")
            if not str(asin_raw).strip():
                continue
            asin_up = str(asin_raw).strip().upper()
            code = str(col(row, "Code01", "code01", "Code 01")).strip()
            cost = _f2_or_none(col(row, "Cost", "cost", "التكلفة", "تكلفة"))
            data_map[asin_up] = {"code": code, "cost": cost}
            if code:
                code_map[asin_up] = code
            if cost is not None:
                cost_map[asin_up] = cost
            if code and cost is not None and code not in code_cost_map:
                code_cost_map[code] = cost
    return {"data": data_map, "code_map": code_map, "cost_map": cost_map, "code_cost_map": code_cost_map}

def _f2(v, default=0.0):
    """تحويل آمن لأي قيمة نصية لرقم عشري | Safe float conversion.
    بيتعامل مع الفاصلة العشرية/فاصلة الآلاف العربية (٫ / ٬) زي اللي بتظهر في خلايا
    جوجل شيت المنسقة بالعربي | Also handles the Arabic decimal/thousands separators
    (٫ / ٬) that show up in Arabic-formatted Google Sheet cells."""
    try:
        s = str(v).strip().replace(",", "").replace("%", "")
        s = s.replace("\u066b", ".").replace("\u066c", "")
        return float(s) if s not in ("", "nan", "none") else default
    except Exception:
        return default

def _f2_or_none(v):
    """زي _f2 بس بترجع None لو الخلية فاضية/مش رقم بدل ما ترجع صفر — عشان صفر حقيقي (سعر=0)
    ميتلخبطش مع "مفيش سعر مسجل"."""
    s = str(v).strip().replace(",", "")
    s = s.replace("\u066b", ".").replace("\u066c", "")
    if s == "" or s.lower() in ("nan", "none"):
        return None
    try:
        return float(s)
    except Exception:
        return None

def get_base_price(sku_up, live_map, fallback_price=None):
    """يرجع (السعر, من_LIVE؟) — بيدّي الأولوية لعمود sale_price في تاب LIVE (سعر البيع الأساسي)،
    ولو مش موجود بيرجع السعر البديل (سعر العرض القديم من الطلبات) لو موجود."""
    live_info = live_map.get(sku_up)
    if live_info and live_info.get("price") is not None:
        return live_info["price"], True
    return fallback_price, False

def compute_net_price_after_fees(price, com_info):
    """يحسب صافي سعر البيع بعد خصم العمولة ومصاريف التوصيل، وبعدها بعد خصم 15% ضريبة.
    | Computes the net selling price after commission + delivery fees, then after 15% VAT.
    مصاريف التوصيل متوقع تكون رقم سالب في الشيت (زي -30) فبنضيفها مباشرة.
    Delivery is expected to be a negative number in the sheet (e.g. -30), so we add it directly.
    بيرجع (الصافي بعد العمولة والتوصيل, الصافي بعد خصم الضريبة كمان) | Returns
    (net after commission+delivery, net after also removing 15% VAT)."""
    if price is None or not com_info:
        return None, None
    commission_pct = com_info.get("commission_pct", 0.0)
    delivery = com_info.get("delivery", 0.0)
    commission_amount = price * commission_pct / 100.0
    net_after_fees = price - commission_amount + delivery
    net_after_tax = net_after_fees * 0.85  # خصم 15% ضريبة القيمة المضافة | remove 15% VAT
    return net_after_fees, net_after_tax

def get_latest_sku_price(r, ordered_dates):
    """يرجع آخر سعر بيع مسجل لهذا الـ SKU — بيدوّر بالترتيب (الأحدث أولاً) وياخد السعر
    اللي بيع بيه أكتر كمية في أقرب يوم فيه سعر. | Returns the most recent recorded selling
    price for this SKU — walks the dates most-recent-first and picks the price with the
    highest quantity on the nearest day that has a price."""
    for d in ordered_dates:
        day_prices_list = r["day_prices"].get(d, [])
        vals = []
        for item in day_prices_list:
            p, qty = item if isinstance(item, tuple) else (item, 1)
            if p and str(p).strip().lower() not in ("", "nan", "none"):
                try:
                    vals.append((float(str(p).replace(",", "")), qty))
                except Exception:
                    pass
        if vals:
            vals.sort(key=lambda x: -x[1])
            return vals[0][0]
    return None

def _build_warehouse_stock_map(raw):
    m = {}
    for row in raw[1:]:
        if len(row) >= 2 and row[0].strip():
            m[row[0].strip()] = row[1].strip()
    return m

def _china_norm_code(v):
    """توحيد شكل الباركود/الكود للمقارنة (بدون مسافات/شرطات/.0 وبدون أصفار في الأول للأرقام) |
    Normalises a barcode/code for matching."""
    s = str(v if v is not None else "").strip()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    s = re.sub(r"[\s\-]", "", s).upper()
    return s.lstrip("0") if s.isdigit() else s

def _china_eta_stats(eta_port, port_arr, cancelled, today):
    """نفس منطق أعمدة الإكسيل days_to_eta / delay_days لكن محسوبة لحظيًا (الإكسيل بيحسبهم وقت آخر
    حفظ فبيتأخروا): days_to_eta = الوصول المتوقع للميناء - النهاردة | delay_days = (وصول الميناء الفعلي أو النهاردة)
    - الوصول المتوقع للميناء لو موجب، وصفر للملغي | Same logic as the Excel columns, computed live."""
    e = _china_date(eta_port)
    if not e:
        return None, None
    days = (e - today).days
    if cancelled:
        return days, 0
    a = _china_date(port_arr)
    return days, max(0, ((a - e).days if a else (today - e).days))

def _china_int(v):
    try:
        return int(float(str(v).replace(",", "")))
    except Exception:
        return 0

def _china_date(s):
    try:
        return datetime.strptime(str(s)[:10], "%Y-%m-%d").date()
    except Exception:
        return None

def _china_merge(existing, new, keyfn):
    order, d = [], {}
    for r in existing:
        k = keyfn(r)
        if k not in d:
            order.append(k)
        d[k] = r
    for r in new:
        k = keyfn(r)
        if k not in d:
            order.append(k)
        d[k] = r
    return [d[k] for k in order]

def _china_rows_from_raw(raw):
    if not raw or len(raw) < 2:
        return []
    hdr = [str(h).strip() for h in raw[0]]
    out = []
    for row in raw[1:]:
        if not any(str(c).strip() for c in row):
            continue
        out.append({h: (str(row[i]).strip() if i < len(row) else "") for i, h in enumerate(hdr) if h})
    return out

def _cs_to_ship(code):
    return {"Allocated to Container": "Booked"}.get(code, code)

def _cs_to_cont(code):
    return {"Allocated to Container": "Booked"}.get(code, code)

def _cs_to_stock(code):
    return {"Delivered to Warehouse": "Warehouse Received"}.get(code, code)

def _cf_g(d, k):
    v = d.get(k, "") if d else ""
    return "" if v is None else str(v).strip()

def _cf_num(v):
    m = re.search(r"\d+", str(v or ""))
    return int(m.group()) if m else 0

def china_ship_record_key(s):
    # مفتاح الطلب = رقم الطلب | PO | قيمة الفاتورة (من غير رقم الشحنة ST — عشان تخصيص ST بعدين يعدّل نفس السجل مش يكرره)
    return "|".join([s.get("internal_order_no", ""), s.get("po_no", ""), s.get("invoice_value", "")])

def _cf_pick(rows, key, conflicts, label):
    """قيمة الحقل لمجموعة صفوف: أحدث last_updated_at ثم الأكتر تكرارًا ثم الأول. لو فيه قيم مختلفة بتتسجل كتعارض."""
    vals = [(_cf_g(r, key), _cf_g(r, "last_updated_at")) for r in rows if _cf_g(r, key)]
    if not vals:
        return ""
    distinct = {v for v, _ in vals}
    if len(distinct) > 1:
        conflicts.append({"السجل": label, "الحقل": key, "القيم": " | ".join(sorted(distinct))})
        latest = max(t for _, t in vals)
        cand = [v for v, t in vals if t == latest]
        best = max(set(cand), key=lambda v: (cand.count(v), -cand.index(v)))
        return best
    return vals[0][0]

def _china_ship_key(r):
    return china_ship_record_key(r)

def _china_line_key(r):
    return r.get("line_id") or r.get("record_key") or "|".join(
        [r.get("shipment_id", ""), r.get("container_no", ""), r.get("barcode", ""), r.get("sku", "")])

def _china_cont_key(r):
    return r.get("container_no") or r.get("record_key") or ""

def _china_avail(l):
    """(تاريخ الإتاحة, الأساس). السطر المستلم/الملغي: (None, ""). الوصول المتوقع لمخزن الرياض غير مكتوب: (None, ...)."""
    if l["_status"] != "incoming":
        return None, ""
    d = _china_date(l["_d"]["eta_dest"])
    return (d, "الوصول المتوقع لمخزن الرياض") if d else (None, "تاريخ الوصول لمخزن الرياض غير مكتوب")

def china_supply_projection(stock, rate, batches, today):
    """محاكاة بسيطة: المخزون الحالي بيتسحب بمعدل يومي ثابت، وكل دفعة بتتضاف في تاريخ إتاحتها.
    batches = [(تاريخ, كمية)]. بيرجع: gaps = [(من, إلى, أيام)] و lost_units (مبيعات معرضة للضياع) و covered_until."""
    if rate <= 0:
        return {"no_rate": True, "gaps": [], "gap_days": 0, "lost_units": 0, "covered_until": None}
    t, S, gaps, lost = 0.0, float(stock), [], 0.0
    for dte, qty in sorted(batches):
        d = max((dte - today).days, 0)
        before = S - rate * (d - t)
        if before < 0:
            zero_at = t + (S / rate if S > 0 else 0)
            gap = d - zero_at
            if gap >= 0.5:
                gaps.append((today + timedelta(days=round(zero_at)), dte, round(gap)))
                lost += rate * gap
            before = 0
        S, t = before + qty, d
    return {"no_rate": False, "gaps": gaps, "gap_days": sum(g[2] for g in gaps), "lost_units": round(lost),
            "covered_until": today + timedelta(days=round(t + S / rate))}

def china_decide(stock, rate, dated, undated_qty, today, cfg):
    """dated = [(تاريخ وصول مخزن الرياض, كمية)] مرتبة. بيرجع الحالة: gap / order_late / order_soon / need_date / safe / no_rate
    + التواريخ المهمة. دالة نقية بدون واجهة."""
    lead, safety, cover, near = (cfg["china_lead_time_days"], cfg["china_reorder_safety_days"],
                                 cfg["china_cover_days"], cfg["china_near_days"])
    if rate <= 0:
        return {"state": "no_rate", "sev": 2}
    dated = sorted(dated)
    days_left = stock / rate
    stockout = today + timedelta(days=int(days_left))
    proj = china_supply_projection(stock, rate, dated, today)
    covered = proj["covered_until"]
    order_by = covered - timedelta(days=lead)
    late_by = (today - order_by).days
    dec = {"stock": stock, "rate": rate, "days_left": days_left, "stockout": stockout, "covered": covered,
           "order_by": order_by, "late_by": late_by, "suggested": round(rate * cover), "undated": undated_qty,
           "next": dated[0] if dated else None, "lead": lead}
    if proj["gaps"]:
        g = proj["gaps"][0]
        dec.update(state="gap", sev=0, gap_from=g[0], gap_to=g[1], gap_days=proj["gap_days"], lost=proj["lost_units"],
                   should_order=g[0] - timedelta(days=lead))
    elif undated_qty and not dated:
        dec.update(state="need_date" if days_left <= lead + safety else "safe", sev=1 if days_left <= lead + safety else 3)
    elif late_by > 0:
        dec.update(state="order_late", sev=0)
    elif late_by >= -safety:
        dec.update(state="order_soon", sev=1)
    else:
        dec.update(state="safe", sev=3)
    return dec

def _china_days(date_str, today):
    d = _china_date(date_str)
    return (d - today).days if d else None

def _to_int(v):
    try:
        return int(float(str(v).replace(",","")))
    except:
        return 0

def parse_count_dates(cell_value):
    """
    يفك خلية بصيغة 'Nx | تاريخ1 | تاريخ2 | ...' ويرجع (العدد الحالي, باقي التواريخ كنص).
    لو الخلية فاضية أو بصيغة قديمة (تاريخ واحد بس)، يتعامل معاها بأمان.
    """
    val = (cell_value or "").strip()
    if not val:
        return 0, ""
    m = re.match(r"^(\d+)x\s*\|\s*(.*)$", val, re.DOTALL)
    if m:
        return int(m.group(1)), m.group(2).strip()
    # صيغة قديمة (تاريخ واحد فقط بدون عداد) — اعتبرها أول مرة
    return 1, val

def append_count_date(rest_dates, new_count, new_date):
    """يبني نص الخلية الجديد بصيغة 'Nx | تاريخ1 | تاريخ2 | ... | تاريخ جديد'."""
    rest_dates = (rest_dates or "").strip()
    if rest_dates:
        return f"{new_count}x | {rest_dates} | {new_date}"
    return f"{new_count}x | {new_date}"

def _find_fulfillment_col_idx(header):
    for ci, h in enumerate(header):
        if str(h).strip().lower() in ("fulfillment model", "fulfillment_model", "fulfillment"):
            return ci
    return None

def _find_status_col_idx(header):
    for ci, h in enumerate(header):
        if str(h).strip().lower() in ("status", "الحالة", "order status"):
            return ci
    return None

def _row_is_fbb(row, fulfillment_col_idx, status_col_idx):
    """يرجع True لو الصف ده طلب FBB — الشرط الوحيد: Fulfillment Model فيه FBP
    (Fulfilled by Partner). شرط الـ Status = Processing اتشال (كان بيمنع تصنيف
    صفوف قديمة كـ FBB لو حالتها اتغيّرت من Processing بعد كده) | Returns True
    when Fulfillment Model contains FBP (Fulfilled by Partner). The
    Status == Processing condition was removed (it was excluding rows whose
    status had since moved on from Processing)."""
    if fulfillment_col_idx is None or len(row) <= fulfillment_col_idx:
        return False
    fulfillment_val = str(row[fulfillment_col_idx]).strip().upper()
    return "FBP" in fulfillment_val or "PARTNER" in fulfillment_val

def _find_col_idx_amz(header, names):
    names_norm = {n.strip().lower() for n in names}
    for ci, h in enumerate(header):
        if str(h).strip().lower() in names_norm:
            return ci
    return None

def _amz_col_indexes(hdr):
    return {
        "sku":   _find_col_idx_amz(hdr, ["رقم تخزين سلعة التاجر msku", "msku", "sku"]),
        "ts":    _find_col_idx_amz(hdr, ["تاريخ شحنة العميل", "order timestamp"]),
        "cont":  _find_col_idx_amz(hdr, ["حاوية كاملة الحمولة", "fulfillment center"]),
        "qty":   _find_col_idx_amz(hdr, ["الكمية", "quantity", "qty"]),
        "price": _find_col_idx_amz(hdr, ["مبلغ المنتج", "price", "amount"]),
        "asin":  _find_col_idx_amz(hdr, ["asin"]),
    }

def _amz_row_is_fba(row, cont_idx):
    """True = تخزين FBA (أي قيمة غير FSAB في عمود حاوية كاملة الحمولة)،
    False = تخزين عادي (FSAB)."""
    if cont_idx is None or len(row) <= cont_idx:
        return False
    val = str(row[cont_idx]).strip().upper()
    return val != "FSAB" and val != ""

def _amz_row_matches(row, cont_idx, mode):
    """mode: 'all' (كل الصفوف) / 'normal' (FSAB بس) / 'fba' (أي حاجة غير FSAB)."""
    if mode == "all":
        return True
    is_fba = _amz_row_is_fba(row, cont_idx)
    return is_fba if mode == "fba" else not is_fba

def _ty_find_col(hdr, *names):
    low = [str(h).strip().lower() for h in hdr]
    for n in names:
        if n in low:
            return low.index(n)
    return None

def _ty_num(v):
    try:
        s = str(v).replace(",", "").strip()
        return float(s) if s not in ("", "-", "None", "nan") else 0.0
    except Exception:
        return 0.0

def _build_ty_products(raw):
    by_cid, img, name = {}, {}, {}
    for row in (raw or [])[1:]:
        if len(row) < 2 or not str(row[0]).strip() or not str(row[1]).strip():
            continue
        cid, bc = str(row[0]).strip(), str(row[1]).strip().upper()
        by_cid.setdefault(cid, [])
        if bc not in by_cid[cid]:
            by_cid[cid].append(bc)
        if len(row) > 2 and row[2] and bc not in name:
            name[bc] = row[2]
        if len(row) > 3 and row[3] and bc not in img:
            img[bc] = row[3]
    return {"by_cid": by_cid, "img": img, "name": name}

def build_warehouse_stockout_sku_lookup(rows):
    """يرجع dict: sku_up -> صف كود المستودع بتاعه، لاستخدام سريع (O(1)) جوه البادچات
    بدل ما نعيد حساب compute_warehouse_stockout_rows لكل SKU على حدة."""
    lookup = {}
    for r in rows:
        for sku_up in r["skus"]:
            lookup[sku_up] = r
    return lookup

def build_warehouse_stockout_asin_lookup(rows):
    """نفس build_warehouse_stockout_sku_lookup بالظبط بس بمفتاح ASIN — لبادچة
    warehouse_available_badge_amazon، والسحب المحسوب في كل صف أصلاً مجمّع من
    نون + أمازون مع بعض (شوف compute_warehouse_stockout_rows) | Same idea as
    build_warehouse_stockout_sku_lookup but keyed by ASIN — the draw in each
    row already combines Noon + Amazon (see compute_warehouse_stockout_rows)."""
    lookup = {}
    for r in rows:
        for asin_up in r["asins"]:
            lookup[asin_up] = r
    return lookup

def _avg_price_from(prices_for_key):
    tot, q = 0.0, 0
    for lst in (prices_for_key or {}).values():
        for it in lst:
            try:
                pr = float(str(it[0]).replace(",", "").strip())
                qt = int(it[1]) if it[1] else 1
            except Exception:
                continue
            if pr > 0:
                tot += pr * qt
                q += qt
    return (tot / q) if q else None

def _gap_qty(v):
    try:
        return float(str(v).replace(",", "").strip() or 0)
    except Exception:
        return 0.0

def _gap_sug(action, dest, whq, avg, cov_days, sched_txt=""):
    """بيرجع (سطور الاقتراح مرتبة تحت بعض). dest مثلا: 'تخزين نون (FBN)' أو 'تخزين أمازون (FBA)'."""
    if action == "send":
        return [f"🚚 يُرسل إلى {dest}", f"المتوفر بالمستودع {whq:,.0f} (يغطي {whq / avg:,.0f} يوم)"]
    if action == "low_wh":
        return [f"⛔ لا يُرسل إلى {dest}", f"كمية المستودع {whq:,.0f} قليلة (تكفي {whq / avg:,.0f} يوم فقط من أصل {cov_days})",
                "يلزم التوريد من المورد الأصلي"]
    if action == "send_nowh":
        return [f"⛔ لا يُرسل إلى {dest}", "غير متوفر بالمستودع", "يلزم التوريد من المورد الأصلي"]
    if action == "pending":
        return ["⏳ لا حاجة للإرسال الآن", "الجدولة قيد الموافقة — بانتظار الاعتماد"]
    if action == "scheduled":
        return ["📅 لا حاجة للإرسال الآن", "الصنف مجدول بالفعل"]
    return []

def parse_check_tokens(text):
    """يطلّع الأكواد من نص لاصق (حتى لو واتساب بتواريخ وأسماء) — بيشيل سطور [الوقت] الاسم: ويفضل الكود بس."""
    text = re.sub(r"\[[^\]\n]*\][^:\n]*:", " ", text or "")
    seen, out = set(), []
    for t in re.findall(r"[A-Za-z0-9][A-Za-z0-9\-]{5,}", text):
        u = t.upper()
        if u not in seen:
            seen.add(u); out.append(u)
    return out

def _ago_txt(n):
    return {0: "مجدولة النهارده", 1: "كانت مجدولة من يوم", 2: "كانت مجدولة من يومين"}.get(n, f"كانت مجدولة من {n} أيام")

def _ct_parse_date(v):
    """نص تاريخ -> date أو None. بيقبل yyyy-mm-dd و dd/mm/yyyy (لو الاتنين <=12 بنعتبره يوم/شهر) و mm/dd/yyyy لو اليوم > 12."""
    s = str(v or "").strip()
    if not s:
        return None
    s = s.replace("T", " ").split(" ")[0]
    for sep in ("-", "/", "."):
        p = s.split(sep)
        if len(p) == 3 and all(x.isdigit() for x in p):
            try:
                if len(p[0]) == 4:
                    y, m, d = int(p[0]), int(p[1]), int(p[2])
                else:
                    a, b, y = int(p[0]), int(p[1]), int(p[2])
                    y = y + 2000 if y < 100 else y
                    d, m = (a, b) if (a > 12 or b <= 12) else (b, a)
                return datetime(y, m, d).date()
            except Exception:
                return None
    return None


# ══════════════════════════════════════════════
# تحويل التواريخ (بيتكرر ملايين المرات على نفس القيم — بنحفظ النتيجة)
# ══════════════════════════════════════════════
def _parse_excel_date_raw(val):
    if val is None:
        return None
    try:
        if isinstance(val,(int,float)):
            return datetime(1899,12,30)+timedelta(days=int(val))
        s = str(val).strip().replace(" ","").replace(" ","")
        try:
            return datetime.strptime(s[:10],"%Y-%m-%d")
        except:
            pass
        try:
            return datetime.strptime(s[:10],"%d/%m/%Y")
        except:
            pass
        try:
            return datetime.strptime(s[:10],"%m/%d/%Y")
        except:
            pass
        return None
    except:
        return None

from functools import lru_cache as _lru_cache

@_lru_cache(maxsize=300000)
def _parse_excel_date_cached(val):
    return _parse_excel_date_raw(val)

def parse_excel_date(val):
    try:
        return _parse_excel_date_cached(val)
    except TypeError:
        return _parse_excel_date_raw(val)


# ══════════════════════════════════════════════
# ثوابت (أعمدة الشيتات، حالات الصين، ...)
# ══════════════════════════════════════════════
CONTRACT_COLS = ["اسم العقد | Contract Name", "تاريخ البداية | Start Date", "تاريخ النهاية | End Date", "ملاحظات | Notes",
                 "تم الاطلاع على | Acknowledged", "وقت الاطلاع | Acknowledged At", "تاريخ الإضافة | Added At"]
TABS_CONFIG = {
    "Requests":          ["SKU","Quantity","Image URL","Date Added","File Name"],
    "Approved":          ["SKU","Quantity Requested","Quantity Approved","Image URL","Date Added","Date Approved"],
    "Unavailable":       ["SKU","Quantity","Image URL","Date Added","Date Marked Unavailable"],
    "Ordered":           ["SKU","Quantity","Image URL","Date Added","Order Count","Notes"],
    # ملاحظة: تاب Scheduled الحقيقي فيه دلوقتي أعمدة زيادة (Store Name/Warehouse Name/Total QTY/
    # Section/Approval) من برنامج الجدولة المكتبي. الليستة هنا بتُستخدم بس للتأكد إن الأعمدة
    # الأساسية دي موجودة (get_or_create_worksheet بتضيف أي عمود ناقص من غيرها في آخر الشيت،
    # ومش بتغيّر ترتيب الأعمدة الموجودة). القراءة/الكتابة الفعلية بتتم حسب اسم العمود مش مكانه
    # (شوف get_scheduled_normalized / build_scheduled_row فوق).
    "Scheduled":         ["ASN","SKU","Quantity","Schedule Date","Image URL","Date Added","Notes","Flag"],
    "CancelledSchedule": ["ASN","SKU","Quantity","Schedule Date","Image URL","Date Added","Cancel Reason","Date Cancelled"],
    "Rescheduled":       ["ASN","SKU","Quantity","Old Schedule Date","Image URL","Date Added","Reschedule Reason","Date Moved"],
    "Expired":           ["ASN","SKU","Quantity","Schedule Date","Image URL","Date Added","Date Expired"],
    "Inventory":         ["SKU","Warehouse","Stock","Monthly Sales","Image URL","Date Uploaded"],
    "DailyOrders":       ["SKU","Order Timestamp","Status","Price","Quantity","Date Uploaded","Family"],
    "Settings":          ["Key","Value"],
    "Check":             ["ASN","SKU","Quantity","Schedule Date","Image URL","Date Added","Notes","Flag"],
    "CancelNotifications": ["ASN","SKUs","Schedule Date","Reason","Timestamp"],
    "Tacweed":           ["SKU","Code01","Cost","Date Uploaded"],
    # تاب تكويد أمازون — نفس فكرة Tacweed بالظبط لكن المفتاح ASIN (هوية منتج
    # أمازون الحقيقية) بدل SKU، عشان منتجات أمازون (FBA أو العادية) اللي
    # مالهاش نفس الـ SKU المستخدم في نون تقدر برضه تتربط بكود01 وتكلفة، وتشارك
    # نفس كمية WarehouseStock | Amazon mapping tab — same idea as Tacweed but
    # keyed by ASIN (Amazon's real product identity) instead of SKU, so Amazon
    # products (FBA or normal) that don't share Noon's SKU can still be linked
    # to a Code01 + cost and share the same WarehouseStock quantity.
    "TacweedAmazon":     ["ASIN","Code01","Cost","Date Uploaded"],
    "WarehouseStock":    ["Code","Quantity","Item Name","Date Uploaded"],
    # تاب "قيد الموافقة" بيتكتب فيه من برنامج الجدولة المكتبي (schedule_entry_app) لما تتحدد
    # علامة Approval — نفس أعمدة تاب Scheduled الجديدة. هنا بنقراه بس (حسب اسم العمود مش
    # مكانه) عشان نعرض ملحوظة في تابات ثانية إن السكو ده لسه قيد الموافقة.
    "PendingApproval":   ["Store Name","Warehouse Name","ASN","Total QTY","Section","Approval",
                           "SKU","Quantity","Schedule Date","Image URL","Date Added","Notes","Flag"],
    # تاب الإعلانات — بيتحدّث يدوي من جوجل شيت، وبنعرض أداء كل SKU منه في تاب المبيعات
    "Advertisements":    ["Campaign Name","Sku","Views","Clicks","Orders","ATC","Spends","Revenue",
                           "CTR","ROAS","CPC","CPS","CVR"],
    # تاب إعلانات أمازون — بيتحدّث يدوي من تصدير أمازون مباشرة (Sponsored Products
    # export)، فشكله مختلف تماماً عن تاب الإعلانات بتاع نون: مفيش عمود "Campaign
    # Name" ولا عمود ASIN مستقل — أول عمود ("المنتجات") ده كود المنتج/العرض
    # (زي B0C62FXJZF-E5-LJNS-ZTDG)، والـ ASIN الحقيقي هو أول 10 حروف/أرقام منه
    # بس (شوف get_ads_map_amazon تحت). الهيدرز هنا لازم تفضل مطابقة تماماً لهيدر
    # الشيت الحقيقي، لأن get_or_create_worksheet بيضيف أي عمود من الليستة دي مش
    # موجود فعلياً في أول صف بالشيت — لو الليستة غلط بتضيف أعمدة زيادة كل جلسة
    # وده كان سبب رئيسي للبطء | Amazon ads sheet — updated manually straight from
    # an Amazon Sponsored Products export, so its shape is completely different
    # from Noon's Advertisements tab: no "Campaign Name" column and no standalone
    # ASIN column — the first column ("المنتجات") is the product/offer code (e.g.
    # B0C62FXJZF-E5-LJNS-ZTDG), and the real ASIN is only its first 10
    # characters (see get_ads_map_amazon below). These headers must match the
    # real sheet's header row exactly, because get_or_create_worksheet appends
    # any header here that isn't already in row 1 — a wrong list here means
    # extra junk columns get appended every session, which was a real source of
    # slowness.
    "AdvertisementsAmz": ["المنتجات","الحالة","النوع","أهلية المنتجات المدعومة","المبيعات(SAR)",
                           "عائد الإنفاق على الإعلانات (ROAS)","معدل التحويل","مرات الظهور","نقرات",
                           "CTR","الإنفاق(SAR)","CPC(SAR)","الطلبات","ACOS",
                           "الطلبات الجديدة بالنسبة للماركة","النسبة المئوية للطلبات الجديدة بالنسبة للماركة",
                           "المبيعات الجديدة بالنسبة للماركة(SAR)","النسبة المئوية للمبيعات الجديدة بالنسبة للماركة",
                           "مرات الظهور القابلة للعرض"],
    # تاب العمولة ومصاريف التوصيل — بيتحدّث يدوي، وبنستخدمه لحساب صافي سعر البيع لكل SKU
    "COM":               ["SKU","مصاريف توصيل","العمولة"],
    # تاب LIVE — نسخة من ملف Noon catalog export، بيتحدّث يدوي بره البرنامج (مش من هنا).
    # sku_child هو المفتاح اللي بيتربط بيه مع باقي الشيتات (زي Inventory وDailyOrders).
    # عمود "price" ده سعر البيع الأساسي اللي بتُحسب عليه الخصومات
    # والعمولة والضريبة والإعلانات (بدل سعر الطلبات القديم، اللي بقى اسمه "سعر العرض").
    # عمود stock_xdock_net مخزون منفصل عن تاب Inventory (بنسميه في الواجهة "مخزون FBN" لما نقارنه)، بيتراقب لوحده في التنبيهات السريعة.
    "LIVE": ["psku_code","country_code","id_partner","partner_sku","family_code","brand_code",
             "partner_barcodes","sku_child","noon_title","noon_brand","active_price","msrp",
             "price","sale_price","noon_price_min","noon_price_max","seller_price_min",
             "seller_price_max","price_engine_min","price_engine_max","sale_start_date",
             "sale_end_date","is_active","warranty","stock_fbn_net","stock_xdock_gross",
             "stock_xdock_net","mp_code","noon_status"],
    # تاب مبيعات أمازون اليومية — بيتحدّث يدوي/من نظام أمازون، نفس فكرة DailyOrders
    # بالظبط بس لشحنات أمازون. الأعمدة هنا بترتيب مختلف عن DailyOrders (نون)،
    # فبنقراها دايمًا بالاسم مش بالمكان (شوف _amz_col_indexes تحت). عمود "حاوية
    # كاملة الحمولة": لو قيمته FSAB يبقى تخزين عادي (Easy Ship/شحن مباشر من
    # التاجر)، ولو أي قيمة تانية يبقى تخزين FBA (مخزون عند أمازون) | Daily Amazon
    # orders sheet — same idea as DailyOrders (Noon) but for Amazon shipments.
    # Column order here differs from DailyOrders, so columns are always looked
    # up by name, not position (see _amz_col_indexes below). "حاوية كاملة
    # الحمولة" column: FSAB = normal storage (merchant-shipped), anything else
    # = FBA storage.
    "DailyOrdersAmazon": ["تاريخ شحنة العميل", "رقم تخزين سلعة التاجر MSKU",
                          "رقم تخزين السلعة لشبكة نظام الشحن (FNSKU)", "ASIN",
                          "حاوية كاملة الحمولة", "الكمية", "رقم الطلب من أمازون",
                          "العملة", "مبلغ المنتج"],
    # مخزون أمازون — تصدير من Amazon Seller Central (Manage Inventory)، مصدر
    # مخزون أمازون المستقل تمامًا عن مخزون نون (Inventory). تاب مبيعات امازون
    # بيبني قايمة منتجاته من الشيت ده (مش من Inventory بتاع نون) عشان تبقى
    # قايمة منتجات أمازون الحقيقية | Amazon inventory export from Seller
    # Central — the independent stock source for the Amazon Sales tab (not
    # Noon's Inventory sheet), so the product list reflects Amazon's actual
    # catalog.
    "StockAmazon": ["seller-sku", "fulfillment-channel-sku", "asin",
                    "condition-type", "Warehouse-Condition-code", "Quantity Available"],
    # مبيعات تريندول — بتتعمل تلقائي أول مرة. Barcode = SKU (نفس تكويد نون). مفيش بيانات عملاء.
    "DailyOrdersTrendyol": ["Row Key", "Barcode", "Order Date", "Order Number", "Order Status", "Quantity",
                            "Net Unit Price", "Invoice Amount", "Commission Rate", "Image URL", "Product Name", "Date Uploaded"],
    # إعلانات تريندول (تقرير Product Ads) — بتتعمل تلقائي أول مرة
    "AdvertisementsTrendyol": ["Ad Name", "Ad Status", "Start Date", "Products", "Content IDs", "Daily Budget",
                               "Cost", "Clicks", "Views", "Orders", "Sales", "ROAS", "Date Uploaded"],
    # ملف منتجات تريندول (لربط Content ID بالـ Barcode) — بيتعمل تلقائي أول مرة
    "TrendyolProducts": ["Content ID", "Barcode", "Product Name", "Image URL", "Date Uploaded"],
    # صور منتجات الصين (باركود -> رابط صورة على Google Drive) — بتترفع من تاب «رفع ملف» > عام
    "ChinaImages":      ["Barcode", "Image URL", "Drive File ID", "Date Uploaded"],
    "Contracts":        CONTRACT_COLS,
}
OLD_SCHED_COLS = ["ASN", "SKU", "Quantity", "Schedule Date", "Image URL", "Date Added", "Notes", "Flag"]
CHINA_SHIP_COLS = ["shipment_id","internal_order_no","po_no","supplier_name","buyer_company","shipment_mode",
                   "incoterm","origin_location","port_of_loading","destination_port","final_destination",
                   "shipping_line","bl_no","booking_no","container_count","currency","invoice_value",
                   "freight_value","planned_etd","actual_etd","eta_port","eta_destination","shipment_status",
                   "tracking_url","last_updated_at","updated_by","notes","row_state","invoice_no","record_key"]
CHINA_LINE_COLS = ["shipment_id","line_id","container_no","sku","model_no","barcode","product_name_en",
                   "product_name_ar","variant","cartons","units_per_carton","total_qty","unit_cost","currency",
                   "line_amount","cbm","gross_weight_kg","net_weight_kg","stock_stage","production_ready_date",
                   "actual_load_date","noon_sku","amazon_sku","trendyol_sku","image_url","notes","last_updated_at",
                   "updated_by","record_key","row_state","internal_order_no","po_no"]
CHINA_CONT_COLS = ["shipment_id","container_no","seal_no","container_type","shipping_line","vessel_name","voyage_no",
                   "bl_no","port_of_loading","destination_port","actual_etd","eta_port","port_arrival_actual",
                   "warehouse_arrival_actual","container_status","last_known_location","tracking_url",
                   "gross_weight_kg","total_cbm","total_cartons","total_units","customs_status","delay_reason",
                   "last_updated_at","updated_by","record_key","row_state","days_to_eta","delay_days","final_destination"]
CHINA_DATE_FIELDS = {"planned_etd","actual_etd","eta_port","eta_destination","last_updated_at","production_ready_date",
                     "actual_load_date","port_arrival_actual","warehouse_arrival_actual"}
CHINA_SKIP_COLS = {"days_to_eta", "delay_days"}
CHINA_STAGES = ["Factory Confirmed", "In Production", "Production Completed", "Ready to Ship", "Allocated to Container",
                "Loaded", "Shipped / In Transit", "Arrived at Port", "Customs Clearance", "Cleared", "Warehouse Received"]
CHINA_STAGE_AR = {"Factory Confirmed": "تأكيد المصنع", "In Production": "في الإنتاج", "Production Completed": "اكتمل الإنتاج",
                  "Ready to Ship": "جاهز للشحن", "Allocated to Container": "تخصيص حاوية", "Loaded": "تم التحميل",
                  "Shipped / In Transit": "في البحر", "Arrived at Port": "وصل الميناء", "Customs Clearance": "تخليص جمركي",
                  "Cleared": "تم التخليص", "Warehouse Received": "استُلم بالمستودع"}
CHINA_SHIPPED_IDX = 5     # من "تم التحميل" فما فوق = خرجت من المصنع | Loaded and beyond = left the factory
CHINA_RECEIVED_IDX = 10
CHINA_CONT_STATUS_IDX = {"Loaded": 5, "Shipped / In Transit": 6, "Arrived at Port": 7, "Customs Clearance": 8,
                         "Cleared": 9, "Delivered to Warehouse": 10, "Warehouse Received": 10}
CHINA_STATUS = [
 ("Factory Confirmed", "تأكيد المصنع | Factory confirmed"), ("In Production", "قيد الإنتاج | In production"),
 ("Production Completed", "اكتمل الإنتاج | Production completed"), ("Ready to Ship", "جاهز للشحن | Ready to ship"),
 ("Allocated to Container", "حاوية متجهزة | Container prepared"), ("Loaded", "تم التحميل | Loaded"),
 ("Shipped / In Transit", "في البحر | In transit"), ("Arrived at Port", "وصلت الميناء | At port"),
 ("Customs Clearance", "تحت التخليص | In customs"), ("Cleared", "تم التخليص | Customs cleared"),
 ("Delivered to Warehouse", "وصلت المستودع | Delivered to warehouse"), ("On Hold", "موقوف | On hold"), ("Cancelled", "ملغي | Cancelled"),
]
CHINA_STATUS_LEGACY = ["shipment_status", "container_status", "stock_stage"]    # أعمدة قديمة بتتقري من الملفات/الشيت القديم بس
_CS_ALIAS = {"order confirmed": "Factory Confirmed", "warehouse received": "Delivered to Warehouse", "booked": "Allocated to Container"}
_CS_RANK = {code: i for i, (code, _) in enumerate(CHINA_STATUS)}
_CS_LABEL = dict(CHINA_STATUS)
_CS_LOOKUP = {}
CHINA_FLAT_COLS = [
 ("ORDER","internal_order_no","رقم الطلب الداخلي NO.xxx",14,"t"),("ORDER","po_no","رقم طلب المصنع PO",16,"t"),
 ("KEY","barcode","الباركود",16,"t"),("KEY","product_image","صورة المنتج (تلقائي من الباركود)",13,"g"),
 ("KEY","status","الحالة | Status",24,"g"),
 ("AUTO","next_action","المطلوب دلوقتي | Next action (تلقائي)",42,"g"),
 ("ORDER","supplier_name","اسم المورد",26,"g"),("ORDER","buyer_company","شركة المشتري",22,"g"),
 ("ORDER","shipment_id","رقم الشحنة STxxxxx",13,"t"),("ORDER","shipment_mode","طريقة الشحن",11,"g"),
 ("ORDER","incoterm","Incoterm",9,"g"),("ORDER","origin_location","منشأ البضاعة",14,"g"),
 ("ORDER","invoice_no","رقم الفاتورة",14,"t"),("ORDER","invoice_value","قيمة الفاتورة",13,"n"),
 ("ORDER","freight_value","قيمة الشحن",12,"n"),("ORDER","currency","العملة",8,"g"),
 ("ORDER","planned_etd","ETD المخطط",12,"d"),("ORDER","eta_destination","ETA الوجهة النهائية",14,"d"),
 ("ORDER","booking_no","رقم الحجز",14,"t"),
 ("ORDER","container_count","عدد الحاويات",9,"n"),("ORDER","final_destination","الوجهة النهائية",11,"g"),
 ("ORDER","order_notes","ملاحظات الطلب",24,"g"),
 ("CONTAINER","container_no","رقم الحاوية",15,"t"),("CONTAINER","seal_no","رقم الختم",13,"t"),
 ("CONTAINER","container_type","نوع الحاوية",10,"g"),("CONTAINER","shipping_line","شركة الشحن",13,"g"),
 ("CONTAINER","vessel_name","اسم السفينة",18,"g"),("CONTAINER","voyage_no","رقم الرحلة",12,"t"),
 ("CONTAINER","bl_no","رقم B/L",18,"t"),("CONTAINER","port_of_loading","ميناء التحميل",13,"g"),
 ("CONTAINER","destination_port","ميناء الوصول",13,"g"),("CONTAINER","actual_etd","ETD الفعلي",12,"d"),
 ("CONTAINER","eta_port","ETA الميناء",12,"d"),("CONTAINER","port_arrival_actual","وصول فعلي للميناء",13,"d"),
 ("CONTAINER","warehouse_arrival_actual","وصول فعلي للمستودع",13,"d"),
 ("CONTAINER","customs_status","حالة التخليص",15,"g"),("CONTAINER","last_known_location","آخر موقع معروف",16,"g"),
 ("CONTAINER","delay_reason","سبب التأخير",18,"g"),("CONTAINER","tracking_url","رابط التتبع",24,"g"),
 ("CONTAINER","ctr_gross_weight_kg","وزن الحاوية كغ",11,"n"),("CONTAINER","ctr_total_cbm","إجمالي CBM الحاوية",11,"n"),
 ("CONTAINER","ctr_total_cartons","إجمالي كراتين الحاوية",11,"n"),("CONTAINER","ctr_total_units","إجمالي قطع الحاوية",11,"n"),
 ("LINE","line_id","معرّف السطر (تلقائي)",13,"t"),("LINE","sku","SKU الداخلي",14,"t"),("LINE","model_no","الموديل",14,"t"),
 ("LINE","product_name_en","اسم المنتج EN",30,"g"),("LINE","product_name_ar","اسم المنتج AR",24,"g"),
 ("LINE","variant","المتغير",10,"g"),("LINE","cartons","الكراتين",9,"n"),("LINE","units_per_carton","قطع/كرتون",9,"n"),
 ("LINE","total_qty","إجمالي الكمية",10,"n"),("LINE","unit_cost","تكلفة الوحدة",10,"n"),("LINE","cbm","CBM",8,"n"),
 ("LINE","gross_weight_kg","وزن إجمالي",9,"n"),("LINE","net_weight_kg","وزن صافي",9,"n"),
 ("LINE","production_ready_date","جاهزية الإنتاج",12,"d"),
 ("LINE","actual_load_date","تاريخ التحميل",12,"d"),("LINE","noon_sku","SKU نون",14,"t"),("LINE","amazon_sku","SKU أمازون",14,"t"),
 ("LINE","trendyol_sku","SKU تريندول",14,"t"),("LINE","image_url","رابط الصورة",20,"g"),("LINE","notes","ملاحظات المنتج",22,"g"),
 ("STAMP","last_updated_at","آخر تحديث",12,"d"),("STAMP","updated_by","حدّث بواسطة",11,"g"),
 # أعمدة محسوبة بمعادلة — الموظف مبيكتبش فيها
 ("AUTO","line_amount","قيمة السطر (محسوب)",12,"g"),("AUTO","days_to_eta","أيام حتى ETA (محسوب)",11,"g"),
 ("AUTO","delay_days","أيام التأخير (محسوب)",11,"g"),("AUTO","row_state","اكتمال الصف (محسوب)",12,"g"),
]
CHINA_FLAT_KEYS = [c[1] for c in CHINA_FLAT_COLS]
CHINA_FLAT_FORMULA_KEYS = {c[1] for c in CHINA_FLAT_COLS if c[0] == "AUTO"} | {"product_image"}   # أعمدة معادلات — الموظف مبيكتبش فيها
CHINA_FLAT_INPUT_KEYS = [c[1] for c in CHINA_FLAT_COLS if c[1] not in CHINA_FLAT_FORMULA_KEYS]
_CF_SHIP_ONLY = ["internal_order_no","po_no","supplier_name","buyer_company","shipment_mode","incoterm","origin_location","invoice_no",
                 "invoice_value","freight_value","currency","planned_etd","eta_destination","booking_no",
                 "container_count","final_destination"]
_CF_SHIP_DATA = ["supplier_name","buyer_company","shipment_mode","incoterm","origin_location","invoice_no","invoice_value",
                 "freight_value","planned_etd","eta_destination","booking_no","container_count","order_notes"]
_CF_CONT_ONLY = ["container_no","seal_no","container_type","vessel_name","voyage_no","port_arrival_actual","warehouse_arrival_actual",
                 "customs_status","last_known_location","delay_reason"]
_CF_SHARED = ["shipping_line","bl_no","port_of_loading","destination_port","actual_etd","eta_port","tracking_url"]   # الحاوية أولاً ثم الطلب
_CF_CONT_RENAME = {"gross_weight_kg": "ctr_gross_weight_kg", "total_cbm": "ctr_total_cbm",
                   "total_cartons": "ctr_total_cartons", "total_units": "ctr_total_units"}
_CF_LINE = ["line_id","sku","model_no","barcode","product_name_en","product_name_ar","variant","cartons","units_per_carton",
            "total_qty","unit_cost","cbm","gross_weight_kg","net_weight_kg","production_ready_date",
            "actual_load_date","noon_sku","amazon_sku","trendyol_sku","image_url","notes"]
_CHINA_LABELS = {c[1]: c[2] for c in CHINA_FLAT_COLS}
_AR_MONTHS = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
TY_EXCLUDED_STATUSES = {"returned", "undelivered", "cancelled", "canceled", "waiting for payment"}
TY_ADS_HEADERS = TABS_CONFIG["AdvertisementsTrendyol"]
FAMILY_HOME_GROUP = {"kitchen_dining", "home_improvement", "gardening", "home_decor", "furniture", "bedding"}
FAMILY_LABELS = {
    # لو عايز تضيف أقسام تانية معروفة، ضيفها هنا: "raw_family_value": "🔖 الاسم بالعربي | Name in English"
    "electronics": "⚡ إلكترونيات | Electronics",
    "fashion": "👗 موضة | Fashion",
    "beauty": "💄 جمال | Beauty",
    "toys": "🧸 ألعاب | Toys",
    "baby": "👶 مستلزمات أطفال | Baby",
    "sports": "🏀 رياضة | Sports",
    "grocery": "🛒 بقالة | Grocery",
    "automotive": "🚗 سيارات | Automotive",
    "books": "📚 كتب | Books",
    "pet_supplies": "🐾 مستلزمات حيوانات | Pet Supplies",
    "office": "🖊️ مستلزمات مكتبية | Office",
    "health": "💊 صحة | Health",
}
_XS_NAMES = {"noon": "نون", "amz": "أمازون", "ty": "تريندول"}
CT_ORD_AR = ["الأول", "الثاني", "الثالث", "الرابع", "الخامس", "السادس", "السابع", "الثامن", "التاسع", "العاشر", "الحادي عشر", "الثاني عشر"]
CT_WINDOW_DAYS = 2        # مدة ظهور الإشعار (اليوم نفسه + اليوم اللي بعده)
for _code, _lab in CHINA_STATUS:
    for _k in {_code, _lab, _lab.split("|")[0], _lab.split("|")[-1]}:
        _CS_LOOKUP[_k.strip().lower()] = _code
_CS_LOOKUP.update(_CS_ALIAS)


# ══════════════════════════════════════════════
# حسابات تعتمد على بيانات: env هو بوابة البيانات (في التطبيق = stock_requests.py)
# ══════════════════════════════════════════════

def get_scheduled_normalized(env, force=False):
    """يرجع نفس شكل get_all_values() (هيدر + صفوف) لكن بترتيب أعمدة ثابت (OLD_SCHED_COLS)
    بغض النظر عن ترتيب الأعمدة الحقيقي في تاب Scheduled دلوقتي — عشان باقي الكود اللي
    بيفترض row[0]=ASN, row[1]=SKU, row[2]=Quantity, row[3]=Schedule Date ...الخ يفضل شغال صح."""
    raw = env.get_cached(env.scheduled_sheet, force=force)
    if not raw:
        return raw
    real_header = raw[0]
    idx_map = {name: real_header.index(name) for name in OLD_SCHED_COLS if name in real_header}
    out_rows = [OLD_SCHED_COLS]
    for row in raw[1:]:
        if len(row) < len(real_header):
            row = row + [""] * (len(real_header) - len(row))
        out_rows.append([
            row[idx_map[c]] if c in idx_map and idx_map[c] < len(row) else ""
            for c in OLD_SCHED_COLS
        ])
    return out_rows

def build_scheduled_row(env, asn, sku, qty, sdate, img, date_added, notes="", flag=""):
    """يبني صف جاهز للإضافة في تاب Scheduled حسب ترتيب أعمدة الشيت الحقيقي دلوقتي
    (بما فيها الأعمدة الجديدة Store Name/Warehouse Name/Total QTY/Section/Approval)،
    وبيسيب الأعمدة الجديدة دي فاضية لأن stock_requests.py مش هو المصدر بتاعها —
    برنامج الجدولة المكتبي هو اللي بيملاها."""
    raw = env.get_cached(env.scheduled_sheet, force=False)
    real_header = raw[0] if raw else OLD_SCHED_COLS
    values = {
        "ASN": asn, "SKU": sku, "Quantity": qty, "Schedule Date": sdate,
        "Image URL": img, "Date Added": date_added, "Notes": notes, "Flag": flag,
    }
    return [values.get(h, "") for h in real_header]

def _tab_lock_setting_key(tab_key):
    return f"tab_lock_{tab_key}"

def get_tacweed_data(env):
    return env._get_tacweed_bundle()["data"]

def get_tacweed_map(env):
    return env._get_tacweed_bundle()["code_map"]

def get_tacweed_reverse_map(env):
    """يرجع dict: كود -> قائمة SKUs المرتبطة بيه (عكس get_tacweed_map) — لأن نفس
    الكود ممكن يتكرر لأكتر من SKU وبالتالي يشتركوا في نفس كمية المستودع الواحدة."""
    rev = {}
    for sku_up, code in get_tacweed_map(env).items():
        rev.setdefault(code, []).append(sku_up)
    return rev

def get_tacweed_cost_map(env):
    """SKU (upper) -> تكلفة مسجلة مباشرة على نفس صف الـ SKU (لو موجودة) | SKU -> cost
    recorded directly on that SKU's own row (if present)."""
    return env._get_tacweed_bundle()["cost_map"]

def get_tacweed_code_cost_map(env):
    """كود01 -> تكلفة (بتاخد أول تكلفة مسجلة لأي SKU مرتبط بالكود ده) | Code01 -> cost
    (first recorded cost among the SKUs sharing that code)."""
    return env._get_tacweed_bundle()["code_cost_map"]

def get_sku_cost(env, sku_up):
    """يرجع تكلفة الوحدة لهذا الـ SKU: تكلفة مسجلة على نفس الـ SKU في تاب التكويد أولاً،
    ولو مش موجودة بيرجع لتكلفة الكود01 المرتبط بيه (لو أي SKU تاني بنفس الكود ليه تكلفة
    مسجلة) — وإلا بيرجع None | Unit cost for this SKU: a cost recorded directly against
    the SKU first, falling back to the cost recorded for its shared Code01, else None."""
    bundle = env._get_tacweed_bundle()
    info = bundle["data"].get(sku_up)
    if not info:
        return None
    if info.get("cost") is not None:
        return info["cost"]
    code = info.get("code")
    if code:
        return bundle["code_cost_map"].get(code)
    return None

def get_tacweed_amazon_data(env):
    return env._get_tacweed_amazon_bundle()["data"]

def get_tacweed_amazon_map(env):
    """ASIN (upper) -> الكود 01 | ASIN -> Code01."""
    return env._get_tacweed_amazon_bundle()["code_map"]

def get_tacweed_amazon_reverse_map(env):
    """كود -> قائمة ASINs المرتبطة بيه (عكس get_tacweed_amazon_map) — نفس فكرة
    get_tacweed_reverse_map بالظبط بس لأمازون، عشان كود واحد ممكن يشترك بين
    ASIN أمازون وSKU نون في نفس كمية المستودع."""
    rev = {}
    for asin_up, code in get_tacweed_amazon_map(env).items():
        rev.setdefault(code, []).append(asin_up)
    return rev

def get_tacweed_amazon_cost_map(env):
    return env._get_tacweed_amazon_bundle()["cost_map"]

def get_tacweed_amazon_code_cost_map(env):
    return env._get_tacweed_amazon_bundle()["code_cost_map"]

def get_asin_cost(env, asin_up):
    """يرجع تكلفة الوحدة لهذا الـ ASIN — نفس منطق get_sku_cost بالظبط بس من
    تاب TacweedAmazon: تكلفة مسجلة على نفس الـ ASIN أولاً، وإلا تكلفة الكود01
    المشترك (سواء مسجلة على ASIN تاني أو حتى SKU نون بنفس الكود)."""
    bundle = env._get_tacweed_amazon_bundle()
    info = bundle["data"].get(asin_up)
    code = info.get("code") if info else None
    cost = info.get("cost") if info else None
    if cost is not None:
        return cost
    if code:
        c = bundle["code_cost_map"].get(code)
        if c is not None:
            return c
        # fallback: لو نفس الكود مسجل بتكلفة في تاب نون التقليدي (Tacweed)
        return get_tacweed_code_cost_map(env).get(code)
    return None

def resolve_unit_cost(env, sku_up, asin_up=None):
    """تكلفة الوحدة الصحيحة سواء المنتج نون أو أمازون: لو معانا ASIN، بيدور
    الأول في تاب TacweedAmazon (get_asin_cost)، وإلا يرجع لتاب Tacweed العادي
    بالـ SKU — عشان أي مكان بيحسب صافي الربح (تاب المبيعات أو الداشبورد)
    يستخدم نفس المصدر الصح بدل ما يفضل بيقرا SKU بس حتى لمنتجات أمازون |
    Correct unit cost whether the product is Noon or Amazon: if an ASIN is
    given, checks TacweedAmazon (via get_asin_cost) first, otherwise falls
    back to the regular Tacweed sheet by SKU."""
    if asin_up:
        c = get_asin_cost(env, asin_up)
        if c is not None:
            return c
    return get_sku_cost(env, sku_up)

def get_ads_map(env):
    data = env.get_cached(env.ads_sheet)
    if not data or len(data) < 2:
        return {}
    header = [h.strip() for h in data[0]]
    idx = {h: i for i, h in enumerate(header)}
    def col(row, *names):
        for name in names:
            i = idx.get(name)
            if i is not None and i < len(row):
                return row[i]
        return ""
    m = {}
    for row in data[1:]:
        sku_raw = col(row, "Sku", "SKU", "sku")
        if not str(sku_raw).strip():
            continue
        sku_up = str(sku_raw).strip().upper()
        entry = {
            "campaign": col(row, "Campaign Name") or "—",
            "views":  _f2(col(row, "Views")),
            "clicks": _f2(col(row, "Clicks")),
            "orders": _f2(col(row, "Orders")),
            "atc":    _f2(col(row, "ATC")),
            "spends": _f2(col(row, "Spends")),
            "revenue":_f2(col(row, "Revenue")),
            "ctr":    _f2(col(row, "CTR")),
            "roas":   _f2(col(row, "ROAS")),
            "cpc":    _f2(col(row, "CPC")),
            "cps":    _f2(col(row, "CPS")),
            "cvr":    _f2(col(row, "CVR")),
        }
        m.setdefault(sku_up, []).append(entry)
    return m

def get_ads_map_amazon(env):
    """بتقرأ من شيت AdvertisementsAmz بشكله الجديد — تصدير إعلانات أمازون
    مباشرة، مفهوش عمود "Campaign Name" ولا عمود ASIN مستقل. عمود "المنتجات"
    فيه كود المنتج/العرض (زي B0C62FXJZF-E5-LJNS-ZTDG)، والـ ASIN الحقيقي هو
    أول 10 حروف/أرقام منه بس، فممكن أكتر من صف (أكتر من عرض/تخزين) يترجعوا
    لنفس الـ ASIN وده طبيعي وبيتجمّعوا مع بعض زي ما كان بيحصل مع الحملات قبل
    كده. مفيش عمود "Campaign Name" في الشكل الجديد، فبنستخدم قيمة عمود
    "المنتجات" الكاملة (كود المنتج/العرض) كبديل لعرض/تجميع كل صف على حدة في
    تاب "تحليل أداء الإعلانات". ومفيش عمود ATC (إضافة للسلة) في الشكل الجديد،
    فبتفضل صفر دايماً. كل الأعمدة بتتقرا بالاسم مش بالمكان عشان تفضل شغالة لو
    ترتيب الأعمدة اتغيّر | Reads the new AdvertisementsAmz layout — a direct
    Amazon ads export with no "Campaign Name" column and no standalone ASIN
    column. The "المنتجات" column holds the product/offer code (e.g.
    B0C62FXJZF-E5-LJNS-ZTDG); the real ASIN is only its first 10 characters,
    so more than one row (more than one offer/storage type) can map to the
    same ASIN — that's expected and they're summed together just like
    multiple campaigns used to be. Since there's no "Campaign Name" column
    anymore, the full "المنتجات" value is used as a stand-in so each row can
    still be shown/grouped individually in the Ads Performance Analysis tab.
    There's no ATC (add-to-cart) column in the new layout, so it's always
    zero. Every column is read by name, not position, so this keeps working
    if the column order changes."""
    data = env.get_cached(env.ads_amz_sheet)
    if not data or len(data) < 2:
        return {}
    header = [h.strip() for h in data[0]]
    idx = {h: i for i, h in enumerate(header)}
    def col(row, *names):
        for name in names:
            i = idx.get(name)
            if i is not None and i < len(row):
                return row[i]
        return ""
    m = {}
    for row in data[1:]:
        product_raw = str(col(row, "المنتجات", "ASIN", "Asin", "asin")).strip()
        if not product_raw:
            continue
        asin_up = product_raw[:10].strip().upper()
        if not asin_up:
            continue
        entry = {
            "campaign": product_raw,
            "views":  _f2(col(row, "مرات الظهور", "Views")),
            "clicks": _f2(col(row, "نقرات", "Clicks")),
            "orders": _f2(col(row, "الطلبات", "Orders")),
            "atc":    0.0,
            "spends": _f2(col(row, "الإنفاق(SAR)", "Spends")),
            "revenue":_f2(col(row, "المبيعات(SAR)", "Revenue")),
            "ctr":    _f2(col(row, "CTR")),
            "roas":   _f2(col(row, "عائد الإنفاق على الإعلانات (ROAS)", "ROAS")),
            "cpc":    _f2(col(row, "CPC(SAR)", "CPC")),
            "cps":    0.0,
            "cvr":    _f2(col(row, "معدل التحويل", "CVR")),
        }
        m.setdefault(asin_up, []).append(entry)
    return m

def get_com_map(env):
    data = env.get_cached(env.com_sheet)
    if not data or len(data) < 2:
        return {}
    header = [h.strip() for h in data[0]]
    idx = {h: i for i, h in enumerate(header)}
    def col(row, *names):
        for name in names:
            i = idx.get(name)
            if i is not None and i < len(row):
                return row[i]
        return ""
    m = {}
    for row in data[1:]:
        sku_raw = col(row, "SKU", "Sku", "sku")
        if not str(sku_raw).strip():
            continue
        sku_up = str(sku_raw).strip().upper()
        delivery = _f2(col(row, "مصاريف توصيل"))
        commission_pct = _f2(col(row, "العمولة"))
        m[sku_up] = {"delivery": delivery, "commission_pct": commission_pct}
    return m

def get_live_map(env):
    """يرجع {SKU (upper) -> {price, sale_price, stock_xdock_net, noon_title, noon_status}}
    من تاب LIVE (نسخة ملف Noon catalog export). المفتاح هو sku_child.
    ملحوظة: مفتاح "price" (اللي بتعتمد عليه كل حسابات المبيعات والداشبورد) بقى مصدره
    عمود sale_price في تاب LIVE بدل عمود price | Note: the "price" key (which all
    Sales/Dashboard calculations rely on) is now sourced from the sale_price column
    in the LIVE sheet instead of the price column."""
    data = env.get_cached(env.live_sheet)
    if not data or len(data) < 2:
        return {}
    header = [h.strip() for h in data[0]]
    idx = {h: i for i, h in enumerate(header)}
    def col(row, *names):
        for name in names:
            i = idx.get(name)
            if i is not None and i < len(row):
                return row[i]
        return ""
    m = {}
    for row in data[1:]:
        sku_raw = col(row, "sku_child", "partner_sku", "SKU", "sku")
        if not str(sku_raw).strip():
            continue
        sku_up = str(sku_raw).strip().upper()
        m[sku_up] = {
            "price": _f2_or_none(col(row, "sale_price")),
            "sale_price": _f2_or_none(col(row, "sale_price")),
            "stock_xdock_net": _to_int(col(row, "stock_xdock_net")),
            "noon_title": col(row, "noon_title"),
            "noon_status": col(row, "noon_status"),
        }
    return m

def _china_clean(v, field=""):
    if v is None or isinstance(v, bool):
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)):
        if field in CHINA_DATE_FIELDS and 20000 < v < 80000:
            return (datetime(1899, 12, 30) + timedelta(days=int(v))).strftime("%Y-%m-%d")
        if float(v) == int(v):
            return str(int(v))
        return f"{v:.4f}".rstrip("0").rstrip(".")
    s = str(v).strip()
    if s.startswith("=") or (s.startswith("#") and s.endswith(("!", "?", "A"))):
        return ""
    return s

def china_status_code(v):
    """نص الحالة (عربي | English أو الكود القديم) -> الكود الداخلي، أو '' لو مش معروف."""
    return _CS_LOOKUP.get(str(v or "").strip().lower(), "")

def china_status_label(code):
    return _CS_LABEL.get(code, "")

def _cs_prog(code):
    r = _CS_RANK.get(code, -1)
    return r if 0 <= r <= _CS_RANK["Delivered to Warehouse"] else -1      # درجة التقدم الفعلية (من غير موقوف/ملغي)

def china_row_status(r):
    """الحالة (كود) للصف: من عمود الحالة، وإلا من الأعمدة القديمة (حاوية ثم مرحلة المخزون ثم الطلب)."""
    code = china_status_code(r.get("status", ""))
    if code:
        return code
    cands = [china_status_code(r.get(k, "")) for k in ("container_status", "stock_stage", "shipment_status")]
    cands = [x for x in cands if x]
    if not cands:
        return ""
    prog = [x for x in cands if _cs_prog(x) >= 0]
    if prog:
        return max(prog, key=_cs_prog)
    return cands[0]

def _cs_floor(r, code):
    """الحالة ماتقلّش عن اللي التواريخ الفعلية بتقوله (مثلاً وصول فعلي للميناء = على الأقل «وصلت الميناء»)."""
    if code and _cs_prog(code) < 0:
        return code                                  # ملغي / موقوف بيفضلوا زي ما هم
    implied = ""
    if _cf_g(r, "warehouse_arrival_actual"):
        implied = "Delivered to Warehouse"
    elif _cf_g(r, "port_arrival_actual"):
        implied = "Arrived at Port"
    elif _cf_g(r, "actual_etd"):
        implied = "Shipped / In Transit"
    elif _cf_g(r, "actual_load_date"):
        implied = "Loaded"
    elif _cf_g(r, "container_no"):
        implied = "Allocated to Container"
    return implied if _cs_prog(implied) > _cs_prog(code) else code

def china_flat_prepare(flat_rows):
    """نسخ من الصفوف فيها status (عنوان عربي | English) محسوب من الأعمدة القديمة لو لازم — للكتابة في الشيت/الإكسيل."""
    out = []
    for r in flat_rows:
        d = dict(r)
        d["status"] = china_status_label(china_row_status(r))
        out.append(d)
    return out

def china_next_action_formula(R, L, link_prefix, row):
    """معادلة «المطلوب دلوقتي» كرابط بيودّي للخلية المطلوبة (عربي | English + [اسم العمود | key]).
    R(key) = مرجع خلية العمود في الصف، L(key) = حرف العمود، link_prefix = بداية الرابط (جوجل: "#gid=..&range=" / إكسيل: "#CHINA!")، row = رقم الصف.
    من غير = (صف واحد، مش ARRAYFORMULA)."""
    S = R("status")
    H = lambda w: f'ISNUMBER(SEARCH("{w}",{S}))'
    cn, ae, pa, wa = R("container_no"), R("actual_etd"), R("port_arrival_actual"), R("warehouse_arrival_actual")
    eta, tr, bl, prd = R("eta_port"), R("tracking_url"), R("bl_no"), R("production_ready_date")
    sl, seal = R("shipping_line"), R("seal_no")
    isline = f'LEN({R("barcode")}&{R("sku")}&{R("product_name_en")})>0'
    orderonly = f'(LEN({R("barcode")}&{R("sku")}&{R("product_name_en")})=0)*({cn}="")'
    blank = f'LEN({R("internal_order_no")}&{R("po_no")}&{R("supplier_name")}&{cn}&{R("barcode")}&{R("sku")})=0'
    cancel, deliv, cleared, customs, port = H("ملغي"), H("وصلت المستودع"), H("تم التخليص"), H("تحت التخليص"), H("وصلت الميناء")
    transit, loaded, prep, ready = H("في البحر"), H("تم التحميل"), H("حاوية متجهزة"), H("جاهز للشحن")
    done, prod, hold, factory = H("اكتمل الإنتاج"), H("قيد الإنتاج"), H("موقوف"), H("تأكيد المصنع")
    q = lambda t: '"' + t + '"'
    label = {x[1]: x[2] for x in CHINA_FLAT_COLS}
    def leaf(text, key=None):
        shown = text if key is None else (f"{text} [{label[key]}]" if "|" in label[key] else f"{text} [{label[key]} | {key}]")
        return ("leaf", shown, L(key or "status"))
    def iff(cond, a, b):
        return ("if", cond, a, b)
    def chain(pairs, default):
        node = default
        for cond, val in reversed(pairs):
            node = iff(cond, val, node)
        return node
    def render(node, which):
        if node[0] == "leaf":
            return q(node[1]) if which == "t" else q(node[2])
        return f'IF({node[1]},{render(node[2], which)},{render(node[3], which)})'
    tree = chain([
        (f'{S}=""', leaf("اختر الحالة | Select status", "status")),
        (cancel, leaf("—")),
        (deliv, leaf("✔ مكتمل | Completed")),
        *[(f'({isline})*({R(k)}="")', leaf(f"أدخل {ar} | Enter {en}", k)) for k, ar, en in (
            ("internal_order_no", "رقم الطلب", "order no."), ("barcode", "الباركود", "barcode"),
            ("total_qty", "الكمية", "quantity"), ("unit_cost", "تكلفة الوحدة", "unit cost"))],
        *[(f'({orderonly})*({R(k)}="")', leaf(f"أدخل {ar} | Enter {en}", k)) for k, ar, en in (
            ("po_no", "رقم PO", "PO no."), ("supplier_name", "المورد", "supplier"))],
        (cleared, iff(f'{wa}=""', leaf("سجّل وصول المستودع | Enter warehouse arrival", "warehouse_arrival_actual"),
                      leaf("غيّر الحالة لـ «وصلت المستودع» | Set status to Delivered", "status"))),
        (customs, iff(f'{pa}=""', leaf("سجّل وصول الميناء الفعلي | Enter port arrival date", "port_arrival_actual"),
                      leaf("تابع التخليص وحدّث الحالة | Follow customs, update status", "status"))),
        (port, iff(f'{pa}=""', leaf("سجّل وصول الميناء الفعلي | Enter port arrival date", "port_arrival_actual"),
                   leaf("حدّث حالة التخليص | Update customs status", "customs_status"))),
        (transit, chain([(f'{ae}=""', leaf("سجّل ETD الفعلي | Enter actual ETD", "actual_etd")),
                         (f'{eta}=""', leaf("سجّل ETA الميناء | Enter port ETA", "eta_port")),
                         (f'{sl}=""', leaf("سجّل شركة الشحن | Enter shipping line", "shipping_line")),
                         (f'{bl}=""', leaf("سجّل B/L | Enter B/L no.", "bl_no")),
                         (f'{tr}=""', leaf("أضف رابط التتبع | Add tracking link", "tracking_url"))],
                        leaf("لما توصل سجّل وصول الميناء | On arrival enter port arrival", "port_arrival_actual"))),
        (loaded, chain([(f'{cn}=""', leaf("سجّل رقم الحاوية | Enter container no.", "container_no")),
                        (f'{seal}=""', leaf("سجّل رقم الختم | Enter seal no.", "seal_no")),
                        (f'{sl}=""', leaf("سجّل شركة الشحن | Enter shipping line", "shipping_line")),
                        (f'{bl}=""', leaf("سجّل B/L | Enter B/L no.", "bl_no"))],
                       leaf("غيّر الحالة لـ «في البحر» عند الإبحار | Set In transit when sailed", "status"))),
        (prep, chain([(f'{cn}=""', leaf("سجّل رقم الحاوية | Enter container no.", "container_no")),
                      (f'{seal}=""', leaf("سجّل رقم الختم | Enter seal no.", "seal_no"))],
                     leaf("غيّر الحالة لـ «تم التحميل» بعد التحميل | Set Loaded after loading", "status"))),
        (ready, leaf("جهّز الحاوية: سجّل رقمها | Prepare container: enter its number", "container_no")),
        (done, leaf("أكّد جاهزية الشحن وغيّر الحالة | Confirm ready to ship, update status", "status")),
        (prod, iff(f'{prd}=""', leaf("سجّل تاريخ جاهزية الإنتاج | Enter production ready date", "production_ready_date"),
                   leaf("لما يخلص الإنتاج غيّر الحالة | Update status when production ends", "status"))),
        (hold, leaf("راجع سبب الإيقاف في الملاحظات | Review hold reason in notes", "order_notes")),
    ], iff(f'{prd}=""', leaf("سجّل تاريخ جاهزية الإنتاج المتوقع | Enter expected production ready date", "production_ready_date"),
           leaf("لما يبدأ الإنتاج غيّر الحالة | Update status when production starts", "status")))
    early_port = f'({transit})+({loaded})+({prep})+({ready})+({done})+({prod})+({factory})'
    early_sail = f'({loaded})+({prep})+({ready})+({done})+({prod})+({factory})'
    W = q("⚠ التواريخ أحدث من الحالة | Dates ahead of status — ")
    warn = (f'IF({S}="","",IF(({wa}<>"")*(1-({deliv}))*(1-({cancel}))>0,{W},'
            f'IF(({pa}<>"")*({early_port})>0,{W},IF(({ae}<>"")*({early_sail})>0,{W},""))))')
    return f'IF({blank},"",HYPERLINK({link_prefix}&{render(tree, "g")}&{row},{warn}&{render(tree, "t")}))'

def china_tables_to_flat(ships, conts, lines):
    """الـ 3 جداول -> صفوف الشيت الواحد. كل سطر منتج بياخد بيانات طلبه وحاويته؛ الطلبات اللي لسه ملهاش منتجات
    والحاويات اللي لسه ملهاش منتجات بتتحط في صفوف لوحدها — فمفيش حاجة بتضيع."""
    cmap = {_cf_g(c, "container_no"): c for c in conts if _cf_g(c, "container_no")}
    by_pair, by_order = {}, {}
    for x in ships:
        o, p = _cf_g(x, "internal_order_no"), _cf_g(x, "po_no")
        by_pair.setdefault((o, p), []).append(x)
        by_order.setdefault(o, []).append(x)
    def pick(cands, sid):
        if not cands:
            return None
        for x in cands:
            if sid and _cf_g(x, "shipment_id") == sid:
                return x
        return cands[0]
    def stamp(r, rec):
        t = _cf_g(rec, "last_updated_at")
        if t and t >= r["last_updated_at"]:
            r["last_updated_at"], r["updated_by"] = t, _cf_g(rec, "updated_by")
    def build(sh, ct, ln):
        r = {k: "" for k in CHINA_FLAT_KEYS}
        if sh:
            for k in _CF_SHIP_ONLY + _CF_SHARED:
                r[k] = _cf_g(sh, k)
            r["order_notes"] = _cf_g(sh, "notes")
            r["shipment_id"] = _cf_g(sh, "shipment_id")
            stamp(r, sh)
        if ct:
            for k in _CF_CONT_ONLY:
                r[k] = _cf_g(ct, k)
            dropped = []
            for k in _CF_SHARED:
                if _cf_g(ct, k):
                    if sh and _cf_g(sh, k) and _cf_g(sh, k) != _cf_g(ct, k):
                        dropped.append(f"{k} (في الطلب): {_cf_g(sh, k)}")      # قيمة الطلب المختلفة عن الحاوية — بنحفظها في الملاحظات بدل ما تضيع
                    r[k] = _cf_g(ct, k)
            if dropped:
                r["order_notes"] = " | ".join(x for x in [r["order_notes"]] + ["; ".join(dropped)] if x)
            for a, b in _CF_CONT_RENAME.items():
                r[b] = _cf_g(ct, a)
            if _cf_g(ct, "shipment_id") and not r["shipment_id"]:
                r["shipment_id"] = _cf_g(ct, "shipment_id")
            if _cf_g(ct, "final_destination") and not r["final_destination"]:
                r["final_destination"] = _cf_g(ct, "final_destination")
            stamp(r, ct)
        if ln:
            for k in _CF_LINE:
                r[k] = _cf_g(ln, k)
            if _cf_g(ln, "shipment_id") and not r["shipment_id"]:
                r["shipment_id"] = _cf_g(ln, "shipment_id")
            if _cf_g(ln, "container_no") and not r["container_no"]:
                r["container_no"] = _cf_g(ln, "container_no")
            for k in ("internal_order_no", "po_no"):
                if not r[k]:
                    r[k] = _cf_g(ln, k)
            if not r["currency"]:
                r["currency"] = _cf_g(ln, "currency")
            stamp(r, ln)
        src = {}
        if ln:
            src = {"stock_stage": _cf_g(ln, "stock_stage"), "container_status": _cf_g(ct, "container_status") if ct else ""}
        elif ct:
            src = {"container_status": _cf_g(ct, "container_status")}
        elif sh:
            src = {"shipment_status": _cf_g(sh, "shipment_status")}
        r["status"] = china_status_label(china_row_status(src))
        return r
    rows, used_ship, used_cont = [], set(), set()
    for l in lines:
        o, p, sid = _cf_g(l, "internal_order_no"), _cf_g(l, "po_no"), _cf_g(l, "shipment_id")
        # ربط السطر بالطلب بتطابق (رقم الطلب + PO) بالظبط — لو الـ PO مختلف بيفضل صف الطلب لوحده ومفيش بيانات بتتدمج غلط
        sh = pick(by_pair.get((o, p)) or [], sid)
        ct = cmap.get(_cf_g(l, "container_no"))
        if sh is not None:
            used_ship.add(id(sh))
        if ct is not None:
            used_cont.add(_cf_g(ct, "container_no"))
        rows.append(build(sh, ct, l))
    for x in ships:
        if id(x) not in used_ship:
            rows.append(build(x, None, None))
    for c in conts:
        if _cf_g(c, "container_no") not in used_cont:
            rows.append(build(None, c, None))
    def sk(r):
        o = r["internal_order_no"]
        n = _cf_num(o)
        return (0 if o else 1, -n, r["po_no"], r["container_no"], r["sku"])
    rows.sort(key=sk)
    return rows

def _cs_container(codes):
    codes = [x for x in codes if x]
    if not codes:
        return ""
    nonc = [x for x in codes if x != "Cancelled"]
    if not nonc:
        return "Cancelled"
    prog = [x for x in nonc if _cs_prog(x) >= 0]
    return max(prog, key=_cs_prog) if prog else nonc[0]

def _cs_shipment(codes):
    codes = [x for x in codes if x]
    if not codes:
        return ""
    nonc = [x for x in codes if x != "Cancelled"]
    if not nonc:
        return "Cancelled"
    if "On Hold" in nonc:
        return "On Hold"
    prog = [x for x in nonc if _cs_prog(x) >= 0]
    return min(prog, key=_cs_prog) if prog else nonc[0]

def china_flat_to_tables(flat_rows):
    """صفوف الشيت الواحد -> (ships, conts, lines, conflicts) بنفس شكل الجداول القديمة."""
    rows = [{k: _cf_g(r, k) for k in CHINA_FLAT_INPUT_KEYS + CHINA_STATUS_LEGACY} for r in flat_rows]
    rows = [r for r in rows if any(r.values())]
    for r in rows:
        r["_code"] = _cs_floor(r, china_row_status(r))
    conflicts = []
    ctr_status = {}
    for r in rows:
        if r["container_no"]:
            ctr_status.setdefault(r["container_no"], []).append(r["_code"])
    ctr_status = {k: _cs_container(v) for k, v in ctr_status.items()}
    pair_codes = {}
    for r in rows:
        pair_codes.setdefault((r["internal_order_no"], r["po_no"]), []).append(r["_code"])
    # ── الحاويات ──
    cgroups = {}
    for r in rows:
        if r["container_no"]:
            cgroups.setdefault(r["container_no"], []).append(r)
    conts = []
    for cno, grp in cgroups.items():
        c = {k: "" for k in CHINA_CONT_COLS}
        for k in _CF_CONT_ONLY + _CF_SHARED:
            c[k] = _cf_pick(grp, k, conflicts, f"حاوية {cno}")
        for a, b in _CF_CONT_RENAME.items():
            c[a] = _cf_pick(grp, b, conflicts, f"حاوية {cno}")
        c["shipment_id"] = _cf_pick(grp, "shipment_id", conflicts, f"حاوية {cno}")
        c["final_destination"] = _cf_pick(grp, "final_destination", [], "")
        c["last_updated_at"] = max([r["last_updated_at"] for r in grp] or [""])
        c["updated_by"] = next((r["updated_by"] for r in grp if r["last_updated_at"] == c["last_updated_at"] and r["updated_by"]), "")
        c["container_status"] = _cs_to_cont(ctr_status.get(cno, ""))
        c["record_key"] = "CTR|" + cno
        conts.append(c)
    # ── الشحنات/الطلبات ──
    sgroups = {}
    for r in rows:
        is_line = bool(r["barcode"] or r["sku"] or r["model_no"] or r["product_name_en"] or r["product_name_ar"] or r["line_id"])
        has_ship = any(r[k] for k in _CF_SHIP_DATA)       # بيانات خاصة بالطلب نفسه (مش مجرد رقم الطلب على سطر المنتج)
        bare = (not r["container_no"]) and (not is_line) and bool(r["internal_order_no"] or r["po_no"] or r["shipment_id"])
        if has_ship or bare:
            key = (r["internal_order_no"], r["po_no"], r["invoice_value"])
            sgroups.setdefault(key, []).append(r)
    ships = []
    for key, grp in sgroups.items():
        s = {k: "" for k in CHINA_SHIP_COLS}
        label = f"طلب {key[0]} / {key[1]}"
        s["shipment_id"] = _cf_pick(grp, "shipment_id", conflicts, label)
        for k in _CF_SHIP_ONLY + _CF_SHARED:
            s[k] = _cf_pick(grp, k, conflicts, label)
        s["notes"] = _cf_pick(grp, "order_notes", conflicts, label)
        s["last_updated_at"] = max([r["last_updated_at"] for r in grp] or [""])
        s["updated_by"] = next((r["updated_by"] for r in grp if r["last_updated_at"] == s["last_updated_at"] and r["updated_by"]), "")
        s["shipment_status"] = _cs_to_ship(_cs_shipment(pair_codes.get((key[0], key[1]), [])))
        s["record_key"] = china_ship_record_key(s)
        ships.append(s)
    # ── أسطر المنتجات ──
    lines, seq = [], {}
    taken = {r["line_id"] for r in rows if r["line_id"]}
    for r in rows:
        if not (r["barcode"] or r["sku"] or r["model_no"] or r["product_name_en"] or r["product_name_ar"] or r["line_id"]):
            continue
        l = {k: "" for k in CHINA_LINE_COLS}
        for k in _CF_LINE:
            l[k] = r[k]
        for k in ("shipment_id", "internal_order_no", "po_no", "currency", "container_no", "last_updated_at", "updated_by"):
            l[k] = r[k]
        code, ccode = r["_code"], ctr_status.get(r["container_no"], "")
        if _cs_prog(code) >= 0 and _cs_prog(ccode) > _cs_prog(code):
            code = ccode                       # الصف بياخد حالة حاويته لو أحدث منه
        l["stock_stage"] = _cs_to_stock(code)
        if not l["line_id"]:
            base = re.sub(r"[^A-Za-z0-9]", "", r["internal_order_no"]) or "LINE"
            n = seq.get(base, 0)
            while True:
                n += 1
                cand = f"{base}-{n:03d}"
                if cand not in taken:
                    break
            seq[base] = n
            taken.add(cand)
            l["line_id"] = cand
        try:
            if l["total_qty"] and l["unit_cost"]:
                v = round(float(l["total_qty"].replace(",", "")) * float(l["unit_cost"].replace(",", "")), 2)
                l["line_amount"] = str(int(v)) if v == int(v) else f"{v:.2f}".rstrip("0").rstrip(".")
        except Exception:
            pass
        l["record_key"] = "SKU|" + l["line_id"]
        lines.append(l)
    return ships, conts, lines, conflicts

def _china_col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def _china_rgb(h):
    return {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255, "blue": int(h[4:6], 16) / 255}

def _china_auto_formulas():
    c = {k: _china_col_letter(i + 1) for i, k in enumerate(CHINA_FLAT_KEYS)}
    R = lambda k: f"{c[k]}4:{c[k]}"
    qty, cost, eta, pa, S = R("total_qty"), R("unit_cost"), R("eta_port"), R("port_arrival_actual"), R("status")
    done = f'(ISNUMBER(SEARCH("وصلت المستودع",{S}))+ISNUMBER(SEARCH("ملغي",{S}))>0)'
    delay_v = f'IF({pa}<>"",{pa}-{eta},IF({done},0,TODAY()-{eta}))'
    ident = f'{R("barcode")}&{R("sku")}&{R("product_name_en")}'
    return {
        "line_amount": f'=ARRAYFORMULA(IF(({qty}<>"")*({cost}<>""),{qty}*{cost},""))',
        "days_to_eta": f'=ARRAYFORMULA(IF({eta}="","",IFERROR({eta}-TODAY(),"")))',
        "delay_days": f'=ARRAYFORMULA(IF({eta}="","",IFERROR(IF({delay_v}>0,{delay_v},0),"")))',
        "row_state": (f'=ARRAYFORMULA(IF(LEN({R("internal_order_no")}&{R("po_no")}&{R("supplier_name")}&{R("container_no")}&{ident})=0,"",'
                      f'IF(LEN({ident})>0,IF(({R("internal_order_no")}="")+({R("barcode")}="")+({qty}="")+({cost}="")+({S}="")>0,"MISSING","OK"),'
                      f'IF({R("container_no")}<>"",IF(({S}="")+(({eta}="")*(ISERROR(SEARCH("ملغي",{S}))))>0,"MISSING","OK"),'
                      f'IF(({R("supplier_name")}="")+({R("po_no")}="")+({S}="")>0,"MISSING","OK")))))'),
    }

def _china_image_formula(r):
    """صورة المنتج من الباركود: رابط الصورة اليدوي لو موجود، وإلا من نفس مصادر المبيعات
    (Tacweed -> links n)، ثم (TacweedAmazon -> links a)، ثم منتجات تريندول. بتتحجّم على الخلية."""
    c = {k: _china_col_letter(i + 1) for i, k in enumerate(CHINA_FLAT_KEYS)}
    bc, u = f"${c['barcode']}{r}", f"${c['image_url']}{r}"
    def via(code_tab, links_tab):
        return (f"INDEX(ARRAYFORMULA(FILTER('{links_tab}'!$B$2:$B,ISNUMBER(MATCH('{links_tab}'!$A$2:$A,"
                f"FILTER({code_tab}!$A$2:$A,{code_tab}!$B$2:$B&\"\"={bc}&\"\"),0))*('{links_tab}'!$B$2:$B<>\"\"))),1)")
    ci = f"INDEX(FILTER(ChinaImages!$B$2:$B,ChinaImages!$A$2:$A&\"\"={bc}&\"\",ChinaImages!$B$2:$B<>\"\"),1)"
    ty = f"INDEX(FILTER(TrendyolProducts!$D$2:$D,TrendyolProducts!$B$2:$B&\"\"={bc}&\"\",TrendyolProducts!$D$2:$D<>\"\"),1)"
    return (f'=IF({bc}="","",IFERROR(IMAGE(IF({u}<>"",{u},IFERROR({ci},IFERROR({via("Tacweed","links n")},'
            f'IFERROR({via("TacweedAmazon","links a")},{ty})))),1),"— لا صورة"))')

def _china_next_action_row(r, gid):
    c = {k: _china_col_letter(i + 1) for i, k in enumerate(CHINA_FLAT_KEYS)}
    return "=" + china_next_action_formula(lambda k: f"{c[k]}{r}", lambda k: c[k], f'"#gid={gid}&range="', r)

def _china_layout_outdated(ws):
    """هل صف المفاتيح (صف 2) مش مطابق للترتيب الحالي؟ (شيت اتعمل بنسخة أقدم من الكود)"""
    try:
        keys = [str(c).strip() for c in ws.row_values(2)]
    except Exception:
        return False
    keys = [k for k in keys if k]
    return "container_no" in keys and keys != CHINA_FLAT_KEYS

def _china_cell(v):
    v = "" if v is None else str(v)
    return ("'" + v) if v[:1] in ("=", "+", "@") else v

def _china_input_segments():
    """مجموعات أعمدة الإدخال المتصلة (عشان أعمدة المعادلات في النص متتمسحش)."""
    segs, cur = [], []
    for i, k in enumerate(CHINA_FLAT_KEYS):
        if k in CHINA_FLAT_FORMULA_KEYS:
            if cur:
                segs.append(cur)
            cur = []
        else:
            cur.append((i, k))
    if cur:
        segs.append(cur)
    return segs

def _china_is_label_row(d):
    """صف العناوين العربي (تحت صف المفاتيح) — مش بيانات."""
    return sum(1 for k, v in d.items() if v and v == _CHINA_LABELS.get(k)) >= 3

def _china_read_flat(raw):
    """الصفوف الخام من الشيت -> قايمة dicts بمفاتيح الأعمدة (بتدوّر على صف المفاتيح، فمش مهم لو اتضاف صفوف فوق)."""
    if not raw:
        return []
    hi = None
    for i, r in enumerate(raw[:8]):
        low = {str(c).strip() for c in r}
        if "container_no" in low and "internal_order_no" in low:
            hi = i
            break
    if hi is None:
        return []
    pos = {str(c).strip(): j for j, c in enumerate(raw[hi]) if str(c).strip() in CHINA_FLAT_INPUT_KEYS + CHINA_STATUS_LEGACY}
    out = []
    for row in raw[hi + 1:]:
        d = {k: (str(row[j]).strip() if j < len(row) else "") for k, j in pos.items()}
        if any(d.values()) and not _china_is_label_row(d):
            out.append(d)
    return out

def _china_current_tables(env, force=False):
    return china_flat_to_tables(_china_read_flat(env._china_flat_raw(force)))

def _china_clean_sheet_row(r, cols):
    return {c: _china_clean(r.get(c), c) for c in cols if c not in CHINA_SKIP_COLS}

def _china_header_row(rows, keys):
    best, best_n = None, 0
    for i, r in enumerate(rows[:15]):
        n = sum(1 for c in r if c is not None and str(c).strip() in keys)
        if n > best_n:
            best, best_n = i, n
    return best if best_n >= 6 else None

def _china_parse_sheet(ws, cols):
    rows = list(ws.iter_rows(values_only=True))
    hdr_i = _china_header_row(rows, set(cols))
    if hdr_i is None:
        return None
    hdr = []
    for j, h in enumerate(rows[hdr_i]):
        h = str(h).strip() if h is not None else ""
        if h not in cols and j < len(cols):     # عنوان تالف (زي 1268 بدل shipment_id): بنعتمد على ترتيب العمود المتوقع
            h = cols[j] if cols[j] not in [str(x).strip() for x in rows[hdr_i] if x is not None] else h
        hdr.append(h)
    out = []
    for r in rows[hdr_i + 1:]:
        if not r or not any(c not in (None, "") for c in r):
            continue
        d = {}
        for j, h in enumerate(hdr):
            if h in cols and h not in CHINA_SKIP_COLS and j < len(r):
                d[h] = _china_clean(r[j], h)
        if any(d.values()):
            out.append(d)
    return out

def _china_diff(env, parsed):
    """مقارنة الملف المرفوع بالموجود في شيت China: جديد / اتغيّر / موجود عندنا ومش في الملف / ثابت."""
    cur_ships, cur_conts, cur_lines, _ = _china_current_tables(env, force=True)
    specs = [("SHIPMENTS", CHINA_SHIP_COLS, "ships", _china_ship_key, cur_ships),
             ("SKU_LINES", CHINA_LINE_COLS, "lines", _china_line_key, cur_lines),
             ("CONTAINERS", CHINA_CONT_COLS, "conts", _china_cont_key, cur_conts)]
    summary, details = [], []
    skip = CHINA_SKIP_COLS | {"record_key", "row_state", "line_amount", "last_updated_at", "updated_by"}
    for label, cols, pk, keyfn, cur in specs:
        new_rows = parsed.get(pk)
        if new_rows is None:
            continue
        old = {keyfn(r): r for r in cur}
        new = {keyfn(r): r for r in new_rows}
        added = [k for k in new if k not in old]
        removed = [k for k in old if k not in new]
        changed = []
        for k in new:
            if k in old:
                diffs = [(c, old[k].get(c, ""), new[k].get(c, "")) for c in cols
                         if c not in skip and old[k].get(c, "") != new[k].get(c, "")]
                if diffs:
                    changed.append(k)
                    for c, ov, nv in diffs:
                        details.append({"الصفحة": label, "المفتاح": k, "الحقل": c, "قبل": ov, "بعد": nv})
        summary.append({"الصفحة": label, "في الملف": len(new), "جديد": len(added), "اتغيّر": len(changed),
                        "ثابت": len(new) - len(added) - len(changed), "عندنا ومش في الملف": len(removed)})
    return summary, details

def _china_enrich(line, conts, ship_by_pair, ship_by_no):
    ctr = conts.get(line.get("container_no", "")) or {}
    shp = (ship_by_pair.get((line.get("internal_order_no", ""), line.get("po_no", "")))
           or ship_by_no.get(line.get("internal_order_no", "")) or {})
    stage = line.get("stock_stage", "")
    cstat = ctr.get("container_status", "")
    wh_arr = ctr.get("warehouse_arrival_actual", "")
    # المرحلة الفعلية = الأبعد بين مرحلة السطر وحالة الحاوية وتواريخ الوصول/التخليص الفعلية
    idx = CHINA_STAGES.index(stage) if stage in CHINA_STAGES else -1
    if ctr:
        idx = max(idx, CHINA_CONT_STATUS_IDX.get(cstat, -1))
        if ctr.get("port_arrival_actual"):
            idx = max(idx, 7)
        cust = ctr.get("customs_status", "")
        if cust == "Cleared":
            idx = max(idx, 9)
        elif cust in ("Submitted", "Under Clearance", "Inspection") and idx >= 6:
            idx = max(idx, 8)
        if wh_arr:
            idx = max(idx, 10)
    if stage == "Cancelled" or cstat == "Cancelled":
        status = "cancelled"
    elif idx >= CHINA_RECEIVED_IDX:
        status = "received"
    else:
        status = "incoming"
    if wh_arr:
        eta_label, eta = "وصل المستودع", wh_arr
    elif ctr.get("port_arrival_actual"):
        eta_label, eta = "وصل الميناء", ctr["port_arrival_actual"]
    elif ctr.get("eta_port"):
        eta_label, eta = "الوصول المتوقع للميناء", ctr["eta_port"]
    elif shp.get("eta_destination"):
        eta_label, eta = "الوصول المتوقع لمخزن الرياض", shp["eta_destination"]
    elif shp.get("eta_port"):
        eta_label, eta = "الوصول المتوقع للميناء", shp["eta_port"]
    else:
        eta_label, eta = "", ""
    e = dict(line)
    # كل التواريخ المهمة للسطر (الحاوية أولاً ثم الشحنة كبديل) | Key dates per line (container first, shipment as fallback)
    e["_d"] = {
        "ready": line.get("production_ready_date", ""),
        "load": line.get("actual_load_date", ""),
        "etd_plan": shp.get("planned_etd", ""),
        "etd": ctr.get("actual_etd") or shp.get("actual_etd", ""),
        "eta_port": ctr.get("eta_port") or shp.get("eta_port", ""),
        "port_arr": ctr.get("port_arrival_actual", ""),
        "eta_dest": shp.get("eta_destination", ""),
        "wh_arr": wh_arr,
    }
    e["_upd"] = ctr.get("last_updated_at") or line.get("last_updated_at", "")
    e["_upd_by"] = ctr.get("updated_by") or line.get("updated_by", "")
    e["_qty"] = _china_int(line.get("total_qty", ""))
    e["_status"] = status
    e["_idx"] = idx
    e["_stage_eff"] = CHINA_STAGE_AR.get(CHINA_STAGES[idx], "") if idx >= 0 else CHINA_STAGE_AR.get(stage, stage)
    e["_eta_label"] = eta_label
    e["_eta"] = eta
    e["_ship"] = shp
    e["_ctr"] = ctr
    e["_shipment_id"] = line.get("shipment_id") or shp.get("shipment_id", "")
    e["_norm"] = _china_norm_code(line.get("barcode", ""))
    return e

def _china_code_to_skus(env):
    """كود (موحّد) -> قائمة SKUs المتجر المرتبطة بيه من تاب التكويد | normalised code -> store SKUs."""
    m = {}
    for sku_up, code in get_tacweed_map(env).items():
        m.setdefault(_china_norm_code(code), []).append(sku_up)
    # ASINs أمازون المرتبطة بنفس الكود (من TacweedAmazon) | Amazon ASINs on the same code
    try:
        for asin_up, code in get_tacweed_amazon_map(env).items():
            m.setdefault(_china_norm_code(code), []).append(asin_up)
    except Exception:
        pass
    return m

def china_lines_for_sku(env, sku_up):
    """أسطر الصين المرتبطة بـ SKU: عن طريق كود التكويد = barcode (الأساسي)، وكمان لو
    noon_sku في الملف = نفس الـ SKU | China lines linked to a SKU: via Tacweed code ==
    barcode (primary), plus an exact noon_sku match."""
    b = env.get_china_bundle()
    if not b["lines"]:
        return []
    cands = []
    code = get_tacweed_map(env).get(sku_up, "") or get_tacweed_amazon_map(env).get(sku_up, "")
    if code:
        cands += b["by_code"].get(_china_norm_code(code), [])
    cands += b["by_noon_sku"].get(sku_up, [])
    out, seen = [], set()
    for l in cands:
        if id(l) not in seen:
            seen.add(id(l))
            out.append(l)
    return out

def _china_group(idx):
    return "في المصنع" if idx < CHINA_SHIPPED_IDX else ("في البحر" if idx < 7 else "ميناء/تخليص")

def china_batches_for_lines(lines):
    """الأسطر القادمة: دفعات بتاريخ وصول الوجهة [{date, group, qty}] + دفعات بدون تاريخ [{group, qty, eta_port}]."""
    dated, undated = {}, {}
    for l in lines:
        if l["_status"] != "incoming":
            continue
        dte, _b = _china_avail(l)
        grp = _china_group(l["_idx"])
        if dte:
            dated[(dte, grp)] = dated.get((dte, grp), 0) + l["_qty"]
        else:
            k = (grp, l["_d"]["eta_port"])
            undated[k] = undated.get(k, 0) + l["_qty"]
    return ([{"date": k[0], "group": k[1], "qty": q} for k, q in sorted(dated.items())],
            [{"group": k[0], "eta_port": k[1], "qty": q} for k, q in sorted(undated.items())])

def _china_stock_rate(env, sku_up, row=None):
    """(مخزون, معدل سحب يومي, مصدر, عدد SKUs في الكود). الأساس = المتوفر بالمستودع ومعدل سحب الكود (مجمّع)؛
    لو الكود مالوش مخزون مستودع مسجل بيرجع مخزون نون كبديل بعلامة واضحة."""
    wh = env._china_wh_lookup().get(sku_up)
    if wh:
        return float(wh.get("qty") or 0), float(wh.get("daily_draw") or 0), "wh", len(wh.get("skus") or []) + len(wh.get("asins") or [])
    if row is not None:
        return float(row.get("stock") or 0), float(row.get("effective_avg") or 0), "noon", 1
    return None

def _china_fmt_date(d, today=None):
    """تاريخ بالعربي: 27 أكتوبر 2026 (بعد 8 يوم) — مفيش مشكلة اتجاه أرقام زي 2026-10-27."""
    if today is not None and (d - today).days > 730:
        return "أكتر من سنتين"
    t = f"{d.day} {_AR_MONTHS[d.month - 1]} {d.year}"
    if today is not None:
        n = (d - today).days
        t += " (النهارده)" if n == 0 else (f" (بعد {n} يوم)" if n > 0 else f" (فات من {-n} يوم)")
    return t

def china_code_decisions(env, bundle, today, cfg):
    """قرار لكل كود صين عنده كمية قادمة: (مخزون المستودع + معدل سحب الكود + دفعات الصين المؤرخة). بيرجع (rows, no_stock_codes)."""
    code_skus = _china_code_to_skus(env)
    lookup = env._china_wh_lookup()
    rows, no_stock = [], 0
    for norm, ls in bundle["by_code"].items():
        inc = [l for l in ls if l["_status"] == "incoming"]
        if not inc:
            continue
        skus = code_skus.get(norm, [])
        wh = next((lookup[s] for s in skus if s in lookup), None)
        if not wh:
            no_stock += 1
            continue
        dated, undated = china_batches_for_lines(inc)
        dec = china_decide(float(wh.get("qty") or 0), float(wh.get("daily_draw") or 0),
                           [(b["date"], b["qty"]) for b in dated], sum(u["qty"] for u in undated), today, cfg)
        dec["near"] = cfg["china_near_days"]
        rows.append({"dec": dec, "code": ls[0].get("barcode", norm), "sku": (skus[0] if skus else ""),
                     "name": ls[0].get("product_name_ar") or ls[0].get("product_name_en", ""),
                     "incoming": sum(l["_qty"] for l in inc), "undated": dec.get("undated", 0)})
    return rows, no_stock

def _china_alert(sev, icon, title, action, rows):
    return {"sev": sev, "icon": icon, "title": title, "action": action, "rows": rows}

def china_build_alerts(env, bundle, today, cfg):
    """كل التنبيهات مرتبة بالخطورة (0 عاجل، 1 تنبيه، 2 معلومة). كلها من تواريخ الملف + مخزون المستودع."""
    near, stale = cfg["china_near_days"], cfg["china_stale_days"]
    F = lambda d: _china_fmt_date(d)
    alerts, on_track = [], 0
    lines, conts = bundle["lines"], bundle["conts"]

    # ── منتجات: فجوة / موعد طلب ──
    prod, no_stock = china_code_decisions(env, bundle, today, cfg)
    def prow(r, extra):
        d = r["dec"]
        base = {"الكود": r["code"], "SKU": r["sku"], "المنتج": r["name"], "المتوفر بالمستودع": f"{d['stock']:,.0f}",
                "سحب/يوم": f"{d['rate']:.1f}", "يكفي لحد": F(d["stockout"])}
        base.update(extra)
        return base
    gap = [r for r in prod if r["dec"]["state"] == "gap"]
    late = [r for r in prod if r["dec"]["state"] == "order_late"]
    soon = [r for r in prod if r["dec"]["state"] == "order_soon"]
    needd = [r for r in prod if r["dec"]["state"] == "need_date"]
    if gap:
        alerts.append(_china_alert(0, "🔴", "منتجات هتخلص قبل وصول شحنة الصين (فجوة بيع)",
            "اضغط على الصين تسرّع الشحنات دي، وفكّر في شحن جوي أو شراء بديل للفجوة، واطلب كمية جديدة.",
            [prow(r, {"تاريخ وصول الشحنة": F(r["dec"]["gap_to"]), "أيام من غير بضاعة": r["dec"]["gap_days"],
                      "مبيعات معرضة للضياع": f"{r['dec']['lost']:,}", "كان لازم نطلب قبل": F(r["dec"]["should_order"]),
                      "كمية مقترحة": f"{r['dec']['suggested']:,}"})
             for r in sorted(gap, key=lambda r: -r["dec"]["lost"])]))
    if late:
        alerts.append(_china_alert(0, "🔴", "كان لازم نطلب من الصين وفات الميعاد",
            "اطلب دلوقتي — كل يوم تأخير بيقرّب الانقطاع.",
            [prow(r, {"المخزون + الشحنات يغطوا لحد": F(r["dec"]["covered"]), "كان لازم نطلب قبل": F(r["dec"]["order_by"]),
                      "فات من (يوم)": r["dec"]["late_by"], "كمية مقترحة": f"{r['dec']['suggested']:,}"})
             for r in sorted(late, key=lambda r: -r["dec"]["late_by"])]))
    if soon:
        alerts.append(_china_alert(1, "🟠", f"قرب موعد الطلب من الصين (خلال {cfg['china_reorder_safety_days']} يوم)",
            "جهّز طلب الشراء دلوقتي عشان متوصلش لمرحلة التأخير.",
            [prow(r, {"المخزون + الشحنات يغطوا لحد": F(r["dec"]["covered"]), "اطلب قبل": F(r["dec"]["order_by"]),
                      "باقي (يوم)": -r["dec"]["late_by"], "كمية مقترحة": f"{r['dec']['suggested']:,}"})
             for r in sorted(soon, key=lambda r: r["dec"]["order_by"])]))
    if needd:
        alerts.append(_china_alert(1, "🟠", "منتجات قرب تخلص وشحنتها من غير تاريخ وصول لمخزن الرياض",
            "اطلب من الصين تكتب تاريخ الوصول لمخزن الرياض في الملف عشان أقدر أحسب الفجوة.",
            [prow(r, {"قطع جاية من غير تاريخ": f"{r['undated']:,}", "اطلب قبل (لو مفيش شحنة تغطي)": F(r["dec"]["order_by"])})
             for r in sorted(needd, key=lambda r: r["dec"]["stockout"])]))

    # ── حاويات ──
    units_by_cont = {}
    for l in lines:
        cn = l.get("container_no", "")
        if cn and l["_status"] != "cancelled":
            units_by_cont[cn] = units_by_cont.get(cn, 0) + l["_qty"]
    c_late, c_soon, c_stuck = [], [], []
    for c in conts:
        if c.get("container_status") == "Cancelled" or _china_date(c.get("warehouse_arrival_actual")):
            continue
        cn = c.get("container_no", "")
        eta, arr = _china_date(c.get("eta_port")), _china_date(c.get("port_arrival_actual"))
        base = {"الحاوية": cn, "الشحنة": c.get("shipment_id", ""), "القطع": (f"{units_by_cont[cn]:,}" if units_by_cont.get(cn) else "—")}
        if arr:
            n = (today - arr).days
            if n > near:
                c_stuck.append(dict(base, **{"وصلت الميناء": F(arr), "بقالها (يوم)": n, "التخليص": c.get("customs_status", "") or "—"}))
            else:
                on_track += 1
        elif eta:
            d = (eta - today).days
            if d < 0:
                c_late.append(dict(base, **{"كان متوقع يوصل الميناء": F(eta), "متأخرة (يوم)": -d,
                                            "السبب المكتوب": c.get("delay_reason", "") or "—"}))
            elif d <= near:
                c_soon.append(dict(base, **{"متوقع يوصل الميناء": F(eta), "باقي (يوم)": d}))
            else:
                on_track += 1
    if c_late:
        alerts.append(_china_alert(0, "🔴", "حاويات اتأخرت عن ميعاد وصولها للميناء",
            "اسأل الصين/شركة الشحن عن الميعاد الجديد، وبلّغ المسؤول عن المنتجات اللي فيها لو مبيعاتها هتتأثر.",
            sorted(c_late, key=lambda r: -r["متأخرة (يوم)"])))
    if c_stuck:
        alerts.append(_china_alert(1, "🟠", f"حاويات وصلت الميناء من أكتر من {near} يوم ولسه ماوصلتش المستودع",
            "تابع التخليص الجمركي والنقل للمستودع — البضاعة واقفة.", sorted(c_stuck, key=lambda r: -r["بقالها (يوم)"])))
    if c_soon:
        alerts.append(_china_alert(2, "🔵", f"حاويات هتوصل الميناء خلال {near} يوم",
            "جهّز التخليص الجمركي والنقل للمستودع.", sorted(c_soon, key=lambda r: r["باقي (يوم)"])))

    # ── أسطر: مصنع / تاريخ الوصول لمخزن الرياض ──
    f_rows, miss, d_late, d_soon = {}, {}, {}, {}
    for l in lines:
        if l["_status"] != "incoming":
            continue
        sid = l["_shipment_id"] or l.get("internal_order_no", "") or "—"
        d = l["_d"]
        if l["_idx"] < CHINA_SHIPPED_IDX:
            ready, plan = _china_date(d["ready"]), _china_date(d["etd_plan"])
            over = max([(today - x).days for x in (ready, plan) if x] or [0])
            if over > 0:
                r = f_rows.setdefault(sid, {"الشحنة": sid, "القطع": 0, "جاهزية الإنتاج": ready, "الخروج المخطط": plan, "over": 0})
                r["القطع"] += l["_qty"]
                r["over"] = max(r["over"], over)
        else:
            ed = _china_date(d["eta_dest"])
            if not ed:
                r = miss.setdefault(sid, {"الشحنة": sid, "القطع": 0, "وصول الميناء المتوقع": d["eta_port"] or "—"})
                r["القطع"] += l["_qty"]
            else:
                n = (ed - today).days
                tgt = d_late if n < 0 else (d_soon if n <= near else None)
                if tgt is not None:
                    r = tgt.setdefault(sid, {"الشحنة": sid, "القطع": 0, "تاريخ الوصول لمخزن الرياض": F(ed), "الأيام": abs(n)})
                    r["القطع"] += l["_qty"]
    if f_rows:
        rows = []
        for r in sorted(f_rows.values(), key=lambda r: -r["over"]):
            rows.append({"الشحنة": r["الشحنة"], "القطع": r["القطع"], "جاهزية الإنتاج": F(r["جاهزية الإنتاج"]) if r["جاهزية الإنتاج"] else "—",
                         "الخروج المخطط": F(r["الخروج المخطط"]) if r["الخروج المخطط"] else "—", "فات من (يوم)": r["over"]})
        sev = 0 if any(r["فات من (يوم)"] > near for r in rows) else 1
        alerts.append(_china_alert(sev, "🔴" if sev == 0 else "🟠", "بضاعة جاهزة أو ميعاد شحنها فات ولسه ماتشحنتش من الصين",
            "كلّم الصين/المصنع وحدد ميعاد التحميل.", rows))
    if d_late:
        alerts.append(_china_alert(0, "🔴", "شحنات فات ميعاد وصولها لمخزن الرياض ولسه ماوصلتش",
            "اسأل عن الموعد الجديد وحدّث الخطة — المنتجات دي ممكن تكون في فجوة بيع.",
            [dict(r, **{"فات من (يوم)": r.pop("الأيام")}) for r in sorted(d_late.values(), key=lambda r: -r["الأيام"])]))
    if d_soon:
        alerts.append(_china_alert(2, "🔵", f"شحنات هتوصل مخزن الرياض خلال {near} يوم",
            "جهّز الاستلام والتخزين.", [dict(r, **{"باقي (يوم)": r.pop("الأيام")}) for r in sorted(d_soon.values(), key=lambda r: r["الأيام"])]))
    if miss:
        alerts.append(_china_alert(1, "🟠", "شحنات خرجت من الصين وتاريخ وصولها لمخزن الرياض غير مكتوب",
            "اطلب من الصين تكتب تاريخ الوصول لمخزن الرياض — من غيره مش هقدر أحسب «متاح للبيع» ولا الفجوة.",
            sorted(miss.values(), key=lambda r: -r["القطع"])))

    # ── تنبيهات على مستوى الكود عبر المتاجر الثلاثة (نون + أمازون + تريندول) ──
    try:
        xrows, _unm, _sw = env._cross_cached()
        wh_codes = env.get_warehouse_stock_map()
        inc_by = {}
        for l in lines:
            if l["_status"] == "incoming" and l["_norm"]:
                inc_by[l["_norm"]] = inc_by.get(l["_norm"], 0) + l["_qty"]
        seen_norm = set(inc_by) | {l["_norm"] for l in lines if l["_norm"]}
        dead_inc, no_ship = [], []
        lead_total = cfg["china_lead_time_days"] + cfg["china_reorder_safety_days"]
        for r in xrows:
            nc = _china_norm_code(r["code"])
            if nc in inc_by and r["total"] == 0 and r["prev_total"] == 0:
                dead_inc.append({"الكود": r["code"], "قطع جاية من الصين": f"{inc_by[nc]:,}", "مبيعات 60 يوم (نون+أمازون+تريندول)": 0,
                                 "مخزون المستودع": f"{r['wh']:,.0f}"})
            elif nc in inc_by and r["total"] == 0 and r["wh"] > 0:
                dead_inc.append({"الكود": r["code"], "قطع جاية من الصين": f"{inc_by[nc]:,}", "مبيعات 60 يوم (نون+أمازون+تريندول)": r["prev_total"],
                                 "مخزون المستودع": f"{r['wh']:,.0f}"})
            if nc not in seen_norm and r["code"] in wh_codes and r["total"] >= 30:
                rate = r["total"] / 30.0
                days = r["wh"] / rate
                if days <= lead_total + 30:
                    no_ship.append({"الكود": r["code"], "مبيعات 30 يوم": r["total"], "نون/أمازون/تريندول": f"{r['noon']}/{r['amz']}/{r['ty']}",
                                    "مخزون المستودع": f"{r['wh']:,.0f}", "يكفي (يوم)": round(days), "مدة التوريد + الأمان": lead_total})
        if no_ship:
            no_ship.sort(key=lambda x: x["يكفي (يوم)"])
            sev = 0 if no_ship[0]["يكفي (يوم)"] <= lead_total else 1
            alerts.append(_china_alert(sev, "🔴" if sev == 0 else "🟠", "كود بيبيع في المتاجر ومفيش له أي شحنة من الصين ومخزونه قرب يخلص",
                "الأكواد دي مالهاش سطر في ملف الصين خالص — اطلب شحنة جديدة أو ضيفها للملف.", no_ship[:200]))
        if dead_inc:
            alerts.append(_china_alert(1, "🟠", "بضاعة جاية من الصين لكود مبيعاته صفر في كل المتاجر آخر 30 يوم",
                "راجع الطلب قبل ما الشحنة توصل: ممكن يتأجل أو يتلغى أو المنتج محتاج تفعيل/إعلان.", dead_inc[:200]))
    except Exception:
        pass

    # ── الملف قديم؟ ──
    upd = [d for d in (_china_date(x.get("last_updated_at")) for x in list(lines) + list(conts) + list(bundle["ships"])) if d]
    if upd and (today - max(upd)).days > stale:
        n = (today - max(upd)).days
        alerts.append(_china_alert(1, "🟠", "ملف الصين قديم",
            "اطلب من الصين نسخة محدّثة وارفعها.", [{"آخر تحديث في الملف": F(max(upd)), "من (يوم)": n}]))
    alerts.sort(key=lambda a: a["sev"])
    return alerts, on_track, no_stock, len(prod)

def build_inv_map(env, excluded_wh: set):
    inv_data = env.get_cached(env.inventory_sheet)
    inv_map = {}
    if len(inv_data) <= 1:
        return inv_map
    for r in inv_data[1:]:
        while len(r) < 6: r.append("")
        sku, wh, stock_raw, sales_raw, img, date_up = r[0].strip(), r[1].strip(), r[2], r[3], r[4], r[5]
        if not sku:
            continue
        sku_up = sku.upper()
        wh_up  = wh.upper()
        stock  = _to_int(stock_raw)
        sales  = _to_int(sales_raw)
        if sku_up not in inv_map:
            inv_map[sku_up] = {"sku":sku,"img":img,"date":date_up,"sales":sales,"warehouses":{},"total_stock":0}
        inv_map[sku_up]["warehouses"][wh] = inv_map[sku_up]["warehouses"].get(wh,0) + stock
        if wh_up not in excluded_wh:
            inv_map[sku_up]["total_stock"] += stock
        if not inv_map[sku_up]["img"] and img:
            inv_map[sku_up]["img"] = img
    return inv_map

def build_amazon_stock_map(env):
    """يبني خريطة مخزون أمازون من شيت StockAmazon (تصدير Seller Central):
    MSKU (seller-sku) -> ASIN + FNSKU + الكمية المتاحة. ده مصدر مخزون أمازون
    المستقل تمامًا عن inv_map بتاع نون — تاب مبيعات امازون بيبني قايمة
    منتجاته من هنا عشان تبقى قايمة منتجات أمازون الحقيقية مش قايمة نون."""
    data = env.get_cached(env.stock_amazon_sheet)
    stock_map = {}
    if len(data) <= 1:
        return stock_map
    hdr = data[0]
    idx_sku   = _find_col_idx_amz(hdr, ["seller-sku", "msku"])
    idx_asin  = _find_col_idx_amz(hdr, ["asin"])
    idx_fnsku = _find_col_idx_amz(hdr, ["fulfillment-channel-sku", "fnsku"])
    idx_qty   = _find_col_idx_amz(hdr, ["quantity available", "quantity"])
    if idx_sku is None:
        return stock_map
    known_idx = [i for i in (idx_sku, idx_asin, idx_fnsku, idx_qty) if i is not None]
    need = (max(known_idx) + 1) if known_idx else 1
    for row in data[1:]:
        if len(row) < need:
            row = row + [""] * (need - len(row))
        sku = row[idx_sku].strip()
        if not sku:
            continue
        sku_up = sku.upper()
        asin  = row[idx_asin].strip()  if idx_asin  is not None and len(row) > idx_asin  else ""
        fnsku = row[idx_fnsku].strip() if idx_fnsku is not None and len(row) > idx_fnsku else ""
        qty_raw = row[idx_qty] if (idx_qty is not None and len(row) > idx_qty) else "0"
        try:
            qty = int(float(str(qty_raw).replace(",", "").strip() or 0))
        except Exception:
            qty = 0
        if sku_up not in stock_map:
            stock_map[sku_up] = {"sku": sku, "asin": asin, "fnsku": fnsku, "stock": 0}
        stock_map[sku_up]["stock"] += qty
        if not stock_map[sku_up]["asin"] and asin:
            stock_map[sku_up]["asin"] = asin
    return stock_map

def build_amazon_stock_map_by_asin(env):
    """يبني خريطة مخزون أمازون من شيت StockAmazon لكن مجمّعة على أساس ASIN
    (مش MSKU زي build_amazon_stock_map فوق) — ده المعرّف الحقيقي لمنتج
    أمازون، فلو نفس الـ ASIN ظاهر تحت أكتر من MSKU (FBA/FBM مختلفين مثلاً)
    مخزونهم بيتجمع مع بعض بدل ما يتقسّم على صفوف منفصلة. تاب مبيعات أمازون
    بيستخدم النسخة دي عشان الأرقام (مخزون/شهري/يومي/نفاد) تطلع صح | Amazon
    stock map from StockAmazon aggregated by ASIN (not MSKU like
    build_amazon_stock_map above) — ASIN is Amazon's real product identity,
    so if the same ASIN appears under more than one MSKU (e.g. different
    FBA/FBM listings) their stock is summed together instead of being split
    across separate rows. The Amazon Sales tabs use this version so the
    numbers (stock/monthly/daily/stockout) come out correct."""
    data = env.get_cached(env.stock_amazon_sheet)
    asin_map = {}
    if len(data) <= 1:
        return asin_map
    hdr = data[0]
    idx_sku   = _find_col_idx_amz(hdr, ["seller-sku", "msku"])
    idx_asin  = _find_col_idx_amz(hdr, ["asin"])
    idx_fnsku = _find_col_idx_amz(hdr, ["fulfillment-channel-sku", "fnsku"])
    idx_qty   = _find_col_idx_amz(hdr, ["quantity available", "quantity"])
    if idx_asin is None:
        return asin_map
    known_idx = [i for i in (idx_sku, idx_asin, idx_fnsku, idx_qty) if i is not None]
    need = (max(known_idx) + 1) if known_idx else 1
    for row in data[1:]:
        if len(row) < need:
            row = row + [""] * (need - len(row))
        asin = row[idx_asin].strip() if len(row) > idx_asin else ""
        if not asin:
            continue  # لازم ASIN عشان التجميع يبقى صح
        asin_up = asin.upper()
        sku   = row[idx_sku].strip()   if idx_sku   is not None and len(row) > idx_sku   else ""
        fnsku = row[idx_fnsku].strip() if idx_fnsku is not None and len(row) > idx_fnsku else ""
        qty_raw = row[idx_qty] if (idx_qty is not None and len(row) > idx_qty) else "0"
        try:
            qty = int(float(str(qty_raw).replace(",", "").strip() or 0))
        except Exception:
            qty = 0
        if asin_up not in asin_map:
            asin_map[asin_up] = {"asin": asin, "fnsku": fnsku, "stock": 0, "skus": []}
        asin_map[asin_up]["stock"] += qty
        if not asin_map[asin_up]["fnsku"] and fnsku:
            asin_map[asin_up]["fnsku"] = fnsku
        if sku and sku not in asin_map[asin_up]["skus"]:
            asin_map[asin_up]["skus"].append(sku)
    return asin_map

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def build_daily_orders_map(env, target_date):
    """يرجع dict: sku_upper -> عدد صفوف الأوردرز لليوم المحدد (= Sum(QTY) بافتراض كل صف = قطعة واحدة)."""
    data = env.get_cached(env.daily_orders_sheet)
    counts = {}
    if len(data) <= 1:
        return counts
    for row in data[1:]:
        while len(row) < 2: row.append("")
        sku, ts = row[0].strip(), row[1].strip()
        if not sku or not ts:
            continue
        d = parse_excel_date(ts)
        if d and d.date() == target_date:
            sku_up = sku.upper()
            counts[sku_up] = counts.get(sku_up, 0) + 1
    return counts

def build_daily_orders_counts(env, dates):
    """يرجع dict: sku_upper -> {date: عدد} لقائمة تواريخ محددة (مرور واحد على الشيت بدل تكرار لكل تاريخ)."""
    data = env.get_cached(env.daily_orders_sheet)
    dates_set = set(dates)
    counts = {}
    if len(data) <= 1:
        return counts
    for row in data[1:]:
        while len(row) < 2: row.append("")
        sku, ts = row[0].strip(), row[1].strip()
        if not sku or not ts:
            continue
        d = parse_excel_date(ts)
        if d and d.date() in dates_set:
            sku_up = sku.upper()
            if sku_up not in counts:
                counts[sku_up] = {dd: 0 for dd in dates}
            counts[sku_up][d.date()] += 1
    return counts

def build_daily_orders_prices(env, dates):
    """يرجع dict: sku_upper -> {date: [(qty, price), ...]} لعرض تفاصيل الأسعار اليومية."""
    data = env.get_cached(env.daily_orders_sheet)
    dates_set = set(dates)
    prices = {}
    if len(data) <= 1:
        return prices
    hdr = data[0] if data else []
    price_col_idx = None
    for ci, h in enumerate(hdr):
        if str(h).strip().lower() in ("price","base_price","سعر","السعر","price_egp","unit_price","sale_price","selling_price"):
            price_col_idx = ci; break
    qty_col_idx = None
    for ci, h in enumerate(hdr):
        if str(h).strip().lower() in ("quantity","qty","كمية","الكمية","count"):
            qty_col_idx = ci; break
    for row in data[1:]:
        while len(row) < 2: row.append("")
        sku, ts = row[0].strip(), row[1].strip()
        if not sku or not ts:
            continue
        d = parse_excel_date(ts)
        if d and d.date() in dates_set:
            sku_up = sku.upper()
            price_val = ""
            if price_col_idx is not None and len(row) > price_col_idx:
                raw = str(row[price_col_idx]).strip()
                price_val = raw if raw and raw.lower() not in ("nan","none","") else ""
            qty_val = 1
            if qty_col_idx is not None and len(row) > qty_col_idx:
                try:
                    qty_val = int(float(str(row[qty_col_idx]).strip()))
                except Exception:
                    qty_val = 1
            if qty_val < 1:
                qty_val = 1
            if sku_up not in prices:
                prices[sku_up] = {dd: [] for dd in dates}
            prices[sku_up][d.date()].append((price_val, qty_val))
    return prices

def build_daily_orders_counts_fbb(env, dates):
    """نفس منطق build_daily_orders_counts بالظبط، لكن بيقتصر بس على صفوف FBB
    (Fulfillment Model = Fulfilled by Partner / FBP فقط) | Same logic as
    build_daily_orders_counts exactly, restricted to FBB rows
    (Fulfillment Model = FBP) only."""
    data = env.get_cached(env.daily_orders_sheet)
    dates_set = set(dates)
    counts = {}
    if len(data) <= 1:
        return counts
    hdr = data[0] if data else []
    fulfillment_col_idx = _find_fulfillment_col_idx(hdr)
    status_col_idx = _find_status_col_idx(hdr)
    for row in data[1:]:
        while len(row) < 2: row.append("")
        sku, ts = row[0].strip(), row[1].strip()
        if not sku or not ts:
            continue
        if not _row_is_fbb(row, fulfillment_col_idx, status_col_idx):
            continue
        d = parse_excel_date(ts)
        if d and d.date() in dates_set:
            sku_up = sku.upper()
            if sku_up not in counts:
                counts[sku_up] = {dd: 0 for dd in dates}
            counts[sku_up][d.date()] += 1
    return counts

def build_daily_orders_prices_fbb(env, dates):
    """نفس منطق build_daily_orders_prices بالظبط، لكن بيقتصر بس على صفوف FBB
    (Fulfillment Model = Fulfilled by Partner / FBP فقط) | Same logic as
    build_daily_orders_prices exactly, restricted to FBB rows
    (Fulfillment Model = FBP) only."""
    data = env.get_cached(env.daily_orders_sheet)
    dates_set = set(dates)
    prices = {}
    if len(data) <= 1:
        return prices
    hdr = data[0] if data else []
    fulfillment_col_idx = _find_fulfillment_col_idx(hdr)
    status_col_idx = _find_status_col_idx(hdr)
    price_col_idx = None
    for ci, h in enumerate(hdr):
        if str(h).strip().lower() in ("price","base_price","سعر","السعر","price_egp","unit_price","sale_price","selling_price"):
            price_col_idx = ci; break
    qty_col_idx = None
    for ci, h in enumerate(hdr):
        if str(h).strip().lower() in ("quantity","qty","كمية","الكمية","count"):
            qty_col_idx = ci; break
    for row in data[1:]:
        while len(row) < 2: row.append("")
        sku, ts = row[0].strip(), row[1].strip()
        if not sku or not ts:
            continue
        if not _row_is_fbb(row, fulfillment_col_idx, status_col_idx):
            continue
        d = parse_excel_date(ts)
        if d and d.date() in dates_set:
            sku_up = sku.upper()
            price_val = ""
            if price_col_idx is not None and len(row) > price_col_idx:
                raw = str(row[price_col_idx]).strip()
                price_val = raw if raw and raw.lower() not in ("nan","none","") else ""
            qty_val = 1
            if qty_col_idx is not None and len(row) > qty_col_idx:
                try:
                    qty_val = int(float(str(row[qty_col_idx]).strip()))
                except Exception:
                    qty_val = 1
            if qty_val < 1:
                qty_val = 1
            if sku_up not in prices:
                prices[sku_up] = {dd: [] for dd in dates}
            prices[sku_up][d.date()].append((price_val, qty_val))
    return prices

def build_amazon_orders_asin_map(env):
    """يبني خريطة MSKU -> ASIN من شيت DailyOrdersAmazon نفسه (عمود ASIN موجود
    جنب كل طلب). ده fallback لما الـ MSKU مش موجود في StockAmazon (يعني
    ملوش ASIN من خريطة المخزون) — عشان الـ ASIN يفضل ظاهر حتى للـ SKUs
    اللي "مش موجودة في المخزون المرفوع" | Builds an MSKU -> ASIN map straight
    from the DailyOrdersAmazon sheet (which has its own ASIN column per
    order) — a fallback for MSKUs missing from StockAmazon, so ASIN still
    shows even for "not in uploaded inventory" rows."""
    data = env.get_cached(env.daily_orders_amazon_sheet)
    asin_map = {}
    if len(data) <= 1:
        return asin_map
    hdr = data[0]
    idx = _amz_col_indexes(hdr)
    if idx["sku"] is None or idx["asin"] is None:
        return asin_map
    need = max(idx["sku"], idx["asin"]) + 1
    for row in data[1:]:
        if len(row) < need:
            row = row + [""] * (need - len(row))
        sku = row[idx["sku"]].strip()
        asin = row[idx["asin"]].strip()
        if not sku or not asin:
            continue
        sku_up = sku.upper()
        if sku_up not in asin_map:
            asin_map[sku_up] = asin
    return asin_map

def build_daily_orders_counts_amazon(env, dates, mode="all", group_by="sku"):
    """نفس منطق build_daily_orders_counts بالظبط، لكن من شيت DailyOrdersAmazon
    وبأعمدة بتتلاقى بالاسم، مع فلترة اختيارية على عمود حاوية كاملة الحمولة.
    group_by='sku' (زي ما كانت دايمًا) أو 'asin' — التجميع بيبقى على عمود
    ASIN في الشيت بدل MSKU، عشان صفوف تاب مبيعات أمازون تتجمع صح | group_by
    controls the aggregation key: 'sku' (unchanged, default) or 'asin' —
    groups by the sheet's own ASIN column instead of MSKU, so the Amazon
    Sales tab rows aggregate correctly by ASIN."""
    data = env.get_cached(env.daily_orders_amazon_sheet)
    dates_set = set(dates)
    counts = {}
    if len(data) <= 1:
        return counts
    hdr = data[0]
    idx = _amz_col_indexes(hdr)
    if idx["sku"] is None or idx["ts"] is None:
        return counts
    if group_by == "asin" and idx["asin"] is None:
        return counts
    need = max(idx["sku"], idx["ts"], idx["asin"] if idx["asin"] is not None else 0) + 1
    for row in data[1:]:
        if len(row) < need:
            row = row + [""] * (need - len(row))
        sku = row[idx["sku"]].strip()
        ts  = row[idx["ts"]].strip()
        if not sku or not ts:
            continue
        if not _amz_row_matches(row, idx["cont"], mode):
            continue
        if group_by == "asin":
            key_raw = row[idx["asin"]].strip() if len(row) > idx["asin"] else ""
            if not key_raw:
                continue
        else:
            key_raw = sku
        d = parse_excel_date(ts)
        if d and d.date() in dates_set:
            key_up = key_raw.upper()
            if key_up not in counts:
                counts[key_up] = {dd: 0 for dd in dates}
            counts[key_up][d.date()] += 1
    return counts

def build_daily_orders_prices_amazon(env, dates, mode="all", group_by="sku"):
    """نفس منطق build_daily_orders_prices بالظبط، لكن من شيت DailyOrdersAmazon.
    عمود "مبلغ المنتج" بيتاخد زي ما هو (سعر الحبة الواحدة) من غير أي قسمة على
    الكمية — بالظبط زي ما بيحصل مع سعر مبيعات نون فوق. group_by زي الدالة
    اللي فوق بالظبط | Same idea as the function above; group_by works the
    same way ('sku' default or 'asin')."""
    data = env.get_cached(env.daily_orders_amazon_sheet)
    dates_set = set(dates)
    prices = {}
    if len(data) <= 1:
        return prices
    hdr = data[0]
    idx = _amz_col_indexes(hdr)
    if idx["sku"] is None or idx["ts"] is None:
        return prices
    if group_by == "asin" and idx["asin"] is None:
        return prices
    need = max(idx["sku"], idx["ts"], idx["asin"] if idx["asin"] is not None else 0) + 1
    for row in data[1:]:
        if len(row) < need:
            row = row + [""] * (need - len(row))
        sku = row[idx["sku"]].strip()
        ts  = row[idx["ts"]].strip()
        if not sku or not ts:
            continue
        if not _amz_row_matches(row, idx["cont"], mode):
            continue
        if group_by == "asin":
            key_raw = row[idx["asin"]].strip() if len(row) > idx["asin"] else ""
            if not key_raw:
                continue
        else:
            key_raw = sku
        d = parse_excel_date(ts)
        if d and d.date() in dates_set:
            key_up = key_raw.upper()
            price_val = ""
            if idx["price"] is not None and len(row) > idx["price"]:
                raw = str(row[idx["price"]]).strip()
                price_val = raw if raw and raw.lower() not in ("nan","none","") else ""
            qty_val = 1
            if idx["qty"] is not None and len(row) > idx["qty"]:
                try:
                    qty_val = int(float(str(row[idx["qty"]]).strip()))
                except Exception:
                    qty_val = 1
            if qty_val < 1:
                qty_val = 1
            if key_up not in prices:
                prices[key_up] = {dd: [] for dd in dates}
            prices[key_up][d.date()].append((price_val, qty_val))
    return prices

def build_daily_orders_counts_amazon_all(env, dates):    return build_daily_orders_counts_amazon(env, dates, "all")

def build_daily_orders_counts_amazon_normal(env, dates): return build_daily_orders_counts_amazon(env, dates, "normal")

def build_daily_orders_counts_amazon_fba(env, dates):    return build_daily_orders_counts_amazon(env, dates, "fba")

def build_daily_orders_prices_amazon_all(env, dates):    return build_daily_orders_prices_amazon(env, dates, "all")

def build_daily_orders_prices_amazon_normal(env, dates): return build_daily_orders_prices_amazon(env, dates, "normal")

def build_daily_orders_prices_amazon_fba(env, dates):    return build_daily_orders_prices_amazon(env, dates, "fba")

def build_daily_orders_counts_amazon_all_asin(env, dates):    return build_daily_orders_counts_amazon(env, dates, "all", "asin")

def build_daily_orders_prices_amazon_all_asin(env, dates):    return build_daily_orders_prices_amazon(env, dates, "all", "asin")

def build_daily_orders_counts_amazon_normal_asin(env, dates): return build_daily_orders_counts_amazon(env, dates, "normal", "asin")

def build_daily_orders_counts_amazon_fba_asin(env, dates):    return build_daily_orders_counts_amazon(env, dates, "fba", "asin")

def build_daily_orders_prices_amazon_normal_asin(env, dates): return build_daily_orders_prices_amazon(env, dates, "normal", "asin")

def build_daily_orders_prices_amazon_fba_asin(env, dates):    return build_daily_orders_prices_amazon(env, dates, "fba", "asin")

def build_amazon_asin_sales_rows(env, multi_counts, prices_map, monthly_counts, sales_dates, sales_display_days):
    """يبني صفوف تاب مبيعات أمازون (عادي/FBA) مجمّعة على أساس ASIN — المعرّف
    الحقيقي لمنتج أمازون، مش MSKU. المخزون بييجي من amz_inv_map_asin (مجمّع
    من كل الـ MSKUs بتاعت نفس الـ ASIN)، والمبيعات اليومية من multi_counts/
    prices_map (مجمّعة بنفس الطريقة من عمود ASIN في DailyOrdersAmazon).
    "شهري" بيتحسب من monthly_counts وهو مجموع حقيقي لآخر 30 يوم فعلي من
    DailyOrdersAmazon — مش تقدير مبني على متوسط الأيام المعروضة (7 أيام
    افتراضيًا)، عشان منتج بطيء الحركة ما بعش خلال آخر 7 أيام بس باع قبل كده
    في نفس الشهر يفضل شهريه صحيح مش صفر. أي ASIN باع خلال آخر 30 يوم بيظهر
    حتى لو مالوش صف مخزون مرفوع في StockAmazon (يبقى مخزونه 0 مع علامة
    توضيحية) — بالظبط زي تابات نون | Builds Amazon sales-tab rows
    (normal/FBA) aggregated by ASIN — Amazon's real product identity, not
    MSKU. Stock comes from amz_inv_map_asin (summed across every MSKU
    sharing that ASIN); daily figures come from multi_counts/prices_map
    (grouped the same way, by DailyOrdersAmazon's own ASIN column).
    "Monthly" is computed from monthly_counts — an actual trailing-30-day
    sum from DailyOrdersAmazon, not an estimate extrapolated from the
    (default 7-day) display window, so a slow-moving product with no sales
    in the last 7 days but sales earlier in the month still shows its real
    monthly total instead of 0. Any ASIN with sales in the last 30 days
    shows up even without an uploaded StockAmazon row for it (stock shown
    as 0 with a note) — same as the Noon tabs' behaviour."""
    links_map_asin = env.get_links_map_asin()
    links_map_sku  = env.get_links_map()
    # MSKU -> ASIN من DailyOrdersAmazon نفسه — بنعكسها لـ ASIN -> MSKUs عشان
    # صفوف الـ ASIN اليتيمة (اللي مالهاش صف في StockAmazon) تفضل عارضة الـ
    # MSKU الحقيقي بتاعها (لشارات المستودع/الجدولة/الإعلانات...إلخ) | MSKU ->
    # ASIN from DailyOrdersAmazon itself — inverted to ASIN -> MSKUs so
    # orphan ASIN rows (no StockAmazon row) still carry their real MSKU for
    # the warehouse/schedule/ads... badges below.
    asin_to_mskus = {}
    for sku_o, asin_o in build_amazon_orders_asin_map(env).items():
        asin_to_mskus.setdefault(asin_o.upper(), []).append(sku_o)

    def _monthly_total(asin_up):
        mc = monthly_counts.get(asin_up)
        return sum(mc.values()) if mc else 0

    rows = []
    for asin_up, info in env.amz_inv_map_asin.items():
        stock = info.get("stock", 0)
        mskus = list(info.get("skus", []))
        for extra_sku in asin_to_mskus.get(asin_up, []):
            if extra_sku not in mskus:
                mskus.append(extra_sku)
        sku_disp    = ", ".join(mskus) if mskus else asin_up
        primary_sku = mskus[0] if mskus else asin_up
        # الصور على أساس الـ SKU/MSKU الأول (زي نون)، وبعدين صورة الـ ASIN كبديل
        img = next((links_map_sku.get(m.upper(), "") for m in mskus if links_map_sku.get(m.upper(), "")), "") or links_map_asin.get(asin_up, "")
        day_counts = multi_counts.get(asin_up, {d: 0 for d in sales_dates})
        day_prices = prices_map.get(asin_up, {d: [] for d in sales_dates})
        total_recent = sum(day_counts.get(d, 0) for d in sales_dates)
        avg_daily = (total_recent / sales_display_days) if sales_display_days > 0 else 0
        effective_avg = avg_daily
        days_to_stockout = round(stock / effective_avg) if effective_avg > 0 else 9999
        rows.append({
            "sku": sku_disp, "sku_up": primary_sku.upper(), "asin": info.get("asin", asin_up),
            "mskus": mskus,
            "stock": stock, "sales_month": _monthly_total(asin_up), "img": img,
            "day_counts": day_counts, "day_prices": day_prices,
            "total_recent": total_recent,
            "effective_avg": effective_avg,
            "days_to_stockout": days_to_stockout,
        })

    # ── ASINs باعت خلال آخر 30 يوم (يومي أو شهري) لكن مالهاش صف مخزون مرفوع
    #    في StockAmazon — بتتضاف كمان عشان كل مبيعة تظهر في القائمة، حتى لو
    #    آخر بيعة ليها كانت قبل الـ 7 أيام المعروضة | ASINs with sales in
    #    the last 30 days (recent or monthly) but no uploaded StockAmazon
    #    row — added too so every sale shows up, even if its last sale was
    #    before the displayed 7-day window.
    asins_seen = set(env.amz_inv_map_asin.keys())
    orphan_asins = (set(multi_counts.keys()) | set(monthly_counts.keys())) - asins_seen
    for asin_up in orphan_asins:
        day_counts = multi_counts.get(asin_up, {d: 0 for d in sales_dates})
        total_recent_orphan = sum(day_counts.get(d, 0) for d in sales_dates)
        monthly_orphan = _monthly_total(asin_up)
        if total_recent_orphan <= 0 and monthly_orphan <= 0:
            continue
        mskus = asin_to_mskus.get(asin_up, [])
        sku_disp    = ", ".join(mskus) if mskus else asin_up
        primary_sku = mskus[0] if mskus else asin_up
        day_prices_orphan = prices_map.get(asin_up, {d: [] for d in sales_dates})
        avg_daily_orphan = (total_recent_orphan / sales_display_days) if sales_display_days > 0 else 0
        rows.append({
            "sku": sku_disp, "sku_up": primary_sku.upper(), "asin": asin_up,
            "mskus": mskus,
            "stock": 0, "sales_month": monthly_orphan,
            "img": next((links_map_sku.get(m.upper(), "") for m in mskus if links_map_sku.get(m.upper(), "")), "") or links_map_asin.get(asin_up, ""),
            "day_counts": day_counts, "day_prices": day_prices_orphan,
            "total_recent": total_recent_orphan,
            "effective_avg": avg_daily_orphan,
            "days_to_stockout": 0 if avg_daily_orphan > 0 else 9999,
            "not_in_inventory": True,
        })

    # ترتيب: الأكتر مبيعاً أمس أولاً — نفس ترتيب تابات نون بالظبط
    rows.sort(key=lambda r: -r["day_counts"].get(sales_dates[0], 0) if sales_dates else 0)
    return rows

def build_daily_orders_family_stats_amazon_all(dates, live_map=None):    return {}

def build_daily_orders_family_stats_amazon_normal(dates, live_map=None): return {}

def build_daily_orders_family_stats_amazon_fba(dates, live_map=None):    return {}

def _build_trendyol_bundle(raw):
    """يحوّل الصفوف الخام لقائمة طلبات جاهزة (تواريخ متحوّلة مرة واحدة) + خرائط العمولة والصور والأسماء."""
    out = {"rows": [], "comm": {}, "img": {}, "name": {}}
    if not raw or len(raw) < 2:
        return out
    idx = {str(h).strip().lower(): i for i, h in enumerate(raw[0])}
    def g(row, name):
        i = idx.get(name)
        return str(row[i]).strip() if (i is not None and i < len(row)) else ""
    latest = {}
    for row in raw[1:]:
        bc = g(row, "barcode").upper()
        d = parse_excel_date(g(row, "order date"))
        if not bc or not d:
            continue
        try:
            qty = max(int(float(g(row, "quantity") or 1)), 1)
        except Exception:
            qty = 1
        try:
            unit = float(g(row, "net unit price").replace(",", ""))
        except Exception:
            unit = None
        status = g(row, "order status")
        out["rows"].append({"bc": bc, "d": d.date(), "qty": qty, "unit": unit, "status": status,
                            "ok": status.lower() not in TY_EXCLUDED_STATUSES})
        cm = g(row, "commission rate")
        if cm and (bc not in latest or d.date() >= latest[bc]):
            try:
                out["comm"][bc] = float(cm.replace(",", ""))
                latest[bc] = d.date()
            except Exception:
                pass
        im = g(row, "image url")
        if im and bc not in out["img"]:
            out["img"][bc] = im
        nm = g(row, "product name")
        if nm and bc not in out["name"]:
            out["name"][bc] = nm
    return out

def get_trendyol_commission_map(env):
    """Barcode (upper) -> نسبة العمولة من آخر طلب | Barcode -> latest commission %."""
    return env.get_trendyol_bundle()["comm"]

def build_daily_orders_counts_trendyol(env, dates):
    """Barcode (upper) -> {تاريخ: كمية القطع} — بيستبعد المرتجع/ما اتسلمش/لسه مدفوعش/ملغي."""
    dates_set = set(dates)
    counts = {}
    for r in env.get_trendyol_bundle()["rows"]:
        if r["ok"] and r["d"] in dates_set:
            if r["bc"] not in counts:
                counts[r["bc"]] = {dd: 0 for dd in dates}
            counts[r["bc"]][r["d"]] += r["qty"]
    return counts

def build_daily_orders_prices_trendyol(env, dates):
    """Barcode (upper) -> {تاريخ: [(سعر الوحدة الصافي، كمية)]} — نفس شكل أسعار نون/أمازون."""
    dates_set = set(dates)
    prices = {}
    for r in env.get_trendyol_bundle()["rows"]:
        if r["ok"] and r["d"] in dates_set:
            if r["bc"] not in prices:
                prices[r["bc"]] = {dd: [] for dd in dates}
            p = ("%.2f" % r["unit"]).rstrip("0").rstrip(".") if r["unit"] is not None else ""
            prices[r["bc"]][r["d"]].append((p, r["qty"]))
    return prices

def build_trendyol_sales_rows(env, multi_counts, prices_map, monthly_counts, sales_dates, sales_display_days):
    """صفوف مبيعات تريندول مجمّعة على Barcode (= SKU). أي Barcode باع خلال آخر 30 يوم بيظهر.
    الصورة على أساس الـ SKU (links n) الأول، وبعدين صورة تريندول من الملف."""
    b = env.get_trendyol_bundle()
    links_sku = env.get_links_map()
    rows = []
    for bc in set(multi_counts) | set(monthly_counts):
        dc = multi_counts.get(bc, {d: 0 for d in sales_dates})
        total_recent = sum(dc.get(d, 0) for d in sales_dates)
        monthly = sum(monthly_counts.get(bc, {}).values())
        if total_recent <= 0 and monthly <= 0:
            continue
        rows.append({
            "sku": bc, "sku_up": bc, "asin": None, "barcode": bc,
            "name": b["name"].get(bc, ""),
            "stock": 0, "sales_month": monthly,
            "img": links_sku.get(bc, "") or b["img"].get(bc, ""),
            "day_counts": dc, "day_prices": prices_map.get(bc, {d: [] for d in sales_dates}),
            "total_recent": total_recent,
            "effective_avg": (total_recent / sales_display_days) if sales_display_days > 0 else 0,
            "days_to_stockout": 9999,
        })
    return rows

def _build_ty_ads(raw):
    camps = []
    for row in (raw or [])[1:]:
        row = list(row) + [""] * (len(TY_ADS_HEADERS) - len(row))
        if not str(row[0]).strip():
            continue
        camps.append({"i": len(camps), "name": row[0], "status": row[1], "start": row[2], "n_products": int(_ty_num(row[3])),
                      "ids": [x.strip() for x in str(row[4]).split(",") if x.strip()],
                      "daily_budget": _ty_num(row[5]), "cost": _ty_num(row[6]), "clicks": int(_ty_num(row[7])),
                      "views": int(_ty_num(row[8])), "sold": int(_ty_num(row[9])),
                      "sales": _ty_num(row[10]), "roas": _ty_num(row[11])})
    return camps

def compute_no_sale_rows_trendyol(env, days=7):
    today = datetime.now().date()
    recent = [today - timedelta(days=i) for i in range(1, days + 1)]
    hist = [today - timedelta(days=i) for i in range(1, 61)]
    c_hist = build_daily_orders_counts_trendyol(env, hist)
    imgs = env.get_trendyol_bundle()["img"]
    links = env.get_links_map()
    cmap = get_tacweed_map(env)
    wh = env.get_warehouse_stock_map()
    rows = []
    for bc, dc in c_hist.items():
        if sum(dc.values()) <= 0 or sum(dc.get(d, 0) for d in recent) > 0:
            continue
        try:
            stock = int(float(wh.get(cmap.get(bc, ""), 0) or 0))
        except Exception:
            stock = 0
        rows.append({"sku": bc, "sku_up": bc, "stock": stock,
                     "sales_month": sum(dc.get(d, 0) for d in hist[:30]),
                     "last": max(d for d, q in dc.items() if q > 0),
                     "img": links.get(bc) or imgs.get(bc, "")})
    rows.sort(key=lambda x: (-x["stock"], -x["sales_month"]))
    return rows

def build_daily_orders_family_stats_trendyol(dates, live_map=None):
    return {}

def family_display_name(raw_family):
    """يحوّل قيمة عمود Family الخام لاسم قسم عربي/إنجليزي — يجمع أقسام المنزل تحت 'الهوم' واحدة."""
    key = str(raw_family).strip().lower().replace(" ", "_")
    if not key or key in ("nan", "none"):
        return None
    if key in FAMILY_HOME_GROUP:
        return "🏠 الهوم | Home"
    if key in FAMILY_LABELS:
        return FAMILY_LABELS[key]
    return f"📦 {str(raw_family).strip().replace('_',' ').title()}"

def build_daily_orders_family_stats(env, dates, live_map=None):
    """يرجع dict: اسم القسم -> {"orders": عدد, "revenue": إيراد} لفترة تواريخ محددة.
    لو عمود Family مش موجود في الشيت، أو الصف مالوش قيمة Family، بيتجاهله بهدوء
    بدون ما يوقف الكود أو يأثر على أي تحليل تاني. الإيراد بيتحسب دايمًا على سعر البيع
    الأساسي من تاب LIVE (sale_price) لو متوفر لهذا الـ SKU، وبيرجع لسعر العرض المسجل
    مع الطلب نفسه بس لو مفيش سعر بيع أساسي مسجل خالص."""
    data = env.get_cached(env.daily_orders_sheet)
    dates_set = set(dates)
    stats = {}
    if len(data) <= 1:
        return stats
    hdr = data[0] if data else []
    family_col_idx = None
    for ci, h in enumerate(hdr):
        if str(h).strip().lower() in ("family", "القسم", "قسم", "department", "category"):
            family_col_idx = ci; break
    if family_col_idx is None:
        return stats
    price_col_idx = None
    for ci, h in enumerate(hdr):
        if str(h).strip().lower() in ("price","base_price","سعر","السعر","price_egp","unit_price","sale_price","selling_price"):
            price_col_idx = ci; break
    qty_col_idx = None
    for ci, h in enumerate(hdr):
        if str(h).strip().lower() in ("quantity","qty","كمية","الكمية","count"):
            qty_col_idx = ci; break
    for row in data[1:]:
        while len(row) <= family_col_idx:
            row.append("")
        sku = row[0].strip() if len(row) > 0 else ""
        ts  = row[1].strip() if len(row) > 1 else ""
        if not sku or not ts:
            continue
        d = parse_excel_date(ts)
        if not d or d.date() not in dates_set:
            continue
        fam_raw = row[family_col_idx].strip() if len(row) > family_col_idx else ""
        if not fam_raw or fam_raw.lower() in ("nan", "none"):
            continue  # لا يوجد قسم لهذا الصف — يتجاهل من تحليل الأقسام فقط
        dept_name = family_display_name(fam_raw)
        if not dept_name:
            continue
        price_val = ""
        if price_col_idx is not None and len(row) > price_col_idx:
            price_val = str(row[price_col_idx]).strip()
        qty_val = 1
        if qty_col_idx is not None and len(row) > qty_col_idx:
            try:
                qty_val = int(float(str(row[qty_col_idx]).strip()))
            except Exception:
                qty_val = 1
        if qty_val < 1:
            qty_val = 1
        rev_val = 0.0
        live_info_fam = live_map.get(sku.upper()) if live_map else None
        live_price_fam = live_info_fam.get("price") if live_info_fam else None
        if live_price_fam is not None:
            rev_val = live_price_fam * qty_val
        elif price_val and price_val.lower() not in ("", "nan", "none"):
            try:
                rev_val = float(price_val.replace(",", "")) * qty_val
            except Exception:
                rev_val = 0.0
        if dept_name not in stats:
            stats[dept_name] = {"orders": 0, "revenue": 0.0}
        stats[dept_name]["orders"]  += 1
        stats[dept_name]["revenue"] += rev_val
    return stats

def build_daily_orders_family_stats_fbb(env, dates, live_map=None):
    """نفس منطق build_daily_orders_family_stats بالظبط، لكن بيقتصر بس على صفوف FBB
    (Fulfillment Model = Fulfilled by Partner / FBP فقط) —
    نسخة منفصلة تمامًا زي build_daily_orders_counts_fbb/build_daily_orders_prices_fbb
    فوق، من غير أي تعديل على الدالة الأصلية | Same logic as
    build_daily_orders_family_stats exactly, restricted to FBB rows only. A fully
    separate copy — mirrors build_daily_orders_counts_fbb / build_daily_orders_prices_fbb
    above — the original function is left completely untouched."""
    data = env.get_cached(env.daily_orders_sheet)
    dates_set = set(dates)
    stats = {}
    if len(data) <= 1:
        return stats
    hdr = data[0] if data else []
    fulfillment_col_idx = _find_fulfillment_col_idx(hdr)
    status_col_idx = _find_status_col_idx(hdr)
    family_col_idx = None
    for ci, h in enumerate(hdr):
        if str(h).strip().lower() in ("family", "القسم", "قسم", "department", "category"):
            family_col_idx = ci; break
    if family_col_idx is None:
        return stats
    price_col_idx = None
    for ci, h in enumerate(hdr):
        if str(h).strip().lower() in ("price","base_price","سعر","السعر","price_egp","unit_price","sale_price","selling_price"):
            price_col_idx = ci; break
    qty_col_idx = None
    for ci, h in enumerate(hdr):
        if str(h).strip().lower() in ("quantity","qty","كمية","الكمية","count"):
            qty_col_idx = ci; break
    for row in data[1:]:
        while len(row) <= family_col_idx:
            row.append("")
        sku = row[0].strip() if len(row) > 0 else ""
        ts  = row[1].strip() if len(row) > 1 else ""
        if not sku or not ts:
            continue
        if not _row_is_fbb(row, fulfillment_col_idx, status_col_idx):
            continue
        d = parse_excel_date(ts)
        if not d or d.date() not in dates_set:
            continue
        fam_raw = row[family_col_idx].strip() if len(row) > family_col_idx else ""
        if not fam_raw or fam_raw.lower() in ("nan", "none"):
            continue
        dept_name = family_display_name(fam_raw)
        if not dept_name:
            continue
        price_val = ""
        if price_col_idx is not None and len(row) > price_col_idx:
            price_val = str(row[price_col_idx]).strip()
        qty_val = 1
        if qty_col_idx is not None and len(row) > qty_col_idx:
            try:
                qty_val = int(float(str(row[qty_col_idx]).strip()))
            except Exception:
                qty_val = 1
        if qty_val < 1:
            qty_val = 1
        rev_val = 0.0
        live_info_fam = live_map.get(sku.upper()) if live_map else None
        live_price_fam = live_info_fam.get("price") if live_info_fam else None
        if live_price_fam is not None:
            rev_val = live_price_fam * qty_val
        elif price_val and price_val.lower() not in ("", "nan", "none"):
            try:
                rev_val = float(price_val.replace(",", "")) * qty_val
            except Exception:
                rev_val = 0.0
        if dept_name not in stats:
            stats[dept_name] = {"orders": 0, "revenue": 0.0}
        stats[dept_name]["orders"]  += 1
        stats[dept_name]["revenue"] += rev_val
    return stats

def build_daily_orders_counts_fbn(env, dates):
    """نفس منطق build_daily_orders_counts بالظبط، لكن بيستبعد صفوف FBB (عكس
    build_daily_orders_counts_fbb) | Same logic as build_daily_orders_counts
    exactly, but excludes FBB rows (inverse of build_daily_orders_counts_fbb)."""
    data = env.get_cached(env.daily_orders_sheet)
    dates_set = set(dates)
    counts = {}
    if len(data) <= 1:
        return counts
    hdr = data[0] if data else []
    fulfillment_col_idx = _find_fulfillment_col_idx(hdr)
    status_col_idx = _find_status_col_idx(hdr)
    for row in data[1:]:
        while len(row) < 2: row.append("")
        sku, ts = row[0].strip(), row[1].strip()
        if not sku or not ts:
            continue
        if _row_is_fbb(row, fulfillment_col_idx, status_col_idx):
            continue
        d = parse_excel_date(ts)
        if d and d.date() in dates_set:
            sku_up = sku.upper()
            if sku_up not in counts:
                counts[sku_up] = {dd: 0 for dd in dates}
            counts[sku_up][d.date()] += 1
    return counts

def build_daily_orders_prices_fbn(env, dates):
    """نفس منطق build_daily_orders_prices بالظبط، لكن بيستبعد صفوف FBB (عكس
    build_daily_orders_prices_fbb) | Same logic as build_daily_orders_prices
    exactly, but excludes FBB rows (inverse of build_daily_orders_prices_fbb)."""
    data = env.get_cached(env.daily_orders_sheet)
    dates_set = set(dates)
    prices = {}
    if len(data) <= 1:
        return prices
    hdr = data[0] if data else []
    fulfillment_col_idx = _find_fulfillment_col_idx(hdr)
    status_col_idx = _find_status_col_idx(hdr)
    price_col_idx = None
    for ci, h in enumerate(hdr):
        if str(h).strip().lower() in ("price","base_price","سعر","السعر","price_egp","unit_price","sale_price","selling_price"):
            price_col_idx = ci; break
    qty_col_idx = None
    for ci, h in enumerate(hdr):
        if str(h).strip().lower() in ("quantity","qty","كمية","الكمية","count"):
            qty_col_idx = ci; break
    for row in data[1:]:
        while len(row) < 2: row.append("")
        sku, ts = row[0].strip(), row[1].strip()
        if not sku or not ts:
            continue
        if _row_is_fbb(row, fulfillment_col_idx, status_col_idx):
            continue
        d = parse_excel_date(ts)
        if d and d.date() in dates_set:
            sku_up = sku.upper()
            price_val = ""
            if price_col_idx is not None and len(row) > price_col_idx:
                raw = str(row[price_col_idx]).strip()
                price_val = raw if raw and raw.lower() not in ("nan","none","") else ""
            qty_val = 1
            if qty_col_idx is not None and len(row) > qty_col_idx:
                try:
                    qty_val = int(float(str(row[qty_col_idx]).strip()))
                except Exception:
                    qty_val = 1
            if qty_val < 1:
                qty_val = 1
            if sku_up not in prices:
                prices[sku_up] = {dd: [] for dd in dates}
            prices[sku_up][d.date()].append((price_val, qty_val))
    return prices

def build_daily_orders_family_stats_fbn(env, dates, live_map=None):
    """نفس منطق build_daily_orders_family_stats بالظبط، لكن بيستبعد صفوف FBB
    (عكس build_daily_orders_family_stats_fbb) | Same logic as
    build_daily_orders_family_stats exactly, but excludes FBB rows (inverse of
    build_daily_orders_family_stats_fbb)."""
    data = env.get_cached(env.daily_orders_sheet)
    dates_set = set(dates)
    stats = {}
    if len(data) <= 1:
        return stats
    hdr = data[0] if data else []
    fulfillment_col_idx = _find_fulfillment_col_idx(hdr)
    status_col_idx = _find_status_col_idx(hdr)
    family_col_idx = None
    for ci, h in enumerate(hdr):
        if str(h).strip().lower() in ("family", "القسم", "قسم", "department", "category"):
            family_col_idx = ci; break
    if family_col_idx is None:
        return stats
    price_col_idx = None
    for ci, h in enumerate(hdr):
        if str(h).strip().lower() in ("price","base_price","سعر","السعر","price_egp","unit_price","sale_price","selling_price"):
            price_col_idx = ci; break
    qty_col_idx = None
    for ci, h in enumerate(hdr):
        if str(h).strip().lower() in ("quantity","qty","كمية","الكمية","count"):
            qty_col_idx = ci; break
    for row in data[1:]:
        while len(row) <= family_col_idx:
            row.append("")
        sku = row[0].strip() if len(row) > 0 else ""
        ts  = row[1].strip() if len(row) > 1 else ""
        if not sku or not ts:
            continue
        if _row_is_fbb(row, fulfillment_col_idx, status_col_idx):
            continue
        d = parse_excel_date(ts)
        if not d or d.date() not in dates_set:
            continue
        fam_raw = row[family_col_idx].strip() if len(row) > family_col_idx else ""
        if not fam_raw or fam_raw.lower() in ("nan", "none"):
            continue
        dept_name = family_display_name(fam_raw)
        if not dept_name:
            continue
        price_val = ""
        if price_col_idx is not None and len(row) > price_col_idx:
            price_val = str(row[price_col_idx]).strip()
        qty_val = 1
        if qty_col_idx is not None and len(row) > qty_col_idx:
            try:
                qty_val = int(float(str(row[qty_col_idx]).strip()))
            except Exception:
                qty_val = 1
        if qty_val < 1:
            qty_val = 1
        rev_val = 0.0
        live_info_fam = live_map.get(sku.upper()) if live_map else None
        live_price_fam = live_info_fam.get("price") if live_info_fam else None
        if live_price_fam is not None:
            rev_val = live_price_fam * qty_val
        elif price_val and price_val.lower() not in ("", "nan", "none"):
            try:
                rev_val = float(price_val.replace(",", "")) * qty_val
            except Exception:
                rev_val = 0.0
        if dept_name not in stats:
            stats[dept_name] = {"orders": 0, "revenue": 0.0}
        stats[dept_name]["orders"]  += 1
        stats[dept_name]["revenue"] += rev_val
    return stats

def is_sku_only_in_excluded_warehouses(env, sku_up, excluded_wh):
    """يتحقق هل كل مستودعات هذا الـ SKU (في ملف المخزون) هي مستودعات مستثناة فقط —
    لو كذلك، يبقى مينفعش يظهر في تابي مراجعة المخزون / مراجعة المبيعات."""
    if not excluded_wh:
        return False
    info = env.inv_map.get(sku_up)
    if not info:
        return False
    whs = info.get("warehouses", {})
    if not whs:
        return False
    return all(w.strip().upper() in excluded_wh for w in whs.keys())

def compute_stock_sales_rows(env, target_date, display_dates=None):
    """يحسب لكل SKU ظهر في أوردرز اليوم المحدد نفس مخرجات استعلامي مراجعة مخزون / مراجعة مبيعات.
    display_dates (اختياري): قائمة تواريخ إضافية تتعرض جنب كل SKU (مثلاً أمس/أول أمس/أول أول أمس)."""
    daily_qty = build_daily_orders_map(env, target_date)
    display_dates = display_dates or [target_date]
    multi_counts = build_daily_orders_counts(env, display_dates)
    rows = []
    for sku_up, qty in daily_qty.items():
        if is_sku_only_in_excluded_warehouses(env, sku_up, env.excluded_wh):
            continue
        info        = env.inv_map.get(sku_up, {})
        stock       = info.get("total_stock", 0)
        sales_month = info.get("sales", 0)
        sku_disp    = info.get("sku", sku_up)
        img         = info.get("img", "")
        threshold_10d = sales_month/30*10
        stock_alert = (stock - threshold_10d) < 0
        day_counts = multi_counts.get(sku_up, {dd:0 for dd in display_dates})

        # ══ مبيعات أعلى من المعتاد — بمتوسط آخر 3 أيام (مش يوم واحد بس) عشان نقلل
        #    الـ noise من طلبية كبيرة عشوائية في يوم واحد، وبشرط إن الارتفاع
        #    يكون مستمر يومين على الأقل من آخر 3 أيام (مش مجرد يوم شاذ) ══
        recent_days      = display_dates[:3] if len(display_dates) >= 3 else display_dates
        recent_vals      = [day_counts.get(dd, 0) for dd in recent_days]
        recent_avg       = (sum(recent_vals) / len(recent_vals)) if recent_vals else 0
        daily_avg_normal = (sales_month / 30) if sales_month > 0 else 0
        elevated_days    = sum(1 for v in recent_vals if daily_avg_normal > 0 and v > daily_avg_normal)
        sales_alert = (
            sales_month > 0
            and recent_avg * 30 > sales_month
            and elevated_days >= 2
        )

        suggested_qty = round(sales_month/30*18) if stock_alert else 0
        days_to_stockout       = round(stock/(sales_month/30)) if sales_month > 0 else 0
        days_to_stockout_today = round(stock/abs(qty)) if abs(qty) > 0 else 0
        rows.append({
            "sku": sku_disp, "sku_up": sku_up, "qty": qty, "stock": stock, "sales_month": sales_month,
            "img": img, "stock_alert": stock_alert, "sales_alert": sales_alert,
            "suggested_qty": suggested_qty, "days_to_stockout": days_to_stockout,
            "days_to_stockout_today": days_to_stockout_today,
            "day_counts": day_counts,
        })
    return rows

def compute_missing_inventory_rows(env, display_dates):
    """SKUs ظهرت في الأوردرز خلال آخر كذا يوم (أمس/أول أمس/أول أول أمس) لكن مالهاش سجل في شيت Inventory
    —  مخزونها انتهى بالكامل وخرجت من ملف المخزون. تظهر بنفس تفاصيل تابي المراجعة."""
    multi_counts = build_daily_orders_counts(env, display_dates)
    links_map_local = env.get_links_map()
    rows = []
    for sku_up, day_counts in multi_counts.items():
        if sku_up in env.inv_map:
            continue
        total_recent = sum(day_counts.values())
        if total_recent <= 0:
            continue
        # مفيش "مبيع شهري" رسمي ليها لأنها مش موجودة في ملف المخزون أصلاً —
        # بنحسب تقدير تقريبي بناءً على متوسط آخر الأيام المعروضة × 30
        est_monthly_sales = round((total_recent / len(display_dates)) * 30)
        rows.append({
            "sku": sku_up, "sku_up": sku_up,
            "img": links_map_local.get(sku_up, ""),
            "day_counts": day_counts,
            "total_recent": total_recent,
            "est_monthly_sales": est_monthly_sales,
        })
    rows.sort(key=lambda r: -r["total_recent"])
    return rows

def get_recent_expired_info(env, sku, days_back=4):
    """يدوّر على SKU في شيت Expired (الجدولة منتهية الصلاحية) خلال آخر days_back يوم — حسب تاريخ الانتهاء.
    يرجع أحدث سجل لو لقى، أو None لو مفيش جدولة منتهية مؤخراً لهذا الـ SKU."""
    sku_up = sku.strip().upper()
    data = env.get_cached(env.expired_sheet)
    if len(data) <= 1:
        return None
    cutoff = datetime.now().date() - timedelta(days=days_back)
    candidates = []
    for row in data[1:]:
        while len(row) < 7: row.append("")
        if row[1].strip().upper() != sku_up:
            continue
        d_exp = parse_excel_date(row[6])
        if d_exp and d_exp.date() >= cutoff:
            candidates.append({"asn": row[0], "schedule_date": row[3], "date_expired": row[6], "parsed_expired": d_exp})
    if not candidates:
        return None
    candidates.sort(key=lambda c: c["parsed_expired"], reverse=True)
    return candidates[0]

def get_recent_schedule_rows(env, days_back=4):
    """يرجع dict: sku_upper -> أحدث سجل جدولة (من Scheduled/Check/Expired) وقع تاريخ جدولته
    خلال آخر days_back يوم (يعني: أمس/أول أمس/أول أول أمس/قبل 4 أيام).
    الهدف: الـ SKU ده لسه في فترة انتظار وصول المخزون بعد الجدولة — حتى لو الجدولة
    خلاص اتنقلت لتاب Expired — فلازم يفضل ظاهر في تابات المراجعة عشان محدش يطلبه تاني بالغلط."""
    cutoff = datetime.now().date() - timedelta(days=days_back)
    today_ = datetime.now().date()
    src_label_map = {"Scheduled": "الجدولة | Scheduled", "Check": "تشييك | Check", "Expired": "منتهية | Expired"}
    by_sku = {}
    for sheet_key in ("Scheduled", "Check", "Expired"):
        data = get_scheduled_normalized(env) if sheet_key == "Scheduled" else env.get_cached(env.sheets[sheet_key])
        if len(data) <= 1:
            continue
        for row in data[1:]:
            while len(row) < 4:
                row.append("")
            sku_up = row[1].strip().upper()
            if not sku_up:
                continue
            d = parse_excel_date(row[3])  # عمود "Schedule Date" في الثلاث شيتات
            if not d:
                continue
            dd = d.date()
            if not (cutoff <= dd <= today_):
                continue
            entry = {
                "sku_up": sku_up, "asn": row[0], "date": row[3], "parsed": dd,
                "qty": row[2] if len(row) > 2 else "",
                "source": sheet_key, "source_label": src_label_map.get(sheet_key, sheet_key),
            }
            prev = by_sku.get(sku_up)
            if not prev or dd > prev["parsed"]:
                by_sku[sku_up] = entry
    return by_sku

def get_pending_approval_skus(env, force=False):
    """يرجع set بكل الـ SKUs الموجودة حالياً في تاب PendingApproval — يعني لسه مستنية اعتماد
    الجدولة من برنامج الجدولة المكتبي (schedule_entry_app) وما اتنقلتش لتاب Scheduled لسه.
    القراءة حسب اسم عمود SKU مش مكانه، عشان تفضل شغالة أياً كان ترتيب الأعمدة الحقيقي في الشيت."""
    raw = env.get_cached(env.pending_approval_sheet, force=force)
    if not raw or len(raw) <= 1:
        return set()
    header = raw[0]
    if "SKU" not in header:
        return set()
    sku_idx = header.index("SKU")
    out = set()
    for row in raw[1:]:
        if sku_idx < len(row):
            sku_up = row[sku_idx].strip().upper()
            if sku_up:
                out.add(sku_up)
    return out

def compute_recent_scheduled_rows(env, exclude_skus, day_dates, days_back=4):
    """يبني صفوف SKUs اتجدولت أو انتهت جدولتها خلال آخر days_back يوم ومش ظاهرة أصلاً
    في قائمة المراجعة الرئيسية (exclude_skus) — عشان تتعرض في سكشن منفصل يفكّر المستخدم
    إنها كانت مجدولة مؤخراً ولسه بتستنى توصل."""
    recent_map = get_recent_schedule_rows(env, days_back=days_back)
    if not recent_map:
        return []
    day_counts_map = build_daily_orders_counts(env, day_dates)
    rows = []
    for sku_up, sched_entry in recent_map.items():
        if sku_up in exclude_skus:
            continue
        info        = env.inv_map.get(sku_up, {})
        stock       = info.get("total_stock", 0)
        sales_month = info.get("sales", 0)
        img         = info.get("img", "")
        sku_disp    = info.get("sku", sku_up)
        day_counts  = day_counts_map.get(sku_up, {d: 0 for d in day_dates})
        rows.append({
            "sku": sku_disp, "sku_up": sku_up, "stock": stock, "sales_month": sales_month,
            "img": img, "day_counts": day_counts, "sched": sched_entry,
        })
    rows.sort(key=lambda r: -r["sched"]["parsed"].toordinal())
    return rows

def get_latest_schedule_info(env, sku):
    """يدوّر على SKU في الجدولة والتشييك ويرجع أقرب جدولة (تاريخ) أو None."""
    sku_up = sku.strip().upper()
    candidates = []
    for sheet_key in ("Scheduled","Check"):
        data = get_scheduled_normalized(env) if sheet_key == "Scheduled" else env.get_cached(env.sheets[sheet_key])
        if len(data) <= 1:
            continue
        for row in data[1:]:
            while len(row) < 4: row.append("")
            if row[1].strip().upper() == sku_up:
                d = parse_excel_date(row[3])
                candidates.append({"asn": row[0], "date": row[3], "qty": row[2], "parsed": d, "source": sheet_key})
    if not candidates:
        return None
    dated = [c for c in candidates if c["parsed"]]
    if dated:
        dated.sort(key=lambda c: c["parsed"])
        return dated[0]
    return candidates[0]

def is_sku_unavailable(env, sku_up):
    """True لو الـ SKU موجود حالياً في تاب Unavailable (يعني متسجل غير متوفر ولسه محدش وافق
    على طلب جديد له). | True if the SKU currently sits in the Unavailable sheet."""
    data_un = env.get_cached(env.unavailable_sheet)
    if len(data_un) > 1:
        for row in data_un[1:]:
            if row and row[0].strip().upper() == sku_up:
                return True
    return False

def get_unavailable_ordered_note(env, sku):
    """لو الـ SKU سبق اتسجل غير متوفر أو تم طلبه، يرجع ملاحظات بالتواريخ."""
    sku_up = sku.strip().upper()
    notes = []
    data_un = env.get_cached(env.unavailable_sheet)
    if len(data_un) > 1:
        for row in data_un[1:]:
            if row and row[0].strip().upper() == sku_up:
                while len(row) < 5: row.append("")
                cnt, dates = parse_count_dates(row[4])
                notes.append(f"❌ غير متوفر سابقاً | Was unavailable ({cnt}x) — {dates}")
                break
    data_ord = env.get_cached(env.ordered_sheet)
    if len(data_ord) > 1:
        for row in data_ord[1:]:
            if row and row[0].strip().upper() == sku_up:
                while len(row) < 6: row.append("")
                cnt, dates = parse_count_dates(row[5])
                notes.append(f"🛒 تم طلبه سابقاً | Was ordered ({cnt}x) — {dates}")
                break
    return notes

def schedule_coverage_badge(env, sku, days_to_stockout, delay_days):
    """يرجع (نص الحالة, لون, معلومات الجدولة) حسب هل الجدولة هتوصل قبل نفاد المخزون أو لأ."""
    sched = get_latest_schedule_info(env, sku)
    if not sched:
        return ("🔴 محتاج جدولة الآن | Needs scheduling now", "#ef4444", None)
    if not sched["parsed"]:
        return (f"⚠️ مجدول (ASN {sched['asn']}) بدون تاريخ واضح | Scheduled, unclear date", "#f59e0b", sched)
    arrival = sched["parsed"] + timedelta(days=delay_days)
    stockout_date = datetime.now() + timedelta(days=days_to_stockout) if days_to_stockout > 0 else datetime.now()
    src_label = "تشييك | Check" if sched["source"]=="Check" else "الجدولة | Scheduled"
    if arrival.date() <= stockout_date.date():
        return (f"✅ مجدول (ASN {sched['asn']}) بتاريخ {sched['date']} [{src_label}] — هيوصل قبل النفاد | Will arrive before stockout", "#22c55e", sched)
    else:
        return (f"🔴 مجدول (ASN {sched['asn']}) بتاريخ {sched['date']} [{src_label}] — لكن متأخر عن موعد النفاد | But too late before stockout", "#ef4444", sched)

def compute_warehouse_stockout_rows(env):
    """لكل كود عنده كمية مسجلة في WarehouseStock، بيحسب معدل السحب اليومي المجمّع من
    كل الـ SKUs التابعة لنفس الكود (فعليًا من شيت DailyOrders، مش من مبيع Inventory
    الشهري)، وبيرجع أيام النفاد المتوقعة وعدد الأيام المتبقية لموعد الطلب.

    3 إعدادات بتتحكم في الحساب، بتتضاف/تتعدّل يدوي في شيت Settings (نفس أسلوب
    schedule_delay_days و xdock_low_stock_threshold الموجودين):
    - warehouse_sales_window_days: عدد أيام نافذة حساب معدل السحب (افتراضي 30)
    - warehouse_lead_time_days: مدة التوريد المتوقعة (افتراضي 7)
    - warehouse_reorder_safety_days: هامش الأمان قبل موعد الطلب (افتراضي 3)
    """
    settings_wh = env.load_settings()
    window_days = int(settings_wh.get("warehouse_sales_window_days", "30") or 30)
    lead_time   = int(settings_wh.get("warehouse_lead_time_days", "7") or 7)
    safety_days = int(settings_wh.get("warehouse_reorder_safety_days", "3") or 3)

    dates = [(datetime.now() - timedelta(days=i)).date() for i in range(max(window_days, 1))]
    multi_counts_wh = build_daily_orders_counts(env, dates)
    reverse_map_wh  = get_tacweed_reverse_map(env)
    stock_map_wh    = env.get_warehouse_stock_map()
    links_map_wh    = env.get_links_map()

    # ── سحب أمازون من نفس مستودعنا (WarehouseStock) — بطلب صريح من المستخدم
    #    بيتحسب من كل طلبات أمازون (عادي FSAB + FBA مع بعض)، مش الـ FSAB بس |
    #    Amazon draw from OUR warehouse (WarehouseStock) — per explicit
    #    request, counts ALL Amazon orders (normal/FSAB + FBA together), not
    #    just FSAB ──
    multi_counts_amz_wh = build_daily_orders_counts_amazon(env, dates, mode="all", group_by="asin")
    reverse_map_amz_wh  = get_tacweed_amazon_reverse_map(env)
    # سحب تريندول من نفس المستودع — Barcode = SKU فبيتحسب على نفس SKUs الكود
    multi_counts_ty_wh  = build_daily_orders_counts_trendyol(env, dates)

    # شحنات الصين القادمة بتتحسب جوه نفس القرار (نفس محرك تاب الصين) عشان الرقمين ميتعارضوش
    try:
        _cn_by_code = env.get_china_bundle()["by_code"]
        _cn_cfg = env._china_cfg()
    except Exception:
        _cn_by_code, _cn_cfg = {}, None
    _today_wh = datetime.now().date()

    rows = []
    for code, qty_str in stock_map_wh.items():
        try:
            qty = float(str(qty_str).replace(",", "").strip())
        except (TypeError, ValueError):
            continue

        skus = reverse_map_wh.get(code, [])
        asins = reverse_map_amz_wh.get(code, [])
        total_sold_noon = sum(sum(multi_counts_wh.get(sku_up, {}).values()) for sku_up in skus)
        total_sold_amz  = sum(sum(multi_counts_amz_wh.get(asin_up, {}).values()) for asin_up in asins)
        total_sold_ty = sum(sum(multi_counts_ty_wh.get(sku_up, {}).values()) for sku_up in skus)
        total_sold = total_sold_noon + total_sold_amz + total_sold_ty
        daily_draw = (total_sold / window_days) if window_days > 0 else 0

        if daily_draw <= 0:
            days_to_stockout, reorder_in = None, None
            status, color = "⚪ لا يوجد سحب مسجل | No recorded draw", "#9ca3af"
        else:
            days_to_stockout = qty / daily_draw
            reorder_in = days_to_stockout - lead_time
            if reorder_in <= 0:
                status, color = "🚨 اطلب الآن | Order now", "#ef4444"
            elif reorder_in <= safety_days:
                status, color = "🟠 قرب موعد الطلب | Reorder soon", "#f59e0b"
            else:
                status, color = "🟢 آمن | Safe", "#22c55e"

        suggested_qty = round(daily_draw * (lead_time + safety_days)) if daily_draw > 0 else 0
        img = next((links_map_wh.get(s) for s in skus if links_map_wh.get(s)), "")

        # ── 🚢 أثر شحنات الصين: الكود اللي ليه شحنات قادمة بيتقيّم بالمخزون + الشحنات المؤرخة ومدة توريد الصين ──
        china = None
        _inc = [l for l in _cn_by_code.get(_china_norm_code(code), []) if l["_status"] == "incoming"] if _cn_cfg else []
        if _inc:
            _dated, _undated = china_batches_for_lines(_inc)
            china = {"incoming": sum(l["_qty"] for l in _inc), "dated_qty": sum(b["qty"] for b in _dated),
                     "undated_qty": sum(u["qty"] for u in _undated),
                     "next_date": _dated[0]["date"] if _dated else None, "next_qty": _dated[0]["qty"] if _dated else 0,
                     "state": None}
            if daily_draw > 0:
                _dec = china_decide(qty, daily_draw, [(b["date"], b["qty"]) for b in _dated],
                                    china["undated_qty"], _today_wh, _cn_cfg)
                china.update(state=_dec["state"], covered=_dec.get("covered"), order_by=_dec.get("order_by"),
                             gap_days=_dec.get("gap_days", 0), lost=_dec.get("lost", 0),
                             gap_from=_dec.get("gap_from"), gap_to=_dec.get("gap_to"))
                _st = _dec["state"]
                _ob = (_dec["order_by"] - _today_wh).days
                if _st in ("gap", "order_late"):
                    status, color = "🚨 اطلب الآن | Order now", "#ef4444"
                    reorder_in = _ob
                    suggested_qty = _dec["suggested"]
                elif _st in ("order_soon", "need_date"):
                    status, color = "🟠 قرب موعد الطلب | Reorder soon", "#f59e0b"
                    reorder_in = _ob
                    suggested_qty = _dec["suggested"]
                else:
                    status, color = "🟢 آمن | Safe", "#22c55e"
                    reorder_in = _ob
                    suggested_qty = 0

        rows.append({
            "code": code, "skus": skus, "asins": asins, "qty": qty,
            "daily_draw": round(daily_draw, 2),
            "draw_noon": round(total_sold_noon / window_days, 2) if window_days > 0 else 0,
            "draw_amz": round(total_sold_amz / window_days, 2) if window_days > 0 else 0,
            "draw_ty": round(total_sold_ty / window_days, 2) if window_days > 0 else 0,
            "days_to_stockout": round(days_to_stockout) if days_to_stockout is not None else None,
            "reorder_in": round(reorder_in) if reorder_in is not None else None,
            "status": status, "color": color, "suggested_qty": suggested_qty, "img": img, "china": china,
        })

    rows.sort(key=lambda r: (r["days_to_stockout"] if r["days_to_stockout"] is not None else 10**9))
    return rows

def compute_no_sale_rows(env, days=7):
    """يرجع قائمة SKUs الموجودة في المخزون (inv_map) ومالهاش أي مبيع في آخر عدد
    أيام محدد — نفس منطق تاب 'مخزون بدون بيع' بالظبط، لكن بيتقرا هنا كمان
    (مش بس جوه تاب16) عشان يتستخدم في شريط التنبيهات السريعة بالداشبورد."""
    today_ns = datetime.now().date()
    dates_ns = [today_ns - timedelta(days=i) for i in range(1, days + 1)]
    counts_ns = build_daily_orders_counts(env, dates_ns)
    rows = []
    for sku_up, info in env.inv_map.items():
        dc = counts_ns.get(sku_up, {})
        if sum(dc.get(d, 0) for d in dates_ns) > 0:
            continue
        rows.append({
            "sku": info.get("sku", sku_up), "sku_up": sku_up,
            "stock": info.get("total_stock", 0), "sales_month": info.get("sales", 0),
            "img": info.get("img", ""),
        })
    rows.sort(key=lambda x: -x["stock"])
    return rows

def compute_no_sale_rows_amazon(env, days=7):
    """نسخة أمازون من compute_no_sale_rows فوق — مستقلة تمامًا عن نون: بتقرا من
    amz_inv_map_asin (StockAmazon) ومبيعات DailyOrdersAmazon مجمّعة على ASIN،
    ومقصورة على تخزين FBA بس (نفس فلتر تاب مراجعة/مبيعات أمازون وتاب مخزون بدون
    بيع) | Amazon counterpart of compute_no_sale_rows above — fully independent
    of Noon: reads amz_inv_map_asin (StockAmazon) and DailyOrdersAmazon sales
    grouped by ASIN, restricted to FBA storage only (matching the Amazon
    review/sales tabs and the No-Sales tab's own Amazon filter)."""
    today_ns_amz = datetime.now().date()
    dates_ns_amz = [today_ns_amz - timedelta(days=i) for i in range(1, days + 1)]
    counts_ns_amz = build_daily_orders_counts_amazon_fba_asin(env, dates_ns_amz)
    rows = []
    for asin_up, info in env.amz_inv_map_asin.items():
        dc = counts_ns_amz.get(asin_up, {})
        if sum(dc.get(d, 0) for d in dates_ns_amz) > 0:
            continue
        mskus = info.get("skus", [])
        rows.append({
            "sku": ", ".join(mskus) if mskus else asin_up, "sku_up": asin_up,
            "asin": info.get("asin", asin_up),
            "stock": info.get("stock", 0), "sales_month": 0,
            "img": "",
        })
    rows.sort(key=lambda x: -x["stock"])
    return rows

def compute_cross_store_analysis(env, win=30):
    today = datetime.now().date()
    d_now = [today - timedelta(days=i) for i in range(1, win + 1)]
    d_prev = [today - timedelta(days=i) for i in range(win + 1, 2 * win + 1)]
    d_all = d_now + d_prev
    sn = set(d_now)
    c_noon = build_daily_orders_counts(env, d_all)
    c_amz = build_daily_orders_counts_amazon(env, d_all, mode="all", group_by="asin")
    c_ty = build_daily_orders_counts_trendyol(env, d_all)
    p_noon = build_daily_orders_prices(env, d_now)
    p_amz = build_daily_orders_prices_amazon_all_asin(env, d_now)
    p_ty = build_daily_orders_prices_trendyol(env, d_now)

    def split(cnts, key):
        dc = cnts.get(key, {})
        a = sum(v for d, v in dc.items() if d in sn)
        return a, sum(dc.values()) - a

    wh = env.get_warehouse_stock_map()
    noon_map, amz_map = get_tacweed_map(env), get_tacweed_amazon_map(env)
    rev_n, rev_a = get_tacweed_reverse_map(env), get_tacweed_amazon_reverse_map(env)
    ty_listed = set(env.get_ty_products()["name"].keys()) | set(env.get_trendyol_bundle()["name"].keys())
    links = env.get_links_map()
    codes = set(rev_n) | set(rev_a) | set(wh)
    rows = []
    for code in codes:
        skus, asins = rev_n.get(code, []), rev_a.get(code, [])
        n_now = n_prev = a_now = a_prev = t_now = t_prev = 0
        for k in skus:
            x, y = split(c_noon, k); n_now += x; n_prev += y
            x, y = split(c_ty, k); t_now += x; t_prev += y
        for k in asins:
            x, y = split(c_amz, k); a_now += x; a_prev += y
        n_stock = sum(env.inv_map.get(k, {}).get("total_stock", 0) or 0 for k in skus)
        a_stock = sum(amz_map_stock(env, k) for k in asins)
        n_listed = any(k in env.inv_map for k in skus)
        a_listed = any(k in env.amz_inv_map_asin for k in asins)
        t_listed = any(k in ty_listed for k in skus) or t_now + t_prev > 0
        try:
            w = float(str(wh.get(code, 0)).replace(",", "").strip() or 0)
        except Exception:
            w = 0.0
        pr = {}
        for nm, mp, ks in (("noon", p_noon, skus), ("amz", p_amz, asins), ("ty", p_ty, skus)):
            vals = [v for v in (_avg_price_from(mp.get(k)) for k in ks) if v]
            if vals:
                pr[nm] = sum(vals) / len(vals)
        rows.append({
            "code": code, "img": next((links.get(k) for k in skus if links.get(k)), ""),
            "skus": skus, "asins": asins, "wh": w,
            "noon": n_now, "amz": a_now, "ty": t_now,
            "noon_prev": n_prev, "amz_prev": a_prev, "ty_prev": t_prev,
            "n_stock": n_stock, "a_stock": a_stock,
            "n_listed": n_listed, "a_listed": a_listed, "t_listed": t_listed,
            "price": pr, "total": n_now + a_now + t_now, "prev_total": n_prev + a_prev + t_prev,
        })
    # مبيعات من غير كود | sales with no mapping
    unm = []
    for k in c_noon:
        if k not in noon_map:
            x = sum(v for d, v in c_noon[k].items() if d in sn)
            if x:
                unm.append({"store": "🟢 نون", "key": k, "qty": x})
    for k in c_amz:
        if k not in amz_map:
            x = sum(v for d, v in c_amz[k].items() if d in sn)
            if x:
                unm.append({"store": "🟠 أمازون", "key": k, "qty": x})
    for k in c_ty:
        if k not in noon_map:
            x = sum(v for d, v in c_ty[k].items() if d in sn)
            if x:
                unm.append({"store": "🟣 تريندول", "key": k, "qty": x})
    unm.sort(key=lambda r: -r["qty"])
    # تذبذب السعر داخل نفس المتجر: أقل وأعلى سعر بيع لنفس المنتج خلال آخر 30 يوم | same-store price swings
    def _swings(since):
        out = []
        for store, pmap, kmap in (("🟢 نون", p_noon, noon_map), ("🟠 أمازون", p_amz, amz_map), ("🟣 تريندول", p_ty, noon_map)):
            for k, byday in pmap.items():
                pts = []
                for d, lst in byday.items():
                    if d < since:
                        continue
                    for it in lst:
                        try:
                            pr = float(str(it[0]).replace(",", "").strip())
                        except Exception:
                            continue
                        if pr > 0:
                            pts.append((pr, d))
                if len(pts) < 2:
                    continue
                lo, hi = min(pts), max(pts)
                if lo[0] > 0 and hi[0] / lo[0] > 1.0:
                    out.append({"store": store, "key": k, "code": kmap.get(k, ""), "lo": lo[0], "lo_d": lo[1], "hi": hi[0], "hi_d": hi[1],
                                "ratio": hi[0] / lo[0], "n": len(pts), "img": links.get(k, "")})
        out.sort(key=lambda r: -r["ratio"])
        return out
    swings = {"30": _swings(today - timedelta(days=win)), "7": _swings(today - timedelta(days=7))}
    # منتجات من غير رابط صورة (نون: links n / المخزون — أمازون: links a / links n — تريندول: links n / ملف تريندول)
    links_a = env.get_links_map_asin()
    ty_img = env.get_trendyol_bundle().get("img", {})
    ty_img2 = env.get_ty_products().get("img", {})
    def _q30(cnts, k):
        return sum(v for d, v in cnts.get(k, {}).items() if d in sn)
    noimg = []
    for k in set(env.inv_map) | set(c_noon):
        if links.get(k) or str((env.inv_map.get(k) or {}).get("img", "")).startswith("http"):
            continue
        q = _q30(c_noon, k)
        stk = (env.inv_map.get(k) or {}).get("total_stock", 0) or 0
        if q or stk:
            noimg.append({"store": "🟢 نون", "key": (env.inv_map.get(k) or {}).get("sku", k), "code": noon_map.get(k, ""), "qty": q, "stock": stk, "sheet": "links n"})
    for k in set(env.amz_inv_map_asin) | set(c_amz):
        info = env.amz_inv_map_asin.get(k) or {}
        if links_a.get(k) or any(links.get(str(m).upper()) for m in info.get("skus", [])):
            continue
        q = _q30(c_amz, k)
        stk = info.get("stock", 0) or 0
        if q or stk:
            noimg.append({"store": "🟠 أمازون", "key": info.get("asin", k), "code": amz_map.get(k, ""), "qty": q, "stock": stk, "sheet": "links a"})
    for k in c_ty:
        if links.get(k) or ty_img.get(k) or ty_img2.get(k):
            continue
        q = _q30(c_ty, k)
        if q:
            noimg.append({"store": "🟣 تريندول", "key": k, "code": noon_map.get(k, ""), "qty": q, "stock": 0, "sheet": "links n"})
    noimg.sort(key=lambda r: (-r["qty"], -r["stock"]))
    swings["noimg"] = noimg
    return rows, unm, swings

def amz_map_stock(env, asin_up):
    return (env.amz_inv_map_asin.get(asin_up, {}) or {}).get("stock", 0) or 0

def build_cross_alerts(rows, min_units=5):
    """يرجّع dict اسم_التنبيه -> قائمة صفوف (code, نص) | alert name -> list of (row, text)."""
    A = {k: [] for k in ("gap", "nolist", "oos", "dead", "drop", "price")}
    for r in rows:
        sell = {k: r[k] for k in ("noon", "amz", "ty")}
        stock = {"noon": r["n_stock"], "amz": r["a_stock"], "ty": None}
        listed = {"noon": r["n_listed"], "amz": r["a_listed"], "ty": r["t_listed"]}
        best = max(sell, key=sell.get)
        if sell[best] >= min_units:
            for o in sell:
                if o == best:
                    continue
                if listed[o] and sell[o] == 0:
                    A["gap"].append((r, f"بيبيع {sell[best]} في {_XS_NAMES[best]} وصفر في {_XS_NAMES[o]} رغم إنه موجود هناك (آخر 30 يوم) — راجع السعر/الصور/الإعلان"))
                elif not listed[o] and (r["wh"] > 0 or stock.get(best) or sell[best] >= 3 * min_units):
                    A["nolist"].append((r, f"بيبيع {sell[best]} في {_XS_NAMES[best]} ومش معروض في {_XS_NAMES[o]} — فرصة نشر"))
            for o in ("noon", "amz"):
                if o != best and listed[o] and stock[o] == 0 and r["wh"] > 0:
                    A["oos"].append((r, f"مخزون {_XS_NAMES[o]} صفر والكود بيبيع {sell[best]} في {_XS_NAMES[best]} والمستودع فيه {r['wh']:,.0f} — حوّل/اشحن"))
        tot_stock = (r["n_stock"] or 0) + (r["a_stock"] or 0) + r["wh"]
        if tot_stock > 0 and r["total"] == 0 and (r["n_listed"] or r["a_listed"] or r["wh"] > 0):
            A["dead"].append((r, f"مخزون {tot_stock:,.0f} (نون {r['n_stock']:,.0f} + أمازون {r['a_stock']:,.0f} + مستودع {r['wh']:,.0f}) ومفيش أي بيع في أي متجر آخر 30 يوم"))
        if r["prev_total"] >= 10 and r["total"] <= r["prev_total"] * 0.5:
            A["drop"].append((r, f"المبيعات نزلت من {r['prev_total']} إلى {r['total']} ({(1 - r['total'] / r['prev_total']) * 100:.0f}%-) مقارنة بالـ30 يوم اللي قبلها (كل المتاجر)"))
        pr = r["price"]
        if len(pr) >= 2:
            lo, hi = min(pr, key=pr.get), max(pr, key=pr.get)
            if pr[lo] > 0 and pr[hi] / pr[lo] >= 1.3:
                A["price"].append((r, f"فرق سعر: {_XS_NAMES[hi]} {pr[hi]:.0f} مقابل {_XS_NAMES[lo]} {pr[lo]:.0f} ({(pr[hi] / pr[lo] - 1) * 100:.0f}%) — تريندول بسعر صافي بعد الخصم"))
    for k in A:
        A[k].sort(key=lambda t: -(t[0]["total"] + t[0]["prev_total"]))
    return A

def compute_channel_gap(env, days=30):
    """بيرجع dict: fbn / fbb / fba / fsab -> قائمة صفوف (اللي باع في الفترة بس) مرتبة بالمبيعات.
    action: send (نودّي — متوفر بالمستودع) | send_nowh (محتاج بس مش متوفر بالمستودع) | pending | scheduled | ok."""
    days = max(int(days or 30), 1)
    today = datetime.now().date()
    dates = [today - timedelta(days=i) for i in range(1, days + 1)]
    yday = dates[0]
    cov_days = int(env.load_settings().get("schedule_coverage_days", "15") or 15)
    wh = env.get_warehouse_stock_map()
    tac_n, tac_a = get_tacweed_map(env), get_tacweed_amazon_map(env)
    pend = get_pending_approval_skus(env)
    links_n, links_a = env.get_links_map(), env.get_links_map_asin()
    out = {"fbn": [], "fbb": [], "fba": [], "fsab": []}

    def _wh(code):
        return _gap_qty(wh.get(code, "")) if code else 0.0

    # ── نون ──
    def _tot(dcs):
        return sum(dcs.get(d, 0) for d in dates) if dcs else 0
    c_fbn, c_fbb = build_daily_orders_counts_fbn(env, dates), build_daily_orders_counts_fbb(env, dates)
    c_fba, c_fsab = build_daily_orders_counts_amazon_fba_asin(env, dates), build_daily_orders_counts_amazon_normal_asin(env, dates)
    for kind, counts in (("fbn", c_fbn), ("fbb", c_fbb)):
        for sku_up, dc in counts.items():
            total = sum(dc.get(d, 0) for d in dates)
            if total <= 0:
                continue
            info = env.inv_map.get(sku_up) or {}
            stock = _gap_qty(info.get("total_stock", 0))
            code = (tac_n.get(sku_up, "") or "").strip()
            whq = _wh(code)
            # FBB: لو اتنقل لـ FBN مخزون FBN هيخدم مبيعات FBN + FBB مع بعض، فالتغطية على المجموع
            other = _tot(c_fbn.get(sku_up)) if kind == "fbb" else _tot(c_fbb.get(sku_up))
            dem = total + other
            avg = dem / days
            cover = gap_cover(stock, avg)
            if sku_up in pend:
                sched, sched_txt = "pending", "⏳ الجدولة قيد الموافقة — بانتظار الاعتماد"
            else:
                si = get_latest_schedule_info(env, sku_up)
                if si:
                    sched = "scheduled"
                    sched_txt = f"📅 مجدولة (ASN {si['asn']}) بتاريخ {si['date']}"
                else:
                    sched, sched_txt = "none", "❌ لا توجد جدولة"
            need = cover < cov_days      # المخزون الصغير (حتى قطعة) مش مخزون: بنقيسه بأيام تغطية المبيعات
            action = gap_action(need, whq, avg, cov_days, sched)
            sug = "  \n".join(_gap_sug(action, "تخزين نون (FBN)", whq, avg, cov_days))
            out[kind].append({
                "kind": kind, "key": info.get("sku", sku_up), "key_up": sku_up, "code": code,
                "img": info.get("img", "") or links_n.get(sku_up, ""),
                "sales": total, "dem": dem, "other": other, "yday": dc.get(yday, 0), "avg": avg,
                "cov_days": cov_days, "need_qty": gap_need_qty(avg, stock, cov_days),
                "stock": stock, "no_stock": need, "cover": cover,
                "sched": sched, "sched_txt": sched_txt, "wh": whq,
                "action": action, "suggestion": sug,
            })
    # ── أمازون ──
    for kind, counts in (("fba", c_fba), ("fsab", c_fsab)):
        for asin_up, dc in counts.items():
            total = sum(dc.get(d, 0) for d in dates)
            if total <= 0:
                continue
            info = env.amz_inv_map_asin.get(asin_up) or {}
            stock = _gap_qty(info.get("stock", 0))
            code = (tac_a.get(asin_up, "") or "").strip()
            whq = _wh(code)
            other = _tot(c_fba.get(asin_up)) if kind == "fsab" else _tot(c_fsab.get(asin_up))
            dem = total + other    # التغطية على مبيعات القناتين (المخزون المنقول بيخدمهم)
            avg = dem / days
            cover = gap_cover(stock, avg)
            need = cover < cov_days
            action = gap_action(need, whq, avg, cov_days)
            sug = "  \n".join(_gap_sug(action, "تخزين أمازون (FBA)", whq, avg, cov_days))
            out[kind].append({
                "kind": kind, "key": asin_up, "key_up": asin_up, "code": code,
                "img": links_a.get(asin_up, ""),
                "sales": total, "dem": dem, "other": other, "yday": dc.get(yday, 0), "avg": avg,
                "cov_days": cov_days, "need_qty": gap_need_qty(avg, stock, cov_days),
                "stock": stock, "no_stock": need, "cover": cover,
                "sched": "", "sched_txt": "", "wh": whq,
                "action": action, "suggestion": sug,
            })
    for k in out:
        out[k].sort(key=lambda r: (-r["sales"], r["key_up"]))
    return out

def compute_transferred_from_sales(env):
    """يحسب SKUs اللي هتتنقل لمراجعة المخزون (محتاج جدولة فقط، بدون أي جدولة أو ملاحظات)."""
    if not env.inv_map:
        return []
    settings_now = env.load_settings()
    sales_days_now  = int(settings_now.get("sales_display_days","7") or 7)
    delay_days_now  = int(settings_now.get("schedule_delay_days","3") or 3)
    cov_days_now    = int(settings_now.get("schedule_coverage_days","15") or 15)
    today_now = datetime.now().date()
    dates_now = [today_now - timedelta(days=i) for i in range(1, sales_days_now + 1)]
    counts_now  = build_daily_orders_counts(env, dates_now)
    result = []
    for sku_up, info in env.inv_map.items():
        stock        = info.get("total_stock", 0)
        sales_month  = info.get("sales", 0)
        img          = info.get("img", "")
        sku_disp     = info.get("sku", sku_up)
        day_counts   = counts_now.get(sku_up, {d: 0 for d in dates_now})
        total_recent = sum(day_counts.values())
        avg_daily    = (total_recent / sales_days_now) if sales_days_now > 0 else (sales_month / 30 if sales_month > 0 else 0)
        eff_avg      = avg_daily if avg_daily > 0 else (sales_month / 30 if sales_month > 0 else 0)
        days_to_so   = round(stock / eff_avg) if eff_avg > 0 else 9999
        # نفس التصحيح اللي اتعمل في تاب المبيعات: لو مفيش أي متوسط بيع خالص،
        # يبقى المخزون مش بينزل أصلاً، فمينفعش يترحّل كـ "محتاج جدولة"
        stock_ok     = (eff_avg <= 0) or (days_to_so >= cov_days_now)
        if stock_ok:
            continue
        badge_text, _, sched = schedule_coverage_badge(env, sku_disp, days_to_so, delay_days_now)
        un_notes = get_unavailable_ordered_note(env, sku_disp)
        is_needs_sched_only = (
            "محتاج جدولة" in badge_text
            and not sched
            and not un_notes
        )
        if is_needs_sched_only:
            result.append({
                "sku": sku_disp, "sku_up": sku_up, "stock": stock,
                "sales_month": sales_month, "img": img,
                "effective_avg": eff_avg, "days_to_stockout": days_to_so,
                "day_counts": day_counts,
            })
    return result

def check_skus(env, tokens, days=30):
    """بيرجع (noon_rows, amazon_rows, notfound). نفس منطق عروض FBN/FBB/FBA/FSAB بس على مدخلات محددة."""
    days = max(int(days), 1)
    today = datetime.now().date()
    dates = [today - timedelta(days=i) for i in range(1, days + 1)]
    cov_days = int(env.load_settings().get("schedule_coverage_days", "15") or 15)
    c_fbn, c_fbb = build_daily_orders_counts_fbn(env, dates), build_daily_orders_counts_fbb(env, dates)
    c_fba, c_fsab = build_daily_orders_counts_amazon_fba_asin(env, dates), build_daily_orders_counts_amazon_normal_asin(env, dates)
    wh = env.get_warehouse_stock_map(); tac_n, tac_a = get_tacweed_map(env), get_tacweed_amazon_map(env)
    rev_a = get_tacweed_amazon_reverse_map(env); pend = get_pending_approval_skus(env); recent = get_recent_schedule_rows(env, days_back=90)
    known = set(env.inv_map) | set(tac_n) | set(c_fbn) | set(c_fbb) | pend
    def _tot(dcs):
        return sum(dcs.values()) if dcs else 0
    def _need(avg, stock):
        return gap_need_qty(avg, stock, cov_days)
    noon_rows, amz_rows, notfound, seen_n, seen_a = [], [], [], set(), set()

    def add_asin(asin_up, tok):
        if (tok, asin_up) in seen_a:
            return
        seen_a.add((tok, asin_up))
        info = env.amz_inv_map_asin.get(asin_up) or {}
        stock = float(info.get("stock", 0) or 0)
        fba, fsab = _tot(c_fba.get(asin_up)), _tot(c_fsab.get(asin_up))
        dem = fba + fsab; avg = dem / days
        cover = gap_cover(stock, avg)
        code = (tac_a.get(asin_up, "") or "").strip(); whq = _gap_qty(wh.get(code, "")) if code else 0.0
        need = dem > 0 and cover < cov_days
        if dem == 0:
            action, lines = "nosales", [f"لا توجد مبيعات في آخر {days} يوم"]
        else:
            action = gap_action(need, whq, avg, cov_days)
            lines = ["مخزون أمازون كافٍ"] if action == "ok" else _gap_sug(action, "تخزين أمازون (FBA)", whq, avg, cov_days)
        nq = _need(avg, stock) if dem > 0 else 0
        amz_rows.append({"المدخل": tok, "ASIN": asin_up, "الكود": code, "مبيعات FBA": fba, "مبيعات FSAB": fsab,
                         "متوسط يومي": round(avg, 2), "مخزون أمازون": stock, "يكفي (يوم)": round(cover) if cover < 9999 else "",
                         "متوفر بالمستودع": whq, f"الكمية المطلوبة لتغطية {cov_days} يوم": nq,
                         "الكمية المقترح إرسالها": nq if action == "send" else 0,
                         "الكمية المطلوب توريدها": max(nq - int(whq), 0) if action in ("low_wh", "send_nowh") else 0,
                         "الاقتراح": lines[0] if lines else "", "السبب": lines[1] if len(lines) > 1 else "", "المطلوب": lines[2] if len(lines) > 2 else ""})

    for tok in tokens:
        t = tok.upper()
        if re.fullmatch(r"B0[A-Z0-9]{8}", t):
            if t in env.amz_inv_map_asin or t in tac_a or t in c_fba or t in c_fsab:
                add_asin(t, tok)
            else:
                notfound.append({"المدخل": tok, "السبب": "ASIN غير موجود في مخزون/تكويد/مبيعات أمازون"})
            continue
        has_suffix = bool(re.search(r"-\d+$", t))
        cands = [t, re.sub(r"-\d+$", "", t)] if has_suffix else ([t + "-1", t] if re.fullmatch(r"Z[A-Z0-9]+Z", t) else [t, t + "-1"])
        sku_up = next((c for c in cands if c in known), None)
        found = sku_up is not None
        if not found:
            sku_up = cands[0]
        if (tok, sku_up) in seen_n:
            continue
        seen_n.add((tok, sku_up))
        info = env.inv_map.get(sku_up) or {}
        stock = float(info.get("total_stock", 0) or 0)
        fbn, fbb = _tot(c_fbn.get(sku_up)), _tot(c_fbb.get(sku_up))
        dem = fbn + fbb; avg = dem / days
        cover = gap_cover(stock, avg)
        code = (tac_n.get(sku_up, "") or "").strip(); whq = _gap_qty(wh.get(code, "")) if code else 0.0
        # الجدولة
        si = get_latest_schedule_info(env, sku_up); rc = recent.get(sku_up)
        if sku_up in pend:
            sched, sched_txt = "pending", "⏳ قيد الموافقة"
        elif si or rc:
            e = si or rc
            ref = (si or {}).get("parsed") or (rc or {}).get("parsed")
            ref = ref.date() if hasattr(ref, "date") else ref
            n = (today - ref).days if ref else None
            when = (_ago_txt(n) if (n is not None and n >= 0) else f"مجدولة بتاريخ {e.get('date', '')}") if ref else f"مجدولة بتاريخ {e.get('date', '')}"
            tag = " [منتهية]" if (not si and rc and rc.get("source") == "Expired") else ""
            sched = "scheduled" if si else "expired"
            sched_txt = f"📅 ASN {e.get('asn', '')} · {e.get('qty', '')} قطعة · {when}{tag}"
        else:
            sched, sched_txt = "none", "❌ لا توجد جدولة"
        need = dem > 0 and cover < cov_days
        if not found and not info and dem == 0 and not code:
            notfound.append({"المدخل": tok, "السبب": "SKU غير موجود في المخزون/التكويد/المبيعات (جربت مع -1 وبدونها)"}); 
            continue
        if dem == 0:
            action, lines = "nosales", [f"لا توجد مبيعات في آخر {days} يوم"]
        else:
            action = gap_action(need, whq, avg, cov_days, sched)
            lines = ["مخزون FBN كافٍ"] if action == "ok" else _gap_sug(action, "تخزين نون (FBN)", whq, avg, cov_days)
        nq = _need(avg, stock) if dem > 0 else 0
        noon_rows.append({"المدخل": tok, "SKU": sku_up, "الكود": code, "مبيعات FBN": fbn, "مبيعات FBB": fbb,
                          "متوسط يومي": round(avg, 2), "مخزون FBN": stock, "يكفي (يوم)": round(cover) if cover < 9999 else "",
                          "حالة الجدولة": sched_txt, "متوفر بالمستودع": whq,
                          f"الكمية المطلوبة لتغطية {cov_days} يوم": nq,
                          "الكمية المقترح إرسالها": nq if action == "send" else 0,
                          "الكمية المطلوب توريدها": max(nq - int(whq), 0) if action in ("low_wh", "send_nowh") else 0,
                          "الاقتراح": lines[0] if lines else "", "السبب": lines[1] if len(lines) > 1 else "", "المطلوب": lines[2] if len(lines) > 2 else ""})
        for a in rev_a.get(code, []) if code else []:      # ASINs أمازون لنفس الكود (FBA / FSAB)
            add_asin(a, tok)
    return noon_rows, amz_rows, notfound

def _t14_build_rows(env, kind, sales_dates, sales_display_days):
    """نفس بناء صفوف تابي FBN/FBB القديمين بالظبط. بيرجع (rows, multi_counts)."""
    if kind == "fbn":
        multi_counts_t14 = build_daily_orders_counts_fbn(env, sales_dates)
        prices_map_t14   = build_daily_orders_prices_fbn(env, sales_dates)
    else:
        multi_counts_t14 = build_daily_orders_counts_fbb(env, sales_dates)
        prices_map_t14   = build_daily_orders_prices_fbb(env, sales_dates)
    rows = []
    for sku_up, info in env.inv_map.items():
        stock       = info.get("total_stock", 0)
        sales_month = info.get("sales", 0)
        img         = info.get("img", "")
        sku_disp    = info.get("sku", sku_up)
        day_counts  = multi_counts_t14.get(sku_up, {d: 0 for d in sales_dates})
        day_prices  = prices_map_t14.get(sku_up, {d: [] for d in sales_dates})
        total_recent = sum(day_counts.get(d, 0) for d in sales_dates)
        avg_daily_t14 = (total_recent / sales_display_days) if sales_display_days > 0 else (sales_month / 30 if sales_month > 0 else 0)
        effective_avg_t14 = avg_daily_t14 if avg_daily_t14 > 0 else (sales_month / 30 if sales_month > 0 else 0)
        days_to_stockout_t14 = round(stock / effective_avg_t14) if effective_avg_t14 > 0 else 9999
        rows.append({
            "sku": sku_disp, "sku_up": sku_up,
            "stock": stock, "sales_month": sales_month, "img": img,
            "day_counts": day_counts, "day_prices": day_prices,
            "total_recent": total_recent,
            "effective_avg": effective_avg_t14,
            "days_to_stockout": days_to_stockout_t14,
        })
    if kind == "fbb":
        # SKUs باعت طلبات FBB لكن مش موجودة في ملف المخزون المرفوع — بتتضاف عشان كل مبيعات FBB تظهر
        inv_skus_seen_fbb = set(env.inv_map.keys())
        links_map_fbb_orphan = env.get_links_map()
        for sku_up, day_counts in multi_counts_t14.items():
            if sku_up in inv_skus_seen_fbb:
                continue
            total_recent_orphan = sum(day_counts.get(d, 0) for d in sales_dates)
            if total_recent_orphan <= 0:
                continue
            day_prices_orphan = prices_map_t14.get(sku_up, {d: [] for d in sales_dates})
            rows.append({
                "sku": sku_up, "sku_up": sku_up,
                "stock": 0, "sales_month": 0, "img": links_map_fbb_orphan.get(sku_up, ""),
                "day_counts": day_counts, "day_prices": day_prices_orphan,
                "total_recent": total_recent_orphan,
                "effective_avg": 0,
                "days_to_stockout": 9999,
                "not_in_inventory": True,
            })
    return rows, multi_counts_t14

def _t14_build_rows_amz(env, kind, sales_dates, sales_display_days):
    """نفس بناء تابي مبيعات أمازون في App.py بالظبط (مجمّع على ASIN، شهري = مجموع 30 يوم فعلي).
    kind: 'FSAB' (تخزين عادي) أو 'FBA'. بيرجع (rows, multi_counts)."""
    today_a = datetime.now().date()
    dates_30d = [today_a - timedelta(days=i) for i in range(1, 31)]
    if kind == "FSAB":
        mc = build_daily_orders_counts_amazon_normal_asin(env, sales_dates)
        pm = build_daily_orders_prices_amazon_normal_asin(env, sales_dates)
        mo = build_daily_orders_counts_amazon_normal_asin(env, dates_30d)
    else:
        mc = build_daily_orders_counts_amazon_fba_asin(env, sales_dates)
        pm = build_daily_orders_prices_amazon_fba_asin(env, sales_dates)
        mo = build_daily_orders_counts_amazon_fba_asin(env, dates_30d)
    rows = build_amazon_asin_sales_rows(env, mc, pm, mo, sales_dates, sales_display_days)
    return rows, mc

def _ct_add_months(d, n):
    m0 = d.month - 1 + n
    y, m = d.year + m0 // 12, m0 % 12 + 1
    return d.replace(year=y, month=m, day=min(d.day, _ct_cal.monthrange(y, m)[1]))

def contract_events(start, end):
    """[(token, date, عنوان عربي | English)] — نهاية كل شهر من العقد، وآخرها نهاية العقد."""
    ev, n = [], 1
    while n <= 600:
        d = _ct_add_months(start, n)
        if d >= end:
            break
        ar = CT_ORD_AR[n - 1] if n <= len(CT_ORD_AR) else str(n)
        ev.append((f"M{n}", d, f"انتهى الشهر {ar} من العقد | Month {n} of the contract has ended"))
        n += 1
    ev.append(("END", end, "انتهى العقد | The contract has ended"))
    return ev

def contract_active_events(env, today=None):
    """الإشعارات الشغالة النهاردة (النهاردة أو إمبارح كان تاريخها) — مع حالة «تم الاطلاع»."""
    today = today or datetime.now().date()
    res = []
    for c in env._ct_read():
        if not c["start"] or not c["end"] or c["end"] <= c["start"]:
            continue
        for tok, d, label in contract_events(c["start"], c["end"]):
            if d <= today <= d + timedelta(days=CT_WINDOW_DAYS - 1):
                res.append({"c": c, "token": tok, "date": d, "label": label, "acked": tok in c["acked"],
                            "days_left": (c["end"] - today).days})
    return res

def contract_missed_events(env, today=None):
    """إشعارات انتهت مدة ظهورها (يومين) ولسه ماتضغطش عليها «تم الاطلاع» — بتفضل ظاهرة كشريط أحمر بدون منع."""
    today = today or datetime.now().date()
    res = []
    for c in env._ct_read():
        if not c["start"] or not c["end"] or c["end"] <= c["start"] or not c["added"]:
            continue
        for tok, d, label in contract_events(c["start"], c["end"]):
            if d + timedelta(days=CT_WINDOW_DAYS) <= today and d >= c["added"] and tok not in c["acked"]:
                res.append({"c": c, "token": tok, "date": d, "label": label, "acked": False,
                            "days_left": (c["end"] - today).days})
    return res


# ══════════════════════════════════════════════
# ثوابت (أعمدة الشيتات، حالات الصين، ...)
# ══════════════════════════════════════════════
_XL_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
_TY_CID_RE = re.compile(r"-p-(\d+)")
_TY_ALLOC_CACHE = {}
_GAP_KINDS = {
    "fbn":  ("🟢 نون FBN",  "noon"),
    "fbb":  ("🟢 نون FBB",  "noon"),
    "fba":  ("🟠 أمازون FBA",  "amz"),
    "fsab": ("🟠 أمازون عادي FSAB", "amz"),
}
_GAP_ACTION_AR = {"send": "إرسال (متوفر بالمستودع)", "low_wh": "لا يُرسل — كمية المستودع قليلة", "send_nowh": "لا يُرسل — غير متوفر بالمستودع",
                  "pending": "قيد الموافقة", "scheduled": "مجدول", "ok": "مخزون متوفر"}



# ══════════════════════════════════════════════
# حسابات تعتمد على بيانات: env هو بوابة البيانات (في التطبيق = stock_requests.py)
# ══════════════════════════════════════════════

def extract_excel_images(fileobj):
    """يرجّع [(barcode, bytes_الصورة)] من إكسيل فيه صور داخل الخلايا (Place in cell) | in-cell pictures keyed by the barcode column."""
    import openpyxl
    import zipfile
    from xml.etree import ElementTree as ET
    raw = fileobj.getvalue() if hasattr(fileobj, 'getvalue') else (fileobj if isinstance(fileobj, (bytes, bytearray)) else fileobj.read())
    z = zipfile.ZipFile(io.BytesIO(raw))
    names = set(z.namelist())
    # vm -> media
    rels = {}
    rp = "xl/richData/_rels/richValueRel.xml.rels"
    for r in ET.fromstring(z.read(rp)):
        rels[r.get("Id")] = "xl/" + r.get("Target").replace("../", "")
    rel_ids = [e.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
               for e in ET.fromstring(z.read("xl/richData/richValueRel.xml"))]
    rv = [[v.text for v in e if v.tag.endswith("}v")] for e in ET.fromstring(z.read("xl/richData/rdrichvalue.xml"))]
    # أي حقل فيه رقم صالح كـ index للصورة: الأول عادةً (LocalImageIdentifier)
    def media_for_vm(vm):
        try:
            idx = int(rv[vm - 1][0])
            return rels[rel_ids[idx]]
        except Exception:
            return None
    # أول ورقة
    wbx = ET.fromstring(z.read("xl/workbook.xml"))
    sheet0 = "xl/worksheets/sheet1.xml"
    wrels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    s = wbx.find("m:sheets", _XL_NS)[0]
    t = wrels.get(s.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"), "")
    sheet0 = "xl/" + t.lstrip("/").replace("xl/", "") if t else sheet0
    # رؤوس الأعمدة: عمود الباركود وعمود الصورة
    ws = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True).worksheets[0]
    hdr = None; vals = {}
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if hdr is None:
            hdr = [str(c).strip() if c is not None else "" for c in row]
            ib = next((j for j, h in enumerate(hdr) if "باركود" in h or h.lower() == "barcode"), 1)
            continue
        v = row[ib] if ib < len(row) else None
        if v is not None and str(v).strip():
            vals[i] = str(int(v)) if isinstance(v, float) and v == int(v) else str(v).strip()
    ii = next((j for j, h in enumerate(hdr) if h == "الصورة" or "image" in h.lower() or "صورة" in h), 4)
    col = ""
    n = ii
    while True:
        col = chr(65 + n % 26) + col
        n = n // 26 - 1
        if n < 0: break
    out = {}
    for _, c in ET.iterparse(io.BytesIO(z.read(sheet0)), events=("end",)):
        if c.tag.endswith("}c") and c.get("vm"):
            ref = c.get("r"); m = re.match(r"([A-Z]+)(\d+)", ref)
            if m and m.group(1) == col:
                bc = vals.get(int(m.group(2)))
                mp = media_for_vm(int(c.get("vm")))
                if bc and mp and mp in names and bc not in out:
                    out[bc] = z.read(mp)
        c.clear()
    # صور عايمة (Drawing) متثبتة على الخلية: from.row (يبدأ من 0) + العمود | floating pictures anchored to cells
    dp = "xl/drawings/drawing1.xml"
    if dp in names:
        dr = {r.get("Id"): "xl/" + r.get("Target").replace("../", "") for r in ET.fromstring(z.read("xl/drawings/_rels/drawing1.xml.rels"))}
        X = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
        A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
        R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
        for anc in ET.fromstring(z.read(dp)):
            fr = anc.find(X + "from"); pic = anc.find(X + "pic")
            if fr is None or pic is None:
                continue
            if int(fr.find(X + "col").text) != ii:
                continue
            blip = pic.find(".//" + A + "blip")
            mp = dr.get(blip.get(R + "embed")) if blip is not None else None
            bc = vals.get(int(fr.find(X + "row").text) + 1)
            if bc and mp and mp in names and bc not in out:
                out[bc] = z.read(mp)
    return list(out.items())

def file_timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def to_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()

def make_empty_template(columns):
    return to_excel(pd.DataFrame(columns=columns))

def build_noon_link(sku: str):
    """يبني لينك منتج نون من الـ SKU — بيشيل أي لاحقة رقمية زي -1 في الآخر (متغيّر/variant)
    قبل ما يحطه في اللينك. | Builds a noon.com product link from a SKU — strips a trailing
    "-<number>" variant suffix (e.g. "-1") before inserting it in the URL.
    مثال | Example: N96556579A -> https://www.noon.com/saudi-ar/N96556579A/p/
    مثال | Example: ZF5A935F52FB29CDD6CB9Z-1 -> https://www.noon.com/saudi-ar/ZF5A935F52FB29CDD6CB9Z/p/"""
    if not sku:
        return None
    s = str(sku).strip().upper()
    if not s:
        return None
    m = re.match(r'^(.*)-\d+$', s)
    if m:
        s = m.group(1)
    return f"https://www.noon.com/saudi-ar/{s}/p/"

def sku_link_html(sku: str, extra_style: str = ""):
    """SKU كنص قابل للنسخ (تحديد/كوبي) جنبه أيقونة 🔗 صغيرة منفصلة تودّي على صفحة المنتج على نون
    في تاب جديد — بالطريقة دي تقدر تنسخ الـ SKU من غير ما تتفتحلك صفحة نون بالغلط.
    | SKU rendered as selectable/copyable text next to a small separate 🔗 link icon that opens
    the noon.com product page in a new tab — keeps the SKU copyable without accidentally
    navigating away."""
    link = build_noon_link(sku)
    sku_span = (
        f'<span style="font-family:monospace;font-weight:700;color:#e2e8f0;'
        f'background:#1e293b;border:1px solid #334155;border-radius:6px;padding:2px 10px;'
        f'user-select:all;-webkit-user-select:all;{extra_style}">{sku}</span>'
    )
    if not link:
        return sku_span
    return (
        f'{sku_span} <a href="{link}" target="_blank" rel="noopener" '
        f'style="text-decoration:none;font-size:14px;" title="فتح صفحة المنتج على نون | Open on noon.com">🔗</a>'
    )

def build_amazon_asin_link(asin: str):
    if not asin:
        return None
    a = str(asin).strip().upper()
    if not a:
        return None
    return f"https://www.amazon.sa/dp/{a}"

def asin_link_html(asin: str, extra_style: str = ""):
    """ASIN كنص قابل للنسخ جنبه أيقونة 🔗 صغيرة تودّي على صفحة المنتج على أمازون
    في تاب جديد — هو المعرّف الأساسي لمنتجات أمازون (زي ما الـ SKU هو المعرّف
    الأساسي لمنتجات نون) | ASIN rendered as copyable text with a small 🔗 link
    to the amazon.sa product page — the primary identifier for Amazon products
    (the way SKU is the primary identifier for Noon products)."""
    asin_disp = asin if asin else "—"
    link = build_amazon_asin_link(asin)
    asin_span = (
        f'<span style="font-family:monospace;font-weight:700;color:#e2e8f0;'
        f'background:#1e293b;border:1px solid #334155;border-radius:6px;padding:2px 10px;'
        f'user-select:all;-webkit-user-select:all;{extra_style}">{asin_disp}</span>'
    )
    if not link:
        return asin_span
    return (
        f'{asin_span} <a href="{link}" target="_blank" rel="noopener" '
        f'style="text-decoration:none;font-size:14px;" title="فتح صفحة المنتج على أمازون | Open on amazon.sa">🔗</a>'
    )

def _ty_parse_orders_xlsx(fileobj):
    """يقرأ ملف طلبات تريندول (Excel) ويطلّع الأعمدة المطلوبة بس بترتيب TABS_CONFIG['DailyOrdersTrendyol'].
    مبيقراش أي عمود فيه بيانات عملاء. بيرجع (rows, عدد الصفوف المتخطّية)."""
    import openpyxl, re as _re
    wb = openpyxl.load_workbook(fileobj, read_only=True)
    ws = wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    hdr = None
    for r in it:
        if r and any(str(c).strip().lower() == "barcode" for c in r if c is not None) and \
           any(str(c).strip().lower() == "order date" for c in r if c is not None):
            hdr = list(r)
            break
    if hdr is None:
        raise ValueError("مش لاقي صف العناوين (Barcode / Order Date) في الملف")
    c_bc  = _ty_find_col(hdr, "barcode")
    c_dt  = _ty_find_col(hdr, "order date")
    c_ord = _ty_find_col(hdr, "order number")
    c_pkg = _ty_find_col(hdr, "package no")
    c_st  = _ty_find_col(hdr, "order status")
    c_q   = _ty_find_col(hdr, "quantity")
    c_inv = _ty_find_col(hdr, "invoice amount")
    c_cm  = _ty_find_col(hdr, "commission rate")
    c_nm  = _ty_find_col(hdr, "product name")
    c_im  = next((i for i, h in enumerate(hdr) if h is not None and str(h).strip().lower().startswith("product image")), None)
    if None in (c_bc, c_dt, c_ord, c_st, c_q, c_inv):
        raise ValueError("الملف ناقص أعمدة مطلوبة (Barcode / Order Date / Order Number / Order Status / Quantity / Invoice Amount)")
    now_s = now_str()
    out, skipped, seen = [], 0, {}
    def cell(r, i):
        return r[i] if (i is not None and i < len(r)) else None
    for r in it:
        bc = str(cell(r, c_bc) or "").strip()
        dv = cell(r, c_dt)
        if not bc or dv is None:
            skipped += 1
            continue
        if isinstance(dv, datetime):
            dt = dv
        else:
            dt = None
            for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y %H:%M:%S", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
                try:
                    dt = datetime.strptime(str(dv).strip(), fmt)
                    break
                except Exception:
                    pass
        if dt is None:
            skipped += 1
            continue
        try:
            qty = max(int(float(str(cell(r, c_q) or 1).replace(",", ""))), 1)
        except Exception:
            qty = 1
        try:
            inv = float(str(cell(r, c_inv) or 0).replace(",", ""))
        except Exception:
            inv = 0.0
        order = str(cell(r, c_ord) or "").strip()
        pkg = str(cell(r, c_pkg) or "").strip() if c_pkg is not None else ""
        base_key = f"{order}|{pkg}|{bc}"
        seen[base_key] = seen.get(base_key, 0) + 1
        key = base_key if seen[base_key] == 1 else f"{base_key}#{seen[base_key]}"
        img = ""
        if c_im is not None:
            m = _re.search(r'https?://[^"\')\s]+', str(cell(r, c_im) or ""))
            img = m.group(0) if m else ""
        out.append([key, bc, dt.strftime("%Y-%m-%d %H:%M:%S"), order, str(cell(r, c_st) or "").strip(),
                    str(qty), f"{inv / qty:.2f}", f"{inv:.2f}", str(cell(r, c_cm) or "").strip(),
                    img, str(cell(r, c_nm) or "").strip()[:120], now_s])
    return out, skipped

def _ty_ads_parse_xlsx(fileobj):
    """تقرير Product Ads من تريندول → صفوف بأعمدة الشيت. بيتقرا بالاسم (أول ظهور لكل عمود)."""
    import openpyxl
    wb = openpyxl.load_workbook(fileobj, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    allrows = list(ws.iter_rows(values_only=True))
    hi = None
    for i, r in enumerate(allrows[:15]):
        low = [str(c).strip().lower() for c in r if c is not None]
        if "ad name" in low and "cost" in low:
            hi = i
            break
    if hi is None:
        raise ValueError("مش لاقي عناوين التقرير (Ad name / Cost)")
    idx = {}
    for j, c in enumerate(allrows[hi]):
        k = str(c).strip().lower() if c is not None else ""
        if k and k not in idx:
            idx[k] = j
    def g(r, name):
        j = idx.get(name)
        v = r[j] if (j is not None and j < len(r)) else ""
        return "" if v is None else str(v).strip()
    out = []
    ts = now_str()
    for r in allrows[hi + 1:]:
        nm = g(r, "ad name")
        if not nm:
            continue
        out.append([nm, g(r, "ad status"), g(r, "start date"), g(r, "number of products"),
                    g(r, "product contentid list"), g(r, "daily budget"), g(r, "cost"), g(r, "clicks"),
                    g(r, "view(s)"), g(r, "total sold"), g(r, "total sales"), g(r, "roas"), ts])
    return out

def _ads_file_rows(fileobj, name, sheet_headers, key_cols):
    """يقرا xlsx/csv ويرجّع صفوف بترتيب هيدر الشيت (المطابقة بالاسم مش بالمكان) | rows aligned to the sheet header by name."""
    import io, csv
    raw = []
    if str(name).lower().endswith(".csv"):
        txt = fileobj.getvalue().decode("utf-8-sig", errors="replace")
        raw = [list(r) for r in csv.reader(io.StringIO(txt))]
    else:
        import openpyxl
        wb = openpyxl.load_workbook(fileobj, read_only=True, data_only=True)
        raw = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    want = {h.strip().lower() for h in key_cols}
    hi = None
    for i, r in enumerate(raw[:15]):
        low = {str(c).strip().lower() for c in r if c is not None}
        if want & low:
            hi = i
            break
    if hi is None:
        raise ValueError("مش لاقي عمود أساسي في الملف: " + " / ".join(key_cols))
    fidx = {}
    for j, c in enumerate(raw[hi]):
        k = str(c).strip().lower() if c is not None else ""
        if k and k not in fidx:
            fidx[k] = j
    cols = [fidx.get(h.strip().lower()) for h in sheet_headers]
    matched = [h for h, c in zip(sheet_headers, cols) if c is not None]
    missing = [h for h, c in zip(sheet_headers, cols) if c is None]
    kc = [fidx[k.strip().lower()] for k in key_cols if k.strip().lower() in fidx][0]
    out = []
    for r in raw[hi + 1:]:
        if kc >= len(r) or r[kc] is None or not str(r[kc]).strip():
            continue
        out.append([("" if (c is None or c >= len(r) or r[c] is None) else r[c]) for c in cols])
    return out, matched, missing

def _ty_products_parse_xlsx(fileobj):
    import openpyxl
    wb = openpyxl.load_workbook(fileobj, data_only=True)
    ws = wb["Products"] if "Products" in wb.sheetnames else wb.worksheets[0]
    allrows = list(ws.iter_rows(values_only=True))
    hi = None
    for i, r in enumerate(allrows[:10]):
        low = [str(c).strip().lower() for c in r if c is not None]
        if "barcode" in low and "trendyol.com link" in low:
            hi = i
            break
    if hi is None:
        raise ValueError("مش لاقي عمودي Barcode و Trendyol.com Link")
    idx = {}
    for j, c in enumerate(allrows[hi]):
        k = str(c).strip().lower() if c is not None else ""
        if k and k not in idx:
            idx[k] = j
    def g(r, name):
        j = idx.get(name)
        v = r[j] if (j is not None and j < len(r)) else ""
        return "" if v is None else str(v).strip()
    out, seen, ts = [], set(), now_str()
    for r in allrows[hi + 1:]:
        bc = g(r, "barcode")
        m = _TY_CID_RE.search(g(r, "trendyol.com link"))
        if not bc or not m or (m.group(1), bc) in seen:
            continue
        seen.add((m.group(1), bc))
        out.append([m.group(1), bc, g(r, "product name")[:120], g(r, "image 1"), ts])
    return out

def _ty_alloc(env):
    """يوزّع أرقام كل حملة على Barcodes منتجاتها. بيرجع (by_bc, by_camp):
    by_bc: BARCODE -> [إدخالات بنفس شكل إعلانات نون]، by_camp: ترتيب الحملة -> تفاصيل التوزيع."""
    if "v" in _TY_ALLOC_CACHE:
        return _TY_ALLOC_CACHE["v"]
    camps = env.get_ads_map_trendyol()
    by_cid = env.get_ty_products()["by_cid"]
    qty_rows = {}
    for r in env.get_trendyol_bundle()["rows"]:
        if r["ok"]:
            qty_rows.setdefault(r["bc"], []).append((r["d"], r["qty"]))
    by_bc, by_camp = {}, {}
    for c in camps:
        bcs, unmapped = [], []
        for cid in c["ids"]:
            if cid in by_cid:
                for bc in by_cid[cid]:
                    if bc not in bcs:
                        bcs.append(bc)
            else:
                unmapped.append(cid)
        try:
            start = datetime.strptime(str(c["start"])[:10], "%Y-%m-%d").date()
        except Exception:
            start = None
        w = {bc: sum(q for d, q in qty_rows.get(bc, []) if (start is None or d >= start)) for bc in bcs}
        tot_w = sum(w.values())
        exact = len(bcs) == 1
        items = []
        for bc in bcs:
            share = 1.0 if exact else ((w[bc] / tot_w) if tot_w > 0 else 1.0 / len(bcs))
            items.append((bc, share, w[bc]))
            sp, rv, od, ck, vw = c["cost"] * share, c["sales"] * share, c["sold"] * share, c["clicks"] * share, c["views"] * share
            by_bc.setdefault(bc, []).append({
                "campaign": c["name"] + ("" if exact else " (تقريبي)"), "views": vw, "clicks": ck, "orders": od, "atc": 0.0,
                "spends": sp, "revenue": rv, "ctr": (ck / vw * 100) if vw else 0.0, "roas": (rv / sp) if sp else 0.0,
                "cpc": (sp / ck) if ck else 0.0, "cps": (sp / od) if od else 0.0, "cvr": (od / ck * 100) if ck else 0.0,
                "exact": exact})
        by_camp[c["i"]] = {"items": items, "exact": exact, "unmapped": unmapped}
    _TY_ALLOC_CACHE["v"] = (by_bc, by_camp)
    return _TY_ALLOC_CACHE["v"]

def get_ads_by_barcode_trendyol(env):
    return _ty_alloc(env)[0]

def channel_gap_alerts(env, res=None):
    """التنبيهات: (1) نون FBB لها مبيعات ومخزون FBN غير كافٍ. (2) أمازون لها مبيعات ومخزونها في أمازون غير كافٍ. كل واحد مقسوم send / send_nowh / مجدول / قيد الموافقة."""
    res = res or env.channel_gap_cached()
    a = {"fbb_nostock": [r for r in res["fbb"] if r["no_stock"] and r["other"] <= 0],
         "amz_nostock": [r for r in res["fba"] + [x for x in res["fsab"] if x["other"] <= 0] if r["no_stock"]]}
    _seen, _u = set(), []          # نفس الـ ASIN لو باع FBA وFSAB بيتعد مرة واحدة
    for r in a["amz_nostock"]:
        if r["key_up"] not in _seen:
            _seen.add(r["key_up"]); _u.append(r)
    a["amz_nostock"] = _u
    return a

def channel_gap_df(rows, kind):
    noon = kind in ("fbn", "fbb")
    recs = []
    for r in rows:
        d = {"الكود": r["code"], ("SKU" if noon else "ASIN"): r["key"], "القناة": _GAP_KINDS[kind][0],
             "مبيعات أمس": r["yday"], "مبيعات الفترة": r["sales"], "متوسط يومي": round(r["avg"], 2),
             ("مخزون نون FBN" if noon else "مخزون أمازون"): r["stock"],
             "إجمالي المبيعات المحسوب عليها": r["dem"],
             "يكفي (يوم)": round(r["cover"]) if r["cover"] < 9999 else "",
             ("مخزون غير كافٍ في نون" if noon else "مخزون غير كافٍ في أمازون"): "نعم" if r["no_stock"] else "لا"}
        if noon:
            d["حالة الجدولة"] = r["sched_txt"]
        d["متوفر بالمستودع"] = r["wh"]
        _nq = r["need_qty"]
        d[f"الكمية المطلوبة لتغطية {r['cov_days']} يوم"] = _nq
        d["الكمية المقترح إرسالها"] = _nq if r["action"] == "send" else 0
        d["الكمية المطلوب توريدها"] = max(_nq - int(r["wh"]), 0) if r["action"] in ("low_wh", "send_nowh") else 0
        d["الإجراء"] = _GAP_ACTION_AR.get(r["action"], "")
        _l = (r["suggestion"] or "").split("  \n")
        d["الاقتراح"] = _l[0] if _l else ""
        d["السبب"] = _l[1] if len(_l) > 1 else ""
        d["المطلوب"] = _l[2] if len(_l) > 2 else ""
        recs.append(d)
    return pd.DataFrame(recs)

def _file_to_text(f):
    name = (f.name or "").lower()
    if name.endswith(".csv"):
        df = pd.read_csv(f, header=None, dtype=str, keep_default_na=False)
        return "\n".join(df.astype(str).values.ravel())
    sheets_ = pd.read_excel(f, sheet_name=None, header=None, dtype=str)
    return "\n".join("\n".join(d.fillna("").astype(str).values.ravel()) for d in sheets_.values())
