# -*- coding: utf-8 -*-
"""
finance_tb.py — تاب 📒 المالية: ميزان المراجعة (TB) وجرد المخزون (INV) من ملف المالية.

الآلية (مصدر واحد، وطريقتين للإدخال):
  ملف جوجل شيت منفصل «Finance» (عشان ملف Complaints قرّب من حد الـ 10 مليون خلية) فيه تابات بنفس شكل ملف المالية:
    TB2024 / TB2025 / TB2026 … : ميزان المراجعة (صف لكل حساب، أعمدة الشهور)
    INV2024 / INV2025 / INV2026 … : جرد المخزون
  1) المحاسب يكتب/يلصق في التابات دي مباشرة في جوجل شيت.
  2) أو نرفع ملف الإكسيل من التطبيق → بيحدّث نفس التابات بالمفتاح (رمز الحساب / الباركود+المادة+المستودع+الفرع) من غير مسح.
  التطبيق بيقرا التابات دي ويعرضها بفلاتر. عمود «الإجمالي» بيتحسب في التطبيق (مش بيتكتب من الرفع).
"""
import re
import io
import time
import datetime as dt

FIN_SHEET_ID = "1PWnYshOIcUDC0_pBVAJoTVjsQeYKDfEGM_a57TPgoD4"
SERVICE_ACCOUNT_EMAIL = "notes-service@notesapp-471512.iam.gserviceaccount.com"

OPENING = "Opening"
TB_KEY = ("رمز الحساب",)
INV_KEY = ("رمز الباركود", "رمز المادة", "المستودع", "الفرع")
SKIP_HDRS = ("الإجمالي",)
TB_RE, INV_RE = re.compile(r"TB\s*(\d{4})", re.I), re.compile(r"INV\s*(\d{4})", re.I)
TB_HEAD = ["رمز الحساب", "اسم الحساب", "الحساب الرئيسي 01", "الحساب الرئيسي 02", "الحساب الختامي", "الرصيد الافتتاحي"]


# ═══════════════════════ أدوات صغيرة ═══════════════════════
def _s(v):
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def _num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace(",", "")
    if t == "" or t.startswith("#"):
        return None
    try:
        return float(t)
    except ValueError:
        return None


def _norm(v):
    """قيمة مفتاح موحّدة (رقم صحيح/نص) عشان تتقارن بين اللي في الملف واللي في الشيت."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return str(int(v)) if float(v) == int(v) else str(v)
    return str(v).strip() if v is not None else ""


# ═══════════════════════ قراءة ملف المالية (خام) ═══════════════════════
def read_sheets(file_like):
    """{اسم الشيت: (عناوين، صفوف)} لشيتات TB#### و INV#### بس. القيم أرقام/نصوص زي ما هي. بيرجع (dict، تحذيرات)."""
    import openpyxl
    wb = openpyxl.load_workbook(file_like, read_only=True, data_only=True)
    out, warn = {}, []
    try:
        for name in wb.sheetnames:
            n = name.strip()
            if not (TB_RE.fullmatch(n) or INV_RE.fullmatch(n)):
                warn.append(f"شيت «{name}» اتجاهل (مش TB#### ولا INV####)")
                continue
            it = wb[name].iter_rows(values_only=True)
            try:
                hdr = [_s(h) for h in next(it)]
            except StopIteration:
                continue
            rows = [list(r) for r in it if any(c is not None and str(c).strip() != "" for c in r)]
            out[re.sub(r"\s+", "", n).upper()] = (hdr, rows)
    finally:
        wb.close()
    return out, warn


def to_long(sheets):
    """من شيتات عريضة {اسم: (hdr, rows)} → (tb_rows, inv_rows) بصيغة طويلة للعرض.
    tb_rows: [سنة، رمز، اسم، رئيسي01، رئيسي02، ختامي، فترة، مبلغ] · inv_rows: [سنة، باركود، رمز، اسم، وحدة، كمية، سعر، قيمة، مستودع، فرع]"""
    tb, inv, info = {}, {}, {"tb_sheets": [], "inv_sheets": [], "warnings": []}
    for name, (hdr, rows) in sheets.items():
        m_tb, m_inv = TB_RE.fullmatch(name), INV_RE.fullmatch(name)
        year = int((m_tb or m_inv).group(1))
        if m_tb:
            _parse_tb(iter(rows), hdr, year, tb, info, name)
        else:
            _parse_inv(iter(rows), hdr, year, inv, info, name)
    tb_rows = [[k[0], k[1], meta[0], meta[1], meta[2], meta[3], k[2], amt] for k, (amt, meta) in tb.items()]
    inv_rows = [[k[0], k[1], k[2], v["name"], v["unit"], v["qty"], (v["val"] / v["qty"] if v["qty"] else v["price"]), v["val"], k[3], k[4]] for k, v in inv.items()]
    return tb_rows, inv_rows, info


def _col(hdr, *prefixes):
    for i, h in enumerate(hdr):
        for p in prefixes:
            if h.startswith(p):
                return i
    return None


def _parse_tb(it, hdr, year, out, info, name):
    i_code, i_name = _col(hdr, "رمز الحساب"), _col(hdr, "اسم الحساب")
    i_m1, i_m2, i_cl = _col(hdr, "الحساب الرئيسي 01"), _col(hdr, "الحساب الرئيسي 02"), _col(hdr, "الحساب الختامي")
    i_op = _col(hdr, "الرصيد الافتتاحي")
    pcols = []
    for i, h in enumerate(hdr):
        m = re.fullmatch(r"(\d{1,2})-(\d{4})", h)
        if m:
            pcols.append((i, f"{m.group(2)}-{int(m.group(1)):02d}"))
    if i_code is None or i_name is None or not pcols:
        info["warnings"].append(f"شيت «{name}»: مش لاقي أعمدة رمز الحساب/اسم الحساب/الشهور")
        return
    info["tb_sheets"].append(name)
    if i_op is not None:
        pcols = [(i_op, OPENING)] + pcols
    for r in it:
        code = _s(r[i_code]) if i_code < len(r) else ""
        if not code:
            continue
        meta = (_s(r[i_name]), _s(r[i_m1]) if i_m1 is not None and i_m1 < len(r) else "",
                _s(r[i_m2]) if i_m2 is not None and i_m2 < len(r) else "", _s(r[i_cl]) if i_cl is not None and i_cl < len(r) else "")
        for i, per in pcols:
            v = _num(r[i]) if i < len(r) else None
            if v is None or v == 0:
                continue
            k = (year, code, per)
            out[k] = (round(out[k][0] + v, 2) if k in out else round(v, 2), meta)


def _parse_inv(it, hdr, year, out, info, name):
    ix = {"bar": _col(hdr, "رمز الباركود"), "code": _col(hdr, "رمز المادة"), "name": _col(hdr, "اسم المادة"), "unit": _col(hdr, "الوحدة"),
          "qty": _col(hdr, "الكمية"), "price": _col(hdr, "السعر"), "val": _col(hdr, "القيمة"), "wh": _col(hdr, "المستودع"), "br": _col(hdr, "الفرع")}
    if ix["code"] is None or ix["qty"] is None:
        info["warnings"].append(f"شيت «{name}»: مش لاقي أعمدة رمز المادة/الكمية")
        return
    info["inv_sheets"].append(name)
    g = lambda r, k: (_s(r[ix[k]]) if ix[k] is not None and ix[k] < len(r) else "")
    for r in it:
        code = g(r, "code")
        if not code:
            continue
        qty = _num(r[ix["qty"]]) or 0.0
        price = (_num(r[ix["price"]]) if ix["price"] is not None and ix["price"] < len(r) else None) or 0.0
        val = _num(r[ix["val"]]) if ix["val"] is not None and ix["val"] < len(r) else None
        val = qty * price if val is None else val
        k = (year, g(r, "bar"), code, g(r, "wh"), g(r, "br"))
        if k in out:
            out[k]["qty"] += qty
            out[k]["val"] += val
        else:
            out[k] = {"name": g(r, "name"), "unit": g(r, "unit"), "qty": qty, "price": price, "val": val}
    for v in out.values():
        v["qty"], v["val"] = round(v["qty"], 4), round(v["val"], 2)


# ═══════════════════════ التحديث بالمفتاح (من غير مسح) ═══════════════════════
def _retry(fn, tries=5):
    for a in range(tries):
        try:
            return fn()
        except Exception as e:
            if a == tries - 1:
                raise
            time.sleep(65 if "429" in str(e) else 3 * (a + 1))


def _read_all(ws):
    try:
        return _retry(lambda: ws.get_all_values(value_render_option="UNFORMATTED_VALUE"))
    except TypeError:
        return _retry(lambda: ws.get_all_values())


def _same(a, b):
    na, nb = _num(a), _num(b)
    if na is not None and nb is not None:
        return abs(na - nb) < 0.005
    return _norm(a) == _norm(b)


def _letter(n):
    t = ""
    while n:
        n, r = divmod(n - 1, 26)
        t = chr(65 + r) + t
    return t


def _keyed(rows, idx):
    """مفتاح لكل صف مع رقم التكرار (لو نفس المفتاح اتكرر): (k..., n)."""
    cnt, out = {}, []
    for r in rows:
        k = tuple(_norm(r[i]) if i < len(r) else "" for i in idx)
        cnt[k] = cnt.get(k, 0) + 1
        out.append(k + (cnt[k],))
    return out


def upsert_wide(ws, f_hdr, f_rows, key_names, log=print, chunk_ranges=400):
    """بيحدّث التاب العريض بمفتاح من غير مسح: الخلية المتغيّرة بتتحدّث، الصف الجديد بيتضاف، وأي صف/عمود في الشيت مش في الملف بيتساب.
    الأعمدة بتتطابق بالاسم (عمود جديد في الملف بيتضاف يمين)، وعمود «الإجمالي» مش بيتكتب. بيرجع (صفوف اتضافت، خلايا اتحدّثت)."""
    f_cols = [(i, h) for i, h in enumerate(f_hdr) if h and h not in SKIP_HDRS]
    vals = _read_all(ws)
    if not vals or not any(str(c).strip() for c in vals[0]):
        hdr = [h for _, h in f_cols]
        vals = [hdr]
        _retry(lambda: ws.update(range_name=f"A1:{_letter(len(hdr))}1", values=[hdr], value_input_option="RAW"))
    s_hdr = [str(h).strip() for h in vals[0]]
    for _, h in f_cols:
        if h not in s_hdr:
            s_hdr.append(h)
    missing_key = [k for k in key_names if k not in s_hdr]
    if missing_key:
        raise RuntimeError(f"تاب {ws.title}: مفيش عمود {', '.join(missing_key)} — مفيش حاجة اتكتبت")
    if len(s_hdr) > len(vals[0]):
        need_c = len(s_hdr)
        if getattr(ws, "col_count", need_c) < need_c:
            _retry(lambda: ws.resize(cols=need_c))
        _retry(lambda: ws.update(range_name=f"A1:{_letter(len(s_hdr))}1", values=[s_hdr], value_input_option="RAW"))
    pos = {h: j for j, h in enumerate(s_hdr)}
    body = [list(r) + [""] * (len(s_hdr) - len(r)) for r in vals[1:]]
    body = [r[:len(s_hdr)] for r in body]
    k_idx = [pos[k] for k in key_names]
    sheet_keys = _keyed(body, k_idx)
    where = {k: j for j, k in enumerate(sheet_keys) if any(k[:-1])}
    f_pos = {h: i for i, h in f_cols}
    f_key_idx = [f_hdr.index(k) for k in key_names if k in f_hdr]
    if len(f_key_idx) != len(key_names):
        raise RuntimeError(f"ملف المالية ناقص عمود من: {', '.join(key_names)}")
    f_keys = _keyed(f_rows, f_key_idx)
    changes, new_rows = {}, []
    for k, r in zip(f_keys, f_rows):
        if not any(k[:-1]):
            continue
        j = where.get(k)
        rowvals = {h: (r[i] if i < len(r) else None) for h, i in ((h, f_pos[h]) for _, h in f_cols)}
        if j is None:
            nr = [""] * len(s_hdr)
            for h, v in rowvals.items():
                nr[pos[h]] = "" if v is None else (_s(v) if isinstance(v, str) else v)
            new_rows.append(nr)
            continue
        for h, v in rowvals.items():
            cur = body[j][pos[h]]
            nv = "" if v is None else v
            if isinstance(nv, str):
                nv = nv.strip()
            if not _same(cur, nv) and not (cur in ("", None) and nv in ("", None, 0, 0.0)):
                changes.setdefault(pos[h], {})[j] = nv
    ranges = []
    for c, rr in changes.items():
        L = _letter(c + 1)
        js = sorted(rr)
        start = prev = js[0]
        run = [rr[start]]
        for j in js[1:]:
            if j == prev + 1:
                run.append(rr[j])
            else:
                ranges.append({"range": f"{L}{start + 2}:{L}{prev + 2}", "values": [[x] for x in run]})
                start, run = j, [rr[j]]
            prev = j
        ranges.append({"range": f"{L}{start + 2}:{L}{prev + 2}", "values": [[x] for x in run]})
    for s0 in range(0, len(ranges), chunk_ranges):
        part = ranges[s0:s0 + chunk_ranges]
        _retry(lambda p=part: ws.batch_update(p, value_input_option="RAW"))
        time.sleep(1.2)
    if new_rows:
        first = len(body) + 2
        need = first + len(new_rows)
        if getattr(ws, "row_count", need) < need:
            _retry(lambda: ws.resize(rows=need + 200))
        L = _letter(len(s_hdr))
        for s0 in range(0, len(new_rows), 5000):
            part = new_rows[s0:s0 + 5000]
            _retry(lambda p=part, a=first + s0: ws.update(range_name=f"A{a}:{L}{a + len(p) - 1}", values=p, value_input_option="RAW"))
            time.sleep(1.2)
    return len(new_rows), sum(len(v) for v in changes.values())


def get_or_create(sh, title, ncols=12):
    try:
        return sh.worksheet(title)
    except Exception as e:
        if "WorksheetNotFound" not in type(e).__name__ and "not found" not in str(e).lower():
            raise
        return sh.add_worksheet(title=title, rows="1500", cols=str(ncols))


def save_to_sheet(sh, sheets, log=print):
    """sheets: ناتج read_sheets. بيحدّث كل تاب بنفس اسمه في ملف Finance. بيرجع {اسم: (اتضاف، خلايا اتحدّثت)}."""
    out = {}
    for name, (hdr, rows) in sheets.items():
        ws = get_or_create(sh, name, max(len([h for h in hdr if h]), 12))
        key = TB_KEY if TB_RE.fullmatch(name) else INV_KEY
        out[name] = upsert_wide(ws, hdr, rows, key, log)
        log(f"{name}: +{out[name][0]} صف جديد · {out[name][1]} خلية اتحدّثت")
    return out


# ═══════════════════════ الواجهة (Streamlit) ═══════════════════════
_SH_REF = {}
try:
    import streamlit as _st

    @_st.cache_data(ttl=300, show_spinner=False)
    def _read(tab):
        return _read_all(_SH_REF["sh"].worksheet(tab))

    @_st.cache_data(ttl=300, show_spinner=False)
    def _tab_names(_v):
        return [w.title for w in _SH_REF["sh"].worksheets()]
except Exception:          # للاختبار من غير streamlit
    class _R:
        def __call__(self, *a):
            raise RuntimeError("streamlit مش متاح")

        def clear(self):
            pass
    _read = _tab_names = _R()


def _open(st, client):
    key = "_fin_sh"
    if key not in st.session_state:
        st.session_state[key] = _retry(lambda: client.open_by_key(FIN_SHEET_ID))
    return st.session_state[key]


def load_long(st, sh):
    """بيقرا كل تابات TB#### و INV#### من ملف Finance ويرجّعها بصيغة طويلة (tb_rows, inv_rows, tabs)."""
    _SH_REF["sh"] = sh
    names = [n for n in _tab_names(0) if TB_RE.fullmatch(n.strip()) or INV_RE.fullmatch(n.strip())]
    sheets = {}
    for n in names:
        vals = _read(n)
        if vals and len(vals) > 1:
            hdr = [_s(h) for h in vals[0]]
            sheets[re.sub(r"\s+", "", n).upper()] = (hdr, vals[1:])
    tb, inv, _ = to_long(sheets)
    return tb, inv, names


def render(st, client, pd):
    st.subheader("📒 المالية | Finance")
    try:
        sh = _open(st, client)
    except Exception as e:
        st.error(f"مش قادر أفتح ملف Finance: {str(e)[:200]}")
        st.info(f"اتأكد إن ملف «Finance» مشارك مع `{SERVICE_ACCOUNT_EMAIL}` بصلاحية Editor.")
        return

    top = st.columns([4, 1])
    top[0].caption("البيانات من ملف جوجل شيت «Finance» (تابات TB#### و INV####). المحاسب يقدر يعدّل فيها مباشرة، أو ترفع الإكسيل من تحت. التحديث بيظهر خلال 5 دقايق أو بزر التحديث.")
    if top[1].button("🔄 تحديث", key="fin_refresh"):
        _read.clear()
        _tab_names.clear()
        st.rerun()

    with st.expander("📤 رفع ملف المالية (Trial Balance)", expanded=False):
        up = st.file_uploader("ملف الإكسيل (شيتات TB#### و INV####)", type=["xlsx", "xlsm"], key="fin_upload")
        if up is not None:
            sheets = None
            try:
                sheets, warn = read_sheets(io.BytesIO(up.getvalue()))
            except Exception as e:
                st.error(f"مقدرتش أقرا الملف: {str(e)[:200]}")
            if sheets is not None:
                if not sheets:
                    st.warning("مفيش شيتات TB#### أو INV#### في الملف.")
                else:
                    st.info("الشيتات: " + " · ".join(f"{n} ({len(r):,} صف)" for n, (h, r) in sheets.items()))
                    for w in warn:
                        st.caption("⚠️ " + w)
                    st.caption("الحفظ بيحدّث الخلايا المتغيّرة ويضيف الصفوف الجديدة بس — مفيش حاجة بتتمسح، وعمود «الإجمالي» مش بيتكتب.")
                    if st.button("📤 حفظ في ملف Finance", type="primary", key="fin_save"):
                        msgs = []
                        with st.spinner("بيكتب في الشيت..."):
                            try:
                                save_to_sheet(sh, sheets, msgs.append)
                                _read.clear()
                                _tab_names.clear()
                                st.success("✅ تم الحفظ")
                            except Exception as e:
                                st.error(f"❌ فشل الحفظ: {str(e)[:250]} (اللي اتكتب قبل الخطأ فاضل، وممكن تعيد الرفع)")
                        for m in msgs:
                            st.write("• " + m)

    try:
        tb_rows, inv_rows, names = load_long(st, sh)
    except Exception as e:
        st.error(f"مقدرتش أقرا تابات المالية: {str(e)[:200]}")
        return
    if not names:
        st.info("مفيش تابات TB#### أو INV#### في ملف Finance لسه — ارفع الملف من فوق أو الصقها في الشيت.")
        return
    tb_df = pd.DataFrame(tb_rows, columns=["Year", "AccountCode", "AccountName", "Main01", "Main02", "Closing", "Period", "Amount"])
    inv_df = pd.DataFrame(inv_rows, columns=["Year", "Barcode", "ItemCode", "ItemName", "Unit", "Qty", "Price", "Value", "Warehouse", "Branch"])
    t_dash, t_tb, t_inv = st.tabs(["🧭 لوحة القيادة", "📒 ميزان المراجعة", "📦 جرد المخزون"])
    with t_dash:
        try:
            import finance_dash
            finance_dash.render(st, pd, tb_df, inv_df)
        except Exception as e:
            st.error(f"تعذر عرض لوحة القيادة: {type(e).__name__}: {str(e)[:250]}")
    with t_tb:
        _render_tb(st, pd, tb_df)
    with t_inv:
        _render_inv(st, pd, inv_df)


def _xlsx(pd, df, sheet="data"):
    b = io.BytesIO()
    with pd.ExcelWriter(b, engine="openpyxl") as w:
        df.to_excel(w, sheet_name=sheet[:30], index=False)
    return b.getvalue()


def _render_tb(st, pd, df):
    if df.empty:
        st.info("مفيش بيانات لسه — ارفع ملف المالية من فوق.")
        return
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0.0)
    df["Year"] = df["Year"].astype(str).str.replace(r"\.0$", "", regex=True)
    years = sorted(df["Year"].unique(), reverse=True)
    c1, c2, c3 = st.columns([1, 2, 2])
    yr = c1.selectbox("السنة المالية (تبدأ أبريل)", years, key="fin_tb_year")
    d = df[df["Year"] == yr]
    cats = sorted([x for x in d["Main01"].unique() if x])
    sel = c2.multiselect("التصنيف الرئيسي", cats, key="fin_tb_cat")
    q = c3.text_input("بحث (رمز أو اسم حساب)", key="fin_tb_q").strip()
    if sel:
        d = d[d["Main01"].isin(sel)]
    if q:
        d = d[d["AccountCode"].astype(str).str.contains(q, case=False, regex=False) | d["AccountName"].str.contains(q, case=False, regex=False)]
    if d.empty:
        st.info("مفيش نتائج بالفلاتر دي.")
        return
    per = sorted(d["Period"].unique(), key=lambda p: ("0" if p == OPENING else "1") + p)
    pv = d.pivot_table(index=["AccountCode", "AccountName", "Main01", "Main02", "Closing"], columns="Period", values="Amount", aggfunc="sum", fill_value=0.0)
    pv = pv.reindex(columns=per).reset_index()
    pv["الإجمالي"] = pv[per].sum(axis=1)
    pv = pv.rename(columns={"AccountCode": "رمز الحساب", "AccountName": "اسم الحساب", "Main01": "الرئيسي 01", "Main02": "الرئيسي 02", "Closing": "الختامي", OPENING: "الرصيد الافتتاحي"})
    k = d.groupby("Main01")["Amount"].sum().sort_values(ascending=False)
    cols = st.columns(min(4, max(1, len(k))))
    for i, (name, v) in enumerate(k.items()):
        if i < len(cols):
            cols[i].metric(name or "—", f"{v:,.0f}")
    st.caption("الأرقام زي ما هي في ملف المالية (رصيد افتتاحي + حركة كل شهر). الإجمالي = مجموع الأعمدة.")
    num = [c for c in pv.columns if c not in ("رمز الحساب", "اسم الحساب", "الرئيسي 01", "الرئيسي 02", "الختامي")]
    st.dataframe(pv, use_container_width=True, hide_index=True,
                 column_config={c: st.column_config.NumberColumn(c, format="%.2f") for c in num})
    st.download_button("⬇️ تحميل (Excel)", data=_xlsx(pd, pv, "TrialBalance"), key="fin_tb_dl",
                       file_name=f"trial_balance_{yr}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def _render_inv(st, pd, df):
    if df.empty:
        st.info("مفيش بيانات لسه — ارفع ملف المالية من فوق.")
        return
    for c in ("Qty", "Price", "Value"):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df["Year"] = df["Year"].astype(str).str.replace(r"\.0$", "", regex=True)
    years = sorted(df["Year"].unique(), reverse=True)
    c1, c2, c3, c4 = st.columns([1, 2, 2, 2])
    yr = c1.selectbox("السنة", years, key="fin_inv_year")
    d = df[df["Year"] == yr]
    brs = sorted([x for x in d["Branch"].unique() if x])
    whs = sorted([x for x in d["Warehouse"].unique() if x])
    sb = c2.multiselect("الفرع", brs, key="fin_inv_br")
    sw = c3.multiselect("المستودع", whs, key="fin_inv_wh")
    q = c4.text_input("بحث (اسم/باركود/رمز)", key="fin_inv_q").strip()
    if sb:
        d = d[d["Branch"].isin(sb)]
    if sw:
        d = d[d["Warehouse"].isin(sw)]
    if q:
        d = d[d["ItemName"].str.contains(q, case=False, regex=False) | d["Barcode"].astype(str).str.contains(q, regex=False) | d["ItemCode"].astype(str).str.contains(q, regex=False)]
    d = d[(d["Qty"] != 0) | (d["Value"] != 0)]
    if d.empty:
        st.info("مفيش نتائج بالفلاتر دي.")
        return
    m = st.columns(3)
    m[0].metric("💰 إجمالي قيمة المخزون", f"{d['Value'].sum():,.0f}")
    m[1].metric("📦 إجمالي الكمية", f"{d['Qty'].sum():,.0f}")
    m[2].metric("🏷️ عدد الأصناف", f"{d[['ItemCode']].drop_duplicates().shape[0]:,}")
    g = d.groupby(["Branch", "Warehouse"], as_index=False).agg(الكمية=("Qty", "sum"), القيمة=("Value", "sum")).sort_values("القيمة", ascending=False)
    g = g.rename(columns={"Branch": "الفرع", "Warehouse": "المستودع"})
    st.markdown("**حسب الفرع والمستودع**")
    st.dataframe(g, use_container_width=True, hide_index=True, column_config={"القيمة": st.column_config.NumberColumn("القيمة", format="%.2f"), "الكمية": st.column_config.NumberColumn("الكمية", format="%.0f")})
    t = d.rename(columns={"Barcode": "الباركود", "ItemCode": "رمز المادة", "ItemName": "المادة", "Unit": "الوحدة", "Qty": "الكمية", "Price": "السعر", "Value": "القيمة", "Warehouse": "المستودع", "Branch": "الفرع"})
    t = t[["الباركود", "رمز المادة", "المادة", "الوحدة", "الكمية", "السعر", "القيمة", "المستودع", "الفرع"]].sort_values("القيمة", ascending=False)
    st.markdown(f"**الأصناف** ({len(t):,})")
    st.dataframe(t.head(2000), use_container_width=True, hide_index=True,
                 column_config={"السعر": st.column_config.NumberColumn("السعر", format="%.2f"), "القيمة": st.column_config.NumberColumn("القيمة", format="%.2f")})
    if len(t) > 2000:
        st.caption("معروض أعلى 2000 صنف بالقيمة — حمّل الإكسيل للكل.")
    st.download_button("⬇️ تحميل (Excel)", data=_xlsx(pd, t, "Inventory"), key="fin_inv_dl",
                       file_name=f"inventory_{yr}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
