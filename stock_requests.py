# -*- coding: utf-8 -*-
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import time
import pandas as pd
import io
import re
import gspread.exceptions

st.set_page_config(page_title="📦 Stock Requests | طلبات المخزون", page_icon="📦", layout="wide")

# ══ اتصال ══
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
client = gspread.authorize(creds)

def open_spreadsheet(retries=5, delay=2):
    for attempt in range(retries):
        try:
            return client.open("Complaints")
        except gspread.exceptions.APIError as e:
            if attempt < retries - 1:
                time.sleep(delay * (2 ** attempt))
            else:
                raise e

ss = open_spreadsheet()

# ══ الأوراق ══
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
    "Tacweed":           ["SKU","Code01","Date Uploaded"],
    "WarehouseStock":    ["Code","Quantity","Item Name","Date Uploaded"],
    # تاب "قيد الموافقة" بيتكتب فيه من برنامج الجدولة المكتبي (schedule_entry_app) لما تتحدد
    # علامة Approval — نفس أعمدة تاب Scheduled الجديدة. هنا بنقراه بس (حسب اسم العمود مش
    # مكانه) عشان نعرض ملحوظة في تابات ثانية إن السكو ده لسه قيد الموافقة.
    "PendingApproval":   ["Store Name","Warehouse Name","ASN","Total QTY","Section","Approval",
                           "SKU","Quantity","Schedule Date","Image URL","Date Added","Notes","Flag"],
    # تاب الإعلانات — بيتحدّث يدوي من جوجل شيت، وبنعرض أداء كل SKU منه في تاب المبيعات
    "Advertisements":    ["Campaign Name","Sku","Views","Clicks","Orders","ATC","Spends","Revenue",
                           "CTR","ROAS","CPC","CPS","CVR"],
    # تاب العمولة ومصاريف التوصيل — بيتحدّث يدوي، وبنستخدمه لحساب صافي سعر البيع لكل SKU
    "COM":               ["SKU","مصاريف توصيل","العمولة"],
}

def get_or_create_worksheet(tab, headers, retries=5, delay=2):
    for attempt in range(retries):
        try:
            ws = ss.worksheet(tab)
            # sync header: أضيف الأعمدة الجديدة لو ناقصة
            try:
                existing_hdr = ws.row_values(1)
                missing = [h for h in headers if h not in existing_hdr]
                if missing:
                    for h in missing:
                        ws.append_cols([[h]], value_input_option="RAW")
            except Exception:
                pass
            return ws
        except gspread.exceptions.WorksheetNotFound:
            try:
                ws = ss.add_worksheet(title=tab, rows="3000", cols="12")
                ws.append_row(headers)
                return ws
            except gspread.exceptions.APIError as e:
                if attempt < retries - 1:
                    time.sleep(delay * (2 ** attempt))
                else:
                    raise e
        except gspread.exceptions.APIError as e:
            if attempt < retries - 1:
                time.sleep(delay * (2 ** attempt))
            else:
                raise e

# ══ تهيئة الشيتات مرة واحدة بس لكل جلسة — مش كل rerun ══
# (قبل كده كان بيتعمل ss.worksheet() + row_values() لكل الـ 13 تاب في كل ضغطة زرار،
#  وده اللي كان بيستهلك الـ quota بسرعة جداً ويسبب الخطأ المتكرر)
if "sheets_initialized" not in st.session_state:
    _sheets = {}
    for tab, headers in TABS_CONFIG.items():
        _sheets[tab] = get_or_create_worksheet(tab, headers)
    st.session_state["sheets_initialized"] = _sheets
sheets = st.session_state["sheets_initialized"]

def get_or_create_links_ws(retries=5, delay=2):
    for attempt in range(retries):
        try:
            return ss.worksheet("links n")
        except gspread.exceptions.WorksheetNotFound:
            try:
                ws = ss.add_worksheet(title="links n", rows="2000", cols="2")
                ws.append_row(["SKU","Image URL"])
                return ws
            except gspread.exceptions.APIError as e:
                if attempt < retries - 1:
                    time.sleep(delay * (2 ** attempt))
                else:
                    raise e
        except gspread.exceptions.APIError as e:
            if attempt < retries - 1:
                time.sleep(delay * (2 ** attempt))
            else:
                raise e

links_ws = get_or_create_links_ws()

requests_sheet    = sheets["Requests"]
approved_sheet    = sheets["Approved"]
unavailable_sheet = sheets["Unavailable"]
ordered_sheet     = sheets["Ordered"]
scheduled_sheet   = sheets["Scheduled"]
pending_approval_sheet = sheets["PendingApproval"]
# ══ توافق تاب الجدولة بعد إضافة أعمدة جديدة (Store Name / Warehouse Name / Total QTY / Section / Approval) ══
# برنامج الجدولة المكتبي (schedule_entry_app) بقى بيكتب في تاب "Scheduled" بترتيب أعمدة جديد:
# Store Name, Warehouse Name, ASN, Total QTY, Section, Approval, SKU, Quantity, Schedule Date,
# Image URL, Date Added, Notes, Flag
# لكن كل الكود هنا في stock_requests.py لسه مبني على الترتيب القديم:
# ASN, SKU, Quantity, Schedule Date, Image URL, Date Added, Notes, Flag
# بدل ما نغيّر كل مكان بيستخدم row[0]/row[1]/... في الملف ده، بنعمل طبقة توافق:
# بنقرا/بنكتب دايماً حسب اسم العمود مش مكانه، فأي ترتيب أعمدة جديد في الشيت هيفضل شغال تمام.
OLD_SCHED_COLS = ["ASN", "SKU", "Quantity", "Schedule Date", "Image URL", "Date Added", "Notes", "Flag"]

def get_scheduled_normalized(force=False):
    """يرجع نفس شكل get_all_values() (هيدر + صفوف) لكن بترتيب أعمدة ثابت (OLD_SCHED_COLS)
    بغض النظر عن ترتيب الأعمدة الحقيقي في تاب Scheduled دلوقتي — عشان باقي الكود اللي
    بيفترض row[0]=ASN, row[1]=SKU, row[2]=Quantity, row[3]=Schedule Date ...الخ يفضل شغال صح."""
    raw = get_cached(scheduled_sheet, force=force)
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

def build_scheduled_row(asn, sku, qty, sdate, img, date_added, notes="", flag=""):
    """يبني صف جاهز للإضافة في تاب Scheduled حسب ترتيب أعمدة الشيت الحقيقي دلوقتي
    (بما فيها الأعمدة الجديدة Store Name/Warehouse Name/Total QTY/Section/Approval)،
    وبيسيب الأعمدة الجديدة دي فاضية لأن stock_requests.py مش هو المصدر بتاعها —
    برنامج الجدولة المكتبي هو اللي بيملاها."""
    raw = get_cached(scheduled_sheet, force=False)
    real_header = raw[0] if raw else OLD_SCHED_COLS
    values = {
        "ASN": asn, "SKU": sku, "Quantity": qty, "Schedule Date": sdate,
        "Image URL": img, "Date Added": date_added, "Notes": notes, "Flag": flag,
    }
    return [values.get(h, "") for h in real_header]
cancelled_sheet   = sheets["CancelledSchedule"]
reschedule_sheet  = sheets["Rescheduled"]
expired_sheet     = sheets["Expired"]
inventory_sheet      = sheets["Inventory"]
daily_orders_sheet   = sheets["DailyOrders"]
settings_sheet       = sheets["Settings"]
cancel_notif_sheet   = sheets["CancelNotifications"]
tacweed_sheet        = sheets["Tacweed"]
warehouse_stock_sheet = sheets["WarehouseStock"]
ads_sheet             = sheets["Advertisements"]
com_sheet             = sheets["COM"]

# ══ كاش ══
def safe_get_all_values(sheet, retries=6, delay=1):
    """زي safe_append بالظبط بس للقراءة — كان ده الناقص اللي بيسبب الكراش."""
    last_err = None
    for attempt in range(retries):
        try:
            return sheet.get_all_values()
        except gspread.exceptions.APIError as e:
            last_err = e
            wait = delay * (2 ** attempt)
            if "429" in str(e) or "quota" in str(e).lower() or "RESOURCE_EXHAUSTED" in str(e):
                st.toast(f"⏳ Google Sheets API limit — جاري إعادة المحاولة ({attempt+1}/{retries})...", icon="⏳")
                time.sleep(wait)
            else:
                time.sleep(delay)
        except Exception as e:
            last_err = e
            time.sleep(delay)
    # خلصت كل المحاولات — ارجع آخر قيمة كانت متخزنة في الكاش لو موجودة بدل ما الابب يقع كله
    key = f"cache_{sheet.title}"
    if key in st.session_state:
        st.warning(f"⚠️ تعذر تحديث '{sheet.title}' من Google Sheets الآن — بيتم عرض آخر نسخة محفوظة | Showing last cached version")
        return st.session_state[key]
    st.error(f"❌ تعذر تحميل '{sheet.title}' من Google Sheets — حاول تاني بعد شوية | Could not load this sheet right now, please retry shortly.")
    st.stop()

def get_cached(sheet, force=False):
    key = f"cache_{sheet.title}"
    if force or key not in st.session_state:
        st.session_state[key] = safe_get_all_values(sheet)
    return st.session_state[key]

def clear_cache(sheet):
    key = f"cache_{sheet.title}"
    if key in st.session_state:
        del st.session_state[key]

# ══ إعدادات ══
def load_settings():
    data = get_cached(settings_sheet)
    s = {}
    for row in data[1:]:
        if len(row) >= 2:
            s[row[0]] = row[1]
    return s

def save_setting(key, value):
    data = get_cached(settings_sheet, force=True)
    for i, row in enumerate(data[1:], start=2):
        if len(row) >= 1 and row[0] == key:
            settings_sheet.update_cell(i, 2, value)
            clear_cache(settings_sheet)
            return
    settings_sheet.append_row([key, value])
    clear_cache(settings_sheet)

def get_excluded_warehouses():
    val = load_settings().get("excluded_warehouses","")
    if not val.strip():
        return set()
    return {w.strip().upper() for w in val.split(",") if w.strip()}

# ══ links map ══
@st.cache_data(ttl=300)
def get_links_map():
    data = safe_get_all_values(links_ws)
    m = {}
    for row in data[1:]:
        if len(row) >= 2 and row[0].strip():
            m[row[0].strip().upper()] = row[1].strip()
    return m

# ══ tacweed map (SKU -> الكود 01) ══
@st.cache_data(ttl=300)
def get_tacweed_map():
    data = safe_get_all_values(tacweed_sheet)
    m = {}
    for row in data[1:]:
        if len(row) >= 2 and row[0].strip():
            code = row[1].strip()
            if code:
                m[row[0].strip().upper()] = code
    return m

def big_note_html(text):
    """نص ملاحظة (غير متوفر سابقاً / تم طلبه سابقاً) بخط أكبر وأسود بولد بدل الكابشن الصغير الرمادي."""
    return f'<span class="status-badge-lg" style="background:#e5e7eb;">{text}</span>'

# ══ خريطة الإعلانات (SKU -> قائمة كامبينات) | Ads map (SKU -> list of campaigns) ══
def _f2(v, default=0.0):
    """تحويل آمن لأي قيمة نصية لرقم عشري | Safe float conversion."""
    try:
        s = str(v).strip().replace(",", "").replace("%", "")
        return float(s) if s not in ("", "nan", "none") else default
    except Exception:
        return default

@st.cache_data(ttl=300)
def get_ads_map():
    data = safe_get_all_values(ads_sheet)
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

# ══ خريطة العمولة/التوصيل (SKU -> {delivery, commission_pct}) | Commission & delivery map ══
@st.cache_data(ttl=300)
def get_com_map():
    data = safe_get_all_values(com_sheet)
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

def tacweed_badge(sku_up):
    """يرجع HTML صغير للكود 01 لو موجود لل SKU ده، وإلا يرجع سلسلة فاضية."""
    code = get_tacweed_map().get(sku_up, "")
    if not code:
        return ""
    return f'<span class="wh-badge" style="background:#3b0764;color:#e9d5ff;">🏷️ تكويد: {code}</span>'

def render_tacweed_upload(key_prefix):
    """واجهة رفع ملف التكويد (تكتشف شيت Noon تلقائيًا وتستخرج SKU + الكود 01)."""
    with st.expander("📤 رفع ملف التكويد | Upload Tacweed (Code Mapping) File", expanded=False):
        upl_tc = st.file_uploader(
            "ارفع ملف التكويد (xlsx/xlsm/csv) | Upload Tacweed File",
            type=["xlsx", "xls", "xlsm", "csv"], key=f"{key_prefix}_tacweed_upload")
        if upl_tc:
            try:
                if upl_tc.name.lower().endswith(".csv"):
                    df_tc = pd.read_csv(upl_tc, dtype=str).fillna("")
                else:
                    xls = pd.ExcelFile(upl_tc)
                    sheet_names = xls.sheet_names
                    chosen_sheet = None
                    for sn in sheet_names:
                        if "noon" in sn.strip().lower():
                            chosen_sheet = sn
                            break
                    if not chosen_sheet:
                        if len(sheet_names) > 1:
                            chosen_sheet = st.selectbox(
                                "اختار الشيت اللي فيه التكويد | Choose the sheet with the mapping",
                                sheet_names, key=f"{key_prefix}_tacweed_sheet_pick")
                        else:
                            chosen_sheet = sheet_names[0]
                    df_tc = pd.read_excel(upl_tc, sheet_name=chosen_sheet, dtype=str).fillna("")

                sku_col_tc = code_col_tc = None
                for c in df_tc.columns:
                    cl = str(c).strip().lower()
                    if sku_col_tc is None and "sku" in cl:
                        sku_col_tc = c
                    if code_col_tc is None and (("كود" in str(c) and "01" in str(c)) or cl in ("code01", "code 01")):
                        code_col_tc = c
                if not sku_col_tc:
                    sku_col_tc = df_tc.columns[0]
                if not code_col_tc:
                    # fallback: أقرب عمود اسمه فيه "كود"
                    for c in df_tc.columns:
                        if "كود" in str(c):
                            code_col_tc = c
                            break

                st.info(f"📊 {len(df_tc)} صف | SKU: `{sku_col_tc}` | الكود 01: `{code_col_tc}`")
                st.dataframe(df_tc[[sku_col_tc] + ([code_col_tc] if code_col_tc else [])].head(10),
                             use_container_width=True, height=180)

                def do_upload_tc(replace=True):
                    dn = now_str()
                    to_add = []
                    for _, row in df_tc.iterrows():
                        sku = str(row[sku_col_tc]).strip()
                        code = str(row[code_col_tc]).strip() if code_col_tc else ""
                        if sku and sku.lower() != "nan":
                            to_add.append([sku, code, dn])
                    if replace:
                        safe_delete_all(tacweed_sheet)
                    safe_batch_append(tacweed_sheet, to_add)
                    clear_cache(tacweed_sheet)
                    get_tacweed_map.clear()
                    return len(to_add)

                ca_tc, cb_tc = st.columns(2)
                with ca_tc:
                    if st.button("📤 إضافة للموجود | Append", type="primary",
                                 use_container_width=True, key=f"{key_prefix}_tc_append"):
                        n = do_upload_tc(replace=False)
                        st.success(f"✅ أُضيف {n} صف | rows added")
                        st.rerun()
                with cb_tc:
                    if st.button("🔄 استبدال الكل | Replace All", type="secondary",
                                 use_container_width=True, key=f"{key_prefix}_tc_replace"):
                        st.session_state[f"{key_prefix}_confirm_replace_tc"] = True
                if st.session_state.get(f"{key_prefix}_confirm_replace_tc"):
                    st.warning("⚠️ هيمسح كل التكويد القديم ويرفع الجديد؟ | Replace all existing mapping?")
                    cy_tc, cn_tc = st.columns(2)
                    if cy_tc.button("✅ نعم | Yes", key=f"{key_prefix}_yes_rep_tc"):
                        n = do_upload_tc(replace=True)
                        st.session_state[f"{key_prefix}_confirm_replace_tc"] = False
                        st.success(f"✅ تم الاستبدال — {n} صف")
                        st.rerun()
                    if cn_tc.button("❌ لا | No", key=f"{key_prefix}_no_rep_tc"):
                        st.session_state[f"{key_prefix}_confirm_replace_tc"] = False
                        st.rerun()
            except Exception as e:
                st.error(f"❌ {e}")

# ══ warehouse stock map (Code -> Quantity) — مربوط بالكود 01 من التكويد ══
@st.cache_data(ttl=300)
def get_warehouse_stock_map():
    data = safe_get_all_values(warehouse_stock_sheet)
    m = {}
    for row in data[1:]:
        if len(row) >= 2 and row[0].strip():
            m[row[0].strip()] = row[1].strip()
    return m

def warehouse_available_badge(sku_up):
    """يعرض الكود 01 + الكمية المتوفرة بالمستودع (لو الكود موجود ف ملف المستودع) جنب الSKU."""
    code = get_tacweed_map().get(sku_up, "")
    if not code:
        return ""
    qty = get_warehouse_stock_map().get(code, "")
    if qty == "":
        return f'<span class="tacweed-badge" style="background:#3b0764;color:#e9d5ff;">🏷️ تكويد: {code}</span>'
    return (
        f'<span class="tacweed-badge" style="background:#3b0764;color:#e9d5ff;">🏷️ تكويد: {code}</span> '
        f'<span class="tacweed-badge" style="background:#78350f;color:#fde68a;">📦 المتوفر بالمستودع: {qty}</span>'
    )

def render_warehouse_stock_upload(key_prefix):
    """واجهة رفع ملف جرد المستودع (بيتكشف شيت فيه عمود باركود وعمود كمية إجمالية تلقائيًا)."""
    with st.expander("📤 رفع ملف جرد المستودع | Upload Warehouse Stock File", expanded=False):
        upl_ws = st.file_uploader(
            "ارفع ملف جرد المستودع (xlsx/xls/csv) | Upload Warehouse Stock File",
            type=["xlsx", "xls", "xlsm", "csv"], key=f"{key_prefix}_wstock_upload")
        if upl_ws:
            try:
                if upl_ws.name.lower().endswith(".csv"):
                    df_ws = pd.read_csv(upl_ws, dtype=str).fillna("")
                else:
                    xls_ws = pd.ExcelFile(upl_ws)
                    sheet_names_ws = xls_ws.sheet_names
                    chosen_sheet_ws = None
                    for sn in sheet_names_ws:
                        try:
                            hdr = pd.read_excel(upl_ws, sheet_name=sn, dtype=str, nrows=0).columns
                        except Exception:
                            continue
                        if any("باركود" in str(c) or "barcode" in str(c).lower() for c in hdr):
                            chosen_sheet_ws = sn
                            break
                    if not chosen_sheet_ws:
                        if len(sheet_names_ws) > 1:
                            chosen_sheet_ws = st.selectbox(
                                "اختار الشيت اللي فيه بيانات الجرد | Choose the sheet with the stock data",
                                sheet_names_ws, key=f"{key_prefix}_wstock_sheet_pick")
                        else:
                            chosen_sheet_ws = sheet_names_ws[0]
                    df_ws = pd.read_excel(upl_ws, sheet_name=chosen_sheet_ws, dtype=str).fillna("")

                code_col_ws = qty_col_ws = name_col_ws = None
                for c in df_ws.columns:
                    cs = str(c).strip()
                    if code_col_ws is None and ("باركود" in cs or "barcode" in cs.lower()):
                        code_col_ws = c
                    if qty_col_ws is None and ("الكمية الإجمالية" in cs or "الكمية الاجمالية" in cs):
                        qty_col_ws = c
                    if name_col_ws is None and ("اسم المادة" in cs or "اسم المنتج" in cs):
                        name_col_ws = c
                if not code_col_ws:
                    code_col_ws = df_ws.columns[0]
                if not qty_col_ws:
                    for c in df_ws.columns:
                        if "كمية" in str(c) and "فعلي" not in str(c):
                            qty_col_ws = c
                            break

                st.info(f"📊 {len(df_ws)} صف | الباركود: `{code_col_ws}` | الكمية: `{qty_col_ws}`" + (f" | الاسم: `{name_col_ws}`" if name_col_ws else ""))
                preview_cols = [code_col_ws] + ([qty_col_ws] if qty_col_ws else []) + ([name_col_ws] if name_col_ws else [])
                st.dataframe(df_ws[preview_cols].head(10), use_container_width=True, height=180)

                def do_upload_ws(replace=True):
                    dn = now_str()
                    to_add = []
                    for _, row in df_ws.iterrows():
                        code = str(row[code_col_ws]).strip()
                        qty = str(row[qty_col_ws]).strip() if qty_col_ws else ""
                        name = str(row[name_col_ws]).strip() if name_col_ws else ""
                        if code and code.lower() != "nan":
                            to_add.append([code, qty, name, dn])
                    if replace:
                        safe_delete_all(warehouse_stock_sheet)
                    safe_batch_append(warehouse_stock_sheet, to_add)
                    clear_cache(warehouse_stock_sheet)
                    get_warehouse_stock_map.clear()
                    return len(to_add)

                ca_ws, cb_ws = st.columns(2)
                with ca_ws:
                    if st.button("📤 إضافة للموجود | Append", type="primary",
                                 use_container_width=True, key=f"{key_prefix}_ws_append"):
                        n = do_upload_ws(replace=False)
                        st.success(f"✅ أُضيف {n} صف | rows added")
                        st.rerun()
                with cb_ws:
                    if st.button("🔄 استبدال الكل | Replace All", type="secondary",
                                 use_container_width=True, key=f"{key_prefix}_ws_replace"):
                        st.session_state[f"{key_prefix}_confirm_replace_ws"] = True
                if st.session_state.get(f"{key_prefix}_confirm_replace_ws"):
                    st.warning("⚠️ هيمسح كل بيانات المستودع القديمة ويرفع الجديد؟ | Replace all existing warehouse stock?")
                    cy_ws, cn_ws = st.columns(2)
                    if cy_ws.button("✅ نعم | Yes", key=f"{key_prefix}_yes_rep_ws"):
                        n = do_upload_ws(replace=True)
                        st.session_state[f"{key_prefix}_confirm_replace_ws"] = False
                        st.success(f"✅ تم الاستبدال — {n} صف")
                        st.rerun()
                    if cn_ws.button("❌ لا | No", key=f"{key_prefix}_no_rep_ws"):
                        st.session_state[f"{key_prefix}_confirm_replace_ws"] = False
                        st.rerun()
            except Exception as e:
                st.error(f"❌ {e}")

# ══ inv_map ══
def _to_int(v):
    try:
        return int(float(str(v).replace(",","")))
    except:
        return 0

def build_inv_map(excluded_wh: set):
    inv_data = get_cached(inventory_sheet)
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

# ══ Sheets helpers ══
def safe_append(sheet, row, retries=5, delay=1):
    for attempt in range(retries):
        try:
            sheet.append_row(row, value_input_option="USER_ENTERED")
            clear_cache(sheet)
            return True
        except gspread.exceptions.APIError as e:
            if "429" in str(e) or "quota" in str(e).lower():
                time.sleep(delay * (2 ** attempt))
            else:
                time.sleep(delay)
        except Exception:
            time.sleep(delay)
    return False

def safe_delete(sheet, row_idx, retries=5, delay=1):
    for attempt in range(retries):
        try:
            sheet.delete_rows(row_idx)
            clear_cache(sheet)
            return True
        except gspread.exceptions.APIError as e:
            if "429" in str(e) or "quota" in str(e).lower():
                time.sleep(delay * (2 ** attempt))
            else:
                time.sleep(delay)
        except Exception:
            time.sleep(delay)
    return False

def safe_delete_all(sheet):
    try:
        data = sheet.get_all_values()
        if len(data) > 1:
            sheet.delete_rows(2, len(data))
        clear_cache(sheet)
        return True
    except Exception:
        return False

def safe_batch_append(sheet, rows_data, retries=5, delay=1):
    if not rows_data:
        return True
    for attempt in range(retries):
        try:
            sheet.append_rows(rows_data, value_input_option="USER_ENTERED")
            clear_cache(sheet)
            return True
        except gspread.exceptions.APIError as e:
            if "429" in str(e) or "quota" in str(e).lower():
                wait = delay * (2 ** attempt)
                st.toast(f"⏳ Google Sheets API limit — waiting {wait}s...", icon="⏳")
                time.sleep(wait)
            else:
                time.sleep(delay)
        except Exception:
            time.sleep(delay)
    return False

def safe_update_row(sheet, row_idx, values, retries=4, delay=1):
    for _ in range(retries):
        try:
            for ci, val in enumerate(values, start=1):
                sheet.update_cell(row_idx, ci, val)
            clear_cache(sheet)
            return True
        except Exception:
            time.sleep(delay)
    return False

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def merge_or_get_existing_row(sheet, sku):
    """
    يبحث عن SKU في العمود الأول لشيت معيّن.
    بيرجع (row_index, row_values) لو لقاه، أو (None, None) لو مش موجود.
    """
    data = get_cached(sheet, force=True)
    sku_up = sku.strip().upper()
    if len(data) > 1:
        for ri, row in enumerate(data[1:], start=2):
            if row and row[0].strip().upper() == sku_up:
                return ri, row
    return None, None

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

def file_timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def to_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()

def make_empty_template(columns):
    return to_excel(pd.DataFrame(columns=columns))

def parse_excel_date(val):
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

def dl_btn(df, prefix, label="⬇️ Excel | Download", key=None):
    st.download_button(label, data=to_excel(df),
        file_name=f"{prefix}_{file_timestamp()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key=key or f"dlbtn_{prefix}")

# ══ CancelNotifications — حفظ/تحميل/تنظيف ══
def load_cancel_notifications():
    """تحمّل الإشعارات من Sheets وتحذف المنتهية تلقائياً"""
    data = get_cached(cancel_notif_sheet, force=False)
    today = datetime.now().date()
    notifs = []
    rows_to_delete = []
    if len(data) <= 1:
        return notifs
    for i, row in enumerate(data[1:], start=2):
        while len(row) < 5: row.append("")
        asn, skus_str, sdate, reason, ts = row[0], row[1], row[2], row[3], row[4]
        if not asn.strip():
            continue
        # تحقق من انتهاء تاريخ الجدولة
        pd_ = parse_excel_date(sdate)
        if pd_ and today > pd_.date():
            rows_to_delete.append(i)
            continue
        notifs.append({
            "asn":    asn.strip(),
            "skus":   [s.strip() for s in skus_str.split("|") if s.strip()],
            "sdate":  sdate.strip(),
            "reason": reason.strip(),
            "ts":     ts.strip(),
        })
    # حذف المنتهية من الشيت
    for idx in sorted(rows_to_delete, reverse=True):
        safe_delete(cancel_notif_sheet, idx)
    if rows_to_delete:
        clear_cache(cancel_notif_sheet)
    return notifs

def save_cancel_notification(asn, skus_list, sdate, reason, ts):
    """يحفظ إشعار كنسل جديد في Sheets"""
    skus_str = "|".join(skus_list)
    safe_append(cancel_notif_sheet, [asn, skus_str, sdate, reason, ts])
    clear_cache(cancel_notif_sheet)

def delete_cancel_notification_by_asn(asn):
    """يحذف إشعار معين بالـ ASN"""
    data = get_cached(cancel_notif_sheet, force=True)
    for i, row in enumerate(data[1:], start=2):
        if row and row[0].strip().upper() == asn.strip().upper():
            safe_delete(cancel_notif_sheet, i)
            clear_cache(cancel_notif_sheet)
            return

def delete_all_cancel_notifications():
    """يمسح كل الإشعارات"""
    safe_delete_all(cancel_notif_sheet)
    clear_cache(cancel_notif_sheet)

def check_expired_scheduled():
    data = get_scheduled_normalized(force=True)
    if len(data) <= 1:
        return
    today = datetime.now().date()
    expired_rows = []
    del_idx = []
    for i, row in enumerate(data[1:], start=2):
        while len(row) < 8: row.append("")
        d = parse_excel_date(row[3])
        # منتهي = تاريخ الجدولة < اليوم (مش <= عشان نفس اليوم ما يتشالش)
        if d and d.date() < today:
            expired_rows.append(row[:7] + [now_str()])
            del_idx.append(i)
    if expired_rows:
        safe_batch_append(expired_sheet, expired_rows)
        # نحذف من الأسفل للأعلى عشان الـ index ما يتغيرش
        for idx in sorted(del_idx, reverse=True):
            safe_delete(scheduled_sheet, idx)
        clear_cache(scheduled_sheet)

# ══ CSS ══
st.markdown("""
<style>
.stTabs [data-baseweb="tab-list"]{gap:5px;flex-wrap:wrap;}
.stTabs [data-baseweb="tab"]{background:#1e293b;color:white;border-radius:8px;padding:6px 12px;font-weight:bold;font-size:11px;}
.stTabs [aria-selected="true"]{background:#3b82f6!important;}
.wh-badge{display:inline-block;border-radius:6px;padding:2px 9px;margin:2px;font-size:12px;}
.tacweed-badge{display:inline-block;border-radius:8px;padding:6px 14px;margin:4px 3px;font-size:17px;font-weight:bold;}
.status-badge-lg{display:inline-block;border-radius:8px;padding:6px 14px;margin:4px 3px;font-size:16px;font-weight:bold;color:#000!important;}
.cancel-notif-card{
    background: linear-gradient(135deg,#2d0a0a,#1a0000);
    border: 1px solid #ef4444;
    border-left: 5px solid #ef4444;
    border-radius:10px;
    padding:10px 14px;
    margin-bottom:8px;
    color:white;
}
.cancel-notif-card .asn-num{font-size:16px;font-weight:bold;color:#fca5a5;}
.cancel-notif-card .sku-chip{
    display:inline-block;
    background:#4b1010;
    color:#fca5a5;
    border-radius:5px;
    padding:1px 7px;
    margin:2px;
    font-size:11px;
}
.cancel-notif-card .reason-text{color:#fcd34d;font-size:12px;}
</style>
""", unsafe_allow_html=True)

# ══ Init ══
# نشغّل فحص المنتهيات كل يوم (مش بس أول تشغيل) — لو بقينا في يوم جديد نعيد الفحص
_today_key = f"expired_checked_{datetime.now().date()}"
if _today_key not in st.session_state:
    # امسح مفاتيح الأيام القديمة
    for _old_key in [k for k in st.session_state if k.startswith("expired_checked_") and k != _today_key]:
        del st.session_state[_old_key]
    check_expired_scheduled()
    st.session_state[_today_key] = True

# تحميل الإشعارات من Sheets عند كل تشغيل (تدوم بعد الإغلاق)
if "cancel_notifs_loaded" not in st.session_state:
    st.session_state["check_cancel_notifications"] = load_cancel_notifications()
    st.session_state["cancel_notifs_loaded"] = True
elif "check_cancel_notifications" not in st.session_state:
    st.session_state["check_cancel_notifications"] = []

excluded_wh = get_excluded_warehouses()
inv_map     = build_inv_map(excluded_wh)

# ══════════════════════════════════════════════
# ══ SIDEBAR — إشعارات الكنسل ══
# ══════════════════════════════════════════════
def render_sidebar_notifications():
    notifs = st.session_state.get("check_cancel_notifications", [])
    if not notifs:
        return
    with st.sidebar:
        st.markdown("## 🔔 إشعارات الكنسل | Cancel Alerts")
        st.markdown(f"**{len(notifs)} إشعار نشط | Active Alerts**")
        st.markdown("---")
        links_map_sb = get_links_map()
        for ni, notif in enumerate(notifs):
            asn   = notif.get("asn","")
            sdate = notif.get("sdate","")
            skus  = notif.get("skus",[])
            reason= notif.get("reason","")
            ts    = notif.get("ts","")

            # بناء الـ SKU chips مع الصور
            sku_chips_html = ""
            for sk in skus[:5]:
                sku_chips_html += f'<span class="sku-chip">{sk}</span>'
            if len(skus) > 5:
                sku_chips_html += f'<span class="sku-chip">+{len(skus)-5} more</span>'

            st.markdown(f"""
<div class="cancel-notif-card">
  <div>🚫 <span class="asn-num">ASN: {asn}</span></div>
  <div style="font-size:12px;color:#94a3b8;">📅 {sdate}</div>
  <div style="margin:6px 0;">{sku_chips_html}</div>
  <div class="reason-text">📝 {reason if reason else '—'}</div>
  <div style="font-size:10px;color:#64748b;margin-top:4px;">🕐 {ts}</div>
</div>
""", unsafe_allow_html=True)

            # عرض الصور بشكل مصغر
            img_cols = st.columns(min(len(skus[:4]), 4))
            for ci2, sk in enumerate(skus[:4]):
                img_url = links_map_sb.get(sk.strip().upper(), "")
                with img_cols[ci2]:
                    if img_url and img_url.startswith("http"):
                        st.image(img_url, width=55, caption=sk[:8])
                    else:
                        st.markdown(f"🖼️ `{sk[:8]}`")

            if st.button(f"✖️ حذف | Remove #{ni+1}", key=f"sb_rm_notif_{ni}", use_container_width=True):
                delete_cancel_notification_by_asn(notif.get("asn",""))
                st.session_state["check_cancel_notifications"].pop(ni)
                st.rerun()
            st.markdown("---")

        if st.button("🗑️ مسح كل الإشعارات | Clear All", key="sb_clear_all_notifs",
                     use_container_width=True, type="secondary"):
            delete_all_cancel_notifications()
            st.session_state["check_cancel_notifications"] = []
            st.rerun()

render_sidebar_notifications()

# ══ UI helpers ══
def show_img(img, width=75):
    if img and str(img).startswith("http"):
        st.image(img, width=width)
    else:
        st.markdown("🖼️")

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
    """SKU كـ لينك قابل للضغط يودّي على صفحة المنتج على نون في تاب جديد
    | SKU rendered as a clickable link that opens the noon.com product page in a new tab."""
    link = build_noon_link(sku)
    if not link:
        return f"`{sku}`"
    return (f'<a href="{link}" target="_blank" rel="noopener" '
            f'style="font-family:monospace;font-weight:700;color:#3b82f6;text-decoration:none;'
            f'background:#1e293b;border:1px solid #334155;border-radius:6px;padding:2px 10px;{extra_style}">'
            f'{sku} 🔗</a>')

def show_sku_inv(sku: str):
    info = inv_map.get(sku.strip().upper())
    if not info:
        return
    total = info["total_stock"]
    sales = info["sales"]
    st.markdown(f"📈 **مبيع شهري | Monthly Sales:** **{sales}** &nbsp;|&nbsp; 📦 **مخزون | Stock:** **{total}**")
    badges = []
    for wh, stk in sorted(info["warehouses"].items()):
        is_ex  = wh.upper() in excluded_wh
        bg     = "#4b1010" if is_ex else "#1e3a5f"
        color  = "#fca5a5" if is_ex else "#93c5fd"
        strike = "text-decoration:line-through;" if is_ex else ""
        badges.append(f'<span class="wh-badge" style="background:{bg};color:{color};{strike}">{wh}: {stk}</span>')
    st.markdown("🏭 " + "".join(badges), unsafe_allow_html=True)

def confirm_clear(key, sheet, label=""):
    if st.session_state.get(f"confirm_{key}"):
        st.warning(f"⚠️ مسح كل {label}؟ | Clear all {label}?")
        cy, cn = st.columns(2)
        if cy.button("✅ نعم | Yes", key=f"yes_{key}"):
            safe_delete_all(sheet)
            st.session_state[f"confirm_{key}"] = False
            st.success("✅ تم المسح | Cleared")
            st.rerun()
        if cn.button("❌ لا | No", key=f"no_{key}"):
            st.session_state[f"confirm_{key}"] = False
            st.rerun()

# ══════════════════════════════════════════════
# ══ مراجعة المخزون / مراجعة المبيعات — نفس منطق استعلامي Access ══
# ══════════════════════════════════════════════
def build_daily_orders_map(target_date):
    """يرجع dict: sku_upper -> عدد صفوف الأوردرز لليوم المحدد (= Sum(QTY) بافتراض كل صف = قطعة واحدة)."""
    data = get_cached(daily_orders_sheet)
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

def build_daily_orders_counts(dates):
    """يرجع dict: sku_upper -> {date: عدد} لقائمة تواريخ محددة (مرور واحد على الشيت بدل تكرار لكل تاريخ)."""
    data = get_cached(daily_orders_sheet)
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

def build_daily_orders_prices(dates):
    """يرجع dict: sku_upper -> {date: [(qty, price), ...]} لعرض تفاصيل الأسعار اليومية."""
    data = get_cached(daily_orders_sheet)
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

# ══ الأقسام | Departments (عمود Family في DailyOrders — اختياري، لو مش موجود الكود بيكمل عادي) ══
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

def build_daily_orders_family_stats(dates):
    """يرجع dict: اسم القسم -> {"orders": عدد, "revenue": إيراد} لفترة تواريخ محددة.
    لو عمود Family مش موجود في الشيت، أو الصف مالوش قيمة Family، بيتجاهله بهدوء
    بدون ما يوقف الكود أو يأثر على أي تحليل تاني."""
    data = get_cached(daily_orders_sheet)
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
        if price_val and price_val.lower() not in ("", "nan", "none"):
            try:
                rev_val = float(price_val.replace(",", "")) * qty_val
            except Exception:
                rev_val = 0.0
        if dept_name not in stats:
            stats[dept_name] = {"orders": 0, "revenue": 0.0}
        stats[dept_name]["orders"]  += 1
        stats[dept_name]["revenue"] += rev_val
    return stats

def is_sku_only_in_excluded_warehouses(sku_up, excluded_wh):
    """يتحقق هل كل مستودعات هذا الـ SKU (في ملف المخزون) هي مستودعات مستثناة فقط —
    لو كذلك، يبقى مينفعش يظهر في تابي مراجعة المخزون / مراجعة المبيعات."""
    if not excluded_wh:
        return False
    info = inv_map.get(sku_up)
    if not info:
        return False
    whs = info.get("warehouses", {})
    if not whs:
        return False
    return all(w.strip().upper() in excluded_wh for w in whs.keys())

def compute_stock_sales_rows(target_date, display_dates=None):
    """يحسب لكل SKU ظهر في أوردرز اليوم المحدد نفس مخرجات استعلامي مراجعة مخزون / مراجعة مبيعات.
    display_dates (اختياري): قائمة تواريخ إضافية تتعرض جنب كل SKU (مثلاً أمس/أول أمس/أول أول أمس)."""
    daily_qty = build_daily_orders_map(target_date)
    display_dates = display_dates or [target_date]
    multi_counts = build_daily_orders_counts(display_dates)
    rows = []
    for sku_up, qty in daily_qty.items():
        if is_sku_only_in_excluded_warehouses(sku_up, excluded_wh):
            continue
        info        = inv_map.get(sku_up, {})
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

def compute_missing_inventory_rows(display_dates):
    """SKUs ظهرت في الأوردرز خلال آخر كذا يوم (أمس/أول أمس/أول أول أمس) لكن مالهاش سجل في شيت Inventory
    — يعني مخزونها انتهى بالكامل وخرجت من ملف المخزون. تظهر بنفس تفاصيل تابي المراجعة."""
    multi_counts = build_daily_orders_counts(display_dates)
    links_map_local = get_links_map()
    rows = []
    for sku_up, day_counts in multi_counts.items():
        if sku_up in inv_map:
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

def render_day_counts_md(day_counts, dates, labels):
    """يبني سطر Markdown بمبيعات كل يوم من التواريخ المعطاة بجانب بعض."""
    parts = [f"**{lbl}:** {day_counts.get(d,0)}" for d, lbl in zip(dates, labels)]
    return " &nbsp;|&nbsp; ".join(parts)

def get_recent_expired_info(sku, days_back=4):
    """يدوّر على SKU في شيت Expired (الجدولة منتهية الصلاحية) خلال آخر days_back يوم — حسب تاريخ الانتهاء.
    يرجع أحدث سجل لو لقى، أو None لو مفيش جدولة منتهية مؤخراً لهذا الـ SKU."""
    sku_up = sku.strip().upper()
    data = get_cached(expired_sheet)
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

def render_recent_expired_note(sku, days_back=4):
    """يعرض ملاحظة لو الـ SKU كانت ليه جدولة انتهت خلال آخر days_back يوم."""
    info = get_recent_expired_info(sku, days_back)
    if not info:
        return
    st.markdown(
        f'<span style="background:#7c2d12;color:#fed7aa;border-radius:6px;padding:3px 10px;font-size:12px;">'
        f'📋 كانت مجدولة (ASN {info["asn"]}) بتاريخ {info["schedule_date"]} وانتهت الجدولة بتاريخ {info["date_expired"]} | '
        f'Was scheduled but expired</span>',
        unsafe_allow_html=True)

def get_recent_schedule_rows(days_back=4):
    """يرجع dict: sku_upper -> أحدث سجل جدولة (من Scheduled/Check/Expired) وقع تاريخ جدولته
    خلال آخر days_back يوم (يعني: أمس/أول أمس/أول أول أمس/قبل 4 أيام).
    الهدف: الـ SKU ده لسه في فترة انتظار وصول المخزون بعد الجدولة — حتى لو الجدولة
    خلاص اتنقلت لتاب Expired — فلازم يفضل ظاهر في تابات المراجعة عشان محدش يطلبه تاني بالغلط."""
    cutoff = datetime.now().date() - timedelta(days=days_back)
    today_ = datetime.now().date()
    src_label_map = {"Scheduled": "الجدولة | Scheduled", "Check": "تشييك | Check", "Expired": "منتهية | Expired"}
    by_sku = {}
    for sheet_key in ("Scheduled", "Check", "Expired"):
        data = get_scheduled_normalized() if sheet_key == "Scheduled" else get_cached(sheets[sheet_key])
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

def recent_schedule_badge_html(entry):
    """شارة توضيحية لسكو اتجدول (أو انتهت جدولته) خلال آخر 4 أيام — تمنع إعادة الطلب بالغلط."""
    color = "#7c3aed" if entry["source"] != "Expired" else "#b45309"
    return (
        f'<span class="status-badge-lg" style="background:{color};">'
        f'📅 مجدول بتاريخ {entry["date"]} (ASN {entry["asn"]}) [{entry["source_label"]}] — '
        f'خلال آخر 4 أيام، لسه في فترة الوصول — لا تطلبه تاني | '
        f'Scheduled within the last 4 days — still within arrival window, don\'t re-order</span>'
    )

def get_pending_approval_skus(force=False):
    """يرجع set بكل الـ SKUs الموجودة حالياً في تاب PendingApproval — يعني لسه مستنية اعتماد
    الجدولة من برنامج الجدولة المكتبي (schedule_entry_app) وما اتنقلتش لتاب Scheduled لسه.
    القراءة حسب اسم عمود SKU مش مكانه، عشان تفضل شغالة أياً كان ترتيب الأعمدة الحقيقي في الشيت."""
    raw = get_cached(pending_approval_sheet, force=force)
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

def pending_approval_badge_html():
    """شارة توضيحية إن السكو موجود في تاب 'قيد الموافقة' — يعني فيه جدولة مقترحة لسه مستنية
    اعتماد من برنامج الجدولة المكتبي، عشان محدش يطلبه أو يجدوله تاني بالغلط."""
    return (
        '<span class="status-badge-lg" style="background:#b45309;">'
        '⏳ قيد الموافقة | Pending Approval — في انتظار اعتماد الجدولة من برنامج الجدولة المكتبي، '
        'لا تطلبه أو تجدوله تاني | Awaiting schedule approval from the desktop scheduling app — '
        'don\'t re-order or re-schedule</span>'
    )

def compute_recent_scheduled_rows(exclude_skus, day_dates, days_back=4):
    """يبني صفوف SKUs اتجدولت أو انتهت جدولتها خلال آخر days_back يوم ومش ظاهرة أصلاً
    في قائمة المراجعة الرئيسية (exclude_skus) — عشان تتعرض في سكشن منفصل يفكّر المستخدم
    إنها كانت مجدولة مؤخراً ولسه بتستنى توصل."""
    recent_map = get_recent_schedule_rows(days_back=days_back)
    if not recent_map:
        return []
    day_counts_map = build_daily_orders_counts(day_dates)
    rows = []
    for sku_up, sched_entry in recent_map.items():
        if sku_up in exclude_skus:
            continue
        info        = inv_map.get(sku_up, {})
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

def render_recent_scheduled_section(rows, day_dates, day_labels, dl_key):
    """يعرض سكشن 'مجدولة خلال آخر 4 أيام' في تابات المراجعة — تنبيه لمنع إعادة الطلب بالغلط."""
    st.divider()
    st.subheader("📅 مجدولة خلال آخر 4 أيام | Recently Scheduled (Last 4 Days)")
    st.caption(
        "SKUs اتجدولت أو انتهت جدولتها خلال آخر 4 أيام ولسه في فترة انتظار وصول المخزون — "
        "بتظهر هنا حتى لو مش محتاجة مراجعة دلوقتي، عشان محدش يطلبها تاني بالغلط | "
        "SKUs scheduled (or whose schedule expired) in the last 4 days and still within the "
        "arrival window — shown here even if not currently flagged, so no one re-requests them by mistake"
    )
    if not rows:
        st.info("لا يوجد SKUs مجدولة مؤخراً غير ظاهرة أعلاه | No recently scheduled SKUs outside the list above")
        return
    df_rs = pd.DataFrame([{
        "SKU": r["sku"], "Stock": r["stock"], "Monthly Sales": r["sales_month"],
        "Schedule Date": r["sched"]["date"], "Source": r["sched"]["source"], "ASN": r["sched"]["asn"],
    } for r in rows])
    c1, c2 = st.columns(2)
    with c1: dl_btn(df_rs, dl_key, key=f"dlbtn_{dl_key}")
    with c2: st.info(f"📅 SKUs مجدولة مؤخراً | Recently scheduled: {len(rows)}")
    for r in rows:
        c_img, c_info = st.columns([1, 6])
        with c_img: show_img(r["img"], 70)
        with c_info:
            st.markdown(f"**SKU:** `{r['sku']}`")
            tc_badge_rs = warehouse_available_badge(r["sku_up"])
            if tc_badge_rs:
                st.markdown(tc_badge_rs, unsafe_allow_html=True)
            st.markdown(f"📦 **المخزون | Stock:** {r['stock']} &nbsp;|&nbsp; 📈 **مبيع شهري | Monthly:** {r['sales_month']}")
            st.markdown("🛒 " + render_day_counts_md(r["day_counts"], day_dates, day_labels))
            st.markdown(recent_schedule_badge_html(r["sched"]), unsafe_allow_html=True)
            for note in get_unavailable_ordered_note(r["sku"]):
                st.markdown(big_note_html(note), unsafe_allow_html=True)
        st.divider()

def get_latest_schedule_info(sku):
    """يدوّر على SKU في الجدولة والتشييك ويرجع أقرب جدولة (تاريخ) أو None."""
    sku_up = sku.strip().upper()
    candidates = []
    for sheet_key in ("Scheduled","Check"):
        data = get_scheduled_normalized() if sheet_key == "Scheduled" else get_cached(sheets[sheet_key])
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

def clear_unavailable_ordered_for_sku(sku):
    """يمسح أي سجل قديم لهذا الـ SKU من شيتات Unavailable و Ordered —
    بيتنفذ وقت الموافقة على طلب جديد لنفس الـ SKU، عشان ملاحظات
    'غير متوفر سابقاً' / 'تم الطلب سابقاً' متفضلش ظاهرة غلط في التابات
    (مراجعة المخزون، مراجعة المبيعات، الموافقة، ...) بعد ما بقى متاح فعلاً
    أو وصل طلبه."""
    if not sku or not str(sku).strip():
        return
    for sh in (unavailable_sheet, ordered_sheet):
        ri, _ = merge_or_get_existing_row(sh, sku)
        if ri:
            safe_delete(sh, ri)

def get_unavailable_ordered_note(sku):
    """لو الـ SKU سبق اتسجل غير متوفر أو تم طلبه، يرجع ملاحظات بالتواريخ."""
    sku_up = sku.strip().upper()
    notes = []
    data_un = get_cached(unavailable_sheet)
    if len(data_un) > 1:
        for row in data_un[1:]:
            if row and row[0].strip().upper() == sku_up:
                while len(row) < 5: row.append("")
                cnt, dates = parse_count_dates(row[4])
                notes.append(f"❌ غير متوفر سابقاً | Was unavailable ({cnt}x) — {dates}")
                break
    data_ord = get_cached(ordered_sheet)
    if len(data_ord) > 1:
        for row in data_ord[1:]:
            if row and row[0].strip().upper() == sku_up:
                while len(row) < 6: row.append("")
                cnt, dates = parse_count_dates(row[5])
                notes.append(f"🛒 تم طلبه سابقاً | Was ordered ({cnt}x) — {dates}")
                break
    return notes

def schedule_coverage_badge(sku, days_to_stockout, delay_days):
    """يرجع (نص الحالة, لون, معلومات الجدولة) حسب هل الجدولة هتوصل قبل نفاد المخزون أو لأ."""
    sched = get_latest_schedule_info(sku)
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

ordinal_map = {1:"الثانية|Second",2:"الثالثة|Third",3:"الرابعة|Fourth",4:"الخامسة|Fifth"}


# ══════════════════════════════════════════════
st.title("📦 Stock Requests | طلبات المخزون")

# ══ حساب المرحلين من المبيعات مسبقاً (قبل رسم التابات) ══
def compute_transferred_from_sales():
    """يحسب SKUs اللي هتتنقل لمراجعة المخزون (محتاج جدولة فقط، بدون أي جدولة أو ملاحظات)."""
    if not inv_map:
        return []
    settings_now = load_settings()
    sales_days_now  = int(settings_now.get("sales_display_days","7") or 7)
    delay_days_now  = int(settings_now.get("schedule_delay_days","3") or 3)
    cov_days_now    = int(settings_now.get("schedule_coverage_days","15") or 15)
    today_now = datetime.now().date()
    dates_now = [today_now - timedelta(days=i) for i in range(1, sales_days_now + 1)]
    counts_now  = build_daily_orders_counts(dates_now)
    result = []
    for sku_up, info in inv_map.items():
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
        badge_text, _, sched = schedule_coverage_badge(sku_disp, days_to_so, delay_days_now)
        un_notes = get_unavailable_ordered_note(sku_disp)
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

if "transferred_skus_t14" not in st.session_state:
    st.session_state["transferred_skus_t14"] = compute_transferred_from_sales()

tabs = st.tabs([
    "🛒 المبيعات | Sales",
    "📊 داشبورد المبيعات | Sales Dashboard",
    "📋 الطلبات | Requests",
    "✅ الموافقة | Approved",
    "❌ غير متوفر | Unavailable",
    "🛒 تم الطلب | Ordered",
    "📅 الجدولة | Scheduled",
    "☑️ تشييك | Check",
    "🚫 جدولة ملغية | Cancelled",
    "🔄 تعديل موعد | Rescheduled",
    "⚠️ تنبيهات | Alerts",
    "📊 المخزون | Inventory",
    "🔴 مراجعة المخزون | Stock Review",
    "🗂️ منتهية | Expired",
    "⚙️ الإعدادات | Settings",
    "📈 مراجعة المبيعات | Sales Review",
    "🗓️ تحليل الجدولة | Schedule Analysis",
    "📦 مخزون بدون بيع | No Sales",
])
(tab14,tab_dash,tab1,tab2,tab3,tab4,tab5,tab_check,tab6,tab7,tab8,tab9,tab10,tab11,tab12,tab13,tab15,tab16) = tabs

# ══ TAB 1 — الطلبات ══
with tab1:
    st.subheader("➕ إضافة طلبات | Add Requests")
    links_map = get_links_map()
    col_m, col_t = st.columns([3,1])
    with col_t:
        st.download_button("⬇️ Template فارغ | Empty Template",
            data=make_empty_template(["SKU","Quantity"]),
            file_name=f"request_template_{file_timestamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with col_m:
        method = st.radio("طريقة الإضافة | Add Method:", ["📂 رفع ملف | Upload","✏️ لصق | Paste"], horizontal=True)

    added_rows, file_name_label = [], ""
    if "Upload" in method:
        uploaded = st.file_uploader("ارفع Excel أو CSV | Upload Excel or CSV", type=["xlsx","xls","csv"])
        if uploaded:
            file_name_label = uploaded.name
            try:
                df_up = pd.read_csv(uploaded,dtype=str).fillna("") if uploaded.name.endswith(".csv") else pd.read_excel(uploaded,dtype=str).fillna("")
                sku_col = qty_col = None
                for c in df_up.columns:
                    cl = c.strip().lower()
                    if cl in ("sku","item","product","item nr","item_nr"): sku_col = c
                    if cl in ("quantity","qty","كمية","الكمية","amount"):  qty_col = c
                if not sku_col: sku_col = df_up.columns[0]
                if not qty_col and len(df_up.columns)>1: qty_col = df_up.columns[1]
                st.info(f"📊 {len(df_up)} صف | rows")
                st.dataframe(df_up[[c for c in [sku_col,qty_col] if c]], use_container_width=True, height=150)
                for _, row in df_up.iterrows():
                    sku = str(row[sku_col]).strip()
                    qty = str(row[qty_col]).strip() if qty_col else ""
                    img = links_map.get(sku.upper(),"")
                    if sku and sku.lower() != "nan":
                        added_rows.append((sku,qty,img))
            except Exception as e:
                st.error(f"❌ {e}")
    else:
        pasted = st.text_area("الصق هنا | Paste here (SKU,Qty):", height=110, placeholder="SKU001,5\nSKU002,3")
        file_name_label = "Manual Entry"
        if pasted.strip():
            for line in pasted.strip().splitlines():
                parts = [p.strip() for p in line.split(",")]
                sku = parts[0] if parts else ""
                qty = parts[1] if len(parts)>1 else ""
                img = links_map.get(sku.upper(),"")
                if sku: added_rows.append((sku,qty,img))
            if added_rows: st.success(f"✅ {len(added_rows)} صف جاهز | rows ready")

    if added_rows:
        if st.button("📤 إضافة | Add", type="primary"):
            dn = now_str()
            if safe_batch_append(requests_sheet, [[s,q,i,dn,file_name_label] for s,q,i in added_rows]):
                st.success(f"✅ أُضيف {len(added_rows)} صف | rows added")
                st.rerun()

    st.divider()
    st.subheader("📋 الطلبات الحالية | Current Requests")
    data = get_cached(requests_sheet)
    if len(data) <= 1:
        st.info("لا توجد طلبات | No requests yet.")
    else:
        rows = data[1:]
        df_req = pd.DataFrame(rows, columns=data[0])
        c1,c2,c3,c4 = st.columns(4)
        with c1: dl_btn(df_req,"requests")
        with c2:
            if st.button("✅ موافقة الكل | Approve All", use_container_width=True):
                st.session_state["confirm_approve_all"] = True
        with c3:
            if st.button("❌ رفض الكل | Reject All", use_container_width=True):
                st.session_state["confirm_reject_all"] = True
        with c4:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", use_container_width=True):
                st.session_state["confirm_clear_req"] = True

        if st.session_state.get("confirm_approve_all"):
            st.warning("⚠️ موافقة على كل الطلبات؟ | Approve all?")
            cy,cn = st.columns(2)
            if cy.button("✅ نعم | Yes", key="yes_app_all"):
                dn = now_str()
                safe_batch_append(approved_sheet, [[r[0],r[1],r[1],r[2] if len(r)>2 else "",r[3] if len(r)>3 else "",dn] for r in rows])
                for r in rows:
                    clear_unavailable_ordered_for_sku(r[0])
                safe_delete_all(requests_sheet)
                st.session_state["confirm_approve_all"] = False
                st.rerun()
            if cn.button("❌ لا | No", key="no_app_all"):
                st.session_state["confirm_approve_all"] = False
                st.rerun()

        if st.session_state.get("confirm_reject_all"):
            st.warning("⚠️ رفض كل الطلبات؟ | Reject all?")
            cy,cn = st.columns(2)
            if cy.button("✅ نعم | Yes", key="yes_rej_all"):
                dn = now_str()
                safe_batch_append(unavailable_sheet, [[r[0],r[1],r[2] if len(r)>2 else "",r[3] if len(r)>3 else "",dn] for r in rows])
                safe_delete_all(requests_sheet)
                st.session_state["confirm_reject_all"] = False
                st.rerun()
            if cn.button("❌ لا | No", key="no_rej_all"):
                st.session_state["confirm_reject_all"] = False
                st.rerun()

        confirm_clear("clear_req", requests_sheet, "الطلبات | Requests")

        ordered_data = get_cached(ordered_sheet)
        ordered_skus = {}
        if len(ordered_data) > 1:
            for r in ordered_data[1:]:
                while len(r) < 6: r.append("")
                sk = r[0].strip().upper()
                ordered_skus[sk] = _to_int(r[4]) if r[4] else 1

        st.write(f"**الإجمالي | Total: {len(rows)}**")

        # ══ تحديد بالعلامة (✓) جمب كل منتج زي المثال | Checkbox next to each item ══
        row_indices = [i for i in range(2, len(rows) + 2)]

        # لو فيه طلب تصفير من إجراء سابق، نفّذه هنا قبل ما أي checkbox يتعمل له render في الـ run ده
        if st.session_state.pop("_req_clear_pending", False):
            for k in [k for k in list(st.session_state.keys()) if k.startswith("chk_req_")]:
                del st.session_state[k]

        # تحديد الكل
        sel_all_col, _ = st.columns([1, 5])
        with sel_all_col:
            select_all = st.checkbox("تحديد الكل | Select All", key="chk_req_select_all")
        if select_all:
            for i in row_indices:
                st.session_state[f"chk_req_{i}"] = True
        elif st.session_state.get("chk_req_select_all_prev") and not select_all:
            for i in row_indices:
                st.session_state[f"chk_req_{i}"] = False
        st.session_state["chk_req_select_all_prev"] = select_all

        selected_idx = [i for i in row_indices if st.session_state.get(f"chk_req_{i}", False)]
        row_by_idx = {i: row for i, row in enumerate(rows, start=2)}

        if selected_idx:
            st.info(f"✅ تم تحديد {len(selected_idx)} منتج | {len(selected_idx)} SKUs selected")
            bc1, bc2, bc3, bc4 = st.columns(4)
            with bc1:
                bulk_approve = st.button("✅ موافقة على المحدد | Approve Selected", use_container_width=True, key="bulk_approve_btn")
            with bc2:
                bulk_unavail = st.button("❌ غير متوفر للمحدد | Mark Unavailable", use_container_width=True, key="bulk_unavail_btn")
            with bc3:
                bulk_check = st.button("🔍 تشيك للمحدد | Move to Check", use_container_width=True, key="bulk_check_btn")
            with bc4:
                bulk_order = st.button("🛒 طلب للمحدد | Order Selected", use_container_width=True, key="bulk_order_btn")

            def _clear_selection(idx_list):
                # ما نقدرش نعدّل قيمة checkbox اتعمله render في نفس الـ run ده (Streamlit بيرفضها).
                # بدل كده نسجل "طلب تصفير" يتنفذ في أول حاجة في الـ run الجاي قبل ما الـ checkboxes تترسم.
                st.session_state["_req_clear_pending"] = True

            if bulk_approve:
                dn = now_str()
                ok_rows = []
                for ri in selected_idx:
                    r = row_by_idx[ri]
                    while len(r) < 5: r.append("")
                    ok_rows.append([r[0], r[1], r[1], r[2], r[3], dn])
                if safe_batch_append(approved_sheet, ok_rows):
                    for ri in sorted(selected_idx, reverse=True):
                        safe_delete(requests_sheet, ri)
                    for ok_r in ok_rows:
                        clear_unavailable_ordered_for_sku(ok_r[0])
                    _clear_selection(selected_idx)
                    st.success(f"✅ تمت الموافقة على {len(selected_idx)} منتج | Approved")
                    st.rerun()

            if bulk_unavail:
                dn = now_str()
                for ri in sorted(selected_idx, reverse=True):
                    r = row_by_idx[ri]
                    while len(r) < 5: r.append("")
                    sku_b, qty_b, img_b, da_b = r[0], r[1], r[2], r[3]
                    un_ri, un_row = merge_or_get_existing_row(unavailable_sheet, sku_b)
                    if un_ri:
                        while len(un_row) < 5: un_row.append("")
                        cur_count, rest_dates = parse_count_dates(un_row[4])
                        merged_dates = append_count_date(rest_dates, cur_count + 1, dn)
                        safe_update_row(unavailable_sheet, un_ri, [un_row[0], qty_b, un_row[2] or img_b, un_row[3], merged_dates])
                    else:
                        safe_append(unavailable_sheet, [sku_b, qty_b, img_b, da_b, append_count_date("", 1, dn)])
                    safe_delete(requests_sheet, ri)
                _clear_selection(selected_idx)
                st.success(f"❌ تم تحويل {len(selected_idx)} منتج لغير متوفر | Marked unavailable")
                st.rerun()

            if bulk_check:
                dn = now_str()
                check_rows = []
                for ri in selected_idx:
                    r = row_by_idx[ri]
                    while len(r) < 5: r.append("")
                    check_rows.append(["", r[0], r[1], "", r[2], dn, "", ""])
                if safe_batch_append(sheets["Check"], check_rows):
                    for ri in sorted(selected_idx, reverse=True):
                        safe_delete(requests_sheet, ri)
                    _clear_selection(selected_idx)
                    st.success(f"🔍 تم نقل {len(selected_idx)} منتج للتشيك | Moved to Check")
                    st.rerun()

            if bulk_order:
                dn = now_str()
                for ri in sorted(selected_idx, reverse=True):
                    r = row_by_idx[ri]
                    while len(r) < 5: r.append("")
                    sku_b, qty_b, img_b = r[0], r[1], r[2]
                    ord_ri, ord_row = merge_or_get_existing_row(ordered_sheet, sku_b)
                    if ord_ri:
                        while len(ord_row) < 6: ord_row.append("")
                        cur_count, rest_notes = parse_count_dates(ord_row[5])
                        merged_note = append_count_date(rest_notes, cur_count + 1, dn)
                        safe_update_row(ordered_sheet, ord_ri, [ord_row[0], qty_b, ord_row[2] or img_b, dn, str(cur_count + 1), merged_note])
                    else:
                        safe_append(ordered_sheet, [sku_b, qty_b, img_b, dn, "1", append_count_date("", 1, dn)])
                    safe_delete(requests_sheet, ri)
                _clear_selection(selected_idx)
                st.success(f"🛒 تم طلب {len(selected_idx)} منتج | Ordered")
                st.rerun()

        pending_approval_skus_t1 = get_pending_approval_skus()

        st.divider()
        for i, row in enumerate(rows, start=2):
            while len(row) < 5: row.append("")
            sku,qty,img,date_added,fname = row[0],row[1],row[2],row[3],row[4]
            c_chk,c_img,c_info,c_act = st.columns([0.5,1,4,3])
            with c_chk:
                st.checkbox("", key=f"chk_req_{i}", label_visibility="collapsed")
            with c_img: show_img(img,75)
            with c_info:
                st.markdown(f"**SKU:** `{sku}`")
                show_sku_inv(sku)
                st.markdown(f"**طلب | Requested Qty:** {qty}")
                st.caption(f"📅 {date_added} | 📁 {fname}")
                prev_count = ordered_skus.get(sku.upper(),0)
                if prev_count > 0:
                    ordn = ordinal_map.get(prev_count, f"{prev_count+1}")
                    st.warning(f"🔁 تم الطلب للمرة {ordn} | Already ordered {prev_count} time(s)")
                if sku.strip().upper() in pending_approval_skus_t1:
                    st.markdown(pending_approval_badge_html(), unsafe_allow_html=True)
            with c_act:
                ca,cb,cc,cd = st.columns(4)
                with ca:
                    with st.popover("✅ وافق\nApprove"):
                        nq = st.text_input("Approved Qty | الكمية الموافقة", value=qty, key=f"aqty_{i}")
                        if st.button("✅ تأكيد | Confirm", key=f"aconf_{i}"):
                            safe_append(approved_sheet, [sku,qty,nq,img,date_added,now_str()])
                            safe_delete(requests_sheet,i)
                            clear_unavailable_ordered_for_sku(sku)
                            st.rerun()
                with cb:
                    if st.button("❌ غير\nمتوفر\nUnavailable", key=f"unavail_{i}"):
                        dn = now_str()
                        un_ri, un_row = merge_or_get_existing_row(unavailable_sheet, sku)
                        if un_ri:
                            while len(un_row) < 5: un_row.append("")
                            cur_count, rest_dates = parse_count_dates(un_row[4])
                            new_count = cur_count + 1
                            merged_dates = append_count_date(rest_dates, new_count, dn)
                            safe_update_row(unavailable_sheet, un_ri, [un_row[0], qty, un_row[2] or img, un_row[3], merged_dates])
                        else:
                            safe_append(unavailable_sheet,[sku,qty,img,date_added,append_count_date("",1,dn)])
                        safe_delete(requests_sheet,i)
                        st.rerun()
                with cc:
                    if st.button("🛒 طلب\nOrder", key=f"order_{i}"):
                        dn = now_str()
                        ord_ri, ord_row = merge_or_get_existing_row(ordered_sheet, sku)
                        if ord_ri:
                            while len(ord_row) < 6: ord_row.append("")
                            cur_count, rest_notes = parse_count_dates(ord_row[5])
                            new_count = cur_count + 1
                            merged_note = append_count_date(rest_notes, new_count, dn)
                            safe_update_row(ordered_sheet, ord_ri, [ord_row[0],qty,ord_row[2] or img,dn,str(new_count),merged_note])
                        else:
                            safe_append(ordered_sheet,[sku,qty,img,dn,"1",append_count_date("",1,dn)])
                        safe_delete(requests_sheet,i)
                        st.rerun()
                with cd:
                    if st.button("🗑️ حذف\nDelete", key=f"del_req_{i}"):
                        safe_delete(requests_sheet,i)
                        st.rerun()
            st.divider()


# ══ TAB 2 — الموافقة ══
with tab2:
    st.subheader("✅ الطلبات الموافق عليها | Approved Requests")
    data_ap = get_cached(approved_sheet)
    if len(data_ap) <= 1:
        st.info("لا توجد موافقات | No approvals yet.")
    else:
        rows_ap = data_ap[1:]
        srch = st.text_input("🔍 بحث SKU | Search SKU", key="srch_ap", placeholder="اكتب SKU...")
        indexed_ap = [(i+2, r) for i, r in enumerate(rows_ap)]
        filtered = [(ri, r) for ri, r in indexed_ap if not srch or srch.strip().upper() in r[0].upper()]
        df_ap = pd.DataFrame(rows_ap, columns=data_ap[0])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_ap,"approved")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_ap", use_container_width=True):
                st.session_state["confirm_clear_ap"] = True
        confirm_clear("clear_ap", approved_sheet, "الموافقة | Approved")
        st.write(f"**عرض | Showing: {len(filtered)} / {len(rows_ap)}**")
        for ri, row in filtered:
            while len(row)<6: row.append("")
            sku,qty_r,qty_a,img,da,dap = row[0],row[1],row[2],row[3],row[4],row[5]
            c_img,c_info,c_del = st.columns([1,5,1])
            with c_img: show_img(img,70)
            with c_info:
                st.markdown(f"**SKU:** `{sku}`")
                show_sku_inv(sku)
                if qty_a and qty_a != qty_r:
                    st.markdown(f"**طلبت | Req:** {qty_r} → **وافقوا | App:** ⚠️ **{qty_a}**")
                else:
                    st.markdown(f"**Quantity | الكمية:** {qty_a}")
                st.caption(f"📅 Requested | طُلب: {da} | ✅ Approved | وُفِق: {dap}")
            with c_del:
                if st.button("🗑️", key=f"del_ap_{ri}"):
                    safe_delete(approved_sheet,ri); st.rerun()
            st.divider()

# ══ TAB 3 — غير متوفر ══
with tab3:
    st.subheader("❌ غير متوفر | Unavailable")
    data_un = get_cached(unavailable_sheet)
    if len(data_un) <= 1:
        st.info("لا يوجد | Nothing unavailable yet.")
    else:
        rows_un = data_un[1:]
        srch = st.text_input("🔍 بحث SKU | Search SKU", key="srch_un", placeholder="اكتب SKU...")
        indexed_un = [(i+2, r) for i, r in enumerate(rows_un)]
        filtered = [(ri, r) for ri, r in indexed_un if not srch or srch.strip().upper() in r[0].upper()]
        df_un = pd.DataFrame(rows_un, columns=data_un[0])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_un,"unavailable")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_un", use_container_width=True):
                st.session_state["confirm_clear_un"] = True
        confirm_clear("clear_un", unavailable_sheet, "غير المتوفر | Unavailable")
        st.write(f"**عرض | Showing: {len(filtered)} / {len(rows_un)}**")
        for ri, row in filtered:
            while len(row)<5: row.append("")
            sku,qty,img,da,dm = row[0],row[1],row[2],row[3],row[4]
            cnt_un, dates_un = parse_count_dates(dm)
            c_img,c_info,c_act = st.columns([1,4,2])
            with c_img: show_img(img,70)
            with c_info:
                st.markdown(f"**SKU:** `{sku}`")
                tc_badge_un = warehouse_available_badge(sku.strip().upper())
                if tc_badge_un:
                    st.markdown(tc_badge_un, unsafe_allow_html=True)
                show_sku_inv(sku)
                st.markdown(f"**Qty طلب | Requested:** {qty}")
                if cnt_un > 1:
                    st.warning(f"🔁 تكرر {cnt_un} مرة | Marked unavailable {cnt_un}x")
                st.caption(f"📅 Requested | طُلب: {da}")
                if dates_un:
                    st.caption(f"❌ غير متوفر بتاريخ | Unavailable on: {dates_un}")
            with c_act:
                with st.popover("↩️ رجّع للموافقة\nReturn to Approved"):
                    nq_un = st.text_input("الكمية المعدّلة | Adjusted Qty", value=qty, key=f"un_ret_qty_{ri}")
                    if st.button("✅ أرسل للموافقة | Send to Approved", key=f"un_ret_conf_{ri}"):
                        safe_append(approved_sheet,[sku,qty,nq_un,img,da,now_str()])
                        safe_delete(unavailable_sheet,ri)
                        clear_unavailable_ordered_for_sku(sku)
                        st.rerun()
                if st.button("🗑️", key=f"del_un_{ri}"):
                    safe_delete(unavailable_sheet,ri); st.rerun()
            st.divider()

# ══ TAB 4 — تم الطلب ══
with tab4:
    st.subheader("🛒 تم الطلب | Ordered Items")
    data_ord = get_cached(ordered_sheet)
    if len(data_ord) <= 1:
        st.info("لا يوجد طلبات منجزة | No ordered items yet.")
    else:
        rows_ord = data_ord[1:]
        srch = st.text_input("🔍 بحث SKU | Search SKU", key="srch_ord", placeholder="اكتب SKU...")
        indexed_ord = [(i+2, r) for i, r in enumerate(rows_ord)]
        filtered = [(ri, r) for ri, r in indexed_ord if not srch or srch.strip().upper() in r[0].upper()]
        df_ord = pd.DataFrame(rows_ord, columns=data_ord[0])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_ord,"ordered")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_ord", use_container_width=True):
                st.session_state["confirm_clear_ord"] = True
        confirm_clear("clear_ord", ordered_sheet, "تم الطلب | Ordered")
        st.write(f"**عرض | Showing: {len(filtered)} / {len(rows_ord)}**")
        for ri, row in filtered:
            while len(row)<6: row.append("")
            sku,qty,img,da,cnt,note = row[0],row[1],row[2],row[3],row[4],row[5]
            cnt_ord, dates_ord = parse_count_dates(note)
            c_img,c_info,c_act = st.columns([1,4,2])
            with c_img: show_img(img,70)
            with c_info:
                st.markdown(f"**SKU:** `{sku}`")
                tc_badge_ord = warehouse_available_badge(sku.strip().upper())
                if tc_badge_ord:
                    st.markdown(tc_badge_ord, unsafe_allow_html=True)
                show_sku_inv(sku)
                st.markdown(f"**Quantity | الكمية:** {qty}")
                if cnt_ord > 1:
                    st.warning(f"🔁 تكرر {cnt_ord} مرة | Ordered {cnt_ord}x")
                if dates_ord:
                    st.caption(f"🗓️ تواريخ الطلب | Order dates: {dates_ord}")
                st.caption(f"📅 آخر تحديث | Last update: {da} | 🔢 عدد الطلبات | Order Count: {cnt}")
            with c_act:
                ca,cb = st.columns(2)
                with ca:
                    with st.popover("↩️ رجّع\nReturn"):
                        nq = st.text_input("الكمية المعدّلة | Adjusted Qty", value=qty, key=f"ret_qty_{ri}")
                        if st.button("✅ أرسل للموافقة | Send to Approved", key=f"ret_conf_{ri}"):
                            safe_append(approved_sheet,[sku,qty,nq,img,da,now_str()])
                            safe_delete(ordered_sheet,ri)
                            clear_unavailable_ordered_for_sku(sku)
                            st.rerun()
                with cb:
                    if st.button("🗑️", key=f"del_ord_{ri}"):
                        safe_delete(ordered_sheet,ri); st.rerun()
            st.divider()


# ══ TAB 5 — الجدولة ══
with tab5:
    st.subheader("📅 الجدولة | Scheduled Items")
    links_map = get_links_map()
    col_t,_ = st.columns([1,3])
    with col_t:
        st.download_button("⬇️ Template الجدولة | Schedule Template",
            data=make_empty_template(["ASN","SKU","qty","تاريخ الجدولة"]),
            file_name=f"schedule_template_{file_timestamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

    upl_sc = st.file_uploader("ارفع ملف الجدولة | Upload Schedule File", type=["xlsx","xls","csv"], key="sched_upload")
    if upl_sc:
        try:
            df_sc = pd.read_csv(upl_sc,dtype=str).fillna("") if upl_sc.name.endswith(".csv") else pd.read_excel(upl_sc,dtype=str).fillna("")
            cm = {}
            for c in df_sc.columns:
                cl = c.strip().lower()
                if cl=="asn": cm["asn"]=c
                if cl in ("sku","item nr","item_nr"): cm["sku"]=c
                if cl in ("qty","quantity","كمية"): cm["qty"]=c
                if "جدول" in cl or "schedule" in cl or "date" in cl: cm["date"]=c
            asn_c  = cm.get("asn",  df_sc.columns[0] if len(df_sc.columns)>0 else None)
            sku_c  = cm.get("sku",  df_sc.columns[1] if len(df_sc.columns)>1 else None)
            qty_c  = cm.get("qty",  df_sc.columns[2] if len(df_sc.columns)>2 else None)
            date_c = cm.get("date", df_sc.columns[3] if len(df_sc.columns)>3 else None)
            st.info(f"📊 {len(df_sc)} صف | rows")
            st.dataframe(df_sc, use_container_width=True, height=150)
            if st.button("📤 إضافة الجدولة | Add Schedule", type="primary"):
                existing = get_scheduled_normalized(force=True)
                ex_pairs = set()
                if len(existing)>1:
                    for r in existing[1:]:
                        while len(r)<2: r.append("")
                        ex_pairs.add((r[0].strip().upper(),r[1].strip().upper()))
                dn = now_str()
                to_add, skipped = [], 0
                for _,row in df_sc.iterrows():
                    asn  = str(row[asn_c]).strip()  if asn_c  else ""
                    sku  = str(row[sku_c]).strip()  if sku_c  else ""
                    qty  = str(row[qty_c]).strip()  if qty_c  else ""
                    dval = str(row[date_c]).strip() if date_c else ""
                    img  = links_map.get(sku.upper(),"")
                    pd_  = parse_excel_date(dval)
                    if pd_:
                        ds = pd_.strftime("%Y-%m-%d")
                    else:
                        ds = str(dval).strip()[:10] if dval else ""
                    pair = (asn.upper(),sku.upper())
                    if asn and asn.lower()!="nan":
                        if pair in ex_pairs:
                            skipped+=1
                        else:
                            to_add.append(build_scheduled_row(asn,sku,qty,ds,img,dn))
                            ex_pairs.add(pair)
                safe_batch_append(scheduled_sheet,to_add)
                msg = f"✅ أُضيف | Added: {len(to_add)}"
                if skipped: msg += f" | ⚠️ مكرر | Duplicates: {skipped}"
                st.success(msg); st.rerun()
        except Exception as e:
            st.error(f"❌ {e}")

    st.divider()
    st.subheader("📋 الجدولة الحالية | Current Schedule")
    data_sch = get_scheduled_normalized()
    if len(data_sch) <= 1:
        st.info("لا توجد جدولة | No scheduled items.")
    else:
        rows_sch = data_sch[1:]

        def sort_key(r):
            d = parse_excel_date(r[3] if len(r)>3 else "")
            return d if d else datetime(2099,1,1)
        rows_sch_sorted = sorted(rows_sch, key=sort_key)

        # جلب الإشعارات لعرض علامة بجانب ASN
        cancel_notif_asns = {
            n["asn"].upper()
            for n in st.session_state.get("check_cancel_notifications", [])
        }

        # جلب ASNs اللي اتشيكت
        chk_data_t5 = get_cached(sheets["Check"])
        checked_asns = set()
        if len(chk_data_t5) > 1:
            for cr in chk_data_t5[1:]:
                if cr: checked_asns.add(cr[0].strip().upper())

        asn_groups = {}
        for r in rows_sch_sorted:
            while len(r)<6: r.append("")
            asn = r[0].strip()
            if asn not in asn_groups:
                asn_groups[asn] = {"date":r[3],"skus":[],"checked": asn.upper() in checked_asns}
            asn_groups[asn]["skus"].append(r)

        df_sch = pd.DataFrame(rows_sch, columns=data_sch[0])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_sch,"scheduled")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_sc", use_container_width=True):
                st.session_state["confirm_clear_sc"] = True
        confirm_clear("clear_sc", scheduled_sheet, "الجدولة | Schedule")

        c_srch1, c_srch2 = st.columns(2)
        with c_srch1:
            srch_asn = st.text_input("🔍 بحث ASN | Search by ASN", key="srch_asn", placeholder="اكتب رقم ASN...")
        with c_srch2:
            srch_sku_sch = st.text_input("🔍 بحث SKU | Search by SKU", key="srch_sku_sch", placeholder="اكتب SKU...")
        today = datetime.now().date()
        st.write(f"**إجمالي ASN | Total ASNs: {len(asn_groups)}**")

        for asn, group in asn_groups.items():
            if srch_asn and srch_asn.strip().upper() not in asn.upper():
                continue
            skus_ = group["skus"]
            if srch_sku_sch and not any(srch_sku_sch.strip().upper() in r[1].strip().upper() for r in skus_):
                continue
            sdate   = group["date"]
            pd_date = parse_excel_date(sdate)
            is_exp  = pd_date and today > pd_date.date()
            has_alert = any(
                inv_map.get(r[1].strip().upper(),{}).get("sales",0) > 0 and
                _to_int(r[2]) > inv_map.get(r[1].strip().upper(),{}).get("sales",0)
                for r in skus_)

            # هل عنده إشعار كنسل؟
            has_cancel_notif = asn.upper() in cancel_notif_asns

            border = "#ef4444" if has_alert else "#f59e0b" if is_exp else "#3b82f6"
            bg     = "#2d1515" if has_alert else "#2d2000" if is_exp else "#0f172a"

            # ══ ASN Header ══
            cancel_badge = ""
            if has_cancel_notif:
                cancel_badge = ' &nbsp;<span style="background:#7f1d1d;color:#fca5a5;border-radius:6px;padding:2px 10px;font-size:12px;font-weight:bold;">🚫 اتشيك واتكنسل | Checked & Cancelled</span>'

            st.markdown(
                f'<div style="border-left:5px solid {border};background:{bg};color:white;border-radius:10px;padding:8px 14px;margin-bottom:4px;">'
                f'<b>ASN:</b> {asn} &nbsp;|&nbsp; 📅 <b>تاريخ الجدولة | Schedule Date:</b> <b>{sdate}</b>'
                f'{cancel_badge}</div>',
                unsafe_allow_html=True)

            # ══ عرض إشعار الكنسل التفصيلي بجانب الـ ASN ══
            if has_cancel_notif:
                # إيجاد الإشعار المناسب
                for notif in st.session_state.get("check_cancel_notifications", []):
                    if notif.get("asn","").upper() == asn.upper():
                        notif_skus_list = notif.get("skus", [])
                        notif_reason    = notif.get("reason","")
                        notif_ts        = notif.get("ts","")

                        with st.container():
                            st.markdown(
                                f'<div style="background:#1a0000;border:1px solid #ef4444;border-radius:8px;'
                                f'padding:8px 12px;margin:4px 0 8px 0;">'
                                f'<span style="color:#fca5a5;font-weight:bold;">🚫 تم الكنسل من التشييك | Cancelled from Check</span><br>'
                                f'<span style="color:#fcd34d;font-size:12px;">📝 السبب | Reason: {notif_reason if notif_reason else "—"}</span><br>'
                                f'<span style="color:#94a3b8;font-size:11px;">🕐 {notif_ts}</span>'
                                f'</div>',
                                unsafe_allow_html=True)

                            # صور الـ SKUs المكنسلة
                            if notif_skus_list:
                                lm_t5 = get_links_map()
                                img_cols_t5 = st.columns(min(len(notif_skus_list[:6]), 6))
                                for ci3, sk3 in enumerate(notif_skus_list[:6]):
                                    img_url3 = lm_t5.get(sk3.strip().upper(), "")
                                    with img_cols_t5[ci3]:
                                        if img_url3 and img_url3.startswith("http"):
                                            st.image(img_url3, width=60, caption=sk3[:10])
                                        else:
                                            st.markdown(f"🖼️ `{sk3[:10]}`")
                        break

            # ══ SKUs ══
            for r in skus_:
                while len(r)<6: r.append("")
                sku,qty,img = r[1].strip(),r[2],r[4]
                info    = inv_map.get(sku.upper(),{})
                monthly = info.get("sales",0)
                is_al   = monthly>0 and _to_int(qty)>monthly
                c_img2,c_info2 = st.columns([1,6])
                with c_img2: show_img(img,60)
                with c_info2:
                    note_badge = ' &nbsp;<span style="background:#8b5cf6;color:white;border-radius:5px;padding:1px 7px;font-size:11px;">☑️ تم تشييكه | Checked</span>' if (len(r)>6 and "تم تشييكه" in str(r[6])) else ""
                    st.markdown(f"&nbsp;&nbsp;**SKU:** `{sku}` | **Qty:** {qty}" + note_badge, unsafe_allow_html=True)
                    show_sku_inv(sku)
                    if is_al:
                        st.markdown(f"&nbsp;&nbsp;🔴 **تنبيه | Alert:** الكمية ({qty}) > المبيع ({monthly})")

            ca,cb,cc,cd = st.columns(4)
            with ca:
                with st.popover("☑️ Check"):
                    st.markdown(f"**ASN:** `{asn}` — اختر SKUs للتشييك | Select SKUs to check")
                    select_all = st.checkbox("تحديد الكل | Select All", key=f"chk_all_{asn}")
                    selected_skus = {}
                    for ri2,r in enumerate(skus_):
                        while len(r)<6: r.append("")
                        sku2 = r[1].strip()
                        default_val = select_all
                        selected_skus[sku2] = st.checkbox(f"`{sku2}` — Qty: {r[2]}", value=default_val, key=f"chk_sku_{asn}_{ri2}")
                    if st.button("✅ أرسل للتشييك | Send to Check", key=f"send_chk_{asn}"):
                        dn = now_str()
                        all_selected = all(selected_skus.values())
                        to_add = []
                        for r in skus_:
                            while len(r)<6: r.append("")
                            sku2 = r[1].strip()
                            if not all_selected:
                                flag = "highlighted" if selected_skus.get(sku2,False) else ""
                            else:
                                flag = ""
                            to_add.append([r[0],r[1],r[2],r[3],r[4],dn,"",flag])
                        safe_batch_append(sheets["Check"], to_add)
                        sch_d = get_scheduled_normalized(force=True)
                        del_i = [i2 for i2,sr in enumerate(sch_d[1:],start=2) if sr[0].strip().upper()==asn.upper()]
                        for i2 in sorted(del_i,reverse=True):
                            safe_delete(scheduled_sheet,i2)
                        st.success(f"☑️ تم الإرسال للتشييك | Sent to Check — ASN: {asn}")
                        st.rerun()
            with cb:
                with st.popover("🚫 كنسل - غير متوفر\nCancel - Unavailable"):
                    reason_u = st.text_input("سبب إضافي | Additional reason", key=f"rsn_u_{asn}", placeholder="اختياري | Optional")
                    if st.button("✅ تأكيد الكنسل | Confirm Cancel", key=f"can_u_{asn}"):
                        dn = now_str()
                        to_add = [[r[0],r[1],r[2],r[3],r[4],r[5],f"غير متوفر | Unavailable — {reason_u}",dn] for r in skus_]
                        safe_batch_append(cancelled_sheet, to_add)
                        sch_data = get_scheduled_normalized(force=True)
                        del_idx = [idx for idx,sr in enumerate(sch_data[1:],start=2) if sr[0].strip().upper()==asn.upper()]
                        for idx in sorted(del_idx, reverse=True):
                            safe_delete(scheduled_sheet,idx)
                        st.success("🚫 تم الكنسل | Cancelled"); st.rerun()
            with cc:
                with st.popover("🔄 كنسل - تغيير موعد\nReschedule"):
                    reason_r = st.text_input("سبب التغيير | Reschedule reason", key=f"rsn_r_{asn}", placeholder="مثال: تأخير مورد")
                    if st.button("✅ تأكيد | Confirm", key=f"can_r_{asn}"):
                        dn = now_str()
                        to_add = [[r[0],r[1],r[2],r[3],r[4],r[5],reason_r,dn] for r in skus_]
                        safe_batch_append(reschedule_sheet, to_add)
                        sch_data = get_scheduled_normalized(force=True)
                        del_idx = [idx for idx,sr in enumerate(sch_data[1:],start=2) if sr[0].strip().upper()==asn.upper()]
                        for idx in sorted(del_idx, reverse=True):
                            safe_delete(scheduled_sheet,idx)
                        st.success("🔄 تم النقل لتعديل الموعد | Moved to Rescheduled"); st.rerun()
            with cd:
                status = "⚠️ منتهي | Expired" if is_exp else "✅ ساري | Active"
                st.markdown(f"&nbsp;{status}")
            st.divider()


# ══ TAB CHECK — تشييك ══
with tab_check:
    st.subheader("☑️ قيد التشييك | Under Check")
    st.caption("ASNs المحولة للتشييك | ASNs moved to check — رجّعها للجدولة أو كنسلها | Return to schedule or cancel")

    # إشعارات الكنسل من التشييك
    if st.session_state.get("check_cancel_notifications"):
        st.markdown("---")
        st.markdown("### 🔔 إشعارات الكنسل الأخيرة | Recent Cancel Notifications")
        for notif in st.session_state["check_cancel_notifications"]:
            asn_n   = notif.get("asn","")
            sdate_n = notif.get("sdate","")
            skus_n  = notif.get("skus",[])
            reason_n= notif.get("reason","")
            ts_n    = notif.get("ts","")
            skus_str = ", ".join(skus_n[:5]) + ("..." if len(skus_n)>5 else "")
            st.error(f"🚫 ASN **{asn_n}** (📅 {sdate_n}) — SKUs: {skus_str} — السبب | Reason: {reason_n} — {ts_n}")
        if st.button("✖️ مسح الإشعارات | Clear Notifications", key="clear_notifs"):
            delete_all_cancel_notifications()
            st.session_state["check_cancel_notifications"] = []
            st.rerun()
        st.markdown("---")

    data_chk = get_cached(sheets["Check"])
    if len(data_chk) <= 1:
        st.info("لا يوجد | No items under check.")
    else:
        rows_chk = data_chk[1:]
        chk_groups = {}
        for idx, r in enumerate(rows_chk, start=2):
            while len(r) < 8: r.append("")
            asn = r[0].strip()
            if asn not in chk_groups:
                chk_groups[asn] = {"date":r[3],"skus":[],"indices":[]}
            chk_groups[asn]["skus"].append(r)
            chk_groups[asn]["indices"].append(idx)

        df_chk = pd.DataFrame(rows_chk, columns=data_chk[0])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_chk,"check")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_chk", use_container_width=True):
                st.session_state["confirm_clear_chk"] = True
        confirm_clear("clear_chk", sheets["Check"], "التشييك | Check")

        st.write(f"**إجمالي ASN | Total ASNs: {len(chk_groups)}**")

        for asn, grp in chk_groups.items():
            sdate = grp["date"]
            skus_ = grp["skus"]
            has_highlighted = any(len(r)>7 and r[7]=="highlighted" for r in skus_)

            st.markdown(
                f'<div style="border-left:5px solid #8b5cf6;background:#1a0a2e;border-radius:10px;padding:8px 14px;margin-bottom:4px;">'
                f'<b>ASN:</b> {asn} &nbsp;|&nbsp; 📅 <b>تاريخ الجدولة | Schedule Date:</b> <b>{sdate}</b>'
                + (' &nbsp; 🔴 <b>يوجد SKUs مميزة | Has highlighted SKUs</b>' if has_highlighted else '') +
                f'</div>', unsafe_allow_html=True)

            for r in skus_:
                while len(r)<8: r.append("")
                sku,qty,img,flag = r[1].strip(),r[2],r[4],r[7]
                is_highlighted = flag=="highlighted"
                bg_color = "#2d0a0a" if is_highlighted else "#0f172a"
                border_c = "#ef4444" if is_highlighted else "#8b5cf6"

                st.markdown(
                    f'<div style="border-left:4px solid {border_c};background:{bg_color};'
                    f'border-radius:8px;padding:6px 10px;margin:4px 0;">',
                    unsafe_allow_html=True)
                c_img2,c_info2 = st.columns([1,6])
                with c_img2: show_img(img,60)
                with c_info2:
                    tag = " 🔴 **مميز | Highlighted**" if is_highlighted else ""
                    st.markdown(f"**SKU:** `{sku}` | **Qty:** {qty}{tag}")
                    show_sku_inv(sku)
                st.markdown('</div>', unsafe_allow_html=True)

            ca,cb = st.columns(2)
            with ca:
                if st.button(f"↩️ رجّع للجدولة | Return to Schedule — {asn}", key=f"ret_chk_{asn}", type="primary"):
                    dn = now_str()
                    lm = get_links_map()
                    to_add = [build_scheduled_row(r[0],r[1],r[2],r[3],lm.get(r[1].strip().upper(),r[4]),dn,"تم تشييكه | Checked","") for r in skus_]
                    safe_batch_append(scheduled_sheet, to_add)
                    for idx in sorted(grp["indices"], reverse=True):
                        safe_delete(sheets["Check"], idx)
                    st.success(f"✅ تم الإرجاع للجدولة | Returned — ASN: {asn}")
                    st.rerun()
            with cb:
                with st.popover(f"🚫 كنسل | Cancel — {asn}"):
                    cancel_reason = st.text_input("سبب الكنسل | Cancel reason", key=f"chk_rsn_{asn}")
                    if st.button("✅ تأكيد الكنسل | Confirm Cancel", key=f"chk_can_{asn}"):
                        dn = now_str()
                        to_add = [[r[0],r[1],r[2],r[3],r[4],r[5],
                                   f"تشييك — {cancel_reason} | Check — {cancel_reason}",dn] for r in skus_]
                        safe_batch_append(cancelled_sheet, to_add)
                        for idx in sorted(grp["indices"], reverse=True):
                            safe_delete(sheets["Check"], idx)

                        # ══ إشعار الكنسل — الصيغة الجديدة كـ dict ══
                        hl_skus = [r[1].strip() for r in skus_ if len(r)>7 and r[7]=="highlighted"]
                        all_skus_list = [r[1].strip() for r in skus_]
                        notif_skus_final = hl_skus if hl_skus else all_skus_list

                        new_notif = {
                            "asn":    asn,
                            "sdate":  sdate,
                            "skus":   notif_skus_final,
                            "reason": cancel_reason,
                            "ts":     dn,
                        }
                        # حفظ في Google Sheets (يدوم بعد الإغلاق)
                        save_cancel_notification(asn, notif_skus_final, sdate, cancel_reason, dn)
                        if "check_cancel_notifications" not in st.session_state:
                            st.session_state["check_cancel_notifications"] = []
                        st.session_state["check_cancel_notifications"].insert(0, new_notif)
                        st.session_state["check_cancel_notifications"] = st.session_state["check_cancel_notifications"][:50]
                        st.success("🚫 تم الكنسل | Cancelled")
                        st.rerun()
            st.divider()


# ══ TAB 6 — جدولة ملغية ══
with tab6:
    st.subheader("🚫 الجدولة الملغية | Cancelled Schedule")
    data_can = get_cached(cancelled_sheet)
    if len(data_can) <= 1:
        st.info("لا يوجد إلغاء | No cancelled schedules.")
    else:
        rows_can = data_can[1:]
        srch = st.text_input("🔍 بحث ASN | Search ASN", key="srch_can", placeholder="اكتب ASN...")
        indexed_can = [(i+2, r) for i, r in enumerate(rows_can)]
        filtered = [(ri, r) for ri, r in indexed_can if not srch or srch.strip().upper() in r[0].upper()]
        df_can = pd.DataFrame(rows_can, columns=data_can[0])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_can,"cancelled")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_can", use_container_width=True):
                st.session_state["confirm_clear_can"] = True
        confirm_clear("clear_can", cancelled_sheet, "الملغية | Cancelled")
        st.write(f"**عرض | Showing: {len(filtered)} / {len(rows_can)}**")
        for ri, row in filtered:
            while len(row)<8: row.append("")
            asn,sku,qty,sd,img,dadd,reason,dcan = row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7]
            c_img,c_info,c_del = st.columns([1,5,1])
            with c_img: show_img(img,70)
            with c_info:
                st.markdown(f"**ASN:** `{asn}` | **SKU:** `{sku}`")
                show_sku_inv(sku)
                st.markdown(f"**Qty | الكمية:** {qty}")
                st.caption(f"📅 Schedule | جدولة: {sd} | 🚫 Cancelled | ألغي: {dcan}")
                if reason: st.caption(f"📝 السبب | Reason: {reason}")
            with c_del:
                if st.button("🗑️", key=f"del_can_{ri}"):
                    safe_delete(cancelled_sheet,ri); st.rerun()
            st.divider()

# ══ TAB 7 — تعديل الموعد ══
with tab7:
    st.subheader("🔄 تعديل الموعد | Rescheduled Items")
    st.caption("عدّل الكميات وأضف ASN جديد وأرجع للجدولة | Edit quantities, add new ASN, return to schedule")
    data_res = get_cached(reschedule_sheet)
    if len(data_res) <= 1:
        st.info("لا يوجد | No rescheduled items.")
    else:
        rows_res = data_res[1:]
        asn_res_groups = {}
        for idx, r in enumerate(rows_res, start=2):
            while len(r)<8: r.append("")
            asn = r[0].strip()
            if asn not in asn_res_groups:
                asn_res_groups[asn] = {"old_date":r[3],"reason":r[6],"date_moved":r[7],"skus":[],"indices":[]}
            asn_res_groups[asn]["skus"].append(r)
            asn_res_groups[asn]["indices"].append(idx)

        df_res = pd.DataFrame(rows_res, columns=data_res[0])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_res,"rescheduled")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_res", use_container_width=True):
                st.session_state["confirm_clear_res"] = True
        confirm_clear("clear_res", reschedule_sheet, "تعديل الموعد | Rescheduled")

        links_map2 = get_links_map()
        for asn, grp in asn_res_groups.items():
            st.markdown(
                f'<div style="border-left:5px solid #f59e0b;background:#1a1500;border-radius:10px;padding:8px 14px;margin-bottom:4px;color:white;">'
                f'<span style="font-size:15px;font-weight:bold;color:white;">ASN: {asn}</span><br>'
                f'<span style="color:white;">📅 <b style="font-size:16px;color:#fcd34d;">موعد قديم | Old Date: {grp["old_date"]}</b></span></div>',
                unsafe_allow_html=True)
            if grp["reason"]:
                st.caption(f"📝 سبب التعديل | Reason: {grp['reason']}")
            with st.expander(f"✏️ تعديل وإرجاع للجدولة | Edit & Reschedule ASN {asn}", expanded=False):
                new_asn  = st.text_input("ASN جديد | New ASN", value=asn, key=f"new_asn_{asn}")
                new_date = st.text_input("تاريخ جديد | New Schedule Date (YYYY-MM-DD)", value="", key=f"new_date_{asn}", placeholder="2025-08-15")
                edited_skus = []
                for ri2, r in enumerate(grp["skus"]):
                    while len(r)<6: r.append("")
                    sku,qty,img = r[1].strip(),r[2],r[4]
                    c_img2,c_s2,c_q2 = st.columns([1,3,2])
                    with c_img2: show_img(img,55)
                    with c_s2:
                        st.markdown(f"**SKU:** `{sku}`")
                        show_sku_inv(sku)
                    with c_q2:
                        new_qty = st.text_input("Qty | الكمية", value=qty, key=f"res_qty_{asn}_{ri2}")
                    edited_skus.append((sku, new_qty, img))
                if st.button("✅ أرجع للجدولة | Return to Schedule", key=f"ret_sch_{asn}", type="primary"):
                    if not new_date.strip():
                        st.error("❌ أدخل تاريخ جديد | Enter new schedule date")
                    else:
                        dn = now_str()
                        to_add = [build_scheduled_row(new_asn, sku, qty, new_date, links_map2.get(sku.upper(), img), dn) for sku,qty,img in edited_skus]
                        safe_batch_append(scheduled_sheet, to_add)
                        for idx in sorted(grp["indices"], reverse=True):
                            safe_delete(reschedule_sheet, idx)
                        st.success(f"✅ تم الإرجاع للجدولة | Returned to schedule — ASN: {new_asn}")
                        st.rerun()
            st.divider()

# ══ TAB 8 — تنبيهات ══
with tab8:
    st.subheader("⚠️ تنبيهات الجدولة | Schedule Alerts")
    st.caption("الكمية المجدولة أعلى من المبيع الشهري | Scheduled qty > Monthly sales")
    data_sc8 = get_scheduled_normalized()
    alerts = []
    if len(data_sc8) > 1:
        for row in data_sc8[1:]:
            while len(row)<6: row.append("")
            asn,sku,qty,sdate,img = row[0],row[1],row[2],row[3],row[4]
            info    = inv_map.get(sku.upper(),{})
            monthly = info.get("sales",0)
            stock   = info.get("total_stock",0)
            try:
                if monthly>0 and _to_int(qty)>monthly:
                    alerts.append((asn,sku,qty,monthly,stock,sdate,img))
            except: pass
    if not inv_map:
        st.info("ارفع ملف المخزون أولاً | Upload Inventory first")
    elif not alerts:
        st.success("✅ لا توجد تنبيهات | No alerts")
    else:
        df_al = pd.DataFrame(alerts, columns=["ASN","SKU","Scheduled Qty","Monthly Sales","Total Stock","Schedule Date","Image URL"])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_al,"alerts")
        with c2: st.error(f"⚠️ تنبيهات | Alerts: {len(alerts)}")
        for asn,sku,qty,monthly,stock,sdate,img in alerts:
            c_img,c_info = st.columns([1,6])
            with c_img: show_img(img,70)
            with c_info:
                st.markdown(f"**ASN:** `{asn}` | **SKU:** `{sku}`")
                show_sku_inv(sku)
                st.markdown(f"🔴 **الكمية المجدولة | Scheduled:** {qty} > **المبيع الشهري | Monthly Sales:** {monthly}")
                st.caption(f"📅 تاريخ الجدولة | Schedule Date: {sdate}")
            st.divider()


# ══ TAB 9 — المخزون ══
with tab9:
    st.subheader("📊 المخزون والمبيع الشهري | Inventory & Monthly Sales")
    links_map = get_links_map()
    render_tacweed_upload("inv")
    render_warehouse_stock_upload("inv")
    col_t,_ = st.columns([1,3])
    with col_t:
        st.download_button("⬇️ Template المخزون | Inventory Template",
            data=make_empty_template(["warehouse_code","sku","STOCCCCK.QTY","مبيع شهر جدول.QTY"]),
            file_name=f"inventory_template_{file_timestamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    upl_inv = st.file_uploader("ارفع ملف المخزون | Upload Inventory File", type=["xlsx","xls","xlsm","csv"], key="inv_upload")
    if upl_inv:
        try:
            df_inv = pd.read_csv(upl_inv,dtype=str).fillna("") if upl_inv.name.endswith(".csv") else pd.read_excel(upl_inv,dtype=str).fillna("")
            wh_col=sku_col=stock_col=sales_col=None
            for c in df_inv.columns:
                cl = c.strip().lower()
                if "warehouse" in cl: wh_col=c
                if cl in ("sku","item nr","item_nr"): sku_col=c
                if "stock" in cl: stock_col=c
                if "مبيع" in cl or "sales" in cl: sales_col=c
                if "qty" in cl and sales_col is None: sales_col=c
            if not wh_col:    wh_col    = df_inv.columns[0]
            if not sku_col:   sku_col   = df_inv.columns[1] if len(df_inv.columns)>1 else df_inv.columns[0]
            if not stock_col: stock_col = df_inv.columns[2] if len(df_inv.columns)>2 else None
            if not sales_col: sales_col = df_inv.columns[3] if len(df_inv.columns)>3 else None
            st.info(f"📊 {len(df_inv)} صف | WH:`{wh_col}` SKU:`{sku_col}` Stock:`{stock_col}` Sales:`{sales_col}`")
            st.dataframe(df_inv.head(10), use_container_width=True, height=180)
            def do_upload(replace=False):
                dn = now_str()
                to_add = []
                for _,row in df_inv.iterrows():
                    wh  = str(row[wh_col]).strip()    if wh_col    else ""
                    sku = str(row[sku_col]).strip()   if sku_col   else ""
                    stk = str(row[stock_col]).strip() if stock_col else ""
                    sal = str(row[sales_col]).strip() if sales_col else ""
                    img = links_map.get(sku.upper(),"")
                    if sku and sku.lower()!="nan":
                        to_add.append([sku,wh,stk,sal,img,dn])
                if replace: safe_delete_all(inventory_sheet)
                safe_batch_append(inventory_sheet,to_add)
                clear_cache(inventory_sheet)
                return len(to_add)
            ca,cb = st.columns(2)
            with ca:
                if st.button("📤 إضافة للموجود | Append", type="primary", use_container_width=True):
                    n = do_upload(replace=False)
                    st.success(f"✅ أُضيف {n} صف | rows added"); st.rerun()
            with cb:
                if st.button("🔄 استبدال الكل | Replace All", type="secondary", use_container_width=True):
                    st.session_state["confirm_replace_inv"] = True
            if st.session_state.get("confirm_replace_inv"):
                st.warning("⚠️ هيمسح الكل ويرفع الجديد؟ | Replace all data?")
                cy,cn = st.columns(2)
                if cy.button("✅ نعم | Yes", key="yes_rep_inv"):
                    n = do_upload(replace=True)
                    st.session_state["confirm_replace_inv"] = False
                    st.success(f"✅ تم الاستبدال — {n} صف"); st.rerun()
                if cn.button("❌ لا | No", key="no_rep_inv"):
                    st.session_state["confirm_replace_inv"] = False; st.rerun()
        except Exception as e:
            st.error(f"❌ {e}")
    st.divider()
    st.subheader("📋 بيانات المخزون الحالية | Current Inventory")
    if not inv_map:
        st.info("لم يُرفع ملف مخزون بعد | No inventory uploaded yet.")
    else:
        if excluded_wh:
            st.info(f"⚙️ مستثنى من الإجمالي | Excluded: **{', '.join(sorted(excluded_wh))}**")
        srch = st.text_input("🔍 بحث SKU | Search SKU", key="srch_inv", placeholder="اكتب SKU...")
        raw_inv = get_cached(inventory_sheet)
        df_inv_dl = pd.DataFrame(raw_inv[1:], columns=raw_inv[0])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_inv_dl,"inventory")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_inv", use_container_width=True):
                st.session_state["confirm_clear_inv"] = True
        confirm_clear("clear_inv", inventory_sheet, "المخزون | Inventory")
        filtered_inv = {k:v for k,v in inv_map.items() if not srch or srch.strip().upper() in k}
        st.write(f"**SKUs: {len(filtered_inv)}**")
        for sku_key,info in filtered_inv.items():
            c_img,c_info = st.columns([1,6])
            with c_img: show_img(info["img"],70)
            with c_info:
                st.markdown(f"**SKU:** `{info['sku']}`")
                tc_badge = warehouse_available_badge(sku_key)
                if tc_badge:
                    st.markdown(tc_badge, unsafe_allow_html=True)
                st.markdown(f"📦 **إجمالي المخزون | Stock:** **{info['total_stock']}** &nbsp;|&nbsp; 📈 **مبيع شهري | Monthly Sales:** **{info['sales']}**")
                badges = []
                for wh,stk in sorted(info["warehouses"].items()):
                    is_ex=wh.upper() in excluded_wh
                    bg="#4b1010" if is_ex else "#1e3a5f"
                    color="#fca5a5" if is_ex else "#93c5fd"
                    strike="text-decoration:line-through;" if is_ex else ""
                    badges.append(f'<span class="wh-badge" style="background:{bg};color:{color};{strike}">{wh}: {stk}</span>')
                st.markdown("🏭 "+"".join(badges), unsafe_allow_html=True)
                st.caption(f"📅 {info['date']}")
            st.divider()

# ══ TAB 10 — مراجعة المخزون ══
with tab10:
    st.subheader("🔴 مراجعة المخزون | Stock Review")
    st.caption("نفس منطق استعلام Access \"مراجعة مخزون\" — المخزون أقل من تغطية 10 أيام بيع | Same logic as the Access \"مراجعة مخزون\" query — stock below 10-day sales coverage")

    with st.expander("📤 رفع بيانات الأوردرز اليومية | Upload Daily Orders", expanded=False):
        st.caption("ارفع ملف الأوردرز (لازم يحتوي على عمودي sku و order_timestamp) — هيتم استبدال البيانات بالكامل في كل رفعة | Upload orders file (needs sku & order_timestamp columns) — fully replaces existing data each time")
        st.download_button("⬇️ Template فارغ | Empty Template",
            data=make_empty_template(["sku","order_timestamp","status","price","quantity"]),
            file_name=f"daily_orders_template_{file_timestamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, key="dlbtn_do_template")
        upl_do = st.file_uploader("ملف الأوردرز | Orders file", type=["xlsx","xls","csv"], key="daily_orders_upload")
        if upl_do:
            try:
                df_do = pd.read_csv(upl_do,dtype=str).fillna("") if upl_do.name.endswith(".csv") else pd.read_excel(upl_do,dtype=str).fillna("")
                sku_col_do = ts_col_do = status_col_do = None
                for c in df_do.columns:
                    cl = c.strip().lower()
                    if cl == "sku": sku_col_do = c
                    if cl == "order_timestamp" or cl == "order timestamp": ts_col_do = c
                    if cl == "status": status_col_do = c
                if not sku_col_do:
                    for c in df_do.columns:
                        if "sku" in c.strip().lower(): sku_col_do = c; break
                if not ts_col_do:
                    for c in df_do.columns:
                        cl = c.strip().lower()
                        if "timestamp" in cl or "date" in cl: ts_col_do = c; break
                st.info(f"📊 {len(df_do)} صف | SKU:`{sku_col_do}` Timestamp:`{ts_col_do}`")
                st.dataframe(df_do.head(10), use_container_width=True, height=150)
                if sku_col_do and ts_col_do:
                    if st.button("🔄 رفع واستبدال | Upload & Replace", type="primary", key="btn_upload_daily_orders"):
                        dn = now_str()
                        to_add = []
                        price_col_do = None
                        for c in df_do.columns:
                            if c.strip().lower() in ("price","base_price","سعر","السعر","price_egp","unit_price","sale_price","selling_price"): price_col_do = c; break
                        qty_col_do = None
                        for c in df_do.columns:
                            if c.strip().lower() in ("quantity","qty","كمية","الكمية","count"): qty_col_do = c; break
                        for _,row in df_do.iterrows():
                            sku_v   = str(row[sku_col_do]).strip()
                            ts_v    = str(row[ts_col_do]).strip()
                            st_v    = str(row[status_col_do]).strip() if status_col_do else ""
                            price_v = str(row[price_col_do]).strip() if price_col_do else ""
                            qty_v   = str(row[qty_col_do]).strip() if qty_col_do else "1"
                            if sku_v and sku_v.lower()!="nan":
                                to_add.append([sku_v, ts_v, st_v, price_v, qty_v, dn])
                        safe_delete_all(daily_orders_sheet)

                        correct_header = ["SKU","Order Timestamp","Status","Price","Quantity","Date Uploaded"]
                        daily_orders_sheet.update("A1", [correct_header])
 
                        if to_add:
                            safe_batch_append(daily_orders_sheet, to_add)

                        clear_cache(daily_orders_sheet)
                        st.success(f"✅ تم رفع {len(to_add)} صف واستبدال البيانات | Uploaded & replaced {len(to_add)} rows")
                        st.rerun()
                else:
                    st.error("❌ مش لاقي أعمدة SKU أو order_timestamp | Couldn't detect SKU or order_timestamp columns")
            except Exception as e:
                st.error(f"❌ {e}")

    today_d = datetime.now().date()
    d1, d2, d3 = today_d - timedelta(days=1), today_d - timedelta(days=2), today_d - timedelta(days=3)
    day_dates  = [d1, d2, d3]
    day_labels = [f"أمس | Yesterday ({d1.strftime('%m-%d')})",
                  f"أول أمس | Day before ({d2.strftime('%m-%d')})",
                  f"أول أول أمس | 3 days ago ({d3.strftime('%m-%d')})"]
    st.caption(f"📅 بيانات يوم | Data for: **{d1.strftime('%Y-%m-%d')}** (أمس | yesterday) — التنبيه نفسه مبني على أمس فقط، والأيام التانية للعرض فقط | Alert itself is based on yesterday only; the other days are for display")

    delay_days = int(load_settings().get("schedule_delay_days","3") or 3)
    # SKUs ليها جدولة (أو جدولة منتهية) خلال آخر 4 أيام ولسه في فترة الوصول —
    # دي أموره تمام بالفعل، فمينفعش تظهر في مراجعة المخزون خالص
    recent_sched_map_t10_all = get_recent_schedule_rows(days_back=4)
    all_review_rows = compute_stock_sales_rows(d1, day_dates)
    stock_review_rows = []
    for r in all_review_rows:
        if not r["stock_alert"]:
            continue
        if r["sku_up"] in recent_sched_map_t10_all:
            continue
        # لو ليه جدولة موجودة أصلاً وهتوصل قبل النفاد ("✅ ... هيوصل قبل النفاد")،
        # يبقى أموره تمام ومش محتاج مراجعة — نستبعده خالص
        cov_badge_text, cov_badge_color, _ = schedule_coverage_badge(r["sku"], r["days_to_stockout"], delay_days)
        if cov_badge_color == "#22c55e":
            continue
        r["_cov_badge_text"], r["_cov_badge_color"] = cov_badge_text, cov_badge_color
        stock_review_rows.append(r)

    # إضافة المرحلين من تاب المبيعات (محتاج جدولة فقط)
    transferred_from_sales = st.session_state.get("transferred_skus_t14", [])
    existing_skus_in_review = {r["sku_up"] for r in stock_review_rows}
    for tr in transferred_from_sales:
        if is_sku_only_in_excluded_warehouses(tr["sku_up"], excluded_wh):
            continue
        if tr["sku_up"] in recent_sched_map_t10_all:
            continue
        if tr["sku_up"] not in existing_skus_in_review:
            avg_tr = tr.get("effective_avg", 0)
            suggested_tr = round(avg_tr * 18) if avg_tr > 0 else 0
            days_to_so_tr = tr.get("days_to_stockout", 0)
            cov_badge_text_tr, cov_badge_color_tr, _ = schedule_coverage_badge(tr["sku"], days_to_so_tr, delay_days)
            if cov_badge_color_tr == "#22c55e":
                continue
            stock_review_rows.append({
                "sku": tr["sku"], "sku_up": tr["sku_up"],
                "stock": tr["stock"], "sales_month": tr["sales_month"],
                "img": tr["img"], "stock_alert": True, "sales_alert": False,
                "suggested_qty": suggested_tr,
                "days_to_stockout": days_to_so_tr,
                "days_to_stockout_today": days_to_so_tr,
                "qty": tr["day_counts"].get(d1, 0) if tr.get("day_counts") else 0,
                "day_counts": tr.get("day_counts", {d: 0 for d in day_dates}),
                "_transferred_from_sales": True,
                "_cov_badge_text": cov_badge_text_tr, "_cov_badge_color": cov_badge_color_tr,
            })

    stock_review_rows.sort(key=lambda r: (-r["qty"], -r["sales_month"]))

    if not inv_map:
        st.info("ارفع ملف المخزون أولاً من تاب المخزون | Upload Inventory first")
    elif not stock_review_rows:
        st.success("✅ لا توجد SKUs محتاجة مراجعة مخزون | No SKUs need stock review")
    else:
        df_sr = pd.DataFrame([{
            "SKU": r["sku"], "Yesterday": r["day_counts"].get(d1,0), "Day Before": r["day_counts"].get(d2,0),
            "3 Days Ago": r["day_counts"].get(d3,0), "Stock": r["stock"], "Monthly Sales": r["sales_month"],
            "Suggested Qty": r["suggested_qty"], "Days to Stockout": r["days_to_stockout"]
        } for r in stock_review_rows])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_sr,"stock_review")
        with c2: st.error(f"🔴 SKUs محتاجة مراجعة | Needs Review: {len(stock_review_rows)}")

        pending_approval_skus_t10 = get_pending_approval_skus()

        for r in stock_review_rows:
            c_img,c_info = st.columns([1,6])
            with c_img: show_img(r["img"],70)
            with c_info:
                st.markdown(f"**SKU:** `{r['sku']}`")
                tc_badge_t10 = warehouse_available_badge(r["sku_up"])
                if tc_badge_t10:
                    st.markdown(tc_badge_t10, unsafe_allow_html=True)
                if r.get("_transferred_from_sales"):
                    st.markdown('<span style="background:#7c3aed;color:white;border-radius:6px;padding:2px 10px;font-size:11px;">📌 مرحّل من تاب المبيعات — محتاج جدولة | Transferred from Sales tab — needs scheduling</span>', unsafe_allow_html=True)
                st.markdown(f"📦 **المخزون | Stock:** {r['stock']} &nbsp;|&nbsp; 📈 **مبيع شهري | Monthly:** {r['sales_month']}")
                st.markdown("🛒 " + render_day_counts_md(r["day_counts"], day_dates, day_labels))
                st.markdown(f"💡 **اقتراح الكمية | Suggested Qty:** **{r['suggested_qty']}** &nbsp;|&nbsp; ⏳ **نفاد خلال | Days to stockout:** {r['days_to_stockout']} يوم")
                if r["sales_alert"]:
                    st.warning("📈 مبيعات أعلى من المعتاد كمان | Also selling faster than usual")
                st.markdown(f'<span class="status-badge-lg" style="background:{r["_cov_badge_color"]};">{r["_cov_badge_text"]}</span>', unsafe_allow_html=True)
                if r["sku_up"] in pending_approval_skus_t10:
                    st.markdown(pending_approval_badge_html(), unsafe_allow_html=True)
                render_recent_expired_note(r["sku"])
                for note in get_unavailable_ordered_note(r["sku"]):
                    st.markdown(big_note_html(note), unsafe_allow_html=True)
            st.divider()

    # ══ قايمة "منتهي بالكامل" — بنستبعد منها السكيوهات المجدولة مؤخراً وكمان اللي ليها
    #    جدولة هتوصل قبل النفاد أصلاً (أموره تمام، مش محتاجة مراجعة) ══
    missing_rows_t10 = compute_missing_inventory_rows(day_dates)
    filtered_missing_t10 = []
    for r in missing_rows_t10:
        if r["sku_up"] in recent_sched_map_t10_all:
            continue
        cov_badge_text, cov_badge_color, _ = schedule_coverage_badge(r["sku"], 0, delay_days)
        if cov_badge_color == "#22c55e":
            continue
        r["_cov_badge_text"], r["_cov_badge_color"] = cov_badge_text, cov_badge_color
        filtered_missing_t10.append(r)
    missing_rows_t10 = filtered_missing_t10

    st.divider()
    st.subheader("⛔ مخزون منتهي بالكامل | Completely Out of Stock")
    st.caption("SKUs باعت في آخر 3 أيام لكن مالهاش سجل في ملف المخزون أصلاً — يبقى مخزونها انتهى وخرجت من الملف | SKUs with sales in the last 3 days but no record in the Inventory file at all — stock fully ran out")
    if not missing_rows_t10:
        st.success("✅ لا يوجد SKUs خارجة عن المخزون | No SKUs missing from inventory")
    else:
        df_miss10 = pd.DataFrame([{
            "SKU": r["sku"], "Yesterday": r["day_counts"].get(d1,0), "Day Before": r["day_counts"].get(d2,0),
            "3 Days Ago": r["day_counts"].get(d3,0), "Estimated Monthly Sales": r["est_monthly_sales"]
        } for r in missing_rows_t10])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_miss10,"out_of_stock", key="dlbtn_oos_t10")
        with c2: st.error(f"⛔ SKUs منتهية | Out of Stock: {len(missing_rows_t10)}")
        for r in missing_rows_t10:
            c_img,c_info = st.columns([1,6])
            with c_img: show_img(r["img"],70)
            with c_info:
                st.markdown(f"**SKU:** `{r['sku']}`")
                tc_badge_t10b = warehouse_available_badge(r["sku_up"])
                if tc_badge_t10b:
                    st.markdown(tc_badge_t10b, unsafe_allow_html=True)
                st.error("⛔ مخزونه انتهى — مش موجود في ملف المخزون | Stock ran out — not found in inventory file")
                st.markdown("🛒 " + render_day_counts_md(r["day_counts"], day_dates, day_labels))
                st.markdown(f"📈 **مبيع شهري تقديري (بناءً على آخر 3 أيام) | Estimated Monthly Sales (based on last 3 days):** **{r['est_monthly_sales']}**")
                st.markdown(f'<span class="status-badge-lg" style="background:{r["_cov_badge_color"]};">{r["_cov_badge_text"]}</span>', unsafe_allow_html=True)
                if r["sku_up"] in pending_approval_skus_t10:
                    st.markdown(pending_approval_badge_html(), unsafe_allow_html=True)
                render_recent_expired_note(r["sku"])
                for note in get_unavailable_ordered_note(r["sku"]):
                    st.markdown(big_note_html(note), unsafe_allow_html=True)
            st.divider()

# ══ TAB 11 — منتهية الصلاحية ══
with tab11:
    st.subheader("🗂️ الجدولة منتهية الصلاحية | Expired Schedule")
    data_ex = get_cached(expired_sheet)
    if len(data_ex) <= 1:
        st.info("لا يوجد منتهي | No expired items.")
    else:
        rows_ex = data_ex[1:]
        df_ex = pd.DataFrame(rows_ex, columns=data_ex[0])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_ex,"expired")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_ex", use_container_width=True):
                st.session_state["confirm_clear_ex"] = True
        confirm_clear("clear_ex", expired_sheet, "المنتهية | Expired")
        st.write(f"**الإجمالي | Total: {len(rows_ex)}**")
        for i,row in enumerate(rows_ex, start=2):
            while len(row)<7: row.append("")
            asn,sku,qty,sd,img,dadd,dexp = row[0],row[1],row[2],row[3],row[4],row[5],row[6]
            c_img,c_info,c_del = st.columns([1,5,1])
            with c_img: show_img(img,70)
            with c_info:
                st.markdown(f"**ASN:** `{asn}` | **SKU:** `{sku}`")
                show_sku_inv(sku)
                st.markdown(f"**Quantity | الكمية:** {qty}")
                st.caption(f"📅 Schedule | جدولة: {sd} | 🗂️ Expired | انتهى: {dexp}")
            with c_del:
                if st.button("🗑️", key=f"del_ex_{i}"):
                    safe_delete(expired_sheet,i); st.rerun()
            st.divider()

# ══ TAB 12 — الإعدادات ══
with tab12:
    st.subheader("⚙️ الإعدادات | Settings")
    st.caption("الإعدادات محفوظة في جوجل شيت وتبقى بعد الإغلاق | Settings saved in Google Sheets and persist")
    current_settings = load_settings()
    st.markdown("### 🏭 المستودعات المستثناة من حساب المخزون | Excluded Warehouses")
    st.caption("المستودعات المستثناة لا تُحسب في الإجمالي وتظهر بشطب | Excluded warehouses are struck-through and not counted")
    all_wh = sorted({r[1].strip() for r in get_cached(inventory_sheet)[1:] if len(r)>1 and r[1].strip()})
    current_ex_str  = current_settings.get("excluded_warehouses","")
    current_ex_list = [w.strip() for w in current_ex_str.split(",") if w.strip()]
    if all_wh:
        st.write("**المستودعات المتاحة | Available Warehouses:**")
        selected_ex = st.multiselect("اختر المستودعات المستثناة | Select excluded warehouses:",
            options=all_wh, default=[w for w in current_ex_list if w in all_wh], key="wh_multi")
    else:
        st.info("ارفع ملف المخزون أولاً لتظهر المستودعات | Upload inventory first to see warehouses")
        manual = st.text_input("أو اكتب يدوياً | Or type manually (comma-separated):", value=current_ex_str, key="wh_manual")
        selected_ex = [w.strip() for w in manual.split(",") if w.strip()]
    if st.button("💾 حفظ الإعدادات | Save Settings", type="primary"):
        save_setting("excluded_warehouses",",".join(selected_ex))
        st.success("✅ تم الحفظ | Saved — ستُطبَّق عند إعادة التحميل | Will apply on next reload")
        st.rerun()
    st.divider()
    st.markdown("### 📋 الإعدادات الحالية | Current Settings")
    if excluded_wh:
        st.warning(f"🚫 مستودعات مستثناة الآن | Currently excluded: **{', '.join(sorted(excluded_wh))}**")
    else:
        st.success("✅ لا توجد مستودعات مستثناة | All warehouses included in totals")
    if inv_map and all_wh:
        st.markdown("### 🏭 ملخص المستودعات | Warehouse Summary")
        wh_totals = {}
        for info in inv_map.values():
            for wh,stk in info["warehouses"].items():
                wh_totals[wh] = wh_totals.get(wh,0)+stk
        wh_df = pd.DataFrame(
            [(wh,stk,"🚫 مستثنى | Excluded" if wh.upper() in excluded_wh else "✅ محسوب | Included")
             for wh,stk in sorted(wh_totals.items())],
            columns=["Warehouse | المستودع","Total Stock | إجمالي المخزون","Status | الحالة"])
        st.dataframe(wh_df, use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("### ⏳ مدة وصول المخزون بعد الجدولة | Stock Arrival Delay After Scheduling")
    st.caption("عدد الأيام اللي ياخدها المخزون عشان يوصل بعد تاريخ الجدولة (مثال: لو جدولت يوم 16، يوصل بعدها بـ 2-3 أيام) — تُستخدم في تابي مراجعة المخزون ومراجعة المبيعات | Days for stock to arrive after the schedule date — used in Stock Review & Sales Review tabs")
    current_delay = int(current_settings.get("schedule_delay_days","3") or 3)
    new_delay = st.number_input("عدد الأيام | Delay Days", min_value=0, max_value=30, value=current_delay, step=1, key="delay_days_input")
    if st.button("💾 حفظ مدة الوصول | Save Delay", key="save_delay_days"):
        save_setting("schedule_delay_days", str(new_delay))
        st.success("✅ تم الحفظ | Saved")
        st.rerun()

    st.divider()
    st.markdown("### 📅 عدد أيام المبيعات المعروضة في تاب المبيعات | Sales Display Days")
    st.caption("عدد الأيام اللي بتتعرض في تاب المبيعات من اليوم للوراء — مثلاً 7 يعني أمس وأول أمس و... إلخ | Number of past days shown in the Sales tab (e.g. 7 = yesterday + 6 days before)")
    current_sales_days = int(current_settings.get("sales_display_days","7") or 7)
    new_sales_days = st.number_input("عدد الأيام | Display Days", min_value=1, max_value=30, value=current_sales_days, step=1, key="sales_days_input")
    if st.button("💾 حفظ عدد الأيام | Save Sales Days", key="save_sales_days"):
        save_setting("sales_display_days", str(new_sales_days))
        st.success("✅ تم الحفظ | Saved")
        st.rerun()

    st.divider()
    st.markdown("### 📦 أيام تغطية الجدولة المقترحة | Suggested Schedule Coverage Days")
    st.caption("عدد الأيام اللي الكمية المقترحة في تحليل الجدولة هتغطيها — مثلاً 15 يعني الكمية = متوسط اليومي × 15 فقط (تجنب رسوم تخزين الكميات الكبيرة) | Days the suggested qty should cover — e.g. 15 means qty = daily_avg × 15 (avoids storage fees for large quantities)")
    current_cov_days = int(current_settings.get("schedule_coverage_days","15") or 15)
    new_cov_days = st.number_input("أيام التغطية | Coverage Days", min_value=5, max_value=90, value=current_cov_days, step=1, key="cov_days_input")
    if st.button("💾 حفظ أيام التغطية | Save Coverage Days", key="save_cov_days"):
        save_setting("schedule_coverage_days", str(new_cov_days))
        st.success("✅ تم الحفظ | Saved")
        st.rerun()

# ══ TAB 13 — مراجعة المبيعات ══
with tab13:
    st.subheader("📈 مراجعة المبيعات | Sales Review")
    st.caption("متوسط مبيعات آخر 3 أيام أعلى من المعتاد بشكل مستمر (يومين على الأقل) لكن المخزون لسه كافي — بديل أقل حساسية للـ noise من مجرد يوم واحد شاذ | Average of the last 3 days consistently above normal (at least 2 elevated days) but stock still sufficient — less sensitive to a single noisy day")

    today_d2 = datetime.now().date()
    e1, e2, e3 = today_d2 - timedelta(days=1), today_d2 - timedelta(days=2), today_d2 - timedelta(days=3)
    day_dates2  = [e1, e2, e3]
    day_labels2 = [f"أمس | Yesterday ({e1.strftime('%m-%d')})",
                   f"أول أمس | Day before ({e2.strftime('%m-%d')})",
                   f"أول أول أمس | 3 days ago ({e3.strftime('%m-%d')})"]
    st.caption(f"📅 بيانات يوم | Data for: **{e1.strftime('%Y-%m-%d')}** (أمس | yesterday) — التنبيه نفسه مبني على أمس فقط، والأيام التانية للعرض فقط | Alert itself is based on yesterday only; the other days are for display")

    delay_days2 = int(load_settings().get("schedule_delay_days","3") or 3)
    # SKUs ليها جدولة (أو جدولة منتهية) خلال آخر 4 أيام ولسه في فترة الوصول —
    # دي أموره تمام بالفعل، فمينفعش تظهر في مراجعة المبيعات خالص (شايفينها بس في سكشن
    # "مجدولة خلال آخر 4 أيام" تحت)
    recent_sched_map_t13_all = get_recent_schedule_rows(days_back=4)
    # أي SKU ظاهر بالفعل في مراجعة المخزون (تاب 10) مينفعش يتكرر هنا كمان —
    # كل تاب له وظيفة مختلفة والسكو يظهر في تاب واحد بس
    stock_review_skus_t10_for_t13 = {r["sku_up"] for r in stock_review_rows}
    all_review_rows2 = compute_stock_sales_rows(e1, day_dates2)
    valid_days_set = {1,2,3,4,5,6,7,8,10}
    sales_review_rows = []
    for r in all_review_rows2:
        if not (r["days_to_stockout_today"] in valid_days_set
                and r["sales_month"] > 0
                and r["sales_alert"]
                and not r["stock_alert"]
                and r["sku_up"] not in recent_sched_map_t13_all
                and r["sku_up"] not in stock_review_skus_t10_for_t13):
            continue
        # لو ليه جدولة موجودة أصلاً وهتوصل قبل النفاد، يبقى أموره تمام ومش محتاج مراجعة
        cov_badge_text, cov_badge_color, _ = schedule_coverage_badge(r["sku"], r["days_to_stockout"], delay_days2)
        if cov_badge_color == "#22c55e":
            continue
        r["_cov_badge_text"], r["_cov_badge_color"] = cov_badge_text, cov_badge_color
        sales_review_rows.append(r)
    sales_review_rows.sort(key=lambda r: (-r["qty"], -r["sales_month"]))

    if not inv_map:
        st.info("ارفع ملف المخزون أولاً من تاب المخزون | Upload Inventory first")
    elif not sales_review_rows:
        st.success("✅ لا توجد SKUs محتاجة مراجعة مبيعات | No SKUs need sales review")
    else:
        df_sales = pd.DataFrame([{
            "SKU": r["sku"], "Yesterday": r["day_counts"].get(e1,0), "Day Before": r["day_counts"].get(e2,0),
            "3 Days Ago": r["day_counts"].get(e3,0), "Stock": r["stock"], "Monthly Sales": r["sales_month"],
            "Days to Stockout (Today's Rate)": r["days_to_stockout_today"]
        } for r in sales_review_rows])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_sales,"sales_review")
        with c2: st.warning(f"📈 SKUs محتاجة مراجعة | Needs Review: {len(sales_review_rows)}")

        pending_approval_skus_t13 = get_pending_approval_skus()

        for r in sales_review_rows:
            c_img,c_info = st.columns([1,6])
            with c_img: show_img(r["img"],70)
            with c_info:
                st.markdown(f"**SKU:** `{r['sku']}`")
                tc_badge_t13 = warehouse_available_badge(r["sku_up"])
                if tc_badge_t13:
                    st.markdown(tc_badge_t13, unsafe_allow_html=True)
                st.markdown(f"📦 **المخزون | Stock:** {r['stock']} &nbsp;|&nbsp; 📈 **مبيع شهري | Monthly:** {r['sales_month']}")
                st.markdown("🛒 " + render_day_counts_md(r["day_counts"], day_dates2, day_labels2))
                st.markdown(f"⚡ **نفاد خلال بيع اليوم | Days to stockout (today's rate):** {r['days_to_stockout_today']} يوم")
                st.markdown(f'<span class="status-badge-lg" style="background:{r["_cov_badge_color"]};">{r["_cov_badge_text"]}</span>', unsafe_allow_html=True)
                if r["sku_up"] in pending_approval_skus_t13:
                    st.markdown(pending_approval_badge_html(), unsafe_allow_html=True)
                render_recent_expired_note(r["sku"])
                for note in get_unavailable_ordered_note(r["sku"]):
                    st.markdown(big_note_html(note), unsafe_allow_html=True)
            st.divider()

    # ══ قايمة "منتهي بالكامل" — بنستبعد منها السكيوهات المجدولة مؤخراً وكمان اللي ليها
    #    جدولة هتوصل قبل النفاد أصلاً (أموره تمام، مش محتاجة مراجعة) ══
    missing_rows_t13 = compute_missing_inventory_rows(day_dates2)
    filtered_missing_t13 = []
    for r in missing_rows_t13:
        if r["sku_up"] in recent_sched_map_t13_all:
            continue
        cov_badge_text, cov_badge_color, _ = schedule_coverage_badge(r["sku"], 0, delay_days2)
        if cov_badge_color == "#22c55e":
            continue
        r["_cov_badge_text"], r["_cov_badge_color"] = cov_badge_text, cov_badge_color
        filtered_missing_t13.append(r)
    missing_rows_t13 = filtered_missing_t13

    st.divider()
    st.subheader("⛔ مخزون منتهي بالكامل | Completely Out of Stock")
    st.caption("SKUs باعت في آخر 3 أيام لكن مالهاش سجل في ملف المخزون أصلاً — يبقى مخزونها انتهى وخرجت من الملف | SKUs with sales in the last 3 days but no record in the Inventory file at all — stock fully ran out")
    if not missing_rows_t13:
        st.success("✅ لا يوجد SKUs خارجة عن المخزون | No SKUs missing from inventory")
    else:
        df_miss13 = pd.DataFrame([{
            "SKU": r["sku"], "Yesterday": r["day_counts"].get(e1,0), "Day Before": r["day_counts"].get(e2,0),
            "3 Days Ago": r["day_counts"].get(e3,0), "Estimated Monthly Sales": r["est_monthly_sales"]
        } for r in missing_rows_t13])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_miss13,"out_of_stock", key="dlbtn_oos_t13")
        with c2: st.error(f"⛔ SKUs منتهية | Out of Stock: {len(missing_rows_t13)}")
        for r in missing_rows_t13:
            c_img,c_info = st.columns([1,6])
            with c_img: show_img(r["img"],70)
            with c_info:
                st.markdown(f"**SKU:** `{r['sku']}`")
                tc_badge_t13b = warehouse_available_badge(r["sku_up"])
                if tc_badge_t13b:
                    st.markdown(tc_badge_t13b, unsafe_allow_html=True)
                st.error("⛔ مخزونه انتهى — مش موجود في ملف المخزون | Stock ran out — not found in inventory file")
                st.markdown("🛒 " + render_day_counts_md(r["day_counts"], day_dates2, day_labels2))
                st.markdown(f"📈 **مبيع شهري تقديري (بناءً على آخر 3 أيام) | Estimated Monthly Sales (based on last 3 days):** **{r['est_monthly_sales']}**")
                st.markdown(f'<span class="status-badge-lg" style="background:{r["_cov_badge_color"]};">{r["_cov_badge_text"]}</span>', unsafe_allow_html=True)
                if r["sku_up"] in pending_approval_skus_t13:
                    st.markdown(pending_approval_badge_html(), unsafe_allow_html=True)
                render_recent_expired_note(r["sku"])
                for note in get_unavailable_ordered_note(r["sku"]):
                    st.markdown(big_note_html(note), unsafe_allow_html=True)
            st.divider()

# ══ TAB 14 — المبيعات ══
with tab14:
    st.subheader("🛒 المبيعات اليومية | Daily Sales")
    st.caption("كل SKU عنده مخزون — مبيعاته اليومية من أمس للوراء بجانب مخزونه ومبيعاته الشهرية وحالة التغطية | All SKUs with inventory — daily sales, stock, monthly sales, and coverage status")
    render_tacweed_upload("sales")
    render_warehouse_stock_upload("sales")

    sales_display_days = int(load_settings().get("sales_display_days","7") or 7)
    today_t14 = datetime.now().date()
    sales_dates = [today_t14 - timedelta(days=i) for i in range(1, sales_display_days + 1)]
    sales_labels = []
    for i, d in enumerate(sales_dates):
        if i == 0:
            sales_labels.append(f"أمس ({d.strftime('%m-%d')})")
        elif i == 1:
            sales_labels.append(f"أول أمس ({d.strftime('%m-%d')})")
        else:
            sales_labels.append(f"قبل {i+1} أيام ({d.strftime('%m-%d')})")

    delay_days_t14 = int(load_settings().get("schedule_delay_days","3") or 3)
    coverage_days_t14 = int(load_settings().get("schedule_coverage_days","15") or 15)

    if not inv_map:
        st.info("ارفع ملف المخزون أولاً من تاب المخزون | Upload Inventory first")
    else:
        multi_counts_t14 = build_daily_orders_counts(sales_dates)
        prices_map_t14   = build_daily_orders_prices(sales_dates)

        # بناء صفوف — كل SKU موجود في المخزون
        sales_tab_rows = []
        for sku_up, info in inv_map.items():
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
            sales_tab_rows.append({
                "sku": sku_disp, "sku_up": sku_up,
                "stock": stock, "sales_month": sales_month, "img": img,
                "day_counts": day_counts, "day_prices": day_prices,
                "total_recent": total_recent,
                "effective_avg": effective_avg_t14,
                "days_to_stockout": days_to_stockout_t14,
            })

        # ترتيب: الأكتر مبيعاً أمس أولاً
        sales_tab_rows.sort(key=lambda r: -r["day_counts"].get(sales_dates[0], 0) if sales_dates else 0)

        # ══ إجماليات اليومية في الأعلى ══
        totals_per_day = {d: sum(r["day_counts"].get(d, 0) for r in sales_tab_rows) for d in sales_dates}
        st.markdown("#### 📊 إجمالي المبيعات اليومية | Daily Sales Totals")
        total_cols = st.columns(min(len(sales_dates), sales_display_days))
        for ci, (d, lbl) in enumerate(zip(sales_dates, sales_labels)):
            if ci < len(total_cols):
                with total_cols[ci]:
                    day_total = totals_per_day.get(d, 0)
                    is_yesterday = (ci == 0)
                    if is_yesterday:
                        bg    = "#14532d" if day_total > 0 else "#7f1d1d"
                        num_color = "#86efac" if day_total > 0 else "#fca5a5"
                        border = "border:2px solid #22c55e;" if day_total > 0 else "border:2px solid #ef4444;"
                    else:
                        bg    = "#1e293b" if day_total == 0 else "#172554"
                        num_color = "#93c5fd" if day_total > 0 else "#64748b"
                        border = ""
                    st.markdown(
                        f'<div style="background:{bg};border-radius:8px;padding:8px 10px;text-align:center;margin:2px;{border}">' +
                        f'<div style="font-size:11px;color:#94a3b8;">{"🔴 " if is_yesterday and day_total==0 else ("🟢 " if is_yesterday else "")}{lbl.split("(")[0].strip()}</div>' +
                        f'<div style="font-size:13px;color:#6b7280;">{d.strftime("%m-%d")}</div>' +
                        f'<div style="font-size:{"28" if is_yesterday else "22"}px;font-weight:bold;color:{num_color};">{day_total}</div>' +
                        '</div>',
                        unsafe_allow_html=True)
        st.divider()

        srch_t14 = st.text_input("🔍 بحث SKU | Search SKU", key="srch_t14", placeholder="اكتب SKU...")
        if srch_t14.strip():
            sales_tab_rows = [r for r in sales_tab_rows if srch_t14.strip().upper() in r["sku_up"]]

        # جدول تحميل
        if sales_tab_rows:
            df_t14 = pd.DataFrame([
                {"SKU": r["sku"], **{sales_labels[i]: r["day_counts"].get(d, 0) for i, d in enumerate(sales_dates)},
                 "مخزون | Stock": r["stock"], "مبيع شهري | Monthly Sales": r["sales_month"]}
                for r in sales_tab_rows
            ])
            c1, c2 = st.columns(2)
            with c1: dl_btn(df_t14, "sales_daily", key="dlbtn_t14")
            with c2: st.info(f"📦 SKUs: {len(sales_tab_rows)} | 📅 {sales_display_days} يوم")

        # ══ قائمة SKUs المرحلة من المبيعات (محتاج جدولة فقط) ══
        # تحديث المرحلين بعد بناء الصفوف الكاملة
        _new_transferred = []

        # ══ خريطة SKUs المجدولة خلال آخر 4 أيام (لعرض ASN + الكمية لو فعلاً ليها جدولة) ══
        recent_sched_map_t14 = get_recent_schedule_rows(days_back=4)
        pending_approval_skus_t14 = get_pending_approval_skus()
        ads_map_t14 = get_ads_map()
        com_map_t14 = get_com_map()

        st.divider()
        for r in sales_tab_rows:
            c_img, c_info = st.columns([1, 7])
            with c_img:
                show_img(r["img"], 70)
            with c_info:
                st.markdown(f"**SKU:** {sku_link_html(r['sku'])}", unsafe_allow_html=True)
                tc_badge_t14 = warehouse_available_badge(r["sku_up"])
                if tc_badge_t14:
                    st.markdown(tc_badge_t14, unsafe_allow_html=True)

                # ══ أداء الإعلانات (لو الـ SKU ده معلن عليه) | Ads performance (if advertised) ══
                ads_entries_t14 = ads_map_t14.get(r["sku_up"])
                if ads_entries_t14:
                    total_spends_t14  = sum(a["spends"] for a in ads_entries_t14)
                    total_revenue_t14 = sum(a["revenue"] for a in ads_entries_t14)
                    total_orders_t14  = sum(a["orders"] for a in ads_entries_t14)
                    with st.expander(f"📢 أداء الإعلانات | Ads Performance — {len(ads_entries_t14)} حملة | campaign(s)"):
                        st.markdown(
                            f"💸 **إجمالي المصروف:** {total_spends_t14:,.2f} ريال &nbsp;|&nbsp; "
                            f"💰 **إجمالي الإيراد:** {total_revenue_t14:,.2f} ريال &nbsp;|&nbsp; "
                            f"🛒 **طلبات من الإعلان:** {total_orders_t14:,.0f}")
                        for ad in ads_entries_t14:
                            st.markdown(
                                f"**{ad['campaign']}**<br>"
                                f"👁️ ظهور: {ad['views']:,.0f} &nbsp;|&nbsp; 🖱️ نقرات: {ad['clicks']:,.0f} &nbsp;|&nbsp; "
                                f"🛒 طلبات: {ad['orders']:,.0f} &nbsp;|&nbsp; ➕ سلة: {ad['atc']:,.0f}<br>"
                                f"💸 مصروف: {ad['spends']:,.2f} ريال &nbsp;|&nbsp; 💰 إيراد: {ad['revenue']:,.2f} ريال<br>"
                                f"📊 CTR: {ad['ctr']:.2f}% &nbsp;|&nbsp; 🎯 ROAS: {ad['roas']:.2f} &nbsp;|&nbsp; "
                                f"CPC: {ad['cpc']:.2f} &nbsp;|&nbsp; CPS: {ad['cps']:.2f} &nbsp;|&nbsp; CVR: {ad['cvr']:.2f}%",
                                unsafe_allow_html=True)
                            st.markdown("<hr style='margin:6px 0;'>", unsafe_allow_html=True)

                # ══ صافي سعر البيع بعد العمولة والتوصيل والضريبة | Net price after commission,
                #    delivery fees, and VAT ══
                com_info_t14 = com_map_t14.get(r["sku_up"])
                if com_info_t14:
                    latest_price_t14 = get_latest_sku_price(r, sales_dates)
                    if latest_price_t14 is not None:
                        net_fees_t14, net_tax_t14 = compute_net_price_after_fees(latest_price_t14, com_info_t14)
                        st.markdown(
                            f'<div style="background:#1e293b;border:1px solid #334155;border-radius:8px;'
                            f'padding:8px 14px;margin:4px 0;">'
                            f'<span style="color:#e2e8f0;font-size:13px;">💵 سعر البيع: <b>{latest_price_t14:,.2f}</b> ريال '
                            f'&nbsp;|&nbsp; 🚚 توصيل: <b>{com_info_t14["delivery"]:,.0f}</b> '
                            f'&nbsp;|&nbsp; 🏷️ عمولة: <b>{com_info_t14["commission_pct"]:,.0f}%</b></span><br>'
                            f'<span style="color:#4ade80;font-size:14px;font-weight:bold;">💳 الصافي بعد خصم العمولة والتوصيل: {net_fees_t14:,.2f} ريال</span><br>'
                            f'<span style="color:#fbbf24;font-size:14px;font-weight:bold;">🧾 الصافي بعد خصم 15% ضريبة: {net_tax_t14:,.2f} ريال</span>'
                            f'</div>',
                            unsafe_allow_html=True)
                    else:
                        st.caption("ℹ️ فيه عمولة وتوصيل مسجلين لكن مفيش سعر بيع حديث لحساب الصافي منهم | Commission & delivery are set but no recent price found to calculate the net")

                # ══ أمس بارز ══
                yesterday_t14 = sales_dates[0] if sales_dates else None
                yesterday_cnt = r["day_counts"].get(yesterday_t14, 0) if yesterday_t14 else 0
                yesterday_prices = r["day_prices"].get(yesterday_t14, []) if yesterday_t14 else []

                def fmt_prices(prices_list):
                    """يجمع الأسعار ويرتبها من الأعلى للأقل، ويتجاهل الفاضي.
                    prices_list: قائمة من (price_str, qty) أو strings."""
                    pc = {}  # price_str -> (total_qty, float_val)
                    for item in prices_list:
                        if isinstance(item, tuple):
                            p, qty = item
                        else:
                            p, qty = item, 1
                        if not p or str(p).strip().lower() in ("","nan","none"):
                            # لو مفيش سعر، نعد الكمية بس بدون سعر
                            pc["__no_price__"] = (pc.get("__no_price__",(0,0))[0] + qty, -1)
                            continue
                        p_str = str(p).strip()
                        try:
                            key = float(p_str.replace(",",""))
                        except Exception:
                            key = 0.0
                        prev_qty, _ = pc.get(p_str, (0, key))
                        pc[p_str] = (prev_qty + qty, key)
                    if not pc:
                        return ""
                    # ترتيب من السعر الأعلى للأقل
                    sorted_prices = sorted(pc.items(), key=lambda x: -x[1][1])
                    parts = []
                    for price_str, (total_qty, _) in sorted_prices:
                        if price_str == "__no_price__":
                            parts.append(f"{total_qty}")
                        else:
                            parts.append(f"{total_qty} × {price_str}")
                    return " | ".join(parts)

                def get_min_max_price(prices_list):
                    """يرجع (أقل سعر, أعلى سعر) كـ float من قائمة (price_str, qty)."""
                    vals = []
                    for item in prices_list:
                        p = item[0] if isinstance(item, tuple) else item
                        if p and str(p).strip().lower() not in ("","nan","none"):
                            try:
                                vals.append(float(str(p).replace(",","")))
                            except Exception:
                                pass
                    if not vals:
                        return None, None
                    return min(vals), max(vals)

                if yesterday_t14:
                    if yesterday_cnt > 0:
                        prices_str_y = fmt_prices(yesterday_prices)
                        min_p_y, max_p_y = get_min_max_price(yesterday_prices)
                        price_lines_y = prices_str_y.split(" | ") if prices_str_y else []
                        price_html_y = ""
                        if price_lines_y:
                            price_html_y = "<br>" + "<br>".join(
                                f'<span style="color:#bbf7d0;font-size:14px;font-weight:bold;">↳ {line}</span>'
                                for line in price_lines_y
                            )
                        minmax_html_y = ""
                        if min_p_y is not None and max_p_y is not None and min_p_y != max_p_y:
                            minmax_html_y = (
                                f'<br><span style="color:#fbbf24;font-size:14px;font-weight:bold;">📉 أقل: {min_p_y:g} &nbsp;|&nbsp; 📈 أعلى: {max_p_y:g}</span>'
                            )
                        elif min_p_y is not None:
                            minmax_html_y = f'<br><span style="color:#fbbf24;font-size:14px;font-weight:bold;">💰 سعر: {min_p_y:g}</span>'
                        yesterday_html = (
                            f'<div style="background:#14532d;border:2px solid #22c55e;border-radius:8px;padding:8px 14px;margin:4px 0;display:inline-block;">' +
                            f'<span style="color:#86efac;font-size:15px;font-weight:bold;">🟢 أمس: {yesterday_cnt}</span>' +
                            minmax_html_y +
                            price_html_y +
                            '</div>'
                        )
                    else:
                        yesterday_html = (
                            '<div style="background:#7f1d1d;border:2px solid #ef4444;border-radius:8px;padding:8px 14px;margin:4px 0;display:inline-block;">' +
                            '<span style="color:#fca5a5;font-size:15px;font-weight:bold;">🔴 أمس: 0</span>' +
                            '</div>'
                        )
                    st.markdown(yesterday_html, unsafe_allow_html=True)

                # باقي الأيام — كل يوم في سطر مع الأسعار تنازلياً + أعلى/أقل
                day_parts = []
                for i, d in enumerate(sales_dates):
                    if i == 0:
                        continue  # أمس اتعرض فوق
                    cnt = r["day_counts"].get(d, 0)
                    day_prices_list = r["day_prices"].get(d, [])
                    color = "#000000" if cnt > 0 else "#475569"
                    lbl_short = sales_labels[i].split("(")[0].strip()
                    prices_str_d = fmt_prices(day_prices_list)
                    min_p_d, max_p_d = get_min_max_price(day_prices_list)
                    minmax_d = ""
                    if min_p_d is not None and max_p_d is not None and min_p_d != max_p_d:
                        minmax_d = f' <span style="color:#b45309;font-size:13px;font-weight:bold;">(📉{min_p_d:g}–📈{max_p_d:g})</span>'
                    elif min_p_d is not None:
                        minmax_d = f' <span style="color:#b45309;font-size:13px;font-weight:bold;">({min_p_d:g})</span>'
                    if prices_str_d:
                        price_lines_d = prices_str_d.split(" | ")
                        price_detail = " &nbsp; ".join(
                            f'<span style="color:#1d4ed8;font-size:13px;font-weight:bold;">↳ {line}</span>'
                            for line in price_lines_d
                        )
                        day_parts.append(
                            f'<span style="color:{color};font-size:15px;font-weight:bold;">{lbl_short}: <b>{cnt}</b>{minmax_d}</span>' +
                            f'<br><span style="padding-right:8px;">{price_detail}</span>'
                        )
                    else:
                        day_parts.append(f'<span style="color:{color};font-size:11px;">{lbl_short}: <b>{cnt}</b>{minmax_d}</span>')
                if day_parts:
                    st.markdown("<br>".join(day_parts), unsafe_allow_html=True)

                # مخزون + مبيع شهري
                st.markdown(
                    f"📦 **مخزون:** {r['stock']} &nbsp;|&nbsp; "
                    f"📈 **شهري:** {r['sales_month']} &nbsp;|&nbsp; "
                    f"📊 **يومي أخير:** {r['effective_avg']:.1f} &nbsp;|&nbsp; "
                    f"⏳ **نفاد خلال:** {r['days_to_stockout'] if r['days_to_stockout'] < 9999 else '—'} يوم"
                )

                # ══ حالة التغطية ══
                badge_text_t14, badge_color_t14, sched_t14 = schedule_coverage_badge(r["sku"], r["days_to_stockout"], delay_days_t14)
                # لو مفيش أي متوسط بيع (لا حديث ولا شهري) يبقى المخزون مش بينزل خالص —
                # فمينفعش نعتبره "غير كافٍ" لمجرد إن مفيش بيانات بيع (كان ده الخلل قبل كده)
                stock_self_ok = (r["effective_avg"] <= 0) or (r["days_to_stockout"] >= coverage_days_t14)
                un_notes = get_unavailable_ordered_note(r["sku"])
                # ══ مجدولة خلال آخر 4 أيام؟ (بنحسبها الأول عشان نستخدمها في قرار عرض شارة التغطية والترحيل) ══
                recent_sched_t14 = recent_sched_map_t14.get(r["sku_up"])

                if stock_self_ok and not sched_t14:
                    if r["effective_avg"] <= 0:
                        cov_badge_text = "✅ لا توجد مبيعات حالياً — لا يحتاج جدولة | No sales recorded — no scheduling needed"
                    else:
                        cov_badge_text = f"✅ مخزون كافٍ ({r['days_to_stockout']} يوم) — لا يحتاج جدولة الآن | Stock sufficient"
                    cov_badge_color = "#15803d"
                elif stock_self_ok and sched_t14:
                    sched_src_t14 = "تشييك" if sched_t14.get("source") == "Check" else "مجدول"
                    arrival_t14 = (sched_t14["parsed"] + timedelta(days=delay_days_t14)).date() if sched_t14.get("parsed") else None
                    stockout_disp_t14 = f"{r['days_to_stockout']} يوم" if r["effective_avg"] > 0 else "لا توجد مبيعات"
                    cov_badge_text  = (f"✅ مخزون كافٍ ({stockout_disp_t14}) + ASN {sched_t14['asn']} بتاريخ {sched_t14['date']}"
                                       + (f" — وصول: {arrival_t14}" if arrival_t14 else "") + f" [{sched_src_t14}]")
                    cov_badge_color = "#15803d"
                else:
                    cov_badge_text  = badge_text_t14
                    cov_badge_color = badge_color_t14

                # لو السكو فعلاً ليه جدولة خلال آخر 4 أيام، منعرضش شارة "محتاج جدولة الآن"
                # المتناقضة جنب شارة الجدولة الحديثة (البنفسجي) اللي هتتعرض تحت — عشان محدش يتلخبط
                show_normal_cov_badge_t14 = not (recent_sched_t14 and "محتاج جدولة" in cov_badge_text)
                if show_normal_cov_badge_t14:
                    st.markdown(
                        f'<span class="status-badge-lg" style="background:{cov_badge_color};">{cov_badge_text}</span>',
                        unsafe_allow_html=True)

                # ══ مجدولة خلال آخر 4 أيام؟ (لو فعلاً ليها جدولة) — تعرض ASN + الكمية + التاريخ ══
                if recent_sched_t14:
                    st.markdown(
                        f'<span class="status-badge-lg" style="background:#7c3aed;">'
                        f'📅 مجدولة خلال آخر 4 أيام | Scheduled in last 4 days — '
                        f'ASN <b>{recent_sched_t14["asn"]}</b> &nbsp;|&nbsp; '
                        f'الكمية | Qty: <b>{recent_sched_t14.get("qty","")}</b> &nbsp;|&nbsp; '
                        f'بتاريخ {recent_sched_t14["date"]} [{recent_sched_t14["source_label"]}]'
                        f'</span>',
                        unsafe_allow_html=True)

                if r["sku_up"] in pending_approval_skus_t14:
                    st.markdown(pending_approval_badge_html(), unsafe_allow_html=True)

                # ══ ترحيل لتاب مخزون بدون بيع إذا كانت الحالة "محتاج جدولة" فقط بدون أي جدولة
                #    وبدون تفاصيل أخرى وبدون جدولة حديثة (آخر 4 أيام) — لو ليه جدولة حديثة
                #    (حتى لو منتهية) يبقى أموره تمام ومينفعش يترحّل لمراجعة المخزون ══
                is_needs_sched_only = (
                    not stock_self_ok
                    and badge_text_t14 and "محتاج جدولة" in badge_text_t14
                    and not sched_t14
                    and not un_notes
                    and not recent_sched_t14
                )
                if is_needs_sched_only:
                    _new_transferred.append({
                        "sku": r["sku"], "sku_up": r["sku_up"], "stock": r["stock"],
                        "sales_month": r["sales_month"], "img": r["img"],
                        "effective_avg": r["effective_avg"], "days_to_stockout": r["days_to_stockout"],
                        "day_counts": r["day_counts"],
                    })
                    st.caption("📌 مرحّل لتاب مراجعة المخزون | Transferred to Stock Review tab")

                if un_notes:
                    for note in un_notes:
                        st.markdown(big_note_html(note), unsafe_allow_html=True)
                render_recent_expired_note(r["sku"])
            st.divider()
        # حفظ المرحلين في session_state بعد اكتمال العرض
        st.session_state["transferred_skus_t14"] = _new_transferred

# ══ TAB DASHBOARD — داشبورد تحليلات المبيعات (تاب جديدة منفصلة) ══
with tab_dash:
    st.header("📊 داشبورد المبيعات | Sales Dashboard")
    st.caption("تحليلات موسّعة على بيانات المبيعات مع صور المنتجات — منفصلة تمامًا عن تاب المبيعات الأصلي ولا تؤثر عليه | Extended sales analytics with product images — fully separate from, and does not affect, the original Sales tab")

    if not inv_map:
        st.info("ارفع ملف المخزون أولاً من تاب المخزون | Upload Inventory first")
    else:
        analysis_period_map_td = {
            "آخر 7 أيام | Last 7 days": 7,
            "آخر 15 يوم | Last 15 days": 15,
            "آخر 30 يوم | Last 30 days": 30,
            "آخر 60 يوم | Last 60 days": 60,
            "آخر 90 يوم | Last 90 days": 90,
        }
        analysis_period_label_td = st.selectbox(
            "🗓️ فترة التحليل | Analysis Period",
            list(analysis_period_map_td.keys()),
            index=2,
            key="dash_period_td")
        analysis_days_td = analysis_period_map_td[analysis_period_label_td]

        today_td  = datetime.now().date()
        cur_dates_td  = [today_td - timedelta(days=i) for i in range(1, analysis_days_td + 1)]
        prev_dates_td = [today_td - timedelta(days=i) for i in range(analysis_days_td + 1, analysis_days_td * 2 + 1)]

        cur_counts_td  = build_daily_orders_counts(cur_dates_td)
        prev_counts_td = build_daily_orders_counts(prev_dates_td)
        cur_prices_td  = build_daily_orders_prices(cur_dates_td)
        prev_prices_td = build_daily_orders_prices(prev_dates_td)

        def _td_total(counts_map, sku_up, dates):
            return sum(counts_map.get(sku_up, {}).get(d, 0) for d in dates)

        def _td_parse_price(p):
            try:
                return float(str(p).replace(",", "").strip())
            except Exception:
                return 0.0

        def _td_revenue_total(prices_map, sku_up, dates):
            total = 0.0
            for d in dates:
                for p, qty in prices_map.get(sku_up, {}).get(d, []):
                    if p and str(p).strip().lower() not in ("", "nan", "none"):
                        total += _td_parse_price(p) * qty
            return total

        rows_td = []
        for sku_up_td, info_td in inv_map.items():
            cur_t_td  = _td_total(cur_counts_td, sku_up_td, cur_dates_td)
            prev_t_td = _td_total(prev_counts_td, sku_up_td, prev_dates_td)
            cur_rev_td  = _td_revenue_total(cur_prices_td, sku_up_td, cur_dates_td)
            prev_rev_td = _td_revenue_total(prev_prices_td, sku_up_td, prev_dates_td)
            rows_td.append({
                "sku_up": sku_up_td, "sku": info_td.get("sku", sku_up_td), "img": info_td.get("img", ""),
                "cur": cur_t_td, "prev": prev_t_td, "stock": info_td.get("total_stock", 0),
                "cur_rev": cur_rev_td, "prev_rev": prev_rev_td,
            })

        total_cur_td  = sum(r["cur"] for r in rows_td)
        total_prev_td = sum(r["prev"] for r in rows_td)
        avg_daily_td  = (total_cur_td / analysis_days_td) if analysis_days_td > 0 else 0
        active_skus_td = sum(1 for r in rows_td if r["cur"] > 0)
        zero_skus_td   = sum(1 for r in rows_td if r["cur"] == 0)
        zero_rows_td   = sorted([r for r in rows_td if r["cur"] == 0], key=lambda r: -r["stock"])
        top_row_td = max(rows_td, key=lambda r: r["cur"], default=None)
        if total_prev_td > 0:
            growth_td = (total_cur_td - total_prev_td) / total_prev_td * 100
        else:
            growth_td = 100.0 if total_cur_td > 0 else 0.0

        total_cur_rev_td  = sum(r["cur_rev"] for r in rows_td)
        total_prev_rev_td = sum(r["prev_rev"] for r in rows_td)
        avg_daily_rev_td  = (total_cur_rev_td / analysis_days_td) if analysis_days_td > 0 else 0
        if total_prev_rev_td > 0:
            growth_rev_td = (total_cur_rev_td - total_prev_rev_td) / total_prev_rev_td * 100
        else:
            growth_rev_td = 100.0 if total_cur_rev_td > 0 else 0.0

        # ── منتجات انخفضت/ارتفعت مبيعاتها (٪20 فأكثر) | Declining / rising SKUs (20%+) ──
        decline_rows_td = sorted(
            [r for r in rows_td if r["prev"] > 0 and r["cur"] < r["prev"]
             and (r["prev"] - r["cur"]) / r["prev"] >= 0.20],
            key=lambda r: -((r["prev"] - r["cur"]) / r["prev"]))
        rise_rows_td = sorted(
            [r for r in rows_td if r["prev"] > 0 and r["cur"] > r["prev"]
             and (r["cur"] - r["prev"]) / r["prev"] >= 0.20],
            key=lambda r: -((r["cur"] - r["prev"]) / r["prev"]))

        # ── أصناف تحتاج انتباه (معرضة لنفاد المخزون) — محسوبة هنا عشان تُستخدم في
        #    شريط التنبيهات السريعة فوق، وبتتعرض بالتفصيل تحت في قسم "أصناف تحتاج انتباه" ──
        delay_days_td = int(load_settings().get("schedule_delay_days", "3") or 3)
        attention_rows_td = []
        for r_td in rows_td:
            avg_d_td = (r_td["cur"] / analysis_days_td) if analysis_days_td > 0 else 0
            if avg_d_td <= 0:
                continue
            days_to_so_td = round(r_td["stock"] / avg_d_td) if avg_d_td > 0 else 9999
            if days_to_so_td > 10:
                continue
            badge_text_td, badge_color_td, sched_td = schedule_coverage_badge(r_td["sku"], days_to_so_td, delay_days_td)
            if "✅" in badge_text_td:
                continue
            status_td = ("⚠️ لديها جدولة — قد تنفد قبل الوصول | Scheduled — may run out before arrival"
                         if sched_td else "🚨 محتاج جدولة الآن | Needs scheduling now")
            attention_rows_td.append({**r_td, "days_to_so": days_to_so_td, "avg_d": avg_d_td, "status": status_td})
        attention_rows_td.sort(key=lambda r: r["days_to_so"])

        # ── كروت رئيسية (نظرة عامة) | Main overview cards ──
        def _kpi_card_html(icon, icon_bg, label, value, delta_text=None, delta_positive=True):
            delta_html = ""
            if delta_text is not None:
                arrow = "↑" if delta_positive else "↓"
                color = "#16a34a" if delta_positive else "#dc2626"
                delta_html = f'<div style="font-size:12px;color:{color};margin-top:6px;font-weight:600;">{arrow} {delta_text}</div>'
            return (
                f'<div style="background:#ffffff;border:1px solid #e5e7eb;border-radius:14px;'
                f'padding:14px 16px;direction:rtl;min-height:118px;box-shadow:0 1px 2px rgba(0,0,0,0.04);">'
                f'<div style="width:34px;height:34px;border-radius:9px;background:{icon_bg}1f;'
                f'display:flex;align-items:center;justify-content:center;font-size:16px;margin-bottom:10px;">{icon}</div>'
                f'<div style="font-size:12px;color:#6b7280;margin-bottom:4px;">{label}</div>'
                f'<div style="font-size:21px;font-weight:800;color:#111827;">{value}</div>'
                f'{delta_html}</div>'
            )

        kc1, kc2, kc3, kc4, kc5 = st.columns(5)
        with kc1:
            st.markdown(_kpi_card_html("📦", "#2563eb", "مبيعات الفترة | Period Sales",
                        f"{total_cur_td:,}", f"{abs(growth_td):.1f}% عن الفترة السابقة", growth_td >= 0), unsafe_allow_html=True)
        with kc2:
            st.markdown(_kpi_card_html("📊", "#0891b2", "متوسط يومي | Daily Avg",
                        f"{avg_daily_td:,.1f}"), unsafe_allow_html=True)
        with kc3:
            st.markdown(_kpi_card_html("💰", "#16a34a", "إجمالي الإيرادات | Total Revenue",
                        f"{total_cur_rev_td:,.0f} ريال", f"{abs(growth_rev_td):.1f}% عن الفترة السابقة", growth_rev_td >= 0), unsafe_allow_html=True)
        with kc4:
            st.markdown(_kpi_card_html("💵", "#9333ea", "متوسط الإيراد اليومي | Daily Avg Revenue",
                        f"{avg_daily_rev_td:,.0f} ريال"), unsafe_allow_html=True)
        with kc5:
            st.markdown(_kpi_card_html("🟢", "#059669", "أصناف نشطة | Active SKUs",
                        f"{active_skus_td:,}"), unsafe_allow_html=True)

        if total_cur_rev_td == 0:
            st.caption("ℹ️ لا توجد أسعار مسجلة لهذه الفترة في شيت الطلبات — الإيراد بيتحسب فقط من الصفوف اللي فيها سعر | No prices recorded for this period — revenue is computed only from rows that include a price")

        st.write("")

        # ── التنبيهات السريعة | Quick alerts strip ──
        st.markdown("##### 🔔 التنبيهات السريعة | Quick Alerts")
        def _alert_chip_html(icon, bg, border, label, value, sub):
            return (
                f'<div style="background:{bg};border:1px solid {border};border-right:4px solid {border};'
                f'border-radius:12px;padding:12px 14px;direction:rtl;min-height:92px;">'
                f'<div style="font-size:12px;color:#374151;margin-bottom:6px;">{icon} {label}</div>'
                f'<div style="font-size:22px;font-weight:800;color:#111827;">{value}</div>'
                f'<div style="font-size:11px;color:#6b7280;">{sub}</div></div>'
            )
        ac1, ac2, ac3, ac4 = st.columns(4)
        with ac1:
            st.markdown(_alert_chip_html("🔴", "#fef2f2", "#ef4444", "منتجات معرضة لنفاد المخزون",
                        f"{len(attention_rows_td):,}", "تغطية أقل من 10 أيام"), unsafe_allow_html=True)
        with ac2:
            st.markdown(_alert_chip_html("🟡", "#fffbeb", "#f59e0b", "منتجات بدون مبيعات في الفترة",
                        f"{zero_skus_td:,}", f"خلال آخر {analysis_days_td} يوم"), unsafe_allow_html=True)
        with ac3:
            st.markdown(_alert_chip_html("🟠", "#fff7ed", "#f97316", "منتجات انخفضت مبيعاتها",
                        f"{len(decline_rows_td):,}", "أكثر من 20% عن الفترة السابقة"), unsafe_allow_html=True)
        with ac4:
            st.markdown(_alert_chip_html("🟢", "#f0fdf4", "#22c55e", "منتجات ارتفعت مبيعاتها",
                        f"{len(rise_rows_td):,}", "أكثر من 20% عن الفترة السابقة"), unsafe_allow_html=True)

        # ── تفاصيل الأصناف تحت كل تنبيه — عشان تبان الـ SKUs نفسها اللي بتكوّن الرقم ──
        def _render_alert_sku_row(r, extra_line=""):
            ci_al, cinfo_al = st.columns([1, 6])
            with ci_al:
                show_img(r["img"], 55)
            with cinfo_al:
                st.markdown(f"{sku_link_html(r['sku'])}", unsafe_allow_html=True)
                if extra_line:
                    st.caption(extra_line)

        with st.expander(f"🔴 عرض منتجات معرضة لنفاد المخزون ({len(attention_rows_td):,}) | Show at-risk SKUs"):
            if attention_rows_td:
                df_al1 = pd.DataFrame([{
                    "SKU": r["sku"], "المخزون | Stock": r["stock"],
                    "متوسط يومي | Daily Avg": round(r["avg_d"], 2),
                    "أيام النفاد | Days to Stockout": r["days_to_so"],
                } for r in attention_rows_td])
                dl_btn(df_al1, "alert_stockout_risk", key="dl_alert_stockout_td")
                for r in attention_rows_td:
                    _render_alert_sku_row(r, f"📦 مخزون: {r['stock']:,} — ⏳ نفاد خلال {r['days_to_so']} يوم")
            else:
                st.caption("لا توجد أصناف معرضة لنفاد المخزون حالياً")

        with st.expander(f"🟡 عرض منتجات بدون مبيعات في الفترة ({zero_skus_td:,}) | Show no-sale SKUs"):
            if zero_rows_td:
                df_al2 = pd.DataFrame([{
                    "SKU": r["sku"], "المخزون | Stock": r["stock"], "مبيعات الفترة | Period Sales": r["cur"],
                } for r in zero_rows_td])
                dl_btn(df_al2, "alert_no_sales", key="dl_alert_nosale_td")
                for r in zero_rows_td:
                    _render_alert_sku_row(r, f"📦 مخزون: {r['stock']:,} — لا يوجد مبيعات خلال {analysis_days_td} يوم")
            else:
                st.caption("كل الأصناف باعت خلال الفترة المحددة")

        with st.expander(f"🟠 عرض منتجات انخفضت مبيعاتها ({len(decline_rows_td):,}) | Show declining SKUs"):
            if decline_rows_td:
                df_al3 = pd.DataFrame([{
                    "SKU": r["sku"], "الفترة الحالية | Current": r["cur"], "الفترة السابقة | Previous": r["prev"],
                    "الانخفاض | Drop %": round((r["prev"] - r["cur"]) / r["prev"] * 100, 1),
                } for r in decline_rows_td])
                dl_btn(df_al3, "alert_declining", key="dl_alert_decline_td")
                for r in decline_rows_td:
                    drop_pct = (r["prev"] - r["cur"]) / r["prev"] * 100
                    _render_alert_sku_row(r, f"📉 {r['cur']:,} مقابل {r['prev']:,} (-{drop_pct:.0f}%)")
            else:
                st.caption("لا توجد أصناف انخفضت مبيعاتها بنسبة 20%+ حالياً")

        with st.expander(f"🟢 عرض منتجات ارتفعت مبيعاتها ({len(rise_rows_td):,}) | Show rising SKUs"):
            if rise_rows_td:
                df_al4 = pd.DataFrame([{
                    "SKU": r["sku"], "الفترة الحالية | Current": r["cur"], "الفترة السابقة | Previous": r["prev"],
                    "الارتفاع | Rise %": round((r["cur"] - r["prev"]) / r["prev"] * 100, 1),
                } for r in rise_rows_td])
                dl_btn(df_al4, "alert_rising", key="dl_alert_rise_td")
                for r in rise_rows_td:
                    rise_pct = (r["cur"] - r["prev"]) / r["prev"] * 100
                    _render_alert_sku_row(r, f"📈 {r['cur']:,} مقابل {r['prev']:,} (+{rise_pct:.0f}%)")
            else:
                st.caption("لا توجد أصناف ارتفعت مبيعاتها بنسبة 20%+ حالياً")

        st.write("")

        # ── كارت أعلى SKU مع صورة | Top SKU card with image ──
        if top_row_td and top_row_td["cur"] > 0:
            st.markdown("##### 🔥 أعلى SKU مبيعًا | Top-Selling SKU")
            ci_top, cinfo_top = st.columns([1, 6])
            with ci_top:
                show_img(top_row_td["img"], 90)
            with cinfo_top:
                st.markdown(f"**`{top_row_td['sku']}`**")
                st.markdown(
                    f"📦 **مبيعات الفترة | Period Sales:** {top_row_td['cur']:,} &nbsp;|&nbsp; "
                    f"📊 **متوسط يومي | Daily Avg:** {(top_row_td['cur']/analysis_days_td):,.1f} &nbsp;|&nbsp; "
                    f"📦 **مخزون | Stock:** {top_row_td['stock']:,}")
                st.markdown(f"💰 **إيراد الفترة | Period Revenue:** {top_row_td['cur_rev']:,.0f} ريال")

        st.divider()

        # ── رسم بياني لاتجاه المبيعات + أهم المنتجات جنب بعض | Trend chart + top products, side by side ──
        cur_dates_sorted_td = sorted(cur_dates_td)
        daily_totals_td = {
            d: sum(cur_counts_td.get(r["sku_up"], {}).get(d, 0) for r in rows_td)
            for d in cur_dates_sorted_td
        }
        chart_df_td = pd.DataFrame({
            "التاريخ | Date": [d.strftime("%Y-%m-%d") for d in cur_dates_sorted_td],
            "المبيعات | Sales": [daily_totals_td.get(d, 0) for d in cur_dates_sorted_td],
        }).set_index("التاريخ | Date")

        col_chart_td, col_top_td = st.columns([1.1, 1])
        with col_chart_td:
            st.markdown(f"##### 📉 اتجاه المبيعات آخر {analysis_days_td} يوم | Sales Trend")
            st.line_chart(chart_df_td)
            growth_icon_td = "📈" if growth_td >= 0 else "📉"
            growth_word_td = "نمو" if growth_td >= 0 else "انخفاض"
            growth_rev_icon_td = "📈" if growth_rev_td >= 0 else "📉"
            growth_rev_word_td = "نمو" if growth_rev_td >= 0 else "انخفاض"
            st.caption(
                f"{growth_icon_td} {growth_word_td} الطلبات: {growth_td:+.1f}% ({total_cur_td:,} مقابل {total_prev_td:,}) "
                f"&nbsp;|&nbsp; {growth_rev_icon_td} {growth_rev_word_td} الإيراد: {growth_rev_td:+.1f}% "
                f"({total_cur_rev_td:,.0f} مقابل {total_prev_rev_td:,.0f} ريال)")

        with col_top_td:
            st.markdown("##### 🏆 أهم المنتجات | Top Products")
            top5_td = sorted(rows_td, key=lambda r: -r["cur"])[:5]
            if top5_td and top5_td[0]["cur"] > 0:
                rows_html_td = ""
                for r in top5_td:
                    if r["cur"] <= 0:
                        continue
                    img_src = r["img"] if r["img"] else ""
                    img_html = (f'<img src="{img_src}" style="width:32px;height:32px;border-radius:6px;'
                                 f'object-fit:cover;margin-left:8px;">') if img_src else "📦"
                    rows_html_td += (
                        '<div style="display:flex;align-items:center;justify-content:space-between;'
                        'padding:8px 4px;border-bottom:1px solid #f1f5f9;direction:rtl;">'
                        f'<div style="display:flex;align-items:center;font-size:12px;color:#111827;">{img_html}'
                        f'<span style="font-family:monospace;">{r["sku"]}</span></div>'
                        f'<div style="text-align:left;font-size:12px;color:#374151;white-space:nowrap;">'
                        f'<b>{r["cur"]:,}</b> طلب &nbsp; <span style="color:#6b7280;">({r["cur_rev"]:,.0f} ريال)</span></div>'
                        '</div>'
                    )
                st.markdown(
                    f'<div style="background:#ffffff;border:1px solid #e5e7eb;border-radius:12px;padding:6px 12px;">{rows_html_td}</div>',
                    unsafe_allow_html=True)
            else:
                st.caption("لا توجد بيانات مبيعات كافية | Not enough sales data")

        st.divider()

        # ── المبيعات حسب القسم (فلوس وعدد) | Sales by Department (revenue & orders) ──
        # بيعتمد على عمود Family (اختياري) في شيت DailyOrders — لو مش موجود أو الصف مالوش
        # قيمة، الكود بيكمل عادي من غير ما يوقف وبيتجاهل هذا الصف من تحليل الأقسام بس
        st.markdown("### 📂 المبيعات حسب القسم | Sales by Department")
        dept_stats_td = build_daily_orders_family_stats(cur_dates_td)
        if not dept_stats_td:
            st.caption("لا توجد بيانات أقسام (عمود Family) لهذه الفترة — العمود اختياري ولا يؤثر على باقي التحليلات | No department (Family) data for this period — the column is optional and does not affect other analytics")
        else:
            dept_sorted_td = sorted(dept_stats_td.items(), key=lambda x: -x[1]["revenue"])
            df_dept_td = pd.DataFrame([{
                "القسم | Department": dept,
                "عدد الطلبات | Orders": v["orders"],
                "الإيراد | Revenue (ريال)": round(v["revenue"], 2),
            } for dept, v in dept_sorted_td])
            dl_btn(df_dept_td, "sales_by_department", key="dl_dept_td")
            st.dataframe(df_dept_td, use_container_width=True, hide_index=True)
            st.bar_chart(df_dept_td.set_index("القسم | Department")[["الإيراد | Revenue (ريال)"]])

        st.divider()

        # ── أعلى 10 أصناف مبيعًا (مع صور) | Top 10 best sellers (with images) ──
        st.markdown("### 🔥 أعلى 10 أصناف مبيعًا | Top 10 Best Sellers")
        top10_td = sorted([r for r in rows_td if r["cur"] > 0], key=lambda r: -r["cur"])[:10]
        if top10_td:
            df_top10_td = pd.DataFrame([{
                "الترتيب | #": i + 1, "SKU": r["sku"], "المبيعات | Sales": r["cur"],
                "متوسط يومي | Daily Avg": round(r["cur"] / analysis_days_td, 2), "المخزون | Stock": r["stock"],
                "الإيراد | Revenue (ريال)": round(r["cur_rev"], 2),
            } for i, r in enumerate(top10_td)])
            dl_btn(df_top10_td, "top_sellers", key="dl_top10_td")
            for i, r in enumerate(top10_td):
                ci_t, cinfo_t = st.columns([1, 6])
                with ci_t:
                    show_img(r["img"], 70)
                with cinfo_t:
                    st.markdown(f"**#{i+1} — `{r['sku']}`**")
                    st.markdown(
                        f"📦 مبيعات | Sales: **{r['cur']:,}** &nbsp;|&nbsp; "
                        f"📊 يومي | Daily: **{r['cur']/analysis_days_td:.1f}** &nbsp;|&nbsp; "
                        f"📦 مخزون | Stock: **{r['stock']:,}** &nbsp;|&nbsp; "
                        f"💰 إيراد | Revenue: **{r['cur_rev']:,.0f} ريال**")
                st.divider()
        else:
            st.caption("لا توجد بيانات مبيعات كافية | Not enough sales data")

        # ── الأصناف البطيئة (مع صور) | Slow moving items (with images) ──
        st.markdown("### 🐌 الأصناف البطيئة | Slow Moving (Bottom 10)")
        slow10_td = sorted(rows_td, key=lambda r: r["cur"])[:10]
        if slow10_td:
            df_slow10_td = pd.DataFrame([{
                "SKU": r["sku"], "المبيعات | Sales": r["cur"],
                "متوسط يومي | Daily Avg": round(r["cur"] / analysis_days_td, 2), "المخزون | Stock": r["stock"],
                "الإيراد | Revenue (ريال)": round(r["cur_rev"], 2),
            } for r in slow10_td])
            dl_btn(df_slow10_td, "slow_movers", key="dl_slow10_td")
            for r in slow10_td:
                ci_s, cinfo_s = st.columns([1, 6])
                with ci_s:
                    show_img(r["img"], 70)
                with cinfo_s:
                    st.markdown(f"**`{r['sku']}`**")
                    st.markdown(
                        f"📦 مبيعات | Sales: **{r['cur']:,}** &nbsp;|&nbsp; "
                        f"📊 يومي | Daily: **{r['cur']/analysis_days_td:.1f}** &nbsp;|&nbsp; "
                        f"📦 مخزون | Stock: **{r['stock']:,}** &nbsp;|&nbsp; "
                        f"💰 إيراد | Revenue: **{r['cur_rev']:,.0f} ريال**")
                st.divider()

        # ── تحليل اتجاه SKU فردي (مع صورة) | Per-SKU trend (with image) ──
        st.markdown("### 🔎 تحليل اتجاه SKU | SKU Trend Analysis")
        sku_options_td = sorted({r["sku"] for r in rows_td})
        selected_sku_td = st.selectbox("اختر SKU للتحليل | Select SKU", ["—"] + sku_options_td, key="dash_sku_td")
        if selected_sku_td and selected_sku_td != "—":
            sel_row_td = next((r for r in rows_td if r["sku"] == selected_sku_td), None)
            if sel_row_td:
                sel_daily_td = {d: cur_counts_td.get(sel_row_td["sku_up"], {}).get(d, 0) for d in cur_dates_sorted_td}
                sel_cur_td, sel_prev_td = sel_row_td["cur"], sel_row_td["prev"]
                sel_cur_rev_td, sel_prev_rev_td = sel_row_td["cur_rev"], sel_row_td["prev_rev"]
                sel_avg_td = (sel_cur_td / analysis_days_td) if analysis_days_td > 0 else 0
                if sel_prev_td > 0:
                    sel_growth_td = (sel_cur_td - sel_prev_td) / sel_prev_td * 100
                else:
                    sel_growth_td = 100.0 if sel_cur_td > 0 else 0.0
                if sel_prev_rev_td > 0:
                    sel_growth_rev_td = (sel_cur_rev_td - sel_prev_rev_td) / sel_prev_rev_td * 100
                else:
                    sel_growth_rev_td = 100.0 if sel_cur_rev_td > 0 else 0.0
                max_day_td = max(sel_daily_td.items(), key=lambda x: x[1], default=(None, 0))
                min_day_td = min(sel_daily_td.items(), key=lambda x: x[1], default=(None, 0))

                ci_sel, cinfo_sel = st.columns([1, 6])
                with ci_sel:
                    show_img(sel_row_td["img"], 90)
                with cinfo_sel:
                    st.markdown(f"**`{sel_row_td['sku']}`** &nbsp;|&nbsp; 📦 مخزون | Stock: **{sel_row_td['stock']:,}**")

                m1_td, m2_td, m3_td = st.columns(3)
                m1_td.metric("مبيعات الفترة | Period Sales", f"{sel_cur_td:,}")
                m2_td.metric("متوسط يومي | Daily Avg", f"{sel_avg_td:,.1f}")
                m3_td.metric("النمو | Growth", f"{sel_growth_td:+.2f}%")
                st.caption(f"مبيعات الفترة السابقة | Previous period sales: **{sel_prev_td:,}**")

                mr1_td, mr2_td, mr3_td = st.columns(3)
                mr1_td.metric("إيراد الفترة | Period Revenue", f"{sel_cur_rev_td:,.0f} ريال")
                mr2_td.metric("متوسط إيراد يومي | Daily Avg Revenue", f"{(sel_cur_rev_td/analysis_days_td if analysis_days_td>0 else 0):,.0f} ريال")
                mr3_td.metric("نمو الإيراد | Revenue Growth", f"{sel_growth_rev_td:+.2f}%")
                st.caption(f"إيراد الفترة السابقة | Previous period revenue: **{sel_prev_rev_td:,.0f} ريال**")

                m4_td, m5_td = st.columns(2)
                with m4_td:
                    st.markdown(f"📈 **أعلى يوم مبيعات | Best Day:** "
                                f"{max_day_td[0].strftime('%Y-%m-%d') if max_day_td[0] else '—'} ({max_day_td[1]})")
                with m5_td:
                    st.markdown(f"📉 **أقل يوم مبيعات | Worst Day:** "
                                f"{min_day_td[0].strftime('%Y-%m-%d') if min_day_td[0] else '—'} ({min_day_td[1]})")
                sku_chart_df_td = pd.DataFrame({
                    "التاريخ | Date": [d.strftime("%Y-%m-%d") for d in cur_dates_sorted_td],
                    "المبيعات | Sales": [sel_daily_td.get(d, 0) for d in cur_dates_sorted_td],
                }).set_index("التاريخ | Date")
                st.line_chart(sku_chart_df_td)

        st.divider()

        # ── أصناف تحتاج انتباه (مع صور) | Needs-attention (with images) ──
        # القائمة اتحسبت فوق قبل شريط التنبيهات السريعة (attention_rows_td) — هنا بس بنعرضها بالتفصيل
        # Already computed above (before the quick-alerts strip) — this just renders the details
        st.markdown("### 🚨 أصناف تحتاج انتباه | Needs Attention")
        if attention_rows_td:
            df_att_td = pd.DataFrame([{
                "SKU": r["sku"], "المخزون | Stock": r["stock"],
                "متوسط يومي | Daily Avg": round(r["avg_d"], 2),
                "أيام النفاد المتوقعة | Days to Stockout": r["days_to_so"],
                "الحالة | Status": r["status"],
            } for r in attention_rows_td])
            dl_btn(df_att_td, "needs_attention", key="dl_attention_td")
            for r in attention_rows_td:
                ci_a, cinfo_a = st.columns([1, 6])
                with ci_a:
                    show_img(r["img"], 70)
                with cinfo_a:
                    st.markdown(f"**`{r['sku']}`**")
                    st.markdown(
                        f"📦 مخزون | Stock: **{r['stock']:,}** &nbsp;|&nbsp; "
                        f"📊 يومي | Daily Avg: **{r['avg_d']:.1f}** &nbsp;|&nbsp; "
                        f"⏳ نفاد خلال | Stockout in: **{r['days_to_so']}** يوم")
                    st.markdown(r["status"])
                st.divider()
        else:
            st.success("✅ لا توجد أصناف محتاجة انتباه حالياً | No items currently need attention")

# ══ TAB 15 — تحليل الجدولة ══
with tab15:
    st.subheader("🗓️ تحليل الجدولة المقترحة | Schedule Analysis")
    st.caption("ارفع أو الصق SKUs وهيجيلك تحليل جدولات مستقبلية مقترحة بناءً على المخزون والمبيعات والجدولات الحالية | Upload or paste SKUs to get suggested future schedules based on stock, sales and existing schedules")

    if not inv_map:
        st.info("ارفع ملف المخزون أولاً من تاب المخزون | Upload Inventory first")
    else:
        method_t15 = st.radio("طريقة الإدخال | Input Method:", ["📂 رفع ملف | Upload", "✏️ لصق | Paste"], horizontal=True, key="method_t15")
        analysis_skus = []

        if "Upload" in method_t15:
            upl_t15 = st.file_uploader("ارفع Excel أو CSV (عمود SKU) | Upload Excel or CSV with SKU column", type=["xlsx","xls","csv"], key="upl_t15")
            if upl_t15:
                try:
                    df_t15_up = pd.read_csv(upl_t15, dtype=str).fillna("") if upl_t15.name.endswith(".csv") else pd.read_excel(upl_t15, dtype=str).fillna("")
                    sku_col_t15 = None
                    for c in df_t15_up.columns:
                        if "sku" in c.strip().lower() or "item" in c.strip().lower():
                            sku_col_t15 = c; break
                    if not sku_col_t15:
                        sku_col_t15 = df_t15_up.columns[0]
                    analysis_skus = [str(r[sku_col_t15]).strip() for _, r in df_t15_up.iterrows()
                                     if str(r[sku_col_t15]).strip() and str(r[sku_col_t15]).strip().lower() != "nan"]
                    st.success(f"✅ {len(analysis_skus)} SKU جاهز | SKUs loaded")
                except Exception as e:
                    st.error(f"❌ {e}")
        else:
            pasted_t15 = st.text_area("الصق SKUs هنا (كل واحد في سطر) | Paste SKUs (one per line):", height=120, key="paste_t15", placeholder="SKU001\nSKU002\nSKU003")
            if pasted_t15.strip():
                analysis_skus = [line.strip() for line in pasted_t15.strip().splitlines() if line.strip()]
                st.success(f"✅ {len(analysis_skus)} SKU | SKUs ready")

        if analysis_skus:
            st.divider()
            today_t15 = datetime.now().date()
            delay_days_t15 = int(load_settings().get("schedule_delay_days","3") or 3)
            sales_days_t15 = int(load_settings().get("sales_display_days","7") or 7)
            coverage_days_t15 = int(load_settings().get("schedule_coverage_days","15") or 15)
            recent_dates_t15 = [today_t15 - timedelta(days=i) for i in range(1, sales_days_t15 + 1)]
            multi_counts_t15 = build_daily_orders_counts(recent_dates_t15)

            st.write(f"**تحليل {len(analysis_skus)} SKU | Analyzing {len(analysis_skus)} SKUs** — أيام التغطية المقترحة: **{coverage_days_t15} يوم**")

            # ══ جمع كل البيانات للإكسيل ══
            excel_rows_t15 = []

            for sku_raw in analysis_skus:
                sku_up = sku_raw.strip().upper()
                info = inv_map.get(sku_up)

                st.markdown(f"### 📦 SKU: `{sku_raw}`")

                if not info:
                    st.error("⛔ هذا الـ SKU مش موجود في المخزون — مخزونه انتهى أو لم يُرفع | Not found in inventory — may be out of stock or not uploaded")
                    day_counts_miss = multi_counts_t15.get(sku_up, {})
                    total_miss = sum(day_counts_miss.values())
                    if total_miss > 0:
                        avg_miss = total_miss / sales_days_t15
                        est_monthly = round(avg_miss * 30)
                        suggested_urgent = round(avg_miss * coverage_days_t15)
                        suggested_urgent = max(suggested_urgent, 1)
                        st.warning(f"📈 باع {total_miss} قطعة في آخر {sales_days_t15} يوم — مبيع شهري تقديري: **{est_monthly}** | Sold {total_miss} units in last {sales_days_t15} days — est. monthly: **{est_monthly}**")
                        urgent_date = today_t15 + timedelta(days=3)
                        st.markdown(
                            f'<div style="background:#1a0000;border:1px solid #ef4444;border-left:5px solid #ef4444;border-radius:8px;padding:10px 14px;color:white;margin:6px 0;">'
                            f'🗓️ <b>جدولة مقترحة عاجلة | Urgent suggested schedule:</b><br>'
                            f'📅 التاريخ المقترح: <b style="color:#fca5a5;">{urgent_date.strftime("%Y-%m-%d")}</b> &nbsp;|&nbsp; '
                            f'📦 الكمية المقترحة ({coverage_days_t15} يوم): <b style="color:#fca5a5;">{suggested_urgent}</b><br>'
                            f'<span style="color:#f87171;font-size:12px;">⚠️ ملاحظة: المنتج خارج المخزون، يُنصح بالجدولة فوراً | Product is out of stock, immediate scheduling recommended</span>'
                            f'</div>', unsafe_allow_html=True)
                        excel_rows_t15.append({
                            "SKU": sku_raw, "المخزون | Stock": 0, "مبيع شهري | Monthly Sales": est_monthly,
                            "متوسط يومي | Daily Avg": round(avg_miss, 2),
                            "نفاد خلال | Days to Stockout": "خلص | Out",
                            "تاريخ الجدولة #1": urgent_date.strftime("%Y-%m-%d"),
                            "وصول #1": (urgent_date + timedelta(days=delay_days_t15)).strftime("%Y-%m-%d"),
                            "كمية #1": suggested_urgent, "ملاحظة #1": "عاجل — مخزون منتهي",
                            "تاريخ الجدولة #2": "", "وصول #2": "", "كمية #2": "", "ملاحظة #2": "",
                            "تاريخ الجدولة #3": "", "وصول #3": "", "كمية #3": "", "ملاحظة #3": "",
                        })
                    else:
                        excel_rows_t15.append({
                            "SKU": sku_raw, "المخزون | Stock": 0, "مبيع شهري | Monthly Sales": 0,
                            "متوسط يومي | Daily Avg": 0, "نفاد خلال | Days to Stockout": "خلص | Out",
                            "تاريخ الجدولة #1": "", "وصول #1": "", "كمية #1": "", "ملاحظة #1": "مخزون منتهي ولا مبيعات",
                            "تاريخ الجدولة #2": "", "وصول #2": "", "كمية #2": "", "ملاحظة #2": "",
                            "تاريخ الجدولة #3": "", "وصول #3": "", "كمية #3": "", "ملاحظة #3": "",
                        })
                    st.divider()
                    continue

                stock       = info.get("total_stock", 0)
                sales_month = info.get("sales", 0)
                img         = info.get("img", "")
                avg_daily   = sales_month / 30 if sales_month > 0 else 0
                day_counts_t15 = multi_counts_t15.get(sku_up, {d: 0 for d in recent_dates_t15})
                recent_total = sum(day_counts_t15.values())
                avg_daily_recent = recent_total / sales_days_t15 if sales_days_t15 > 0 else avg_daily

                effective_avg = avg_daily_recent if avg_daily_recent > 0 else avg_daily
                if effective_avg > 0:
                    days_to_stockout_t15 = round(stock / effective_avg)
                    stockout_date_t15 = today_t15 + timedelta(days=days_to_stockout_t15)
                else:
                    days_to_stockout_t15 = 0
                    stockout_date_t15 = today_t15

                c_img_t15, c_info_t15 = st.columns([1, 6])
                with c_img_t15:
                    show_img(img, 65)
                with c_info_t15:
                    st.markdown(
                        f"📦 **مخزون | Stock:** **{stock}** &nbsp;|&nbsp; "
                        f"📈 **مبيع شهري | Monthly:** **{sales_month}** &nbsp;|&nbsp; "
                        f"📊 **متوسط يومي أخير | Recent daily avg:** **{avg_daily_recent:.1f}**"
                    )
                    if effective_avg > 0:
                        st.markdown(
                            f"⏳ **متوقع النفاد | Estimated stockout:** "
                            f"**{days_to_stockout_t15} يوم** ({stockout_date_t15.strftime('%Y-%m-%d')})"
                        )
                    else:
                        st.caption("⚠️ لا توجد مبيعات مسجلة — لا يمكن تقدير يوم النفاد | No sales data — cannot estimate stockout")

                # الجدولات الحالية
                existing_schedules = []
                for sheet_key in ("Scheduled", "Check"):
                    sdata = get_scheduled_normalized() if sheet_key == "Scheduled" else get_cached(sheets[sheet_key])
                    if len(sdata) <= 1:
                        continue
                    for row in sdata[1:]:
                        while len(row) < 4: row.append("")
                        if row[1].strip().upper() == sku_up:
                            d_parsed = parse_excel_date(row[3])
                            existing_schedules.append({
                                "asn": row[0], "qty": row[2], "date": row[3],
                                "parsed": d_parsed, "source": sheet_key
                            })
                existing_schedules.sort(key=lambda s: s["parsed"] or datetime.max)

                if existing_schedules:
                    st.markdown("**📋 الجدولات الحالية | Existing Schedules:**")
                    for es in existing_schedules:
                        arrival_es = (es["parsed"] + timedelta(days=delay_days_t15)).date() if es["parsed"] else None
                        src_label = "تشييك" if es["source"] == "Check" else "مجدول"
                        st.markdown(
                            f'<span style="background:#1e3a5f;color:#93c5fd;border-radius:6px;padding:3px 10px;font-size:12px;margin:2px;">'
                            f'ASN {es["asn"]} | {es["qty"]} قطعة | {es["date"]} | {src_label}'
                            f'{f" | وصول متوقع: {arrival_es}" if arrival_es else ""}'
                            f'</span>',
                            unsafe_allow_html=True)

                st.markdown("---")
                st.markdown(f"**🗓️ الجدولات المقترحة | Suggested Schedules** — كل جدولة تغطي **{coverage_days_t15} يوم** فقط:")

                # تاريخ الوصول الفعلي لآخر جدولة موجودة
                last_covered_date = today_t15
                if existing_schedules:
                    for es in existing_schedules:
                        if es["parsed"]:
                            arr = (es["parsed"] + timedelta(days=delay_days_t15)).date()
                            if arr > last_covered_date:
                                last_covered_date = arr

                # مخزون + جدولات موجودة
                total_incoming = sum(_to_int(es["qty"]) for es in existing_schedules)
                adjusted_stock = stock + total_incoming
                if effective_avg > 0:
                    adjusted_days = round(adjusted_stock / effective_avg)
                    adjusted_stockout = today_t15 + timedelta(days=adjusted_days)
                else:
                    adjusted_days = 999
                    adjusted_stockout = today_t15 + timedelta(days=999)

                if existing_schedules:
                    st.caption(
                        f"📦 بعد الجدولات الحالية | After existing schedules: مخزون فعلي = {adjusted_stock} "
                        f"→ نفاد متوقع بعد {adjusted_days} يوم ({adjusted_stockout.strftime('%Y-%m-%d')})"
                    )

                # ══ توليد الجدولات المقترحة بكميات تغطي coverage_days_t15 فقط ══
                BUFFER_DAYS = 3   # هامش أمان — يبدأ الجدولة قبل النفاد بـ 3 أيام + الـ delay
                suggested_list = []
                running_stock = adjusted_stock
                running_date = last_covered_date

                for sg_i in range(3):
                    if effective_avg <= 0:
                        break
                    # الكمية = متوسط يومي × أيام التغطية فقط
                    suggested_qty = max(round(effective_avg * coverage_days_t15), 1)
                    # متى هيخلص الـ running_stock؟
                    days_until_running_out = round(running_stock / effective_avg) if effective_avg > 0 else 999
                    # موعد الجدولة: قبل النفاد بـ (delay + buffer) يوم على الأقل
                    days_before_arrival_needed = delay_days_t15 + BUFFER_DAYS
                    days_to_next_schedule = max(days_until_running_out - days_before_arrival_needed, 1)
                    target_schedule_date = running_date + timedelta(days=days_to_next_schedule)
                    arrival_date = target_schedule_date + timedelta(days=delay_days_t15)
                    stock_at_arrival = max(round(running_stock - effective_avg * (arrival_date - running_date).days), 0)

                    note = ""
                    if sg_i == 0 and existing_schedules:
                        note = "⚠️ يوجد جدولة حالية — هذا اقتراح الجدولة التالية بعدها | Existing schedule found — this is the NEXT suggested schedule"
                    elif sg_i == 0 and days_to_stockout_t15 <= coverage_days_t15:
                        note = "🔴 المخزون قريب على الخلاص — يُنصح بالجدولة العاجلة | Stock nearly out — urgent scheduling recommended"

                    suggested_list.append({
                        "num": sg_i + 1,
                        "schedule_date": target_schedule_date,
                        "arrival_date": arrival_date,
                        "qty": suggested_qty,
                        "note": note,
                        "stock_at_arrival": stock_at_arrival,
                    })

                    # المخزون بعد وصول هذه الجدولة
                    running_stock = stock_at_arrival + suggested_qty
                    running_date = arrival_date

                colors_sg = ["#14532d", "#1e3a5f", "#3b0764"]
                border_sg = ["#22c55e", "#3b82f6", "#a855f7"]
                for sg in suggested_list:
                    st.markdown(
                        f'<div style="background:{colors_sg[sg["num"]-1]};border:1px solid {border_sg[sg["num"]-1]};'
                        f'border-left:5px solid {border_sg[sg["num"]-1]};border-radius:8px;padding:10px 14px;color:white;margin:6px 0;">'
                        f'🗓️ <b>الجدولة {sg["num"]} | Schedule #{sg["num"]}:</b><br>'
                        f'📅 تاريخ الجدولة المقترح: <b style="color:#86efac;">{sg["schedule_date"].strftime("%Y-%m-%d")}</b>'
                        f' &nbsp;→&nbsp; وصول متوقع: <b style="color:#93c5fd;">{sg["arrival_date"].strftime("%Y-%m-%d")}</b><br>'
                        f'📦 كمية مقترحة ({coverage_days_t15} يوم): <b style="color:#c4b5fd;">{sg["qty"]}</b>'
                        f' &nbsp;|&nbsp; مخزون متوقع عند الوصول: <b>{sg["stock_at_arrival"]}</b><br>'
                        + (f'<span style="color:#fcd34d;font-size:12px;">📝 {sg["note"]}</span>' if sg["note"] else "")
                        + '</div>',
                        unsafe_allow_html=True)

                render_recent_expired_note(sku_raw)
                for note in get_unavailable_ordered_note(sku_raw):
                    st.markdown(big_note_html(note), unsafe_allow_html=True)

                # تجميع صف الإكسيل
                excel_row_t15 = {
                    "SKU": sku_raw,
                    "المخزون | Stock": stock,
                    "مبيع شهري | Monthly Sales": sales_month,
                    "متوسط يومي | Daily Avg": round(effective_avg, 2),
                    f"نفاد خلال | Days to Stockout": days_to_stockout_t15 if effective_avg > 0 else "—",
                }
                for sg in suggested_list:
                    n = sg["num"]
                    excel_row_t15[f"تاريخ الجدولة #{n}"] = sg["schedule_date"].strftime("%Y-%m-%d")
                    excel_row_t15[f"وصول #{n}"] = sg["arrival_date"].strftime("%Y-%m-%d")
                    excel_row_t15[f"كمية #{n}"] = sg["qty"]
                    excel_row_t15[f"مخزون عند الوصول #{n}"] = sg["stock_at_arrival"]
                    excel_row_t15[f"ملاحظة #{n}"] = sg["note"]
                excel_rows_t15.append(excel_row_t15)

                st.divider()

            # ══ زر تحميل الإكسيل ══
            if excel_rows_t15:
                st.divider()
                df_excel_t15 = pd.DataFrame(excel_rows_t15)
                dl_btn(df_excel_t15, "schedule_analysis", label="⬇️ تحميل تحليل الجدولة Excel | Download Schedule Analysis", key="dlbtn_t15_excel")

# ══ TAB 16 — مخزون بدون بيع ══
with tab16:
    st.subheader("📦 مخزون بدون بيع | Stock With No Sales")
    st.caption("SKUs موجودة في المخزون لكن ما بيعت في الفترة المحددة | SKUs in inventory with no sales in the selected period")

    if not inv_map:
        st.info("ارفع ملف المخزون أولاً من تاب المخزون | Upload Inventory first")
    else:
        today_t16 = datetime.now().date()
        sales_display_days_t16 = int(load_settings().get("sales_display_days","7") or 7)

        # تواريخ الفترات الثلاث
        dates_1d  = [today_t16 - timedelta(days=1)]
        dates_3d  = [today_t16 - timedelta(days=i) for i in range(1, 4)]
        dates_7d  = [today_t16 - timedelta(days=i) for i in range(1, 8)]

        all_dates_t16 = list({d for d in dates_1d + dates_3d + dates_7d})
        counts_t16 = build_daily_orders_counts(all_dates_t16)

        def sku_sold_in(sku_up, dates_list):
            dc = counts_t16.get(sku_up, {})
            return sum(dc.get(d, 0) for d in dates_list) > 0

        # بناء القوائم الثلاث
        no_sale_1d, no_sale_3d, no_sale_7d = [], [], []
        for sku_up, info in inv_map.items():
            stock       = info.get("total_stock", 0)
            sales_month = info.get("sales", 0)
            img         = info.get("img", "")
            sku_disp    = info.get("sku", sku_up)
            row_t16 = {"sku": sku_disp, "sku_up": sku_up, "stock": stock, "sales_month": sales_month, "img": img}
            if not sku_sold_in(sku_up, dates_1d):
                no_sale_1d.append(row_t16)
            if not sku_sold_in(sku_up, dates_3d):
                no_sale_3d.append(row_t16)
            if not sku_sold_in(sku_up, dates_7d):
                no_sale_7d.append(row_t16)
        # ترتيب من الأعلى مخزوناً للأقل
        no_sale_1d.sort(key=lambda x: -x["stock"])
        no_sale_3d.sort(key=lambda x: -x["stock"])
        no_sale_7d.sort(key=lambda x: -x["stock"])

        def render_no_sale_list(rows, period_label, dl_key):
            if not rows:
                st.success(f"✅ لا يوجد SKUs بدون مبيعات في {period_label} | No SKUs without sales in {period_label}")
                return
            df_ns = pd.DataFrame([{
                "SKU": r["sku"], "مخزون | Stock": r["stock"],
                "مبيع شهري | Monthly Sales": r["sales_month"],
            } for r in rows])
            c1, c2 = st.columns(2)
            with c1: dl_btn(df_ns, dl_key, key=f"dlbtn_{dl_key}")
            with c2: st.warning(f"⚠️ {len(rows)} SKU بدون مبيعات | SKUs without sales")
            for r in rows:
                c_img, c_info = st.columns([1, 6])
                with c_img:
                    show_img(r["img"], 60)
                with c_info:
                    st.markdown(f"**SKU:** `{r['sku']}`", unsafe_allow_html=True)
                    st.markdown(
                        f"📦 **مخزون:** {r['stock']} &nbsp;|&nbsp; 📈 **شهري:** {r['sales_month']}",
                    )
                    sched_ns = get_latest_schedule_info(r["sku"])
                    if sched_ns:
                        arrival_ns = (sched_ns["parsed"] + timedelta(days=int(load_settings().get("schedule_delay_days","3") or 3))).date() if sched_ns.get("parsed") else None
                        st.caption(f"📅 ASN {sched_ns['asn']} بتاريخ {sched_ns['date']}" + (f" — وصول: {arrival_ns}" if arrival_ns else ""))
                    for note in get_unavailable_ordered_note(r["sku"]):
                        st.markdown(big_note_html(note), unsafe_allow_html=True)
                st.divider()

        sub1, sub2, sub3 = st.tabs([
            f"📅 بدون مبيع أمس ({len(no_sale_1d)})",
            f"📅 بدون مبيع آخر 3 أيام ({len(no_sale_3d)})",
            f"📅 بدون مبيع آخر أسبوع ({len(no_sale_7d)})",
        ])
        with sub1:
            render_no_sale_list(no_sale_1d, "أمس", "no_sale_1d")
        with sub2:
            render_no_sale_list(no_sale_3d, "آخر 3 أيام", "no_sale_3d")
        with sub3:
            render_no_sale_list(no_sale_7d, "آخر أسبوع", "no_sale_7d")
