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
live_sheet            = sheets["LIVE"]

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

# ══ قفل التابات بكلمة سر | Tab password lock ══
# ملحوظة: تخزين كلمة السر هنا نص عادي في جوجل شيت (زي باقي الإعدادات) — مش تشفير حقيقي،
# الهدف بس منع الدخول العرضي مش حماية أمنية قوية | Note: the password is stored as
# plain text in the Settings sheet like every other setting — this is a basic access
# gate, not strong security.
TAB_LOCK_OPTIONS = [
    ("tab14",     "🛒 المبيعات | Sales"),
    ("tab_dash",  "📊 داشبورد المبيعات | Sales Dashboard"),
    ("tab1",      "📋 الطلبات | Requests"),
    ("tab2",      "✅ الموافقة | Approved"),
    ("tab3",      "❌ غير متوفر | Unavailable"),
    ("tab4",      "🛒 تم الطلب | Ordered"),
    ("tab5",      "📅 الجدولة | Scheduled"),
    ("tab_check", "☑️ تشييك | Check"),
    ("tab6",      "🚫 جدولة ملغية | Cancelled"),
    ("tab7",      "🔄 تعديل موعد | Rescheduled"),
    ("tab8",      "⚠️ تنبيهات | Alerts"),
    ("tab9",      "📊 المخزون | Inventory"),
    ("tab10",     "🔴 مراجعة المخزون | Stock Review"),
    ("tab11",     "🗂️ منتهية | Expired"),
    ("tab12",     "⚙️ الإعدادات | Settings"),
    ("tab13",     "📈 مراجعة المبيعات | Sales Review"),
    ("tab15",     "🗓️ تحليل الجدولة | Schedule Analysis"),
    ("tab16",     "📦 مخزون بدون بيع | No Sales"),
]
TAB_LOCK_LABELS = dict(TAB_LOCK_OPTIONS)

def _tab_lock_setting_key(tab_key):
    return f"tab_lock_{tab_key}"

def get_tab_lock_map():
    """يرجع {tab_key: password} لكل تاب عليه كلمة سر محفوظة (مش فاضية) | Returns
    {tab_key: password} for every tab that currently has a non-empty password."""
    s = load_settings()
    out = {}
    for tab_key, _label in TAB_LOCK_OPTIONS:
        pw = s.get(_tab_lock_setting_key(tab_key), "")
        if str(pw).strip():
            out[tab_key] = pw
    return out

def set_tab_lock(tab_key, password):
    """يحفظ/يعدّل كلمة سر التاب. لو password فاضية، ده معناه إلغاء القفل | Saves or
    updates the tab's password. An empty password removes the lock."""
    save_setting(_tab_lock_setting_key(tab_key), password)

def _tab_gate(tab_key, tab_label):
    """لو التاب ده عليه كلمة سر ومحدش دخلها صح لسه في الجلسة دي، بيعرض فورم كلمة
    السر ويرجّع False (يعني منوع من عرض باقي محتوى التاب). لو مفيش كلمة سر عليه،
    أو المستخدم دخلها صح قبل كده في نفس الجلسة، بيرجّع True | If this tab has a
    password and it hasn't been entered correctly yet this session, shows a
    password form and returns False (blocking the rest of the tab's content).
    Returns True when the tab is unlocked (no password set, or already entered
    correctly this session)."""
    locks = get_tab_lock_map()
    pw = locks.get(tab_key, "")
    if not pw:
        return True
    unlocked_flag = f"_tab_unlocked_{tab_key}"
    if st.session_state.get(unlocked_flag):
        return True
    st.subheader(f"🔒 {tab_label}")
    st.info("🔐 هذا القسم محمي بكلمة سر — اكتبها للدخول | This section is password-protected — enter the password to continue")
    entered_pw = st.text_input("🔑 كلمة السر | Password", type="password", key=f"_tab_pw_input_{tab_key}")
    if st.button("دخول | Unlock", key=f"_tab_pw_btn_{tab_key}"):
        if entered_pw == pw:
            st.session_state[unlocked_flag] = True
            st.rerun()
        else:
            st.error("❌ كلمة السر غلط | Wrong password")
    return False

def get_excluded_warehouses():
    val = load_settings().get("excluded_warehouses","")
    if not val.strip():
        return set()
    return {w.strip().upper() for w in val.split(",") if w.strip()}

# ══ links map ══
def get_links_map():
    data = get_cached(links_ws)
    m = {}
    for row in data[1:]:
        if len(row) >= 2 and row[0].strip():
            m[row[0].strip().upper()] = row[1].strip()
    return m

# ══ tacweed map (SKU -> الكود 01) ══
def get_tacweed_map():
    data = get_cached(tacweed_sheet)
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

def get_ads_map():
    data = get_cached(ads_sheet)
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
def get_com_map():
    data = get_cached(com_sheet)
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

# ══ خريطة LIVE (SKU -> سعر أساسي/سعر عرض/مخزون xdock من ملف Noon catalog export) ══
def _f2_or_none(v):
    """زي _f2 بس بترجع None لو الخلية فاضية/مش رقم بدل ما ترجع صفر — عشان صفر حقيقي (سعر=0)
    ميتلخبطش مع "مفيش سعر مسجل"."""
    s = str(v).strip().replace(",", "")
    if s == "" or s.lower() in ("nan", "none"):
        return None
    try:
        return float(s)
    except Exception:
        return None

def get_live_map():
    """يرجع {SKU (upper) -> {price, sale_price, stock_xdock_net, noon_title, noon_status}}
    من تاب LIVE (نسخة ملف Noon catalog export). المفتاح هو sku_child.
    ملحوظة: مفتاح "price" (اللي بتعتمد عليه كل حسابات المبيعات والداشبورد) بقى مصدره
    عمود sale_price في تاب LIVE بدل عمود price | Note: the "price" key (which all
    Sales/Dashboard calculations rely on) is now sourced from the sale_price column
    in the LIVE sheet instead of the price column."""
    data = get_cached(live_sheet)
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

def get_fulfillment_model_map():
    """يرجع {SKU (upper) -> fulfillment_model} من تاب DailyOrders (عمود fulfillment_model لو موجود).
    بياخد أول قيمة غير فاضية يلاقيها لكل SKU. المفتاح ده بيتستخدم لاستبعاد أي SKU من نوع
    Fulfilled by Partner (FBP) من تابي مراجعة المخزون / مراجعة المبيعات (بما فيها قسم
    'مخزون منتهي بالكامل') لأنه مش مسؤولية المخزون عندنا."""
    data = get_cached(daily_orders_sheet)
    m = {}
    if len(data) <= 1:
        return m
    header = [h.strip() for h in data[0]]
    fm_idx = None
    for ci, h in enumerate(header):
        h_norm = h.strip().lower().replace(" ", "_")
        if h_norm == "fulfillment_model":
            fm_idx = ci
            break
    if fm_idx is None:
        return m
    for row in data[1:]:
        if not row:
            continue
        sku = row[0].strip() if len(row) > 0 else ""
        if not sku:
            continue
        sku_up = sku.upper()
        if m.get(sku_up):
            continue
        val = row[fm_idx].strip() if len(row) > fm_idx else ""
        if val:
            m[sku_up] = val
    return m

def is_fbp_sku(sku_up, fulfillment_map):
    """يتحقق هل الـ SKU ده Fulfilled by Partner (FBP) حسب خريطة fulfillment_model
    (من get_fulfillment_model_map) — لو كذلك، يُستبعد من تابي مراجعة المخزون / مراجعة المبيعات."""
    if not fulfillment_map:
        return False
    val = str(fulfillment_map.get(sku_up, "")).strip().upper()
    return "FBP" in val or "FULFILLED BY PARTNER" in val

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
def get_warehouse_stock_map():
    data = get_cached(warehouse_stock_sheet)
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

def build_daily_orders_family_stats(dates, live_map=None):
    """يرجع dict: اسم القسم -> {"orders": عدد, "revenue": إيراد} لفترة تواريخ محددة.
    لو عمود Family مش موجود في الشيت، أو الصف مالوش قيمة Family، بيتجاهله بهدوء
    بدون ما يوقف الكود أو يأثر على أي تحليل تاني. الإيراد بيتحسب دايمًا على سعر البيع
    الأساسي من تاب LIVE (sale_price) لو متوفر لهذا الـ SKU، وبيرجع لسعر العرض المسجل
    مع الطلب نفسه بس لو مفيش سعر بيع أساسي مسجل خالص."""
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
    fulfillment_map_css = get_fulfillment_model_map()
    rows = []
    for sku_up, qty in daily_qty.items():
        if is_sku_only_in_excluded_warehouses(sku_up, excluded_wh):
            continue
        if is_fbp_sku(sku_up, fulfillment_map_css):
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
    —  مخزونها انتهى بالكامل وخرجت من ملف المخزون. تظهر بنفس تفاصيل تابي المراجعة."""
    multi_counts = build_daily_orders_counts(display_dates)
    links_map_local = get_links_map()
    fulfillment_map_cmir = get_fulfillment_model_map()
    rows = []
    for sku_up, day_counts in multi_counts.items():
        if sku_up in inv_map:
            continue
        if is_fbp_sku(sku_up, fulfillment_map_cmir):
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

def get_recent_schedule_rows(days_back=4, include_expired=True):
    """يرجع dict: sku_upper -> أحدث سجل جدولة (من Scheduled/Check، وكمان Expired لو
    include_expired=True) وقع تاريخ جدولته خلال آخر days_back يوم (يعني: أمس/أول
    أمس/أول أول أمس/قبل 4 أيام).
    الهدف: الـ SKU ده لسه في فترة انتظار وصول المخزون بعد الجدولة — حتى لو الجدولة
    خلاص اتنقلت لتاب Expired — فلازم يفضل ظاهر في تابات المراجعة عشان محدش يطلبه تاني بالغلط.
    ملحوظة مهمة: include_expired=False بيستخدم في قرار "استبعاد الصنف من مراجعة
    المخزون/المبيعات" — لأن جدولة Expired معناها إنها فعلاً ماوصلتش ولسه محتاجة
    جدولة تانية، فمينفعش تستخدم كسبب لإخفاء الصنف من المراجعة. الاستخدام بـ
    include_expired=True (الافتراضي) لسكشن العرض الإعلامي 'مجدولة خلال آخر 4 أيام'
    اللي غرضه إظهار كل حاجة اتحركت مؤخرًا (حتى المنتهية) للمعلومية بس."""
    cutoff = datetime.now().date() - timedelta(days=days_back)
    today_ = datetime.now().date()
    src_label_map = {"Scheduled": "الجدولة | Scheduled", "Check": "تشييك | Check", "Expired": "منتهية | Expired"}
    sheet_keys = ("Scheduled", "Check", "Expired") if include_expired else ("Scheduled", "Check")
    by_sku = {}
    for sheet_key in sheet_keys:
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

def is_sku_unavailable(sku_up):
    """True لو الـ SKU موجود حالياً في تاب Unavailable (يعني متسجل غير متوفر ولسه محدش وافق
    على طلب جديد له). | True if the SKU currently sits in the Unavailable sheet."""
    data_un = get_cached(unavailable_sheet)
    if len(data_un) > 1:
        for row in data_un[1:]:
            if row and row[0].strip().upper() == sku_up:
                return True
    return False

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
    "📋 الطلبات | Requests",
    "✅ الموافقة | Approved",
    "❌ غير متوفر | Unavailable",
    "🛒 تم الطلب | Ordered",
    "📅 الجدولة | Scheduled",
    "☑️ تشييك | Check",
    "🚫 جدولة ملغية | Cancelled",
    "🔄 تعديل موعد | Rescheduled",
    "⚠️ تنبيهات | Alerts",
    "🔴 مراجعة المخزون | Stock Review",
    "🗂️ منتهية | Expired",
    "⚙️ الإعدادات | Settings",
    "📈 مراجعة المبيعات | Sales Review",
    "🗓️ تحليل الجدولة | Schedule Analysis",
])
(tab1,tab2,tab3,tab4,tab5,tab_check,tab6,tab7,tab8,tab10,tab11,tab12,tab13,tab15) = tabs

# ── تحسينات بسيطة لشكل شريط التابات الأصلي (مسافات، حواف مدوّرة، خط أوضح للتاب النشط) ──
st.markdown("""
<style>
div[data-baseweb="tab-list"] {
    gap: 4px;
    flex-wrap: wrap;
    border-bottom: 1px solid rgba(150,150,150,0.25);
}
div[data-baseweb="tab-list"] button[data-baseweb="tab"] {
    border-radius: 8px 8px 0 0;
    padding: 8px 14px;
    font-size: 14px;
}
div[data-baseweb="tab-list"] button[data-baseweb="tab"][aria-selected="true"] {
    font-weight: 700;
    border-bottom: 3px solid #ef4444;
}
</style>
""", unsafe_allow_html=True)

# ══ TAB 1 — الطلبات ══
# ══ TAB 1 — الطلبات ══
with tab1:
    if _tab_gate("tab1", "📋 الطلبات | Requests"):
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
    # ══ TAB 2 — الموافقة ══
with tab2:
    if _tab_gate("tab2", "✅ الموافقة | Approved"):
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
    # ══ TAB 3 — غير متوفر ══
with tab3:
    if _tab_gate("tab3", "❌ غير متوفر | Unavailable"):
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
    # ══ TAB 4 — تم الطلب ══
with tab4:
    if _tab_gate("tab4", "🛒 تم الطلب | Ordered"):
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
    # ══ TAB 5 — الجدولة ══
with tab5:
    if _tab_gate("tab5", "📅 الجدولة | Scheduled"):
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
    # ══ TAB CHECK — تشييك ══
with tab_check:
    if _tab_gate("tab_check", "☑️ تشييك | Check"):
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
    # ══ TAB 6 — جدولة ملغية ══
with tab6:
    if _tab_gate("tab6", "🚫 جدولة ملغية | Cancelled"):
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
    # ══ TAB 7 — تعديل الموعد ══
with tab7:
    if _tab_gate("tab7", "🔄 تعديل موعد | Rescheduled"):
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
    # ══ TAB 8 — تنبيهات ══
with tab8:
    if _tab_gate("tab8", "⚠️ تنبيهات | Alerts"):
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
    # ══ TAB 10 — مراجعة المخزون ══
with tab10:
    if _tab_gate("tab10", "🔴 مراجعة المخزون | Stock Review"):
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
        # SKUs ليها جدولة نشطة (Scheduled/Check) خلال آخر 4 أيام ولسه في فترة الوصول —
        # دي أموره تمام بالفعل، فمينفعش تظهر في مراجعة المخزون خالص. ملحوظة: هنا
        # include_expired=False عمدًا — جدولة اتنقلت لـ Expired معناها إنها فعلاً
        # ماوصلتش، فمينفعش تُستخدم كسبب لإخفاء صنف لسه محتاج جدولة تانية فعليًا
        # (كان بيحصل قبل كده تناقض: البادچ بيقول "محتاج جدولة الآن" والصنف مستبعد
        # من القايمة في نفس الوقت بسبب جدولة قديمة منتهية). صنف زي ده لسه هيظهر هنا
        # مع ملاحظة "كانت مجدولة لكن انتهت" (render_recent_expired_note تحت).
        recent_sched_map_t10_all = get_recent_schedule_rows(days_back=4, include_expired=False)
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
        fulfillment_map_t10_tr = get_fulfillment_model_map()
        for tr in transferred_from_sales:
            if is_sku_only_in_excluded_warehouses(tr["sku_up"], excluded_wh):
                continue
            if is_fbp_sku(tr["sku_up"], fulfillment_map_t10_tr):
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
    # ══ TAB 11 — منتهية الصلاحية ══
with tab11:
    if _tab_gate("tab11", "🗂️ منتهية | Expired"):
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
    # ══ TAB 12 — الإعدادات ══
with tab12:
    if _tab_gate("tab12", "⚙️ الإعدادات | Settings"):
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
        st.markdown("### 🔒 حماية التابات بكلمة سر | Tab Password Lock")
        st.caption(
            "اختر تاب وحط له كلمة سر — أي حد يفتح التاب ده هيلاقي فورم كلمة سر ومش هيقدر "
            "يشوف محتواه غير لو دخلها صح. سيب كلمة السر فاضية عشان تلغي القفل | Pick a tab "
            "and set a password for it — anyone who opens that tab will see a password form "
            "and won't see its content unless they enter it correctly. Leave the password "
            "empty to remove the lock.")
        st.caption(
            "⚠️ لو حطيت قفل على تاب \"⚙️ الإعدادات\" نفسه ونسيت كلمة السر، الطريقة الوحيدة "
            "لإلغاء القفل هي حذف السطر tab_lock_tab12 يدويًا من تاب Settings في الشيت | If you "
            "lock the Settings tab itself and forget the password, the only way back in is "
            "to manually delete the tab_lock_tab12 row from the Settings sheet.")

        current_locks_map = get_tab_lock_map()
        lock_labels_disp = [
            f"{('🔒' if k in current_locks_map else '🔓')} {lbl}" for k, lbl in TAB_LOCK_OPTIONS
        ]
        picked_idx = st.selectbox(
            "اختر التاب | Choose tab", options=range(len(TAB_LOCK_OPTIONS)),
            format_func=lambda i: lock_labels_disp[i], key="tab_lock_picker")
        picked_key, picked_label = TAB_LOCK_OPTIONS[picked_idx]
        is_locked_now = picked_key in current_locks_map

        lc1, lc2 = st.columns([2, 1])
        with lc1:
            new_pw_val = st.text_input(
                f"كلمة سر {picked_label} | Password for {picked_label}"
                + (" (محدد حاليًا | currently set)" if is_locked_now else " (لا يوجد قفل حاليًا | not locked now)"),
                type="password", key=f"tab_lock_pw_field_{picked_key}",
                placeholder="اكتب كلمة سر جديدة أو سيبها فاضية لإلغاء القفل | Type a new password, or leave empty to unlock")
        with lc2:
            st.write("")
            st.write("")
            if st.button("💾 حفظ | Save", key=f"tab_lock_save_{picked_key}"):
                set_tab_lock(picked_key, new_pw_val.strip())
                # لو المستخدم غيّر/شال كلمة السر، نلغي فتح الجلسة الحالية للتاب ده عشان
                # القفل/الفتح الجديد يتفعّل فورًا | Reset this session's unlocked flag so
                # the new lock/unlock state takes effect immediately.
                st.session_state.pop(f"_tab_unlocked_{picked_key}", None)
                if new_pw_val.strip():
                    st.success(f"✅ تم قفل {picked_label} | {picked_label} is now locked")
                else:
                    st.success(f"✅ تم إلغاء قفل {picked_label} | {picked_label} is now unlocked")
                st.rerun()

        if current_locks_map:
            st.caption("🔒 التابات المقفولة حاليًا | Currently locked tabs: " +
                       "، ".join(TAB_LOCK_LABELS.get(k, k) for k in current_locks_map))
        else:
            st.caption("🔓 لا يوجد أي تاب مقفول حاليًا | No tabs are currently locked")
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

        st.divider()
        st.markdown("### 🟣 حد تنبيه مخزون Xdock | Xdock Low-Stock Alert Threshold")
        st.caption("لو مخزون Xdock (عمود stock_xdock_net في تاب LIVE) وصل للرقم ده أو أقل، يظهر في التنبيهات السريعة بالداشبورد وتاب المبيعات كمحتاج تزويد | If Xdock stock (stock_xdock_net in the LIVE sheet) reaches this number or less, it shows up in the Dashboard/Sales quick alerts as needing restock")
        current_xdock_th = int(current_settings.get("xdock_low_stock_threshold","10") or 10)
        new_xdock_th = st.number_input("الحد | Threshold", min_value=0, max_value=1000, value=current_xdock_th, step=1, key="xdock_th_input")
        if st.button("💾 حفظ الحد | Save Threshold", key="save_xdock_th"):
            save_setting("xdock_low_stock_threshold", str(new_xdock_th))
            st.success("✅ تم الحفظ | Saved")
            st.rerun()

    # ══ TAB 13 — مراجعة المبيعات ══
    # ══ TAB 13 — مراجعة المبيعات ══
with tab13:
    if _tab_gate("tab13", "📈 مراجعة المبيعات | Sales Review"):
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
        # نفس المنطق المصحّح في مراجعة المخزون: بنستبعد بس على أساس جدولة نشطة
        # (Scheduled/Check) خلال آخر 4 أيام — مش جدولة Expired، عشان صنف جدولته
        # انتهت من غير ما توصل يفضل يظهر ويحتاج مراجعة فعلاً، مش يختفي بالغلط
        # (شايفينها بس في سكشن "مجدولة خلال آخر 4 أيام" تحت لو لسه نشطة)
        recent_sched_map_t13_all = get_recent_schedule_rows(days_back=4, include_expired=False)
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

        pending_approval_skus_t13 = get_pending_approval_skus()

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
    # ══ TAB 15 — تحليل الجدولة ══
with tab15:
    if _tab_gate("tab15", "🗓️ تحليل الجدولة | Schedule Analysis"):
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
