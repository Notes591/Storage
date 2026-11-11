# -*- coding: utf-8 -*-
import os
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import time
import pandas as pd
import io
import re
import gspread.exceptions
import calc
import calc
from calc import (
    gap_cover, gap_need_qty, gap_action, parse_excel_date,
    _build_tacweed_bundle, _build_tacweed_amazon_bundle, _f2, _f2_or_none,
    get_base_price, compute_net_price_after_fees, get_latest_sku_price, _build_warehouse_stock_map,
    _china_norm_code, _china_eta_stats, _china_int, _china_date,
    _china_merge, _china_rows_from_raw, _cs_to_ship, _cs_to_cont,
    _cs_to_stock, _cf_g, _cf_num, china_ship_record_key,
    _cf_pick, _china_ship_key, _china_line_key, _china_cont_key,
    _china_avail, china_supply_projection, china_decide, _china_days,
    _to_int, parse_count_dates, append_count_date, _find_fulfillment_col_idx,
    _find_status_col_idx, _row_is_fbb, _find_col_idx_amz, _amz_col_indexes,
    _amz_row_is_fba, _amz_row_matches, _ty_find_col, _ty_num,
    _build_ty_products, build_warehouse_stockout_sku_lookup, build_warehouse_stockout_asin_lookup, _avg_price_from,
    _gap_qty, _gap_sug, parse_check_tokens, _ago_txt,
    _ct_parse_date, _build_trendyol_bundle, _build_ty_ads, _china_alert,
    _china_auto_formulas, _china_cell, _china_clean, _china_clean_sheet_row,
    _china_col_letter, _china_enrich, _china_fmt_date, _china_group,
    _china_header_row, _china_image_formula, _china_input_segments, _china_is_label_row,
    _china_layout_outdated, _china_next_action_row, _china_parse_sheet, _china_read_flat,
    _china_rgb, _cs_container, _cs_floor, _cs_prog,
    _cs_shipment, _ct_add_months, _tab_lock_setting_key, build_cross_alerts,
    build_daily_orders_family_stats_amazon_all, build_daily_orders_family_stats_amazon_fba, build_daily_orders_family_stats_amazon_normal, build_daily_orders_family_stats_trendyol,
    china_batches_for_lines, china_flat_prepare, china_flat_to_tables, china_next_action_formula,
    china_row_status, china_status_code, china_status_label, china_tables_to_flat,
    contract_events, family_display_name, now_str, CHINA_CONT_COLS,
    CHINA_CONT_STATUS_IDX, CHINA_DATE_FIELDS, CHINA_FLAT_COLS, CHINA_FLAT_FORMULA_KEYS,
    CHINA_FLAT_INPUT_KEYS, CHINA_FLAT_KEYS, CHINA_LINE_COLS, CHINA_RECEIVED_IDX,
    CHINA_SHIPPED_IDX, CHINA_SHIP_COLS, CHINA_SKIP_COLS, CHINA_STAGES,
    CHINA_STAGE_AR, CHINA_STATUS, CHINA_STATUS_LEGACY, CONTRACT_COLS,
    CT_ORD_AR, CT_WINDOW_DAYS, FAMILY_HOME_GROUP, FAMILY_LABELS,
    OLD_SCHED_COLS, TABS_CONFIG, TY_ADS_HEADERS, TY_EXCLUDED_STATUSES,
    CT_SHEET_NAME, ct_read_tracking, ct_sync_plan,
    china_incoming_badge_text, skus_asins_links_html,
    _AR_MONTHS, _CF_CONT_ONLY, _CF_CONT_RENAME, _CF_LINE,
    _CF_SHARED, _CF_SHIP_DATA, _CF_SHIP_ONLY, _CHINA_LABELS,
    _CS_ALIAS, _CS_LABEL, _CS_LOOKUP, _CS_RANK,
    _XS_NAMES, _ads_file_rows, _file_to_text, _ty_ads_parse_xlsx,
    _ty_parse_orders_xlsx, _ty_products_parse_xlsx, asin_link_html, build_amazon_asin_link,
    build_noon_link, channel_gap_df, extract_excel_images, file_timestamp,
    make_empty_template, sku_link_html, to_excel, _GAP_ACTION_AR,
    _GAP_KINDS, _TY_ALLOC_CACHE, _TY_CID_RE, _XL_NS,
)

class _Env:
    """بوابة calc.py للبيانات: أي اسم بيتجاب من متغيرات الملف وقت الاستدعاء (فيتحدّث مع كل تحميل)."""
    def __getattr__(self, name):
        return globals()[name]
_ENV = _Env()

class _Env:
    """بوابة calc.py للبيانات: أي اسم بيتجاب من متغيرات الملف وقت الاستدعاء (فيتحدّث مع كل تحميل)."""
    def __getattr__(self, name):
        return globals()[name]
_ENV = _Env()

st.set_page_config(page_title="🏢 عالم الرشاقة للتجارة | Fitness World Trading", page_icon="🏢", layout="wide")

# ══ منع إيماءة "اسحب للرجوع" (swipe-back) من متصفح الموبايل | Prevent the
#    mobile browser's own edge swipe-back/forward gesture ══
# لمس/سحب قريب من حافة الشاشة (زي حافة القايمة الجانبية) بيخلي المتصفح يفسّرها
# كإيماءة "ارجع للصفحة اللي قبلها"، فبيعرض لمحة سريعة من صفحة تانية (شكلها شبه
# صفحتنا لأنها نفس الصفحة بحالة قبل كده) قبل م يرجع — وده اللي كان حاسس إنه
# "قفزة" بتغطي فورم الدخول. الإيماءة دي مالهاش علاقة بالسكريبت بتاعنا، هي جزء
# من المتصفح نفسه | A tap/drag near the screen edge (like the sidebar's edge)
# gets read by the mobile browser as a "go back" swipe gesture, which briefly
# previews another page (looking like ours since it's the same page in an
# earlier state) before snapping back — that's the "jump" that appeared to
# cover the login form. This gesture is a browser-level feature, unrelated to
# our own script.
st.markdown("""
<style>
html, body {
    overscroll-behavior-x: none;
    touch-action: pan-y;
}
</style>
""", unsafe_allow_html=True)

# ══ اتصال ══
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
if "_gs_client" not in st.session_state:
    _creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
    st.session_state["_gs_client"] = gspread.authorize(_creds)
client = st.session_state["_gs_client"]

def open_spreadsheet(retries=5, delay=2):
    for attempt in range(retries):
        try:
            return client.open("Complaints")
        except gspread.exceptions.APIError as e:
            if attempt < retries - 1:
                time.sleep(delay * (2 ** attempt))
            else:
                raise e

if "_gs_ss" not in st.session_state:
    st.session_state["_gs_ss"] = open_spreadsheet()
ss = st.session_state["_gs_ss"]

# ══ الأوراق ══


def get_or_create_worksheet(tab, headers, retries=5, delay=2):
    for attempt in range(retries):
        try:
            ws = ss.worksheet(tab)
            # sync header: أضيف الأعمدة الجديدة لو ناقصة فعلاً — المقارنة بتتجاهل حالة
            # الحروف والمسافات الزيادة، عشان عمود موجود فعلاً (زي "Cost") متتضافش نسخة
            # منه تاني في كل جلسة جديدة بسبب فرق بسيط في التنسيق | Header-sync check
            # ignores case/extra whitespace, so a column that already exists (e.g. "Cost")
            # doesn't get a duplicate appended every new session just from a formatting
            # mismatch.
            try:
                existing_hdr = ws.row_values(1)
                existing_hdr_norm = {str(h).strip().lower() for h in existing_hdr}
                missing = [h for h in headers if h.strip().lower() not in existing_hdr_norm]
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
def _init_all_sheets():
    """أول جلسة بس: طلب واحد لأسماء كل التابات + طلب واحد (values_batch_get) لقراءة كل بياناتها وتعبئة الكاش،
    بدل ما كان بيتعمل طلبين (ss.worksheet + row_values) لكل تاب + قراءة منفصلة لكل تاب. الهيدر بيتراجع من أول صف
    في نفس القراءة. أي تاب مش موجود بيتعمل بالطريقة القديمة | One metadata call + one batch read for ALL tabs."""
    try:
        existing = {w.title: w for w in ss.worksheets()}
    except Exception:
        existing = {}
    st.session_state["_ws_existing"] = existing
    wanted = [t for t in list(TABS_CONFIG.keys()) + ["links n", "links a"] if t in existing]
    vals = {}
    try:
        res = ss.values_batch_get(["'" + t.replace("'", "''") + "'" for t in wanted])
        for t, vr in zip(wanted, res.get("valueRanges", [])):
            rows = vr.get("values", []) or []
            w = max((len(r) for r in rows), default=0)
            vals[t] = [list(r) + [""] * (w - len(r)) for r in rows]
    except Exception:
        vals = {}
    out = {}
    for tab, headers in TABS_CONFIG.items():
        if tab in existing and tab in vals:
            ws = existing[tab]
            hdr = vals[tab][0] if vals[tab] else []
            norm = {str(h).strip().lower() for h in hdr}
            missing = [h for h in headers if h.strip().lower() not in norm]
            if missing:
                try:
                    for h in missing:
                        ws.append_cols([[h]], value_input_option="RAW")
                except Exception:
                    pass
            else:
                st.session_state[f"cache_{tab}"] = vals[tab]
            out[tab] = ws
        else:
            out[tab] = get_or_create_worksheet(tab, headers)
    for t in ("links n", "links a"):
        if t in vals:
            st.session_state[f"cache_{t}"] = vals[t]
    return out

if "sheets_initialized" not in st.session_state:
    st.session_state["sheets_initialized"] = _init_all_sheets()
sheets = st.session_state["sheets_initialized"]

def get_or_create_links_ws(retries=5, delay=2):
    for attempt in range(retries):
        try:
            return st.session_state.get("_ws_existing", {}).get("links n") or ss.worksheet("links n")
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

if "_links_ws" not in st.session_state:
    st.session_state["_links_ws"] = get_or_create_links_ws()
links_ws = st.session_state["_links_ws"]

def get_or_create_links_asin_ws(retries=5, delay=2):
    """شيت روابط صور أمازون — مفتاح بالـ ASIN (مش SKU زي 'links n' فوق)، عشان
    ده المعرّف الحقيقي لمنتج أمازون وصورته واحدة بغض النظر عن الـ MSKU |
    Amazon image-links sheet keyed by ASIN (not SKU like 'links n' above),
    since ASIN is Amazon's real product identity and its photo is the same
    regardless of which MSKU sold it."""
    for attempt in range(retries):
        try:
            return st.session_state.get("_ws_existing", {}).get("links a") or ss.worksheet("links a")
        except gspread.exceptions.WorksheetNotFound:
            try:
                ws = ss.add_worksheet(title="links a", rows="2000", cols="2")
                ws.append_row(["Asin","Image"])
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

if "_links_asin_ws" not in st.session_state:
    st.session_state["_links_asin_ws"] = get_or_create_links_asin_ws()
links_asin_ws = st.session_state["_links_asin_ws"]

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

def get_scheduled_normalized(*args, **kwargs):
    return calc.get_scheduled_normalized(_ENV, *args, **kwargs)

def build_scheduled_row(*args, **kwargs):
    return calc.build_scheduled_row(_ENV, *args, **kwargs)
cancelled_sheet   = sheets["CancelledSchedule"]
reschedule_sheet  = sheets["Rescheduled"]
expired_sheet     = sheets["Expired"]
def _ws_once(tab, headers):
    """نفس get_or_create_worksheet بس مرة واحدة في الجلسة (بدل ما تتنادى مع كل تفاعل وتعمل طلبات جوجل زيادة)."""
    c = st.session_state.setdefault("_ws_once_cache", {})
    if tab not in c:
        c[tab] = get_or_create_worksheet(tab, headers)
    return c[tab]


# نسخة التجربة: secret اسمه INVENTORY_TAB (مثلاً Inventory_API) يخلي المخزون يتقرا منه بدل Inventory.
_INV_TAB = str(os.environ.get("INVENTORY_TAB") or st.secrets.get("INVENTORY_TAB", "") or "Inventory_API").strip()
if _INV_TAB and _INV_TAB != "Inventory":
    inventory_sheet = _ws_once(_INV_TAB, TABS_CONFIG["Inventory"])
else:
    _INV_TAB = ""
    inventory_sheet = sheets["Inventory"]
# نسخة الـ API: الطلبات من تاب DailyOrders_API والمخزون من Inventory_API (تقدر تغيّرهم بـ secret ORDERS_TAB / INVENTORY_TAB).
_ORDERS_TAB = str(os.environ.get("ORDERS_TAB") or st.secrets.get("ORDERS_TAB", "") or "DailyOrders_API").strip()
if _ORDERS_TAB and _ORDERS_TAB != "DailyOrders":
    daily_orders_sheet = _ws_once(_ORDERS_TAB, TABS_CONFIG["DailyOrders"])
else:
    _ORDERS_TAB = ""
    daily_orders_sheet = sheets["DailyOrders"]
daily_orders_amazon_sheet = sheets["DailyOrdersAmazon"]
stock_amazon_sheet   = sheets["StockAmazon"]
daily_orders_trendyol_sheet = sheets["DailyOrdersTrendyol"]
ads_ty_sheet = sheets["AdvertisementsTrendyol"]
ty_products_sheet = sheets["TrendyolProducts"]
china_img_sheet = sheets["ChinaImages"]
contracts_sheet = sheets["Contracts"]
settings_sheet       = sheets["Settings"]
cancel_notif_sheet   = sheets["CancelNotifications"]
tacweed_sheet        = sheets["Tacweed"]
tacweed_amazon_sheet = sheets["TacweedAmazon"]
warehouse_stock_sheet = sheets["WarehouseStock"]
ads_sheet             = sheets["Advertisements"]
ads_amz_sheet         = sheets["AdvertisementsAmz"]
# العمولة والتوصيل: الاعتماد كله على تاب COM_API (من المعاملات المالية للـ API) — من غير رجوع لتاب COM القديم.
_COM_TAB = str(os.environ.get("COM_TAB") or st.secrets.get("COM_TAB", "") or "COM_API").strip()
com_sheet = _ws_once(_COM_TAB, TABS_CONFIG["COM"])
# الكتالوج: LIVE_TAB (env/secret) — لو التاب موجود وفيه بيانات بيتقرا منه، وإلا بيرجع لتاب LIVE العادي.
_LIVE_TAB = str(os.environ.get("LIVE_TAB") or st.secrets.get("LIVE_TAB", "") or "LIVE_API").strip()
live_sheet = sheets["LIVE"]
if _LIVE_TAB and _LIVE_TAB != "LIVE":
    try:
        _lw = st.session_state.get("_live_api_ws")        # متخزّن في الجلسة: من غير قراءة العمود كله مع كل تفاعل
        if _lw is None:
            _lw = get_or_create_worksheet(_LIVE_TAB, TABS_CONFIG["LIVE"])
            if str(_lw.acell("A2").value or "").strip():   # خلية واحدة بدل col_values (كانت بتقرا العمود كله)
                st.session_state["_live_api_ws"] = _lw
            else:
                _lw = None
        if _lw is not None:
            live_sheet = _lw
        else:
            _LIVE_TAB = ""
    except Exception:
        _LIVE_TAB = ""
else:
    _LIVE_TAB = ""

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

# ══ كاش لأي "خريطة مشتقة" من شيت معيّن (زي SKU -> كود، أو كود -> كمية) — بيتم إعادة
#    بنائها بس أول مرة بعد ما بيانات الشيت الخام تتغيّر (بعد clear_cache)، مش في كل مرة
#    الدالة بتتنادى. من غير الكاش ده، أي دالة زي دي بتتنادى جوه loop بيلف على كل الـ SKUs
#    المعروضة (زي تاب المبيعات) كانت بتعيد بناء نفس الخريطة من الصفر مرة لكل SKU، وده
#    اللي بيسبب بطء ملحوظ لما يزيد عدد الـ SKUs | Cache for any "derived map" built from a
#    sheet's raw rows. It's only rebuilt the first time after the underlying raw data
#    changes (after clear_cache), not on every call. Without this, a map-building function
#    called inside a per-SKU loop (like the Sales tab) would re-scan the whole sheet from
#    scratch for every single SKU rendered — the actual cause of the slowdown.
def _memo_from_sheet(sheet, cache_key, builder):
    raw = get_cached(sheet)
    src_key = f"{cache_key}__src"
    if src_key in st.session_state and st.session_state[src_key] is raw:
        return st.session_state[cache_key]
    result = builder(raw)
    st.session_state[cache_key] = result
    st.session_state[src_key] = raw
    return result

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
    ("tab_wh",    "🏭 مخزون المستودع | Warehouse Stockout"),
    ("tab_china", "🚢 شحنات الصين | China Shipments"),
    ("tab_contracts", "📑 العقود | Contracts"),
    ("tab10",     "🔴 مراجعة المخزون | Stock Review"),
    ("tab11",     "🗂️ منتهية | Expired"),
    ("tab12",     "⚙️ الإعدادات | Settings"),
    ("tab13",     "📈 مراجعة المبيعات | Sales Review"),
    ("tab15",     "🗓️ تحليل الجدولة | Schedule Analysis"),
    ("tab16",     "📦 مخزون بدون بيع | No Sales"),
    ("tab_upload", "📤 رفع ملف | Upload Files"),
]
TAB_LOCK_LABELS = dict(TAB_LOCK_OPTIONS)


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
    أو المستخدم دخلها صح قبل كده (لأي تاب) في نفس الجلسة، بيرجّع True | If this tab
    has a password and it hasn't been entered correctly yet this session, shows a
    password form and returns False (blocking the rest of the tab's content).
    Returns True when the tab is unlocked (no password set, or a password for
    ANY locked tab was already entered correctly this session).

    ── الباسورد بيتكتب مرة واحدة بس لكل الجلسة | Enter once, unlocks every locked
    tab: أول ما المستخدم يدخل كلمة سر صح لأي تاب محمي، بنحفظ فلاج عام
    (_all_tabs_unlocked) في الجلسة، وأي تاب تاني عليه كلمة سر بيتفتح تلقائي من
    غير ما يُطلب يدخلها تاني | The first time the user enters any locked tab's
    password correctly, we set a single session-wide flag; every other locked
    tab then opens automatically without asking again this session."""
    locks = get_tab_lock_map()
    pw = locks.get(tab_key, "")
    if not pw:
        return True
    if st.session_state.get("_all_tabs_unlocked"):
        return True
    unlocked_flag = f"_tab_unlocked_{tab_key}"
    if st.session_state.get(unlocked_flag):
        return True
    st.subheader(f"🔒 {tab_label}")
    st.info("🔐 هذا القسم محمي بكلمة سر — اكتبها للدخول (هتفتح كل الأقسام المحمية بعد كده لنفس الجلسة) | This section is password-protected — enter the password to continue (unlocks every protected section for the rest of this session)")
    entered_pw = st.text_input("🔑 كلمة السر | Password", type="password", key=f"_tab_pw_input_{tab_key}")
    if st.button("دخول | Unlock", key=f"_tab_pw_btn_{tab_key}"):
        if entered_pw == pw:
            st.session_state[unlocked_flag] = True
            st.session_state["_all_tabs_unlocked"] = True
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

def get_links_map_asin():
    """خريطة ASIN -> رابط صورة من شيت 'links a' — المصدر الأساسي لصور تابات
    مبيعات أمازون | ASIN -> image URL map from the 'links a' sheet — the
    primary image source for the Amazon Sales tabs."""
    data = get_cached(links_asin_ws)
    m = {}
    for row in data[1:]:
        if len(row) >= 2 and row[0].strip():
            m[row[0].strip().upper()] = row[1].strip()
    return m

# ══ tacweed data (SKU -> {code, cost}) — مبني على أسماء الأعمدة (مش ترتيبها) عشان يفضل
#    شغال حتى لو اتغير ترتيب الأعمدة في الشيت، ومكاش بـ _memo_from_sheet عشان متتبنيش
#    الخريطة من الصفر لكل SKU بيتعرض | Built from column names (not position), so it keeps
#    working even if the sheet's column order changes, and cached via _memo_from_sheet so
#    it isn't rebuilt from scratch for every displayed SKU ══

def _get_tacweed_bundle():
    return _memo_from_sheet(tacweed_sheet, "tacweed_bundle", _build_tacweed_bundle)

def get_tacweed_data(*args, **kwargs):
    return calc.get_tacweed_data(_ENV, *args, **kwargs)

# ══ tacweed map (SKU -> الكود 01) ══
def get_tacweed_map(*args, **kwargs):
    return calc.get_tacweed_map(_ENV, *args, **kwargs)

def get_tacweed_reverse_map(*args, **kwargs):
    return calc.get_tacweed_reverse_map(_ENV, *args, **kwargs)

# ══ تكلفة الوحدة من تاب التكويد — عن طريق الـ SKU مباشرة، أو عن طريق الكود01 المشترك ══
# ══ Unit cost from the Tacweed sheet — directly by SKU, or via the shared Code01 ══
def get_tacweed_cost_map(*args, **kwargs):
    return calc.get_tacweed_cost_map(_ENV, *args, **kwargs)

def get_tacweed_code_cost_map(*args, **kwargs):
    return calc.get_tacweed_code_cost_map(_ENV, *args, **kwargs)

def get_sku_cost(*args, **kwargs):
    return calc.get_sku_cost(_ENV, *args, **kwargs)

# ══ نفس منطق تكويد نون بالظبط (_build_tacweed_bundle وما بعدها) لكن بالمفتاح
#    ASIN بدل SKU، مبني على تاب TacweedAmazon | Same logic as the Noon Tacweed
#    bundle above, but keyed by ASIN instead of SKU, built from TacweedAmazon ══

def _get_tacweed_amazon_bundle():
    return _memo_from_sheet(tacweed_amazon_sheet, "tacweed_amazon_bundle", _build_tacweed_amazon_bundle)

def get_tacweed_amazon_data(*args, **kwargs):
    return calc.get_tacweed_amazon_data(_ENV, *args, **kwargs)

def get_tacweed_amazon_map(*args, **kwargs):
    return calc.get_tacweed_amazon_map(_ENV, *args, **kwargs)

def get_tacweed_amazon_reverse_map(*args, **kwargs):
    return calc.get_tacweed_amazon_reverse_map(_ENV, *args, **kwargs)

def get_tacweed_amazon_cost_map(*args, **kwargs):
    return calc.get_tacweed_amazon_cost_map(_ENV, *args, **kwargs)

def get_tacweed_amazon_code_cost_map(*args, **kwargs):
    return calc.get_tacweed_amazon_code_cost_map(_ENV, *args, **kwargs)

def get_asin_cost(*args, **kwargs):
    return calc.get_asin_cost(_ENV, *args, **kwargs)

def resolve_unit_cost(*args, **kwargs):
    return calc.resolve_unit_cost(_ENV, *args, **kwargs)

def big_note_html(text):
    """نص ملاحظة (غير متوفر سابقاً / تم طلبه سابقاً) بخط أكبر وأسود بولد بدل الكابشن الصغير الرمادي."""
    return f'<span class="status-badge-lg" style="background:#e5e7eb;">{text}</span>'

# ══ خريطة الإعلانات (SKU -> قائمة كامبينات) | Ads map (SKU -> list of campaigns) ══

def get_ads_map(*args, **kwargs):
    """إعلانات نون: المصدر الأساسي تاب AdsMonthly (البوت بيحدّثه كل تشغيل)، ولو مش موجود/فاضي بيرجع لشيت Advertisements اليدوي."""
    rows = _ads_monthly_rows()
    if rows and len(rows) > 1:
        return calc.get_ads_map(_ENV, data=rows)
    return calc.get_ads_map(_ENV, *args, **kwargs)

def _ads_monthly_rows():
    c = st.session_state.get("_ads_monthly_rows")
    if c and time.time() - c[0] < 600:
        return c[1]
    rows = None
    try:
        rows = ss.worksheet("AdsMonthly").get_all_values()
    except Exception:
        rows = c[1] if c else None
    st.session_state["_ads_monthly_rows"] = (time.time(), rows)
    return rows

# ══ خريطة إعلانات أمازون (ASIN -> قائمة صفوف) | Amazon ads map (ASIN -> list of rows) ══
def get_ads_map_amazon(*args, **kwargs):
    return calc.get_ads_map_amazon(_ENV, *args, **kwargs)

# ══ خريطة العمولة/التوصيل (SKU -> {delivery, commission_pct}) | Commission & delivery map ══
def get_com_map(*args, **kwargs):
    return calc.get_com_map(_ENV, *args, **kwargs)

# ══ خريطة LIVE (SKU -> سعر أساسي/سعر عرض/مخزون xdock من ملف Noon catalog export) ══

def get_live_map(*args, **kwargs):
    return calc.get_live_map(_ENV, *args, **kwargs)




def render_price_profit_block_t14(r, com_map_t14, live_map_t14, sales_dates, ads_entries_t14, total_orders_t14, total_spends_t14, part="all", reco_noon=False):
    """يعرض بلوك السعر الأساسي/التوصيل/العمولة/الصافي/الربح لصف SKU واحد في تاب المبيعات
    (بيتنادى من الـ FBN والـ FBB على حد سواء). بيعرض البلوك دايمًا حتى لو بيانات العمولة/
    التوصيل (من تاب COM) أو السعر (من LIVE/الطلبات) ناقصة لهذا الـ SKU — وبيحط صفر في
    القيمة الناقصة مع تنبيه واضح ⚠️، بدل ما البلوك يختفي بصمت زي قبل كده | Renders the
    price/delivery/commission/net/profit block for one SKU row in the Sales tab (called
    from both FBN and FBB). Always shows the block even when COM (commission/delivery) or
    price data is missing for this SKU — missing values default to zero with a visible
    warning, instead of silently hiding the whole block like before."""
    com_info_t14 = com_map_t14.get(r["sku_up"])
    com_missing_t14 = com_info_t14 is None
    if com_missing_t14:
        com_info_t14 = {"delivery": 0.0, "commission_pct": 0.0}
    # ── عمولة/توصيل حسب نوع التنفيذ: نوع التنفيذ اللي اتباع منه الـ SKU مؤخرًا (أحدث يوم بيع في فترة العرض،
    #    ولو الاتنين في نفس اليوم فاللي قطعه أكتر) — بنستخدم عمولة وتوصيل **آخر بيع مكتمل** للنوع ده من COM_API.
    #    لو النوع ملوش رقم، بنرجع للرقم العام (آخر بيع مكتمل من أي نوع) ──
    _mix = r.get("_mix") or {}
    if _mix and not com_missing_t14:
        _best = None
        for _k, _v in _mix.items():
            if _v.get("idx") is None:
                continue
            _score = (_v["idx"], -_v.get("units", 0))
            if _best is None or _score < _best[0]:
                _best = (_score, _k)
        if _best:
            _sfx = "fbn" if _best[1] == "FBN" else "fbb"
            _d = com_info_t14.get("delivery_" + _sfx)
            _c = com_info_t14.get("commission_pct_" + _sfx)
            com_info_t14 = dict(com_info_t14,
                                delivery=com_info_t14.get("delivery", 0.0) if _d is None else _d,
                                commission_pct=com_info_t14.get("commission_pct", 0.0) if _c is None else _c)

    offer_price_t14 = get_latest_sku_price(r, sales_dates)
    latest_price_t14, price_from_live_t14 = get_base_price(r["sku_up"], live_map_t14, offer_price_t14)
    price_missing_t14 = latest_price_t14 is None
    if price_missing_t14:
        latest_price_t14 = 0.0

    net_fees_t14, net_tax_t14 = compute_net_price_after_fees(latest_price_t14, com_info_t14)

    # ── تكلفة الوحدة من تاب التكويد (لو مسجلة) وصافي الربح الفعلي بعدها | Unit
    #    cost from the Tacweed sheet (if recorded) and the actual net profit
    #    after subtracting it — لو التكلفة مش مسجلة بتتحط صفر مع تنبيه، بدل ما
    #    سطر التكلفة يختفي بصمت زي قبل كده | if cost isn't recorded it defaults
    #    to zero with a warning, instead of the cost line silently disappearing ──
    cost_t14 = resolve_unit_cost(r["sku_up"], r.get("asin", "").upper() if r.get("asin") else None)
    cost_missing_t14 = cost_t14 is None
    if cost_missing_t14:
        cost_t14 = 0.0
    net_profit_t14 = net_tax_t14 - cost_t14
    # الرقم المستخدم في مقارنة أداء الإعلان: صافي الربح الفعلي بعد التكلفة | Figure
    # used for the ad ROI comparison: the actual net profit after cost.
    profit_basis_t14 = net_profit_t14

    # ── فلوس الإعلان دي مفديه ولا لأ؟ | Did the ad spend pay off overall? ──
    # مقارنة إجمالية: صافي الربح الكلي من الطلبات اللي جابها الإعلان مقابل
    # إجمالي اللي اتصرف على الإعلان — مش مقارنة لكل طلب لوحده
    # الحكم على الإعلان بيظهر كسطر قصير (مربح / غير مربح)، والتفاصيل بالأرقام جوه سلايدر | Short
    # verdict visible, the numbers live in an expander.
    ad_verdict_t14, ad_detail_t14, ad_ok_t14 = "", "", True
    if ads_entries_t14:
        if total_orders_t14 <= 0:
            ad_ok_t14 = False
            ad_verdict_t14 = "🚨 الاعلان لسه ما جابش طلبات"
            ad_detail_t14 = (f"اتصرف {total_spends_t14:,.2f} ريال على الإعلان ده ولسه ما جابش أي طلبات فعلية — "
                             "يستاهل تراجع تفاصيل الحملة فوق 👆")
        else:
            total_net_from_ads_t14 = total_orders_t14 * profit_basis_t14
            net_result_t14 = total_net_from_ads_t14 - total_spends_t14
            if net_result_t14 >= 0:
                ad_verdict_t14 = "🎯 الاعلان مربح"
                ad_detail_t14 = (f"عدد طلبات الاعلان {total_orders_t14:,.0f} طلب بصافي ربح إجمالي {total_net_from_ads_t14:,.2f} ريال "
                                 f"مقابل {total_spends_t14:,.2f} ريال مدفوع — حقق **{net_result_t14:,.2f} ريال** 👌")
            else:
                ad_ok_t14 = False
                ad_verdict_t14 = "🚨 الاعلان غير مربح"
                ad_detail_t14 = (f"مدفوع {total_spends_t14:,.2f} ريال، لكن صافي الربح من {total_orders_t14:,.0f} طلب بس "
                                 f"{total_net_from_ads_t14:,.2f} ريال — خسران **{abs(net_result_t14):,.2f} ريال إجمالي**. "
                                 "افتح تفاصيل الحملة فوق 👆 وتراجعها")

    offer_price_line_t14 = ""
    if offer_price_t14 is not None and (not price_from_live_t14 or round(offer_price_t14, 2) != round(latest_price_t14, 2)):
        offer_price_line_t14 = f' &nbsp;|&nbsp; 🏷️ سعر العرض (معلومة فقط، غير مستخدم في الحسابات): <b>{offer_price_t14:,.2f}</b> ريال'
    price_label_t14 = "سعر البيع الأساسي" if price_from_live_t14 else "سعر البيع (آخر طلب)"

    profit_color_t14 = "#4ade80" if net_profit_t14 >= 0 else "#f87171"

    # ── تنبيه واضح لو أي جزء من البيانات (سعر / عمولة وتوصيل / تكلفة المنتج) كان
    #    ناقص واتحط صفر بدلاً منه | Visible warning when price, COM data, or product
    #    cost was missing and defaulted to zero ──
    missing_parts_t14 = []
    if price_missing_t14:
        missing_parts_t14.append("سعر بيع (LIVE/الطلبات)")
    if com_missing_t14:
        missing_parts_t14.append("عمولة وتوصيل (تاب COM)")
    if cost_missing_t14:
        missing_parts_t14.append("تكلفة المنتج (تاب التكويد)")
    missing_note_t14 = ""
    if missing_parts_t14:
        missing_note_t14 = (
            f'<div style="color:#ef4444;font-size:13px;font-weight:800;margin-top:4px;">⚠️ ناقص: '
            f'{" و".join(missing_parts_t14)} — عشان كده صافي الربح مش معروض</div>'
        )

    verdict_html_t14 = ""
    if ad_verdict_t14:
        _vc = ("#052e1655", "#16a34a", "#4ade80") if ad_ok_t14 else ("#4c051655", "#dc2626", "#f87171")
        verdict_html_t14 = (f'<div dir="rtl" style="margin-top:8px;padding:6px 11px;background:{_vc[0]};border:1px solid {_vc[1]};'
                            f'border-radius:7px;"><span style="color:{_vc[2]};font-size:13px;font-weight:800;">{ad_verdict_t14}</span></div>')

    # الظاهر: سطور عمودية (سعر / توصيل / عمولة) + صافي الربح + حكم الإعلان. الباقي جوه سلايدرات.
    _row = ('<div style="display:flex;justify-content:space-between;align-items:center;gap:8px;'
            'font-size:13px;color:#e2e8f0;padding:3px 0;"><span>{k}</span><b dir="ltr">{v}</b></div>')
    if part in ("all", "box"):
      # نسخة مضغوطة: سعر | توصيل | عمولة في سطر واحد (كبسولات صغيرة) + صافي الربح وحكم الإعلان في سطر تاني
      _chip_b = ('<span style="background:#0f172a;border:1px solid #334155;border-radius:6px;padding:1px 8px;'
                 'font-size:12px;color:#e2e8f0;white-space:nowrap;">{k} <b dir="ltr">{v}</b></span>')
      _pill_v = ""
      if ad_verdict_t14:
          _vc2 = ("#16a34a", "#4ade80") if ad_ok_t14 else ("#dc2626", "#f87171")
          _pill_v = (f'<span style="border:1px solid {_vc2[0]};color:{_vc2[1]};border-radius:999px;padding:1px 9px;'
                     f'font-size:11px;font-weight:800;white-space:nowrap;">{ad_verdict_t14}</span>')
      _reco_pill, _reco_line = ("", "")
      if reco_noon and ad_verdict_t14 and r.get("sku_up"):
          _reco_pill, _reco_line = ads_reco_inline_html(r["sku_up"])
      st.markdown(
        '<div style="background:#1e293b;border:1px solid #334155;border-radius:8px;padding:5px 10px;margin:4px 0;">'
        '<div style="display:flex;flex-wrap:wrap;gap:5px;align-items:center;">'
        + _chip_b.format(k=f"💵 {price_label_t14}", v=f"{latest_price_t14:,.2f}")
        + _chip_b.format(k="🚚", v=f'{com_info_t14["delivery"]:,.0f}')
        + _chip_b.format(k="🏷️", v=f'{com_info_t14["commission_pct"]:,.0f}%')
        + '</div>'
        + f'<div style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;justify-content:space-between;margin-top:4px;'
          f'color:{profit_color_t14};font-size:13px;font-weight:800;">'
          + (f'<span>🧮 صافي الربح <span dir="ltr">{net_profit_t14:,.2f}</span> ريال</span>' if not missing_parts_t14
             else '<span style="color:#ef4444;">🧮 صافي الربح غير محسوب</span>') +
          f'<span style="display:flex;gap:5px;flex-wrap:wrap;align-items:center;">{_pill_v}{_reco_pill}</span></div>{_reco_line}'
        + missing_note_t14 + '</div>',
        unsafe_allow_html=True)
    if part in ("all", "calc"):
      with st.expander("💳 الحساب"):
        st.markdown(
            f"💳 **الصافي بعد خصم العمولة والتوصيل:** {net_fees_t14:,.2f} ريال  \n"
            f"🧾 **الصافي بعد خصم 15% ضريبة:** {net_tax_t14:,.2f} ريال  \n"
            f"📦 **التكلفة | Cost:** {cost_t14:,.2f} ريال"
            + (f"  \n🏷️ **سعر العرض (معلومة فقط):** {offer_price_t14:,.2f} ريال" if offer_price_t14 is not None else ""))
    if ad_detail_t14 and part in ("all", "verdict"):
        with st.expander("📊 الحكم"):
            st.markdown(ad_detail_t14)

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

                sku_col_tc = code_col_tc = cost_col_tc = None
                for c in df_tc.columns:
                    cl = str(c).strip().lower()
                    if sku_col_tc is None and "sku" in cl:
                        sku_col_tc = c
                    if code_col_tc is None and (("كود" in str(c) and "01" in str(c)) or cl in ("code01", "code 01")):
                        code_col_tc = c
                    if cost_col_tc is None and ("cost" in cl or "تكلفة" in str(c)):
                        cost_col_tc = c
                if not sku_col_tc:
                    sku_col_tc = df_tc.columns[0]
                if not code_col_tc:
                    # fallback: أقرب عمود اسمه فيه "كود"
                    for c in df_tc.columns:
                        if "كود" in str(c):
                            code_col_tc = c
                            break

                st.info(f"📊 {len(df_tc)} صف | SKU: `{sku_col_tc}` | الكود 01: `{code_col_tc}` | التكلفة: `{cost_col_tc or '—'}`")
                _preview_cols_tc = [sku_col_tc] + ([code_col_tc] if code_col_tc else []) + ([cost_col_tc] if cost_col_tc else [])
                st.dataframe(df_tc[_preview_cols_tc].head(10),
                             use_container_width=True, height=180)

                def do_upload_tc(replace=True):
                    dn = now_str()
                    to_add = []
                    for _, row in df_tc.iterrows():
                        sku = str(row[sku_col_tc]).strip()
                        code = str(row[code_col_tc]).strip() if code_col_tc else ""
                        cost = str(row[cost_col_tc]).strip() if cost_col_tc else ""
                        if cost.lower() == "nan":
                            cost = ""
                        if sku and sku.lower() != "nan":
                            to_add.append([sku, code, cost, dn])
                    if replace:
                        safe_delete_all(tacweed_sheet)
                    safe_batch_append(tacweed_sheet, to_add)
                    clear_cache(tacweed_sheet)
                    st.session_state.pop("warehouse_stockout_rows", None)
                    st.session_state.pop("warehouse_stockout_sku_lookup", None)
                    st.session_state.pop("warehouse_stockout_asin_lookup", None)
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

# ══ warehouse stock map (Code -> Quantity) — مربوط بالكود 01 من التكويد، ومكاش برضه
#    عشان مبيتبنيش من الصفر لكل SKU بيتعرض | Code -> Quantity, tied to the Tacweed
#    Code01, and cached too so it isn't rebuilt from scratch for every displayed SKU ══

def get_warehouse_stock_map():
    return _memo_from_sheet(warehouse_stock_sheet, "wh_stock_map", _build_warehouse_stock_map)

def tacweed_cost_badge_html(sku_up):
    """شارة صغيرة لتكلفة الوحدة لهذا الـ SKU (مباشرة من صفه في تاب التكويد، أو عن طريق
    الكود01 المشترك) — وترجع سلسلة فاضية لو مفيش تكلفة مسجلة | Small badge for this SKU's
    unit cost (directly, or via its shared Code01) — empty string if none recorded."""
    cost = get_sku_cost(sku_up)
    if cost is None:
        return ""
    return f'<span class="tacweed-badge" style="background:#0c4a6e;color:#7dd3fc;">💰 التكلفة: {cost:,.2f} ريال</span>'

def warehouse_available_badge(sku_up):
    """يعرض الكود 01 + تكلفته (لو مسجلة في تاب التكويد) + الكمية المتوفرة بالمستودع (لو
    الكود موجود ف ملف المستودع) جنب الSKU + شارة توقع النفاد/موعد الطلب لو محسوبة
    (compute_warehouse_stockout_rows) — الشارة دي محسوبة على مستوى الكود ككل (مجموع
    مبيعات كل الـ SKUs المشتركة فيه)، مش على SKU لوحده."""
    code = get_tacweed_map().get(sku_up, "")
    if not code:
        return ""
    qty = get_warehouse_stock_map().get(code, "")
    cost_badge_html = tacweed_cost_badge_html(sku_up)
    cost_part = f" {cost_badge_html}" if cost_badge_html else ""

    stockout_html = ""
    wh_row = st.session_state.get("warehouse_stockout_sku_lookup", {}).get(sku_up)
    if wh_row:
        dts = wh_row["days_to_stockout"]
        dts_txt = f"نفاد خلال {dts} يوم" if dts is not None else "لا يوجد سحب مسجل"
        status_short = wh_row["status"].split(" | ")[0]
        _cn = wh_row.get("china")
        _cn_txt = china_incoming_badge_text(_cn, datetime.now().date())
        stockout_html = (
            f' <span class="tacweed-badge" style="background:{wh_row["color"]}22;'
            f'color:{wh_row["color"]};border:1px solid {wh_row["color"]}55;">'
            f'{status_short} — {dts_txt}{_cn_txt}</span>'
        )

    if qty == "":
        return f'<span class="tacweed-badge" style="background:#3b0764;color:#e9d5ff;">🏷️ تكويد: {code}</span>{cost_part}{stockout_html}'
    return (
        f'<span class="tacweed-badge" style="background:#3b0764;color:#e9d5ff;">🏷️ تكويد: {code}</span>{cost_part} '
        f'<span class="tacweed-badge" style="background:#78350f;color:#fde68a;">📦 المتوفر بالمستودع: {qty}</span>'
        f'{stockout_html}'
    )

def tacweed_cost_badge_html_amazon(asin_up):
    """نفس tacweed_cost_badge_html بالظبط بس لتكلفة أمازون (get_asin_cost)."""
    cost = get_asin_cost(asin_up)
    if cost is None:
        return ""
    return f'<span class="tacweed-badge" style="background:#0c4a6e;color:#7dd3fc;">💰 التكلفة: {cost:,.2f} ريال</span>'

def warehouse_available_badge_amazon(asin_up):
    """نفس warehouse_available_badge بالظبط لكن لمنتج أمازون بالـ ASIN — بتقرا
    الكود01 والتكلفة من تاب TacweedAmazon، وبتاخد نفس كمية WarehouseStock
    المشتركة مع نون (نفس الكود = نفس المستودع الفعلي)، وشارة توقع النفاد بتتقرا
    من warehouse_stockout_asin_lookup اللي بيحسب السحب مجمّع من نون + أمازون
    مع بعض | Same as warehouse_available_badge but for an Amazon ASIN — reads
    Code01/cost from TacweedAmazon, shares the same WarehouseStock quantity as
    Noon (same code = same physical stock), and its stockout badge comes from
    warehouse_stockout_asin_lookup, which already combines Noon + Amazon draw."""
    code = get_tacweed_amazon_map().get(asin_up, "")
    if not code:
        return ""
    qty = get_warehouse_stock_map().get(code, "")
    cost_badge_html = tacweed_cost_badge_html_amazon(asin_up)
    cost_part = f" {cost_badge_html}" if cost_badge_html else ""

    stockout_html = ""
    wh_row = st.session_state.get("warehouse_stockout_asin_lookup", {}).get(asin_up)
    if wh_row:
        dts = wh_row["days_to_stockout"]
        dts_txt = f"نفاد خلال {dts} يوم" if dts is not None else "لا يوجد سحب مسجل"
        status_short = wh_row["status"].split(" | ")[0]
        _cn = wh_row.get("china")
        _cn_txt = china_incoming_badge_text(_cn, datetime.now().date())
        stockout_html = (
            f' <span class="tacweed-badge" style="background:{wh_row["color"]}22;'
            f'color:{wh_row["color"]};border:1px solid {wh_row["color"]}55;">'
            f'{status_short} — {dts_txt}{_cn_txt}</span>'
        )

    if qty == "":
        return f'<span class="tacweed-badge" style="background:#3b0764;color:#e9d5ff;">🏷️ تكويد: {code}</span>{cost_part}{stockout_html}'
    return (
        f'<span class="tacweed-badge" style="background:#3b0764;color:#e9d5ff;">🏷️ تكويد: {code}</span>{cost_part} '
        f'<span class="tacweed-badge" style="background:#78350f;color:#fde68a;">📦 المتوفر بالمستودع: {qty}</span>'
        f'{stockout_html}'
    )

def render_tacweed_amazon_upload(key_prefix):
    """واجهة رفع ملف تكويد أمازون (ASIN + Code01 + التكلفة) — نفس شكل
    render_tacweed_upload بالظبط بس بتدور على عمود ASIN بدل SKU، ومفيهاش
    اكتشاف تلقائي لشيت \"noon\" لأن الملف أصلاً بتاع أمازون."""
    with st.expander("📤 رفع ملف تكويد أمازون | Upload Amazon Tacweed (ASIN Mapping) File", expanded=False):
        upl_tca = st.file_uploader(
            "ارفع ملف تكويد أمازون (xlsx/xlsm/csv) | Upload Amazon Tacweed File",
            type=["xlsx", "xls", "xlsm", "csv"], key=f"{key_prefix}_tacweed_amz_upload")
        if upl_tca:
            try:
                if upl_tca.name.lower().endswith(".csv"):
                    df_tca = pd.read_csv(upl_tca, dtype=str).fillna("")
                else:
                    xls_a = pd.ExcelFile(upl_tca)
                    sheet_names_a = xls_a.sheet_names
                    chosen_sheet_a = None
                    for sn in sheet_names_a:
                        if "amazon" in sn.strip().lower() or "امازون" in sn.strip():
                            chosen_sheet_a = sn
                            break
                    if not chosen_sheet_a:
                        if len(sheet_names_a) > 1:
                            chosen_sheet_a = st.selectbox(
                                "اختار الشيت اللي فيه التكويد | Choose the sheet with the mapping",
                                sheet_names_a, key=f"{key_prefix}_tacweed_amz_sheet_pick")
                        else:
                            chosen_sheet_a = sheet_names_a[0]
                    df_tca = pd.read_excel(upl_tca, sheet_name=chosen_sheet_a, dtype=str).fillna("")

                asin_col_tca = code_col_tca = cost_col_tca = None
                for c in df_tca.columns:
                    cl = str(c).strip().lower()
                    if asin_col_tca is None and "asin" in cl:
                        asin_col_tca = c
                    if code_col_tca is None and (("كود" in str(c) and "01" in str(c)) or cl in ("code01", "code 01")):
                        code_col_tca = c
                    if cost_col_tca is None and ("cost" in cl or "تكلفة" in str(c)):
                        cost_col_tca = c
                if not asin_col_tca:
                    asin_col_tca = df_tca.columns[0]
                if not code_col_tca:
                    for c in df_tca.columns:
                        if "كود" in str(c):
                            code_col_tca = c
                            break

                st.info(f"📊 {len(df_tca)} صف | ASIN: `{asin_col_tca}` | الكود 01: `{code_col_tca}` | التكلفة: `{cost_col_tca or '—'}`")
                _preview_cols_tca = [asin_col_tca] + ([code_col_tca] if code_col_tca else []) + ([cost_col_tca] if cost_col_tca else [])
                st.dataframe(df_tca[_preview_cols_tca].head(10),
                             use_container_width=True, height=180)

                def do_upload_tca(replace=True):
                    dn = now_str()
                    to_add = []
                    for _, row in df_tca.iterrows():
                        asin = str(row[asin_col_tca]).strip()
                        code = str(row[code_col_tca]).strip() if code_col_tca else ""
                        cost = str(row[cost_col_tca]).strip() if cost_col_tca else ""
                        if cost.lower() == "nan":
                            cost = ""
                        if asin and asin.lower() != "nan":
                            to_add.append([asin, code, cost, dn])
                    if replace:
                        safe_delete_all(tacweed_amazon_sheet)
                    safe_batch_append(tacweed_amazon_sheet, to_add)
                    clear_cache(tacweed_amazon_sheet)
                    st.session_state.pop("warehouse_stockout_rows", None)
                    st.session_state.pop("warehouse_stockout_sku_lookup", None)
                    st.session_state.pop("warehouse_stockout_asin_lookup", None)
                    return len(to_add)

                ca_tca, cb_tca = st.columns(2)
                with ca_tca:
                    if st.button("📤 إضافة للموجود | Append", type="primary",
                                 use_container_width=True, key=f"{key_prefix}_tca_append"):
                        n = do_upload_tca(replace=False)
                        st.success(f"✅ أُضيف {n} صف | rows added")
                        st.rerun()
                with cb_tca:
                    if st.button("🔄 استبدال الكل | Replace All", type="secondary",
                                 use_container_width=True, key=f"{key_prefix}_tca_replace"):
                        st.session_state[f"{key_prefix}_confirm_replace_tca"] = True
                if st.session_state.get(f"{key_prefix}_confirm_replace_tca"):
                    st.warning("⚠️ هيمسح كل تكويد أمازون القديم ويرفع الجديد؟ | Replace all existing Amazon mapping?")
                    cy_tca, cn_tca = st.columns(2)
                    if cy_tca.button("✅ نعم | Yes", key=f"{key_prefix}_yes_rep_tca"):
                        n = do_upload_tca(replace=True)
                        st.session_state[f"{key_prefix}_confirm_replace_tca"] = False
                        st.success(f"✅ تم الاستبدال — {n} صف")
                        st.rerun()
                    if cn_tca.button("❌ لا | No", key=f"{key_prefix}_no_rep_tca"):
                        st.session_state[f"{key_prefix}_confirm_replace_tca"] = False
                        st.rerun()
            except Exception as e:
                st.error(f"❌ {e}")





def _drive_call(http, method, url, body=None, headers=None):
    import json as _json
    resp, content = http.request(url, method=method, body=body, headers=headers or {})
    if resp.status >= 300:
        raise RuntimeError(f"Drive {resp.status}: {content[:200]!r}")
    return _json.loads(content or b"{}")

def _drive_http():
    import httplib2
    _c = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
    return _c.authorize(httplib2.Http(timeout=60))

def _shrink_jpeg(raw, side=500):
    from PIL import Image
    im = Image.open(io.BytesIO(raw))
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGB", im.size, (255, 255, 255))
        bg.paste(im, mask=im.split()[-1])
        im = bg
    else:
        im = im.convert("RGB")
    im.thumbnail((side, side))
    out = io.BytesIO()
    im.save(out, "JPEG", quality=85)
    return out.getvalue()

def _drive_list_folder(folder_id):
    """أسماء وIDs الملفات في مجلد Drive (لازم المجلد يكون متشارك مع إيميل حساب الخدمة أو بـ«أي حد معاه الرابط») | list files in a Drive folder."""
    import urllib.parse as _up
    http = _drive_http()
    out, tok = [], ""
    while True:
        q = _up.quote(f"'{folder_id}' in parents and trashed=false")
        url = (f"https://www.googleapis.com/drive/v3/files?q={q}&pageSize=1000&fields=nextPageToken,files(id,name)"
               f"&supportsAllDrives=true&includeItemsFromAllDrives=true" + (f"&pageToken={tok}" if tok else ""))
        d = _drive_call(http, "GET", url)
        out += d.get("files", [])
        tok = d.get("nextPageToken", "")
        if not tok:
            return out

def render_china_images_upload(key_prefix):
    """صور منتجات الصين: (1) نطلّع الصور من ملف الجرد ونجهّزها ZIP، (2) انت ترفعها على مجلد Drive بتاعك، (3) نربط المجلد بشيت ChinaImages.
    حساب الخدمة مالوش مساحة تخزين على Drive، عشان كده الرفع بيتم من عندك | service accounts have no Drive quota, so you upload the ZIP contents yourself."""
    with st.expander("🖼️ صور منتجات الصين | China product images", expanded=False):
        st.markdown("**الخطوة 1 — جهّز الصور:** ارفع ملف الجرد (فيه عمود «رمز الباركود» وعمود «الصورة») ونطلّع لك ZIP فيه كل صورة باسم `الباركود.jpg` مصغّرة.")
        upl = st.file_uploader("ملف الجرد (xlsx)", type=["xlsx", "xlsm"], key=f"{key_prefix}_chimg_up")
        if upl:
            if st.session_state.get("_chimg_name") != (upl.name, upl.size):
                try:
                    with st.spinner("جاري استخراج الصور..."):
                        pics = extract_excel_images(upl)
                        import zipfile as _zf
                        zb = io.BytesIO()
                        with _zf.ZipFile(zb, "w", _zf.ZIP_STORED) as zf:
                            for bc, raw in pics:
                                try:
                                    zf.writestr(f"{bc}.jpg", _shrink_jpeg(raw))
                                except Exception:
                                    pass
                    st.session_state["_chimg_name"] = (upl.name, upl.size)
                    st.session_state["_chimg_zip"] = (zb.getvalue(), len(pics))
                except Exception as e:
                    st.error(f"❌ تعذر قراءة الملف: {e}")
                    return
            zdata, zn = st.session_state["_chimg_zip"]
            if zn:
                st.success(f"🖼️ {zn} صورة جاهزة")
                st.download_button("⬇️ نزّل ZIP الصور", zdata, file_name="china_product_images.zip", mime="application/zip",
                                   key=f"{key_prefix}_chimg_dl", use_container_width=True)
            else:
                st.warning("مالقيتش صور مرتبطة بباركود في الملف")
        st.markdown("**الخطوة 2 — ارفعها على Drive:** اعمل مجلد في Google Drive بتاعك، فك الـ ZIP وارفع الصور جواه (سحب وإفلات). "
                    "بعدين اضغط مشاركة على المجلد واختار «أي شخص معه الرابط — مشاهد» (عشان الصور تظهر في التطبيق والشيت).")
        try:
            _sa = st.secrets["gcp_service_account"]["client_email"]
        except Exception:
            _sa = ""
        st.markdown("**الخطوة 3 — اربط المجلد:** شارك المجلد كمان مع إيميل التطبيق (حساب الخدمة) بصلاحية «مشاهد» عشان يقدر يقرا أسماء الملفات، "
                    "وبعدين الصق رابط المجلد أو الـ ID بتاعه.")
        if _sa:
            st.code(_sa, language=None)
        cur = load_settings().get("china_images_folder_id", "")
        raw_id = st.text_input("رابط/ID مجلد Drive", value=cur, key=f"{key_prefix}_chimg_folder").strip()
        m = re.search(r"folders/([A-Za-z0-9_-]+)", raw_id)
        fid = m.group(1) if m else raw_id
        if fid and st.button("🔗 اربط الصور من المجلد | Link images", key=f"{key_prefix}_chimg_go", type="primary", use_container_width=True):
            try:
                with st.spinner("جاري قراءة المجلد..."):
                    files = _drive_list_folder(fid)
            except Exception as e:
                st.error(f"❌ مقدرتش أقرا المجلد: {e}")
                st.info("اتأكد إن المجلد متشارك «أي شخص معه الرابط» أو متشارك مع إيميل حساب الخدمة (client_email في الـ secrets).")
                return
            ts, rows, seen = now_str(), [], set()
            for f in files:
                nm = f["name"].rsplit(".", 1)[0].strip()
                if nm and nm not in seen and re.search(r"\.(jpe?g|png|webp)$", f["name"], re.I):
                    seen.add(nm)
                    rows.append([nm, f"https://drive.google.com/thumbnail?id={f['id']}&sz=w500", f["id"], ts])
            if not rows:
                st.warning("مالقيتش صور في المجلد. غالبًا المجلد مش متشارك مع إيميل التطبيق (حساب الخدمة) — شاركه بصلاحية «مشاهد» — أو الصور جوه مجلد فرعي أو الصيغة مش jpg/png.")
                return
            if fid != cur:
                save_setting("china_images_folder_id", fid)
            safe_delete_all(china_img_sheet)
            for i in range(0, len(rows), 2000):
                china_img_sheet.append_rows(rows[i:i + 2000], value_input_option="RAW")
            clear_cache(china_img_sheet)
            st.session_state.pop("china_bundle", None)
            st.success(f"✅ اتربطت {len(rows)} صورة بشيت ChinaImages")
        st.divider()
        st.caption("بعد أول ربط، اضغط مرة واحدة عشان شيت China يبدأ ياخد الصور من ChinaImages:")
        if st.button("🔄 حدّث معادلة صور شيت China | Refresh image formulas", key=f"{key_prefix}_chimg_formula"):
            with st.spinner("جاري التحديث..."):
                _china_write_image_formulas(get_china_flat_ws())
            st.success("✅ تم")

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
                    st.session_state.pop("warehouse_stockout_rows", None)
                    st.session_state.pop("warehouse_stockout_sku_lookup", None)
                    st.session_state.pop("warehouse_stockout_asin_lookup", None)
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

# ══════════════════════════════════════════════════════════════════════════════
# ══ 🚢 شحنات الصين | China Shipments ══
# ملف إكسيل (Smart_Store_China_Data_Template) فيه 3 صفحات بيانات: SHIPMENTS / SKU_LINES /
# CONTAINERS. بيتخزنوا في 3 تابات جديدة في نفس جوجل شيت (ChinaShipments / ChinaLines /
# ChinaContainers) عشان يفضلوا محفوظين بين الجلسات. الربط مع المتجر بيتم عن طريق التكويد:
# الكود01 في تاب التكويد = عمود barcode في صفحة SKU_LINES. كل ده إضافة فقط — مفيش أي
# دالة أو تاب قديم اتغيّر منطقه | The China Excel's 3 data sheets are stored in 3 new Google
# Sheet tabs. Linking is done through Tacweed: Code01 == the barcode column in SKU_LINES.
# Purely additive — no existing function or tab had its logic changed.
# ══════════════════════════════════════════════════════════════════════════════
CHINA_TABS = {"ChinaShipments": CHINA_SHIP_COLS, "ChinaLines": CHINA_LINE_COLS, "ChinaContainers": CHINA_CONT_COLS}
# حقول محسوبة بالإكسيل بتتأخر (متحسوبة وقت آخر حفظ) — بنحسبها إحنا لحظيًا بدلها
# مراحل رحلة البضاعة من أول تأكيد المصنع لحد استلام المستودع | The full journey, in order

def _china_get_or_create_ws(tab, cols, retries=5, delay=2):
    for attempt in range(retries):
        try:
            return ss.worksheet(tab)
        except gspread.exceptions.WorksheetNotFound:
            try:
                ws = ss.add_worksheet(title=tab, rows="1000", cols=str(len(cols)))
                ws.append_row(cols, value_input_option="RAW")
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


# ── تنظيف القيم / Value cleaning ──





# ── مفاتيح السجلات (للتحديث بدل التكرار عند رفع نفس الملف تاني) | Row keys for upsert ──




# ── قراءة ملف الإكسيل | Parse the uploaded workbook ──


# ── قراءة/كتابة تابات الصين في جوجل شيت | Read / write the China tabs ──






# ══════════════════════════════════════════════════════════════════════════════
# ══ شيت الصين الواحد | Single China sheet ══
# الموظف بيدخل كل حاجة في شيت واحد اسمه China — كل صف = منتج/دفعة داخل طلب، وبيانات الطلب والحاوية بتتكرر على الصف.
# التطبيق بيفكّه تلقائيًا لـ (شحنات / حاويات / منتجات) نفس الـ 3 جداول اللي باقي الكود شغال عليهم.
# One row = one product lot; order + container data repeat on the row. The app splits it back into the same
# 3 tables (shipments / containers / lines) that the rest of the code already works with.
# ══════════════════════════════════════════════════════════════════════════════
CHINA_FLAT_TAB = "China"
CHINA_FLAT_HDR_ROWS = 3          # صف 1 مجموعات | صف 2 مفاتيح | صف 3 عناوين عربي | البيانات من صف 4
# (المجموعة, المفتاح, العنوان العربي, العرض, النوع: t نص / d تاريخ / n رقم / g عام)
# ── الحالة الواحدة (بتحل محل حالة الطلب + حالة الحاوية + مرحلة المخزون) ──
# (الكود الإنجليزي الداخلي, العنوان عربي | English) — الترتيب = درجة التقدم










CHINA_FLAT_LISTS = {
 "status": [x[1] for x in CHINA_STATUS],
 "shipment_mode": ["SEA - FCL","SEA - LCL","AIR","COURIER","TRUCK"],
 "incoterm": ["EXW","FCA","FOB","CFR","CIF","CPT","CIP","DAP","DPU","DDP"],
 "currency": ["USD","CNY","SAR","EUR","HKD","AED"],
 "container_type": ["20GP","40GP","40HQ","45HQ","LCL","AIR"],
 "customs_status": ["Not Started","Documents Ready","Submitted","Under Clearance","Inspection","Cleared","On Hold"],
 "shipping_line": ["MSC","CMA CGM","Maersk","COSCO","Hapag-Lloyd","OOCL","Evergreen","ONE","PIL","TSL","ECON","WHL","ZG中古","Other"],
}











# ── دليل الموظف + سكريبت اختياري | Staff guide + optional Apps Script ──
CHINA_STAFF_GUIDE = """📋 دليل تعبئة شيت الصين

القاعدة: الشيت واحد (China). كل صف = منتج واحد داخل طلب، وبيانات الطلب والحاوية بتتكرر على صف كل منتج.
ألوان الأعمدة: 🔑 الباركود والصورة والحالة | 🔵 الطلب/الشحنة | 🟠 الحاوية والتتبع | 🟢 المنتج | ⚪ التحديث | ⚙️ رمادي = محسوب تلقائي، لا تكتب فيه.
الأعمدة المتثبتة (رقم الطلب، PO، الباركود، الصورة) بتفضل قدامك وإنت بتمشي يمين.

⭐ عمود «الحالة | Status» هو العمود الوحيد للحالة (اختر من القايمة). كل صف حالته لوحده، والتطبيق بيحسب منها حالة الحاوية وحالة الشحنة (الطلب + PO) لوحده.
عمود «المطلوب دلوقتي» بيقولك تدخل إيه بعد كده، ولو ظهر ⚠ يبقى فيه تاريخ أحدث من الحالة المختارة.

الخطوة 1 — طلب جديد من المصنع (صف لكل منتج):
 1) 🔵 رقم الطلب الداخلي (NO.xxx) + رقم PO + اسم المورد + شركة المشتري.
 2) 🔑 الباركود (الصورة بتظهر لوحدها) + الحالة: تأكيد المصنع ← قيد الإنتاج ← اكتمل الإنتاج ← جاهز للشحن.
 3) 🔵 طريقة الشحن + Incoterm + العملة + قيمة الفاتورة.
 4) 🟢 SKU + الموديل + اسم المنتج EN/AR + الكراتين + قطع/كرتون + إجمالي الكمية + تكلفة الوحدة + CBM + الأوزان + تاريخ جاهزية الإنتاج، وSKU نون/أمازون/تريندول لو معروفين.
 5) لو الطلب فيه أكتر من منتج: صف لكل منتج، وانسخ بيانات 🔵 الطلب على كل صفوفه.
 معرّف السطر اتركه فاضي — بيتولّد تلقائي.

الخطوة 2 — تجهيز الحاوية:
 - على صف كل منتج داخل الحاوية: 🟠 رقم الحاوية + رقم الختم + نوع الحاوية (40HQ ...).
 - الحالة: «حاوية متجهزة» ثم «تم التحميل» بعد التحميل.
 - المنتج الموزع على أكتر من حاوية: صف لكل حاوية، وكل صف بكمية وCBM الحاوية دي بس.

الخطوة 3 — تم الشحن:
 - 🔵 رقم الشحنة STxxxxx (بعد تخصيصها فقط — لا تكتب NO.xxx هنا).
 - 🟠 شركة الشحن + B/L + اسم السفينة + رقم الرحلة + ميناء التحميل + ميناء الوصول + ETD الفعلي + ETA الميناء + رابط التتبع.
 - الحالة: «في البحر».

الخطوة 4 — تأخير أو وصول:
 - ETA اتغيّر: عدّل ETA واكتب سبب التأخير.
 - عند الوصول فعلاً: الحالة «وصلت الميناء» ← «تحت التخليص» ← «تم التخليص» ← «وصلت المستودع»، وحالة التخليص الجمركي.
 - الوصول الفعلي يُكتب بعد الوصول الحقيقي فقط، مش ETA.

التواريخ والحالة الأوتوماتيك (لو سكريبت Apps Script متركّب مرة واحدة): أي تعديل في الصف بيسجّل «آخر تحديث» لوحده. وأول ما تكتب تاريخ فعلي (التحميل / ETD الفعلي / وصول الميناء / وصول المستودع) أو رقم حاوية أو حالة تخليص، «الحالة» بتتحدّث لوحدها (للأمام بس). ولو غيّرت «الحالة» بإيدك بيتملا تاريخ الحالة الفعلي لو فاضي (تم التحميل ← تاريخ التحميل، في البحر ← ETD الفعلي، وصلت الميناء ← وصول الميناء، وصلت المستودع ← وصول المستودع)، وبتتوحّد الحالة وبيانات الحاوية على كل صفوف نفس الحاوية. من غير السكريبت اكتب التواريخ بإيدك.

قواعد مهمة:
 - لا تغيّر الصفوف 1–3 (العناوين) ولا ترتّب الأعمدة.
 - NO.xxx = رقم الطلب الداخلي فقط. STxxxxx = رقم الشحنة بعد تخصيصها.
 - RIYADH = الوجهة النهائية وليست ميناء. الميناء = JEDDAH أو DAMMAM.
 - عمود «اكتمال الصف»: MISSING = ناقص بيانات مطلوبة، OK = تمام.
 - الطلب أو المنتج الملغي: الحالة = ملغي."""

CHINA_APPS_SCRIPT = """// اختياري: Extensions ▸ Apps Script ▸ الصق الكود ▸ Save ▸ (أول مرة هيطلب صلاحية).
// 1) بيختم «آخر تحديث» لوحده. 2) أول ما تكتب تاريخ فعلي أو رقم حاوية بتتحدّث «الحالة» لوحدها (للأمام بس، عمرها ما بترجّع).
// 3) أول ما تغيّر «الحالة» بيتملا تاريخ الحالة لو فاضي. 4) بيوحّد الحالة وبيانات الحاوية على كل صفوف نفس الحاوية.
var CONT_COLS = ["container_type","seal_no","shipping_line","vessel_name","voyage_no","bl_no","port_of_loading","destination_port",
  "actual_etd","eta_port","port_arrival_actual","warehouse_arrival_actual","customs_status","last_known_location",
  "delay_reason","tracking_url","ctr_gross_weight_kg","ctr_total_cbm","ctr_total_cartons","ctr_total_units"];
var FULL = ["تأكيد المصنع | Factory confirmed","قيد الإنتاج | In production","اكتمل الإنتاج | Production completed","جاهز للشحن | Ready to ship",
  "حاوية متجهزة | Container prepared","تم التحميل | Loaded","في البحر | In transit","وصلت الميناء | At port","تحت التخليص | In customs",
  "تم التخليص | Customs cleared","وصلت المستودع | Delivered to warehouse"];
var STATUS_DATES = [["تم التحميل","actual_load_date"],["في البحر","actual_etd"],["وصلت الميناء","port_arrival_actual"],
  ["وصلت المستودع","warehouse_arrival_actual"],["اكتمل الإنتاج","production_ready_date"],["جاهز للشحن","production_ready_date"]];
var AUTO_BY_FIELD = {actual_load_date: 5, actual_etd: 6, port_arrival_actual: 7, warehouse_arrival_actual: 10, container_no: 4};
var CUSTOMS = {"Under Clearance": 8, "Inspection": 8, "Submitted": 8, "Cleared": 9};
function rankOf(s) {
  s = String(s);
  for (var i = FULL.length - 1; i >= 0; i--) if (s.indexOf(FULL[i].split(" | ")[0]) >= 0) return i;
  return -1;
}
function onEdit(e) {
  var sh = e.range.getSheet();
  if (sh.getName() !== "China" || e.range.getRow() < 4) return;
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;
  var keys = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0];
  var K = function (k) { return keys.indexOf(k) + 1; };
  var row = e.range.getRow(), col = e.range.getColumn(), key = keys[col - 1];
  if (key === "last_updated_at" || key === "updated_by") return;
  var stamp = K("last_updated_at"), by = K("updated_by"), sc = K("status"), now = new Date();
  var who = ""; try { who = Session.getActiveUser().getEmail(); } catch (x) {}
  var touch = function (r) {
    if (stamp > 0) sh.getRange(r, stamp).setValue(now);
    if (by > 0 && who) sh.getRange(r, by).setValue(who);
  };
  var fillDate = function (r, status) {
    for (var i = 0; i < STATUS_DATES.length; i++) {
      if (String(status).indexOf(STATUS_DATES[i][0]) >= 0) {
        var c = K(STATUS_DATES[i][1]);
        if (c > 0 && !sh.getRange(r, c).getValue()) sh.getRange(r, c).setValue(now);
      }
    }
  };
  var advance = function (r, rank) {                      // الحالة تتحرك للأمام بس، وماتلمسش «ملغي/موقوف»
    var cell = sh.getRange(r, sc), cur = String(cell.getValue());
    if (cur.indexOf("ملغي") >= 0 || cur.indexOf("موقوف") >= 0 || rankOf(cur) >= rank) return false;
    cell.setValue(FULL[rank]); fillDate(r, FULL[rank]); touch(r); return true;
  };
  touch(row);
  var val = e.range.getValue(), isStatus = key === "status", isCont = CONT_COLS.indexOf(key) >= 0, rank = -1;
  if (val !== "" && sc > 0) {
    if (AUTO_BY_FIELD[key] !== undefined) rank = AUTO_BY_FIELD[key];
    else if (key === "customs_status" && CUSTOMS[val] !== undefined) rank = CUSTOMS[val];
    if (rank >= 0) advance(row, rank);
  }
  if (isStatus) fillDate(row, val);
  if (!isStatus && !isCont) return;
  var cn = K("container_no");
  var no = sh.getRange(row, cn).getValue(), last = sh.getLastRow();
  if (!no || last < 4) return;
  var nos = sh.getRange(4, cn, last - 3, 1).getValues();
  var sts = sh.getRange(4, sc, last - 3, 1).getValues();
  for (var i = 0; i < nos.length; i++) {
    var r = i + 4;
    if (nos[i][0] !== no || r === row || String(sts[i][0]).indexOf("ملغي") >= 0) continue;
    sh.getRange(r, col).setValue(val);
    if (isStatus) fillDate(r, val); else if (rank >= 0) advance(r, rank);
    touch(r);
  }
}"""



CHINA_SYNC_SCRIPT = r"""// ═══ مزامنة Container tracking → China (داخل الجوجل شيت نفسه) ═══
// التركيب (مرة واحدة): Extensions ▸ Apps Script ▸ ملف جديد (+) ▸ الصق الكود ده ▸ Save ▸ اختار الدالة ctSetup من القايمة ▸ Run ▸ وافق على الصلاحيات.
// بعدها: أي مرجع تكتبه في China (رقم الحاوية / البوليصة / الشحنة) بيتملا الباقي بعد ثواني، وكمان بيشتغل لوحده كل 5 دقايق.
// الكود ده بيكمّل مزامنة التطبيق ومش بيعارضها (نفس القواعد ونفس الأعمدة).
var CT_NAME = "Container tracking";
var CT_CHINA_TAB = "China";
var CT_STATUS = [["Factory Confirmed", "تأكيد المصنع | Factory confirmed"], ["In Production", "قيد الإنتاج | In production"],
  ["Production Completed", "اكتمل الإنتاج | Production completed"], ["Ready to Ship", "جاهز للشحن | Ready to ship"],
  ["Allocated to Container", "حاوية متجهزة | Container prepared"], ["Loaded", "تم التحميل | Loaded"],
  ["Shipped / In Transit", "في البحر | In transit"], ["Arrived at Port", "وصلت الميناء | At port"],
  ["Customs Clearance", "تحت التخليص | In customs"], ["Cleared", "تم التخليص | Customs cleared"],
  ["Delivered to Warehouse", "وصلت المستودع | Delivered to warehouse"], ["On Hold", "موقوف | On hold"], ["Cancelled", "ملغي | Cancelled"]];
var CT_ALIAS = {"order confirmed": "Factory Confirmed", "warehouse received": "Delivered to Warehouse", "booked": "Allocated to Container"};
var CT_HEAD = {
  loading: ["LOADING DATE", "تاريخ التحميل"], invoice: ["INVOICE NO", "رقم الفاتورة"], origin: ["ADD", "عنوان"],
  container: ["CONTAINER NO", "رقم الحاوية"], seal: ["SEAL NO", "رقم الختم"], bl: ["D/R NO", "رقم بوليصة الشحن"],
  sailing: ["SAILING DATE", "تاريخ الإبحار"], transfer_loc: ["TRANSFER PORT LOCATION", "موقع ميناء النقل"],
  transfer_date: ["DATE OF TRANSFER PORT", "تاريخ العبور"], dest: ["DESTINATION", "مقصد"], eta: ["ARRIVAL DATE", "تاريخ الوصول"],
  url: ["QUERY WEBSITE", "موقع الاستعلام"], actual: ["ACTUAL ARRIVAL DATE", "تاريخ الوصول الفعلي"], status: ["الحالة"],
  received: ["تاريخ الاستلام في المستودع"]};
var CT_MAP = [["shipment_id", "invoice"], ["container_no", "container"], ["seal_no", "seal"], ["bl_no", "bl"], ["actual_load_date", "loading"],
  ["actual_etd", "sailing"], ["destination_port", "dest"], ["eta_port", "eta"], ["vessel_name", "vessel"], ["voyage_no", "voyage"],
  ["tracking_url", "url"], ["last_known_location", "transfer_loc"], ["port_arrival_actual", "actual"],
  ["warehouse_arrival_actual", "received"], ["origin_location", "origin"]];
var CT_IDENT = {shipment_id: 1, container_no: 1}, CT_FILL_ONLY = {origin_location: 1};
var CT_DATES = {actual_load_date: 1, actual_etd: 1, eta_port: 1, port_arrival_actual: 1, warehouse_arrival_actual: 1};
var CT_FORMULA = {next_action: 1, line_amount: 1, days_to_eta: 1, delay_days: 1, row_state: 1, product_image: 1};
var CT_AR = "٠١٢٣٤٥٦٧٨٩";

function ctNorm(v) { return String(v == null ? "" : v).replace(/\s+/g, "").toUpperCase(); }
function ctTrim(v) { return String(v == null ? "" : v).trim(); }
function ctDigits(s) { return String(s).replace(/[٠-٩]/g, function (d) { return CT_AR.indexOf(d); }).replace(/٫/g, "."); }
function ctPad(n) { return (n < 10 ? "0" : "") + n; }
function ctMk(y, m, d) {
  var t = new Date(y, m - 1, d);
  if (t.getFullYear() !== y || t.getMonth() !== m - 1 || t.getDate() !== d) return "";
  return y + "-" + ctPad(m) + "-" + ctPad(d);
}
function ctParseDate(v, year, anchor) {
  var s = ctDigits(ctTrim(v));
  if (!s) return "";
  var m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(?:\D.*)?$/);
  if (m) return ctMk(+m[1], +m[2], +m[3]);
  m = s.match(/^(\d{1,2})[.\/\-](\d{1,2})$/);
  if (m && year) {
    var day = +m[2];
    if (day >= 40 && day <= 90 && day % 10 === 0) day = day / 10;     // 5٫8 بتتعرض 5٫80 (خلية بخانتين عشريتين) = 8 مايو
    var r = ctMk(+year, +m[1], day);
    if (r && anchor) {
      var a = new Date(anchor + "T00:00:00"), d = new Date(r + "T00:00:00");
      if (d < new Date(a.getTime() - 60 * 86400000)) r = ctMk(+year + 1, +m[1], day);
    }
    return r;
  }
  return "";
}
function ctLookup(txt) {
  var k = ctTrim(txt).toLowerCase();
  for (var i = 0; i < CT_STATUS.length; i++) {
    var c = CT_STATUS[i][0], l = CT_STATUS[i][1], parts = l.split("|");
    var ks = [c, l, parts[0], parts[parts.length - 1]];
    for (var j = 0; j < ks.length; j++) if (ks[j].trim().toLowerCase() === k) return c;
  }
  return CT_ALIAS[k] || "";
}
function ctProg(code) {
  for (var i = 0; i < CT_STATUS.length; i++) if (CT_STATUS[i][0] === code) return i <= 10 ? i : -1;
  return -1;
}
function ctLabel(code) { for (var i = 0; i < CT_STATUS.length; i++) if (CT_STATUS[i][0] === code) return CT_STATUS[i][1]; return ""; }
function ctFloor(r, code) {
  if (code && ctProg(code) < 0) return code;
  var im = "";
  if (r.warehouse_arrival_actual) im = "Delivered to Warehouse";
  else if (r.port_arrival_actual) im = "Arrived at Port";
  else if (r.actual_etd) im = "Shipped / In Transit";
  else if (r.actual_load_date) im = "Loaded";
  else if (r.container_no) im = "Allocated to Container";
  return ctProg(im) > ctProg(code) ? im : code;
}

function ctHeadMap(raw) {
  for (var hi = 0; hi < Math.min(12, raw.length); hi++) {
    var pos = {};
    for (var j = 0; j < raw[hi].length; j++) {
      var lines = String(raw[hi][j]).replace(/\r/g, "\n").split("\n").map(function (x) { return x.trim().toUpperCase().replace(/\.$/, ""); }).filter(String);
      if (pos.ship === undefined && lines.indexOf("اسم السفينة") >= 0) { pos.ship = j; continue; }
      for (var f in CT_HEAD) {
        if (pos[f] !== undefined) continue;
        var hit = CT_HEAD[f].some(function (n) { return lines.indexOf(n) >= 0; });
        if (hit) { pos[f] = j; break; }
      }
    }
    if (pos.container !== undefined && pos.invoice !== undefined) return {hi: hi, pos: pos};
  }
  return null;
}
function ctReadTracking(raw, nowYear) {
  var h = ctHeadMap(raw), out = [];
  if (!h) return out;
  var g = function (row, f) { return h.pos[f] !== undefined && h.pos[f] < row.length ? ctTrim(row[h.pos[f]]) : ""; };
  for (var i = h.hi + 1; i < raw.length; i++) {
    var row = raw[i], cno = ctNorm(g(row, "container")), inv = ctNorm(g(row, "invoice"));
    if (!cno && !inv) continue;
    var m = inv.match(/^ST(\d{2})/), year = m ? 2000 + parseInt(m[1], 10) : nowYear;
    var load = ctParseDate(g(row, "loading"), year), anchor = load || null;
    var sail = ctParseDate(g(row, "sailing"), year, anchor); anchor = anchor || sail || null;
    var eta = ctParseDate(g(row, "eta"), year, anchor);
    var ship = g(row, "ship"), voy = "", toks = ship.split(/\s+/).filter(String);
    if (toks.length > 1 && /^[A-Z]{0,3}\d{2,5}[A-Z]?$/.test(toks[toks.length - 1].toUpperCase())) {
      voy = toks[toks.length - 1].toUpperCase(); ship = toks.slice(0, -1).join(" ");
    }
    out.push({shipment_id: inv, container_no: cno, seal_no: g(row, "seal"), bl_no: ctNorm(g(row, "bl")), actual_load_date: load, actual_etd: sail,
      destination_port: g(row, "dest").toUpperCase(), eta_port: eta, vessel_name: ship, voyage_no: voy, tracking_url: g(row, "url"),
      last_known_location: g(row, "transfer_loc"), port_arrival_actual: ctParseDate(g(row, "actual"), year),
      warehouse_arrival_actual: ctParseDate(g(row, "received"), year), origin_location: g(row, "origin"), _status: g(row, "status")});
  }
  return out;
}

// china = getDisplayValues() لشيت China كله، tracking = recs من ctReadTracking. بيرجّع [{row, key, value}] (row من 1).
function ctPlan(china, recs) {
  var hi = -1, keysRow = [];
  for (var i = 0; i < Math.min(8, china.length); i++) {
    var low = china[i].map(ctTrim);
    if (low.indexOf("container_no") >= 0 && low.indexOf("internal_order_no") >= 0) { hi = i; keysRow = low; break; }
  }
  if (hi < 0 || !recs.length) return {updates: [], pos: {}};
  var pos = {};
  keysRow.forEach(function (k, j) { if (k && pos[k] === undefined) pos[k] = j; });
  var byC = {}, byBl = {}, byInv = {};
  recs.forEach(function (t) {
    if (t.container_no) (byC[t.container_no] = byC[t.container_no] || []).push(t);
    if (t.bl_no) (byBl[t.bl_no] = byBl[t.bl_no] || []).push(t);
    if (t.shipment_id) (byInv[t.shipment_id] = byInv[t.shipment_id] || []).push(t);
  });
  var inputKeys = Object.keys(pos).filter(function (k) { return !CT_FORMULA[k]; }), updates = [];
  for (var r = hi + 2; r < china.length; r++) {                    // hi+1 = صف العناوين العربي
    var row = china[r], cur = {};
    Object.keys(pos).forEach(function (k) { cur[k] = pos[k] < row.length ? ctTrim(row[pos[k]]) : ""; });
    if (!inputKeys.some(function (k) { return cur[k]; })) continue;
    var cno = ctNorm(cur.container_no), bl = ctNorm(cur.bl_no), sid = ctNorm(cur.shipment_id);
    var recsFor = cno ? byC[cno] : null;
    if (!recsFor && bl) recsFor = byBl[bl];
    if (!recsFor && sid) recsFor = byInv[sid];
    if (!recsFor) continue;
    var vals = {}, st = "";
    if (recsFor.length === 1) {
      CT_MAP.forEach(function (p) { vals[p[0]] = recsFor[0][p[0]] || ""; });
      st = recsFor[0]._status;
    } else {
      CT_MAP.forEach(function (p) {
        var set = {}; recsFor.forEach(function (x) { set[x[p[0]] || ""] = 1; });
        var ks = Object.keys(set); vals[p[0]] = ks.length === 1 ? ks[0] : "";
      });
      var ss = {}; recsFor.forEach(function (x) { ss[x._status] = 1; });
      var sk = Object.keys(ss); st = sk.length === 1 ? sk[0] : "";
    }
    var nw = {}, ch = {};
    Object.keys(cur).forEach(function (k) { nw[k] = cur[k]; });
    CT_MAP.forEach(function (p) {
      var k = p[0], tv = vals[k], cv = cur[k] || "";
      if (!tv || pos[k] === undefined) return;
      var same;
      if (CT_DATES[k]) same = ctParseDate(cv, parseInt(tv.slice(0, 4), 10)) === tv;
      else if (CT_IDENT[k] || k === "seal_no" || k === "bl_no") same = ctNorm(cv) === ctNorm(tv);
      else same = cv.trim().toUpperCase() === tv.trim().toUpperCase();
      if (same) return;
      if (CT_IDENT[k] && cv) return;
      if (CT_FILL_ONLY[k] && cv) return;
      ch[k] = tv; nw[k] = tv;
    });
    var code = ctLookup(nw.status || ""), floor = ctFloor(nw, code);
    if (st.indexOf("مستلم") >= 0 && (!code || ctProg(code) >= 0) && ctProg("Delivered to Warehouse") > ctProg(floor)) floor = "Delivered to Warehouse";
    if (floor && floor !== code && ctProg(floor) > ctProg(code) && pos.status !== undefined) ch.status = ctLabel(floor);
    Object.keys(ch).forEach(function (k) { updates.push({row: r + 1, key: k, value: ch[k]}); });
  }
  return {updates: updates, pos: pos};
}

function ctSyncAll() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;
  try {
    var files = DriveApp.getFilesByName(CT_NAME);
    if (!files.hasNext()) { Logger.log("مش لاقي جوجل شيت اسمه " + CT_NAME); return; }
    var trk = SpreadsheetApp.open(files.next()).getSheets()[0].getDataRange().getDisplayValues();
    var sh = SpreadsheetApp.getActive().getSheetByName(CT_CHINA_TAB);
    if (!sh) { Logger.log("مفيش تاب اسمه " + CT_CHINA_TAB); return; }
    var plan = ctPlan(sh.getDataRange().getDisplayValues(), ctReadTracking(trk, new Date().getFullYear()));
    plan.updates.forEach(function (u) { sh.getRange(u.row, plan.pos[u.key] + 1).setValue(u.value); });
    SpreadsheetApp.flush();
    Logger.log("اتحدّث " + plan.updates.length + " خلية");
  } finally { lock.releaseLock(); }
}
function ctOnEdit(e) {                                              // installable trigger: لما تكتب مرجع في China
  try {
    var sh = e.range.getSheet();
    if (sh.getName() !== CT_CHINA_TAB || e.range.getRow() < 4) return;
    var key = String(sh.getRange(2, e.range.getColumn()).getValue()).trim();
    if (["container_no", "bl_no", "shipment_id"].indexOf(key) >= 0) ctSyncAll();
  } catch (err) { Logger.log(err); }
}
function ctSetup() {                                                // شغّلها مرة واحدة بس
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === "ctSyncAll" || f === "ctOnEdit") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("ctSyncAll").timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger("ctOnEdit").forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  ctSyncAll();
}
"""


_CHINA_GROUP_STYLE = {   # (عنوان المجموعة, لون الشريط, لون العناوين)
    "ORDER": ("📋 الطلب / الشحنة — يتكرر على كل منتج في نفس الطلب", "1F4E79", "DDEBF7"),
    "CONTAINER": ("🚢 الحاوية والتتبع — يتكرر على كل منتج في نفس الحاوية", "C55A11", "FCE4D6"),
    "LINE": ("📦 المنتج / الدفعة — صف لكل منتج", "2E7D32", "E2EFDA"),
    "STAMP": ("🕓 التحديث", "595959", "EDEDED"),
    "KEY": ("🔑 الباركود · الصورة · الشحنة · الحاوية · الحالة", "6A1B9A", "E4D4F0"),
    "AUTO": ("⚙️ محسوب تلقائي — لا تكتب هنا", "7F7F7F", "D9D9D9"),
}
CHINA_FLAT_MAX_ROWS = 3000




def _china_retry(fn, *args, **kwargs):
    """بيعيد محاولة الكتابة في جوجل شيت لو اتعدّت حدّ الطلبات (429) — عشان تحديث شيت China مايقفش في النص."""
    for attempt in range(7):
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            if attempt < 6 and ("429" in str(e) or "quota" in str(e).lower() or "RESOURCE_EXHAUSTED" in str(e)):
                time.sleep(10 * (attempt + 1))
            else:
                raise


def _china_write_image_formulas(ws):
    """معادلات الصف الواحد (مش ARRAYFORMULA): صورة المنتج + «المطلوب دلوقتي» (رابط للخلية)."""
    for key, mk in (("product_image", lambda r: _china_image_formula(r)), ("next_action", lambda r: _china_next_action_row(r, ws.id))):
        ic = _china_col_letter(CHINA_FLAT_KEYS.index(key) + 1)
        for r0 in range(4, CHINA_FLAT_MAX_ROWS + 1, 250):          # على دفعات عشان حجم الطلب يفضل صغير
            r1 = min(r0 + 249, CHINA_FLAT_MAX_ROWS)
            _china_retry(ws.update, values=[[mk(r)] for r in range(r0, r1 + 1)], range_name=f"{ic}{r0}:{ic}{r1}", value_input_option="USER_ENTERED")

def _china_format_flat(ws):
    """بيكتب العناوين والمعادلات وينسّق الشيت (ألوان/تثبيت/قوايم/تلوين ذكي) — مرة واحدة عند الإنشاء الأول."""
    sid, N = ws.id, len(CHINA_FLAT_KEYS)
    band, labels = [""] * N, []
    for i, (g, k, lab, w, kind) in enumerate(CHINA_FLAT_COLS):
        if i == 0 or CHINA_FLAT_COLS[i - 1][0] != g:
            band[i] = _CHINA_GROUP_STYLE[g][0]
        labels.append(lab)
    auto = _china_auto_formulas()
    _china_retry(ws.update, values=[band, CHINA_FLAT_KEYS, labels], range_name="A1", value_input_option="USER_ENTERED")
    _china_write_image_formulas(ws)
    for k, f in auto.items():
        _china_retry(ws.update, values=[[f]], range_name=f"{_china_col_letter(CHINA_FLAT_KEYS.index(k) + 1)}4", value_input_option="USER_ENTERED")
    rg = lambda r0, r1, c0, c1: {"sheetId": sid, "startRowIndex": r0, "endRowIndex": r1, "startColumnIndex": c0, "endColumnIndex": c1}
    M = CHINA_FLAT_MAX_ROWS
    reqs = []
    i = 0
    while i < N:
        g = CHINA_FLAT_COLS[i][0]
        j = i
        while j + 1 < N and CHINA_FLAT_COLS[j + 1][0] == g:
            j += 1
        st_ = _CHINA_GROUP_STYLE[g]
        reqs.append({"repeatCell": {"range": rg(0, 1, i, j + 1), "cell": {"userEnteredFormat": {
            "backgroundColor": _china_rgb(st_[1]), "horizontalAlignment": "LEFT", "verticalAlignment": "MIDDLE", "wrapStrategy": "OVERFLOW_CELL",
            "textFormat": {"bold": True, "fontSize": 12, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,wrapStrategy,textFormat)"}})
        reqs.append({"repeatCell": {"range": rg(2, 3, i, j + 1), "cell": {"userEnteredFormat": {
            "backgroundColor": _china_rgb(st_[2]), "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE", "wrapStrategy": "WRAP",
            "textFormat": {"bold": True, "fontSize": 10}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,wrapStrategy,textFormat)"}})
        if g == "AUTO":
            reqs.append({"repeatCell": {"range": rg(3, M, i, j + 1), "cell": {"userEnteredFormat": {
                "backgroundColor": _china_rgb("F2F2F2"), "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE", "wrapStrategy": "WRAP"}},
                "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,wrapStrategy)"}})
        i = j + 1
    reqs.append({"repeatCell": {"range": rg(1, 2, 0, N), "cell": {"userEnteredFormat": {
        "backgroundColor": _china_rgb("F7F7F7"), "horizontalAlignment": "CENTER",
        "textFormat": {"fontSize": 8, "foregroundColor": _china_rgb("7F7F7F")}}},
        "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)"}})
    for i, (g, k, lab, w, kind) in enumerate(CHINA_FLAT_COLS):
        reqs.append({"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": i, "endIndex": i + 1},
                                                   "properties": {"pixelSize": int(w * 7 + 12)}, "fields": "pixelSize"}})
        if kind in ("t", "d") and g != "AUTO":
            nf = {"type": "TEXT"} if kind == "t" else {"type": "DATE", "pattern": "yyyy-mm-dd"}
            reqs.append({"repeatCell": {"range": rg(3, M, i, i + 1), "cell": {"userEnteredFormat": {"numberFormat": nf, "horizontalAlignment": "CENTER"}},
                                        "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"}})
        elif kind == "n":
            reqs.append({"repeatCell": {"range": rg(3, M, i, i + 1), "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                                        "fields": "userEnteredFormat.horizontalAlignment"}})
        if k in CHINA_FLAT_LISTS:
            reqs.append({"setDataValidation": {"range": rg(3, M, i, i + 1), "rule": {"condition": {
                "type": "ONE_OF_LIST", "values": [{"userEnteredValue": x} for x in CHINA_FLAT_LISTS[k]]}, "showCustomUi": True, "strict": False}}})
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
                                               "properties": {"pixelSize": 30}, "fields": "pixelSize"}})
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 2, "endIndex": 3},
                                               "properties": {"pixelSize": 44}, "fields": "pixelSize"}})
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 3, "endIndex": M},
                                               "properties": {"pixelSize": 64}, "fields": "pixelSize"}})
    pi = CHINA_FLAT_KEYS.index("product_image")
    reqs.append({"repeatCell": {"range": rg(3, M, pi, pi + 1), "cell": {"userEnteredFormat": {
        "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE", "backgroundColor": _china_rgb("F2F2F2")}},
        "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment,backgroundColor)"}})
    reqs.append({"updateSheetProperties": {"properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 3, "frozenColumnCount": 6}},
                                           "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}})
    reqs.append({"setBasicFilter": {"filter": {"range": rg(2, M, 0, N)}}})
    c = {k: _china_col_letter(i + 1) for i, k in enumerate(CHINA_FLAT_KEYS)}
    def cf(key, formula, bg, fg=None, bold=False):
        fmt = {"backgroundColor": _china_rgb(bg)}
        if fg or bold:
            fmt["textFormat"] = {"bold": bold, **({"foregroundColor": _china_rgb(fg)} if fg else {})}
        i = CHINA_FLAT_KEYS.index(key)
        reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {"ranges": [rg(3, M, i, i + 1)], "booleanRule": {
            "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": formula}]}, "format": fmt}}}})
    E, A, S_ = c["eta_port"], c["port_arrival_actual"], c["status"]
    cf("eta_port", f'=AND(${E}4<>"",${E}4<TODAY(),${A}4="",NOT(REGEXMATCH(${S_}4,"ملغي|وصلت المستودع")))', "F8CBAD", "9C0006", True)
    cf("eta_port", f'=AND(${E}4<>"",${E}4>=TODAY(),${E}4-TODAY()<=7,${A}4="")', "FFF2CC")
    DD = c["delay_days"]
    cf("status", f'=AND(N(${DD}4)>0,NOT(REGEXMATCH(${S_}4,"ملغي|وصلت المستودع")))', "F8CBAD", "9C0006", True)
    cf("status", f'=REGEXMATCH(${S_}4,"ملغي")', "D9D9D9", "7F7F7F")
    cf("status", f'=REGEXMATCH(${S_}4,"وصلت المستودع")', "C6EFCE")
    cf("status", f'=REGEXMATCH(${S_}4,"وصلت الميناء|التخليص")', "F8CBAD")
    cf("status", f'=REGEXMATCH(${S_}4,"في البحر")', "DDEBF7")
    cf("status", f'=REGEXMATCH(${S_}4,"تم التحميل|متجهزة")', "FCE4D6")
    cf("status", f'=REGEXMATCH(${S_}4,"قيد الإنتاج|تأكيد المصنع|اكتمل|جاهز")', "FFF2CC")
    cf("status", f'=REGEXMATCH(${S_}4,"موقوف")', "FFEB9C", "9C5700", True)
    cf("next_action", f'=LEFT(${c["next_action"]}4,1)="⚠"', "FFEB9C", "9C5700", True)
    cf("row_state", f'=${c["row_state"]}4="MISSING"', "F8CBAD", "9C0006", True)
    cf("row_state", f'=${c["row_state"]}4="OK"', "C6EFCE")
    _china_retry(ss.batch_update, {"requests": reqs})


@st.cache_resource
def _china_checked_flags():
    """علم على مستوى السيرفر كله: فحص ترتيب الأعمدة والمعادلات بيتعمل مرة واحدة بس، مش مع كل جلسة."""
    return {}

def get_china_flat_ws():
    """شيت China الواحد — بيتنشأ تلقائي أول مرة (بالتنسيق والمعادلات)، ولو كان فيه بيانات قديمة في التابات
    الـ3 (ChinaShipments/ChinaLines/ChinaContainers) بتتنقل له مرة واحدة."""
    if "china_flat_ws" in st.session_state:
        return st.session_state["china_flat_ws"]
    ws, created, last = None, False, None
    _done = _china_checked_flags()
    ws = st.session_state.get("_ws_existing", {}).get(CHINA_FLAT_TAB)
    if ws is not None and _done.get("ok"):
        st.session_state["china_flat_ws"] = ws
        return ws
    ws = None
    for attempt in range(5):
        try:
            ws = ss.worksheet(CHINA_FLAT_TAB)
            break
        except gspread.exceptions.WorksheetNotFound:
            try:
                ws = ss.add_worksheet(title=CHINA_FLAT_TAB, rows=str(CHINA_FLAT_MAX_ROWS), cols=str(len(CHINA_FLAT_KEYS)))
                created = True
                break
            except gspread.exceptions.APIError as e:
                last = e
                time.sleep(1 * (2 ** attempt))
        except gspread.exceptions.APIError as e:
            last = e
            time.sleep(1 * (2 ** attempt))
    if ws is None:
        raise last
    if created:
        try:
            _china_format_flat(ws)
        except Exception as e:
            st.warning(f"⚠️ اتعمل شيت China لكن التنسيق ما اكتملش: {e}")
        moved = 0
        try:
            old = {}
            for tab in ("ChinaShipments", "ChinaLines", "ChinaContainers"):
                try:
                    old[tab] = _china_rows_from_raw(ss.worksheet(tab).get_all_values())
                except gspread.exceptions.WorksheetNotFound:
                    old[tab] = []
            if any(old.values()):
                flat = china_tables_to_flat(old["ChinaShipments"], old["ChinaContainers"], old["ChinaLines"])
                _china_write_flat(ws, flat)
                moved = len(flat)
        except Exception as e:
            st.warning(f"⚠️ مقدرتش أنقل البيانات القديمة: {e}")
        st.session_state["china_flash"] = ("✅ تم إنشاء شيت China في Google Sheet" + (f" ونقل {moved} صف من البيانات القديمة له." if moved else ".")
                                           + " الموظف تقدر تعدّل فيه مباشرة.")
    elif _china_layout_outdated(ws):
        try:
            old_rows = _china_read_flat(ws.get_all_values())
            try:                                   # بقايا محاولة تحديث سابقة وقفت في النص — الشيت الأصلي لسه سليم فنمسحها ونبدأ من جديد
                ss.del_worksheet(ss.worksheet("China_new_tmp"))
            except gspread.exceptions.WorksheetNotFound:
                pass
            tmp = ss.add_worksheet(title="China_new_tmp", rows=str(CHINA_FLAT_MAX_ROWS), cols=str(len(CHINA_FLAT_KEYS)))
            _china_format_flat(tmp)
            if old_rows:
                _china_write_flat(tmp, old_rows)
            _bk = f"China_backup_{time.strftime('%Y%m%d_%H%M')}"
            ws.update_title(_bk)
            try:
                tmp.update_title(CHINA_FLAT_TAB)
            except Exception:
                ws.update_title(CHINA_FLAT_TAB)    # لو التسمية فشلت نرجّع الأصلي مكانه عشان ما نفقدش شيت China
                raise
            ws = tmp
            st.session_state["china_flash"] = (f"✅ اتحدّث ترتيب أعمدة شيت China (رقم الشحنة ورقم الحاوية بعد الصورة وبقوا في الجزء المتثبت) ونقل {len(old_rows)} صف. "
                                               "النسخة القديمة اتحفظت كتاب باسم China_backup_… ومسموح تمسحها بعد ما تتأكد.")
        except Exception as e:
            st.warning(f"⚠️ تحديث ترتيب أعمدة شيت China فشل: {e}")
    elif not getattr(ws, "frozen_row_count", 0):
        try:
            _china_format_flat(ws)
            st.session_state["china_flash"] = "✅ اتكمّل تنسيق شيت China."
        except Exception as e:
            st.warning(f"⚠️ تنسيق شيت China ما اكتملش: {e}")
    try:   # لو معادلة الصورة أو «المطلوب دلوقتي» مش موجودة (أو الرابط بتاعها لتاب قديم) بنكتبها تاني
        ni = CHINA_FLAT_KEYS.index("next_action") + 1
        cur = ws.cell(4, ni, value_render_option="FORMULA").value or ""
        if f"gid={ws.id}&" not in str(cur):
            _china_write_image_formulas(ws)
    except Exception as e:
        st.warning(f"⚠️ مقدرتش أتأكد من معادلات الصورة/المطلوب دلوقتي: {e}")
    else:
        _done["ok"] = True
    st.session_state["china_flat_ws"] = ws
    if not _done.get("seed_checked"):
        _done["seed_checked"] = True
        china_seed_once(ws)
    return ws


CHINA_SEED_FILE = "China_initial.xlsx"       # الملف اللي جنب App.py — بيتستورد مرة واحدة بس

def china_seed_once(ws):
    """أول مرة بس: لو شيت China فاضي ولسه ما اتعملش seed، بيستورد ملف China_initial.xlsx (لو موجود جنب App.py).
    بعد كده بيتسجّل علم china_seeded في شيت Settings فمايتكررش أبدًا (حتى لو الشيت اتفضّى بعد كده).
    لو الشيت فيه بيانات أصلًا بيتسجّل العلم من غير ما يلمس حاجة."""
    import os
    try:
        if load_settings().get("china_seeded"):
            return
        if _china_read_flat(get_cached(ws, force=True)):
            save_setting("china_seeded", "existing-data")
            return
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), CHINA_SEED_FILE)
        if not os.path.exists(path):
            return                              # الملف لسه مترفعش على GitHub — هنحاول تاني مع أول تشغيل جاي
        with open(path, "rb") as f:
            parsed = _china_parse_workbook(f.read())
        r = _china_do_import(parsed, True, f"{CHINA_SEED_FILE} (أول مرة)")
        save_setting("china_seeded", now_str())
        st.session_state["china_flash"] = (f"✅ أول مرة: اتضاف محتوى {CHINA_SEED_FILE} لشيت China — {r['SHIPMENTS']} شحنة، "
                                           f"{r['SKU_LINES']} منتج، {r['CONTAINERS']} حاوية ({r['rows']} صف).")
    except Exception as e:
        st.warning(f"⚠️ استيراد {CHINA_SEED_FILE} أول مرة فشل: {e}")



def _china_write_flat(ws, flat_rows):
    """بيكتب أعمدة الإدخال بس من صف 4 (أعمدة المعادلات والصورة بتفضل زي ما هي)."""
    flat_rows = china_flat_prepare(flat_rows)
    n = len(flat_rows)
    if ws.row_count < n + CHINA_FLAT_HDR_ROWS + 50:
        ws.add_rows(n + CHINA_FLAT_HDR_ROWS + 200 - ws.row_count)
    segs = _china_input_segments()
    _china_retry(ws.batch_clear, [f"{_china_col_letter(s[0][0] + 1)}4:{_china_col_letter(s[-1][0] + 1)}{max(ws.row_count, 4)}" for s in segs])
    for i in range(0, n, 800):
        part = flat_rows[i:i + 800]
        for s in segs:
            chunk = [[_china_cell(r.get(k, "")) for _, k in s] for r in part]
            _china_retry(ws.update, values=chunk, range_name=f"{_china_col_letter(s[0][0] + 1)}{4 + i}", value_input_option="USER_ENTERED")
    clear_cache(ws)




@st.cache_resource
def _ct_state():
    """حالة مزامنة Container tracking على مستوى السيرفر كله (آخر تشغيل + نتيجته + الشيت المفتوح)."""
    return {"ts": 0, "ws": None, "last": None}


def china_sync_tracking(force=False):
    """بيقرأ شيت Container tracking (جوجل شيت منفصل بنفس الـ service account) وبيملا/يحدّث شيت China من أي مرجع مكتوب
    (رقم حاوية / بوليصة / شحنة STxxxxx). بيشتغل تلقائي كل 5 دقايق على الأكتر أو بالزرار — وعمره ما بيرمي خطأ للتطبيق."""
    s = _ct_state()
    if not force and time.time() - s["ts"] < 300:
        return s["last"]
    s["ts"] = time.time()
    res = {"time": time.strftime("%Y-%m-%d %H:%M"), "rows": 0, "cells": 0, "notes": [], "error": ""}
    try:
        if s["ws"] is None:
            s["ws"] = client.open(CT_SHEET_NAME).sheet1
        recs = ct_read_tracking(safe_get_all_values(s["ws"]))
        if not recs:
            res["error"] = "ما لقيتش في Container tracking صفوف فيها رقم حاوية أو شحنة (أو عناوين الأعمدة مش مفهومة)."
        else:
            ws = get_china_flat_ws()
            plan = ct_sync_plan(safe_get_all_values(ws), recs)
            res["notes"], res["rows"], res["cells"] = plan["notes"], plan["rows"], len(plan["updates"])
            pos = plan["pos"]      # مكان العمود من صف العناوين الفعلي في الشيت (مش من الترتيب المتوقع) — آمن حتى لو الشيت لسه بالترتيب القديم
            ups = [{"range": f"{_china_col_letter(pos[k] + 1)}{r}", "values": [[_china_cell(v)]]} for r, k, v in plan["updates"] if k in pos]
            for i in range(0, len(ups), 400):
                ws.batch_update(ups[i:i + 400], value_input_option="USER_ENTERED")
            if ups:
                clear_cache(ws)
    except gspread.exceptions.SpreadsheetNotFound:
        s["ws"] = None
        try:
            mail = st.secrets["gcp_service_account"].get("client_email", "")
        except Exception:
            mail = ""
        res["error"] = f"مش لاقي جوجل شيت اسمه «{CT_SHEET_NAME}» — اعمله Share (Editor أو Viewer) للإيميل: {mail}"
    except Exception as e:
        s["ws"] = None
        res["error"] = f"{type(e).__name__}: {e}"
    s["last"] = res
    return res


def _china_flat_raw(force=False):
    """الخام من الشيت — بيتجدّد تلقائي كل دقيقتين (عشان تعديلات الموظف تبان) أو بالزرار."""
    ws = get_china_flat_ws()
    ts = st.session_state.get("china_flat_ts", 0)
    if force or time.time() - ts > 300:
        china_sync_tracking()                      # الأول نملا China من Container tracking، وبعدين نقرأ
        raw = get_cached(ws, force=True)
        st.session_state["china_flat_ts"] = time.time()
        return raw
    return get_cached(ws)

def _china_current_tables(*args, **kwargs):
    return calc._china_current_tables(_ENV, *args, **kwargs)





# ── قراءة ملف الإكسيل: بيقبل الملف القديم (3 صفحات) أو شيت China الواحد | Parse either format ──


@st.cache_data(show_spinner=False)
def _china_parse_workbook(file_bytes):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    res = {"schema": "", "ships": None, "lines": None, "conts": None}
    # 1) شيت China الواحد: صفحة فيها مفاتيح الطلب والحاوية والمنتج مع بعض (الملف القديم بيفصلهم على 3 صفحات)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(min_row=1, max_row=8, values_only=True))
        hit = None
        for i, r in enumerate(rows):
            low = {str(c).strip() for c in r if c is not None}
            if {"container_no", "line_id", "internal_order_no", "supplier_name", "barcode"} <= low and ("status" in low or "container_status" in low):
                hit = i
                break
        if hit is None:
            continue
        keys = [str(c).strip() if c is not None else "" for c in rows[hit]]
        flat = []
        for r in ws.iter_rows(min_row=hit + 2, values_only=True):
            d = {k: _china_clean(r[j], k) for j, k in enumerate(keys) if k in CHINA_FLAT_INPUT_KEYS + CHINA_STATUS_LEGACY and j < len(r)}
            if any(d.values()) and not _china_is_label_row(d):
                flat.append(d)
        ships, conts, lines, _ = china_flat_to_tables(flat)
        res.update({"schema": "SMARTSTORE_CHINA_V5.1", "ships": ships, "lines": lines, "conts": conts})
        return res
    # 2) الملف القديم: SHIPMENTS / SKU_LINES / CONTAINERS
    by_name = {n.strip().upper(): n for n in wb.sheetnames}
    for sheet_key, out_key, cols in (("SHIPMENTS", "ships", CHINA_SHIP_COLS), ("SKU_LINES", "lines", CHINA_LINE_COLS),
                                     ("CONTAINERS", "conts", CHINA_CONT_COLS)):
        real = by_name.get(sheet_key)
        if not real:
            continue
        ws = wb[real]
        if not res["schema"]:
            for r in ws.iter_rows(min_row=1, max_row=6, values_only=True):
                if r and str(r[0] or "").strip() == "schema_version" and len(r) > 1 and r[1]:
                    res["schema"] = str(r[1]).strip()
                    break
        rows = _china_parse_sheet(ws, cols)
        if rows is None:
            continue
        if out_key == "ships":
            rows = [r for r in rows if r.get("shipment_id") or r.get("internal_order_no")]
        elif out_key == "lines":
            rows = [r for r in rows if r.get("barcode") or r.get("sku") or r.get("model_no")
                    or r.get("product_name_en") or r.get("product_name_ar") or r.get("line_id")]
        else:
            rows = [r for r in rows if r.get("container_no")]
        res[out_key] = rows
    for r in res["ships"] or []:
        r["record_key"] = china_ship_record_key(r)
    for r in res["conts"] or []:
        r["record_key"] = "CTR|" + r.get("container_no", "")
    return res

def _china_diff(*args, **kwargs):
    return calc._china_diff(_ENV, *args, **kwargs)

def _china_do_import(parsed, replace, label=""):
    ws = get_china_flat_ws()
    cur_ships, cur_conts, cur_lines, _ = _china_current_tables(force=True)
    pieces = {}
    for pk, cur, keyfn in (("ships", cur_ships, _china_ship_key), ("lines", cur_lines, _china_line_key), ("conts", cur_conts, _china_cont_key)):
        new_rows = parsed.get(pk)
        if new_rows is None:          # الصفحة دي مش في الملف — سيبها زي ما هي
            pieces[pk] = cur
        elif replace:
            pieces[pk] = new_rows
        else:
            pieces[pk] = _china_merge(cur, new_rows, keyfn)
    flat = china_tables_to_flat(pieces["ships"], pieces["conts"], pieces["lines"])
    _china_write_flat(ws, flat)
    st.session_state["china_flat_ts"] = 0
    try:
        save_setting("china_last_import", f"{now_str()} | {label} | {'استبدال' if replace else 'دمج'}")
    except Exception:
        pass
    return {"SHIPMENTS": len(pieces["ships"]), "SKU_LINES": len(pieces["lines"]), "CONTAINERS": len(pieces["conts"]), "rows": len(flat)}

def _china_delete_all():
    _china_write_flat(get_china_flat_ws(), [])
    st.session_state["china_flat_ts"] = 0

def _china_delete_order(order_no):
    """يحذف طلب (internal_order_no) كامل: شحناته وأسطره، والحاويات اللي كانت مربوطة بأسطره بس."""
    ships, conts, lines, _ = _china_current_tables(force=True)
    gone = {l.get("container_no", "") for l in lines if l.get("internal_order_no") == order_no} - {""}
    lines2 = [l for l in lines if l.get("internal_order_no") != order_no]
    still = {l.get("container_no", "") for l in lines2}
    conts2 = [c for c in conts if not (c.get("container_no", "") in gone and c.get("container_no", "") not in still)]
    ships2 = [s for s in ships if s.get("internal_order_no") != order_no]
    _china_write_flat(get_china_flat_ws(), china_tables_to_flat(ships2, conts2, lines2))
    st.session_state["china_flat_ts"] = 0


def get_china_images_map():
    """باركود موحّد -> رابط صورة (شيت ChinaImages) | normalised barcode -> image URL."""
    def _b(raw):
        m = {}
        for r in (raw or [])[1:]:
            if len(r) > 1 and str(r[0]).strip() and str(r[1]).strip().startswith("http"):
                m[_china_norm_code(r[0])] = str(r[1]).strip()
        return m
    return _memo_from_sheet(china_img_sheet, "china_img_map", _b)

def get_china_bundle():
    raw = _china_flat_raw()
    if (st.session_state.get("china_bundle_src") is raw and "china_bundle" in st.session_state
            and st.session_state.get("china_bundle_img_src") is get_cached(china_img_sheet)):
        return st.session_state["china_bundle"]
    ships, conts_list, lines, conflicts = china_flat_to_tables(_china_read_flat(raw))
    conts = {c["container_no"]: c for c in conts_list if c.get("container_no")}
    ship_by_pair, ship_by_no = {}, {}
    for s in ships:
        ship_by_pair.setdefault((s.get("internal_order_no", ""), s.get("po_no", "")), s)
        ship_by_no.setdefault(s.get("internal_order_no", ""), s)
    enriched = [_china_enrich(l, conts, ship_by_pair, ship_by_no) for l in lines]
    _cimg = get_china_images_map()
    for e in enriched:
        if not e.get("image_url") and e.get("_norm") in _cimg:
            e["image_url"] = _cimg[e["_norm"]]
    by_code, by_noon_sku = {}, {}
    for e in enriched:
        if e["_norm"]:
            by_code.setdefault(e["_norm"], []).append(e)
        ns = e.get("noon_sku", "").strip().upper()
        if ns:
            by_noon_sku.setdefault(ns, []).append(e)
    bundle = {"ships": ships, "conts": conts_list, "conts_map": conts, "lines": enriched,
              "by_code": by_code, "by_noon_sku": by_noon_sku, "conflicts": conflicts}
    st.session_state["china_bundle"] = bundle
    st.session_state["china_bundle_src"] = raw
    st.session_state["china_bundle_img_src"] = get_cached(china_img_sheet)
    return bundle

# ── البيانات المجمّعة والمربوطة (متكاشة) | The joined data bundle (cached) ──


def _china_code_to_skus(*args, **kwargs):
    return calc._china_code_to_skus(_ENV, *args, **kwargs)

def _china_wh_lookup():
    """مفتاح (SKU نون أو ASIN أمازون) -> صف المستودع على مستوى الكود (سحب مجمّع نون + أمازون)."""
    out = dict(st.session_state.get("warehouse_stockout_asin_lookup", {}))
    out.update(st.session_state.get("warehouse_stockout_sku_lookup", {}))
    return out

def china_lines_for_sku(*args, **kwargs):
    return calc.china_lines_for_sku(_ENV, *args, **kwargs)

# ── 📅 تاريخ الإتاحة للبيع + فجوة الوصول | Available-for-sale date & arrival gap ──
# من بيانات الإكسيل بس، من غير أي افتراضات أو تقديرات: متاح للبيع = الوصول المتوقع لمخزن الرياض (eta_destination = وصول
# المستودع بالرياض) اللي بتكتبه موظفة الصين. لو مش مكتوب، الدفعة بتظهر "بدون تاريخ وصول للوجهة" مع ETA
# الميناء كمعلومة بس، ومبتدخلش في حساب الفجوة. الدوال دي مفيهاش اعتماد على واجهة العرض — فتتنقل لاحقًا
# لقاعدة بيانات/API من غير تعديل | Data-only, no assumptions. Pure calculations, no UI dependency.




def china_badge_html(lines):
    inc = [l for l in lines if l["_status"] == "incoming"]
    if not inc:
        return ""
    qty = sum(l["_qty"] for l in inc)
    html = f'<span class="tacweed-badge" style="background:#7c2d12;color:#fed7aa;">🚢 من الصين: {qty:,} قطعة</span>'
    etas = [d for d in (_china_date(l["_eta"]) for l in inc) if d]
    if etas:
        html += f' <span class="tacweed-badge" style="background:#1e3a5f;color:#bfdbfe;">📅 أقرب وصول: {min(etas)}</span>'
    return html

def _china_lines_table_html(lines):
    from html import escape as _esc
    head = ["shipment_id", "container_no", "الكمية", "المرحلة", "خروج الشحنة من الصين", "الوصول المتوقع للميناء", "وصل الميناء فعليًا",
            "الوصول المتوقع لمخزن الرياض", "وصل المستودع", "متاح للبيع (وصول مخزن الرياض)", "باقي كم يوم للوصول للميناء", "أيام التأخير", "آخر تحديث", "الطلب | PO"]
    th = "".join(f'<th style="padding:3px 8px;border-bottom:1px solid #94a3b8;text-align:right;">{h}</th>' for h in head)
    body = ""
    _today = datetime.now().date()
    for l in lines:
        d = l["_d"]
        _av, _basis = _china_avail(l)
        av_txt = str(_av) if _av else ("وصل" if l["_status"] == "received" else "—")
        dte, dly = _china_eta_stats(d["eta_port"], d["port_arr"], l["_status"] == "cancelled", _today)
        upd = f'{l["_upd"]} ({l["_upd_by"]})' if l["_upd"] and l["_upd_by"] else (l["_upd"] or "—")
        cells = [l["_shipment_id"] or "—", l.get("container_no") or "—", f'{l["_qty"]:,}',
                 l.get("_stage_eff") or l.get("stock_stage") or "—",
                 d["etd"] or "—", d["eta_port"] or "—", d["port_arr"] or "—", (("≈ " + d["eta_dest"]) if d["eta_dest"] and d.get("eta_dest_est") else (d["eta_dest"] or "—")), d["wh_arr"] or "—", av_txt,
                 "—" if dte is None else dte, "—" if dly is None else dly, upd,
                 f'{l.get("internal_order_no") or "—"} | {l.get("po_no") or "—"}']
        body += "<tr>" + "".join(f'<td style="padding:3px 8px;">{_esc(str(c))}</td>' for c in cells) + "</tr>"
    return f'<div style="overflow-x:auto;"><table style="font-size:12px;border-collapse:collapse;"><tr>{th}</tr>{body}</table></div>'

# ── 🎯 محرك القرار: مخزون المستودع + موعد الطلب من الصين | Decision engine ──
# كله مبني على: المتوفر بالمستودع ومعدل سحب الكود (نفس أساس شارة "آمن — نفاد خلال N يوم") + تواريخ ملف الصين.
# الإعدادات (مدة التوريد وغيرها) بتتخزن في شيت Settings وتتغير من تاب شحنات الصين.
CHINA_CFG_DEFAULTS = {"china_lead_time_days": 60, "china_reorder_safety_days": 14, "china_cover_days": 60,
                      "china_near_days": 7, "china_stale_days": 3}
CHINA_CFG_LABELS = {
    "china_lead_time_days": "مدة توريد الصين (من يوم الطلب لحد التوفر في الرياض) — يوم",
    "china_reorder_safety_days": "نبّهني قبل موعد الطلب بـ — يوم",
    "china_cover_days": "الكمية المقترحة تغطي — يوم مبيعات",
    "china_near_days": "اعتبر الوصول «قريب» لو خلال — يوم",
    "china_stale_days": "نبّهني لو ملف الصين ما اتحدّثش من — يوم",
}

def _china_cfg():
    try:
        s = load_settings()
    except Exception:
        s = {}
    out = {}
    for k, dv in CHINA_CFG_DEFAULTS.items():
        try:
            v = str(s.get(k, "")).strip()
            out[k] = max(int(float(v)), 0) if v else dv
        except Exception:
            out[k] = dv
    return out

def _china_stock_rate(*args, **kwargs):
    return calc._china_stock_rate(_ENV, *args, **kwargs)


def _china_decision_boxes(dec, today):
    """صناديق HTML بالعربي: الحالة + المطلوب + (وصول قريب/متأخر) لو فيه."""
    F = lambda d: _china_fmt_date(d, today)
    st_, sug = dec["state"], f"{dec.get('suggested', 0):,}"
    box = lambda bg, fg, t: f"<div class='cn-box' style='background:{bg};color:{fg};'>{t}</div>"
    out = ""
    if st_ == "gap":
        out += box("#7f1d1d", "#fecaca",
                   f"🔴 <b>فجوة في البيع:</b> المخزون هيخلص يوم {F(dec['gap_from'])} والشحنة مش هتتوفر قبل يوم {F(dec['gap_to'])} — "
                   f"يعني {dec['gap_days']} يوم من غير بضاعة ونخسر مبيعات تقريبًا {dec['lost']:,} قطعة.<br>"
                   f"🧾 كان لازم الشحنة توصل قبل {F(dec['gap_from'])}، يعني كان لازم نطلب من الصين قبل {F(dec['should_order'])} "
                   f"(مدة التوريد {dec['lead']} يوم).<br>"
                   f"🎯 <b>المطلوب:</b> اضغط على الصين تسرّع الشحنة اللي في الطريق، واطلب شحنة جديدة حوالي {sug} قطعة دلوقتي.")
    elif st_ == "order_late":
        out += box("#7f1d1d", "#fecaca",
                   f"🔴 <b>كان لازم نطلب من الصين:</b> المخزون وشحنات الصين الجاية بيغطوا لحد {F(dec['covered'])}، ومدة التوريد "
                   f"{dec['lead']} يوم — يعني كان لازم نطلب الشحنة الجاية قبل {F(dec['order_by'])}.<br>"
                   f"🎯 <b>المطلوب:</b> اطلب دلوقتي حوالي {sug} قطعة.")
    elif st_ == "order_soon":
        out += box("#78350f", "#fde68a",
                   f"🟠 <b>قرب موعد الطلب من الصين:</b> المخزون وشحنات الصين بيغطوا لحد {F(dec['covered'])}، "
                   f"ولازم نطلب الشحنة الجاية قبل {F(dec['order_by'])}.<br>🎯 <b>المطلوب:</b> جهّز طلب حوالي {sug} قطعة.")
    elif st_ == "need_date":
        out += box("#78350f", "#fde68a",
                   f"🟠 <b>مش قادر أحكم:</b> المخزون هيخلص يوم {F(dec['stockout'])} وفيه {dec['undated']:,} قطعة جاية من الصين "
                   f"بس تاريخ وصولها لمخزن الرياض غير مكتوب.<br>🎯 <b>المطلوب:</b> اطلب من موظف التخليص الجمركي يكتب تاريخ الوصول لمخزن الرياض، "
                   f"ولو مفيش شحنة هتغطي اطلب قبل {F(dec['order_by'])}.")
    elif st_ == "safe":
        extra = f" (مش شامل {dec['undated']:,} قطعة تاريخ وصولها غير مكتوب)" if dec.get("undated") else ""
        out += box("#14532d", "#bbf7d0",
                   f"🟢 <b>آمن:</b> المخزون وشحنات الصين بيغطوا لحد {F(dec['covered'])}{extra}. "
                   f"موعد طلب الشحنة الجاية من الصين: قبل {F(dec['order_by'])}.")
    elif st_ == "no_rate":
        out += box("#374151", "#d1d5db", "ℹ️ مفيش سحب مسجل للمنتج ده، فمش هقدر أحسب موعد الطلب أو الفجوة")
    nxt = dec.get("next")
    if nxt:
        n = (nxt[0] - today).days
        if n < 0:
            out += box("#78350f", "#fde68a", f"🟠 <b>شحنة اتأخرت:</b> {nxt[1]:,} قطعة كان متوقع تتوفر في مخزن الرياض يوم "
                       f"{F(nxt[0])} ولسه ماوصلتش — اسأل الصين عن الموعد الجديد.")
        elif n <= dec.get("near", 7):
            out += box("#1e3a8a", "#bfdbfe", f"🔵 <b>شحنة قربت:</b> {nxt[1]:,} قطعة هتتوفر في مخزن الرياض يوم {F(nxt[0])} — جهّز الاستلام.")
    return out

def _china_decision_html(inc, sku_up, row, today):
    """(سطور المخزون, صناديق القرار) لكارت الـSKU."""
    sr = _china_stock_rate(sku_up, row)
    if sr is None:
        return [], "", ""
    stock, rate, src, n_sku = sr
    cfg = _china_cfg()
    dated, undated = china_batches_for_lines(inc)
    dec = china_decide(stock, rate, [(b["date"], b["qty"]) for b in dated], sum(u["qty"] for u in undated), today, cfg)
    dec["near"] = cfg["china_near_days"]
    src_txt = "المتوفر بالمستودع" if src == "wh" else "مخزون نون (مفيش مخزون مستودع مسجل للكود)"
    items = []
    if rate > 0:
        days_left = stock / rate
        _whr = _china_wh_lookup().get(sku_up) or {}
        if _whr.get("asins"):
            who = (f" (نون {_whr.get('draw_noon', 0):g} + أمازون {_whr.get('draw_amz', 0):g} + تريندول {_whr.get('draw_ty', 0):g} — سحب الكود كله من {n_sku} منتج)")
        else:
            who = f" (سحب الكود كله من {n_sku} SKU)" if n_sku > 1 else ""
        if days_left > 730:
            cover_txt = "يكفي أكتر من سنتين"
        else:
            cover_txt = f"يكفي حوالي {days_left:,.0f} يوم (لحد {_china_fmt_date(today + timedelta(days=int(days_left)))})"
        items.append(f"📦 {src_txt} <b>{stock:,.0f}</b> قطعة بمعدل سحب {rate:.1f} قطعة/يوم{who} — {cover_txt}")
    _lbl = {"gap": "🔴 فجوة في البيع", "order_late": "🔴 كان لازم نطلب من الصين", "order_soon": "🟠 قرب موعد الطلب",
            "need_date": "🟠 مش قادر أحكم", "safe": "🟢 آمن", "no_rate": "ℹ️ مفيش سحب مسجل"}.get(dec["state"], "")
    return items, _china_decision_boxes(dec, today), _lbl

# ── 📝 ملخص الصين بالعربي الواضح | Plain-Arabic summary card ──
_CH_PORTS_AR = {"JEDDAH": "جدة", "DAMMAM": "الدمام", "RIYADH": "الرياض"}


def _china_range(dates, today=None):
    ds = sorted({d for d in dates if d})
    if not ds:
        return ""
    if len(ds) == 1:
        return _china_fmt_date(ds[0], today)
    return f"من {_china_fmt_date(ds[0])} إلى {_china_fmt_date(ds[-1])}"

_OWN_CN = "موظف الصين"
_OWN_CU = "موظف التخليص الجمركي"

def _own_tag(owner, qty=None, what=""):
    """وسم صغير: مين يراجع + كام قطعة + على إيه"""
    t = f"👤 مراجعة {owner}"
    if qty:
        t += f": {qty:,} قطعة {what}".rstrip()
    return f" <span style='color:#fbbf24;font-weight:700;'>[{t}]</span>"

def _china_avail_sentence(g, today, owner=_OWN_CU, hide_missing=False):
    dated, undated = {}, 0
    for l in g:
        d = _china_date(l["_d"]["eta_dest"])
        if d:
            q0, e0 = dated.get(d, (0, False))
            dated[d] = (q0 + l["_qty"], e0 or bool(l["_d"].get("eta_dest_est")))
        else:
            undated += l["_qty"]
    est_note = f"≈ = تقديري (تاريخ وصول الميناء الفعلي + {calc.CHINA_RIYADH_LEAD_DAYS} يوم، مش تاريخ مكتوب)"
    parts = [f"{q:,} قطعة يوم {'≈ ' if e else ''}{_china_fmt_date(d, today)}" for d, (q, e) in sorted(dated.items())]
    _tag = (f" <span style='color:#9ca3af;'>({est_note})</span>" if any(e for _, e in dated.values()) else "")
    if hide_missing:   # لسه في خطوة قبل الوصول — طبيعي ماله تاريخ رياض، منذكرش اللي بعدها
        undated = 0
        if not parts:
            return ""
    if parts and undated:
        return ("📅 متوقع تكون متاحة للبيع في مخزن الرياض: " + "، ".join(parts) + _tag
                + f"، و{undated:,} قطعة لسه تاريخ وصولها لمخزن الرياض غير مكتوب في ملف الصين"
                + _own_tag(owner, undated, "تاريخ وصولها ناقص"))
    if parts:
        return "📅 متوقع تكون متاحة للبيع في مخزن الرياض: " + "، ".join(parts) + _tag
    return ("📅 تاريخ وصولها لمخزن الرياض لسه غير مكتوب في ملف الصين"
            + _own_tag(owner, undated, "تاريخ وصولها ناقص"))

def _china_card_parts(inc, today, row=None, sku_up=""):
    """كارت واحد يشرح الكمية القادمة من الصين: كام في المصنع، كام في البحر، كام وصل الميناء، وإمتى كل جزء
    هيبقى متاح للبيع (من تواريخ ملف الصين بس) — وتنبيه الانقطاع لو المخزون هيخلص قبل الوصول."""
    from html import escape as _esc
    groups = {"factory": [], "sea": [], "port": []}
    for l in inc:
        groups["factory" if l["_idx"] < CHINA_SHIPPED_IDX else ("sea" if l["_idx"] < 7 else "port")].append(l)
    q = {k: sum(x["_qty"] for x in v) for k, v in groups.items()}
    total = sum(q.values())
    bits = []
    if q["factory"]:
        bits.append(f"{q['factory']:,} في المصنع")
    if q["sea"]:
        bits.append(f"{q['sea']:,} في البحر")
    if q["port"]:
        bits.append(f"{q['port']:,} وصلت الميناء")
    head = f"🚢 من الصين {total:,} قطعة"
    sub = ("منها " + "، و".join(bits)) if bits else ""
    items = []

    g = groups["factory"]
    if g:
        by = {}
        for l in g:
            by[l["_stage_eff"] or "—"] = by.get(l["_stage_eff"] or "—", 0) + l["_qty"]
        parts = "، ".join((f"{n:,} {_esc(k)}" if len(by) > 1 else _esc(k)) for k, n in by.items())
        ready = _china_range([_china_date(l["_d"]["ready"]) for l in g])
        line = f"🏭 <b>{q['factory']:,} قطعة</b> لسه في المصنع ({parts})"
        if ready:
            line += f" — الإنتاج جاهز {ready}"
        _late_q = sum(l["_qty"] for l in g if _china_date(l["_d"]["ready"]) and _china_date(l["_d"]["ready"]) < today)
        if _late_q:
            line += " — ولسه ماتشحنتش من الصين" + _own_tag(_OWN_CN, _late_q, "فات ميعاد جاهزيتها ولسه ماتشحنتش")
        _av = _china_avail_sentence(g, today, _OWN_CN, hide_missing=True)
        items.append(line + (f"<br><span class='cn-sub'>{_av}</span>" if _av else ""))

    g = groups["sea"]
    if g:
        etd = _china_range([_china_date(l["_d"]["etd"]) for l in g])
        etas = [_china_date(l["_d"]["eta_port"]) for l in g]
        ports = sorted({_CH_PORTS_AR.get((l["_ctr"].get("destination_port") or l["_ship"].get("destination_port") or "").upper(),
                                         l["_ctr"].get("destination_port") or l["_ship"].get("destination_port") or "") for l in g} - {""})
        port_txt = " " + " و".join(_esc(x) for x in ports) if ports else ""
        line = f"🚢 <b>{q['sea']:,} قطعة</b> في البحر"
        if etd:
            line += f" — خرجت من الصين {etd}"
        if any(etas):
            late = max(e for e in etas if e) < today
            line += (f"، وكان متوقع تصل ميناء{port_txt} {_china_range(etas, today)} ولسه ماوصلتش" if late
                     else f"، ومتوقع تصل ميناء{port_txt} {_china_range(etas, today)}")
            if late:
                line += _own_tag(_OWN_CU, q['sea'], "فات ميعاد وصولها للميناء")
        else:
            line += "، وموعد وصولها للميناء غير مكتوب" + _own_tag(_OWN_CU, q['sea'], "موعد وصولها للميناء ناقص")
        _av = _china_avail_sentence(g, today, _OWN_CU, hide_missing=True)
        items.append(line + (f"<br><span class='cn-sub'>{_av}</span>" if _av else ""))

    g = groups["port"]
    if g:
        arr = _china_range([_china_date(l["_d"]["port_arr"]) for l in g])
        n_clr = sum(l["_qty"] for l in g if l["_idx"] >= 9)
        n_prc = sum(l["_qty"] for l in g if l["_idx"] == 8)
        n_wait = sum(l["_qty"] for l in g if l["_idx"] == 7)
        cust = [f"{n:,} {t}" for n, t in ((n_clr, "تم تخليصها"), (n_prc, "جاري تخليصها"), (n_wait, "في انتظار التخليص الجمركي")) if n]
        line = f"⚓ <b>{q['port']:,} قطعة</b> وصلت الميناء" + (f" يوم {arr}" if arr else "")
        if cust:
            line += " — " + "، ".join(cust)
        if n_prc + n_wait:
            line += _own_tag(_OWN_CU, n_prc + n_wait, "باقي لها تخليص")
        items.append(line + "<br><span class='cn-sub'>" + _china_avail_sentence(g, today) + "</span>")

    alert, dec_items, dec_label = "", [], ""
    if row is not None or sku_up:
        dec_items, alert, dec_label = _china_decision_html(inc, sku_up or (row or {}).get("sku_up", ""), row, today)
    return {"title": head, "sub": sub, "items": items, "dec_items": dec_items, "alert": alert, "dec_label": dec_label}

_CN_CSS = ("<style>.cn-sub{color:#93c5fd;font-size:13px;}.cn-box{margin-top:8px;padding:7px 11px;border-radius:7px;"
           "font-weight:700;font-size:13px;}.cn-li{margin:4px 0;}</style>")

def _china_wrap(inner):
    return ("<div dir='rtl' style='background:#1e293b;border:1px solid #334155;border-radius:10px;padding:10px 14px;margin:2px 0;"
            f"color:#e2e8f0;font-size:14px;line-height:1.9;'>{_CN_CSS}{inner}</div>")

def _china_render_card(inc, today, row=None, sku_up="", cols=None):
    """الكارت اتقسم لعناوين قابلة للفتح (كلها مقفولة): «🚢 من الصين N قطعة» بتفتح التفصيل (منها كذا في المصنع وكذا في البحر
    + كل مجموعة وتاريخ إتاحتها)، و«📦 المخزون وموعد الطلب» بتفتح مخزون المستودع والتنبيه — والحالة ظاهرة في العنوان."""
    p = _china_card_parts(inc, today, row, sku_up)
    import contextlib
    _c0 = cols[0] if cols else contextlib.nullcontext()
    _c1 = cols[1] if cols else contextlib.nullcontext()
    with _c0, st.expander(p["title"], expanded=False):
        body = (f"<div style='font-size:15px;font-weight:800;color:#fed7aa;'>{p['sub']}</div>" if p["sub"] else "")
        body += "".join(f"<div class='cn-li'>{i}</div>" for i in p["items"])
        st.markdown(_china_wrap(body), unsafe_allow_html=True)
    if p["dec_items"] or p["alert"]:
        lbl = "📦 المخزون وموعد الطلب" + (f" — {p['dec_label']}" if p["dec_label"] else "")
        with _c1, st.expander(lbl, expanded=False):
            st.markdown(_china_wrap("".join(f"<div class='cn-li'>{i}</div>" for i in p["dec_items"]) + p["alert"]),
                        unsafe_allow_html=True)

def render_china_for_sku(sku_up, row=None):
    """بيتنادى جوه تابات مبيعات FBN وFBB: كمية الصين القادمة + دفعاتها بتاريخ وصول الوجهة (من الإكسيل بس) +
    تنبيه فجوة الانقطاع لو المخزون هيخلص قبل وصول الشحنة (بيستخدم مخزون ومعدل بيع الصف row) + تفاصيل
    الأسطر. أي خطأ بيظهر كتحذير صغير ومبيوقفش تاب المبيعات."""
    try:
        lines = china_lines_for_sku(sku_up)
        if not lines:
            return
        today = datetime.now().date()
        inc = [l for l in lines if l["_status"] == "incoming"]
        _cc = st.columns(3)        # العناوين التلاتة جنب بعض عشان نوفّر مساحة الصفحة
        if inc:
            _china_render_card(inc, today, row, sku_up, cols=_cc)
        recv = sum(l["_qty"] for l in lines if l["_status"] == "received")
        canc = sum(l["_qty"] for l in lines if l["_status"] == "cancelled")
        extra = ""
        if recv:
            extra += f" | وصل المستودع: {recv:,}"
        if canc:
            extra += f" | ملغي: {canc:,}"
        with _cc[2], st.expander(f"🚢 تفاصيل شحنات الصين | China shipments — {len(lines)} سطر{extra}"):
            st.markdown(_china_lines_table_html(lines), unsafe_allow_html=True)
            for l in lines:
                st.markdown(_china_stepper_html(l["_idx"], l["_status"] == "cancelled"), unsafe_allow_html=True)
                st.caption(f'{l["_shipment_id"] or "—"} | {l.get("container_no") or "—"} | {l["_qty"]:,} قطعة')
            st.caption("موعد الطلب والفجوة محسوبين على المتوفر بالمستودع ومعدل سحب الكود (نفس أساس شارة النفاد) + تواريخ ملف الصين.")
    except Exception as e:
        st.caption(f"⚠️ تعذر عرض بيانات الصين لهذا الـ SKU: {e}")

def render_china_sales_status(sku_ups):
    """سطر حالة صغير فوق قائمة مبيعات FBN/FBB: بيوضح هل بيانات الصين اتحمّلت وكام SKU في القائمة
    اتربط بيها — عشان لو مفيش شارات يبان السبب فورًا | One status line above the FBN/FBB list."""
    try:
        b = get_china_bundle()
        if not b["lines"]:
            st.caption("🚢 بيانات شحنات الصين: لا توجد بيانات محمّلة — ارفع الملف من تاب «رفع ملف» | no China data loaded yet")
            return
        n = sum(1 for s in sku_ups if china_lines_for_sku(s))
        no_code = sum(1 for s in sku_ups if not (get_tacweed_map().get(s, "") or get_tacweed_amazon_map().get(s, "")))
        st.caption(f"🚢 بيانات الصين: {len(b['lines'])} سطر محمّل | مربوط بمنتجات هذه القائمة: {n} SKU"
                   + (f" | {no_code} SKU بدون كود تكويد" if no_code else ""))
    except Exception as e:
        st.caption(f"⚠️ تعذر تحميل بيانات الصين: {e}")

# ── مسار الكود من البداية للنهاية | Full journey of one code, start to finish ──
def _china_stepper_html(idx, cancelled=False):
    if cancelled:
        return '<div style="padding:6px 10px;border-radius:8px;background:#7f1d1d;color:#fecaca;font-size:12px;display:inline-block;">🚫 ملغي | Cancelled</div>'
    items = ""
    for i, st_name in enumerate(CHINA_STAGES):
        if i < idx or (i == idx and idx == CHINA_RECEIVED_IDX):
            bg, fg, mark = "#166534", "#dcfce7", "✓"
        elif i == idx:
            bg, fg, mark = "#c2410c", "#ffffff", "●"
        else:
            bg, fg, mark = "#374151", "#9ca3af", str(i + 1)
        items += (f'<div style="flex:1 1 74px;min-width:74px;text-align:center;">'
                  f'<div style="margin:0 auto;width:24px;height:24px;border-radius:50%;background:{bg};color:{fg};'
                  f'font-size:12px;line-height:24px;font-weight:700;">{mark}</div>'
                  f'<div style="font-size:11px;color:{"#e5e7eb" if i <= idx else "#9ca3af"};margin-top:2px;">{CHINA_STAGE_AR[st_name]}</div></div>')
    return f'<div style="display:flex;flex-wrap:wrap;gap:4px;margin:4px 0 8px 0;">{items}</div>'

def _china_kv_html(pairs):
    from html import escape as _esc
    rows = "".join(
        f'<tr><td style="padding:2px 10px;color:#9ca3af;white-space:nowrap;">{_esc(k)}</td>'
        f'<td style="padding:2px 10px;">{_esc(str(v)) if v not in (None, "") else "—"}</td></tr>' for k, v in pairs)
    return f'<table style="font-size:13px;border-collapse:collapse;">{rows}</table>'

def _china_line_journey(l, today):
    """كارت كامل لسطر واحد: شريط المراحل + الطلب + الكميات + التواريخ + الشحنة + الحاوية."""
    s, c = l["_ship"], l["_ctr"]
    st.markdown(_china_stepper_html(l["_idx"], l["_status"] == "cancelled"), unsafe_allow_html=True)
    st.markdown(f"**المرحلة الحالية:** {l['_stage_eff'] or '—'} &nbsp;|&nbsp; **الكمية:** {l['_qty']:,} قطعة "
                f"&nbsp;|&nbsp; **shipment_id:** `{l['_shipment_id'] or '—'}` &nbsp;|&nbsp; **container_no:** `{l.get('container_no') or '—'}`")
    eta_days = ""
    if l["_status"] == "incoming" and l["_eta"]:
        dd = _china_days(l["_eta"], today)
        if dd is not None:
            eta_days = f"بعد {dd} يوم" if dd >= 0 else f"متأخرة {-dd} يوم"
    _st = _china_eta_stats(l["_d"]["eta_port"], l["_d"]["port_arr"], l["_status"] == "cancelled", today)
    _av, _basis = _china_avail(l)
    g1, g2, g3 = st.columns(3)
    with g1:
        st.markdown("**🏭 الطلب والمصنع | Order**")
        st.markdown(_china_kv_html([
            ("رقم الطلب الداخلي", l.get("internal_order_no")), ("PO المصنع", l.get("po_no")),
            ("المورد", s.get("supplier_name")), ("المشتري", s.get("buyer_company")),
            ("Incoterm", s.get("incoterm")), ("منشأ", s.get("origin_location")),
            ("حالة الشحنة", s.get("shipment_status")), ("آخر تحديث", l.get("last_updated_at")),
            ("حدّث بواسطة", l.get("updated_by"))]), unsafe_allow_html=True)
        st.markdown("**📦 الكميات والتكلفة | Quantity & cost**")
        st.markdown(_china_kv_html([
            ("الكراتين", l.get("cartons")), ("قطع/كرتون", l.get("units_per_carton")), ("الإجمالي", f"{l['_qty']:,}"),
            ("CBM", l.get("cbm")), ("وزن إجمالي كغ", l.get("gross_weight_kg")), ("وزن صافي كغ", l.get("net_weight_kg")),
            ("تكلفة الوحدة", f"{l.get('unit_cost')} {l.get('currency')}".strip() if l.get("unit_cost") else ""),
            ("قيمة السطر", f"{l.get('line_amount')} {l.get('currency')}".strip() if l.get("line_amount") else "")]),
            unsafe_allow_html=True)
    with g2:
        st.markdown("**📅 التواريخ | Timeline**")
        st.markdown(_china_kv_html([
            ("جاهزية الإنتاج", l.get("production_ready_date")), ("تاريخ التحميل", l.get("actual_load_date")),
            ("خروج مخطط من الصين", s.get("planned_etd")), ("خروج الشحنة من الصين", c.get("actual_etd") or s.get("actual_etd")),
            ("الوصول المتوقع للميناء", c.get("eta_port") or s.get("eta_port")), ("وصول فعلي للميناء", c.get("port_arrival_actual")),
            ("الوصول المتوقع لمخزن الرياض", s.get("eta_destination")), ("وصول فعلي للمستودع", c.get("warehouse_arrival_actual")),
            ("الموعد المعتمد", f"{l['_eta_label']}: {l['_eta']} {('(' + eta_days + ')') if eta_days else ''}".strip() if l["_eta"] else ""),
            ("باقي كم يوم للوصول للميناء (days_to_eta)", _st[0]), ("أيام التأخير (delay_days)", _st[1]),
            ("متاح للبيع (وصول مخزن الرياض)", str(_av) if _av else ("وصل المستودع" if l["_status"] == "received" else _basis))]),
            unsafe_allow_html=True)
    with g3:
        st.markdown("**🚢 الشحنة والحاوية | Shipment & container**")
        url = c.get("tracking_url") or s.get("tracking_url") or ""
        bl = c.get("bl_no") or s.get("bl_no") or ""
        if url.endswith("=") and bl:
            url += bl
        st.markdown(_china_kv_html([
            ("شركة الشحن", c.get("shipping_line") or s.get("shipping_line")), ("طريقة الشحن", s.get("shipment_mode")),
            ("B/L", bl), ("رقم الحجز", s.get("booking_no")), ("السفينة", c.get("vessel_name")), ("الرحلة", c.get("voyage_no")),
            ("ميناء التحميل", c.get("port_of_loading") or s.get("port_of_loading")),
            ("ميناء الوصول", c.get("destination_port") or s.get("destination_port")),
            ("الوجهة النهائية", c.get("final_destination") or s.get("final_destination")),
            ("نوع الحاوية", c.get("container_type")), ("رقم الختم", c.get("seal_no")),
            ("حالة الحاوية", c.get("container_status")), ("التخليص", c.get("customs_status")),
            ("آخر موقع", c.get("last_known_location")), ("سبب التأخير", c.get("delay_reason")),
            ("آخر تحديث (الحاوية)", c.get("last_updated_at")), ("حدّث بواسطة (الحاوية)", c.get("updated_by"))]), unsafe_allow_html=True)
        if url.startswith("http"):
            st.markdown(f"[🔗 تتبع الشحنة | Track]({url})")
    if l.get("notes") or s.get("notes"):
        st.caption("📝 " + " | ".join(x for x in (l.get("notes"), s.get("notes")) if x))

def render_china_journey(lines, store_skus_fn, today):
    """تاب مسار الكود: اختار كود (باركود) تشوف كل اللي حصله من تأكيد المصنع لحد المستودع."""
    if not lines:
        st.info("لا توجد أسطر | No lines")
        return
    groups = {}
    for l in lines:
        groups.setdefault(l["_norm"] or f"NOBC|{l.get('sku', '')}", []).append(l)
    def _label(k):
        g = groups[k]
        nm = g[0].get("product_name_ar") or g[0].get("product_name_en") or g[0].get("model_no") or g[0].get("sku", "") or g[0].get("line_id", "")
        inc = sum(x["_qty"] for x in g if x["_status"] == "incoming")
        return f"{g[0].get('barcode') or '(بدون باركود)'} — {nm} — قادم {inc:,} / إجمالي {sum(x['_qty'] for x in g):,}"
    keys = sorted(groups, key=lambda k: -sum(x["_qty"] for x in groups[k] if x["_status"] == "incoming"))
    q = st.text_input("🔍 ابحث عن كود/منتج/SKU/شحنة/حاوية | Search", key="china_jq").strip().upper()
    if q:
        def _hit(k):
            for x in groups[k]:
                hay = " ".join([x.get("barcode", ""), x.get("sku", ""), x.get("model_no", ""), x.get("product_name_en", ""),
                                x.get("product_name_ar", ""), x["_shipment_id"], x.get("container_no", ""),
                                x.get("internal_order_no", ""), x.get("po_no", ""), " ".join(store_skus_fn(x))]).upper()
                if q in hay:
                    return True
            return False
        keys = [k for k in keys if _hit(k)]
    if not keys:
        st.info("لا توجد نتائج | No results")
        return
    pick = st.selectbox(f"الكود | Code ({len(keys)})", keys, format_func=_label, key="china_jpick")
    g = groups[pick]
    first = g[0]
    nm = first.get("product_name_ar") or first.get("product_name_en") or first.get("model_no") or first.get("line_id") or ""
    skus = sorted({s for x in g for s in store_skus_fn(x)})
    h1, h2 = st.columns([1, 6])
    with h1:
        show_img(first.get("image_url", ""), 90)
    with h2:
        st.markdown(f"### {nm}")
        st.markdown(f'<span class="tacweed-badge" style="background:#3b0764;color:#e9d5ff;">🏷️ تكويد: {first.get("barcode") or "—"}</span>',
                    unsafe_allow_html=True)
        st.markdown("**SKUs المتجر:** " + (", ".join(f"`{s}`" for s in skus) if skus else "— غير مربوط بالتكويد"))
        st.caption(f"الموديل: {first.get('model_no') or first.get('sku') or '—'} | EN: {first.get('product_name_en') or '—'}")
    tot = sum(x["_qty"] for x in g)
    by = {"factory": 0, "sea": 0, "port": 0, "recv": 0, "canc": 0}
    for x in g:
        if x["_status"] == "cancelled":
            by["canc"] += x["_qty"]
        elif x["_idx"] >= CHINA_RECEIVED_IDX:
            by["recv"] += x["_qty"]
        elif x["_idx"] >= 7:
            by["port"] += x["_qty"]
        elif x["_idx"] >= CHINA_SHIPPED_IDX:
            by["sea"] += x["_qty"]
        else:
            by["factory"] += x["_qty"]
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("إجمالي الكمية | Total", f"{tot:,}")
    m2.metric("في المصنع | Factory", f"{by['factory']:,}")
    m3.metric("في البحر | At sea", f"{by['sea']:,}")
    m4.metric("ميناء/تخليص | Port/Customs", f"{by['port']:,}")
    m5.metric("وصل المستودع | Received", f"{by['recv']:,}")
    if by["canc"]:
        st.caption(f"🚫 ملغي: {by['canc']:,} قطعة")
    st.markdown(f"**عدد الأسطر/الدفعات: {len(g)}** — كل دفعة (طلب/حاوية) ليها مسارها:")
    for n, l in enumerate(sorted(g, key=lambda x: (x["_status"] == "cancelled", x["_idx"], x.get("line_id", ""))), start=1):
        lbl = (f"دفعة {n}: {l['_qty']:,} قطعة — {l['_stage_eff'] or '—'} — {l['_shipment_id'] or l.get('internal_order_no') or '—'}"
               f" — {l.get('container_no') or 'بدون حاوية'}")
        with st.expander(lbl, expanded=(len(g) <= 2)):
            _china_line_journey(l, today)



# ── واجهة الرفع والحذف | Upload / delete UI ──
def render_china_upload(key_prefix, bundle):
    has_data = bool(bundle["lines"] or bundle["ships"] or bundle["conts"])
    ver = st.session_state.get("china_upl_ver", 0)
    with st.expander("📤 رفع / تحديث من ملف إكسيل (اختياري — الإدخال الأساسي من شيت China) | Upload from Excel (optional)", expanded=not has_data):
        st.caption("بيقبل ملف V5.1 القديم (3 صفحات) أو نسخة الشيت الواحد، ويدمجها في شيت China بدون تكرار | Accepts the old 3-sheet file or the single-sheet version")
        upl = st.file_uploader("ارفع ملف الإكسيل (xlsx/xlsm) | Upload the China Excel file",
                               type=["xlsx", "xlsm"], key=f"{key_prefix}_upl_{ver}")
        if upl:
            try:
                parsed = _china_parse_workbook(upl.getvalue())
            except Exception as e:
                parsed = None
                st.error(f"❌ تعذر قراءة الملف | Could not read the file: {e}")
            if parsed is not None:
                if all(parsed[k] is None for k in ("ships", "lines", "conts")):
                    st.error("❌ مفيش صفحات SHIPMENTS / SKU_LINES / CONTAINERS في الملف ده | None of the expected sheets found")
                else:
                    n_s = len(parsed["ships"]) if parsed["ships"] is not None else "—"
                    n_l = len(parsed["lines"]) if parsed["lines"] is not None else "—"
                    n_c = len(parsed["conts"]) if parsed["conts"] is not None else "—"
                    st.info(f"📊 SHIPMENTS: {n_s} | SKU_LINES: {n_l} | CONTAINERS: {n_c}"
                            + (f" | schema: `{parsed['schema']}`" if parsed["schema"] else ""))
                    if parsed["schema"] and parsed["schema"] != "SMARTSTORE_CHINA_V5.1":
                        st.warning(f"⚠️ نسخة الملف `{parsed['schema']}` مختلفة عن V5.1 المتوقعة — راجع النتيجة بعد الرفع | schema differs from V5.1")
                    _dsum, _ddet = _china_diff(parsed)
                    if _dsum:
                        st.markdown("**🔍 مقارنة الملف بالمحفوظ حاليًا | Compared with what is stored**")
                        st.dataframe(pd.DataFrame(_dsum), use_container_width=True, hide_index=True)
                        if any(x["عندنا ومش في الملف"] for x in _dsum):
                            st.warning("⚠️ فيه سجلات عندنا ومش في الملف الجديد. لو الملف نسخة كاملة محدّثة استخدم «استبدال الكل» عشان تتشال، "
                                       "ولو تحديث جزئي استخدم «إضافة / تحديث الموجود».")
                        if _ddet:
                            with st.expander(f"📝 تفاصيل التغييرات ({len(_ddet)} حقل)"):
                                st.dataframe(pd.DataFrame(_ddet[:300]), use_container_width=True, hide_index=True)
                    if parsed["lines"]:
                        nb = sum(1 for r in parsed["lines"] if not r.get("barcode"))
                        if nb:
                            st.warning(f"⚠️ {nb} سطر في SKU_LINES من غير barcode — مش هيتربط بالتكويد | line(s) without barcode won't link to Tacweed")
                        st.dataframe(pd.DataFrame(parsed["lines"][:10])[["shipment_id", "container_no", "barcode", "sku",
                                                                         "product_name_en", "total_qty", "stock_stage"]],
                                     use_container_width=True, height=200)

                    def _after(msg):
                        st.session_state["china_upl_ver"] = ver + 1
                        st.session_state.pop("china_bundle", None)
                        st.session_state.pop("china_bundle_src", None)
                        st.session_state["china_flash"] = msg
                        st.rerun()

                    ca, cb = st.columns(2)
                    with ca:
                        if st.button("📤 إضافة / تحديث الموجود | Merge", type="primary",
                                     use_container_width=True, key=f"{key_prefix}_merge"):
                            with st.spinner("جاري الحفظ في Google Sheets..."):
                                counts = _china_do_import(parsed, replace=False, label=upl.name)
                            _after(f"✅ تم الدمج — الإجمالي الآن: {counts}")
                    with cb:
                        if st.button("🔄 استبدال الكل | Replace All", type="secondary",
                                     use_container_width=True, key=f"{key_prefix}_replace"):
                            st.session_state[f"{key_prefix}_confirm_replace"] = True
                    if st.session_state.get(f"{key_prefix}_confirm_replace"):
                        st.warning("⚠️ هيمسح بيانات الصين القديمة (للصفحات الموجودة في الملف) ويرفع الجديدة؟ | Replace existing China data?")
                        cy, cn = st.columns(2)
                        if cy.button("✅ نعم | Yes", key=f"{key_prefix}_yes_replace"):
                            st.session_state[f"{key_prefix}_confirm_replace"] = False
                            with st.spinner("جاري الحفظ في Google Sheets..."):
                                counts = _china_do_import(parsed, replace=True, label=upl.name)
                            _after(f"✅ تم الاستبدال — {counts}")
                        if cn.button("❌ لا | No", key=f"{key_prefix}_no_replace"):
                            st.session_state[f"{key_prefix}_confirm_replace"] = False
                            st.rerun()

        if has_data:
            st.markdown("---")
            st.markdown("**🗑️ حذف | Delete**")
            orders = sorted({l.get("internal_order_no", "") for l in bundle["lines"] if l.get("internal_order_no")}
                            | {s.get("internal_order_no", "") for s in bundle["ships"] if s.get("internal_order_no")})
            d1, d2 = st.columns([3, 1])
            with d1:
                pick = st.selectbox("حذف طلب محدد (NO.xxx) بكل أسطره | Delete one order", ["—"] + orders,
                                    key=f"{key_prefix}_del_pick")
            with d2:
                st.write("")
                if st.button("🗑️ حذف الطلب | Delete", key=f"{key_prefix}_del_order_btn",
                             use_container_width=True, disabled=(pick == "—")):
                    st.session_state[f"{key_prefix}_confirm_del_order"] = pick
            pend = st.session_state.get(f"{key_prefix}_confirm_del_order")
            if pend:
                st.warning(f"⚠️ هيحذف الطلب {pend} بكل أسطره وشحناته؟ | Delete order {pend}?")
                yo, no_ = st.columns(2)
                if yo.button("✅ نعم | Yes", key=f"{key_prefix}_yes_del_order"):
                    st.session_state[f"{key_prefix}_confirm_del_order"] = None
                    with st.spinner("جاري الحذف..."):
                        _china_delete_order(pend)
                    st.session_state.pop("china_bundle", None)
                    st.session_state.pop("china_bundle_src", None)
                    st.session_state["china_flash"] = f"✅ تم حذف الطلب {pend}"
                    st.rerun()
                if no_.button("❌ لا | No", key=f"{key_prefix}_no_del_order"):
                    st.session_state[f"{key_prefix}_confirm_del_order"] = None
                    st.rerun()
            if st.button("🗑️ مسح كل بيانات الصين | Delete ALL China data", key=f"{key_prefix}_del_all_btn"):
                st.session_state[f"{key_prefix}_confirm_del_all"] = True
            if st.session_state.get(f"{key_prefix}_confirm_del_all"):
                st.warning("⚠️ هيمسح كل بيانات الصين (الشحنات والمنتجات والحاويات)؟ | Delete ALL China data?")
                ya, na = st.columns(2)
                if ya.button("✅ نعم | Yes", key=f"{key_prefix}_yes_del_all"):
                    st.session_state[f"{key_prefix}_confirm_del_all"] = False
                    with st.spinner("جاري الحذف..."):
                        _china_delete_all()
                    st.session_state.pop("china_bundle", None)
                    st.session_state.pop("china_bundle_src", None)
                    st.session_state["china_flash"] = "✅ تم مسح كل بيانات الصين"
                    st.rerun()
                if na.button("❌ لا | No", key=f"{key_prefix}_no_del_all"):
                    st.session_state[f"{key_prefix}_confirm_del_all"] = False
                    st.rerun()

# ── تاب شحنات الصين | The China tab body ──
# ── 🚨 لوحة تنبيهات تاب شحنات الصين (من تواريخ الإكسيل) | Alerts board ──
def china_code_decisions(*args, **kwargs):
    return calc.china_code_decisions(_ENV, *args, **kwargs)


def china_build_alerts(*args, **kwargs):
    return calc.china_build_alerts(_ENV, *args, **kwargs)

def render_china_alerts_board(bundle, today):
    """أول حاجة بتظهر في تاب شحنات الصين: عدد التنبيهات + كل تنبيه مع المطلوب عمله. أي خطأ بيظهر تحذير ومبيوقفش التاب."""
    try:
        cfg = _china_cfg()
        alerts, on_track, no_stock, n_prod = china_build_alerts(bundle, today, cfg)
        n0 = sum(1 for a in alerts if a["sev"] == 0)
        n1 = sum(1 for a in alerts if a["sev"] == 1)
        n2 = sum(1 for a in alerts if a["sev"] == 2)
        st.markdown("### 🚨 التنبيهات والمطلوب عمله | Alerts")
        if n0:
            st.error(f"🔴 فيه {n0} تنبيه عاجل محتاج تحرك دلوقتي، و{n1} تنبيه تاني محتاج متابعة.")
        elif n1:
            st.warning(f"🟠 مفيش حاجة عاجلة، بس فيه {n1} تنبيه محتاج متابعة.")
        else:
            st.success("🟢 كل حاجة ماشية في معادها حسب تواريخ ملف الصين.")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("🔴 عاجل", n0)
        m2.metric("🟠 متابعة", n1)
        m3.metric("🔵 قريب", n2)
        m4.metric("🟢 حاويات ماشية في معادها", on_track)
        for a in alerts:
            with st.expander(f"{a['icon']} {a['title']} — {len(a['rows'])}", expanded=a["sev"] <= 1):
                st.markdown(f"🎯 **المطلوب:** {a['action']}")
                st.dataframe(pd.DataFrame(a["rows"]), use_container_width=True, hide_index=True,
                             height=min(420, 60 + 36 * len(a["rows"])))
        st.caption(f"موعد الطلب والفجوة محسوبين على المتوفر بالمستودع ومعدل سحب الكود | منتجات محللة: {n_prod}"
                   + (f" | {no_stock} كود صين بدون مخزون مستودع مسجل (مش داخلين في الحساب)" if no_stock else "")
                   + f" | مدة التوريد المستخدمة {cfg['china_lead_time_days']} يوم")
        with st.expander("⚙️ إعدادات التنبيهات | Alert settings", expanded=False):
            cols = st.columns(len(CHINA_CFG_DEFAULTS))
            new = {}
            for col, (k, dv) in zip(cols, CHINA_CFG_DEFAULTS.items()):
                new[k] = col.number_input(CHINA_CFG_LABELS[k], min_value=0, max_value=999, value=int(cfg[k]), step=1, key=f"china_cfg_{k}")
            if st.button("💾 حفظ الإعدادات | Save", key="china_cfg_save"):
                for k, v in new.items():
                    if int(v) != int(cfg[k]):
                        save_setting(k, str(int(v)))
                st.rerun()
    except Exception as e:
        st.warning(f"⚠️ تعذر بناء لوحة التنبيهات: {e}")

CHINA_STATUS_OPTS = {"🚢 قادمة | Incoming": "incoming", "✅ وصلت المستودع | Received": "received",
                     "🚫 ملغاة | Cancelled": "cancelled", "الكل | All": None}


def render_china_tab():
    st.subheader("🚢 شحنات الصين | China Shipments")
    st.caption("بضاعة جاية من الصين — بيانات شيت China (شيت واحد في Google Sheet بتعبّيه الموظف مباشرة) مربوطة بالتكويد: "
               "الكود01 في تاب التكويد = barcode في SKU_LINES. الكمية القادمة بتظهر كمان جوه تابات مبيعات FBN وFBB | "
               "Incoming goods from China, linked to Tacweed (Code01 = barcode). Also shown inside the FBN/FBB sales tabs.")
    flash = st.session_state.pop("china_flash", None)
    if flash:
        st.success(flash)

    bundle = get_china_bundle()
    _cfw = get_china_flat_ws()
    _cb1, _cb2, _cb3 = st.columns([2, 2, 3])
    with _cb1:
        st.link_button("📝 افتح شيت الصين للتعديل | Open sheet", f"https://docs.google.com/spreadsheets/d/{ss.id}/edit#gid={_cfw.id}",
                       use_container_width=True)
    with _cb2:
        if st.button("🔄 حدّث من الشيت | Refresh", key="china_refresh_flat", use_container_width=True):
            st.session_state["china_flat_ts"] = 0
            st.session_state.pop("china_bundle", None)
            st.session_state.pop("china_bundle_src", None)
            st.rerun()
    with _cb3:
        st.caption("التعديل بيتم في الشيت مباشرة، والتطبيق بيقرأه تلقائي (كل دقيقتين أو بزرار التحديث) | Edit in the sheet; the app re-reads it automatically")
    _ctr = _ct_state().get("last")
    with st.expander("🔗 مزامنة Container tracking" + (f" — آخر مرة {_ctr['time']}" if _ctr else ""),
                     expanded=bool(_ctr and (_ctr["error"] or _ctr["notes"]))):
        st.caption(f"اكتب في China أي مرجع (رقم الحاوية، رقم البوليصة، أو رقم الشحنة STxxxxx) والباقي بيتملا من جوجل شيت «{CT_SHEET_NAME}» "
                   "(الختم، البوليصة، التواريخ، السفينة، الوصول الفعلي، الاستلام، والحالة بتتقدم للأمام). بيشتغل تلقائي كل 5 دقايق أو بالزرار. "
                   "قيم التتبع بتكتب فوق قيم China — فعدّل في Container tracking مش في China.")
        if st.button("🔄 زامن الآن | Sync now", key="ct_sync_now"):
            china_sync_tracking(force=True)
            st.session_state["china_flat_ts"] = 0
            st.session_state.pop("china_bundle", None)
            st.session_state.pop("china_bundle_src", None)
            st.rerun()
        if _ctr:
            if _ctr["error"]:
                st.warning("⚠️ " + _ctr["error"])
            else:
                st.success(f"✅ اتحدّث {_ctr['rows']} صف ({_ctr['cells']} خلية) في آخر مزامنة")
            for _n in _ctr["notes"]:
                st.caption(_n)
    with st.expander("⚡ تحديث فوري جوه الجوجل شيت (Apps Script — تركيب مرة واحدة)", expanded=False):
        st.markdown("بدل ما تستنى فتح التطبيق: الكود ده بيملا China من Container tracking **بعد ثواني من كتابة المرجع**، وكمان لوحده كل 5 دقايق.  \n"
                    "**التركيب:** ① في شيت China: Extensions ▸ Apps Script ② ملف جديد (+) ▸ Script ③ الصق الكود وSave "
                    "④ اختار الدالة **ctSetup** من القايمة فوق ▸ Run ▸ وافق على الصلاحيات (بيحتاج يقرأ Container tracking من Drive). "
                    "لو عندك السكريبت القديم (onEdit) سيبه — الأسماء مختلفة ومش بيتعارضوا.")
        st.code(CHINA_SYNC_SCRIPT, language="javascript")
    with st.expander("📖 دليل تعبئة الشيت للموظفة | Staff guide", expanded=False):
        st.code(CHINA_STAFF_GUIDE, language=None)
        st.markdown("**اختياري — سكريبت Google Apps Script** بيوحّد بيانات الحاوية على كل صفوفها وبيختم آخر تحديث تلقائي:")
        st.code(CHINA_APPS_SCRIPT, language="javascript")
    if bundle.get("conflicts"):
        with st.expander(f"⚠️ {len(bundle['conflicts'])} تعارض في الشيت (نفس الحاوية/الطلب بقيم مختلفة) — راجعها", expanded=False):
            st.caption("التطبيق أخد القيمة الأحدث (حسب «آخر تحديث») مؤقتًا. صحّح القيمة على كل صفوف نفس الحاوية/الطلب في الشيت.")
            st.dataframe(pd.DataFrame(bundle["conflicts"]), use_container_width=True, hide_index=True)
    st.caption("📤 رفع ملف الصين (اختياري) من تاب «رفع ملف» | China file upload is in the «Upload Files» tab")
    _li = load_settings().get("china_last_import", "")
    if _li:
        st.caption(f"🕒 آخر رفع: {_li}")
    lines = bundle["lines"]
    if not (lines or bundle["ships"] or bundle["conts"]):
        st.info("لا توجد بيانات صين بعد — ارفع ملف الإكسيل من تاب «رفع ملف» | No China data yet — upload the Excel file in the «Upload Files» tab")
        return

    today = datetime.now().date()
    render_china_alerts_board(bundle, today)
    st.divider()
    code_skus = _china_code_to_skus()

    def _store_skus(l):
        s = list(code_skus.get(l["_norm"], [])) if l["_norm"] else []
        ns = l.get("noon_sku", "").strip().upper()
        if ns and ns not in s:
            s.append(ns)
        return s

    # ── مؤشرات | KPIs ──
    inc = [l for l in lines if l["_status"] == "incoming"]
    rec = [l for l in lines if l["_status"] == "received"]
    q_inc = sum(l["_qty"] for l in inc)
    q_ship = sum(l["_qty"] for l in inc if l["_idx"] >= CHINA_SHIPPED_IDX)
    q_fact = q_inc - q_ship
    q_rec = sum(l["_qty"] for l in rec)
    n_orders = len({s.get("internal_order_no") for s in bundle["ships"] if s.get("internal_order_no")}
                   | {l.get("internal_order_no") for l in lines if l.get("internal_order_no")})
    n_ship = len({s.get("shipment_id") for s in bundle["ships"] if s.get("shipment_id")}
                 | {l["_shipment_id"] for l in lines if l["_shipment_id"]})
    n_cont = len([c for c in bundle["conts"] if c.get("container_status") != "Cancelled"])
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("الطلبات | Orders", f"{n_orders:,}")
    k2.metric("الشحنات | Shipments", f"{n_ship:,}")
    k3.metric("الحاويات | Containers", f"{n_cont:,}")
    k4.metric("قطع قادمة | Incoming units", f"{q_inc:,}")
    k5.metric("قطع وصلت | Received units", f"{q_rec:,}")
    st.caption(f"📦 من القطع القادمة: {q_ship:,} في الطريق (محمّلة/مشحونة/ميناء/تخليص) | {q_fact:,} لسه في المصنع/قبل الشحن")

    t_jour, t_prod, t_ship, t_cont, t_map = st.tabs(["🧭 مسار الكود | Code Journey", "📦 المنتجات | Products",
                                                     "🧾 الشحنات | Shipments", "🧱 الحاويات | Containers",
                                                     "🔗 الربط بالتكويد | Mapping"])

    # ═══ مسار الكود: كل اللي حصل لكل كود من البداية للنهاية ═══
    with t_jour:
        render_china_journey(lines, _store_skus, today)

    # ═══ المنتجات ═══
    with t_prod:
        f1, f2, f3 = st.columns([3, 2, 3])
        q_txt = f1.text_input("🔍 بحث (باركود/SKU/اسم/شحنة/حاوية/طلب) | Search", key="china_q").strip().upper()
        status_lbl = f2.selectbox("الحالة | Status", list(CHINA_STATUS_OPTS.keys()), key="china_status")
        stage_opts = sorted({l.get("stock_stage", "") for l in lines if l.get("stock_stage")})
        stage_sel = f3.multiselect("المرحلة | Stage", stage_opts, key="china_stage")
        view = st.radio("العرض | View", ["📦 ملخص بالمنتج | By product", "📋 تفصيل الأسطر | Line details"],
                        horizontal=True, key="china_view")
        want = CHINA_STATUS_OPTS[status_lbl]
        flt = []
        for l in lines:
            if want and l["_status"] != want:
                continue
            if stage_sel and l.get("stock_stage") not in stage_sel:
                continue
            if q_txt:
                hay = " ".join([l.get("barcode", ""), l.get("sku", ""), l.get("model_no", ""), l.get("product_name_en", ""),
                                l.get("product_name_ar", ""), l["_shipment_id"], l.get("container_no", ""),
                                l.get("internal_order_no", ""), l.get("po_no", ""), " ".join(_store_skus(l))]).upper()
                if q_txt not in hay:
                    continue
            flt.append(l)
        if not flt:
            st.info("لا توجد نتائج | No results")
        elif view.startswith("📦"):
            agg = {}
            for l in flt:
                k = l["_norm"] or f"NOBC|{l.get('sku', '')}"
                a = agg.setdefault(k, {"barcode": l.get("barcode", ""), "name": l.get("product_name_ar") or l.get("product_name_en", ""),
                                       "model": l.get("model_no") or l.get("sku", ""), "qty": 0, "conts": set(), "ships": set(),
                                       "stages": set(), "etas": [], "skus": set(), "etds": [], "eta_ports": [], "avails": []})
                a["qty"] += l["_qty"]
                if l.get("container_no"):
                    a["conts"].add(l["container_no"])
                if l["_shipment_id"]:
                    a["ships"].add(l["_shipment_id"])
                elif l.get("internal_order_no"):
                    a["ships"].add(l["internal_order_no"])
                if l.get("stock_stage"):
                    a["stages"].add(l["stock_stage"])
                d = _china_date(l["_eta"])
                if d:
                    a["etas"].append(d)
                _a = _china_avail(l)[0]
                if _a:
                    a["avails"].append(_a)
                d_etd = _china_date(l["_d"]["etd"])
                if d_etd:
                    a["etds"].append(d_etd)
                d_ep = _china_date(l["_d"]["eta_port"])
                if d_ep:
                    a["eta_ports"].append(d_ep)
                a["skus"].update(_store_skus(l))
            df_p = pd.DataFrame([{
                "تكويد/Barcode": a["barcode"], "الموديل | Model": a["model"], "المنتج | Product": a["name"],
                "SKUs المتجر | Store SKUs": ", ".join(sorted(a["skus"])) or "— غير مربوط",
                "إجمالي الكمية | Qty": a["qty"], "عدد الحاويات | Containers": len(a["conts"]),
                "الشحنات/الطلبات | Shipments": ", ".join(sorted(a["ships"])),
                "أول خروج من الصين": str(min(a["etds"])) if a["etds"] else "",
                "أقرب وصول متوقع للميناء": str(min(a["eta_ports"])) if a["eta_ports"] else "",
                "أقرب وصول": str(min(a["etas"])) if a["etas"] else "",
                "أقرب إتاحة للبيع (وصول مخزن الرياض)": str(min(a["avails"])) if a["avails"] else "",
                "المراحل | Stages": ", ".join(sorted(a["stages"])),
            } for a in sorted(agg.values(), key=lambda x: -x["qty"])])
            c1, c2 = st.columns(2)
            with c1:
                dl_btn(df_p, "china_products", key="dlbtn_china_products")
            with c2:
                st.info(f"📦 منتجات: {len(df_p)} | إجمالي القطع: {int(df_p['إجمالي الكمية | Qty'].sum()):,}")
            st.dataframe(df_p, use_container_width=True, height=520)
        else:
            rows_d = []
            for l in flt:
                c, s = l["_ctr"], l["_ship"]
                dd = _china_days(l["_eta"], today) if l["_status"] == "incoming" else None
                _st = _china_eta_stats(l["_d"]["eta_port"], l["_d"]["port_arr"], l["_status"] == "cancelled", today)
                _av = _china_avail(l)
                rows_d.append({
                    "shipment_id": l["_shipment_id"], "container_no": l.get("container_no", ""),
                    "تكويد/Barcode": l.get("barcode", ""), "SKU": l.get("sku", ""), "الموديل | Model": l.get("model_no", ""),
                    "المنتج | Product": l.get("product_name_ar") or l.get("product_name_en", ""),
                    "SKUs المتجر | Store SKUs": ", ".join(_store_skus(l)),
                    "الكمية | Qty": l["_qty"], "كراتين | Cartons": _china_int(l.get("cartons", "")),
                    "قطع/كرتون": _china_int(l.get("units_per_carton", "")), "CBM": l.get("cbm", ""),
                    "المرحلة | Stage": l["_stage_eff"], "الحالة | Status": l["_status"],
                    "جاهزية الإنتاج": l["_d"]["ready"], "تاريخ التحميل": l["_d"]["load"],
                    "خروج مخطط من الصين": l["_d"]["etd_plan"], "خروج الشحنة من الصين (actual_etd)": l["_d"]["etd"],
                    "الوصول المتوقع للميناء (eta_port)": l["_d"]["eta_port"], "وصول فعلي للميناء": l["_d"]["port_arr"],
                    "الوصول المتوقع لمخزن الرياض": l["_d"]["eta_dest"], "وصول فعلي للمستودع": l["_d"]["wh_arr"],
                    "الموعد المعتمد": l["_eta_label"], "التاريخ | Date": l["_eta"], "أيام للوصول | Days": dd,
                    "متاح للبيع (وصول مخزن الرياض)": str(_av[0]) if _av[0] else _av[1],
                    "باقي كم يوم للوصول للميناء (days_to_eta)": _st[0], "أيام التأخير (delay_days)": _st[1],
                    "آخر تحديث (last_updated_at)": l["_upd"], "حدّث بواسطة (updated_by)": l["_upd_by"],
                    "حالة الحاوية | Container": c.get("container_status", ""), "السفينة | Vessel": c.get("vessel_name", ""),
                    "B/L": c.get("bl_no") or s.get("bl_no", ""), "المورد | Supplier": s.get("supplier_name", ""),
                    "الطلب | Order": l.get("internal_order_no", ""), "PO": l.get("po_no", ""),
                    "ملاحظات | Notes": l.get("notes", ""),
                })
            df_d = pd.DataFrame(rows_d)
            c1, c2 = st.columns(2)
            with c1:
                dl_btn(df_d, "china_lines", key="dlbtn_china_lines")
            with c2:
                st.info(f"📋 أسطر: {len(df_d)} | إجمالي القطع: {int(df_d['الكمية | Qty'].sum()):,}")
            st.dataframe(df_d, use_container_width=True, height=520)

    # ═══ الشحنات ═══
    with t_ship:
        ships = bundle["ships"]
        if not ships:
            st.info("لا توجد صفوف في SHIPMENTS | No shipment rows")
        else:
            units_by_pair = {}
            for l in lines:
                if l["_status"] != "cancelled":
                    k = (l.get("internal_order_no", ""), l.get("po_no", ""))
                    units_by_pair[k] = units_by_pair.get(k, 0) + l["_qty"]
            st_opts = sorted({s.get("shipment_status", "") for s in ships if s.get("shipment_status")})
            sf1, sf2 = st.columns([3, 3])
            sq = sf1.text_input("🔍 بحث | Search", key="china_ship_q").strip().upper()
            ssel = sf2.multiselect("حالة الشحنة | Status", st_opts, key="china_ship_status")
            rows_s = []
            for s in ships:
                if ssel and s.get("shipment_status") not in ssel:
                    continue
                if sq and sq not in " ".join([s.get("shipment_id", ""), s.get("internal_order_no", ""), s.get("po_no", ""),
                                              s.get("supplier_name", ""), s.get("bl_no", "")]).upper():
                    continue
                rows_s.append({
                    "shipment_id": s.get("shipment_id", ""), "الطلب | Order": s.get("internal_order_no", ""),
                    "PO": s.get("po_no", ""), "المورد | Supplier": s.get("supplier_name", ""),
                    "الحالة | Status": s.get("shipment_status", ""), "الشحن | Mode": s.get("shipment_mode", ""),
                    "ميناء التحميل": s.get("port_of_loading", ""), "ميناء الوصول": s.get("destination_port", ""),
                    "شركة الشحن": s.get("shipping_line", ""), "B/L": s.get("bl_no", ""),
                    "حاويات | Cont.": s.get("container_count", ""),
                    "القطع | Units": units_by_pair.get((s.get("internal_order_no", ""), s.get("po_no", "")), 0),
                    "خروج الشحنة من الصين": s.get("actual_etd", ""), "الوصول المتوقع للميناء": s.get("eta_port", ""),
                    "الوصول المتوقع لمخزن الرياض": s.get("eta_destination", ""),
                    "آخر تحديث (last_updated_at)": s.get("last_updated_at", ""), "حدّث بواسطة (updated_by)": s.get("updated_by", ""),
                    "ملاحظات | Notes": s.get("notes", ""),
                })
            df_s = pd.DataFrame(rows_s)
            c1, c2 = st.columns(2)
            with c1:
                dl_btn(df_s, "china_shipments", key="dlbtn_china_shipments")
            with c2:
                st.info(f"🧾 صفوف: {len(df_s)}")
            st.dataframe(df_s, use_container_width=True, height=520)

    # ═══ الحاويات ═══
    with t_cont:
        conts = bundle["conts"]
        if not conts:
            st.info("لا توجد صفوف في CONTAINERS | No container rows")
        else:
            units_by_cont, prods_by_cont = {}, {}
            for l in lines:
                cn = l.get("container_no", "")
                if cn and l["_status"] != "cancelled":
                    units_by_cont[cn] = units_by_cont.get(cn, 0) + l["_qty"]
                    prods_by_cont.setdefault(cn, set()).add(l.get("model_no") or l.get("sku") or l.get("barcode", ""))
            cs_opts = sorted({c.get("container_status", "") for c in conts if c.get("container_status")})
            cf1, cf2 = st.columns([3, 3])
            cq = cf1.text_input("🔍 بحث | Search", key="china_cont_q").strip().upper()
            csel = cf2.multiselect("حالة الحاوية | Status", cs_opts, key="china_cont_status")
            rows_c = []
            for c in conts:
                if csel and c.get("container_status") not in csel:
                    continue
                if cq and cq not in " ".join([c.get("container_no", ""), c.get("shipment_id", ""), c.get("bl_no", ""),
                                              c.get("vessel_name", "")]).upper():
                    continue
                dd, dly = _china_eta_stats(c.get("eta_port", ""), c.get("port_arrival_actual", ""),
                                           c.get("container_status") == "Cancelled", today)
                rows_c.append({
                    "container_no": c.get("container_no", ""), "shipment_id": c.get("shipment_id", ""),
                    "الحالة | Status": c.get("container_status", ""), "النوع": c.get("container_type", ""),
                    "السفينة | Vessel": c.get("vessel_name", ""), "الرحلة": c.get("voyage_no", ""), "B/L": c.get("bl_no", ""),
                    "ميناء التحميل": c.get("port_of_loading", ""), "ميناء الوصول": c.get("destination_port", ""),
                    "خروج الشحنة من الصين": c.get("actual_etd", ""), "الوصول المتوقع للميناء": c.get("eta_port", ""),
                    "وصول فعلي للميناء": c.get("port_arrival_actual", ""), "وصول فعلي للمستودع": c.get("warehouse_arrival_actual", ""),
                    "باقي كم يوم للوصول للميناء": dd,
                    "أيام التأخير | Delay": dly,
                    "آخر تحديث (last_updated_at)": c.get("last_updated_at", ""), "حدّث بواسطة (updated_by)": c.get("updated_by", ""),
                    "الوجهة النهائية": c.get("final_destination", ""),
                    "التخليص | Customs": c.get("customs_status", ""),
                    "القطع بالنظام | Units": units_by_cont.get(c.get("container_no", ""), 0),
                    "منتجات | Products": ", ".join(sorted(prods_by_cont.get(c.get("container_no", ""), []))),
                    "آخر موقع": c.get("last_known_location", ""), "سبب التأخير": c.get("delay_reason", ""),
                })
            df_c = pd.DataFrame(rows_c)
            c1, c2 = st.columns(2)
            with c1:
                dl_btn(df_c, "china_containers", key="dlbtn_china_containers")
            with c2:
                st.info(f"🧱 حاويات: {len(df_c)}")
            st.dataframe(df_c, use_container_width=True, height=520)

    # ═══ الربط ═══
    with t_map:
        if not lines:
            st.info("لا توجد أسطر منتجات | No product lines")
        else:
            linked = [l for l in lines if _store_skus(l)]
            unlinked = [l for l in lines if not _store_skus(l)]
            codes_all = {l["_norm"] for l in lines if l["_norm"]}
            codes_linked = {l["_norm"] for l in linked if l["_norm"]}
            m1, m2, m3 = st.columns(3)
            m1.metric("باركودات الصين | China barcodes", f"{len(codes_all):,}")
            m2.metric("مربوطة بالتكويد | Linked", f"{len(codes_linked):,}")
            m3.metric("غير مربوطة | Unlinked lines", f"{len(unlinked):,}")
            if not get_tacweed_map():
                st.warning("⚠️ تاب التكويد فاضي — ارفع ملف التكويد من تاب «رفع ملف» | Tacweed is empty — upload it first")
            if unlinked:
                st.markdown("**أسطر مفيهاش SKU في المتجر بنفس الكود | Lines with no store SKU for their code:**")
                df_u = pd.DataFrame([{
                    "تكويد/Barcode": l.get("barcode", ""), "SKU": l.get("sku", ""), "الموديل | Model": l.get("model_no", ""),
                    "المنتج | Product": l.get("product_name_ar") or l.get("product_name_en", ""),
                    "الكمية | Qty": l["_qty"], "shipment_id": l["_shipment_id"], "container_no": l.get("container_no", ""),
                    "المرحلة | Stage": l.get("stock_stage", ""),
                } for l in unlinked])
                dl_btn(df_u, "china_unlinked", key="dlbtn_china_unlinked")
                st.dataframe(df_u, use_container_width=True, height=420)
            else:
                st.success("✅ كل أسطر الصين مربوطة بالتكويد | All China lines are linked")

# ══ inv_map ══

def build_inv_map(*args, **kwargs):
    return calc.build_inv_map(_ENV, *args, **kwargs)

def build_amazon_stock_map(*args, **kwargs):
    return calc.build_amazon_stock_map(_ENV, *args, **kwargs)

def build_amazon_stock_map_by_asin(*args, **kwargs):
    return calc.build_amazon_stock_map_by_asin(_ENV, *args, **kwargs)


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







_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

def dl_btn(df, prefix, label="⬇️ Excel | Download", key=None):
    """الجداول الصغيرة: تحميل مباشر. الكبيرة (+300 صف): زرار «جهّز» الأول — عشان بناء ملف Excel كل مرة بيبطّئ الصفحة."""
    k = key or f"dlbtn_{prefix}"
    if len(df) > 300:
        if st.button(f"📥 جهّز Excel ({len(df):,} صف) | Prepare", key=k + "__prep", use_container_width=True):
            st.session_state[k + "__xl"] = (to_excel(df), f"{prefix}_{file_timestamp()}.xlsx")
        _xl = st.session_state.get(k + "__xl")
        if _xl:
            st.download_button(label, data=_xl[0], file_name=_xl[1], mime=_XLSX_MIME, use_container_width=True, key=k)
        return
    st.download_button(label, data=to_excel(df),
        file_name=f"{prefix}_{file_timestamp()}.xlsx",
        mime=_XLSX_MIME,
        use_container_width=True,
        key=k)

# ══ عرض القوايم الطويلة على دفعات (الرسم هو اللي بيبطّئ مش الحساب) ══
def page_slice(rows, key, n=25):
    return rows[:int(st.session_state.get(key, n))]

def page_more(rows, key, n=25):
    lim = int(st.session_state.get(key, n))
    if len(rows) > lim:
        st.caption(f"معروض {lim} من {len(rows):,} | Showing {lim} of {len(rows):,}")
        _c1, _c2 = st.columns(2)
        with _c1:
            if st.button(f"⬇️ كمان {n} | {n} more", key=key + "_more", use_container_width=True):
                st.session_state[key] = lim + n
                st.rerun()
        with _c2:
            if st.button("📜 الكل (أبطأ) | All", key=key + "_all", use_container_width=True):
                st.session_state[key] = len(rows)
                st.rerun()

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
def build_daily_orders_map(*args, **kwargs):
    return calc.build_daily_orders_map(_ENV, *args, **kwargs)

def build_daily_orders_counts(*args, **kwargs):
    return calc.build_daily_orders_counts(_ENV, *args, **kwargs)

def build_daily_orders_prices(*args, **kwargs):
    return calc.build_daily_orders_prices(_ENV, *args, **kwargs)

# ══════════════════════════════════════════════════════════════════════════
# ══ نسخ منفصلة تمامًا لمبيعات FBB (Fulfilled by Partner) — بتفلتر على عمود
#    "Fulfillment Model" في شيت DailyOrders، وبتاخد بس الصفوف اللي قيمتها فيها
#    "FBP" (زي "Fulfilled by Partner (FBP)"). الدالتين دول نسخة كاملة منفصلة عن
#    build_daily_orders_counts/build_daily_orders_prices الأصليين فوق — مفيش أي
#    تعديل على الأصليين خالص، وتاب "مبيعات نون FBN" لسه بيستخدمهم زي ما هم من
#    غير أي تغيير | Fully separate FBP/FBB-filtered copies — used only by the
#    new FBB sub-tab. The originals above are completely untouched, and the
#    FBN sub-tab keeps using them exactly as before.
# ══════════════════════════════════════════════════════════════════════════



def build_daily_orders_counts_fbb(*args, **kwargs):
    return calc.build_daily_orders_counts_fbb(_ENV, *args, **kwargs)

def build_daily_orders_prices_fbb(*args, **kwargs):
    return calc.build_daily_orders_prices_fbb(_ENV, *args, **kwargs)

# ══════════════════════════════════════════════════════════════════════════
# ══ مبيعات أمازون — نسخة منفصلة تمامًا عن مبيعات نون فوق، بتقرا من شيت
#    DailyOrdersAmazon (أعمدة عربي بترتيب مختلف تمامًا عن DailyOrders)، فكل
#    عمود بيتلاقى بالاسم مش بمكانه. عمود "حاوية كاملة الحمولة": FSAB = تخزين
#    عادي، أي قيمة تانية = تخزين FBA. مفيش أي تعديل على دوال نون فوق خالص |
#    Amazon sales — fully separate from the Noon sales above, reads
#    DailyOrdersAmazon (different Arabic headers, completely different column
#    order), so every column is located by name, not position. "حاوية كاملة
#    الحمولة" column: FSAB = normal storage, anything else = FBA storage. The
#    Noon functions above are completely untouched.
# ══════════════════════════════════════════════════════════════════════════




def build_amazon_orders_asin_map(*args, **kwargs):
    return calc.build_amazon_orders_asin_map(_ENV, *args, **kwargs)

def build_daily_orders_counts_amazon(*args, **kwargs):
    return calc.build_daily_orders_counts_amazon(_ENV, *args, **kwargs)

def build_daily_orders_prices_amazon(*args, **kwargs):
    return calc.build_daily_orders_prices_amazon(_ENV, *args, **kwargs)

def build_daily_orders_counts_amazon_all(*args, **kwargs):
    return calc.build_daily_orders_counts_amazon_all(_ENV, *args, **kwargs)
def build_daily_orders_counts_amazon_normal(*args, **kwargs):
    return calc.build_daily_orders_counts_amazon_normal(_ENV, *args, **kwargs)
def build_daily_orders_counts_amazon_fba(*args, **kwargs):
    return calc.build_daily_orders_counts_amazon_fba(_ENV, *args, **kwargs)
def build_daily_orders_prices_amazon_all(*args, **kwargs):
    return calc.build_daily_orders_prices_amazon_all(_ENV, *args, **kwargs)
def build_daily_orders_prices_amazon_normal(*args, **kwargs):
    return calc.build_daily_orders_prices_amazon_normal(_ENV, *args, **kwargs)
def build_daily_orders_prices_amazon_fba(*args, **kwargs):
    return calc.build_daily_orders_prices_amazon_fba(_ENV, *args, **kwargs)
# ── نسخ مجمّعة على ASIN — تاب مبيعات أمازون (عادي/FBA) بيستخدمها بدل النسخ
#    اللي فوق المجمّعة على MSKU | ASIN-grouped versions — used by the Amazon
#    Sales tab (normal/FBA) instead of the MSKU-grouped ones above ──
def build_daily_orders_counts_amazon_all_asin(*args, **kwargs):
    return calc.build_daily_orders_counts_amazon_all_asin(_ENV, *args, **kwargs)
def build_daily_orders_prices_amazon_all_asin(*args, **kwargs):
    return calc.build_daily_orders_prices_amazon_all_asin(_ENV, *args, **kwargs)
def build_daily_orders_counts_amazon_normal_asin(*args, **kwargs):
    return calc.build_daily_orders_counts_amazon_normal_asin(_ENV, *args, **kwargs)
def build_daily_orders_counts_amazon_fba_asin(*args, **kwargs):
    return calc.build_daily_orders_counts_amazon_fba_asin(_ENV, *args, **kwargs)
def build_daily_orders_prices_amazon_normal_asin(*args, **kwargs):
    return calc.build_daily_orders_prices_amazon_normal_asin(_ENV, *args, **kwargs)
def build_daily_orders_prices_amazon_fba_asin(*args, **kwargs):
    return calc.build_daily_orders_prices_amazon_fba_asin(_ENV, *args, **kwargs)

def build_amazon_asin_sales_rows(*args, **kwargs):
    return calc.build_amazon_asin_sales_rows(_ENV, *args, **kwargs)

# ══ داشبورد أمازون — مفيش عمود Family في شيت DailyOrdersAmazon، فبنرجع إحصائية
#    أقسام فاضية دايمًا (نفس فكرة "لو العمود مش موجود الكود بيكمل عادي" تحت) ══



# ══ خريطة مخزون أمازون (StockAmazon) — بتتبنى مرة واحدة بس زي inv_map بالظبط،
#    وهي مصدر قايمة منتجات تاب مبيعات امازون المستقل عن inv_map بتاع نون ══
# ══ خريطة مخزون أمازون مجمّعة على ASIN — المصدر اللي تاب مبيعات أمازون
#    (عادي/FBA) بيبني منه صفوفه، عشان مخزون نفس الـ ASIN يتجمع حتى لو ظاهر
#    تحت أكتر من MSKU | Amazon stock map aggregated by ASIN — what the
#    Amazon Sales tab (normal/FBA) builds its rows from, so the same ASIN's
#    stock is combined even when it appears under more than one MSKU ══
amz_inv_map_asin = build_amazon_stock_map_by_asin()

# ══════════════════════════════════════════════════════════════════════════
# ══ تريندول | Trendyol — مبيعات تريندول من شيت DailyOrdersTrendyol ══
# الشيت بيتعمل تلقائي أول مرة. الـ SKU في تريندول = عمود Barcode، ونفس تكويد نون (تاب Tacweed)
# بيشتغل عليه. بنخزّن بس الأعمدة المطلوبة للمبيعات — مفيش أسماء عملاء/عناوين/تليفونات/إيميلات.
# Trendyol sales from the DailyOrdersTrendyol sheet (auto-created). SKU = Barcode, and the Noon Tacweed
# mapping applies to it. Only the columns needed for sales are stored — no customer personal data.
# ══════════════════════════════════════════════════════════════════════════
# حالات الطلب اللي مش بتتحسب كمبيعات (مرتجع/ما اتسلمش/لسه مدفوعش/ملغي)


def get_trendyol_bundle():
    return _memo_from_sheet(daily_orders_trendyol_sheet, "trendyol_bundle", _build_trendyol_bundle)

def get_trendyol_commission_map(*args, **kwargs):
    return calc.get_trendyol_commission_map(_ENV, *args, **kwargs)

def build_daily_orders_counts_trendyol(*args, **kwargs):
    return calc.build_daily_orders_counts_trendyol(_ENV, *args, **kwargs)

def build_daily_orders_prices_trendyol(*args, **kwargs):
    return calc.build_daily_orders_prices_trendyol(_ENV, *args, **kwargs)

def build_trendyol_sales_rows(*args, **kwargs):
    return calc.build_trendyol_sales_rows(_ENV, *args, **kwargs)



def render_trendyol_upload(key_prefix):
    """رفع ملف طلبات تريندول (Excel) → بيتحدّث شيت DailyOrdersTrendyol: الطلبات اللي رقمها موجود في الملف
    بتتبدّل بنسختها الجديدة (عشان الحالة تتحدّث)، والباقي بيفضل، وأي طلب أقدم من 60 يوم بيتشال."""
    with st.expander("📤 رفع ملف طلبات تريندول | Upload Trendyol Orders File", expanded=False):
        st.caption("بيتاخد من الملف بس: Barcode (= SKU) / تاريخ الطلب / رقم الطلب / الحالة / الكمية / المبلغ / العمولة / الصورة / اسم المنتج — "
                   "مفيش أسماء عملاء أو عناوين أو تليفونات بتتخزّن.")
        upl = st.file_uploader("ارفع ملف الطلبات (xlsx) | Upload orders (xlsx)", type=["xlsx", "xlsm"], key=f"{key_prefix}_ty_upload")
        if upl:
            try:
                rows, skipped = _ty_parse_orders_xlsx(upl)
            except Exception as e:
                st.error(f"❌ تعذر قراءة الملف: {e}")
                return
            if not rows:
                st.warning("الملف مفيهوش طلبات صالحة | no valid rows")
                return
            dts = [r[2][:10] for r in rows]
            st.info(f"📊 {len(rows):,} سطر | من {min(dts)} إلى {max(dts)} | "
                    f"{len({r[3] for r in rows}):,} طلب | {len({r[1] for r in rows}):,} Barcode"
                    + (f" | تم تخطي {skipped} سطر ناقص" if skipped else ""))
            if st.button("⬆️ رفع / تحديث مبيعات تريندول | Upload", key=f"{key_prefix}_ty_go", type="primary", use_container_width=True):
                headers = TABS_CONFIG["DailyOrdersTrendyol"]
                existing = get_cached(daily_orders_trendyol_sheet, force=True)
                new_orders = {r[3] for r in rows}
                cutoff = (datetime.now() - timedelta(days=60)).strftime("%Y-%m-%d")
                keep = []
                if len(existing) > 1:
                    e_hdr = [str(h).strip().lower() for h in existing[0]]
                    i_ord = e_hdr.index("order number") if "order number" in e_hdr else 3
                    i_dt = e_hdr.index("order date") if "order date" in e_hdr else 2
                    for er in existing[1:]:
                        er = list(er) + [""] * (len(headers) - len(er))
                        if er[i_ord] in new_orders or str(er[i_dt])[:10] < cutoff:
                            continue
                        keep.append(er[:len(headers)])
                final = keep + [r for r in rows if r[2][:10] >= cutoff]
                with st.spinner("جاري الرفع..."):
                    safe_delete_all(daily_orders_trendyol_sheet)
                    ok = True
                    for i in range(0, len(final), 3000):
                        try:
                            daily_orders_trendyol_sheet.append_rows(final[i:i + 3000], value_input_option="RAW")
                        except Exception as e:
                            ok = False
                            st.error(f"❌ فشل الرفع: {e}")
                            break
                clear_cache(daily_orders_trendyol_sheet)
                for _k in ("warehouse_stockout_rows", "warehouse_stockout_sku_lookup", "warehouse_stockout_asin_lookup"):
                    st.session_state.pop(_k, None)
                if ok:
                    st.success(f"✅ تم — الشيت دلوقتي فيه {len(final):,} سطر")
                    st.rerun()


# ══════════════════════════════════════════════════════════════════════════
# ══ أيقونات المنصات + إعلانات تريندول + "مفيش مبيعات" تريندول ══
# أيقونة المنصة: لو حطيت رابط صورة اللوجو بتاعك في Settings (logo_noon / logo_amazon / logo_trendyol)
# بتظهر صورتك، وإلا بتظهر شارة مصغّرة بلون المنصة.
# Platform icons: set your own logo image URL in Settings (logo_noon / logo_amazon / logo_trendyol);
# otherwise a small colored badge is shown.
# ══════════════════════════════════════════════════════════════════════════
_PLAT_STYLE = {
    "noon":     ("n", "#fde047", "#111827", "نون"),
    "amazon":   ("a", "#232f3e", "#ff9900", "أمازون"),
    "trendyol": ("t", "#f27a1a", "#ffffff", "تريندول"),
}

def plat_badge(key, size=22):
    """أيقونة صغيرة للمنصة (HTML inline) | small inline platform icon."""
    letter, bg, fg, _nm = _PLAT_STYLE[key]
    try:
        url = str(load_settings().get("logo_" + key, "") or "").strip()
    except Exception:
        url = ""
    if url.startswith("http"):
        return (f'<img src="{url}" alt="{key}" style="height:{size}px;width:{size}px;object-fit:contain;'
                f'border-radius:50%;vertical-align:middle;margin-inline-end:6px;">')
    return (f'<span style="display:inline-flex;align-items:center;justify-content:center;width:{size}px;height:{size}px;'
            f'border-radius:50%;background:{bg};color:{fg};font-weight:900;font-size:{int(size*0.62)}px;'
            f'font-family:Arial,sans-serif;vertical-align:middle;margin-inline-end:6px;">{letter}</span>')

def plat_header(key, title=None):
    nm = _PLAT_STYLE[key][3]
    st.markdown(
        f'<div dir="rtl" style="display:flex;align-items:center;margin:2px 0 10px;font-size:19px;font-weight:800;">'
        f'{plat_badge(key, 28)}{title or nm}</div>', unsafe_allow_html=True)

# ── إعلانات تريندول | Trendyol ads (تقرير الحملات: Ad name / Cost / Clicks / view(s) / Total sold / Total sales / ROAS) ──




def get_ads_map_trendyol():
    return _memo_from_sheet(ads_ty_sheet, "ty_ads", _build_ty_ads)


def render_generic_ads_upload(key_prefix, sheet, sheet_name, title, key_cols):
    """رفع ملف إعلانات (نون / أمازون) — بيستبدل محتوى الشيت بالكامل بالملف الجديد | replace the ads sheet with the uploaded file."""
    with st.expander(title, expanded=False):
        st.caption(f"الملف بيتقرا بأسماء الأعمدة (مش أماكنها) وبيستبدل كل محتوى شيت {sheet_name}. الهيدر في الشيت زي ما هو.")
        upl = st.file_uploader("ارفع الملف (xlsx / csv)", type=["xlsx", "xlsm", "csv"], key=f"{key_prefix}_{sheet_name}_up")
        if not upl:
            return
        hdr = [h.strip() for h in sheet.row_values(1) if str(h).strip()] or list(TABS_CONFIG[sheet_name])
        try:
            rows, matched, missing = _ads_file_rows(upl, upl.name, hdr, key_cols)
        except Exception as e:
            st.error(f"❌ تعذر قراءة الملف: {e}")
            return
        if not rows:
            st.warning("الملف مفيهوش صفوف | no rows")
            return
        st.info(f"📊 {len(rows)} صف | أعمدة اتطابقت {len(matched)} من {len(hdr)}")
        if missing:
            st.warning("أعمدة مش موجودة في الملف وهتتساب فاضية: " + "، ".join(missing[:12]))
        if len(matched) < 3:
            st.error("❌ الملف مش شبه شكل الشيت (أقل من 3 أعمدة اتطابقت) — مش هرفع.")
            return
        if st.button("⬆️ رفع واستبدال | Upload & replace", key=f"{key_prefix}_{sheet_name}_go", type="primary", use_container_width=True):
            with st.spinner("جاري الرفع..."):
                ok = safe_delete_all(sheet)
                if ok:
                    try:
                        for i in range(0, len(rows), 2000):
                            sheet.append_rows(rows[i:i + 2000], value_input_option="RAW")
                    except Exception as e:
                        ok = False
                        st.error(f"❌ فشل الرفع: {e}")
                else:
                    st.error("❌ مقدرتش أمسح الشيت القديم — محصلش رفع.")
            clear_cache(sheet)
            if ok:
                st.success(f"✅ تم — {len(rows)} صف")
                st.rerun()

def render_trendyol_ads_upload(key_prefix):
    """رفع تقرير إعلانات تريندول — بيستبدل محتوى شيت AdvertisementsTrendyol (التقرير لقطة كاملة بإجماليات الحملات)."""
    with st.expander("📤 رفع تقرير إعلانات تريندول | Upload Trendyol Ads Report", expanded=False):
        st.caption("التقرير بيتخزّن كما هو (لقطة بإجماليات كل حملة من أول تشغيلها) وبيستبدل الموجود. الشيت بيتعمل تلقائي أول مرة.")
        upl = st.file_uploader("ارفع ملف Product Ads Report (xlsx)", type=["xlsx", "xlsm"], key=f"{key_prefix}_tyads_upload")
        if upl:
            try:
                rows = _ty_ads_parse_xlsx(upl)
            except Exception as e:
                st.error(f"❌ تعذر قراءة الملف: {e}")
                return
            if not rows:
                st.warning("الملف مفيهوش حملات | no campaigns")
                return
            st.info(f"📊 {len(rows)} حملة | إجمالي الصرف {sum(_ty_num(r[6]) for r in rows):,.2f}")
            if st.button("⬆️ رفع / تحديث إعلانات تريندول | Upload", key=f"{key_prefix}_tyads_go", type="primary", use_container_width=True):
                with st.spinner("جاري الرفع..."):
                    safe_delete_all(ads_ty_sheet)
                    try:
                        ads_ty_sheet.append_rows(rows, value_input_option="RAW")
                        ok = True
                    except Exception as e:
                        ok = False
                        st.error(f"❌ فشل الرفع: {e}")
                clear_cache(ads_ty_sheet)
                if ok:
                    st.success(f"✅ تم — {len(rows)} حملة")
                    st.rerun()

def _render_ads_performance_tab_trendyol():
    """تاب إعلانات تريندول — على مستوى الحملة (التقرير مفيهوش Barcode، فيه ContentId بس)."""
    plat_header("trendyol", "إعلانات تريندول | Trendyol Ads")
    camps = get_ads_map_trendyol()
    if not camps:
        st.info("ارفع تقرير إعلانات تريندول من تاب «رفع ملف» | Upload the Trendyol ads report in the «Upload Files» tab")
        return
    try:
        min_roas = float(load_settings().get("trendyol_min_roas", "3") or 3)
    except Exception:
        min_roas = 3.0
    st.caption(f"أرقام الحملات تراكمية من تاريخ بدايتها. التقييم: 🔴 صرف بدون مبيعات أو ROAS أقل من 1 | 🟠 ROAS أقل من {min_roas:g} "
               f"(غيّره من Settings بالمفتاح trendyol_min_roas) | 🟢 ROAS {min_roas:g} أو أكتر.")
    stat_opts = ["الكل | All"] + sorted({c["status"] for c in camps if c["status"]})
    pick = st.selectbox("الحالة | Status", stat_opts, key="ty_ads_status")
    shown = [c for c in camps if pick.startswith("الكل") or c["status"] == pick]
    def verdict(c):
        if c["cost"] > 0 and c["sold"] == 0:
            return "🔴 صرف بدون مبيعات"
        if c["cost"] > 0 and c["roas"] < 1:
            return "🔴 خسارة"
        if c["cost"] > 0 and c["roas"] < min_roas:
            return "🟠 ضعيف"
        return "🟢 رابح" if c["cost"] > 0 else "⚪ لسه"
    cost = sum(c["cost"] for c in shown); sales = sum(c["sales"] for c in shown)
    clicks = sum(c["clicks"] for c in shown); views = sum(c["views"] for c in shown); sold = sum(c["sold"] for c in shown)
    k = st.columns(5)
    k[0].metric("💸 الصرف | Spend", f"{cost:,.2f}")
    k[1].metric("💰 المبيعات | Sales", f"{sales:,.2f}")
    k[2].metric("📈 ROAS", f"{(sales / cost) if cost else 0:.2f}")
    k[3].metric("🛒 الطلبات | Sold", f"{sold:,}")
    k[4].metric("🖱️ تكلفة النقرة | CPC", f"{(cost / clicks) if clicks else 0:.2f}")
    bad = [c for c in shown if verdict(c).startswith("🔴")]
    weak = [c for c in shown if verdict(c).startswith("🟠")]
    st.markdown(f"**🔴 {len(bad)} حملة خسارة** &nbsp;|&nbsp; **🟠 {len(weak)} ضعيفة** &nbsp;|&nbsp; **🟢 {sum(1 for c in shown if verdict(c).startswith('🟢'))} رابحة**")
    df = pd.DataFrame([{
        "الحملة | Campaign": c["name"], "الحالة": c["status"], "التقييم": verdict(c), "منتجات": c["n_products"],
        "الصرف": c["cost"], "المبيعات": c["sales"], "ROAS": c["roas"], "طلبات": c["sold"],
        "نقرات": c["clicks"], "مرات ظهور": c["views"], "ميزانية يومية": c["daily_budget"], "بدأت": c["start"],
    } for c in sorted(shown, key=lambda x: -x["cost"])])
    dl_btn(df, "trendyol_ads", key="dlbtn_trendyol_ads")
    _by_bc = get_ads_by_barcode_trendyol()
    if _by_bc:
        _pn2 = get_ty_products()["name"]
        _prow = sorted(({"Barcode": bc, "المنتج": _pn2.get(bc, "")[:50], "الصرف": sum(e["spends"] for e in es),
                         "الإيراد": sum(e["revenue"] for e in es), "الطلبات": round(sum(e["orders"] for e in es), 1),
                         "ROAS": (sum(e["revenue"] for e in es) / sum(e["spends"] for e in es)) if sum(e["spends"] for e in es) else 0,
                         "دقة": "دقيق" if all(e["exact"] for e in es) else "تقريبي"} for bc, es in _by_bc.items()),
                       key=lambda x: -x["الصرف"])
        with st.expander(f"📦 الإعلانات لكل منتج (Barcode) — {len(_prow)} منتج"):
            st.dataframe(pd.DataFrame(_prow), use_container_width=True, hide_index=True)
            dl_btn(pd.DataFrame(_prow), "trendyol_ads_by_barcode", key="dlbtn_trendyol_ads_bc")
    for c in sorted(shown, key=lambda x: -x["cost"]):
        with st.container(border=True):
            h1, h2 = st.columns([4, 2])
            h1.markdown(f"**{c['name']}** &nbsp; {verdict(c)}")
            h2.caption(f"{c['status']} · بدأت {c['start']} · ميزانية يومية {c['daily_budget']:g}")
            m = st.columns(6)
            m[0].markdown(f"💸 **{c['cost']:,.2f}**")
            m[1].markdown(f"💰 **{c['sales']:,.2f}**")
            m[2].markdown(f"📈 ROAS **{c['roas']:.2f}**")
            m[3].markdown(f"🛒 **{c['sold']}**")
            m[4].markdown(f"🖱️ **{c['clicks']:,}**")
            m[5].markdown(f"👁️ **{c['views']:,}**")
            _det = _ty_alloc()[1].get(c["i"])
            if _det and _det["items"]:
                _pn = get_ty_products()["name"]
                with st.expander(f"🧩 منتجات الحملة ({len(_det['items'])}) — " + ("توزيع دقيق (منتج واحد)" if _det["exact"] else "توزيع تقريبي بنسبة المبيعات")):
                    for bc, share, wq in _det["items"]:
                        st.markdown(f"`{bc}` — {_pn.get(bc, '')[:50]} &nbsp;|&nbsp; نصيب {share * 100:.0f}% = صرف **{c['cost'] * share:,.2f}** &nbsp;|&nbsp; مبيعات المنتج من بداية الحملة: {wq}")
                    if _det["unmapped"]:
                        st.caption(f"⚠️ {len(_det['unmapped'])} Content ID مش موجودين في ملف المنتجات: " + ", ".join(_det["unmapped"][:8]))
            elif c["ids"]:
                st.caption(f"{c['n_products']} منتج — ContentId: " + ", ".join(c["ids"][:12]) + (" …" if len(c["ids"]) > 12 else "") + " — ارفع ملف المنتجات من تاب «رفع ملف» لربطهم بالـ Barcode")

# ── "مفيش مبيعات" لتريندول: Barcodes اتباعت في آخر 60 يوم ومفيهاش مبيع في الفترة ──
def compute_no_sale_rows_trendyol(*args, **kwargs):
    return calc.compute_no_sale_rows_trendyol(_ENV, *args, **kwargs)


def render_no_sale_list_trendyol(rows, period_label, dl_key):
    if not rows:
        st.success(f"✅ كل منتجات تريندول اتباعت في {period_label} | all Trendyol products sold in {period_label}")
        return
    df = pd.DataFrame([{"Barcode": r["sku"], "مخزون المستودع | WH stock": r["stock"],
                        "مبيع شهري | Monthly": r["sales_month"], "آخر بيع | Last sale": str(r["last"])} for r in rows])
    c1, c2 = st.columns(2)
    with c1: dl_btn(df, dl_key, key=f"dlbtn_{dl_key}")
    with c2: st.warning(f"⚠️ {len(rows)} منتج بدون مبيعات | products without sales")
    _pgk = f"pg_ns_{dl_key}"
    _pg_items = list(rows)
    for r in page_slice(_pg_items, _pgk):
        ci, cx = st.columns([1, 6])
        with ci:
            show_img(r["img"], 60)
        with cx:
            st.markdown(f"**Barcode:** `{r['sku']}`")
            st.markdown(f"📦 **مخزون المستودع:** {r['stock']} &nbsp;|&nbsp; 📈 **شهري:** {r['sales_month']} &nbsp;|&nbsp; 🕓 **آخر بيع:** {r['last']}")
        st.divider()


    page_more(_pg_items, _pgk)
# ══════════════════════════════════════════════════════════════════════════
# ══ ربط إعلانات تريندول بالـ Barcode — عن طريق ملف المنتجات (Content ID من رابط المنتج) ══
# ملف Products من لوحة تريندول: عمود "Trendyol.com Link" فيه رقم المنتج بعد "-p-" (= Content ID في تقرير الإعلانات).
# بنخزّن بس: Content ID / Barcode / اسم المنتج / صورة 1. الحملة اللي فيها منتج واحد دقيقة، والمتعددة بيتوزع صرفها
# على منتجاتها بنسبة مبيعاتها (من ملف الطلبات) من تاريخ بداية الحملة — ومكتوب عليها "تقريبي".
# ══════════════════════════════════════════════════════════════════════════



def get_ty_products():
    return _memo_from_sheet(ty_products_sheet, "ty_products", _build_ty_products)

def render_trendyol_products_upload(key_prefix):
    with st.expander("📤 رفع ملف منتجات تريندول (لربط الإعلانات بالـ Barcode) | Upload Trendyol Products File", expanded=False):
        st.caption("من لوحة تريندول: ملف Products. بيتاخد منه بس: رقم المنتج (من عمود Trendyol.com Link) / Barcode / اسم المنتج / صورة 1. "
                   "بيستبدل الموجود. الشيت TrendyolProducts بيتعمل تلقائي أول مرة.")
        upl = st.file_uploader("ارفع ملف المنتجات (xlsx) | Upload Products (xlsx)", type=["xlsx", "xlsm"], key=f"{key_prefix}_typrod_upload")
        if upl:
            try:
                rows = _ty_products_parse_xlsx(upl)
            except Exception as e:
                st.error(f"❌ تعذر قراءة الملف: {e}")
                return
            if not rows:
                st.warning("مفيش منتجات صالحة في الملف | no valid rows")
                return
            st.info(f"📊 {len(rows):,} منتج | {len({r[0] for r in rows}):,} Content ID")
            if st.button("⬆️ رفع / تحديث ملف المنتجات | Upload", key=f"{key_prefix}_typrod_go", type="primary", use_container_width=True):
                with st.spinner("جاري الرفع..."):
                    safe_delete_all(ty_products_sheet)
                    ok = True
                    for i in range(0, len(rows), 3000):
                        try:
                            ty_products_sheet.append_rows(rows[i:i + 3000], value_input_option="RAW")
                        except Exception as e:
                            ok = False
                            st.error(f"❌ فشل الرفع: {e}")
                            break
                clear_cache(ty_products_sheet)
                if ok:
                    st.success(f"✅ تم — {len(rows):,} منتج")
                    st.rerun()


def _ty_alloc(*args, **kwargs):
    return calc._ty_alloc(_ENV, *args, **kwargs)

def get_ads_by_barcode_trendyol(*args, **kwargs):
    return calc.get_ads_by_barcode_trendyol(_ENV, *args, **kwargs)


# ══ الأقسام | Departments (عمود Family في DailyOrders — اختياري، لو مش موجود الكود بيكمل عادي) ══


def build_daily_orders_family_stats(*args, **kwargs):
    return calc.build_daily_orders_family_stats(_ENV, *args, **kwargs)

def build_daily_orders_family_stats_fbb(*args, **kwargs):
    return calc.build_daily_orders_family_stats_fbb(_ENV, *args, **kwargs)

# ══════════════════════════════════════════════════════════════════════════
# ══ نسخ "FBN فقط" — عكس فلتر FBB بالظبط (يستبعد أي صف Fulfillment Model فيه
#    FBP/Partner)، عشان الداشبورد بس (تاب "الكل" + تاب "FBN") يبقى صحيح
#    رياضيًا: FBN فقط + FBB = الكل، من غير أي تكرار أو نقص. الدوال دي مستخدمة
#    في داشبورد المبيعات فقط — تاب المبيعات الأصلي (tab14) لسه بيستخدم النسخة
#    غير المفلترة زي ما هو، بدون أي تغيير | "FBN-only" copies — the exact
#    inverse of the FBB filter (excludes any row whose Fulfillment Model is
#    FBP/Partner), so the
#    dashboard's "All Combined" + "FBN" tabs add up correctly: FBN-only + FBB =
#    Combined, with no double-counting or gaps. Used only by the Sales
#    Dashboard — the original Sales tab (tab14) keeps using the unfiltered
#    functions exactly as before, untouched.
# ══════════════════════════════════════════════════════════════════════════
def build_daily_orders_counts_fbn(*args, **kwargs):
    return calc.build_daily_orders_counts_fbn(_ENV, *args, **kwargs)

def build_daily_orders_prices_fbn(*args, **kwargs):
    return calc.build_daily_orders_prices_fbn(_ENV, *args, **kwargs)

def build_daily_orders_family_stats_fbn(*args, **kwargs):
    return calc.build_daily_orders_family_stats_fbn(_ENV, *args, **kwargs)

def is_sku_only_in_excluded_warehouses(*args, **kwargs):
    return calc.is_sku_only_in_excluded_warehouses(_ENV, *args, **kwargs)

def compute_stock_sales_rows(*args, **kwargs):
    return calc.compute_stock_sales_rows(_ENV, *args, **kwargs)

def compute_missing_inventory_rows(*args, **kwargs):
    return calc.compute_missing_inventory_rows(_ENV, *args, **kwargs)

def render_day_counts_md(day_counts, dates, labels):
    """يبني سطر Markdown بمبيعات كل يوم من التواريخ المعطاة بجانب بعض."""
    parts = [f"**{lbl}:** {day_counts.get(d,0)}" for d, lbl in zip(dates, labels)]
    return " &nbsp;|&nbsp; ".join(parts)

def get_recent_expired_info(*args, **kwargs):
    return calc.get_recent_expired_info(_ENV, *args, **kwargs)

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

def get_recent_schedule_rows(*args, **kwargs):
    return calc.get_recent_schedule_rows(_ENV, *args, **kwargs)

def recent_schedule_badge_html(entry):
    """شارة توضيحية لسكو اتجدول (أو انتهت جدولته) خلال آخر 4 أيام — تمنع إعادة الطلب بالغلط."""
    color = "#7c3aed" if entry["source"] != "Expired" else "#b45309"
    return (
        f'<span class="status-badge-lg" style="background:{color};">'
        f'📅 مجدول بتاريخ {entry["date"]} (ASN {entry["asn"]}) [{entry["source_label"]}] — '
        f'خلال آخر 4 أيام، لسه في فترة الوصول — لا تطلبه تاني | '
        f'Scheduled within the last 4 days — still within arrival window, don\'t re-order</span>'
    )

def get_pending_approval_skus(*args, **kwargs):
    return calc.get_pending_approval_skus(_ENV, *args, **kwargs)

def pending_approval_badge_html():
    """شارة توضيحية إن السكو موجود في تاب 'قيد الموافقة' — يعني فيه جدولة مقترحة لسه مستنية
    اعتماد من برنامج الجدولة المكتبي، عشان محدش يطلبه أو يجدوله تاني بالغلط."""
    return (
        '<span class="status-badge-lg" style="background:#b45309;">'
        '⏳ قيد الموافقة | Pending Approval — في انتظار اعتماد الجدولة من برنامج الجدولة المكتبي، '
        'لا تطلبه أو تجدوله تاني | Awaiting schedule approval from the desktop scheduling app — '
        'don\'t re-order or re-schedule</span>'
    )

def compute_recent_scheduled_rows(*args, **kwargs):
    return calc.compute_recent_scheduled_rows(_ENV, *args, **kwargs)

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

def get_latest_schedule_info(*args, **kwargs):
    return calc.get_latest_schedule_info(_ENV, *args, **kwargs)

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

def is_sku_unavailable(*args, **kwargs):
    return calc.is_sku_unavailable(_ENV, *args, **kwargs)

def get_unavailable_ordered_note(*args, **kwargs):
    return calc.get_unavailable_ordered_note(_ENV, *args, **kwargs)

def schedule_coverage_badge(*args, **kwargs):
    return calc.schedule_coverage_badge(_ENV, *args, **kwargs)

def compute_warehouse_stockout_rows(*args, **kwargs):
    return calc.compute_warehouse_stockout_rows(_ENV, *args, **kwargs)

def wh_china_note_html(r):
    """سطر الصين لصف كود المستودع: الكمية القادمة، أقرب وصول، وهل تغطي ولا فيه فجوة | China note for a warehouse row."""
    c = r.get("china")
    if not c:
        return ""
    F = lambda d: _china_fmt_date(d, datetime.now().date())
    parts = [f"🚢 <b>من الصين {c['incoming']:,} قطعة</b>"]
    if c.get("next_date"):
        parts.append(f"أقرب وصول لمخزن الرياض {c['next_qty']:,} قطعة يوم {F(c['next_date'])}")
    if c.get("undated_qty"):
        parts.append(f"{c['undated_qty']:,} قطعة تاريخ وصولها غير مكتوب")
    st_ = c.get("state")
    if st_ == "gap":
        parts.append(f"🔴 فجوة بيع {c['gap_days']} يوم (مبيعات معرضة للضياع ~{c['lost']:,})")
    elif st_ in ("order_late", "order_soon") and c.get("covered"):
        parts.append(f"المخزون + الشحنات يغطوا لحد {F(c['covered'])}")
    elif st_ == "safe" and c.get("covered"):
        parts.append(f"🟢 المخزون + الشحنات يغطوا لحد {F(c['covered'])}")
    return '<div dir="rtl" style="font-size:13px;color:#fed7aa;margin:3px 0;">' + " — ".join(parts) + "</div>"



def compute_no_sale_rows(*args, **kwargs):
    return calc.compute_no_sale_rows(_ENV, *args, **kwargs)

def compute_no_sale_rows_amazon(*args, **kwargs):
    return calc.compute_no_sale_rows_amazon(_ENV, *args, **kwargs)


# ══ 🔀 تحليل الأكواد عبر المتاجر | Cross-store code analysis ══
# كل كود (من تاب التكويد) = منتج واحد: SKU نون = Barcode تريندول، وASIN أمازون من TacweedAmazon.
# بنجمع مبيعات آخر 30 يوم وآخر 30 قبلها لكل متجر ونطلّع تنبيهات المقارنة.

def compute_cross_store_analysis(*args, **kwargs):
    return calc.compute_cross_store_analysis(_ENV, *args, **kwargs)

def amz_map_stock(*args, **kwargs):
    return calc.amz_map_stock(_ENV, *args, **kwargs)



# ══════════════════════════════════════════════
# تحليل فجوات القنوات: FBN / FBB (نون) و FBA / FSAB (أمازون)
# كل صف: المبيعات + مخزون القناة + حالة الجدولة + المتوفر بالمستودع + الاقتراح
# ══════════════════════════════════════════════



def compute_channel_gap(*args, **kwargs):
    return calc.compute_channel_gap(_ENV, *args, **kwargs)

def channel_gap_cached():
    days = max(int(load_settings().get("sales_display_days", "7") or 7), 30)
    sig = (len(get_cached(daily_orders_sheet)), len(get_cached(daily_orders_amazon_sheet)), len(get_tacweed_map()),
           len(get_tacweed_amazon_map()), len(get_warehouse_stock_map()), len(inv_map), len(amz_inv_map_asin),
           len(get_pending_approval_skus()), datetime.now().date(), days, "g1")
    c = st.session_state.get("_gap_cache")
    if c and c[0] == sig:
        return c[1]
    res = compute_channel_gap(days)
    st.session_state["_gap_cache"] = (sig, res)
    return res

def channel_gap_alerts(*args, **kwargs):
    return calc.channel_gap_alerts(_ENV, *args, **kwargs)



def render_channel_gap_view(kind, q=""):
    """العرض المخصص لقناة واحدة في تاب المبيعات + تحميل Excel."""
    res = channel_gap_cached()
    rows = res[kind]
    if q:
        rows = [r for r in rows if q in f"{r['key_up']} {r['code']}".upper()]
    noon = kind in ("fbn", "fbb")
    title = _GAP_KINDS[kind][0]
    if kind in ("fbb", "fsab"):
        # لو الصنف بيبيع كمان في FBN / FBA يبقى هيظهر هناك (كافي أو مش كافي) فمبنعرضوش هنا تاني؛
        # وهنا بيتعرض بس اللي مالوش مبيعات FBN / FBA ومخزونه مش كافي
        rows = [r for r in rows if r["action"] != "ok" and r["other"] <= 0]
    else:
        only_gap = st.checkbox("🔎 الأصناف ذات المخزون غير الكافي فقط | Only insufficient stock", value=False, key=f"gap_only_{kind}")
        if only_gap:
            rows = [r for r in rows if r["action"] != "ok"]
    n_send = sum(1 for r in rows if r["action"] == "send")
    n_nowh = sum(1 for r in rows if r["action"] in ("send_nowh", "low_wh"))
    n_pend = sum(1 for r in rows if r["action"] in ("pending", "scheduled"))
    st.markdown(f"**{title}** — {len(rows)} منتج باع في آخر 30 يوم" + (" ومخزونه غير كافٍ ومالوش مبيعات في " + ("FBN" if kind == "fbb" else "FBA") if kind in ("fbb", "fsab") else "") + f" &nbsp;|&nbsp; 🚚 يُرسل: **{n_send}** &nbsp;|&nbsp; "
                f"⛔ لا يُرسل (المستودع قليل/غير متوفر): **{n_nowh}** &nbsp;|&nbsp; ⏳ مجدول/قيد الموافقة: **{n_pend}**")
    if not rows:
        st.info("مفيش بيانات للقناة دي في الفترة")
        return
    dl_btn(channel_gap_df(rows, kind), f"gap_{kind}", key=f"dlbtn_gap_{kind}")
    if True:      # كل القنوات الأربع (FBN / FBB / FBA / FSAB)
        # ملف «قرب يخلص»: شيت 1 = يُرسل (المستودع يغطي) | شيت 2 = يلزم توريد (المستودع قليل/غير متوفر) — من غير المجدول وقيد الموافقة
        _dest = "FBN" if kind in ("fbn", "fbb") else "FBA"      # FBB بيتنقل لـ FBN، وFSAB بيتبعت لـ FBA
        _all = [r for r in res[kind] if not (kind in ("fbb", "fsab") and r["other"] > 0)]
        _send = [r for r in _all if r["action"] == "send"]
        _supply = [r for r in _all if r["action"] in ("low_wh", "send_nowh")]
        st.caption(f"📑 ملف «قرب يخلص» — 🚚 يُرسل إلى {_dest}: **{len(_send)}** | ⛔ يلزم توريد (غير متوفر/قليل بالمستودع): **{len(_supply)}** "
                   "— (مستبعد منهم المجدول وقيد الموافقة)")
        _kx = f"gap_low_{kind}"
        if st.button(f"📥 جهّز ملف «قرب يخلص» ({len(_send) + len(_supply)} صنف) | Prepare", key=_kx + "__prep", use_container_width=True):
            _buf = io.BytesIO()
            with pd.ExcelWriter(_buf, engine="openpyxl") as _w:
                channel_gap_df(_send, kind).to_excel(_w, sheet_name=f"يُرسل إلى {_dest}"[:31], index=False)
                channel_gap_df(_supply, kind).to_excel(_w, sheet_name="يلزم توريد", index=False)
            st.session_state[_kx + "__xl"] = (_buf.getvalue(), f"low_stock_{kind}_{file_timestamp()}.xlsx")
        _xl = st.session_state.get(_kx + "__xl")
        if _xl:
            st.download_button("⬇️ تحميل ملف «قرب يخلص» | Download", data=_xl[0], file_name=_xl[1], mime=_XLSX_MIME,
                               use_container_width=True, key=_kx)
    for r in page_slice(rows, f"pg_gap_{kind}", 30):
        c1, c2 = st.columns([1, 7])
        with c1:
            show_img(r["img"], 60)
        with c2:
            lk = sku_link_html(r["key"]) if noon else asin_link_html(r["key"])
            st.markdown(f"{lk} &nbsp;|&nbsp; كود `{r['code'] or '—'}`", unsafe_allow_html=True)
            st.markdown(f"📅 أمس **{r['yday']:g}** &nbsp;|&nbsp; 30 يوم **{r['sales']:g}** &nbsp;|&nbsp; "
                        f"{'مخزون FBN' if noon else 'مخزون أمازون'} **{r['stock']:g}**" + ((" 🔴 لا يوجد مخزون" if r["stock"] <= 0 else f" 🔴 غير كافٍ (يكفي {r['cover']:.0f} يوم فقط)") if r["no_stock"] else f" ✅ كافٍ (يكفي {r['cover']:.0f} يوم)" if r["cover"] < 9999 else " ✅ كافٍ")
                        + f" &nbsp;|&nbsp; 🏭 المستودع **{r['wh']:,.0f}**")
            if r["dem"] != r["sales"]:
                st.caption(f"التغطية محسوبة على إجمالي مبيعات {'FBN + FBB' if noon else 'FBA + FSAB'} = {r['dem']:g} في 30 يوم "
                           f"(لأن المخزون المنقول لـ {'FBN' if noon else 'FBA'} هيخدم القناتين)")
            if noon:
                st.markdown(f"**الجدولة:**  \n{r['sched_txt']}")
            if r["suggestion"]:
                (st.success if r["action"] == "send" else st.warning if r["action"] in ("send_nowh", "low_wh") else st.info)(r["suggestion"])
        st.divider()
    page_more(rows, f"pg_gap_{kind}", 30)

def render_cross_store_tab():
    st.caption("كل كود = منتج واحد عبر نون وأمازون وتريندول (من تاب التكويد). آخر 30 يوم مقابل الـ30 اللي قبلها. المنتجات اللي من غير كود مش هتظهر هنا إلا في تنبيه «مبيعات من غير كود».")
    with st.spinner("جاري تحليل الأكواد..."):
        rows, unm, swings = _cross_cached()
    if not rows:
        st.info("لا توجد أكواد — املأ تاب التكويد أولاً | No codes found")
        return
    min_u = st.number_input("أقل عدد مبيعات يعتبر «بيبيع» | Min units", min_value=1, max_value=100, value=5, key="xs_min")
    A = build_cross_alerts(rows, int(min_u))
    try:
        _ga = channel_gap_alerts()
        A["fbbgap"] = [(r, r["suggestion"]) for r in _ga["fbb_nostock"]]
        A["amzgap"] = [(r, r["suggestion"]) for r in _ga["amz_nostock"]]
    except Exception:
        A["fbbgap"], A["amzgap"] = [], []
    titles = [
        ("fbbgap", "🟢 نون FBB لها مبيعات ومخزون FBN غير كافٍ"),
        ("amzgap", "🟠 أمازون لها مبيعات ومخزونها في أمازون غير كافٍ"),
        ("gap", "🟠 بيبيع في متجر وصفر في متجر تاني (موجود فيه)"),
        ("nolist", "🔵 بيبيع وغير معروض في متجر تاني — فرصة نشر"),
        ("oos", "🔴 مخزون صفر في متجر والمستودع فيه بضاعة والكود بيبيع في غيره"),
        ("dead", "⚫ مخزون راكد: مفيش بيع في أي متجر"),
        ("drop", "📉 هبوط مبيعات الكود (كل المتاجر)"),
        ("price", "💲 فرق سعر كبير بين المتاجر"),
        ("unm", "❓ مبيعات من غير كود (مش مربوطة بالتكويد)"),
        ("noimg", "🖼️ منتجات من غير رابط صورة (نزّلها Excel وجيب الروابط)"),
        ("swing", "💱 تذبذب السعر داخل نفس المتجر — آخر 30 يوم"),
        ("swing7", "⚡ تذبذب السعر داخل نفس المتجر — الأسبوع ده (آخر 7 أيام)"),
    ]
    sw_pct = st.number_input("أقل فرق % بين أقل وأعلى سعر بيع في نفس المتجر يعتبر تنبيه | Min same-store swing %", min_value=5, max_value=500,
                             value=30, step=5, key="xs_swing")
    noimg = swings.get("noimg", [])
    swings7 = [w for w in swings["7"] if (w["ratio"] - 1) * 100 >= sw_pct]
    swings = [w for w in swings["30"] if (w["ratio"] - 1) * 100 >= sw_pct]
    cnt = {k: len(v) for k, v in A.items()}
    cnt["unm"] = len(unm)
    cnt["swing"] = len(swings)
    cnt["swing7"] = len(swings7)
    cnt["noimg"] = len(noimg)
    st.markdown(" &nbsp;|&nbsp; ".join(f"{t.split(' ')[0]} **{cnt[k]}**" for k, t in titles))
    pick = st.radio("التنبيه | Alert", [f"{t} ({cnt[k]})" for k, t in titles], key="xs_pick")
    key = titles[[f"{t} ({cnt[k]})" for k, t in titles].index(pick)][0]
    if key == "noimg":
        if not noimg:
            st.success("✅ كل المنتجات النشطة ليها رابط صورة")
            return
        st.caption("منتجات اتباعت في آخر 30 يوم أو ليها مخزون ومالهاش رابط صورة. الإكسيل فيه عمود «Image URL» فاضي تملاه وتلصقه في الشيت المكتوب جنب كل صف "
                   "(links n لنون وتريندول بالـ SKU/الباركود، links a لأمازون بالـ ASIN).")
        _pick = st.radio("المتجر | Store", ["الكل", "🟢 نون", "🟠 أمازون", "🟣 تريندول"], horizontal=True, key="xs_noimg_store")
        _it = [r for r in noimg if _pick == "الكل" or r["store"] == _pick]
        df = pd.DataFrame([{"SKU / ASIN / Barcode": r["key"], "Image URL": "", "المتجر": r["store"], "الكود": r["code"],
                            "مبيع 30 يوم": r["qty"], "مخزون": r["stock"], "حطه في شيت": r["sheet"]} for r in _it])
        dl_btn(df, "missing_images", key="dlbtn_xs_noimg")
        st.caption(f"{len(_it)} منتج")
        for r in page_slice(_it, "pg_xs_noimg", 30):
            st.markdown(f"{r['store']} &nbsp; `{r['key']}`" + (f" &nbsp;|&nbsp; كود `{r['code']}`" if r["code"] else "")
                        + f" &nbsp;—&nbsp; مبيع 30 يوم **{r['qty']}** | مخزون **{r['stock']}** | شيت `{r['sheet']}`")
        page_more(_it, "pg_xs_noimg", 30)
        return
    if key in ("swing", "swing7"):
        _per = "آخر 7 أيام (الأسبوع ده)" if key == "swing7" else "آخر 30 يوم"
        swings = swings7 if key == "swing7" else swings
        if not swings:
            st.success("✅ مفيش تذبذب سعر كبير داخل نفس المتجر")
            return
        st.caption(f"بيقارن أقل وأعلى سعر بيع لنفس المنتج في {_per} داخل نفس المتجر. فرق كبير ممكن يكون خصم/عرض مقصود أو سعر اتسجل غلط.")
        df = pd.DataFrame([{"المتجر": w["store"], "SKU/ASIN": w["key"], "الكود": w["code"], "أقل سعر": round(w["lo"], 2),
                            "يوم أقل سعر": str(w["lo_d"]), "أعلى سعر": round(w["hi"], 2), "يوم أعلى سعر": str(w["hi_d"]),
                            "الفرق %": round((w["ratio"] - 1) * 100), "عدد المبيعات": w["n"]} for w in swings])
        dl_btn(df, "price_swings", key=f"dlbtn_xs_{key}")
        for w in page_slice(swings, f"pg_xs_{key}", 25):
            c1, c2 = st.columns([1, 7])
            with c1:
                show_img(w["img"], 60)
            with c2:
                st.markdown(f"{w['store']} &nbsp; `{w['key']}`" + (f" &nbsp;|&nbsp; كود `{w['code']}`" if w["code"] else ""))
                st.caption(f"💲 أقل سعر {w['lo']:.2f} ({w['lo_d']}) ← أعلى سعر {w['hi']:.2f} ({w['hi_d']}) — فرق {(w['ratio'] - 1) * 100:.0f}% على {w['n']} عملية بيع")
            st.divider()
        page_more(swings, f"pg_xs_{key}", 25)
        return
    if key == "unm":
        if not unm:
            st.success("✅ كل المبيعات مربوطة بكود")
            return
        df = pd.DataFrame([{"المتجر": u["store"], "SKU / ASIN / Barcode": u["key"], "مبيع 30 يوم": u["qty"]} for u in unm])
        dl_btn(df, "unmapped_sales", key="dlbtn_xs_unm")
        for u in page_slice(unm, "pg_xs_unm", 30):
            st.markdown(f"{u['store']} &nbsp; `{u['key']}` &nbsp;—&nbsp; **{u['qty']}** قطعة")
        page_more(unm, "pg_xs_unm", 30)
        return
    items = A[key]
    if not items:
        st.success("✅ مفيش حالات")
        return
    if key in ("fbbgap", "amzgap"):
        df = pd.concat([channel_gap_df([r for r, _t in items if r["kind"] == k], k) for k in dict.fromkeys(r["kind"] for r, _t in items)],
                       ignore_index=True)
        dl_btn(df, f"xs_{key}", key=f"dlbtn_xs_{key}")
        for r, t in page_slice(items, f"pg_xs_{key}", 25):
            c1, c2 = st.columns([1, 7])
            with c1:
                show_img(r["img"], 60)
            with c2:
                st.markdown(f"{_GAP_KINDS[r['kind']][0]} &nbsp; `{r['key']}`" + (f" &nbsp;|&nbsp; كود `{r['code']}`" if r["code"] else "")
                            + f" &nbsp;—&nbsp; مبيع 30 يوم **{r['sales']:g}** &nbsp;|&nbsp; 🏭 مستودع **{r['wh']:,.0f}**")
                if r["sched_txt"]:
                    st.markdown(f"**الجدولة:**  \n{r['sched_txt']}")
                st.markdown(t)
            st.divider()
        page_more(items, f"pg_xs_{key}", 25)
        return
    df = pd.DataFrame([{"الكود": r["code"], "نون": r["noon"], "أمازون": r["amz"], "تريندول": r["ty"],
                        "مخزون نون": r["n_stock"], "مخزون أمازون": r["a_stock"], "مستودع": r["wh"], "ملاحظة": t} for r, t in items])
    dl_btn(df, f"xs_{key}", key=f"dlbtn_xs_{key}")
    for r, t in page_slice(items, f"pg_xs_{key}", 25):
        c1, c2 = st.columns([1, 7])
        with c1:
            show_img(r["img"], 60)
        with c2:
            st.markdown(f"**الكود:** `{r['code']}` &nbsp;|&nbsp; 🟢 {r['noon']} &nbsp; 🟠 {r['amz']} &nbsp; 🟣 {r['ty']} &nbsp;|&nbsp; 🏭 مستودع {r['wh']:,.0f}")
            st.caption(t)
        st.divider()
    page_more(items, f"pg_xs_{key}", 25)

def _cross_cached():
    sig = (len(get_cached(daily_orders_sheet)), len(get_cached(daily_orders_amazon_sheet)),
           len(get_trendyol_bundle()["rows"]), len(get_tacweed_map()), len(get_warehouse_stock_map()), datetime.now().date(), "v5")
    c = st.session_state.get("_xs_cache")
    if c and c[0] == sig and len(c[1]) == 3:
        return c[1]
    res = compute_cross_store_analysis()
    st.session_state["_xs_cache"] = (sig, res)
    return res

ordinal_map = {1:"الثانية|Second",2:"الثالثة|Third",3:"الرابعة|Fourth",4:"الخامسة|Fifth"}


# ══════════════════════════════════════════════
st.title("🏢 عالم الرشاقة للتجارة | Fitness World Trading")

# ══ حساب المرحلين من المبيعات مسبقاً (قبل رسم التابات) ══
def compute_transferred_from_sales(*args, **kwargs):
    return calc.compute_transferred_from_sales(_ENV, *args, **kwargs)

if "transferred_skus_t14" not in st.session_state:
    st.session_state["transferred_skus_t14"] = compute_transferred_from_sales()

try:
    _wh_cn_raw = _china_flat_raw()
except Exception:
    _wh_cn_raw = None
if st.session_state.get("_wh_china_src") is not _wh_cn_raw:      # ملف الصين اتغيّر → أعد حساب المستودع
    for _k in ("warehouse_stockout_rows", "warehouse_stockout_sku_lookup", "warehouse_stockout_asin_lookup"):
        st.session_state.pop(_k, None)
    st.session_state["_wh_china_src"] = _wh_cn_raw
if "warehouse_stockout_rows" not in st.session_state:
    st.session_state["warehouse_stockout_rows"] = compute_warehouse_stockout_rows()
    st.session_state["warehouse_stockout_sku_lookup"] = build_warehouse_stockout_sku_lookup(
        st.session_state["warehouse_stockout_rows"])
    st.session_state["warehouse_stockout_asin_lookup"] = build_warehouse_stockout_asin_lookup(
        st.session_state["warehouse_stockout_rows"])

NAV_ITEMS = [
    ("tab14",    "🛒", "المبيعات | Sales"),
    ("tab_dash", "📊", "داشبورد المبيعات | Dashboard"),
    ("tab9",     "📦", "المخزون | Inventory"),
    ("tab_wh",   "🏭", "مخزون المستودع | Warehouse Stockout"),
    ("tab_china", "🚢", "شحنات الصين | China Shipments"),
    ("tab_contracts", "📑", "العقود | Contracts"),
    ("tab16",    "🗂️", "مخزون بدون بيع | No Sales"),
    ("tab_ads",  "📢", "الإعلانات | Ads"),
    ("tab_upload", "📤", "رفع ملف | Upload Files"),
]
NAV_KEY_LABELS = {k: lbl for k, _i, lbl in NAV_ITEMS}

if "nav_page" not in st.session_state:
    st.session_state["nav_page"] = "tab14"
if "sidebar_open" not in st.session_state:
    st.session_state["sidebar_open"] = True

_sb_open = st.session_state["sidebar_open"]
_sb_width = "250px" if _sb_open else "58px"

# ── Sidebar مخصصة داكنة على يمين الشاشة بدل شريط التابات العلوي — قابلة للفتح
#    والإغلاق بزر. كل عنصر فيها بيمثل نفس الـ tab القديم بالظبط، والمحتوى نفسه
#    (كل الأكواد اللي تحت "with tabX:") فاضل زي ما هو من غير أي تغيير — الاختلاف
#    الوحيد إن tab14/tab_dash/tab9/tab16 بقوا containers بمفتاح ثابت بدل ما يكونوا
#    ناتج st.tabs()، وبيتم إخفاء غير النشط منهم بـ CSS | Custom dark right-side
#    sidebar replacing the old top tab bar — collapsible via a button. Each item
#    maps to exactly the same tab as before, and every line of code under each
#    "with tabX:" block below is 100% unchanged. The only difference is that
#    tab14/tab_dash/tab9/tab16 are now keyed containers instead of st.tabs()
#    output, with the inactive ones hidden via CSS.
st.markdown(f"""
<style>
[data-testid="stAppViewContainer"] > .main {{ order: 1 !important; }}
[data-testid="stSidebar"] {{
    order: 2 !important;
    width: {_sb_width} !important;
    min-width: {_sb_width} !important;
    max-width: {_sb_width} !important;
    transition: width .2s ease, min-width .2s ease, max-width .2s ease;
    background: linear-gradient(180deg,#0f172a 0%,#111827 100%) !important;
    border-left: 1px solid rgba(255,255,255,.08);
}}
[data-testid="stSidebar"] > div:first-child {{ padding-top: 10px; }}
[data-testid="collapsedControl"] {{ display: none !important; }}
.nav-brand {{
    display:flex; align-items:center; justify-content:center; gap:8px;
    padding: 4px 4px 14px 4px; border-bottom: 1px solid rgba(255,255,255,.08);
    margin-bottom: 10px;
}}
.nav-brand .logo-box {{
    width:36px; height:36px; border-radius:12px; background:#1e293b;
    display:flex; align-items:center; justify-content:center; font-size:18px; flex:0 0 auto;
}}
.nav-brand .brand-text {{ color:#e2e8f0; font-weight:800; font-size:14px; text-align:center; white-space:nowrap; }}
[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] div.stButton > button {{
    width: 100%;
    direction: rtl;
    text-align: right !important;
    justify-content: flex-start !important;
    background: transparent !important;
    color: #cbd5e1 !important;
    border: none !important;
    border-radius: 10px !important;
    font-size: 14px !important;
    padding: 10px 14px !important;
    margin-bottom: 4px !important;
    box-shadow: none !important;
}}
[data-testid="stSidebar"] div.stButton > button:hover {{
    background: rgba(255,255,255,.07) !important;
    color: #f1f5f9 !important;
}}
[data-testid="stSidebar"] div.stButton > button[kind="primary"] {{
    background: #ef4444 !important;
    color: #ffffff !important;
    font-weight: 700 !important;
}}
[data-testid="stSidebar"] div.stButton > button[kind="primary"]:hover {{
    background: #dc2626 !important;
}}
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    if _sb_open:
        bcol1, bcol2 = st.columns([1, 5])
        with bcol1:
            if st.button("☰", key="nav_toggle_btn_open"):
                st.session_state["sidebar_open"] = False
                st.rerun()
        with bcol2:
            st.markdown(
                '<div class="nav-brand"><span class="logo-box">🏬</span>'
                '<span class="brand-text">المتجر الذكي<br>منظومة إدارة المتاجر</span></div>',
                unsafe_allow_html=True)
        for _key, _icon, _label in NAV_ITEMS:
            _active = st.session_state["nav_page"] == _key
            if st.button(f"{_icon}  {_label}", key=f"nav_item_{_key}",
                         type="primary" if _active else "secondary",
                         use_container_width=True):
                st.session_state["nav_page"] = _key
                st.rerun()
    else:
        if st.button("☰", key="nav_toggle_btn_closed"):
            st.session_state["sidebar_open"] = True
            st.rerun()
        for _key, _icon, _label in NAV_ITEMS:
            _active = st.session_state["nav_page"] == _key
            if st.button(_icon, key=f"nav_item_c_{_key}", help=NAV_KEY_LABELS[_key],
                         type="primary" if _active else "secondary",
                         use_container_width=True):
                st.session_state["nav_page"] = _key
                st.rerun()

# ══ قايمة أفقية فوق للموبايل فقط | Mobile-only horizontal top nav ══
# القايمة الجانبية فوق دي overlay بتاخد جزء من عرض الشاشة وهي مفتوحة على
# الموبايل — وده أصل كل مشاكل التغطية/القفز اللي واجهناها، لأن باقي الشاشة
# (زي فورم كلمة السر) فاضل ظاهر وقابل للمس جنبها. الحل الأنضف: نخبيها خالص على
# الموبايل (بـ CSS media query تحت)، ونستبدلها بشريط أفقي عادي فوق الصفحة —
# جزء طبيعي من تدفق الصفحة (مش overlay ولا عليه فتح/قفل)، فمفيش أي حاجة تتغطي
# أو تتقفز تاني. الشريط ده بيتحكم في نفس nav_page بالظبط زي القايمة الجانبية،
# فمفيش تعارض بينهم أبداً | The sidebar above is an overlay that eats part of
# the screen width while open on mobile — the root cause of every
# covering/jumping issue we ran into, since the rest of the screen (like the
# password form) stays visible and tappable right beside it. Cleaner fix: hide
# it completely on mobile (CSS media query below) and replace it with an
# ordinary horizontal bar at the top of the page — a normal part of page flow,
# with no open/close and nothing to overlay. It drives the exact same
# nav_page state as the sidebar, so the two can never conflict.
st.markdown("""
<style>
[data-testid="stSidebar"] { display: block; }
.st-key-mobile_topnav { display: none; }
@media (max-width: 767px) {
    [data-testid="stSidebar"] { display: none !important; }
    .st-key-mobile_topnav { display: block !important; }
}
.st-key-mobile_topnav div.stButton > button {
    padding: 8px 4px !important;
    font-size: 20px !important;
}
</style>
""", unsafe_allow_html=True)

with st.container(key="mobile_topnav"):
    st.markdown(
        '<div style="text-align:center;font-weight:800;color:#0f172a;'
        'padding:6px 0 10px 0;">🏬 المتجر الذكي | منظومة إدارة المتاجر</div>',
        unsafe_allow_html=True)
    _mt_cols = st.columns(len(NAV_ITEMS))
    for _mt_col, (_key, _icon, _label) in zip(_mt_cols, NAV_ITEMS):
        with _mt_col:
            _active = st.session_state["nav_page"] == _key
            if st.button(_icon, key=f"nav_item_mobile_{_key}", help=_label,
                         type="primary" if _active else "secondary",
                         use_container_width=True):
                st.session_state["nav_page"] = _key
                st.rerun()

tab14    = st.container(key="navpanel_tab14")
tab_dash = st.container(key="navpanel_tab_dash")
tab9     = st.container(key="navpanel_tab9")
tab_wh   = st.container(key="navpanel_tab_wh")
tab_china = st.container(key="navpanel_tab_china")
tab_contracts = st.container(key="navpanel_tab_contracts")
tab16    = st.container(key="navpanel_tab16")
tab_ads  = st.container(key="navpanel_tab_ads")
tab_upload = st.container(key="navpanel_tab_upload")

for _key, _i, _l in NAV_ITEMS:
    _disp = "block" if st.session_state["nav_page"] == _key else "none"
    st.markdown(f'<style>.st-key-navpanel_{_key} {{ display: {_disp} !important; }}</style>', unsafe_allow_html=True)

def render_inventory_upload():
    """رفع ملف مخزون نون (كان جوه تاب المخزون) | Noon inventory file upload (moved out of the Inventory tab)."""
    links_map = get_links_map()
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

# ══════════════════════════════════════════════
# فحص SKUs محددة: الصق أو ارفع Excel → حالة الجدولة + المبيعات + المخزون + المستودع + الاقتراح (نون FBN/FBB وأمازون FBA/FSAB)
# ══════════════════════════════════════════════



def check_skus(*args, **kwargs):
    return calc.check_skus(_ENV, *args, **kwargs)

def render_sku_check_tab():
    st.subheader("🔎 فحص SKUs محددة | Check specific SKUs")
    st.caption("الصق الأكواد (حتى لو منسوخة من واتساب) أو ارفع ملف Excel/CSV. لو الـ SKU بصيغة Z…Z بنضيف له -1 تلقائي، وممكن تبعته كامل. "
               "بيطلع: الجدولة (قيد الموافقة / مجدول / كانت مجدولة من N أيام)، مبيعات FBN وFBB، المخزون، المستودع، والكمية المقترحة — وأمازون FBA وFSAB لنفس الكود.")
    txt = st.text_area("الصق الأكواد هنا | Paste SKUs", height=170, key="skuchk_txt", placeholder="Z82C7E3332A832BBA3075Z\nN12520014A\nZF0D2CB93BA8D8247F487Z-1")
    upl = st.file_uploader("أو ارفع ملف Excel / CSV | Or upload", type=["xlsx", "xls", "csv"], key="skuchk_file")
    if st.button("🔍 افحص | Check", key="skuchk_go", type="primary", use_container_width=True):
        body = txt or ""
        if upl is not None:
            try:
                body += "\n" + _file_to_text(upl)
            except Exception as e:
                st.error(f"❌ تعذر قراءة الملف: {e}")
        toks = parse_check_tokens(body)
        if not toks:
            st.warning("مفيش أكواد في المدخلات")
        else:
            with st.spinner(f"جاري فحص {len(toks)} كود..."):
                nr, ar, nf = check_skus(toks)
            _buf = io.BytesIO()
            with pd.ExcelWriter(_buf, engine="openpyxl") as _w:
                pd.DataFrame(nr).to_excel(_w, sheet_name="نون", index=False)
                pd.DataFrame(ar).to_excel(_w, sheet_name="أمازون", index=False)
                if nf:
                    pd.DataFrame(nf).to_excel(_w, sheet_name="غير موجود", index=False)
            st.session_state["skuchk_res"] = (nr, ar, nf, _buf.getvalue(), f"sku_check_{file_timestamp()}.xlsx", len(toks))
    res_ = st.session_state.get("skuchk_res")
    if res_:
        nr, ar, nf, xl, fn, n_in = res_
        st.success(f"✅ {n_in} كود مدخل — نون: {len(nr)} | أمازون: {len(ar)} | غير موجود: {len(nf)}")
        st.download_button("⬇️ تحميل Excel | Download", data=xl, file_name=fn, mime=_XLSX_MIME, use_container_width=True, key="skuchk_dl")
        t1, t2, t3 = st.tabs([f"🟢 نون ({len(nr)})", f"🟠 أمازون ({len(ar)})", f"❓ غير موجود ({len(nf)})"])
        with t1:
            st.dataframe(pd.DataFrame(nr), use_container_width=True, hide_index=True)
        with t2:
            st.dataframe(pd.DataFrame(ar), use_container_width=True, hide_index=True)
        with t3:
            st.dataframe(pd.DataFrame(nf), use_container_width=True, hide_index=True)

def render_upload_tab():
    """كل رفع الملفات في مكان واحد — نون مع بعض، أمازون مع بعض، تريندول مع بعض (+ عام: مخزون المستودع والصين)."""
    st.header("📤 رفع ملف | Upload Files")
    st.caption("كل رفع الملفات هنا: نون مع بعض، أمازون مع بعض، تريندول مع بعض | All file uploads live here, grouped by platform")
    _grp = st.radio("القسم | Section", ["🅽 نون | Noon", "📦 أمازون | Amazon", "🟣 تريندول | Trendyol", "🏭 عام (مستودع + صين) | General", "🔎 فحص SKUs | Check SKUs"],
                    horizontal=True, key="up_group", label_visibility="collapsed")
    if _grp.startswith("🅽"):
        render_tacweed_upload("up")
        render_inventory_upload()
        if _ads_monthly_rows() and len(_ads_monthly_rows()) > 1:
            st.info("📊 إعلانات نون بتتحدّث تلقائي من البوت (تاب AdsMonthly) — الرفع اليدوي اتقفل.")
        else:
            render_generic_ads_upload("up", ads_sheet, "Advertisements", "📤 رفع إعلانات نون | Upload Noon Ads", ["Sku", "SKU"])
    elif _grp.startswith("📦"):
        render_tacweed_amazon_upload("up")
        render_generic_ads_upload("up", ads_amz_sheet, "AdvertisementsAmz", "📤 رفع إعلانات أمازون | Upload Amazon Ads", ["المنتجات"])
        st.caption("مخزون أمازون (StockAmazon) بيتحدّث يدويًا في الشيت نفسه | Amazon stock (StockAmazon) is updated manually in the sheet")
    elif _grp.startswith("🟣"):
        render_trendyol_upload("up")
        render_trendyol_ads_upload("up")
        render_trendyol_products_upload("up")
    elif _grp.startswith("🔎"):
        render_sku_check_tab()
    else:
        render_warehouse_stock_upload("up")
        render_china_upload("up", get_china_bundle())
        render_china_images_upload("up")
        _li = load_settings().get("china_last_import", "")
        if _li:
            st.caption(f"🕒 آخر رفع للصين: {_li}")

if st.session_state["nav_page"] == "tab9":
    with tab9:
        if _tab_gate("tab9", "📊 المخزون | Inventory"):
            noon_stock_subtab, amazon_stock_subtab = st.tabs(
                ["🅽 مخزون نون | Noon Stock", "📦 مخزون أمازون | Amazon Stock"])
            with noon_stock_subtab:
                st.subheader("📊 المخزون والمبيع الشهري | Inventory & Monthly Sales")
                links_map = get_links_map()
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
                    _pgk = f"pg_inv_noon_{abs(hash(srch or ''))}"
                    _pg_items = list(filtered_inv.items())
                    for sku_key,info in page_slice(_pg_items, _pgk):
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

                    page_more(_pg_items, _pgk)
            with amazon_stock_subtab:
                st.subheader("📦 مخزون أمازون | Amazon Stock")
                st.caption("مصدره شيت StockAmazon (تصدير Amazon Seller Central) — مستقل تمامًا عن مخزون نون فوق، ومجمّع على أساس ASIN (معرّف منتج أمازون الحقيقي) مش MSKU | Sourced from the StockAmazon sheet (Amazon Seller Central export) — fully independent from Noon's stock above, aggregated by ASIN (Amazon's real product identity), not MSKU.")
                if not amz_inv_map_asin:
                    st.info("لا يوجد مخزون أمازون مرفوع بعد — الشيت StockAmazon بيتحدّث يدوي بره البرنامج | No Amazon stock uploaded yet — the StockAmazon sheet is updated manually outside the app")
                else:
                    links_map_asin_inv = get_links_map_asin()
                    srch_amz_inv = st.text_input("🔍 بحث ASIN / SKU | Search ASIN / SKU", key="srch_amz_inv", placeholder="اكتب ASIN أو SKU...")
                    df_amz_inv_dl = pd.DataFrame([{
                        "ASIN": info.get("asin", asin_up),
                        "MSKUs": ", ".join(info.get("skus", [])),
                        "FNSKU": info.get("fnsku", ""),
                        "الكمية المتاحة | Quantity Available": info.get("stock", 0),
                    } for asin_up, info in amz_inv_map_asin.items()])
                    c1_amz_inv, c2_amz_inv = st.columns(2)
                    with c1_amz_inv: dl_btn(df_amz_inv_dl, "inventory_amazon")
                    with c2_amz_inv:
                        srch_up_amz_inv = srch_amz_inv.strip().upper()
                        filtered_amz_inv = {
                            asin_up: info for asin_up, info in amz_inv_map_asin.items()
                            if not srch_up_amz_inv
                            or srch_up_amz_inv in asin_up
                            or any(srch_up_amz_inv in s.upper() for s in info.get("skus", []))
                        }
                        st.info(f"📦 ASINs: {len(filtered_amz_inv)}")
                    _pgk = f"pg_inv_amz_{abs(hash(srch_up_amz_inv))}"
                    _pg_items = list(filtered_amz_inv.items())
                    for asin_up, info in page_slice(_pg_items, _pgk):
                        c_img_amz_inv, c_info_amz_inv = st.columns([1, 6])
                        mskus_amz_inv = info.get("skus", [])
                        img_amz_inv = links_map_asin_inv.get(asin_up, "") or (
                            links_map.get(mskus_amz_inv[0].upper(), "") if mskus_amz_inv else "")
                        with c_img_amz_inv:
                            show_img(img_amz_inv, 70)
                        with c_info_amz_inv:
                            st.markdown(asin_link_html(info.get("asin", asin_up)), unsafe_allow_html=True)
                            if mskus_amz_inv:
                                st.markdown(f"**MSKU:** `{', '.join(mskus_amz_inv)}`")
                            tc_badge_amz_inv = warehouse_available_badge_amazon(asin_up)
                            if tc_badge_amz_inv:
                                st.markdown(tc_badge_amz_inv, unsafe_allow_html=True)
                            st.markdown(f"📦 **الكمية المتاحة | Quantity Available:** **{info.get('stock', 0)}**")
                            if info.get("fnsku"):
                                st.caption(f"FNSKU: {info['fnsku']}")
                        st.divider()


                    page_more(_pg_items, _pgk)
        # ══ TAB 10 — مراجعة المخزون ══

# ══ TAB WH — توقع نفاد مخزون المستودع (على مستوى كود التكويد) ══
if st.session_state["nav_page"] == "tab_wh":
    with tab_wh:
        if _tab_gate("tab_wh", "🏭 مخزون المستودع | Warehouse Stockout"):
            st.subheader("🏭 توقع نفاد مخزون المستودع | Warehouse Stockout Forecast")
            st.caption(
                "الكمية المتوفرة بالمستودع مسجلة على مستوى كود التكويد، وكود واحد ممكن "
                "يشمل أكتر من SKU — فمعدل السحب هنا بيتحسب من مجموع مبيعات كل الـ SKUs "
                "المشتركة في نفس الكود (فعليًا من DailyOrders)، مش SKU لوحده. الفترة/مدة "
                "التوريد/هامش الأمان بتتغير من شيت الإعدادات (Settings) بإضافة/تعديل "
                "الصفوف: warehouse_sales_window_days, warehouse_lead_time_days, "
                "warehouse_reorder_safety_days."
            )

            wh_settings_now = load_settings()
            wcol1, wcol2, wcol3, wcol4 = st.columns(4)
            wcol1.metric("نافذة حساب السحب | Sales Window",
                          f"{wh_settings_now.get('warehouse_sales_window_days','30')} يوم")
            wcol2.metric("مدة التوريد | Lead Time",
                          f"{wh_settings_now.get('warehouse_lead_time_days','7')} يوم")
            wcol3.metric("هامش الأمان | Safety Margin",
                          f"{wh_settings_now.get('warehouse_reorder_safety_days','3')} يوم")
            with wcol4:
                st.write("")
                if st.button("🔄 إعادة الحساب | Recalculate", use_container_width=True, key="wh_recalc_btn"):
                    st.session_state.pop("warehouse_stockout_rows", None)
                    st.session_state.pop("warehouse_stockout_sku_lookup", None)
                    st.session_state.pop("warehouse_stockout_asin_lookup", None)
                    st.rerun()

            wh_rows = st.session_state.get("warehouse_stockout_rows", [])
            if not wh_rows:
                st.info("لا توجد بيانات مخزون مستودع مرفوعة — ارفع ملف التكويد وملف جرد "
                         "المستودع من تاب «رفع ملف» أولاً | No warehouse stock data uploaded yet "
                         "— upload the Tacweed and Warehouse Stock files from the Inventory tab first")
            else:
                urgent_n = sum(1 for r in wh_rows if r["status"].startswith("🚨"))
                soon_n   = sum(1 for r in wh_rows if r["status"].startswith("🟠"))
                st.markdown(
                    '<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px;">'
                    f'<span class="tacweed-badge" style="background:#7f1d1d;color:#fecaca;">🚨 محتاج طلب الآن: {urgent_n}</span>'
                    f'<span class="tacweed-badge" style="background:#78350f;color:#fde68a;">🟠 قرب موعد الطلب: {soon_n}</span>'
                    f'<span class="tacweed-badge" style="background:#14532d;color:#bbf7d0;">📦 إجمالي الأكواد: {len(wh_rows)}</span>'
                    '</div>', unsafe_allow_html=True)

                sort_cols = st.columns([2, 2])
                with sort_cols[0]:
                    status_filter = st.radio(
                        "فلترة حسب الحالة | Filter by status",
                        ["الكل | All", "🚨 محتاج طلب الآن", "🟠 قرب موعد الطلب", "🟢 آمن", "🚢 فيها شحنات صين"],
                        horizontal=True, key="wh_status_filter")
                with sort_cols[1]:
                    sort_choice = st.radio(
                        "الترتيب | Sort by",
                        ["📈 الأعلى مبيعات", "⏳ الأقرب للنفاد"],
                        horizontal=True, key="wh_sort_choice")

                wh_rows_display = list(wh_rows)
                if sort_choice == "📈 الأعلى مبيعات":
                    wh_rows_display.sort(key=lambda r: -r["daily_draw"])
                else:
                    wh_rows_display.sort(key=lambda r: (r["days_to_stockout"] if r["days_to_stockout"] is not None else 10**9))

                for r in wh_rows_display:
                    if status_filter == "🚢 فيها شحنات صين":
                        if not r.get("china"):
                            continue
                    elif status_filter != "الكل | All" and not r["status"].startswith(status_filter.split(" ")[0]):
                        continue
                    dts_txt = f"{r['days_to_stockout']} يوم" if r["days_to_stockout"] is not None else "—"
                    c1, c2 = st.columns([1, 5])
                    with c1:
                        show_img(r["img"], 60)
                    with c2:
                        st.markdown(
                            f'<span class="tacweed-badge" style="background:#3b0764;color:#e9d5ff;">🏷️ كود: {r["code"]}</span> '
                            f'<span class="tacweed-badge" style="background:{r["color"]}22;color:{r["color"]};border:1px solid {r["color"]}55;">{r["status"]}</span>',
                            unsafe_allow_html=True)
                        st.markdown(
                            f"📦 **الكمية المتوفرة:** {r['qty']:,.0f} &nbsp;|&nbsp; "
                            f"📉 **معدل السحب اليومي:** {r['daily_draw']:g} &nbsp;|&nbsp; "
                            f"⏳ **نفاد خلال:** {dts_txt}")
                        if r["reorder_in"] is not None:
                            rin = r["reorder_in"]
                            rin_txt = f"اليوم أو فات ({rin})" if rin <= 0 else f"{rin} يوم"
                            st.markdown(f"🗓️ **الأيام المتبقية لموعد الطلب:** {rin_txt}")
                        if r["suggested_qty"] > 0:
                            st.markdown(f"💡 **الكمية المقترحة للطلب:** {r['suggested_qty']:,}")
                        _cn_html = wh_china_note_html(r)
                        if _cn_html:
                            st.markdown(_cn_html, unsafe_allow_html=True)
                        with st.expander(f"📋 الـ SKUs/ASINs التابعة لهذا الكود ({len(r['skus']) + len(r.get('asins', []))})"):
                            st.markdown(skus_asins_links_html(r["skus"], r.get("asins"), r.get("sku_sales"), r.get("asin_sales"), r.get("sales_days")), unsafe_allow_html=True)
                    st.divider()

        # ══ TAB 14 — المبيعات ══
    # ══ TAB DASHBOARD — داشبورد تحليلات المبيعات (تاب جديدة منفصلة) ══
    # ══ TAB DASHBOARD — داشبورد تحليلات المبيعات (تاب جديدة منفصلة) ══
def apply_noon_our_names(track):
    """بيسجّل أسماء متاجرنا (إعداد + اللي البوت علّمه) عشان متظهرش كمنافسين."""
    try:
        txt = load_settings().get("noon_our_store_names", "") or ""
        names = [x for x in re.split(r"[,،;\n]+", txt) if x.strip()]
        calc.set_noon_our_names(names + list(calc.noon_learn_our_names(track or {})))
    except AttributeError:
        pass

def render_ads_reco_dashboard():
    """🧠 توصيات الإعلانات في داشبورد المبيعات (اقتراحات فقط)."""
    D = _ads_reco_compute()
    if D.get("status") != "ok":
        return
    res = D["results"]
    st.markdown("### 🧠 توصيات الإعلانات")
    st.caption("اقتراحات فقط — مفيش أي تغيير تلقائي. التفاصيل الكاملة والتجارب في تبويب الإعلانات.")
    with st.expander("ℹ️ معنى TACOS وROAS وباقي المصطلحات"):
        st.markdown(calc.ADS_GLOSSARY_HTML, unsafe_allow_html=True)
    lv = calc.ADS_RECO_LEVELS
    cnt = {}
    for r in res:
        cnt[r["level"]] = cnt.get(r["level"], 0) + 1
    cols = st.columns(len(lv))
    for col, (lk, (lab, colr, bg)) in zip(cols, lv.items()):
        col.markdown(f'<div style="background:{bg};border:1px solid {colr};border-radius:10px;padding:6px;text-align:center;direction:rtl;">'
                     f'<div style="font-size:12px;color:{colr};font-weight:800;">{lab}</div>'
                     f'<div style="font-size:20px;font-weight:800;color:#111827;">{cnt.get(lk, 0)}</div></div>', unsafe_allow_html=True)
    act = [r for r in res if r["level"] in ("urgent", "test_pause", "trim", "increase")]
    if act:
        st.markdown(calc.ads_reco_table_html(act, limit=15), unsafe_allow_html=True)
        if len(act) > 15:
            st.caption(f"معروض أول 15 من {len(act)} — الباقي في Excel.")
    else:
        st.caption("مفيش توصيات تحتاج تدخل دلوقتي.")
    dl_btn(pd.DataFrame(calc.ads_reco_rows(res)), "ads_recommendations_dash", key="ads_reco_dl_dash")


def render_noon_price_alerts():
    """🚨 تنبيهات الأسعار في داشبورد المبيعات: اسكيوهات نون فيها منافس أرخص مننا + Excel."""
    nt = get_noon_track()
    if not nt:
        return
    apply_noon_our_names(nt)
    with st.expander("⚙️ أسماء متاجرنا على نون (عشان متظهرش كمنافسين)", expanded=False):
        cur = load_settings().get("noon_our_store_names", "") or ""
        new = st.text_input("الأسماء زي ما بتظهر على noon.com، مفصولة بفاصلة", value=cur, key="nt_our_names_in",
                            placeholder="Five-Star1, ...")
        if st.button("💾 حفظ", key="nt_our_names_save") and new.strip() != cur.strip():
            save_setting("noon_our_store_names", new.strip())
            st.rerun()
    try:
        price = {k: (v or {}).get("price") for k, v in get_live_map().items()}
        rows = calc.noon_price_alerts(nt, list(nt.keys()), price)
    except AttributeError:
        st.warning("⚠️ ارفع calc.py الجديد عشان تظهر تنبيهات الأسعار")
        return
    with st.expander(f"🚨 تنبيهات الأسعار — منافس أرخص مننا ({len(rows)})", expanded=bool(rows)):
        if rows:
            df = pd.DataFrame(rows)
            st.dataframe(df.drop(columns=["الرابط"]), use_container_width=True, hide_index=True)
            dl_btn(df, "noon_price_alerts", key="dl_noon_alerts")
        else:
            st.success("✅ مفيش منافس أرخص مننا")

def _render_sales_dashboard_body(counts_fn, prices_fn, family_stats_fn, key_suffix, channel="noon"):
    """يعرض كل محتوى داشبورد المبيعات (KPIs + تنبيهات سريعة + تحليل إعلانات + رسوم
    بيانية + أعلى/أبطأ الأصناف + تحليل SKU فردي) — بيتفعّل 6 مرات: 3 لنون
    (الكل/FBN/FBB) و3 لأمازون (الكل/عادي/FBA)، نفس الكود بالظبط في كل الحالات،
    غير بس مصدر البيانات و key_suffix (عشان مفاتيح عناصر Streamlit متتكررش).
    channel="noon" أو "amazon" بيتحكم في مصدر المنتجات (مخزون/صور) اللي
    الأرقام بتتربط بيه — عشان تابات داشبورد أمازون تفضل عن منتجات أمازون بس
    ومتظهرش أي منتج أو تنبيه خاص بنون | Renders the entire sales dashboard
    content — called 6 times: 3 for Noon (All/FBN/FBB) and 3 for Amazon
    (All/Normal/FBA). Identical code every time; only the data source and the
    widget key_suffix differ. channel="noon" or "amazon" controls which
    product/stock/image source the numbers are joined against, so the Amazon
    dashboard tabs stay scoped to Amazon products only and never surface a
    Noon-only product or alert."""
    is_amz_chan = (channel == "amazon")
    if is_amz_chan:
        _links_map_chan = get_links_map_asin()
        chan_product_map = {
            asin_up: {
                "sku": ", ".join(info.get("skus", [])) or info.get("asin", asin_up),
                "img": _links_map_chan.get(asin_up, ""),
                "total_stock": info.get("stock", 0),
            }
            for asin_up, info in amz_inv_map_asin.items()
        }
    elif channel == "trendyol":
        _links_map_chan = get_links_map()
        _ty_b = get_trendyol_bundle()
        chan_product_map = {
            bc: {"sku": bc, "img": _links_map_chan.get(bc) or _ty_b["img"].get(bc, ""), "total_stock": 0}
            for bc in {r["bc"] for r in _ty_b["rows"] if r["ok"]}
        }
    else:
        _links_map_chan = get_links_map()
        chan_product_map = inv_map
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
        key=f"dash_period_td_{key_suffix}")
    analysis_days_td = analysis_period_map_td[analysis_period_label_td]

    today_td  = datetime.now().date()
    cur_dates_td  = [today_td - timedelta(days=i) for i in range(1, analysis_days_td + 1)]
    prev_dates_td = [today_td - timedelta(days=i) for i in range(analysis_days_td + 1, analysis_days_td * 2 + 1)]

    cur_counts_td  = counts_fn(cur_dates_td)
    prev_counts_td = counts_fn(prev_dates_td)
    cur_prices_td  = prices_fn(cur_dates_td)
    prev_prices_td = prices_fn(prev_dates_td)
    live_map_dash  = {} if channel == "trendyol" else get_live_map()

    def _td_total(counts_map, sku_up, dates):
        return sum(counts_map.get(sku_up, {}).get(d, 0) for d in dates)

    def _td_parse_price(p):
        try:
            return float(str(p).replace(",", "").strip())
        except Exception:
            return 0.0

    def _td_revenue_total(prices_map, sku_up, dates, live_map=None):
        """إجمالي الإيراد لهذا الـ SKU خلال التواريخ دي — بيعتمد على سعر البيع
        الأساسي من تاب LIVE (sale_price) لكل طلب، وبيرجع لسعر العرض المسجل مع
        الطلب نفسه بس لو مفيش سعر بيع أساسي مسجل لهذا الـ SKU خالص | Total revenue
        for this SKU over these dates — every order is valued at the LIVE base
        (sale_price) price; only falls back to that order's own recorded offer
        price when no LIVE base price exists for this SKU at all."""
        total = 0.0
        live_price = None
        if live_map:
            live_info = live_map.get(sku_up)
            if live_info:
                live_price = live_info.get("price")
        for d in dates:
            for p, qty in prices_map.get(sku_up, {}).get(d, []):
                if live_price is not None:
                    total += live_price * qty
                elif p and str(p).strip().lower() not in ("", "nan", "none"):
                    total += _td_parse_price(p) * qty
        return total

    rows_td = []
    for sku_up_td, info_td in chan_product_map.items():
        cur_t_td  = _td_total(cur_counts_td, sku_up_td, cur_dates_td)
        prev_t_td = _td_total(prev_counts_td, sku_up_td, prev_dates_td)
        cur_rev_td  = _td_revenue_total(cur_prices_td, sku_up_td, cur_dates_td, live_map_dash)
        prev_rev_td = _td_revenue_total(prev_prices_td, sku_up_td, prev_dates_td, live_map_dash)
        rows_td.append({
            "sku_up": sku_up_td, "sku": info_td.get("sku", sku_up_td), "img": info_td.get("img", ""),
            "cur": cur_t_td, "prev": prev_t_td, "stock": info_td.get("total_stock", 0),
            "cur_rev": cur_rev_td, "prev_rev": prev_rev_td,
        })

    # ── SKUs باعت لكن مش موجودة في ملف المخزون المرفوع — بتتضاف كمان عشان
    #    إجماليات الداشبورد (الكل/FBN/FBB) تعكس عدد الطلبات الحقيقي، مش بس
    #    اللي عندها مخزون مرفوع. الصورة بتتجاب من شيت "links n" (نفس مصدر
    #    الصور في تاب المخزون) لأن الـ SKU ده أصلاً مش في ملف المخزون فمفيش
    #    عمود صورة نجيبها منه هناك | SKUs that sold but aren't in the uploaded
    #    inventory file — added too so the dashboard totals (All/FBN/FBB)
    #    reflect the true order count, not only SKUs with uploaded stock. The
    #    image is looked up from the "links n" sheet (the same source
    #    Inventory itself pulls images from), since this SKU has no Inventory
    #    row to carry an image column at all.
    _inv_skus_seen_td = set(chan_product_map.keys())
    _orphan_skus_td = (set(cur_counts_td.keys()) | set(prev_counts_td.keys())) - _inv_skus_seen_td
    _links_map_dash_orphan = _links_map_chan
    for sku_up_td in _orphan_skus_td:
        cur_t_td  = _td_total(cur_counts_td, sku_up_td, cur_dates_td)
        prev_t_td = _td_total(prev_counts_td, sku_up_td, prev_dates_td)
        if cur_t_td <= 0 and prev_t_td <= 0:
            continue
        cur_rev_td  = _td_revenue_total(cur_prices_td, sku_up_td, cur_dates_td, live_map_dash)
        prev_rev_td = _td_revenue_total(prev_prices_td, sku_up_td, prev_dates_td, live_map_dash)
        rows_td.append({
            "sku_up": sku_up_td, "sku": sku_up_td, "img": _links_map_dash_orphan.get(sku_up_td, ""),
            "cur": cur_t_td, "prev": prev_t_td, "stock": 0,
            "cur_rev": cur_rev_td, "prev_rev": prev_rev_td,
            "not_in_inventory": True,
        })

    total_cur_td  = sum(r["cur"] for r in rows_td)
    total_prev_td = sum(r["prev"] for r in rows_td)
    avg_daily_td  = (total_cur_td / analysis_days_td) if analysis_days_td > 0 else 0
    active_skus_td = sum(1 for r in rows_td if r["cur"] > 0)
    # أصناف بدون مبيعات، لكن استبعدنا اللي مخزونها صفر (طبيعي متبعش لو مفيش مخزون أصلاً — مش تنبيه مفيد)
    zero_rows_td   = sorted([r for r in rows_td if r["cur"] == 0 and r["stock"] > 0], key=lambda r: -r["stock"])
    zero_skus_td   = len(zero_rows_td)
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
        if r_td.get("not_in_inventory"):
            continue  # مفيش بيانات مخزون حقيقية له، فمينفعش نحسب له نفاد مخزون | no real stock data to judge stockout from
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

    # ── بيانات مساعدة لعرض تفاصيل أوضح تحت كل تنبيه (جدولة حديثة / غير متوفر / اعتماد معلّق) ──
    recent_sched_map_td = get_recent_schedule_rows(days_back=4)
    pending_approval_skus_dash = get_pending_approval_skus()

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
        st.caption("ℹ️ لا توجد أسعار مسجلة لهذه الفترة (لا في شيت الطلبات ولا في LIVE) — الإيراد بيتحسب من الصفوف اللي فيها سعر، وبيستخدم سعر LIVE الأساسي كتقدير لو الطلب من غير سعر مسجل | No prices found for this period (neither in the orders sheet nor in LIVE) — revenue is computed from priced rows, falling back to the LIVE base price as an estimate for orders recorded without a price")

    st.write("")


    # ── مخزون Xdock قارب على النفاد (من تاب LIVE) — مخزون منفصل عن Inventory، محتاج تزويد
    #    لو قرب يخلص، مش جدولة | Xdock stock running low (from LIVE sheet) — a separate
    #    stock pool from Inventory; low means it needs restocking, not scheduling ──
    # Xdock مخزون خاص بنون بس (عمود stock_xdock_net في تاب LIVE بتاع نون) — مبيتحسبش
    # ولا بيتعرض خالص في تابات داشبورد أمازون | Xdock is a Noon-only stock pool
    # (the stock_xdock_net column in Noon's LIVE sheet) — never computed or shown
    # on the Amazon dashboard tabs.
    xdock_threshold_dash = int(load_settings().get("xdock_low_stock_threshold", "10") or 10)
    xdock_low_rows_td = []
    if channel == "noon":
        for sku_up_x, live_info_x in live_map_dash.items():
            xnet_x = live_info_x.get("stock_xdock_net")
            if xnet_x is None or xnet_x > xdock_threshold_dash:
                continue
            inv_info_x = inv_map.get(sku_up_x, {})
            other_stock_x = inv_info_x.get("total_stock", 0)
            sales_month_x = inv_info_x.get("sales", 0)
            xdock_low_rows_td.append({
                "sku_up": sku_up_x,
                "sku": inv_info_x.get("sku", sku_up_x),
                "img": inv_info_x.get("img", ""),
                "stock_xdock_net": xnet_x,
                "other_stock": other_stock_x,
                "has_other_stock": other_stock_x > 0,
                "sales_month": sales_month_x,
                "price": live_info_x.get("price"),
                "noon_title": live_info_x.get("noon_title", ""),
            })
        # ترتيب من الأعلى مبيعاً للأقل — عشان نعرف الصنف مهم (بيتباع كتير) ولا لأ
        # قبل ما نقرر مدى إلحاح تزويد مخزون Xdock بتاعه | Sort by monthly sales
        # (highest → lowest) so it's clear which low-Xdock-stock SKUs actually matter
        xdock_low_rows_td.sort(key=lambda r: -r["sales_month"])
    xdock_low_with_other_td = sum(1 for r in xdock_low_rows_td if r["has_other_stock"])
    xdock_low_without_other_td = len(xdock_low_rows_td) - xdock_low_with_other_td

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

    # ── تنبيهات التابين الجداد (مخزون بدون بيع + مخزون المستودع) — نفس منطق
    #    الحساب المستخدم جوه التابين نفسهم بالظبط، عشان الرقم هنا يطابق اللي
    #    هتلاقيه لو فتحت التاب مباشرة | Alerts for the two newer tabs (No Sales
    #    stock + Warehouse Stockout) — using the exact same computation as
    #    inside those tabs themselves, so the number here always matches what
    #    you'd see by opening the tab directly.
    no_sale_rows_dash = compute_no_sale_rows_amazon(days=7) if is_amz_chan else (compute_no_sale_rows_trendyol(days=7) if channel == "trendyol" else compute_no_sale_rows(days=7))
    wh_stockout_rows_dash = st.session_state.get("warehouse_stockout_rows", [])
    wh_urgent_rows_dash = [r for r in wh_stockout_rows_dash if "🚨" in r["status"] or "🟠" in r["status"]]
    wh_now_dash = sum(1 for r in wh_urgent_rows_dash if "🚨" in r["status"])
    wh_soon_dash = len(wh_urgent_rows_dash) - wh_now_dash

    ac5, ac6 = st.columns(2)
    with ac5:
        st.markdown(_alert_chip_html("🗂️", "#fefce8", "#ca8a04", "مخزون بدون بيع",
                    f"{len(no_sale_rows_dash):,}", "SKUs في المخزون بدون أي مبيع آخر 7 أيام"), unsafe_allow_html=True)
    with ac6:
        st.markdown(_alert_chip_html("🏭", "#fdf4ff", "#c026d3", "مخزون مستودع محتاج تزويد",
                    f"{len(wh_urgent_rows_dash):,}",
                    f"{wh_now_dash} اطلب الآن، {wh_soon_dash} قرب موعد الطلب — محسوبة بعد شحنات الصين"), unsafe_allow_html=True)

    # ── فجوات القنوات: FBB لها مبيعات ومخزون FBN غير كافٍ / أمازون كذلك — مع مراعاة المتوفر بالمستودع ──
    try:
        _gap_al = channel_gap_alerts()
        _fbb_ns, _amz_ns = _gap_al["fbb_nostock"], _gap_al["amz_nostock"]
        def _gap_sub(lst, ch):
            return (f"🚚 {sum(1 for r in lst if r['action'] == 'send')} يُرسل إلى {ch} • "
                    f"⛔ {sum(1 for r in lst if r['action'] in ('send_nowh', 'low_wh'))} لا يُرسل (المستودع قليل/غير متوفر) • "
                    f"⏳ {sum(1 for r in lst if r['action'] in ('pending', 'scheduled'))} مجدول/قيد الموافقة")
        ac8, ac9 = st.columns(2)
        with ac8:
            st.markdown(_alert_chip_html("🟢", "#ecfdf5", "#10b981", "نون FBB لها مبيعات ومخزون FBN غير كافٍ",
                        f"{len(_fbb_ns):,}", _gap_sub(_fbb_ns, "FBN")), unsafe_allow_html=True)
        with ac9:
            st.markdown(_alert_chip_html("🟠", "#fff7ed", "#fb923c", "أمازون لها مبيعات ومخزونها في أمازون غير كافٍ",
                        f"{len(_amz_ns):,}", _gap_sub(_amz_ns, "FBA")), unsafe_allow_html=True)
        _gap_show = [("نون FBB", r) for r in _fbb_ns] + [("أمازون", r) for r in _amz_ns]
        _gap_show = [(n, r) for n, r in _gap_show if r["action"] in ("send", "send_nowh", "low_wh")]
        if _gap_show:
            with st.expander(f"🚚 تفاصيل فجوات القنوات ({len(_gap_show)}) — التفاصيل والتحميل في تاب المبيعات"):
                st.dataframe(pd.DataFrame([{"القناة": n, "الكود": r["code"], "SKU/ASIN": r["key"], "مبيعات 30 يوم": r["sales"],
                                            "المستودع": r["wh"], "الاقتراح": " — ".join((r["suggestion"] or "").split("  \n"))} for n, r in _gap_show[:40]]),
                             use_container_width=True, hide_index=True)
    except Exception as _e_gap:
        st.caption(f"تعذر حساب فجوات القنوات: {_e_gap}")

    if channel == "noon":
        ac7 = st.columns(1)[0]
        with ac7:
            st.markdown(_alert_chip_html("🟣", "#faf5ff", "#a855f7", "مخزون Xdock قارب على النفاد",
                        f"{len(xdock_low_rows_td):,}",
                        f"{xdock_threshold_dash} قطعة أو أقل — منهم {xdock_low_with_other_td} عندهم مخزون FBN و {xdock_low_without_other_td} من غيره"),
                        unsafe_allow_html=True)

    # ── تفاصيل الأصناف تحت كل تنبيه — عشان تبان الـ SKUs نفسها اللي بتكوّن الرقم،
    #    مع سياق كافي (غير متوفر؟ ليها جدولة حديثة؟ في انتظار اعتماد؟) عشان التحليل يكون مفهوم ──
    def _render_alert_sku_row(r, lines=None, badges_html=""):
        ci_al, cinfo_al = st.columns([1, 6])
        with ci_al:
            show_img(r["img"], 55)
        with cinfo_al:
            _link_html = asin_link_html(r["sku_up"]) if is_amz_chan else (f'<code>{r["sku"]}</code>' if channel == "trendyol" else sku_link_html(r["sku"]))
            st.markdown(f"{_link_html}", unsafe_allow_html=True)
            for line in (lines or []):
                st.caption(line)
            if badges_html:
                st.markdown(badges_html, unsafe_allow_html=True)

    def _extra_context_badges(r, include_schedule=True):
        """شارات إضافية مشتركة: غير متوفر حالياً / (لو include_schedule) جدولة خلال آخر 4 أيام
        أو في انتظار اعتماد الجدولة."""
        parts = []
        for note in get_unavailable_ordered_note(r["sku"]):
            color = "#f87171" if "غير متوفر" in note else "#38bdf8"
            parts.append(f'<div dir="rtl" style="color:{color};font-size:11px;margin-top:3px;">{note}</div>')
        if include_schedule:
            sched_entry = recent_sched_map_td.get(r["sku_up"])
            if sched_entry:
                color_sc = "#7c3aed" if sched_entry["source"] != "Expired" else "#b45309"
                parts.append(
                    f'<div dir="rtl" style="background:{color_sc}1a;border:1px solid {color_sc};'
                    f'border-radius:6px;padding:5px 9px;margin-top:4px;font-size:12px;'
                    f'color:{color_sc};font-weight:700;line-height:1.7;">'
                    f'📅 مجدول بتاريخ {sched_entry["date"]} — ASN {sched_entry["asn"]} — [{sched_entry["source_label"]}]'
                    f'<div style="font-size:11px;color:#374151;margin-top:2px;font-weight:500;">خلال آخر 4 أيام، لسه في فترة الوصول — لا تطلبه تاني</div>'
                    f'</div>')
            if r["sku_up"] in pending_approval_skus_dash:
                parts.append('<div dir="rtl" style="color:#7dd3fc;font-size:11px;margin-top:3px;">⏳ في انتظار اعتماد الجدولة | Pending schedule approval</div>')
        return "".join(parts)

    with st.expander(f"🔴 عرض منتجات معرضة لنفاد المخزون ({len(attention_rows_td):,}) | Show at-risk SKUs"):
        if attention_rows_td:
            df_al1 = pd.DataFrame([{
                "SKU": r["sku"], "المخزون | Stock": r["stock"],
                "متوسط يومي | Daily Avg": round(r["avg_d"], 2),
                "أيام النفاد | Days to Stockout": r["days_to_so"],
                "الحالة | Status": r["status"],
            } for r in attention_rows_td])
            dl_btn(df_al1, "alert_stockout_risk", key=f"dl_alert_stockout_td_{key_suffix}")
            _pgk_al1 = f"pg_al1_{key_suffix}"
            for r in page_slice(attention_rows_td, _pgk_al1):
                _render_alert_sku_row(
                    r,
                    lines=[f"📦 مخزون: {r['stock']:,} — ⏳ نفاد خلال {r['days_to_so']} يوم", r["status"]],
                    badges_html=_extra_context_badges(r, include_schedule=True))
            page_more(attention_rows_td, _pgk_al1)
        else:
            st.caption("لا توجد أصناف معرضة لنفاد المخزون حالياً")

    with st.expander(f"🟡 عرض منتجات بدون مبيعات في الفترة ({zero_skus_td:,}) | Show no-sale SKUs"):
        st.caption("ℹ️ الأصناف اللي مخزونها صفر مستبعدة من القائمة دي — طبيعي متبعش لو مفيش مخزون أصلاً | SKUs with zero stock are excluded — no stock naturally means no sales")
        if zero_rows_td:
            df_al2 = pd.DataFrame([{
                "SKU": r["sku"], "المخزون | Stock": r["stock"], "مبيعات الفترة | Period Sales": r["cur"],
            } for r in zero_rows_td])
            dl_btn(df_al2, "alert_no_sales", key=f"dl_alert_nosale_td_{key_suffix}")
            _pgk_al2 = f"pg_al2_{key_suffix}"
            for r in page_slice(zero_rows_td, _pgk_al2):
                _render_alert_sku_row(
                    r,
                    lines=[f"📦 مخزون: {r['stock']:,} — لا يوجد مبيعات خلال {analysis_days_td} يوم رغم توفر المخزون"],
                    badges_html=_extra_context_badges(r, include_schedule=False))
            page_more(zero_rows_td, _pgk_al2)
        else:
            st.caption("كل الأصناف اللي معاها مخزون باعت خلال الفترة المحددة")

    with st.expander(f"🟠 عرض منتجات انخفضت مبيعاتها ({len(decline_rows_td):,}) | Show declining SKUs"):
        st.caption("ℹ️ المخزون المعروض هنا بعد استبعاد المستودعات اللي مستبعدة من إعدادات النظام | Stock shown here already excludes warehouses excluded in settings")
        if decline_rows_td:
            df_al3 = pd.DataFrame([{
                "SKU": r["sku"], "الفترة الحالية | Current": r["cur"], "الفترة السابقة | Previous": r["prev"],
                "الانخفاض | Drop %": round((r["prev"] - r["cur"]) / r["prev"] * 100, 1),
                "المخزون | Stock": r["stock"],
                "غير متوفر حالياً؟ | Unavailable now?": "نعم | Yes" if is_sku_unavailable(r["sku_up"]) else "لا | No",
            } for r in decline_rows_td])
            dl_btn(df_al3, "alert_declining", key=f"dl_alert_decline_td_{key_suffix}")
            _pgk_al3 = f"pg_al3_{key_suffix}"
            for r in page_slice(decline_rows_td, _pgk_al3):
                drop_pct = (r["prev"] - r["cur"]) / r["prev"] * 100
                stock_available_dec = r["stock"] > 0
                if not stock_available_dec:
                    reason = '<span style="color:#f87171;font-size:12px;">🔴 لا يوجد مخزون الآن — الانخفاض غالباً بسبب نفاد الكمية | No stock currently — decline is likely stockout-driven</span>'
                elif is_sku_unavailable(r["sku_up"]):
                    reason = '<span style="color:#f87171;font-size:12px;">❌ مسجل حالياً "غير متوفر" — ده ممكن يكون سبب الانخفاض | Currently marked Unavailable — likely explains the decline</span>'
                else:
                    reason = '<span style="color:#4ade80;font-size:12px;">✅ المخزون متاح — الانخفاض مش بسبب نقص المخزون، محتاج مراجعة (سعر/منافسة/إعلانات..) | Stock is available — decline isn\'t stock-related, worth reviewing (price/competition/ads..)</span>'
                # جدولة/انتظار اعتماد بتظهر بس لو مفيش مخزون فعلاً — لو المخزون متاح مفيش داعي نعرض جدولته
                badges = _extra_context_badges(r, include_schedule=not stock_available_dec)
                _render_alert_sku_row(
                    r,
                    lines=[f"📉 {r['cur']:,} مقابل {r['prev']:,} (-{drop_pct:.0f}%)", f"📦 مخزون حالي: {r['stock']:,}"],
                    badges_html=(reason + badges))
            page_more(decline_rows_td, _pgk_al3)
        else:
            st.caption("لا توجد أصناف انخفضت مبيعاتها بنسبة 20%+ حالياً")

    with st.expander(f"🟢 عرض منتجات ارتفعت مبيعاتها ({len(rise_rows_td):,}) | Show rising SKUs"):
        if rise_rows_td:
            df_al4 = pd.DataFrame([{
                "SKU": r["sku"], "الفترة الحالية | Current": r["cur"], "الفترة السابقة | Previous": r["prev"],
                "الارتفاع | Rise %": round((r["cur"] - r["prev"]) / r["prev"] * 100, 1),
            } for r in rise_rows_td])
            dl_btn(df_al4, "alert_rising", key=f"dl_alert_rise_td_{key_suffix}")
            _pgk_al4 = f"pg_al4_{key_suffix}"
            for r in page_slice(rise_rows_td, _pgk_al4):
                rise_pct = (r["cur"] - r["prev"]) / r["prev"] * 100
                _render_alert_sku_row(
                    r,
                    lines=[f"📈 {r['cur']:,} مقابل {r['prev']:,} (+{rise_pct:.0f}%)", f"📦 مخزون حالي: {r['stock']:,}"],
                    badges_html=_extra_context_badges(r, include_schedule=False))
            page_more(rise_rows_td, _pgk_al4)
        else:
            st.caption("لا توجد أصناف ارتفعت مبيعاتها بنسبة 20%+ حالياً")

    if channel == "noon":
        with st.expander(f"🟣 عرض أصناف مخزون Xdock قارب على النفاد ({len(xdock_low_rows_td):,} — {xdock_low_with_other_td} عندهم مخزون FBN | {xdock_low_without_other_td} من غيره) | Show low Xdock-stock SKUs"):
            st.caption("ℹ️ ده مخزون Xdock من تاب LIVE، منفصل عن مخزون Inventory العادي — لو قرب يخلص محتاج تزويد (مش جدولة) لو متوفر عندنا. الأصناف اللي معاها مخزون FBN (Inventory) أقل إلحاحاً من اللي مفيش عندها غير مخزون Xdock بس. الترتيب هنا من الأعلى مبيعاً للأقل عشان تعرف الصنف مهم ولا لأ | This is Xdock stock from the LIVE sheet, separate from regular Inventory — running low means it needs restocking (not scheduling) if available with us. SKUs that also have FBN (Inventory) stock are less urgent than ones relying on Xdock stock alone. Sorted by monthly sales (highest → lowest) so you can tell if it actually matters")
            if xdock_low_rows_td:
                df_al7 = pd.DataFrame([{
                    "SKU": r["sku"], "مخزون Xdock | Xdock Stock": r["stock_xdock_net"],
                    "مخزون FBN | FBN Stock": r["other_stock"],
                    "مبيع شهري | Monthly Sales": r["sales_month"],
                    "السعر | Price": r["price"] if r["price"] is not None else "—",
                } for r in xdock_low_rows_td])
                dl_btn(df_al7, "alert_xdock_low", key=f"dl_alert_xdock_td_{key_suffix}")
                _pgk_al7 = f"pg_al7_{key_suffix}"
                for r in page_slice(xdock_low_rows_td, _pgk_al7):
                    if r["has_other_stock"]:
                        other_badge = f'<span style="color:#4ade80;font-size:12px;">📦 عنده مخزون FBN (Inventory): {r["other_stock"]:,} — أقل إلحاحاً</span>'
                    else:
                        other_badge = '<span style="color:#f87171;font-size:12px;font-weight:700;">🚫 لا يوجد مخزون FBN — الاعتماد على Xdock بس</span>'
                    _render_alert_sku_row(
                        r,
                        lines=[f"🟣 مخزون Xdock: {r['stock_xdock_net']:,} — 📈 مبيع شهري: {r['sales_month']:,}"
                               + (f" — 💵 {r['price']:,.2f} ريال" if r["price"] is not None else "")],
                        badges_html=other_badge)
                page_more(xdock_low_rows_td, _pgk_al7)
            else:
                st.caption(f"لا توجد أصناف مخزون Xdock عندها {xdock_threshold_dash} قطعة أو أقل حالياً")

    def _render_wh_alert_row(r):
        """يعرض صف كود مستودع واحد جوه تفاصيل تنبيه — نفس شكل عرضه في تاب مخزون
        المستودع بالظبط | Renders one warehouse code row inside an alert detail —
        matching the exact same layout used in the Warehouse Stockout tab."""
        dts_txt_wh = f"{r['days_to_stockout']} يوم" if r["days_to_stockout"] is not None else "—"
        c1_wh, c2_wh = st.columns([1, 5])
        with c1_wh:
            show_img(r["img"], 55)
        with c2_wh:
            st.markdown(
                f'<span class="tacweed-badge" style="background:#3b0764;color:#e9d5ff;">🏷️ كود: {r["code"]}</span> '
                f'<span class="tacweed-badge" style="background:{r["color"]}22;color:{r["color"]};border:1px solid {r["color"]}55;">{r["status"]}</span>',
                unsafe_allow_html=True)
            st.markdown(
                f"📦 **الكمية المتوفرة:** {r['qty']:,.0f} &nbsp;|&nbsp; "
                f"📉 **معدل السحب اليومي:** {r['daily_draw']:g} &nbsp;|&nbsp; "
                f"⏳ **نفاد خلال:** {dts_txt_wh}")
            if r["reorder_in"] is not None:
                rin_wh = r["reorder_in"]
                rin_txt_wh = f"اليوم أو فات ({rin_wh})" if rin_wh <= 0 else f"{rin_wh} يوم"
                st.markdown(f"🗓️ **الأيام المتبقية لموعد الطلب:** {rin_txt_wh}")
            if r["suggested_qty"] > 0:
                st.markdown(f"💡 **الكمية المقترحة للطلب:** {r['suggested_qty']:,}")
            _cn_html = wh_china_note_html(r)
            if _cn_html:
                st.markdown(_cn_html, unsafe_allow_html=True)
            st.markdown(skus_asins_links_html(r["skus"], r.get("asins"), r.get("sku_sales"), r.get("asin_sales"), r.get("sales_days")), unsafe_allow_html=True)
        st.divider()

    wh_order_now_rows = sorted(
        [r for r in wh_urgent_rows_dash if "🚨" in r["status"]], key=lambda r: -r["daily_draw"])
    wh_reorder_soon_rows = sorted(
        [r for r in wh_urgent_rows_dash if "🟠" in r["status"]], key=lambda r: -r["daily_draw"])

    with st.expander(f"🚨 عرض أكواد مستودع محتاجة طلب الآن ({len(wh_order_now_rows):,}) | Show warehouse codes to order now"):
        st.caption("ℹ️ مرتبة من الأعلى معدل سحب (مبيعات) للأقل — الكود الأكثر بيعاً يظهر أولاً | Sorted by daily draw rate (highest → lowest) — the best-selling code appears first")
        if wh_order_now_rows:
            df_wh1 = pd.DataFrame([{
                "الكود | Code": r["code"], "الكمية المتوفرة | Qty": r["qty"],
                "معدل السحب اليومي | Daily Draw": r["daily_draw"],
                "الكمية المقترحة | Suggested Qty": r["suggested_qty"],
            } for r in wh_order_now_rows])
            dl_btn(df_wh1, "alert_wh_order_now", key=f"dl_alert_wh_now_td_{key_suffix}")
            _pgk_wh1 = f"pg_wh1_{key_suffix}"
            for r in page_slice(wh_order_now_rows, _pgk_wh1):
                _render_wh_alert_row(r)
            page_more(wh_order_now_rows, _pgk_wh1)
        else:
            st.caption("لا توجد أكواد مستودع محتاجة طلب الآن حالياً")

    with st.expander(f"🟠 عرض أكواد مستودع قرب موعد الطلب ({len(wh_reorder_soon_rows):,}) | Show warehouse codes reordering soon"):
        st.caption("ℹ️ مرتبة من الأعلى معدل سحب (مبيعات) للأقل — الكود الأكثر بيعاً يظهر أولاً | Sorted by daily draw rate (highest → lowest) — the best-selling code appears first")
        if wh_reorder_soon_rows:
            df_wh2 = pd.DataFrame([{
                "الكود | Code": r["code"], "الكمية المتوفرة | Qty": r["qty"],
                "معدل السحب اليومي | Daily Draw": r["daily_draw"],
                "الأيام المتبقية لموعد الطلب | Days to Reorder": r["reorder_in"],
            } for r in wh_reorder_soon_rows])
            dl_btn(df_wh2, "alert_wh_reorder_soon", key=f"dl_alert_wh_soon_td_{key_suffix}")
            _pgk_wh2 = f"pg_wh2_{key_suffix}"
            for r in page_slice(wh_reorder_soon_rows, _pgk_wh2):
                _render_wh_alert_row(r)
            page_more(wh_reorder_soon_rows, _pgk_wh2)
        else:
            st.caption("لا توجد أكواد مستودع قرب موعد الطلب حالياً")

    # ── 🚢 تنبيهات الصين على مستوى كود المستودع (نفس محرك تاب شحنات الصين) ──
    _wh_cn_gap = sorted([r for r in wh_stockout_rows_dash if (r.get("china") or {}).get("state") == "gap"],
                        key=lambda r: -(r["china"].get("lost") or 0))
    _wh_cn_nodate = sorted([r for r in wh_stockout_rows_dash if (r.get("china") or {}).get("state") == "need_date"],
                           key=lambda r: -r["daily_draw"])
    _cn_lead = _china_cfg()["china_lead_time_days"]
    _wh_cn_cov = sorted([r for r in wh_stockout_rows_dash
                         if r.get("china") and r["china"].get("state") in ("safe", "order_soon", "order_late")
                         and r["china"].get("dated_qty") and r["days_to_stockout"] is not None and r["days_to_stockout"] <= _cn_lead],
                        key=lambda r: r["days_to_stockout"])
    for _ttl, _lst, _kk, _empty in (
            (f"🚢🟢 أكواد قربت تخلص وشحنة الصين مغطياها ({len(_wh_cn_cov):,}) | Running low but China covers", _wh_cn_cov, "cnc",
             "لا توجد حالات"),
            (f"🚢🔴 أكواد هتخلص قبل وصول شحنة الصين — فجوة بيع ({len(_wh_cn_gap):,}) | China gap", _wh_cn_gap, "cng",
             "مفيش أكواد هتخلص قبل وصول شحنتها"),
            (f"🚢🟠 أكواد مخزونها قرب يخلص وشحنة الصين من غير تاريخ وصول ({len(_wh_cn_nodate):,}) | China undated", _wh_cn_nodate, "cnd",
             "لا توجد حالات")):
        with st.expander(_ttl):
            if _kk == "cnc":
                st.caption(f"ℹ️ أكواد مخزون المستودع بيكفيها {_cn_lead} يوم أو أقل، بس الشحنات اللي ليها تاريخ وصول لمخزن الرياض هتوصل قبل ما تخلص "
                           "ومفيش فجوة بيع. مرتبة بالأقرب للنفاد. المطلوب: متابعة الشحنة بس (ولو اتأخرت هتتحول لتنبيه فجوة).")
            elif _kk == "cng":
                st.caption("ℹ️ مرتبة بالأكتر مبيعات معرضة للضياع. المطلوب: سرّع الشحنة مع الصين أو اشحن جوي/بديل، واطلب كمية جديدة.")
            else:
                st.caption("ℹ️ المطلوب: اطلب من موظفة الصين تكتب تاريخ الوصول لمخزن الرياض في شيت China عشان نقدر نحسب الفجوة.")
            if _lst:
                dl_btn(pd.DataFrame([{
                    "الكود": r["code"], "المتوفر": r["qty"], "سحب/يوم": r["daily_draw"], "من الصين": r["china"]["incoming"],
                    "قطع بدون تاريخ": r["china"]["undated_qty"], "أيام الفجوة": r["china"].get("gap_days", 0),
                    "مبيعات معرضة للضياع": r["china"].get("lost", 0)} for r in _lst]), f"alert_{_kk}", key=f"dl_alert_{_kk}_{key_suffix}")
                _pgk_cn = f"pg_{_kk}_{key_suffix}"
                for r in page_slice(_lst, _pgk_cn):
                    _render_wh_alert_row(r)
                page_more(_lst, _pgk_cn)
            else:
                st.caption(_empty)

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
    dept_stats_td = family_stats_fn(cur_dates_td, live_map_dash)
    if not dept_stats_td:
        st.caption("لا توجد بيانات أقسام (عمود Family) لهذه الفترة — العمود اختياري ولا يؤثر على باقي التحليلات | No department (Family) data for this period — the column is optional and does not affect other analytics")
    else:
        dept_sorted_td = sorted(dept_stats_td.items(), key=lambda x: -x[1]["revenue"])
        df_dept_td = pd.DataFrame([{
            "القسم | Department": dept,
            "عدد الطلبات | Orders": v["orders"],
            "الإيراد | Revenue (ريال)": round(v["revenue"], 2),
        } for dept, v in dept_sorted_td])
        dl_btn(df_dept_td, "sales_by_department", key=f"dl_dept_td_{key_suffix}")
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
        dl_btn(df_top10_td, "top_sellers", key=f"dl_top10_td_{key_suffix}")
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
        dl_btn(df_slow10_td, "slow_movers", key=f"dl_slow10_td_{key_suffix}")
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
    selected_sku_td = st.selectbox("اختر SKU للتحليل | Select SKU", ["—"] + sku_options_td, key=f"dash_sku_td_{key_suffix}")
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

    if channel == "noon" and key_suffix == "all":
        render_noon_price_alerts()
        render_ads_reco_dashboard()

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
        dl_btn(df_att_td, "needs_attention", key=f"dl_attention_td_{key_suffix}")
        _pgk_att = f"pg_att_{key_suffix}"
        for r in page_slice(attention_rows_td, _pgk_att):
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
        page_more(attention_rows_td, _pgk_att)
    else:
        st.success("✅ لا توجد أصناف محتاجة انتباه حالياً | No items currently need attention")

ADS_RECO_TABS = ("AdsHistory", "AdsMonthly", "AdsCampaignHist", "AdsExperiments")
ADS_EXP_SHEET = "AdsExperiments"
ADS_RECO_TTL = 600
ADS_RECO_VER = 5      # بيزيد مع أي تغيير في شكل نتيجة الحساب عشان الكاش القديم ما يتستخدمش


def _ads_reco_load(force=False):
    """يقرا تابات تاريخ الإعلانات والتجارب من الشيت (كاش 10 دقايق)."""
    c = st.session_state.get("_ads_reco_raw")
    if c and not force and time.time() - c[0] < ADS_RECO_TTL:
        return c[1]
    raw, ok = {}, False
    for _try in range(2):          # محاولة تانية لو جوجل رفض مؤقتاً
        try:
            have = {w.title for w in ss.worksheets()}
            names = [t for t in ADS_RECO_TABS if t in have]
            if names:
                res = ss.values_batch_get(["'" + t + "'" for t in names])
                for t, vr in zip(names, res.get("valueRanges", [])):
                    raw[t] = vr.get("values", []) or []
            ok = True
            break
        except Exception:
            time.sleep(1.5)
    if not ok:
        # فشل القراءة: ما نخزّنش نتيجة فاضية في الكاش (كانت بتخلّي التوصيات ترجع للشهري 10 دقايق)
        return c[1] if c else {}
    st.session_state["_ads_reco_raw"] = (time.time(), raw)
    return raw


def _ads_exp_append(row):
    try:
        try:
            ws = ss.worksheet(ADS_EXP_SHEET)
        except Exception:
            ws = ss.add_worksheet(title=ADS_EXP_SHEET, rows=200, cols=len(calc.ADS_EXP_HEADERS))
            ws.update(range_name="A1", values=[calc.ADS_EXP_HEADERS])
        ws.append_row(row, value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        st.error(f"تعذر حفظ التجربة: {e}")
        return False


def _ads_exp_close(exp_id, note):
    try:
        ws = ss.worksheet(ADS_EXP_SHEET)
        vals = ws.get_all_values()
        hdr = vals[0]
        ci, si, di, ni = hdr.index("ID"), hdr.index("Status"), hdr.index("Closed Date"), hdr.index("Result Note")
        for i, r in enumerate(vals[1:], start=2):
            if len(r) > ci and r[ci] == exp_id:
                ws.update_cell(i, si + 1, "اتقفلت")
                ws.update_cell(i, di + 1, datetime.now().date().isoformat())
                ws.update_cell(i, ni + 1, note)
                return True
    except Exception as e:
        st.error(f"تعذر قفل التجربة: {e}")
    return False


def _ads_daily_rows():
    c = st.session_state.get("_ads_daily_rows")
    if c and time.time() - c[0] < 600:
        return c[1]
    rows = None
    try:
        rows = ss.worksheet("AdsDaily").get_all_values()
    except Exception:
        rows = c[1] if c else None
    st.session_state["_ads_daily_rows"] = (time.time(), rows)
    return rows


def get_ads_map_daily():
    """إعلانات أمس (من تاب AdsDaily اللي البوت بيكتبه) بنفس شكل get_ads_map."""
    rows = _ads_daily_rows()
    if rows and len(rows) > 1:
        try:
            return calc.get_ads_map(_ENV, data=rows)
        except Exception:
            return {}
    return {}


def _ads_yday_for(sku_up):
    return get_ads_map_daily().get(sku_up) or []


def _ads_row_key(a):
    return (a.get("store", ""), a.get("cid", ""), a.get("psku", ""), a.get("offer", ""))


def _ads_yday_assign(ent, yl):
    """يربط كل صف شهري بصف أمس بتاعه (واحد لواحد): بمفتاح (متجر، حملة، PSKU، كود العرض) ثم بالحملة+العرض ثم بالاسم لو فريد."""
    used, out = set(), {}
    def take(pred):
        for idx, ad in enumerate(ent):
            if idx in out:
                continue
            for k, y in enumerate(yl):
                if k not in used and pred(ad, y):
                    out[idx] = y; used.add(k)
                    break
    take(lambda a, y: _ads_row_key(a) == _ads_row_key(y) and any(_ads_row_key(a)))
    take(lambda a, y: (a.get("cid"), a.get("offer")) == (y.get("cid"), y.get("offer")) and a.get("cid"))
    take(lambda a, y: a.get("cid") and a.get("cid") == y.get("cid") and a.get("store") == y.get("store"))
    return out


def _ads_line(ad, cps=True):
    return (f"👁️ ظهور: {ad['views']:,.0f} &nbsp;|&nbsp; 🖱️ نقرات: {ad['clicks']:,.0f} &nbsp;|&nbsp; "
            f"🛒 طلبات: {ad['orders']:,.0f} &nbsp;|&nbsp; ➕ سلة: {ad['atc']:,.0f}<br>"
            f"💸 مصروف: {ad['spends']:,.2f} ريال &nbsp;|&nbsp; 💰 إيراد: {ad['revenue']:,.2f} ريال<br>"
            f"📊 CTR (نسبة النقر): {ad['ctr']:.2f}% &nbsp;|&nbsp; 🎯 ROAS (عائد الإعلان): {ad['roas']:.2f}<br>"
            f"CPC (تكلفة النقرة): {ad['cpc']:.2f} &nbsp;|&nbsp; CPS (تكلفة الطلب): {ad['cps']:.2f} &nbsp;|&nbsp; CVR (نسبة التحويل): {ad['cvr']:.2f}%")


def render_noon_ads_expander(sku_up, ent, tot_s, tot_r, tot_o):
    """إعلانات نون في كارت المبيعات: تبويب لأمس (من AdsDaily) وتبويب للشهر (من AdsMonthly) — كل واحد لوحده."""
    yl = calc.ads_merge_by_sku(_ads_yday_for(sku_up))
    ent = calc.ads_merge_by_sku(ent)
    def _tot(lst):
        return (sum(a["spends"] for a in lst), sum(a["revenue"] for a in lst),
                sum(a["orders"] for a in lst), sum(a["clicks"] for a in lst))
    def _tacos_txt(spend, dates, ad_orders=0):
        try:
            D = _ads_reco_compute()
            inp = ((D.get("inp_by_sku") or {}).get(str(sku_up).upper()) or (D.get("inp_by_key") or {}).get(calc.noon_track_key(sku_up))) if D.get("status") == "ok" else None
            if not inp or not inp.get("price"):
                return ""
            units = sum((inp["units_by_date"] or {}).get(d, 0) for d in dates)
            sales = units * inp["price"]
            if sales <= 0:
                return "<br>📉 TACOS (تكلفة الإعلان من المبيعات): — <span style='color:#9ca3af;'>(مفيش مبيعات للـ SKU ده في الفترة في شيت الطلبات)</span>"
            warn = []
            if ad_orders and ad_orders > units:
                warn.append(f"طلبات الإعلان ({ad_orders:,.0f}) أكتر من وحدات الطلبات ({units:,.0f})")
            of = D.get("orders_from")
            if of and dates and of > min(dates):
                warn.append(f"بيانات الطلبات عندنا بتبدأ من {of}")
            w = (f" <span style='color:#f59e0b;'>⚠️ غير دقيق — {' و'.join(warn)}</span>" if warn else "")
            org = max(units - (ad_orders or 0), 0)
            split = f"<br>🧮 من الإعلان: <b>{min(ad_orders or 0, units):,.0f}</b> طلب | عضوي: <b>{org:,.0f}</b> طلب"
            return (f"<br>📉 <b>TACOS (تكلفة الإعلان من المبيعات): {spend / sales * 100:.1f}%</b> "
                    f"<span style='color:#9ca3af;'>(مصروف ÷ مبيعات الـ SKU الكلية ≈ {sales:,.0f} ريال من {units:,.0f} وحدة)</span>{w}" + split)
        except Exception:
            return ""
    _today = datetime.now().date()
    _yday = [_today - timedelta(days=1)]
    _mdays = [_today.replace(day=1) + timedelta(days=i) for i in range(max((_today - _today.replace(day=1)).days, 0))]
    def _show(lst, label):
        _dates = _yday if label == "أمس" else _mdays
        for ad in lst:
            st.markdown(f"🏷️ **الحملات:** {ad['campaign']}<br>{_ads_line(ad)}" + _tacos_txt(ad["spends"], _dates, ad.get("orders", 0)), unsafe_allow_html=True)
    with st.expander("📢 إعلانات"):
        t_y, t_m = st.tabs(["🕒 أمس", "📅 الشهر (من أول الشهر لأمس)"])
        with t_y:
            if yl:
                _show(yl, "أمس")
            else:
                st.caption("مفيش إعلان اشتغل أمس على الـ SKU ده (أو البوت لسه ماسحبش أمس).")
        with t_m:
            _show(ent, "الشهر")


def _ads_reco_get(sku_up):
    try:
        D = _ads_reco_compute()
        if D.get("status") != "ok":
            return None
        return (D.get("by_sku") or {}).get(str(sku_up).upper()) or (D.get("by_key") or {}).get(calc.noon_track_key(sku_up))
    except Exception:
        return None


def _ads_reco_metrics_line(r):
    """TACOS | ROAS | الإعلان من الربح — من آخر 7 أيام."""
    m = r["m"]
    bits = []
    if m.get("tacos") is not None:
        bits.append(f"TACOS <small>(تكلفة الإعلان من المبيعات)</small> <b>{m['tacos'] * 100:.1f}%</b>")
    if m.get("roas") is not None:
        bits.append(f"ROAS <small>(عائد الإعلان)</small> <b>{m['roas']:.2f}</b>")
    if m.get("units") is not None and m.get("ad_orders") is not None:
        _u = m["units"] or 0
        _ao = min(m["ad_orders"] or 0, _u)
        _d = m.get("cov_days")
        _per = f" <small>(مبيعات آخر {_d:.0f} {'أيام' if 3 <= _d <= 10 else 'يوم'})</small>" if _d else ""
        bits.append(f"🛒 إجمالي <b>{_u:,.0f}</b> طلب{_per}: من الإعلان <b>{_ao:,.0f}</b> | بدون إعلان <b>{max(_u - _ao, 0):,.0f}</b>")
    if m.get("ad_share") is not None:
        bits.append(f"الإعلان من الربح <b>{m['ad_share'] * 100:.0f}%</b>")
    conf = {"high": "عالية", "medium": "متوسطة", "low": "منخفضة"}.get(r.get("confidence"), "")
    days = m.get("cov_days")
    if conf:
        bits.append(f"ثقة <b>{conf}</b>" + (f" ({days:.0f} يوم بيانات)" if days is not None else ""))
    return " &nbsp;|&nbsp; ".join(bits)


def ads_reco_inline_html(sku_up):
    """(كبسولة التوصية، سطر الاقتراح والأرقام) — بيتحطوا جنب «الإعلان مربح/غير مربح» في كارت المبيعات."""
    r = _ads_reco_get(sku_up)
    if not r:
        try:
            D = _ads_reco_compute()
            stt = D.get("status")
            if stt == "error":
                why = "تعذر حساب التوصيات: " + D.get("err", "")
            elif stt == "nodata":
                why = "مفيش بيانات تاريخ إعلانات في الشيت (AdsHistory/AdsMonthly) — شغّل البوت واضغط تحديث"
            elif stt == "nokeys":
                why = "مفيش SKU عليه مصروف إعلاني في بيانات التاريخ"
            else:
                _bk = list((D.get("by_key") or {}).keys())
                why = (f"مفيش توصية للـ SKU ده (مفتاح البحث {calc.noon_track_key(sku_up)}) — عدد SKUs اللي ليها توصية: {len(_bk)}"
                       + (f"، أمثلة: {', '.join(_bk[:3])}" if _bk else "") + f" — {D.get('diag', '')} [نسخة {D.get('v')}]")
        except Exception as e:
            why = "خطأ: " + str(e)[:120]
        return "", f'<div dir="rtl" style="font-size:11px;margin-top:4px;color:#9ca3af;">🧠 التوصية: {calc._nt_h(why)}</div>'
    pill = calc._ads_pill(r["label"], r["color"], r["bg"])
    ml = _ads_reco_metrics_line(r)
    line = ('<div dir="rtl" style="font-size:12px;margin-top:4px;color:#e2e8f0;">'
            + (f'<div>📉 {ml} </div>' if ml else "")
            + (f'<div>🔎 {calc._nt_h(r["reasons"][0])}</div>' if r.get("reasons") else "")
            + f'<div>💡 {calc._nt_h(r["suggestion"])}</div></div>')
    return pill, line


def render_ads_reco_chip(sku_up):
    """كبسولة توصية الإعلان جنب صورة المنتج في المبيعات."""
    try:
        D = _ads_reco_compute()
        if D.get("status") == "error":
            st.caption("🧠 توصيات الإعلان: تعذر الحساب — " + D.get("err", ""))
            return
        if D.get("status") == "nodata":
            st.caption("🧠 توصيات الإعلان: مفيش بيانات تاريخ في الشيت لسه (شغّل بوت نون واضغط تحديث).")
            return
        if D.get("status") != "ok":
            return
        r = (D.get("by_sku") or {}).get(str(sku_up).upper()) or (D.get("by_key") or {}).get(calc.noon_track_key(sku_up))
        if not r:
            return
        conf = {"high": "عالية", "medium": "متوسطة", "low": "منخفضة"}[r["confidence"]]
        why = r["reasons"][0] if r["reasons"] else ""
        st.markdown(
            '<div style="direction:rtl;margin:4px 0;">' + calc._ads_pill(r["label"], r["color"], r["bg"]) +
            f' <span style="font-size:12px;color:#9ca3af;">ثقة {conf}</span>'
            f'<div style="font-size:12px;margin-top:3px;">{_ads_reco_metrics_line(r)}</div>'
            f'<div style="font-size:12px;margin-top:3px;">{calc._nt_h(why)}</div>'
            f'<div style="font-size:12px;color:#9ca3af;">{calc._nt_h(r["suggestion"])}</div></div>',
            unsafe_allow_html=True)
    except Exception:
        pass


def _ads_reco_compute(force=False):
    """بيحسب توصيات كل الـ SKUs مرة واحدة ويكاشيها 10 دقايق (بتستخدمها شاشة التوصيات وكارت المبيعات)."""
    c = st.session_state.get("_ads_reco_calc")
    if c and not force and time.time() - c[0] < ADS_RECO_TTL and c[1].get("v") == ADS_RECO_VER:
        return c[1]
    try:
        out = _ads_reco_compute_inner()
    except Exception as e:
        out = {"status": "error", "err": str(e)[:200]}
    out["v"] = ADS_RECO_VER
    st.session_state["_ads_reco_calc"] = (time.time(), out)
    return out


def _ads_reco_compute_inner():
    today = datetime.now().date()
    raw = _ads_reco_load()
    hist, known = calc.ads_hist_parse(raw.get("AdsHistory"))
    mtd, mtd_period = calc.ads_mtd_parse(raw.get("AdsMonthly"))
    if not mtd:        # لو القراءة الجماعية ماجابتش الشهري، نقراه مباشرة من التاب
        try:
            mtd, mtd_period = calc.ads_mtd_parse(ss.worksheet("AdsMonthly").get_all_values())
        except Exception:
            pass
    _diag = (f"AdsHistory: {len(raw.get('AdsHistory') or [])} صف / {len(hist)} SKU · "
             f"AdsMonthly: {len(raw.get('AdsMonthly') or [])} صف / {len(mtd)} SKU "
             f"({sum(1 for v in mtd.values() if v.get('spends', 0) > 0)} بمصروف)")
    if not hist and not mtd:
        return {"status": "nodata", "raw": raw}
    live_m, com_m = get_live_map(), get_com_map()
    nt = get_noon_track() or {}
    apply_noon_our_names(nt)
    key_to_sku = {}
    for src in (live_m, com_m, inv_map):
        for s in list(src.keys()):
            key_to_sku.setdefault(calc.noon_track_key(s), s)
    keys = set(hist.keys()) | set(mtd.keys())
    keys = [k for k in keys if (sum(d.get("spends", 0) for d in hist.get(k, {}).values()) > 0
                                 or (mtd.get(k) or {}).get("spends", 0) > 0)]
    if not keys:
        return {"status": "nokeys", "raw": raw}
    days45 = [today - timedelta(days=i) for i in range(1, 46)]
    try:
        counts = build_daily_orders_counts(days45)
    except Exception:
        counts = {}
    _of = [d for dd in counts.values() for d, n in dd.items() if n]
    orders_from = min(_of) if _of else None
    units_by_key = {}
    for s, dd in counts.items():
        k = calc.noon_track_key(s)
        for d, n in dd.items():
            units_by_key.setdefault(k, {})
            units_by_key[k][d] = units_by_key[k].get(d, 0) + (n or 0)
    asin_map = build_amazon_orders_asin_map()
    inputs, extra = [], {}
    for k in keys:
        sku_up = key_to_sku.get(k)
        price = (live_m.get(sku_up) or {}).get("price") if sku_up else None
        ci = com_m.get(sku_up) if sku_up else None
        unit_profit, cost_missing = None, False
        if price is not None and ci:
            _, net_tax = compute_net_price_after_fees(price, ci)
            cost = resolve_unit_cost(sku_up, (asin_map.get(sku_up, "") or "").upper() or None)
            if cost is None:
                cost_missing, cost = True, 0.0
            unit_profit = net_tax - cost
        stock = (inv_map.get(sku_up) or {}).get("total_stock") if sku_up else None
        comp = None
        e = nt.get(k)
        if e and price:
            try:
                r = calc.noon_track_rivals(e, price)
                cheaper = [o["price"] for o in r["others"] if o["price"] < price]
                box = r["box"]
                if box and box.get("price") and box["price"] < price and not box.get("ours"):
                    cheaper.append(box["price"])
                comp = {"box_ours": (box["ours"] if box else None),
                        "gap_lowest": (min(cheaper) - price) if cheaper else 0,
                        "n_cheaper": len(cheaper),
                        "rating": next((o["rating"] for o in r["others"] if o.get("rating")), None)}
            except Exception:
                comp = None
        inputs.append({"sku": sku_up or k, "price": price, "unit_profit": unit_profit, "cost_missing": cost_missing,
                       "com_missing": ci is None, "units_by_date": units_by_key.get(k, {}),
                       "ads_by_date": hist.get(k, {}), "known_days": known, "ads_mtd": mtd.get(k),
                       "stock": stock, "comp": comp, "today": today})
        extra[inputs[-1]["sku"]] = (k, e)
    results = calc.ads_reco_all(inputs)
    return {"status": "ok", "diag": _diag, "orders_from": orders_from, "raw": raw, "results": results, "inputs": inputs, "known": known,
            "by_sku": {str(r["sku"]).upper(): r for r in results},
            "inp_by_sku": {str(i["sku"]).upper(): i for i in inputs},
            "by_key": {calc.noon_track_key(r["sku"]): r for r in results},
            "inp_by_key": {calc.noon_track_key(i["sku"]): i for i in inputs}}



def render_ads_reco_section():
    """🧠 تحليل ربحية الإعلانات والتوصيات (اقتراحات فقط — مفيش أي تغيير تلقائي في الميزانيات أو الحملات)."""
    st.markdown("---")
    st.markdown("### 🧠 تحليل ربحية الإعلانات والتوصيات")
    st.caption("اقتراحات فقط — النظام مش بيغيّر أي ميزانية أو حملة. القرار ليك. "
               "التقييم على المنتج كله (مبيعات عضوية + إعلان) مش على أرباح الطلبات المنسوبة للإعلان بس.")
    with st.expander("ℹ️ معنى TACOS وROAS وباقي المصطلحات"):
        st.markdown(calc.ADS_GLOSSARY_HTML, unsafe_allow_html=True)
    if st.button("🔄 تحديث بيانات التوصيات", key="ads_reco_refresh"):
        _ads_reco_load(force=True)
        st.session_state.pop("_ads_monthly_rows", None)
        get_noon_track(force=True)
        st.session_state.pop("_ads_reco_calc", None)
        st.session_state.pop("_ads_daily_rows", None)
    D = _ads_reco_compute()
    if D.get("status") == "error":
        st.error("تعذر حساب التوصيات: " + D.get("err", ""))
        return
    if D["status"] == "nodata":
        st.info("لسه مفيش بيانات تاريخ إعلانات في الشيت. شغّل بوت نون (NOONServer / NOONAds) مرة، "
                "هيعمل تابات AdsHistory وAdsMonthly وAdsCampaignHist، وبعدها اضغط تحديث.")
        return
    if D["status"] == "nokeys":
        st.info("مفيش SKU عليه مصروف إعلاني في البيانات الحالية.")
        return
    raw, results, inputs, known = D["raw"], D["results"], D["inputs"], D["known"]
    today = datetime.now().date()
    # ── ملخص ──
    lv = calc.ADS_RECO_LEVELS
    cnt = {}
    for r in results:
        cnt[r["level"]] = cnt.get(r["level"], 0) + 1
    cols = st.columns(len(lv))
    for col, (lk, (lab, colr, bg)) in zip(cols, lv.items()):
        col.markdown(f'<div style="background:{bg};border:1px solid {colr};border-radius:10px;padding:8px;text-align:center;direction:rtl;">'
                     f'<div style="font-size:12px;color:{colr};font-weight:800;">{lab}</div>'
                     f'<div style="font-size:22px;font-weight:800;color:#111827;">{cnt.get(lk, 0)}</div></div>',
                     unsafe_allow_html=True)
    if known:
        st.caption(f"تاريخ يومي متاح: {len(known)} يوم (من {min(known)} إلى {max(known)}). كل ما التاريخ يزيد، الثقة بتزيد.")
    else:
        st.caption("مفيش تاريخ يومي لسه — التوصيات مبنية على لقطة الشهر وثقتها متوسطة كحد أقصى.")

    f1, f2 = st.columns([2, 3])
    sel = f1.multiselect("فلتر التوصية", [v[0] for v in lv.values()], key="ads_reco_lvfilter")
    q = f2.text_input("بحث بالـ SKU", key="ads_reco_q").strip().upper()
    shown = [r for r in results if (not sel or r["label"] in sel) and (not q or q in str(r["sku"]).upper())]
    st.markdown(calc.ads_reco_table_html(shown), unsafe_allow_html=True)
    if len(shown) > 80:
        st.caption(f"معروض أول 80 من {len(shown)} — الباقي في ملف الإكسيل.")
    dl_btn(pd.DataFrame(calc.ads_reco_rows(shown)), "ads_recommendations", key="ads_reco_dl")

    # ── تفاصيل SKU + بدء تجربة تخفيض ──
    st.markdown("#### 🔍 تفاصيل SKU")
    pick = st.selectbox("اختار SKU", [r["sku"] for r in shown] or [""], key="ads_reco_pick")
    cur = next((r for r in shown if r["sku"] == pick), None)
    if cur:
        st.markdown(calc.ads_reco_card_html(cur), unsafe_allow_html=True)
        inp = next(i for i in inputs if i["sku"] == pick)
        if cur["level"] in ("trim", "test_pause", "urgent") and inp.get("unit_profit") is not None:
            with st.expander("🧪 ابدأ تجربة تخفيض تدريجي (بتتسجل للمتابعة بس — إنت اللي بتغيّر الميزانية بإيدك)"):
                c1, c2 = st.columns(2)
                step = c1.number_input("نسبة التخفيض %", 5, 80, 20, 5, key="ads_exp_step")
                every = c2.number_input("نقيّم كل كام يوم", 2, 14, 3, key="ads_exp_every")
                if st.button("📌 سجّل التجربة (البداية النهاردة)", key="ads_exp_go"):
                    m = cur["m"]
                    box_o = (inp.get("comp") or {}).get("box_ours")
                    row = [f"E{datetime.now():%y%m%d%H%M%S}", datetime.now().isoformat(timespec="seconds"), pick,
                           "تخفيض تدريجي", today.isoformat(), round(m["units_pd"], 2),
                           round(m["profit_after_pd"], 2) if m.get("profit_after_pd") is not None else "",
                           round(m["spend_pd"], 2), round(m["tacos"], 4) if m.get("tacos") is not None else "",
                           "نعم" if box_o is True else ("لا" if box_o is False else ""), step, every, "جارية", "", "", ""]
                    if _ads_exp_append(row):
                        _ads_reco_load(force=True)
                        st.success("اتسجلت. غيّر الميزانية بنفسك في لوحة نون، وراجع النتيجة تحت بعد كام يوم.")

    # ── التجارب الجارية ──
    exps = [x for x in calc.ads_exp_parse(raw.get("AdsExperiments")) if x["active"]]
    if exps:
        st.markdown("#### 🧪 التجارب الجارية")
        by_sku = {i["sku"]: i for i in inputs}
        for x in exps:
            inp = by_sku.get(x.get("SKU"))
            if not inp or inp.get("unit_profit") is None:
                st.warning(f"{x.get('SKU')}: مفيش بيانات كفاية لتقييم التجربة.")
                continue
            box_now = (inp.get("comp") or {}).get("box_ours")
            ev = calc.ads_experiment_eval(x, inp["units_by_date"], inp["ads_by_date"], known, inp["price"],
                                          inp["unit_profit"], box_now, today)
            st.markdown(calc.ads_exp_card_html(x, ev), unsafe_allow_html=True)
            if st.button("إقفال التجربة", key="ads_exp_close_" + x["ID"]):
                if _ads_exp_close(x["ID"], " | ".join(ev["reasons"])[:300]):
                    _ads_reco_load(force=True)
                    st.rerun()
    st.caption("🔎 ترتيب الكلمات (عضوي/مدفوع) وظهور إعلانات المنافسين: مش متاح لسه — محتاج بوت ترتيب (المرحلة 2). "
               "التوصيات مش بتعتمد عليه، وبيتكتب لك في كل كارت إنه 'لسه محتاج اختبار'.")


def _render_ads_performance_tab():
    """تاب "الإعلانات" المستقل — منقول بالكامل برة داشبورد المبيعات، وبيتعرض مرة واحدة
    بس (مش 3 مرات زي الأول جوه تابات الكل/FBN/FBB) — لأن بيانات الإعلانات (Views/
    Clicks/Orders/Spends/Revenue) جايه من شيت الإعلانات نفسه واللي مفيهوش عمود
    Fulfillment Model أصلاً، يعني مفيش طريقة تقنية نقسّمها على FBN/FBB. الأرقام هنا
    إجمالية لكل SKU/حملة، زي ما نون بيوريها بالظبط | The standalone "Ads" tab — moved
    entirely out of the sales dashboard and rendered once (not 3x inside the All/FBN/
    FBB sub-tabs) — because ad data (Views/Clicks/Orders/Spends/Revenue) comes from
    the Advertisements sheet, which has no Fulfillment Model column at all, so there's
    no way to split it by FBN/FBB. Numbers here are the SKU/campaign totals as-is,
    exactly as Noon reports them."""
    live_map_dash = get_live_map()

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

    # ── ربحية الإعلانات لكل SKU (نفس منطق تاب المبيعات: صافي الربح الفعلي بعد خصم
    #    تكلفة المنتج من طلبات الإعلان، مقابل إجمالي المصروف عليه) — عشان تظهر في
    #    التنبيهات السريعة تحت | Same logic as the Sales tab: actual net profit
    #    after subtracting product cost from ad orders, vs total ad spend ──
    ads_map_dash = get_ads_map()
    com_map_dash = get_com_map()

    _ads_fallback_dates = [datetime.now().date() - timedelta(days=i) for i in range(1, 91)]
    _ads_fallback_prices = build_daily_orders_prices(_ads_fallback_dates)

    def _ads_latest_price_for_sku(sku_up):
        """السعر الأساسي لهذا الـ SKU: عمود sale_price من تاب LIVE أولاً (سعر البيع الأساسي)،
        ولو مش موجود بيدوّر في أسعار آخر 90 يوم من الطلبات (كل الطلبات، غير مقسومة FBN/FBB
        لأن الإعلانات نفسها مش مقسومة كده) كبديل (سعر العرض) | Base price for this SKU:
        LIVE sheet's sale_price column first, falling back to the last 90 days of order
        prices (all orders, unsplit — ads data itself isn't split by fulfillment type)."""
        live_info = live_map_dash.get(sku_up)
        if live_info and live_info.get("price") is not None:
            return live_info["price"]
        vals = []
        for d in _ads_fallback_dates:
            for p, qty in _ads_fallback_prices.get(sku_up, {}).get(d, []):
                if p and str(p).strip().lower() not in ("", "nan", "none"):
                    try:
                        vals.append((float(str(p).replace(",", "")), qty))
                    except Exception:
                        pass
        if vals:
            vals.sort(key=lambda x: -x[1])
            return vals[0][0]
        return None

    ads_profit_rows_ap, ads_loss_rows_ap = [], []
    _links_map_ads_orphan = get_links_map()
    _sku_to_asin_ads = build_amazon_orders_asin_map()
    for sku_up_ad, ads_entries_ad in ads_map_dash.items():
        com_info_ad = com_map_dash.get(sku_up_ad)
        if not ads_entries_ad or not com_info_ad:
            continue
        latest_price_ad = _ads_latest_price_for_sku(sku_up_ad)
        if latest_price_ad is None:
            continue
        _, net_tax_ad = compute_net_price_after_fees(latest_price_ad, com_info_ad)
        # ── تكلفة الوحدة من تاب التكويد (نون أو أمازون حسب المنتج) وصافي
        #    الربح الفعلي بعدها، زي بالظبط منطق تاب المبيعات — لو التكلفة مش
        #    مسجلة بتتحط صفر مع تنبيه، بدل ما تختفي بصمت من الحساب | Unit cost
        #    from the right Tacweed sheet (Noon or Amazon, depending on the
        #    product) and the actual net profit after subtracting it — same
        #    logic as the Sales tab. Missing cost defaults to zero with a flag,
        #    instead of silently dropping out of the calculation ──
        asin_for_ad = _sku_to_asin_ads.get(sku_up_ad, "").upper() or None
        cost_ad = resolve_unit_cost(sku_up_ad, asin_for_ad)
        cost_missing_ad = cost_ad is None
        if cost_missing_ad:
            cost_ad = 0.0
        net_profit_ad = net_tax_ad - cost_ad
        total_spends_ad = sum(a["spends"] for a in ads_entries_ad)
        total_orders_ad = sum(a["orders"] for a in ads_entries_ad)
        total_net_ad = total_orders_ad * net_profit_ad
        result_ad = total_net_ad - total_spends_ad
        inv_info_ad = inv_map.get(sku_up_ad, {})
        # الصورة من المخزون لو الـ SKU موجود فيه، وإلا من شيت links n مباشرة |
        # Image from Inventory if the SKU is there, otherwise straight from
        # the links n sheet
        img_ad = inv_info_ad.get("img", "") or _links_map_ads_orphan.get(sku_up_ad, "")
        entry_ad = {"sku_up": sku_up_ad, "sku": inv_info_ad.get("sku", sku_up_ad),
                    "img": img_ad,
                    "spends": total_spends_ad, "orders": total_orders_ad,
                    "net_total": total_net_ad, "result": result_ad,
                    "cost": cost_ad, "cost_missing": cost_missing_ad}
        if total_orders_ad <= 0 or result_ad < 0:
            ads_loss_rows_ap.append(entry_ad)
        else:
            ads_profit_rows_ap.append(entry_ad)
    ads_profit_rows_ap.sort(key=lambda r: -r["result"])
    ads_loss_rows_ap.sort(key=lambda r: r["result"])  # الأكثر خسارة أولاً

    st.markdown("##### 🔔 تنبيهات الإعلانات | Ads Alerts")
    acp1, acp2 = st.columns(2)
    with acp1:
        st.markdown(
            f'<div style="background:#f0fdf4;border:1px solid #22c55e;border-right:4px solid #22c55e;'
            f'border-radius:12px;padding:12px 14px;direction:rtl;min-height:92px;">'
            f'<div style="font-size:12px;color:#374151;margin-bottom:6px;">🎯 منتجات ربحانة من الإعلانات</div>'
            f'<div style="font-size:22px;font-weight:800;color:#111827;">{len(ads_profit_rows_ap):,}</div>'
            f'<div style="font-size:11px;color:#6b7280;">صافي ربح الطلبات &gt; المصروف على الإعلان</div></div>',
            unsafe_allow_html=True)
    with acp2:
        st.markdown(
            f'<div style="background:#fef2f2;border:1px solid #ef4444;border-right:4px solid #ef4444;'
            f'border-radius:12px;padding:12px 14px;direction:rtl;min-height:92px;">'
            f'<div style="font-size:12px;color:#374151;margin-bottom:6px;">🚨 منتجات خسرانة من الإعلانات</div>'
            f'<div style="font-size:22px;font-weight:800;color:#111827;">{len(ads_loss_rows_ap):,}</div>'
            f'<div style="font-size:11px;color:#6b7280;">المصروف على الإعلان أكبر من صافي الربح (أو من غير طلبات)</div></div>',
            unsafe_allow_html=True)

    def _render_ads_alert_sku_row(r, badges_html=""):
        ci_al, cinfo_al = st.columns([1, 6])
        with ci_al:
            show_img(r["img"], 55)
        with cinfo_al:
            st.markdown(f"{sku_link_html(r['sku'])}", unsafe_allow_html=True)
            if badges_html:
                st.markdown(badges_html, unsafe_allow_html=True)

    def _ads_insight_html(r):
        # ── سطر التكلفة + تنبيه لو غير مسجلة | Cost line + warning if unrecorded ──
        cost_note_ad = f' &nbsp;|&nbsp; 📦 التكلفة | Cost: <b>{r["cost"]:,.2f}</b> ريال'
        missing_note_ad = ""
        if r.get("cost_missing"):
            missing_note_ad = (' <span style="color:#fbbf24;">⚠️ تكلفة المنتج غير مسجلة '
                                '(تاب التكويد) — اعتُبرت صفر</span>')
        if r["orders"] <= 0:
            return (f'<span style="color:#f87171;font-size:12px;font-weight:700;">🚨 مفدتش لحد دلوقتي: '
                    f'اتصرف {r["spends"]:,.2f} ريال ولسه ما جابش أي طلبات فعلية</span>'
                    f'{cost_note_ad}{missing_note_ad}')
        if r["result"] >= 0:
            return (f'<span style="color:#4ade80;font-size:12px;font-weight:700;">🎯 الاعلان مربح: '
                    f'عدد طلبات الاعلان {r["orders"]:,.0f} طلب بصافي ربح فعلي بعد التكلفة إجمالي {r["net_total"]:,.2f} ريال مقابل '
                    f'{r["spends"]:,.2f} ريال مدفوع — حقق {r["result"]:,.2f} ريال 👌</span>'
                    f'{cost_note_ad}{missing_note_ad}')
        return (f'<span style="color:#f87171;font-size:12px;font-weight:700;">🚨 الاعلان غير مربح: '
                f'مدفوع {r["spends"]:,.2f} ريال، لكن صافي الربح الفعلي بعد التكلفة من {r["orders"]:,.0f} طلب بس {r["net_total"]:,.2f} ريال — '
                f'خسران {abs(r["result"]):,.2f} ريال إجمالي</span>'
                f'{cost_note_ad}{missing_note_ad}')

    with st.expander(f"🎯 عرض منتجات ربحانة من الإعلانات ({len(ads_profit_rows_ap):,}) | Show profitable-ads SKUs"):
        if ads_profit_rows_ap:
            df_al5 = pd.DataFrame([{
                "SKU": r["sku"], "طلبات الإعلان | Ad Orders": r["orders"],
                "المصروف | Spends": round(r["spends"], 2),
                "التكلفة | Cost": round(r["cost"], 2),
                "صافي الربح بعد التكلفة | Net Total After Cost": round(r["net_total"], 2),
                "النتيجة | Result": round(r["result"], 2),
            } for r in ads_profit_rows_ap])
            dl_btn(df_al5, "alert_ads_profit", key="dl_alert_ads_profit_ap")
            for r in ads_profit_rows_ap:
                _render_ads_alert_sku_row(r, badges_html=_ads_insight_html(r))
        else:
            st.caption("لا توجد أصناف ربحانة من الإعلانات حالياً")

    with st.expander(f"🚨 عرض منتجات خسرانة من الإعلانات ({len(ads_loss_rows_ap):,}) | Show losing-ads SKUs"):
        if ads_loss_rows_ap:
            df_al6 = pd.DataFrame([{
                "SKU": r["sku"], "طلبات الإعلان | Ad Orders": r["orders"],
                "المصروف | Spends": round(r["spends"], 2),
                "التكلفة | Cost": round(r["cost"], 2),
                "صافي الربح بعد التكلفة | Net Total After Cost": round(r["net_total"], 2),
                "النتيجة | Result": round(r["result"], 2),
            } for r in ads_loss_rows_ap])
            dl_btn(df_al6, "alert_ads_loss", key="dl_alert_ads_loss_ap")
            for r in ads_loss_rows_ap:
                _render_ads_alert_sku_row(r, badges_html=_ads_insight_html(r))
        else:
            st.caption("لا توجد أصناف خسرانة من الإعلانات حالياً 🎉")

    st.divider()

    # ══════════════════════════════════════════════════════════════════════
    # 📢 تحليل أداء الإعلانات (على مستوى الحملة) | Ads Performance Analysis
    # (Campaign level) — ده معتمد بس على بيانات الإعلانات الموجودة فعليًا
    # (Views/Clicks/Orders/ATC/Spends/Revenue) من غير أي افتراض لتكلفة المنتج أو
    # هامش ربح حقيقي. كل حساب هنا بيتجمّع من *كل* الحملات/الصفوف الخاصة بالـ SKU
    # أو الحملة قبل ما يتحسب — مش بياخد رقم من صف واحد بس لو فيه أكتر من حملة |
    # Based only on ad data that actually exists, no product-cost or profit-margin
    # assumptions. Every number here is summed across *all* matching campaign rows
    # before any ratio is computed — never taken from a single row when more than
    # one campaign exists.
    st.markdown("---")
    st.markdown("## 📢 تحليل أداء الإعلانات | Ads Performance Analysis")
    st.caption("مبني فقط على بيانات الإعلانات الموجودة في النظام حاليًا — بدون أي افتراض لتكلفة المنتج أو هامش الربح | Based only on ad data currently in the system — no product cost or profit margin assumptions")

    def _apa_ratios(views, clicks, atc, orders, spends, revenue):
        """يعيد حساب كل النسب من الأرقام الخام المجمّعة (مش من عمود جاهز في صف واحد)
        عشان أي SKU/حملة ليها أكتر من صف تتحسب صح | Recomputes every ratio from the
        summed raw totals (never from a single pre-computed column), so multi-row
        SKUs/campaigns are calculated correctly."""
        ctr = (clicks / views * 100) if views > 0 else 0.0
        cpc = (spends / clicks) if clicks > 0 else 0.0
        cpa = (spends / orders) if orders > 0 else 0.0
        cvr = (orders / clicks * 100) if clicks > 0 else 0.0
        roas = (revenue / spends) if spends > 0 else 0.0
        click_to_atc = (atc / clicks * 100) if clicks > 0 else 0.0
        atc_to_order = (orders / atc * 100) if atc > 0 else 0.0
        return {"ctr": ctr, "cpc": cpc, "cpa": cpa, "cvr": cvr, "roas": roas,
                "click_to_atc": click_to_atc, "atc_to_order": atc_to_order}

    # ── تجميع كل صفوف الإعلانات (Sku × Campaign) على مستوى الحملة نفسها —
    #    عشان أي حملة بتستهدف أكتر من SKU تتحسب مجمّعة صح ومتاخدش من صف واحد ──
    campaigns_apa = {}
    for sku_up_c, entries_c in ads_map_dash.items():
        for e in entries_c:
            cname_c = e["campaign"] or "—"
            agg_c = campaigns_apa.setdefault(cname_c, {
                "campaign": cname_c, "views": 0.0, "clicks": 0.0, "orders": 0.0,
                "atc": 0.0, "spends": 0.0, "revenue": 0.0, "skus": set(),
            })
            agg_c["views"]   += e["views"]
            agg_c["clicks"]  += e["clicks"]
            agg_c["orders"]  += e["orders"]
            agg_c["atc"]     += e["atc"]
            agg_c["spends"]  += e["spends"]
            agg_c["revenue"] += e["revenue"]
            agg_c["skus"].add(sku_up_c)

    for _cname, _agg in campaigns_apa.items():
        _agg.update(_apa_ratios(_agg["views"], _agg["clicks"], _agg["atc"], _agg["orders"], _agg["spends"], _agg["revenue"]))
        _agg["sku_count"] = len(_agg["skus"])
        # نتيجة الإعلان بعد الإنفاق الإعلاني فقط — مش ربح حقيقي (مفيش تكلفة منتج) |
        # Ad result after ad spend only — not real profit (no product cost known)
        _agg["ad_result"] = _agg["revenue"] - _agg["spends"]

    campaigns_list_apa = list(campaigns_apa.values())

    def _apa_render_related_skus(sku_up_set, max_show=6):
        """يعرض الأصناف (SKU) المرتبطة بالحملة مع صورها | Renders the SKUs linked
        to this campaign, each with its product image."""
        sku_list_r = sorted(sku_up_set)
        shown_r = sku_list_r[:max_show]
        _links_map_apa_orphan = get_links_map()
        for sku_up_r in shown_r:
            inv_info_r = inv_map.get(sku_up_r, {})
            # الصورة من المخزون لو الـ SKU موجود فيه، وإلا من شيت links n مباشرة |
            # Image from Inventory if the SKU is there, otherwise straight
            # from the links n sheet
            img_r = inv_info_r.get("img", "") or _links_map_apa_orphan.get(sku_up_r, "")
            ci_r, cinfo_r = st.columns([1, 6])
            with ci_r:
                show_img(img_r, 45)
            with cinfo_r:
                st.markdown(sku_link_html(inv_info_r.get("sku", sku_up_r)), unsafe_allow_html=True)
        if len(sku_list_r) > max_show:
            st.caption(f"+ {len(sku_list_r) - max_show} SKU إضافي | more SKUs")

    if not campaigns_list_apa:
        st.info("لا توجد بيانات إعلانات مرفوعة حالياً | No ad data uploaded yet")
    else:
        # ── 1) مؤشرات أداء الإعلانات | Ad Performance Metrics (إجمالي كل الحملات) ──
        tot_views_apa   = sum(c["views"] for c in campaigns_list_apa)
        tot_clicks_apa  = sum(c["clicks"] for c in campaigns_list_apa)
        tot_atc_apa     = sum(c["atc"] for c in campaigns_list_apa)
        tot_orders_apa  = sum(c["orders"] for c in campaigns_list_apa)
        tot_spends_apa  = sum(c["spends"] for c in campaigns_list_apa)
        tot_revenue_apa = sum(c["revenue"] for c in campaigns_list_apa)
        tot_ratios_apa  = _apa_ratios(tot_views_apa, tot_clicks_apa, tot_atc_apa, tot_orders_apa, tot_spends_apa, tot_revenue_apa)

        st.markdown("#### 📊 مؤشرات أداء الإعلانات | Ad Performance Metrics")
        st.caption("ℹ️ كل النسب (CTR/CPC/CPS/CVR/ROAS) بتتحسب من إجمالي الأرقام الخام لكل الحملات مجمّعة — مش من عمود جاهز في صف واحد | All ratios are computed from the raw totals across every campaign combined — never from a single pre-computed column")
        mrow1 = st.columns(4)
        with mrow1[0]:
            st.markdown(_kpi_card_html("🛒", "#2563eb", "الطلبات | Orders", f"{tot_orders_apa:,.0f}"), unsafe_allow_html=True)
        with mrow1[1]:
            st.markdown(_kpi_card_html("👁️", "#0891b2", "مرات الظهور | Impressions", f"{tot_views_apa:,.0f}"), unsafe_allow_html=True)
        with mrow1[2]:
            st.markdown(_kpi_card_html("🖱️", "#7c3aed", "النقرات | Clicks", f"{tot_clicks_apa:,.0f}"), unsafe_allow_html=True)
        with mrow1[3]:
            st.markdown(_kpi_card_html("➕", "#059669", "الإضافة إلى السلة | Add to Cart", f"{tot_atc_apa:,.0f}"), unsafe_allow_html=True)
        mrow2 = st.columns(4)
        with mrow2[0]:
            st.markdown(_kpi_card_html("📈", "#0891b2", "معدل النقر | CTR", f"{tot_ratios_apa['ctr']:.2f}%"), unsafe_allow_html=True)
        with mrow2[1]:
            st.markdown(_kpi_card_html("💵", "#f59e0b", "تكلفة النقرة | CPC", f"{tot_ratios_apa['cpc']:.2f} ريال"), unsafe_allow_html=True)
        with mrow2[2]:
            st.markdown(_kpi_card_html("🎯", "#dc2626", "تكلفة الطلب | CPS / CPA", f"{tot_ratios_apa['cpa']:.2f} ريال"), unsafe_allow_html=True)
        with mrow2[3]:
            st.markdown(_kpi_card_html("📊", "#9333ea", "معدل التحويل | CVR", f"{tot_ratios_apa['cvr']:.2f}%"), unsafe_allow_html=True)
        mrow3 = st.columns(3)
        with mrow3[0]:
            st.markdown(_kpi_card_html("🎯", "#16a34a", "العائد على الإنفاق الإعلاني | ROAS", f"{tot_ratios_apa['roas']:.2f}"), unsafe_allow_html=True)
        with mrow3[1]:
            st.markdown(_kpi_card_html("💰", "#16a34a", "الإيراد | Revenue", f"{tot_revenue_apa:,.2f} ريال"), unsafe_allow_html=True)
        with mrow3[2]:
            st.markdown(_kpi_card_html("💸", "#dc2626", "الإنفاق الإعلاني | Ad Spend", f"{tot_spends_apa:,.2f} ريال"), unsafe_allow_html=True)

        st.write("")

        # ── 2) تحليل مسار الإعلان | Advertising Funnel ──
        st.markdown("#### 🔻 تحليل مسار الإعلان | Advertising Funnel")
        st.caption("عشان تعرف أين يحدث انخفاض الأداء في مسار الإعلان | See exactly where performance drops along the funnel")
        fcols_apa = st.columns(4)
        funnel_stages_apa = [
            ("👁️ مرات الظهور | Impressions", tot_views_apa, None),
            ("🖱️ النقرات | Clicks", tot_clicks_apa, tot_ratios_apa["ctr"]),
            ("➕ الإضافة إلى السلة | Add to Cart", tot_atc_apa, tot_ratios_apa["click_to_atc"]),
            ("🛒 الطلبات | Orders", tot_orders_apa, tot_ratios_apa["atc_to_order"]),
        ]
        for fc_apa, (label_f, val_f, rate_f) in zip(fcols_apa, funnel_stages_apa):
            with fc_apa:
                rate_html_f = (f'<div style="font-size:11px;color:#f59e0b;margin-top:4px;">↓ {rate_f:.1f}%</div>'
                               if rate_f is not None else "")
                st.markdown(
                    f'<div style="background:#1e293b;border:1px solid #334155;border-radius:10px;'
                    f'padding:12px 10px;text-align:center;">'
                    f'<div style="font-size:11px;color:#94a3b8;">{label_f}</div>'
                    f'<div style="font-size:20px;font-weight:800;color:#e2e8f0;">{val_f:,.0f}</div>'
                    f'{rate_html_f}</div>', unsafe_allow_html=True)
        st.markdown(
            f'<div style="margin-top:8px;font-size:12px;color:#94a3b8;">'
            f'📈 معدل النقر | CTR: <b style="color:#e2e8f0;">{tot_ratios_apa["ctr"]:.2f}%</b> &nbsp;|&nbsp; '
            f'النقر → السلة | Click → Cart: <b style="color:#e2e8f0;">{tot_ratios_apa["click_to_atc"]:.1f}%</b> &nbsp;|&nbsp; '
            f'السلة → الطلب | Cart → Order: <b style="color:#e2e8f0;">{tot_ratios_apa["atc_to_order"]:.1f}%</b> &nbsp;|&nbsp; '
            f'معدل التحويل الكلي | CVR: <b style="color:#e2e8f0;">{tot_ratios_apa["cvr"]:.2f}%</b>'
            f'</div>', unsafe_allow_html=True)

        st.write("")

        # ── دوال التحليل التلقائي / التصنيف / التوصية — مبنية على أكتر من مؤشر مع
        #    بعض (ROAS+CPA+CVR+CTR+CPC+Orders+Spend+Revenue) مش مؤشر واحد بس ──
        def _apa_insight(c):
            if c["orders"] <= 0:
                return ("🔴", "أداء ضعيف | Poor Performance",
                        f"اتصرف {c['spends']:,.2f} ريال على الحملة ولسه ما جابتش أي طلبات فعلية.")
            score_c = 0
            if c["roas"] >= 3: score_c += 2
            elif c["roas"] >= 1.5: score_c += 1
            elif c["roas"] < 1: score_c -= 2
            if c["cvr"] >= 3: score_c += 1
            elif c["cvr"] < 1: score_c -= 1
            if c["ctr"] >= 1: score_c += 1
            elif c["ctr"] < 0.3: score_c -= 1
            if c["spends"] > 0 and c["revenue"] < c["spends"]:
                score_c -= 2
            if score_c >= 3:
                return ("🟢", "أداء جيد | Good Performance",
                        f"الحملة تحقق ROAS {c['roas']:.2f} مع معدل تحويل {c['cvr']:.2f}% جيد.")
            elif score_c >= 0:
                return ("🟡", "يحتاج إلى تحسين | Needs Improvement",
                        f"الحملة بتاخد نقرات معقولة (CTR {c['ctr']:.2f}%)، لكن التحويل للطلبات ({c['cvr']:.2f}%) أو الـ ROAS ({c['roas']:.2f}) لسه محتاج تحسين.")
            else:
                return ("🔴", "أداء ضعيف | Poor Performance",
                        f"تكلفة الطلب {c['cpa']:,.2f} ريال مرتفعة مقارنة بعدد الطلبات ({c['orders']:,.0f}) والـ ROAS {c['roas']:.2f}.")

        def _apa_classification(c):
            icon_i, _t, _d = _apa_insight(c)
            if c["orders"] <= 0:
                return "🔴", "أداء ضعيف | Poor Performance"
            if icon_i == "🟢":
                return "🟢", "أداء قوي | Strong Performance"
            if icon_i == "🟡":
                if c["roas"] < 1.5 and c["cvr"] < 2:
                    return "🟠", "يحتاج إلى تحسين | Needs Improvement"
                return "🟡", "يحتاج إلى مراقبة | Needs Monitoring"
            return "🔴", "أداء ضعيف | Poor Performance"

        def _apa_recommendation(c):
            cls_icon_c, _l = _apa_classification(c)
            if c["orders"] <= 0 and c["spends"] > 0:
                return "قلل الإنفاق | Reduce Spend"
            if cls_icon_c == "🟢":
                return "استمر | Continue"
            if cls_icon_c == "🟡":
                return "راقب | Monitor"
            if cls_icon_c == "🟠":
                return "حسّن الحملة | Optimize"
            return "راجع الحملة | Review"

        # ── 5) مقارنة الحملات | Campaign Comparison (قبل التفاصيل عشان تبان الأهم فوق) ──
        st.markdown("#### 🏆 مقارنة الحملات | Campaign Comparison")
        camps_with_orders_apa = [c for c in campaigns_list_apa if c["orders"] > 0]
        camps_with_clicks_apa = [c for c in campaigns_list_apa if c["clicks"] > 0]
        comp_specs_apa = [
            ("🏆", "أفضل حملة حسب ROAS | Best by ROAS", camps_with_orders_apa, lambda c: c["roas"], lambda c: f"ROAS {c['roas']:.2f}"),
            ("💰", "أعلى إيراد | Highest Revenue", campaigns_list_apa, lambda c: c["revenue"], lambda c: f"{c['revenue']:,.2f} ريال"),
            ("🛒", "أكثر طلبات | Most Orders", campaigns_list_apa, lambda c: c["orders"], lambda c: f"{c['orders']:,.0f} طلب"),
            ("💸", "أعلى إنفاق إعلاني | Highest Ad Spend", campaigns_list_apa, lambda c: c["spends"], lambda c: f"{c['spends']:,.2f} ريال"),
            ("🎯", "أفضل تكلفة طلب | Best CPA", camps_with_orders_apa, lambda c: -c["cpa"], lambda c: f"{c['cpa']:.2f} ريال"),
            ("📈", "أفضل معدل تحويل | Best CVR", camps_with_clicks_apa, lambda c: c["cvr"], lambda c: f"{c['cvr']:.2f}%"),
            ("👁️", "أفضل معدل نقر | Best CTR", campaigns_list_apa, lambda c: c["ctr"], lambda c: f"{c['ctr']:.2f}%"),
        ]
        comp_cols_apa = st.columns(2)
        for i_apa, (icon_s, label_s, pool_s, key_s, fmt_s) in enumerate(comp_specs_apa):
            best_c_apa = max(pool_s, key=key_s, default=None)
            with comp_cols_apa[i_apa % 2]:
                if best_c_apa:
                    st.markdown(_kpi_card_html(icon_s, "#2563eb", label_s, best_c_apa["campaign"]), unsafe_allow_html=True)
                    st.caption(fmt_s(best_c_apa))
                    with st.expander(f"🏷️ الأصناف | SKUs ({best_c_apa['sku_count']})"):
                        _apa_render_related_skus(best_c_apa["skus"])
                else:
                    st.markdown(_kpi_card_html(icon_s, "#6b7280", label_s, "—"), unsafe_allow_html=True)

        st.write("")

        # ── 3+6+7) التحليل التلقائي + التصنيف + التوصية لكل حملة | Automatic
        #    insight + classification + recommendation per campaign ──
        st.markdown("#### 🔎 تحليل كل حملة | Per-Campaign Analysis")
        for c_apa in sorted(campaigns_list_apa, key=lambda x: -x["spends"]):
            icon_i, title_i, desc_i = _apa_insight(c_apa)
            icon_c, label_c = _apa_classification(c_apa)
            rec_c = _apa_recommendation(c_apa)
            bg_i = {"🟢": "#052e1655", "🟡": "#78350f33", "🔴": "#4c051655"}[icon_i]
            border_i = {"🟢": "#16a34a", "🟡": "#f59e0b", "🔴": "#dc2626"}[icon_i]
            with st.expander(f"{icon_i} {c_apa['campaign']} — {c_apa['sku_count']} SKU | {c_apa['orders']:,.0f} طلب"):
                st.markdown(
                    f'<div dir="rtl" style="background:{bg_i};border:1px solid {border_i};border-radius:8px;padding:8px 12px;margin-bottom:8px;">'
                    f'<b>{icon_i} {title_i}</b><br><span style="font-size:13px;">{desc_i}</span>'
                    f'</div>', unsafe_allow_html=True)
                st.markdown(
                    f"👁️ ظهور: {c_apa['views']:,.0f} &nbsp;|&nbsp; 🖱️ نقرات: {c_apa['clicks']:,.0f} &nbsp;|&nbsp; "
                    f"➕ سلة: {c_apa['atc']:,.0f} &nbsp;|&nbsp; 🛒 طلبات: {c_apa['orders']:,.0f}<br>"
                    f"📊 CTR: {c_apa['ctr']:.2f}% &nbsp;|&nbsp; 💵 CPC: {c_apa['cpc']:.2f} &nbsp;|&nbsp; "
                    f"🎯 CPS/CPA: {c_apa['cpa']:.2f} &nbsp;|&nbsp; 📈 CVR: {c_apa['cvr']:.2f}% &nbsp;|&nbsp; 🎯 ROAS: {c_apa['roas']:.2f}<br>"
                    f"💸 إنفاق: {c_apa['spends']:,.2f} ريال &nbsp;|&nbsp; 💰 إيراد: {c_apa['revenue']:,.2f} ريال &nbsp;|&nbsp; "
                    f"📉 نتيجة الإعلان بعد الإنفاق الإعلاني | Ad Result After Ad Spend: <b>{c_apa['ad_result']:,.2f} ريال</b>")
                st.markdown(f"🏷️ التصنيف | Classification: **{icon_c} {label_c}**")
                st.markdown(f"✅ التوصية | Recommendation: **{rec_c}**")
                st.markdown("🏷️ **الأصناف المرتبطة | Related SKUs**")
                _apa_render_related_skus(c_apa["skus"])

        st.caption(
            "ℹ️ مقارنة الفترة الحالية بالفترة السابقة (📈/📉 Revenue, Orders, Ad Spend, ROAS, CPA, CTR, CVR, "
            "Clicks, Add to Cart) مش متاحة هنا لسه — لأن تاب الإعلانات بيحفظ إجمالي كل حملة لحظيًا من غير "
            "تاريخ يومي، فمفيش فترة سابقة نقارن بيها. لو حبينا نفعّلها، محتاجين نبدأ نسجّل نسخة/تاريخ لكل "
            "تحديث في شيت الإعلانات | Current-vs-previous-period comparison isn't available yet because the "
            "Advertisements sheet only stores each campaign's live cumulative totals, with no daily date "
            "history to compare against. Enabling it would require snapshotting the ads sheet with dates.")

    render_ads_reco_section()


def _render_ads_performance_tab_amazon():
    """نسخة أمازون من تاب "الإعلانات" المستقل — نفس بالظبط منطق وتصميم
    _render_ads_performance_tab (تنبيهات ربح/خسارة + تحليل أداء الحملات)، بس
    بيانات الإعلانات هنا جايه من شيت AdvertisementsAmz ومفتاحها ASIN (معرّف
    منتج أمازون الحقيقي) بدل SKU، زي بالظبط ما بيانات نون بتُعرض في تابها |
    Amazon version of the standalone "Ads" tab — identical logic and layout to
    _render_ads_performance_tab (profit/loss alerts + campaign performance
    analysis), but ad data here comes from the AdvertisementsAmz sheet and is
    keyed by ASIN (Amazon's real product identity) instead of SKU, shown next
    to Noon's tab exactly the same way."""
    live_map_dash = get_live_map()

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

    # ── ربحية الإعلانات لكل SKU (نفس منطق تاب المبيعات: صافي الربح الفعلي بعد خصم
    #    تكلفة المنتج من طلبات الإعلان، مقابل إجمالي المصروف عليه) — عشان تظهر في
    #    التنبيهات السريعة تحت | Same logic as the Sales tab: actual net profit
    #    after subtracting product cost from ad orders, vs total ad spend ──
    ads_map_dash = get_ads_map_amazon()
    com_map_dash = get_com_map()

    _ads_fallback_dates = [datetime.now().date() - timedelta(days=i) for i in range(1, 91)]
    _ads_fallback_prices = build_daily_orders_prices_amazon_all_asin(_ads_fallback_dates)

    def _ads_latest_price_for_asin(asin_up, primary_sku_up):
        """السعر الأساسي لهذا الـ ASIN: عمود sale_price من تاب LIVE أولاً (لو المنتج
        موجود فيه تحت الـ MSKU الأساسي بتاعه)، ولو مش موجود بيدوّر في أسعار آخر 90
        يوم من طلبات أمازون (مجمّعة على ASIN، كل الطلبات، غير مقسومة Normal/FBA
        لأن الإعلانات نفسها مش مقسومة كده) كبديل | Base price for this ASIN: the
        LIVE sheet's sale_price column first (if the product exists there under
        its primary MSKU), falling back to the last 90 days of Amazon order
        prices (grouped by ASIN, all orders, unsplit — ads data itself isn't
        split by fulfillment type)."""
        live_info = live_map_dash.get(primary_sku_up)
        if live_info and live_info.get("price") is not None:
            return live_info["price"]
        vals = []
        for d in _ads_fallback_dates:
            for p, qty in _ads_fallback_prices.get(asin_up, {}).get(d, []):
                if p and str(p).strip().lower() not in ("", "nan", "none"):
                    try:
                        vals.append((float(str(p).replace(",", "")), qty))
                    except Exception:
                        pass
        if vals:
            vals.sort(key=lambda x: -x[1])
            return vals[0][0]
        return None

    ads_profit_rows_ap, ads_loss_rows_ap = [], []
    _links_map_ads_orphan = get_links_map_asin()
    for asin_up_ad, ads_entries_ad in ads_map_dash.items():
        asin_info_ad = amz_inv_map_asin.get(asin_up_ad, {})
        mskus_ad = list(asin_info_ad.get("skus", []))
        primary_sku_ad = mskus_ad[0].upper() if mskus_ad else asin_up_ad
        com_info_ad = com_map_dash.get(primary_sku_ad)
        if not ads_entries_ad or not com_info_ad:
            continue
        latest_price_ad = _ads_latest_price_for_asin(asin_up_ad, primary_sku_ad)
        if latest_price_ad is None:
            continue
        _, net_tax_ad = compute_net_price_after_fees(latest_price_ad, com_info_ad)
        # ── تكلفة الوحدة من تاب التكويد الخاص بأمازون (TacweedAmazon) عن طريق
        #    الـ ASIN، وصافي الربح الفعلي بعدها، زي بالظبط منطق تاب المبيعات —
        #    لو التكلفة مش مسجلة بتتحط صفر مع تنبيه، بدل ما تختفي بصمت من
        #    الحساب | Unit cost from the Amazon Tacweed sheet (TacweedAmazon)
        #    via the ASIN, and the actual net profit after subtracting it —
        #    same logic as the Sales tab. Missing cost defaults to zero with a
        #    flag, instead of silently dropping out of the calculation ──
        cost_ad = resolve_unit_cost(primary_sku_ad, asin_up_ad)
        cost_missing_ad = cost_ad is None
        if cost_missing_ad:
            cost_ad = 0.0
        net_profit_ad = net_tax_ad - cost_ad
        total_spends_ad = sum(a["spends"] for a in ads_entries_ad)
        total_orders_ad = sum(a["orders"] for a in ads_entries_ad)
        total_net_ad = total_orders_ad * net_profit_ad
        result_ad = total_net_ad - total_spends_ad
        # الصورة من خريطة ASIN مباشرة (شيت links a) | Image straight from the
        # ASIN links map ("links a" sheet)
        img_ad = _links_map_ads_orphan.get(asin_up_ad, "")
        sku_disp_ad = ", ".join(mskus_ad) if mskus_ad else asin_up_ad
        entry_ad = {"sku_up": asin_up_ad, "sku": sku_disp_ad, "asin": asin_up_ad,
                    "img": img_ad,
                    "spends": total_spends_ad, "orders": total_orders_ad,
                    "net_total": total_net_ad, "result": result_ad,
                    "cost": cost_ad, "cost_missing": cost_missing_ad}
        if total_orders_ad <= 0 or result_ad < 0:
            ads_loss_rows_ap.append(entry_ad)
        else:
            ads_profit_rows_ap.append(entry_ad)
    ads_profit_rows_ap.sort(key=lambda r: -r["result"])
    ads_loss_rows_ap.sort(key=lambda r: r["result"])  # الأكثر خسارة أولاً

    st.markdown("##### 🔔 تنبيهات الإعلانات | Ads Alerts")
    acp1, acp2 = st.columns(2)
    with acp1:
        st.markdown(
            f'<div style="background:#f0fdf4;border:1px solid #22c55e;border-right:4px solid #22c55e;'
            f'border-radius:12px;padding:12px 14px;direction:rtl;min-height:92px;">'
            f'<div style="font-size:12px;color:#374151;margin-bottom:6px;">🎯 منتجات ربحانة من الإعلانات</div>'
            f'<div style="font-size:22px;font-weight:800;color:#111827;">{len(ads_profit_rows_ap):,}</div>'
            f'<div style="font-size:11px;color:#6b7280;">صافي ربح الطلبات &gt; المصروف على الإعلان</div></div>',
            unsafe_allow_html=True)
    with acp2:
        st.markdown(
            f'<div style="background:#fef2f2;border:1px solid #ef4444;border-right:4px solid #ef4444;'
            f'border-radius:12px;padding:12px 14px;direction:rtl;min-height:92px;">'
            f'<div style="font-size:12px;color:#374151;margin-bottom:6px;">🚨 منتجات خسرانة من الإعلانات</div>'
            f'<div style="font-size:22px;font-weight:800;color:#111827;">{len(ads_loss_rows_ap):,}</div>'
            f'<div style="font-size:11px;color:#6b7280;">المصروف على الإعلان أكبر من صافي الربح (أو من غير طلبات)</div></div>',
            unsafe_allow_html=True)

    def _render_ads_alert_sku_row(r, badges_html=""):
        ci_al, cinfo_al = st.columns([1, 6])
        with ci_al:
            show_img(r["img"], 55)
        with cinfo_al:
            st.markdown(f"{asin_link_html(r.get('asin', r['sku']))}", unsafe_allow_html=True)
            st.caption(f"MSKU: {r['sku']}")
            if badges_html:
                st.markdown(badges_html, unsafe_allow_html=True)

    def _ads_insight_html(r):
        # ── سطر التكلفة + تنبيه لو غير مسجلة | Cost line + warning if unrecorded ──
        cost_note_ad = f' &nbsp;|&nbsp; 📦 التكلفة | Cost: <b>{r["cost"]:,.2f}</b> ريال'
        missing_note_ad = ""
        if r.get("cost_missing"):
            missing_note_ad = (' <span style="color:#fbbf24;">⚠️ تكلفة المنتج غير مسجلة '
                                '(تاب التكويد) — اعتُبرت صفر</span>')
        if r["orders"] <= 0:
            return (f'<span style="color:#f87171;font-size:12px;font-weight:700;">🚨 مفدتش لحد دلوقتي: '
                    f'اتصرف {r["spends"]:,.2f} ريال ولسه ما جابش أي طلبات فعلية</span>'
                    f'{cost_note_ad}{missing_note_ad}')
        if r["result"] >= 0:
            return (f'<span style="color:#4ade80;font-size:12px;font-weight:700;">🎯 الاعلان مربح: '
                    f'عدد طلبات الاعلان {r["orders"]:,.0f} طلب بصافي ربح فعلي بعد التكلفة إجمالي {r["net_total"]:,.2f} ريال مقابل '
                    f'{r["spends"]:,.2f} ريال مدفوع — حقق {r["result"]:,.2f} ريال 👌</span>'
                    f'{cost_note_ad}{missing_note_ad}')
        return (f'<span style="color:#f87171;font-size:12px;font-weight:700;">🚨 الاعلان غير مربح: '
                f'مدفوع {r["spends"]:,.2f} ريال، لكن صافي الربح الفعلي بعد التكلفة من {r["orders"]:,.0f} طلب بس {r["net_total"]:,.2f} ريال — '
                f'خسران {abs(r["result"]):,.2f} ريال إجمالي</span>'
                f'{cost_note_ad}{missing_note_ad}')

    with st.expander(f"🎯 عرض منتجات ربحانة من الإعلانات ({len(ads_profit_rows_ap):,}) | Show profitable-ads SKUs"):
        if ads_profit_rows_ap:
            df_al5 = pd.DataFrame([{
                "ASIN": r.get("asin", ""), "MSKU": r["sku"], "طلبات الإعلان | Ad Orders": r["orders"],
                "المصروف | Spends": round(r["spends"], 2),
                "التكلفة | Cost": round(r["cost"], 2),
                "صافي الربح بعد التكلفة | Net Total After Cost": round(r["net_total"], 2),
                "النتيجة | Result": round(r["result"], 2),
            } for r in ads_profit_rows_ap])
            dl_btn(df_al5, "alert_ads_profit_amz", key="dl_alert_ads_profit_ap_amz")
            for r in ads_profit_rows_ap:
                _render_ads_alert_sku_row(r, badges_html=_ads_insight_html(r))
        else:
            st.caption("لا توجد أصناف ربحانة من الإعلانات حالياً")

    with st.expander(f"🚨 عرض منتجات خسرانة من الإعلانات ({len(ads_loss_rows_ap):,}) | Show losing-ads SKUs"):
        if ads_loss_rows_ap:
            df_al6 = pd.DataFrame([{
                "ASIN": r.get("asin", ""), "MSKU": r["sku"], "طلبات الإعلان | Ad Orders": r["orders"],
                "المصروف | Spends": round(r["spends"], 2),
                "التكلفة | Cost": round(r["cost"], 2),
                "صافي الربح بعد التكلفة | Net Total After Cost": round(r["net_total"], 2),
                "النتيجة | Result": round(r["result"], 2),
            } for r in ads_loss_rows_ap])
            dl_btn(df_al6, "alert_ads_loss_amz", key="dl_alert_ads_loss_ap_amz")
            for r in ads_loss_rows_ap:
                _render_ads_alert_sku_row(r, badges_html=_ads_insight_html(r))
        else:
            st.caption("لا توجد أصناف خسرانة من الإعلانات حالياً 🎉")

    st.divider()

    # ══════════════════════════════════════════════════════════════════════
    # 📢 تحليل أداء الإعلانات (على مستوى الحملة) | Ads Performance Analysis
    # (Campaign level) — ده معتمد بس على بيانات الإعلانات الموجودة فعليًا
    # (Views/Clicks/Orders/ATC/Spends/Revenue) من غير أي افتراض لتكلفة المنتج أو
    # هامش ربح حقيقي. كل حساب هنا بيتجمّع من *كل* الحملات/الصفوف الخاصة بالـ SKU
    # أو الحملة قبل ما يتحسب — مش بياخد رقم من صف واحد بس لو فيه أكتر من حملة |
    # Based only on ad data that actually exists, no product-cost or profit-margin
    # assumptions. Every number here is summed across *all* matching campaign rows
    # before any ratio is computed — never taken from a single row when more than
    # one campaign exists.
    st.markdown("---")
    st.markdown("## 📢 تحليل أداء الإعلانات | Ads Performance Analysis")
    st.caption("مبني فقط على بيانات الإعلانات الموجودة في النظام حاليًا — بدون أي افتراض لتكلفة المنتج أو هامش الربح | Based only on ad data currently in the system — no product cost or profit margin assumptions")

    def _apa_ratios(views, clicks, atc, orders, spends, revenue):
        """يعيد حساب كل النسب من الأرقام الخام المجمّعة (مش من عمود جاهز في صف واحد)
        عشان أي SKU/حملة ليها أكتر من صف تتحسب صح | Recomputes every ratio from the
        summed raw totals (never from a single pre-computed column), so multi-row
        SKUs/campaigns are calculated correctly."""
        ctr = (clicks / views * 100) if views > 0 else 0.0
        cpc = (spends / clicks) if clicks > 0 else 0.0
        cpa = (spends / orders) if orders > 0 else 0.0
        cvr = (orders / clicks * 100) if clicks > 0 else 0.0
        roas = (revenue / spends) if spends > 0 else 0.0
        click_to_atc = (atc / clicks * 100) if clicks > 0 else 0.0
        atc_to_order = (orders / atc * 100) if atc > 0 else 0.0
        return {"ctr": ctr, "cpc": cpc, "cpa": cpa, "cvr": cvr, "roas": roas,
                "click_to_atc": click_to_atc, "atc_to_order": atc_to_order}

    # ── تجميع كل صفوف الإعلانات (Sku × Campaign) على مستوى الحملة نفسها —
    #    عشان أي حملة بتستهدف أكتر من SKU تتحسب مجمّعة صح ومتاخدش من صف واحد ──
    campaigns_apa = {}
    for sku_up_c, entries_c in ads_map_dash.items():
        for e in entries_c:
            cname_c = e["campaign"] or "—"
            agg_c = campaigns_apa.setdefault(cname_c, {
                "campaign": cname_c, "views": 0.0, "clicks": 0.0, "orders": 0.0,
                "atc": 0.0, "spends": 0.0, "revenue": 0.0, "skus": set(),
            })
            agg_c["views"]   += e["views"]
            agg_c["clicks"]  += e["clicks"]
            agg_c["orders"]  += e["orders"]
            agg_c["atc"]     += e["atc"]
            agg_c["spends"]  += e["spends"]
            agg_c["revenue"] += e["revenue"]
            agg_c["skus"].add(sku_up_c)

    for _cname, _agg in campaigns_apa.items():
        _agg.update(_apa_ratios(_agg["views"], _agg["clicks"], _agg["atc"], _agg["orders"], _agg["spends"], _agg["revenue"]))
        _agg["sku_count"] = len(_agg["skus"])
        # نتيجة الإعلان بعد الإنفاق الإعلاني فقط — مش ربح حقيقي (مفيش تكلفة منتج) |
        # Ad result after ad spend only — not real profit (no product cost known)
        _agg["ad_result"] = _agg["revenue"] - _agg["spends"]

    campaigns_list_apa = list(campaigns_apa.values())

    def _apa_render_related_skus(asin_up_set, max_show=6):
        """يعرض المنتجات (ASIN) المرتبطة بالحملة مع صورها | Renders the ASINs linked
        to this campaign, each with its product image."""
        asin_list_r = sorted(asin_up_set)
        shown_r = asin_list_r[:max_show]
        _links_map_apa_orphan = get_links_map_asin()
        for asin_up_r in shown_r:
            asin_info_r = amz_inv_map_asin.get(asin_up_r, {})
            # الصورة من خريطة ASIN مباشرة (شيت links a) | Image straight from
            # the ASIN links map ("links a" sheet)
            img_r = _links_map_apa_orphan.get(asin_up_r, "")
            ci_r, cinfo_r = st.columns([1, 6])
            with ci_r:
                show_img(img_r, 45)
            with cinfo_r:
                st.markdown(asin_link_html(asin_up_r), unsafe_allow_html=True)
                mskus_r = asin_info_r.get("skus", [])
                if mskus_r:
                    st.caption(f"MSKU: {', '.join(mskus_r)}")
        if len(asin_list_r) > max_show:
            st.caption(f"+ {len(asin_list_r) - max_show} ASIN إضافي | more ASINs")

    if not campaigns_list_apa:
        st.info("لا توجد بيانات إعلانات مرفوعة حالياً | No ad data uploaded yet")
    else:
        # ── 1) مؤشرات أداء الإعلانات | Ad Performance Metrics (إجمالي كل الحملات) ──
        tot_views_apa   = sum(c["views"] for c in campaigns_list_apa)
        tot_clicks_apa  = sum(c["clicks"] for c in campaigns_list_apa)
        tot_atc_apa     = sum(c["atc"] for c in campaigns_list_apa)
        tot_orders_apa  = sum(c["orders"] for c in campaigns_list_apa)
        tot_spends_apa  = sum(c["spends"] for c in campaigns_list_apa)
        tot_revenue_apa = sum(c["revenue"] for c in campaigns_list_apa)
        tot_ratios_apa  = _apa_ratios(tot_views_apa, tot_clicks_apa, tot_atc_apa, tot_orders_apa, tot_spends_apa, tot_revenue_apa)

        st.markdown("#### 📊 مؤشرات أداء الإعلانات | Ad Performance Metrics")
        st.caption("ℹ️ كل النسب (CTR/CPC/CPS/CVR/ROAS) بتتحسب من إجمالي الأرقام الخام لكل الحملات مجمّعة — مش من عمود جاهز في صف واحد | All ratios are computed from the raw totals across every campaign combined — never from a single pre-computed column")
        mrow1 = st.columns(4)
        with mrow1[0]:
            st.markdown(_kpi_card_html("🛒", "#2563eb", "الطلبات | Orders", f"{tot_orders_apa:,.0f}"), unsafe_allow_html=True)
        with mrow1[1]:
            st.markdown(_kpi_card_html("👁️", "#0891b2", "مرات الظهور | Impressions", f"{tot_views_apa:,.0f}"), unsafe_allow_html=True)
        with mrow1[2]:
            st.markdown(_kpi_card_html("🖱️", "#7c3aed", "النقرات | Clicks", f"{tot_clicks_apa:,.0f}"), unsafe_allow_html=True)
        with mrow1[3]:
            st.markdown(_kpi_card_html("➕", "#059669", "الإضافة إلى السلة | Add to Cart", f"{tot_atc_apa:,.0f}"), unsafe_allow_html=True)
        mrow2 = st.columns(4)
        with mrow2[0]:
            st.markdown(_kpi_card_html("📈", "#0891b2", "معدل النقر | CTR", f"{tot_ratios_apa['ctr']:.2f}%"), unsafe_allow_html=True)
        with mrow2[1]:
            st.markdown(_kpi_card_html("💵", "#f59e0b", "تكلفة النقرة | CPC", f"{tot_ratios_apa['cpc']:.2f} ريال"), unsafe_allow_html=True)
        with mrow2[2]:
            st.markdown(_kpi_card_html("🎯", "#dc2626", "تكلفة الطلب | CPS / CPA", f"{tot_ratios_apa['cpa']:.2f} ريال"), unsafe_allow_html=True)
        with mrow2[3]:
            st.markdown(_kpi_card_html("📊", "#9333ea", "معدل التحويل | CVR", f"{tot_ratios_apa['cvr']:.2f}%"), unsafe_allow_html=True)
        mrow3 = st.columns(3)
        with mrow3[0]:
            st.markdown(_kpi_card_html("🎯", "#16a34a", "العائد على الإنفاق الإعلاني | ROAS", f"{tot_ratios_apa['roas']:.2f}"), unsafe_allow_html=True)
        with mrow3[1]:
            st.markdown(_kpi_card_html("💰", "#16a34a", "الإيراد | Revenue", f"{tot_revenue_apa:,.2f} ريال"), unsafe_allow_html=True)
        with mrow3[2]:
            st.markdown(_kpi_card_html("💸", "#dc2626", "الإنفاق الإعلاني | Ad Spend", f"{tot_spends_apa:,.2f} ريال"), unsafe_allow_html=True)

        st.write("")

        # ── 2) تحليل مسار الإعلان | Advertising Funnel ──
        st.markdown("#### 🔻 تحليل مسار الإعلان | Advertising Funnel")
        st.caption("عشان تعرف أين يحدث انخفاض الأداء في مسار الإعلان | See exactly where performance drops along the funnel")
        fcols_apa = st.columns(4)
        funnel_stages_apa = [
            ("👁️ مرات الظهور | Impressions", tot_views_apa, None),
            ("🖱️ النقرات | Clicks", tot_clicks_apa, tot_ratios_apa["ctr"]),
            ("➕ الإضافة إلى السلة | Add to Cart", tot_atc_apa, tot_ratios_apa["click_to_atc"]),
            ("🛒 الطلبات | Orders", tot_orders_apa, tot_ratios_apa["atc_to_order"]),
        ]
        for fc_apa, (label_f, val_f, rate_f) in zip(fcols_apa, funnel_stages_apa):
            with fc_apa:
                rate_html_f = (f'<div style="font-size:11px;color:#f59e0b;margin-top:4px;">↓ {rate_f:.1f}%</div>'
                               if rate_f is not None else "")
                st.markdown(
                    f'<div style="background:#1e293b;border:1px solid #334155;border-radius:10px;'
                    f'padding:12px 10px;text-align:center;">'
                    f'<div style="font-size:11px;color:#94a3b8;">{label_f}</div>'
                    f'<div style="font-size:20px;font-weight:800;color:#e2e8f0;">{val_f:,.0f}</div>'
                    f'{rate_html_f}</div>', unsafe_allow_html=True)
        st.markdown(
            f'<div style="margin-top:8px;font-size:12px;color:#94a3b8;">'
            f'📈 معدل النقر | CTR: <b style="color:#e2e8f0;">{tot_ratios_apa["ctr"]:.2f}%</b> &nbsp;|&nbsp; '
            f'النقر → السلة | Click → Cart: <b style="color:#e2e8f0;">{tot_ratios_apa["click_to_atc"]:.1f}%</b> &nbsp;|&nbsp; '
            f'السلة → الطلب | Cart → Order: <b style="color:#e2e8f0;">{tot_ratios_apa["atc_to_order"]:.1f}%</b> &nbsp;|&nbsp; '
            f'معدل التحويل الكلي | CVR: <b style="color:#e2e8f0;">{tot_ratios_apa["cvr"]:.2f}%</b>'
            f'</div>', unsafe_allow_html=True)

        st.write("")

        # ── دوال التحليل التلقائي / التصنيف / التوصية — مبنية على أكتر من مؤشر مع
        #    بعض (ROAS+CPA+CVR+CTR+CPC+Orders+Spend+Revenue) مش مؤشر واحد بس ──
        def _apa_insight(c):
            if c["orders"] <= 0:
                return ("🔴", "أداء ضعيف | Poor Performance",
                        f"اتصرف {c['spends']:,.2f} ريال على الحملة ولسه ما جابتش أي طلبات فعلية.")
            score_c = 0
            if c["roas"] >= 3: score_c += 2
            elif c["roas"] >= 1.5: score_c += 1
            elif c["roas"] < 1: score_c -= 2
            if c["cvr"] >= 3: score_c += 1
            elif c["cvr"] < 1: score_c -= 1
            if c["ctr"] >= 1: score_c += 1
            elif c["ctr"] < 0.3: score_c -= 1
            if c["spends"] > 0 and c["revenue"] < c["spends"]:
                score_c -= 2
            if score_c >= 3:
                return ("🟢", "أداء جيد | Good Performance",
                        f"الحملة تحقق ROAS {c['roas']:.2f} مع معدل تحويل {c['cvr']:.2f}% جيد.")
            elif score_c >= 0:
                return ("🟡", "يحتاج إلى تحسين | Needs Improvement",
                        f"الحملة بتاخد نقرات معقولة (CTR {c['ctr']:.2f}%)، لكن التحويل للطلبات ({c['cvr']:.2f}%) أو الـ ROAS ({c['roas']:.2f}) لسه محتاج تحسين.")
            else:
                return ("🔴", "أداء ضعيف | Poor Performance",
                        f"تكلفة الطلب {c['cpa']:,.2f} ريال مرتفعة مقارنة بعدد الطلبات ({c['orders']:,.0f}) والـ ROAS {c['roas']:.2f}.")

        def _apa_classification(c):
            icon_i, _t, _d = _apa_insight(c)
            if c["orders"] <= 0:
                return "🔴", "أداء ضعيف | Poor Performance"
            if icon_i == "🟢":
                return "🟢", "أداء قوي | Strong Performance"
            if icon_i == "🟡":
                if c["roas"] < 1.5 and c["cvr"] < 2:
                    return "🟠", "يحتاج إلى تحسين | Needs Improvement"
                return "🟡", "يحتاج إلى مراقبة | Needs Monitoring"
            return "🔴", "أداء ضعيف | Poor Performance"

        def _apa_recommendation(c):
            cls_icon_c, _l = _apa_classification(c)
            if c["orders"] <= 0 and c["spends"] > 0:
                return "قلل الإنفاق | Reduce Spend"
            if cls_icon_c == "🟢":
                return "استمر | Continue"
            if cls_icon_c == "🟡":
                return "راقب | Monitor"
            if cls_icon_c == "🟠":
                return "حسّن الحملة | Optimize"
            return "راجع الحملة | Review"

        # ── 5) مقارنة الحملات | Campaign Comparison (قبل التفاصيل عشان تبان الأهم فوق) ──
        st.markdown("#### 🏆 مقارنة الحملات | Campaign Comparison")
        camps_with_orders_apa = [c for c in campaigns_list_apa if c["orders"] > 0]
        camps_with_clicks_apa = [c for c in campaigns_list_apa if c["clicks"] > 0]
        comp_specs_apa = [
            ("🏆", "أفضل حملة حسب ROAS | Best by ROAS", camps_with_orders_apa, lambda c: c["roas"], lambda c: f"ROAS {c['roas']:.2f}"),
            ("💰", "أعلى إيراد | Highest Revenue", campaigns_list_apa, lambda c: c["revenue"], lambda c: f"{c['revenue']:,.2f} ريال"),
            ("🛒", "أكثر طلبات | Most Orders", campaigns_list_apa, lambda c: c["orders"], lambda c: f"{c['orders']:,.0f} طلب"),
            ("💸", "أعلى إنفاق إعلاني | Highest Ad Spend", campaigns_list_apa, lambda c: c["spends"], lambda c: f"{c['spends']:,.2f} ريال"),
            ("🎯", "أفضل تكلفة طلب | Best CPA", camps_with_orders_apa, lambda c: -c["cpa"], lambda c: f"{c['cpa']:.2f} ريال"),
            ("📈", "أفضل معدل تحويل | Best CVR", camps_with_clicks_apa, lambda c: c["cvr"], lambda c: f"{c['cvr']:.2f}%"),
            ("👁️", "أفضل معدل نقر | Best CTR", campaigns_list_apa, lambda c: c["ctr"], lambda c: f"{c['ctr']:.2f}%"),
        ]
        comp_cols_apa = st.columns(2)
        for i_apa, (icon_s, label_s, pool_s, key_s, fmt_s) in enumerate(comp_specs_apa):
            best_c_apa = max(pool_s, key=key_s, default=None)
            with comp_cols_apa[i_apa % 2]:
                if best_c_apa:
                    st.markdown(_kpi_card_html(icon_s, "#2563eb", label_s, best_c_apa["campaign"]), unsafe_allow_html=True)
                    st.caption(fmt_s(best_c_apa))
                    with st.expander(f"🏷️ الأصناف (ASIN) | ASINs ({best_c_apa['sku_count']})"):
                        _apa_render_related_skus(best_c_apa["skus"])
                else:
                    st.markdown(_kpi_card_html(icon_s, "#6b7280", label_s, "—"), unsafe_allow_html=True)

        st.write("")

        # ── 3+6+7) التحليل التلقائي + التصنيف + التوصية لكل حملة | Automatic
        #    insight + classification + recommendation per campaign ──
        st.markdown("#### 🔎 تحليل كل حملة | Per-Campaign Analysis")
        for c_apa in sorted(campaigns_list_apa, key=lambda x: -x["spends"]):
            icon_i, title_i, desc_i = _apa_insight(c_apa)
            icon_c, label_c = _apa_classification(c_apa)
            rec_c = _apa_recommendation(c_apa)
            bg_i = {"🟢": "#052e1655", "🟡": "#78350f33", "🔴": "#4c051655"}[icon_i]
            border_i = {"🟢": "#16a34a", "🟡": "#f59e0b", "🔴": "#dc2626"}[icon_i]
            with st.expander(f"{icon_i} {c_apa['campaign']} — {c_apa['sku_count']} ASIN | {c_apa['orders']:,.0f} طلب"):
                st.markdown(
                    f'<div dir="rtl" style="background:{bg_i};border:1px solid {border_i};border-radius:8px;padding:8px 12px;margin-bottom:8px;">'
                    f'<b>{icon_i} {title_i}</b><br><span style="font-size:13px;">{desc_i}</span>'
                    f'</div>', unsafe_allow_html=True)
                st.markdown(
                    f"👁️ ظهور: {c_apa['views']:,.0f} &nbsp;|&nbsp; 🖱️ نقرات: {c_apa['clicks']:,.0f} &nbsp;|&nbsp; "
                    f"➕ سلة: {c_apa['atc']:,.0f} &nbsp;|&nbsp; 🛒 طلبات: {c_apa['orders']:,.0f}<br>"
                    f"📊 CTR: {c_apa['ctr']:.2f}% &nbsp;|&nbsp; 💵 CPC: {c_apa['cpc']:.2f} &nbsp;|&nbsp; "
                    f"🎯 CPS/CPA: {c_apa['cpa']:.2f} &nbsp;|&nbsp; 📈 CVR: {c_apa['cvr']:.2f}% &nbsp;|&nbsp; 🎯 ROAS: {c_apa['roas']:.2f}<br>"
                    f"💸 إنفاق: {c_apa['spends']:,.2f} ريال &nbsp;|&nbsp; 💰 إيراد: {c_apa['revenue']:,.2f} ريال &nbsp;|&nbsp; "
                    f"📉 نتيجة الإعلان بعد الإنفاق الإعلاني | Ad Result After Ad Spend: <b>{c_apa['ad_result']:,.2f} ريال</b>")
                st.markdown(f"🏷️ التصنيف | Classification: **{icon_c} {label_c}**")
                st.markdown(f"✅ التوصية | Recommendation: **{rec_c}**")
                st.markdown("🏷️ **الأصناف المرتبطة | Related ASINs**")
                _apa_render_related_skus(c_apa["skus"])

        st.caption(
            "ℹ️ مقارنة الفترة الحالية بالفترة السابقة (📈/📉 Revenue, Orders, Ad Spend, ROAS, CPA, CTR, CVR, "
            "Clicks, Add to Cart) مش متاحة هنا لسه — لأن تاب الإعلانات بيحفظ إجمالي كل حملة لحظيًا من غير "
            "تاريخ يومي، فمفيش فترة سابقة نقارن بيها. لو حبينا نفعّلها، محتاجين نبدأ نسجّل نسخة/تاريخ لكل "
            "تحديث في شيت الإعلانات | Current-vs-previous-period comparison isn't available yet because the "
            "Advertisements sheet only stores each campaign's live cumulative totals, with no daily date "
            "history to compare against. Enabling it would require snapshotting the ads sheet with dates.")



# ══ تاب المبيعات الموحد — المفتاح الأساسي = كود التكويد | Unified Sales tab keyed by Tacweed code ══
# بدل تابين (FBN / FBB): كل كود تكويد له كارت واحد فيه (صورة، شارات المستودع/التكلفة، شحنات الصين)،
# وتحته كل SKU تابع للكود، وتحت كل SKU مبيعات FBN ومبيعات FBB (لو فيه) بنفس التفاصيل القديمة بالظبط.
def _t14_build_rows(*args, **kwargs):
    return calc._t14_build_rows(_ENV, *args, **kwargs)

def _t14_sku_part(r):
    """الجزء الخاص بالـ SKU نفسه (مرة واحدة لكل SKU): الرابط، الإعلانات، Xdock، السعر والربح وحكم الإعلان."""
    st.markdown(f"**SKU:** {sku_link_html(r['sku'])}", unsafe_allow_html=True)
    if r.get("not_in_inventory"):
        st.markdown(
            '<span style="background:#78350f;color:#fde68a;padding:2px 8px;'
            'border-radius:6px;font-size:12px;">⚠️ مش موجود في ملف المخزون المرفوع | Not in uploaded inventory</span>',
            unsafe_allow_html=True)
    # ══ أداء الإعلانات (لو الـ SKU ده معلن عليه) | Ads performance (if advertised) ══
    ads_entries_t14 = ads_map_t14.get(r["sku_up"])
    # ── الإجماليات دي بتتحسب دايمًا (صفر لو مفيش إعلانات) عشان متفضلش
    #    قيم قديمة من الـ SKU اللي قبله في اللستة | Always computed
    #    (zero when no ads) so stale values from the previous SKU in
    #    the loop never leak through ──
    total_spends_t14  = sum(a["spends"]  for a in ads_entries_t14) if ads_entries_t14 else 0.0
    total_revenue_t14 = sum(a["revenue"] for a in ads_entries_t14) if ads_entries_t14 else 0.0
    total_orders_t14  = sum(a["orders"]  for a in ads_entries_t14) if ads_entries_t14 else 0.0
    # ══ مخزون Xdock (من تاب LIVE) — مخزون منفصل عن Inventory، لو قرب يخلص محتاج تزويد ══
    live_info_t14 = live_map_t14.get(r["sku_up"])
    if live_info_t14 is not None:
        xnet_t14 = live_info_t14.get("stock_xdock_net", 0)
        xlow_t14 = xnet_t14 <= xdock_threshold_t14
        st.markdown(
            f'<span class="wh-badge" style="background:{"#7f1d1d" if xlow_t14 else "#3b0764"};'
            f'color:{"#fca5a5" if xlow_t14 else "#e9d5ff"};">'
            f'{"🔴" if xlow_t14 else "🟣"} مخزون Xdock: {xnet_t14:,}'
            f'{" — قارب على النفاد" if xlow_t14 else ""}</span>',
            unsafe_allow_html=True)

    # ══ صافي سعر البيع بعد العمولة والتوصيل والضريبة (يعرض دايمًا، وبيحط صفر في أي
    #    قيمة ناقصة مع تنبيه) | Net price after commission/delivery/VAT (always shown;
    #    missing values default to zero with a warning) ══
    render_price_profit_block_t14(
        r, com_map_t14, live_map_t14, sales_dates,
        ads_entries_t14, total_orders_t14, total_spends_t14, reco_noon=True)
    if ads_entries_t14:
        render_noon_ads_expander(r["sku_up"], ads_entries_t14, total_spends_t14, total_revenue_t14, total_orders_t14)

_COV_SINK = [None]      # لو فيها قائمة: الكبسولات بتتجمع فيها وتتعرض كلها في سطر واحد فوق

def _cov_pill(text, color, title=""):
    """شارة صغيرة (كبسولة) بخط صغير — بدل الصناديق الكبيرة؛ النص الكامل في التلميح (hover)."""
    import html as _h
    _html = (f'<span title="{_h.escape(title or text, quote=True)}" style="display:inline-block;background:{color};color:#fff;'
             f'border-radius:999px;padding:2px 10px;margin:2px 3px 2px 0;font-size:11px;font-weight:700;line-height:1.5;">{text}</span>')
    if _COV_SINK[0] is not None:
        _COV_SINK[0].append(_html)
    else:
        st.markdown(_html, unsafe_allow_html=True)

def _t14_channel_part(r, ch, with_sku_level, part="all"):
    """مبيعات قناة واحدة (FBN أو FBB) للـ SKU: أمس، الأيام، المخزون، حالة التغطية، الجدولة.
    with_sku_level=True بيعرض كمان الجدولة الحديثة/الموافقة المعلقة/الملاحظات (اللي هي على مستوى الـ SKU)
    مرة واحدة بس."""
    if ch == "TY" and part in ("stock", "cov"):
        if part == "stock":
            st.markdown(f"📈 **شهري:** {r['sales_month']} &nbsp;|&nbsp; 📊 **يومي أخير:** {r['effective_avg']:.1f}")
        return
    _chip = {"FBN": ("#064e3b", "#6ee7b7", "🅽 مبيعات نون FBN"), "FBB": ("#1e3a8a", "#bfdbfe", "🅱 مبيعات نون FBB"),
             "FSAB": ("#7c2d12", "#fed7aa", "🅰 مبيعات أمازون عادي FSAB"), "FBA": ("#78350f", "#fde68a", "📦 مبيعات أمازون FBA"),
             "TY": ("#4c1d95", "#ddd6fe", "🟣 مبيعات تريندول")}[ch]
    if part in ("all", "sales"):
        st.markdown(f'<span class="wh-badge" style="background:{_chip[0]};color:{_chip[1]};">{_chip[2]}</span>', unsafe_allow_html=True)
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
        if day_parts:      # أول أمس وما قبله: في شريط مقفول يتفتح بالضغط (توفير مساحة)
            with st.expander(f"📅 أول أمس وما قبله ({len(day_parts)} يوم)", expanded=False):
                st.markdown("<br>".join(day_parts), unsafe_allow_html=True)

    if part in ("all", "stock"):
        # مخزون + مبيع شهري
        st.markdown(
            f"📦 **مخزون:** {r['stock']} &nbsp;|&nbsp; "
            f"📈 **شهري:** {r['sales_month']} &nbsp;|&nbsp; "
            f"📊 **يومي أخير:** {r['effective_avg']:.1f} &nbsp;|&nbsp; "
            f"⏳ **نفاد خلال:** {r['days_to_stockout'] if r['days_to_stockout'] < 9999 else '—'} يوم"
        )

    if part in ("all", "cov"):
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
            _short = cov_badge_text.split(" | ")[0].split(" — ")[0]      # نسخة مختصرة في الكبسولة، والنص الكامل في التلميح
            if stock_self_ok and r["effective_avg"] > 0 and not sched_t14:
                _short = f"✅ كافٍ {r['days_to_stockout']} يوم"
            _cov_pill(_short, cov_badge_color, cov_badge_text)

        # ══ مجدولة خلال آخر 4 أيام؟ (لو فعلاً ليها جدولة) — تعرض ASN + الكمية + التاريخ ══
        if recent_sched_t14 and with_sku_level:
            _dn = max((datetime.now().date() - recent_sched_t14["parsed"]).days, 0)
            _ago = {0: "النهارده", 1: "من يوم", 2: "من يومين"}.get(_dn, f"من {_dn} أيام")
            _cov_pill(f'📅 ASN {recent_sched_t14["asn"]} · {recent_sched_t14.get("qty","")} قطعة · {"كانت " if _dn >= 1 else ""}مجدولة {_ago}',
                      "#7c3aed",
                      f'مجدولة خلال آخر 4 أيام — ASN {recent_sched_t14["asn"]} | الكمية {recent_sched_t14.get("qty","")} | بتاريخ {recent_sched_t14["date"]} [{recent_sched_t14["source_label"]}]')

        if with_sku_level and r["sku_up"] in pending_approval_skus_t14:
            _cov_pill("⏳ قيد الموافقة", "#b45309",
                      "قيد الموافقة — في انتظار اعتماد الجدولة من برنامج الجدولة المكتبي، لا تطلبه أو تجدوله تاني")

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
        if is_needs_sched_only and ch in ("FBN", "FBB"):
            _new_transferred.append({
                "sku": r["sku"], "sku_up": r["sku_up"], "stock": r["stock"],
                "sales_month": r["sales_month"], "img": r["img"],
                "effective_avg": r["effective_avg"], "days_to_stockout": r["days_to_stockout"],
                "day_counts": r["day_counts"],
            })
            st.caption("📌 مرحّل لتاب مراجعة المخزون | Transferred to Stock Review tab")

        if with_sku_level and un_notes:
            for note in un_notes:
                st.markdown(big_note_html(note), unsafe_allow_html=True)
        if with_sku_level:
            render_recent_expired_note(r["sku"])


# ══ أمازون داخل تاب المبيعات الموحد — صفوف ASIN بقناتين: عادي (FSAB) و FBA ══
# Amazon inside the unified Sales tab — ASIN rows with two channels: normal (FSAB) and FBA.
def _t14_build_rows_amz(*args, **kwargs):
    return calc._t14_build_rows_amz(_ENV, *args, **kwargs)

# ══ متابعة المنافسين (من بوت noon_tracker) | Noon competitor tracking ══
# التابات NoonTrack_* بيعملها البوت المكتبي أول مرة يشتغل؛ هنا بنقراها (لو موجودة) في طلب واحد
# ونحتفظ بيها 10 دقايق في الجلسة عشان مانستهلكش الـ quota.
NOON_TRACK_TTL = 600

def get_noon_track(force=False):
    c = st.session_state.get("_noon_track")
    if c and not force and time.time() - c[0] < NOON_TRACK_TTL:
        return c[1]
    data = {}
    try:
        have = {w.title for w in ss.worksheets()}
        if calc.NOON_TRACK_TABS[0] in have:
            names = [t for t in calc.NOON_TRACK_TABS if t in have]
            res = ss.values_batch_get(["'" + t + "'" for t in names])
            rows = {t: (vr.get("values", []) or []) for t, vr in zip(names, res.get("valueRanges", []))}
            data = calc.build_noon_track(rows.get(calc.NOON_TRACK_TABS[0]), rows.get(calc.NOON_TRACK_TABS[1]),
                                         rows.get(calc.NOON_TRACK_TABS[2]))
    except Exception:
        data = c[1] if c else {}
    st.session_state["_noon_track"] = (time.time(), data)
    return data

NT_SKUS_TAB = "NoonTrack_SKUs"

def publish_noon_track_skus(skus):
    """بيكتب في تاب NoonTrack_SKUs (بيتعمل أول مرة) اسكيوهات نون الظاهرة في تاب المبيعات (بدون -1)،
    والبوت المكتبي بيسحب المنافسين لنفس القايمة دي بس. بيتكتب لما القايمة تتغير أو كل ساعة على الأكتر."""
    try:
        keys = sorted({calc.noon_track_key(x) for x in skus if x})
    except AttributeError:      # calc.py قديم على GitHub — ارفع calc.py مع App.py
        return
    if not keys:
        return
    sig = hash(tuple(keys))
    last = st.session_state.get("_nt_sku_pub")
    if last and last[0] == sig and time.time() - last[1] < 3600:
        return
    st.session_state["_nt_sku_pub"] = (sig, time.time())    # حتى لو فشل، منعيدش المحاولة في كل ضغطة
    try:
        try:
            ws = ss.worksheet(NT_SKUS_TAB)
        except gspread.exceptions.WorksheetNotFound:
            ws = ss.add_worksheet(title=NT_SKUS_TAB, rows=str(len(keys) + 50), cols="2")
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        ws.clear()
        ws.append_rows([["SKU", "آخر تحديث"]] + [[k, stamp] for k in keys], value_input_option="RAW")
    except Exception:
        pass

def noon_track_entry(sku_up):
    try:
        return get_noon_track().get(calc.noon_track_key(sku_up))
    except AttributeError:      # calc.py قديم على GitHub — ارفع calc.py مع App.py
        return None

def render_noon_rivals(sku_up, our_price=None):
    """كبسولات جنب صورة المنتج: صاحب البوكس + المنافسين الأرخص مننا (ولو مفيش، بتقول السبب)."""
    try:
        all_tr = get_noon_track()
        e = all_tr.get(calc.noon_track_key(sku_up))
        if not e:
            if all_tr:      # البوت اشتغل بس الـ SKU ده مسحبش (مش في تاب NoonTrack_SKUs وقت التشغيل أو فشل سحبه)
                st.caption("🔎 المنافسين: الـ SKU ده لسه مسحبش — يتسحب في تشغيلة البوت الجاية")
            return
        h = calc.noon_track_rivals_html(e, our_price)
        if h:
            st.markdown(h, unsafe_allow_html=True)
        else:
            st.caption("🔎 المنافسين: " + (e["p"].get("الحالة") or "مفيش بائعين على المنتج") + " — مفيش بوكس أو أسعار نقارن بيها")
    except AttributeError:      # calc.py قديم على GitHub
        st.caption("⚠️ ارفع calc.py الجديد مع App.py علشان تظهر كبسولات المنافسين")

def render_noon_track(sku_up, our_price=None):
    """سلايدر المنافسين والتقييمات تحت SKU نون (لو البوت جاب بياناته)."""
    e = noon_track_entry(sku_up)
    if not e:
        return
    with st.expander(calc.noon_track_label(e)):
        st.markdown(calc.noon_track_html(e, our_price), unsafe_allow_html=True)

def _t14_noon_part(r, part):
    """جزء من كارت نون: ident (SKU + تنبيهات + Xdock) / price (السعر والربح وحكم الإعلان) / ads (أداء الإعلانات).
    مقسّم لأجزاء عشان تتحاذى مع أمازون في صفوف متوازية."""
    ent = ads_map_t14.get(r["sku_up"])
    tot_s = sum(a["spends"] for a in ent) if ent else 0.0
    tot_r = sum(a["revenue"] for a in ent) if ent else 0.0
    tot_o = sum(a["orders"] for a in ent) if ent else 0.0
    if part == "ident":
        st.markdown(f"**SKU:** {sku_link_html(r['sku'])}", unsafe_allow_html=True)
        if r.get("not_in_inventory"):
            st.markdown(
                '<span style="background:#78350f;color:#fde68a;padding:2px 8px;'
                'border-radius:6px;font-size:12px;">⚠️ مش موجود في ملف المخزون المرفوع | Not in uploaded inventory</span>',
                unsafe_allow_html=True)
        live_info = live_map_t14.get(r["sku_up"])
        if live_info is not None:
            xnet = live_info.get("stock_xdock_net", 0)
            xlow = xnet <= xdock_threshold_t14
            st.markdown(
                f'<span class="wh-badge" style="background:{"#7f1d1d" if xlow else "#3b0764"};'
                f'color:{"#fca5a5" if xlow else "#e9d5ff"};">'
                f'{"🔴" if xlow else "🟣"} مخزون Xdock: {xnet:,}'
                f'{" — قارب على النفاد" if xlow else ""}</span>',
                unsafe_allow_html=True)
    elif part in ("box", "calc", "verdict"):
        render_price_profit_block_t14(r, com_map_t14, live_map_t14, sales_dates, ent, tot_o, tot_s, part=part, reco_noon=True)
    elif part == "comp":
        render_noon_track(r["sku_up"])
    elif part == "ads" and ent:
        render_noon_ads_expander(r["sku_up"], ent, tot_s, tot_r, tot_o)

def _t14_amz_part(r, part):
    """نفس _t14_noon_part لكن لـ ASIN أمازون."""
    asin_up = (r.get("asin") or "").upper()
    ent = ads_map_amz_t14.get(asin_up)
    tot_s = sum(a["spends"] for a in ent) if ent else 0.0
    tot_r = sum(a["revenue"] for a in ent) if ent else 0.0
    tot_o = sum(a["orders"] for a in ent) if ent else 0.0
    if part == "ident":
        st.markdown(f"**ASIN:** {asin_link_html(r.get('asin', ''))}", unsafe_allow_html=True)
        st.caption(f"MSKU: {r['sku']}")
        if r.get("not_in_inventory"):
            st.markdown(
                '<span style="background:#78350f;color:#fde68a;padding:2px 8px;'
                'border-radius:6px;font-size:12px;">⚠️ مش موجود في مخزون أمازون المرفوع | Not in uploaded Amazon stock</span>',
                unsafe_allow_html=True)
    elif part in ("box", "calc", "verdict"):
        render_price_profit_block_t14(r, com_map_t14, live_map_t14, sales_dates, ent, tot_o, tot_s, part=part)
    elif part == "ads" and ent:
        with st.expander(f"📢 إعلانات ({len(ent)})"):
            st.markdown(
                f"💸 **إجمالي المصروف:** {tot_s:,.2f} ريال &nbsp;|&nbsp; "
                f"💰 **إجمالي الإيراد:** {tot_r:,.2f} ريال &nbsp;|&nbsp; "
                f"🛒 **طلبات من الإعلان:** {tot_o:,.0f}")
            for ad in ent:
                st.markdown(
                    f"**{ad['campaign']}**<br>"
                    f"👁️ ظهور: {ad['views']:,.0f} &nbsp;|&nbsp; 🖱️ نقرات: {ad['clicks']:,.0f} &nbsp;|&nbsp; "
                    f"🛒 طلبات: {ad['orders']:,.0f}<br>"
                    f"💸 مصروف: {ad['spends']:,.2f} ريال &nbsp;|&nbsp; 💰 إيراد: {ad['revenue']:,.2f} ريال<br>"
                    f"📊 CTR: {ad['ctr']:.2f}% &nbsp;|&nbsp; 🎯 ROAS: {ad['roas']:.2f} &nbsp;|&nbsp; "
                    f"CPC: {ad['cpc']:.2f} &nbsp;|&nbsp; CVR: {ad['cvr']:.2f}%",
                    unsafe_allow_html=True)
                st.markdown("<hr style='margin:6px 0;'>", unsafe_allow_html=True)

def _t14_trendyol_placeholder():
    """مكان تريندول لما مفيش بيانات: بيوضح السبب | Shown when there's no Trendyol data."""
    st.caption("🔜 مفيش بيانات تريندول — ارفع ملف الطلبات من تاب «رفع ملف» | no Trendyol data yet")

def _t14_ty_part(r, part):
    """جزء من كارت تريندول: ident (Barcode + اسم المنتج) / box / calc / verdict (السعر والربح)."""
    if part == "ident":
        _bc = r["sku"]
        st.markdown(
            f'**Barcode:** <span style="font-family:monospace;font-weight:700;color:#e2e8f0;background:#1e293b;'
            f'border:1px solid #334155;border-radius:6px;padding:2px 10px;user-select:all;-webkit-user-select:all;">{_bc}</span>',
            unsafe_allow_html=True)
        if r.get("name"):
            st.caption(r["name"][:70])
    elif part in ("box", "calc", "verdict", "ads"):
        ent = get_ads_by_barcode_trendyol().get(r["sku_up"])
        tot_s = sum(a["spends"] for a in ent) if ent else 0.0
        tot_r = sum(a["revenue"] for a in ent) if ent else 0.0
        tot_o = sum(a["orders"] for a in ent) if ent else 0.0
        if part == "ads":
            if ent:
                with st.expander(f"📢 إعلانات ({len(ent)})" + ("" if all(a["exact"] for a in ent) else " ≈")):
                    st.markdown(f"💸 **المصروف:** {tot_s:,.2f} &nbsp;|&nbsp; 💰 **الإيراد:** {tot_r:,.2f} &nbsp;|&nbsp; 🛒 **طلبات:** {tot_o:,.1f}")
                    for ad in ent:
                        st.markdown(f"**{ad['campaign']}**<br>👁️ {ad['views']:,.0f} &nbsp;|&nbsp; 🖱️ {ad['clicks']:,.0f} &nbsp;|&nbsp; 🛒 {ad['orders']:,.1f} "
                                    f"&nbsp;|&nbsp; 💸 {ad['spends']:,.2f} &nbsp;|&nbsp; 💰 {ad['revenue']:,.2f} &nbsp;|&nbsp; 🎯 ROAS {ad['roas']:.2f}", unsafe_allow_html=True)
                        st.markdown("<hr style='margin:6px 0;'>", unsafe_allow_html=True)
        else:
            render_price_profit_block_t14(r, com_map_ty_t14, {}, sales_dates, ent, tot_o, tot_s, part=part)
            if part == "box" and not ty_fee_set_t14:
                st.caption("ℹ️ توصيل تريندول محسوب 0 — ضيف trendyol_delivery_fee في Settings")

def _t14_yday_line(parts, store="noon"):
    """سطر تحت اسم المتجر: «📅 أمس: N (تخزين … | عادي …)» — شريحة بخلفية غامقة وكتابة بيضاء (واضحة في الثيم الفاتح والغامق).
    بيظهر بس لو فيه مبيعات أمس | Shown only when yesterday > 0."""
    tot = sum(y for _n, y in parts)
    if tot <= 0:
        return
    _bg, _bd = {"noon": ("#064e3b", "#34d399"), "amz": ("#7c2d12", "#fb923c"), "ty": ("#4c1d95", "#a78bfa")}[store]
    detail = "  |  ".join(f"{n} <b>{y:g}</b>" for n, y in parts if n)
    st.markdown(f'<div dir="rtl" style="display:inline-block;background:{_bg};border:1px solid {_bd};border-radius:6px;'
                f'padding:4px 10px;margin:4px 0;font-size:14px;font-weight:700;color:#ffffff;line-height:1.6;">'
                f'📅 أمس: <b style="font-size:16px;">{tot:g}</b>' + (f' <span style="font-weight:600;color:#f1f5f9;">( {detail} )</span>' if detail else '') + '</div>',
                unsafe_allow_html=True)

def _t14_day_tile(bg, border, num_color, head, date_s, total, noon, fbn, fbb, amz, fsab, fba, ty=0):
    """تايل إجمالي يوم — كل سطر بخط بولد وبترتيب ثابت (المتجر ثم قنواته)."""
    return (
        f'<div dir="rtl" style="background:{bg};border-radius:8px;padding:8px 10px;text-align:center;margin:2px;{border}">'
        f'<div style="font-size:12px;font-weight:700;color:#e2e8f0;">{head}</div>'
        f'<div style="font-size:12px;font-weight:600;color:#94a3b8;">{date_s}</div>'
        f'<div style="font-size:26px;font-weight:800;color:{num_color};line-height:1.2;">{total}</div>'
        f'<div style="margin-top:4px;padding-top:4px;border-top:1px solid #33415588;">'
        f'<div style="font-size:13px;font-weight:800;color:#34d399;">{plat_badge("noon", 16)}نون: {noon}</div>'
        f'<div style="font-size:12px;font-weight:700;color:#d1fae5;">FBN {fbn} · FBB {fbb}</div></div>'
        f'<div style="margin-top:4px;">'
        f'<div style="font-size:13px;font-weight:800;color:#fb923c;">{plat_badge("amazon", 16)}أمازون: {amz}</div>'
        f'<div style="font-size:12px;font-weight:700;color:#ffedd5;">FSAB {fsab} · FBA {fba}</div></div>'
        f'<div style="margin-top:4px;">'
        f'<div style="font-size:13px;font-weight:800;color:#a78bfa;">{plat_badge("trendyol", 16)}تريندول: {ty}</div></div>'
        f'</div>')
# ══════════════════════════════════════════════════════════════════════════════
# ══ تنبيهات العقود | Contract milestone alerts ══
# جدول العقود (تاب Contracts في جوجل شيت — بيتعمل تلقائي): اسم العقد + تاريخ البداية + تاريخ النهاية.
# كل شهر من بداية العقد (ولحد النهاية) بيظهر في أعلى تاب المبيعات «انتهى الشهر N من العقد» لمدة يومين،
# وخلال اليومين دول تصفّح المبيعات متوقف لحد ما تضغط «تم الاطلاع». ممكن تضيف/تحذف العقود من هنا أو من الشيت مباشرة.
# Every month after the contract start (up to the end date) a "month N ended" notice shows at the top of the Sales
# tab for 2 days; while it's unacknowledged the Sales tab is blocked until «Acknowledged» is pressed.
# ══════════════════════════════════════════════════════════════════════════════
import calendar as _ct_cal





def _ct_read():
    """صفوف العقد من الشيت (بالاسم مش بالمكان): قايمة dicts فيها row (رقم الصف في الشيت)."""
    raw = get_cached(contracts_sheet) or []
    if not raw:
        return []
    pos = {str(h).strip(): i for i, h in enumerate(raw[0])}
    g = lambda row, col: (row[pos[col]].strip() if col in pos and pos[col] < len(row) and row[pos[col]] is not None else "")
    out = []
    for i, row in enumerate(raw[1:], start=2):
        name = g(row, CONTRACT_COLS[0])
        if not name and not g(row, CONTRACT_COLS[1]):
            continue
        out.append({"row": i, "name": name, "start": _ct_parse_date(g(row, CONTRACT_COLS[1])), "end": _ct_parse_date(g(row, CONTRACT_COLS[2])),
                    "start_raw": g(row, CONTRACT_COLS[1]), "end_raw": g(row, CONTRACT_COLS[2]), "notes": g(row, CONTRACT_COLS[3]),
                    "acked": {t.strip() for t in g(row, CONTRACT_COLS[4]).split(",") if t.strip()},
                    "added": _ct_parse_date(g(row, CONTRACT_COLS[6])), "pos": pos})
    return out

def contract_active_events(*args, **kwargs):
    return calc.contract_active_events(_ENV, *args, **kwargs)

def _ct_stamp_added(rows):
    """العقود اللي اتكتبت في الشيت مباشرة من غير «تاريخ الإضافة» بنختمها بتاريخ أول مرة التطبيق يشوفها —
    عشان الإشعارات الفايتة قبل الإضافة ما تتحسبش «فايتة»."""
    col = rows[0]["pos"].get(CONTRACT_COLS[6]) if rows else None
    if col is None:
        return
    today = str(datetime.now().date())
    changed = False
    for c in rows:
        if c["start"] and not c["added"]:
            contracts_sheet.update_cell(c["row"], col + 1, today)
            c["added"] = datetime.now().date()
            changed = True
    if changed:
        clear_cache(contracts_sheet)

def contract_missed_events(*args, **kwargs):
    return calc.contract_missed_events(_ENV, *args, **kwargs)

def _ct_ack(events):
    ts = now_str()
    for e in events:
        c = e["c"]
        col_ack, col_ts = c["pos"].get(CONTRACT_COLS[4]), c["pos"].get(CONTRACT_COLS[5])
        toks = sorted(c["acked"] | {x["token"] for x in events if x["c"]["row"] == c["row"]})
        if col_ack is not None:
            contracts_sheet.update_cell(c["row"], col_ack + 1, ",".join(toks))
        if col_ts is not None:
            contracts_sheet.update_cell(c["row"], col_ts + 1, ts)
    clear_cache(contracts_sheet)

def render_contracts_tab():
    """تاب «العقود» المنفصل: إدخال وتعديل وحذف العقود (نفس بيانات تاب Contracts في جوجل شيت)."""
    if True:
        st.header("📑 العقود | Contracts")
        st.caption("كل عقد: اسم + تاريخ بداية + تاريخ نهاية (سنوي أو شهري أو أي مدة). كل شهر من البداية بيظهر إشعار «انتهى الشهر N» يومين في أعلى المبيعات، ولحد ما تضغط «تم الاطلاع» التصفّح متوقف. "
                   "الإشعارات بتظهر في تاب المبيعات. ممكن تعدّل العقود هنا أو مباشرة في تاب Contracts في جوجل شيت (التاريخ بصيغة yyyy-mm-dd) | Notices appear in the Sales tab. Edit here or directly in the Contracts Google Sheet tab")
        rows = _ct_read()
        today = datetime.now().date()
        if rows:
            table = []
            for c in rows:
                if not c["start"] or not c["end"] or c["end"] <= c["start"]:
                    status, nxt = "⚠️ تاريخ غير صحيح", ""
                else:
                    months = (c["end"].year - c["start"].year) * 12 + c["end"].month - c["start"].month
                    status = "لم يبدأ" if today < c["start"] else ("منتهي" if today > c["end"] else "جاري")
                    upcoming = [d for _, d, _ in contract_events(c["start"], c["end"]) if d >= today]
                    nxt = str(upcoming[0]) if upcoming else "—"
                    status += f" · {months} شهر"
                table.append({"العقد | Contract": c["name"], "البداية | Start": str(c["start"] or c["start_raw"]), "النهاية | End": str(c["end"] or c["end_raw"]),
                              "الحالة | Status": status, "الإشعار القادم | Next notice": nxt, "ملاحظات | Notes": c["notes"]})
            st.dataframe(pd.DataFrame(table), use_container_width=True, hide_index=True)
        else:
            st.info("لا يوجد عقود مسجلة بعد | No contracts yet")
        with st.form("ct_add_form", clear_on_submit=True):
            c1, c2, c3 = st.columns([2, 1, 1])
            nm = c1.text_input("اسم العقد | Contract name")
            sd = c2.date_input("تاريخ البداية | Start date", value=today, key="ct_sd")
            ed = c3.date_input("تاريخ النهاية | End date", value=today + timedelta(days=365), key="ct_ed")
            nt = st.text_input("ملاحظات (اختياري) | Notes (optional)")
            if st.form_submit_button("➕ إضافة عقد | Add contract"):
                if not nm.strip():
                    st.error("اكتب اسم العقد | Enter a contract name")
                elif ed <= sd:
                    st.error("تاريخ النهاية لازم يكون بعد البداية | End must be after start")
                else:
                    contracts_sheet.append_row([nm.strip(), str(sd), str(ed), nt.strip(), "", "", str(today)], value_input_option="USER_ENTERED")
                    clear_cache(contracts_sheet)
                    st.success("✅ تمت الإضافة | Added")
                    st.rerun()
        if rows:
            d1, d2 = st.columns([3, 1])
            pick = d1.selectbox("حذف عقد | Delete a contract", [""] + [f"{c['row']} — {c['name']}" for c in rows], key="ct_del_pick")
            if d2.button("🗑️ حذف | Delete", key="ct_del_btn", disabled=not pick):
                contracts_sheet.delete_rows(int(pick.split(" — ")[0]))
                clear_cache(contracts_sheet)
                st.rerun()

def _contracts_gate():
    """بتتعرض في أعلى تاب المبيعات. بترجّع False (يعني امنع باقي التاب) لو فيه إشعار عقد شغال ولسه ماتضغطش عليه «تم الاطلاع»."""
    import html as _html
    try:
        events = contract_active_events()
    except Exception as e:
        st.warning(f"⚠️ مقدرتش أقرأ تاب العقود: {e}")
        events = []
    try:
        _ct_stamp_added(_ct_read())
        missed = contract_missed_events()
    except Exception:
        missed = []
    pending = [e for e in events if not e["acked"]]
    for e in events:
        if e["acked"]:
            st.success(f"✅ عقد «{e['c']['name']}» — {e['label']} (تم الاطلاع) | Acknowledged")
    if not pending:
        if missed:
            items = "".join(
                f'<div style="margin:4px 0;">⚠️ فات إشعار عقد «<b>{_html.escape(e["c"]["name"])}</b>»: {_html.escape(e["label"])} '
                f'(بتاريخ <b>{e["date"]}</b>) ولم يتم الاطلاع عليه</div>' for e in missed)
            st.markdown('<div style="background:#b91c1c;color:#fff;border-radius:10px;padding:12px 16px;margin:6px 0;font-size:16px;font-weight:600;">'
                        + items + '</div>', unsafe_allow_html=True)
            if st.button("✅ تم الاطلاع | Acknowledged", key="ct_ack_missed_btn"):
                try:
                    _ct_ack(missed)
                except Exception as e:
                    st.error(f"❌ مقدرتش أسجّل الاطلاع: {e}")
                else:
                    st.rerun()
        return True
    cards = "".join(
        f'<div style="background:#fff;color:#7f1d1d;border-radius:12px;padding:14px 18px;margin:10px auto;max-width:760px;'
        f'text-align:center;border:3px solid #fecaca;">'
        f'<div style="font-size:26px;font-weight:900;">{_html.escape(e["c"]["name"])}</div>'
        f'<div style="font-size:19px;font-weight:700;margin:4px 0;">{_html.escape(e["label"])}</div>'
        f'<div style="font-size:16px;">📅 تاريخ الإشعار <b>{e["date"]}</b> &nbsp;|&nbsp; بداية العقد <b>{e["c"]["start"]}</b> '
        f'&nbsp;|&nbsp; نهاية العقد <b>{e["c"]["end"]}</b> &nbsp;|&nbsp; المتبقي <b>{max(e["days_left"], 0)}</b> يوم</div></div>'
        for e in pending)
    st.markdown(
        '<div style="background:#b91c1c;border:6px solid #7f1d1d;border-radius:18px;padding:22px 16px;margin:6px 0 14px 0;'
        'text-align:center;color:#fff;box-shadow:0 6px 24px rgba(127,29,29,.45);">'
        '<div style="font-size:96px;line-height:1;">🛑</div>'
        '<div style="font-size:54px;font-weight:900;letter-spacing:6px;line-height:1.1;">STOP | قِف</div>'
        '<div style="font-size:22px;font-weight:800;margin:6px 0 10px 0;">إشعار عقد — تصفّح المبيعات متوقف | Contract notice — Sales browsing is blocked</div>'
        + cards +
        '<div style="font-size:16px;margin-top:8px;">اضغط «تم الاطلاع» عشان تكمل | Press «Acknowledged» to continue</div></div>',
        unsafe_allow_html=True)
    if st.button("✅ تم الاطلاع | Acknowledged", key="ct_ack_btn", type="primary", use_container_width=True):
        try:
            _ct_ack(pending)
        except Exception as e:
            st.error(f"❌ مقدرتش أسجّل الاطلاع: {e}")
            return False
        st.rerun()
    return False

if st.session_state["nav_page"] == "tab14":
    with tab14:
        if _tab_gate("tab14", "🛒 المبيعات | Sales") and _contracts_gate():
            st.subheader("🛒 المبيعات اليومية | Daily Sales")
            st.caption("المفتاح الأساسي هو كود التكويد 🏷️ — كل بلوك = SKU نون (مرتب تنازليًا بمبيعات أمس) + ASIN أمازون لنفس الكود جنبه (والكود بيتكرر): "
                       "نون 🟢 | أمازون 🟠 | تريندول 🟣 | "
                       "One block per Noon SKU (sorted by yesterday's sales) with Noon | Amazon | Trendyol side by side")

            _nt_all = get_noon_track()
            apply_noon_our_names(_nt_all)
            _nt_c1, _nt_c2 = st.columns([1, 3])
            with _nt_c1:
                if st.button("🔄 تحديث بيانات المنافسين", key="nt_refresh_btn", use_container_width=True):
                    _nt_all = get_noon_track(force=True)
            with _nt_c2:
                st.caption(f"🔎 بيانات المنافسين من بوت noon_tracker: {len(_nt_all)} SKU نون" if _nt_all
                           else "🔎 لسه مفيش بيانات منافسين — شغّل بوت noon_tracker (بيعمل التابات NoonTrack_* أول مرة)")

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

            if not inv_map and not amz_inv_map_asin and not get_trendyol_bundle()["rows"]:
                st.info("ارفع ملف المخزون أولاً (نون: تاب «رفع ملف» / أمازون: StockAmazon) | Upload Inventory first")
            else:
                rows_fbn_t14, multi_fbn_t14 = _t14_build_rows("fbn", sales_dates, sales_display_days)
                rows_fbb_all_t14, multi_fbb_t14 = _t14_build_rows("fbb", sales_dates, sales_display_days)
                # FBB: بنعرض بس SKUs اللي باعت FBB في الفترة (أو مش في ملف المخزون) — عشان الكارت ميتملاش بصفوف صفرية
                rows_fbb_t14 = [r for r in rows_fbb_all_t14 if r["total_recent"] > 0 or r.get("not_in_inventory")]
                # أمازون: عادي FSAB (كل ASIN في مخزون أمازون + اليتامى اللي باعت) و FBA (اللي باع بس)
                rows_fsab_t14, multi_fsab_t14 = _t14_build_rows_amz("FSAB", sales_dates, sales_display_days)
                rows_fba_all_t14, multi_fba_t14 = _t14_build_rows_amz("FBA", sales_dates, sales_display_days)
                rows_fba_t14 = [r for r in rows_fba_all_t14 if r["total_recent"] > 0 or r.get("not_in_inventory")]
                # تريندول: Barcode = SKU (نفس تكويد نون)
                _dates30_ty = [today_t14 - timedelta(days=i) for i in range(1, 31)]
                multi_ty_t14 = build_daily_orders_counts_trendyol(sales_dates)
                rows_ty_t14 = build_trendyol_sales_rows(multi_ty_t14, build_daily_orders_prices_trendyol(sales_dates),
                                                        build_daily_orders_counts_trendyol(_dates30_ty), sales_dates, sales_display_days)

                # ══ إجماليات اليومية في الأعلى (نون FBN+FBB + أمازون FSAB+FBA) من صفوف الأوردرز الخام ══
                tot_fbn = {d: sum(dc.get(d, 0) for dc in multi_fbn_t14.values()) for d in sales_dates}
                tot_fbb = {d: sum(dc.get(d, 0) for dc in multi_fbb_t14.values()) for d in sales_dates}
                tot_fsab = {d: sum(dc.get(d, 0) for dc in multi_fsab_t14.values()) for d in sales_dates}
                tot_fba = {d: sum(dc.get(d, 0) for dc in multi_fba_t14.values()) for d in sales_dates}
                tot_ty = {d: sum(dc.get(d, 0) for dc in multi_ty_t14.values()) for d in sales_dates}
                st.markdown("#### 📊 إجمالي المبيعات اليومية | Daily Sales Totals")
                total_cols = st.columns(min(len(sales_dates), sales_display_days))
                for ci, (d, lbl) in enumerate(zip(sales_dates, sales_labels)):
                    if ci < len(total_cols):
                        with total_cols[ci]:
                            day_noon = tot_fbn.get(d, 0) + tot_fbb.get(d, 0)
                            day_amz = tot_fsab.get(d, 0) + tot_fba.get(d, 0)
                            day_total = day_noon + day_amz + tot_ty.get(d, 0)
                            is_yesterday = (ci == 0)
                            if is_yesterday:
                                bg    = "#14532d" if day_total > 0 else "#7f1d1d"
                                num_color = "#86efac" if day_total > 0 else "#fca5a5"
                                border = "border:2px solid #22c55e;" if day_total > 0 else "border:2px solid #ef4444;"
                            else:
                                bg    = "#1e293b" if day_total == 0 else "#172554"
                                num_color = "#93c5fd" if day_total > 0 else "#64748b"
                                border = ""
                            st.markdown(_t14_day_tile(
                                bg, border, num_color,
                                f'{"🔴 " if is_yesterday and day_total==0 else ("🟢 " if is_yesterday else "")}{lbl.split("(")[0].strip()}',
                                d.strftime("%m-%d"), day_total, day_noon, tot_fbn.get(d, 0), tot_fbb.get(d, 0),
                                day_amz, tot_fsab.get(d, 0), tot_fba.get(d, 0), tot_ty.get(d, 0)), unsafe_allow_html=True)
                st.divider()

                # ══ تجميع بالكود: code -> {SKUs نون} + {ASINs أمازون} ══
                tac_map_t14 = get_tacweed_map()
                tac_amz_map_t14 = get_tacweed_amazon_map()
                sku_rows_t14 = {}
                for r in rows_fbn_t14:
                    sku_rows_t14.setdefault(r["sku_up"], {})["FBN"] = r
                for r in rows_fbb_t14:
                    sku_rows_t14.setdefault(r["sku_up"], {})["FBB"] = r
                publish_noon_track_skus(list(sku_rows_t14.keys()))
                ty_rows_t14 = {r["sku_up"]: {"TY": r} for r in rows_ty_t14}
                asin_rows_t14 = {}
                for r in rows_fsab_t14:
                    asin_rows_t14.setdefault((r.get("asin") or r["sku_up"]).upper(), {})["FSAB"] = r
                for r in rows_fba_t14:
                    asin_rows_t14.setdefault((r.get("asin") or r["sku_up"]).upper(), {})["FBA"] = r

                srch_t14 = st.text_input("🔍 بحث بالكود أو SKU أو ASIN | Search code / SKU / ASIN", key="srch_t14",
                                         placeholder="اكتب كود التكويد أو SKU أو ASIN...")
                q_t14 = srch_t14.strip().upper()
                # ══ عرض متجر واحد في المرة (أريح للعين) أو الكل زي الشكل القديم ══
                _VIEWS = {"🧩 الكل | All": "all", "🟢 نون | Noon": "noon", "🟠 أمازون | Amazon": "amz",
                          "🟣 تريندول | Trendyol": "ty",
                          "🟢 نون FBN فقط": "g_fbn", "🟢 نون FBB فقط": "g_fbb",
                          "🟠 أمازون FBA فقط": "g_fba", "🟠 أمازون عادي FSAB فقط": "g_fsab"}
                _vw = _VIEWS[st.radio("🏬 اعرض متجر | Show store", list(_VIEWS), horizontal=True, key="t14_view")]
                if _vw.startswith("g_"):
                    render_channel_gap_view(_vw[2:], q_t14)

                def _yv(chs, ch):
                    rr = chs.get(ch)
                    return rr["day_counts"].get(sales_dates[0], 0) if (rr and sales_dates) else 0
                def _tv(chs, ch):
                    rr = chs.get(ch)
                    return rr["total_recent"] if rr else 0
                def _sku_sales(s):
                    c = sku_rows_t14[s]
                    return (_yv(c, "FBN") + _yv(c, "FBB"), _tv(c, "FBN") + _tv(c, "FBB"))

                # ══ كل بلوك = SKU نون واحد + ASIN أمازون + Barcode تريندول لنفس الكود (بالترتيب حسب المبيعات)، والكود بيتكرر عادي ══
                # الترتيب النهائي تنازلي بمبيعات أمس لـ SKU نون (الأكبر فوق). تريندول: زي أمازون بالكود والأعلى مبيعًا (الاسم بس لو مفيش كود).
                # ASIN/Barcode زيادة أو بدون كود = بلوك لوحده. أي ASIN/Barcode بدون مبيعات أمس مبيتعرضش (إلا لو بتبحث).
                def _asin_sales(a):
                    c = asin_rows_t14[a]
                    return (_yv(c, "FSAB") + _yv(c, "FBA"), _tv(c, "FSAB") + _tv(c, "FBA"))
                def _ty_sales(k):
                    c = ty_rows_t14[k]
                    return (_yv(c, "TY"), _tv(c, "TY"))
                asin_sel = [a for a in asin_rows_t14 if q_t14 or _asin_sales(a)[0] > 0]
                ty_sel = [k for k in ty_rows_t14 if q_t14 or _ty_sales(k)[0] > 0]
                blocks_t14, blocks_by_sku = [], {}
                for sku_up in sku_rows_t14:
                    blk = {"code": (tac_map_t14.get(sku_up, "") or "").strip(), "sku": sku_up, "asin": None, "tb": None}
                    blocks_by_sku[sku_up] = blk
                    blocks_t14.append(blk)
                # تريندول زي أمازون: بيتطابق بالكود وبالأعلى مبيعًا. التطابق بالاسم (Barcode = SKU) بس لو الـ Barcode ملوش كود
                ty_left = []
                for k in ty_sel:
                    if (tac_map_t14.get(k, "") or "").strip() or k not in blocks_by_sku:
                        ty_left.append(k)
                    else:
                        blocks_by_sku[k]["tb"] = k
                code_blocks = {}
                for blk in blocks_t14:
                    if blk["code"]:
                        code_blocks.setdefault(blk["code"], []).append(blk)
                for _c in code_blocks:
                    code_blocks[_c].sort(key=lambda x: _sku_sales(x["sku"]), reverse=True)
                asin_by_code, ty_by_code, _pool = {}, {}, []
                for a in asin_sel:
                    _c = (tac_amz_map_t14.get(a, "") or "").strip()
                    if _c:
                        asin_by_code.setdefault(_c, []).append(a)
                    else:
                        _pool.append({"code": "", "sku": None, "asin": a, "tb": None})
                for k in ty_left:
                    _c = (tac_map_t14.get(k, "") or "").strip()
                    if _c:
                        ty_by_code.setdefault(_c, []).append(k)
                    else:
                        _pool.append({"code": "", "sku": None, "asin": None, "tb": k})
                for _c in set(asin_by_code) | set(ty_by_code):
                    aa_ = sorted(asin_by_code.get(_c, []), key=_asin_sales, reverse=True)
                    tt_ = sorted(ty_by_code.get(_c, []), key=_ty_sales, reverse=True)
                    for blk in code_blocks.get(_c, []):
                        if aa_:
                            blk["asin"] = aa_.pop(0)
                        if tt_ and not blk["tb"]:
                            blk["tb"] = tt_.pop(0)
                    for k in range(max(len(aa_), len(tt_))):
                        blocks_t14.append({"code": _c, "sku": None, "asin": aa_[k] if k < len(aa_) else None,
                                           "tb": tt_[k] if k < len(tt_) else None})
                blocks_t14 += _pool
                # أفضل ASIN أمازون لكل كود (حتى لو مباعش أمس) — لعرض رقمه في خانة "مفيش مبيعات أمس"
                _best_asin_code = {}
                for _a in asin_rows_t14:
                    _c = (tac_amz_map_t14.get(_a, "") or "").strip()
                    if not _c:
                        continue
                    _r0 = asin_rows_t14[_a].get("FSAB") or asin_rows_t14[_a].get("FBA")
                    _k = (_tv(asin_rows_t14[_a], "FSAB") + _tv(asin_rows_t14[_a], "FBA"), _r0.get("sales_month", 0))
                    if _c not in _best_asin_code or _k > _best_asin_code[_c][0]:
                        _best_asin_code[_c] = (_k, _a)
                for b in blocks_t14:
                    b["amz_hint"] = _best_asin_code.get(b["code"], (None, None))[1] if (b["code"] and not b["asin"]) else None
                # أفضل Barcode تريندول لكل كود (حتى لو مباعش أمس) — لعرضه في خانة "مفيش مبيعات تريندول أمس"
                _ty_s60 = {}
                for _tr in get_trendyol_bundle()["rows"]:
                    if _tr["ok"]:
                        _ty_s60[_tr["bc"]] = _ty_s60.get(_tr["bc"], 0) + _tr["qty"]
                _ty_known = set(_ty_s60) | {_bb for _v in get_ty_products()["by_cid"].values() for _bb in _v}
                _best_ty_code = {}
                for _bb in sorted(_ty_known):
                    _c = (tac_map_t14.get(_bb, "") or "").strip()
                    if _c and (_c not in _best_ty_code or _ty_s60.get(_bb, 0) > _ty_s60.get(_best_ty_code[_c], 0)):
                        _best_ty_code[_c] = _bb
                for b in blocks_t14:
                    b["ty_hint"] = _best_ty_code.get(b["code"]) if (b["code"] and not b["tb"]) else None
                for b in blocks_t14:
                    ns_ = _sku_sales(b["sku"]) if b["sku"] else (0, 0)
                    as_ = _asin_sales(b["asin"]) if b["asin"] else (0, 0)
                    ts_ = _ty_sales(b["tb"]) if b["tb"] else (0, 0)
                    b["sort"] = (-ns_[0], -ns_[1], -(as_[0] + ts_[0]), -(as_[1] + ts_[1]))
                    keys = [b["code"], b["sku"] or "", b["asin"] or "", b["tb"] or ""]
                    if b["asin"]:
                        r0 = asin_rows_t14[b["asin"]].get("FSAB") or asin_rows_t14[b["asin"]].get("FBA")
                        keys += list(r0.get("mskus") or [r0["sku_up"]])
                    b["hay"] = " ".join(str(k).upper() for k in keys)
                blocks_t14 = [b for b in blocks_t14 if not q_t14 or q_t14 in b["hay"]]
                blocks_t14.sort(key=lambda b: b["sort"])
                if _vw.startswith("g_"):      # عروض القنوات المنفردة بتترسم فوق (render_channel_gap_view) — مفيش بلوكات
                    blocks_t14 = []
                elif _vw == "noon":
                    blocks_t14 = [b for b in blocks_t14 if b["sku"]]
                elif _vw == "amz":
                    blocks_t14 = [b for b in blocks_t14 if b["asin"]]
                    blocks_t14.sort(key=lambda b: (-_asin_sales(b["asin"])[0], -_asin_sales(b["asin"])[1]))
                elif _vw == "ty":
                    blocks_t14 = [b for b in blocks_t14 if b["tb"]]
                    blocks_t14.sort(key=lambda b: (-_ty_sales(b["tb"])[0], -_ty_sales(b["tb"])[1]))

                # ══ تقسيم لصفحات: نرسم أول N بلوك بس (الرسم هو اللي بيبطّئ الصفحة) وزرار «كمان» يزوّد ══
                _sig_t14 = (_vw, q_t14)
                if st.session_state.get("t14_sig") != _sig_t14:
                    st.session_state["t14_sig"] = _sig_t14
                    st.session_state["t14_limit"] = 30
                _lim_t14 = int(st.session_state.get("t14_limit", 30))
                blocks_all_t14 = blocks_t14
                blocks_show_t14 = blocks_all_t14[:_lim_t14]
                # جدول تحميل — بيتبني عند الطلب بس (بدل كل مرة) عشان الصفحة تفتح أسرع
                if blocks_t14:
                    c1, c2 = st.columns(2)
                    with c1:
                        if st.button("📥 جهّز ملف Excel | Prepare Excel", key="t14_dl_prep", use_container_width=True):
                            _dl_rows = []
                            for b in blocks_t14:
                                if b["sku"]:
                                    chs = sku_rows_t14[b["sku"]]
                                    rn, rb = chs.get("FBN"), chs.get("FBB")
                                    base = rn or rb
                                    _dl_rows.append({
                                        "المتجر | Store": "Noon", "تكويد | Code": b["code"], "SKU / ASIN": base["sku"],
                                        **{"FBN " + sales_labels[i]: (rn["day_counts"].get(d, 0) if rn else 0) for i, d in enumerate(sales_dates)},
                                        **{"FBB " + sales_labels[i]: (rb["day_counts"].get(d, 0) if rb else 0) for i, d in enumerate(sales_dates)},
                                        "مخزون | Stock": base["stock"], "مبيع شهري | Monthly Sales": base["sales_month"]})
                                if b["asin"]:
                                    chs = asin_rows_t14[b["asin"]]
                                    rs, rf = chs.get("FSAB"), chs.get("FBA")
                                    base = rs or rf
                                    _dl_rows.append({
                                        "المتجر | Store": "Amazon", "تكويد | Code": b["code"], "SKU / ASIN": f"{b['asin']} ({base['sku']})",
                                        **{"FSAB " + sales_labels[i]: (rs["day_counts"].get(d, 0) if rs else 0) for i, d in enumerate(sales_dates)},
                                        **{"FBA " + sales_labels[i]: (rf["day_counts"].get(d, 0) if rf else 0) for i, d in enumerate(sales_dates)},
                                        "مخزون | Stock": base["stock"], "مبيع شهري | Monthly Sales": base["sales_month"]})
                                if b["tb"]:
                                    chs = ty_rows_t14[b["tb"]]
                                    rt = chs["TY"]
                                    _dl_rows.append({
                                        "المتجر | Store": "Trendyol", "تكويد | Code": b["code"], "SKU / ASIN": rt["sku"],
                                        **{"TY " + sales_labels[i]: rt["day_counts"].get(d, 0) for i, d in enumerate(sales_dates)},
                                        "مخزون | Stock": 0, "مبيع شهري | Monthly Sales": rt["sales_month"]})
                            st.session_state["t14_dl_bytes"] = to_excel(pd.DataFrame(_dl_rows).fillna(0))
                            st.session_state["t14_dl_name"] = f"sales_daily_{file_timestamp()}.xlsx"
                        if st.session_state.get("t14_dl_bytes"):
                            st.download_button("⬇️ تحميل الملف الجاهز | Download", data=st.session_state["t14_dl_bytes"],
                                               file_name=st.session_state.get("t14_dl_name", "sales_daily.xlsx"),
                                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                               use_container_width=True, key="dlbtn_t14")
                    with c2: st.info(f"📦 بلوكات: {len(blocks_t14)} | 📅 {sales_display_days} يوم")

                # ══ قائمة SKUs المرحلة من المبيعات (محتاج جدولة فقط) ══
                _new_transferred = []

                # ══ خريطة SKUs المجدولة خلال آخر 4 أيام (لعرض ASN + الكمية لو فعلاً ليها جدولة) ══
                recent_sched_map_t14 = get_recent_schedule_rows(days_back=4)
                pending_approval_skus_t14 = get_pending_approval_skus()
                ads_map_t14 = get_ads_map()
                ads_map_amz_t14 = get_ads_map_amazon()
                com_map_t14 = get_com_map()
                live_map_t14 = get_live_map()
                xdock_threshold_t14 = int(load_settings().get("xdock_low_stock_threshold","10") or 10)
                # تريندول: العمولة من الملف، والتوصيل من إعداد trendyol_delivery_fee في Settings (لو مش موجود = 0)
                _ty_fee_raw = str(load_settings().get("trendyol_delivery_fee", "")).strip()
                ty_fee_set_t14 = bool(_ty_fee_raw)
                try:
                    _ty_fee_v = -abs(float(_ty_fee_raw)) if _ty_fee_raw else 0.0
                except Exception:
                    _ty_fee_v, ty_fee_set_t14 = 0.0, False
                com_map_ty_t14 = {bc: {"delivery": _ty_fee_v, "commission_pct": pct} for bc, pct in get_trendyol_commission_map().items()}

                def _t14_noon_shown(noon_chs):
                    """أي قنوات نون بتتعرض للبلوك (FBN/FBB) — نفس القاعدة بالظبط للرسم وللحساب من غير رسم."""
                    _sn = {k: bool(noon_chs and k in noon_chs) for k in ("FBN", "FBB")}
                    # نون: القناة اللي مفيهاش مبيع أمس مبتتعرضش لو القناة التانية باعت؛ ولو الاتنين صفر بنعرض FBN (أو FBB لو هي الوحيدة)
                    if noon_chs and not q_t14:
                        _yn = {k: (_yv(noon_chs, k) if _sn[k] else 0) for k in ("FBN", "FBB")}
                        if _yn["FBN"] > 0 or _yn["FBB"] > 0:
                            _sn = {k: _sn[k] and _yn[k] > 0 for k in ("FBN", "FBB")}
                        elif _sn["FBN"]:
                            _sn["FBB"] = False
                    return _sn

                def _t14_transfer_entry(r):
                    """لو الـ SKU "محتاج جدولة فقط" بيرجّع بياناته للترحيل (نفس شرط _t14_channel_part من غير رسم)."""
                    badge_text, _c, sched = schedule_coverage_badge(r["sku"], r["days_to_stockout"], delay_days_t14)
                    stock_self_ok = (r["effective_avg"] <= 0) or (r["days_to_stockout"] >= coverage_days_t14)
                    if (not stock_self_ok and badge_text and "محتاج جدولة" in badge_text and not sched
                            and not get_unavailable_ordered_note(r["sku"]) and not recent_sched_map_t14.get(r["sku_up"])):
                        return {"sku": r["sku"], "sku_up": r["sku_up"], "stock": r["stock"],
                                "sales_month": r["sales_month"], "img": r["img"],
                                "effective_avg": r["effective_avg"], "days_to_stockout": r["days_to_stockout"],
                                "day_counts": r["day_counts"]}
                    return None

                _t14_bn = [0]
                st.markdown("<style>[class*='st-key-t14blk_'] [data-testid='stVerticalBlock']{gap:.35rem !important;}"
                            "[class*='st-key-t14blk_'] [data-testid='stMarkdownContainer'] p{margin-bottom:.1rem;}</style>", unsafe_allow_html=True)
                render_china_sales_status([k for b in blocks_t14 for k in (b["sku"], b["asin"], b["tb"]) if k])  # 🚢 حالة ربط بيانات الصين
                st.divider()
                wh_lookup_t14 = _china_wh_lookup()
                for b in blocks_show_t14:
                    noon_chs = sku_rows_t14.get(b["sku"]) if b["sku"] else None
                    amz_chs = asin_rows_t14.get(b["asin"]) if b["asin"] else None
                    ty_chs = ty_rows_t14.get(b["tb"]) if b["tb"] else None
                    if _vw != "all":      # متجر واحد بس: نخفي بيانات المتاجر التانية من البلوك
                        noon_chs = noon_chs if _vw == "noon" else None
                        amz_chs = amz_chs if _vw == "amz" else None
                        ty_chs = ty_chs if _vw == "ty" else None
                    first_key = (b["sku"] if noon_chs else None) or (b["asin"] if amz_chs else None) or (b["tb"] if ty_chs else None) or b["sku"] or b["asin"] or b["tb"]
                    first_row = ((noon_chs.get("FBN") or noon_chs.get("FBB")) if noon_chs
                                 else ((amz_chs.get("FSAB") or amz_chs.get("FBA")) if amz_chs else ty_chs["TY"]))
                    # الصور على أساس الـ SKU: صورة SKU نون، وبعدين صورة أمازون (MSKU ثم ASIN)
                    img_t14 = ((first_row.get("img") if noon_chs else "")
                               or (next((amz_chs.get(k).get("img") for k in ("FSAB", "FBA") if amz_chs.get(k) and amz_chs.get(k).get("img")), "") if amz_chs else "")
                               or (ty_chs["TY"].get("img") if ty_chs else ""))
                    y_n = _yv(noon_chs, "FBN") if noon_chs else 0; y_b = _yv(noon_chs, "FBB") if noon_chs else 0
                    t_n = _tv(noon_chs, "FBN") if noon_chs else 0; t_b = _tv(noon_chs, "FBB") if noon_chs else 0
                    y_a = _yv(amz_chs, "FSAB") if amz_chs else 0; y_f = _yv(amz_chs, "FBA") if amz_chs else 0
                    t_a = _tv(amz_chs, "FSAB") if amz_chs else 0; t_f = _tv(amz_chs, "FBA") if amz_chs else 0
                    y_t = _yv(ty_chs, "TY") if ty_chs else 0; t_t = _tv(ty_chs, "TY") if ty_chs else 0
                    c_img, c_info = st.columns([1, 7])
                    with c_img:
                        show_img(img_t14, 70)
                    with c_info:
                        # ── مستوى الكود: شارات الكود/المستودع/التكلفة/النفاد (محسوبة من سحب نون + أمازون) ──
                        tc_badge_t14 = warehouse_available_badge_amazon(first_key) if (not b["sku"] and b["asin"]) else warehouse_available_badge(first_key)
                        if tc_badge_t14:
                            st.markdown(tc_badge_t14, unsafe_allow_html=True)
                        else:
                            st.markdown('<span class="tacweed-badge" style="background:#374151;color:#e5e7eb;">🏷️ بدون تكويد</span>',
                                        unsafe_allow_html=True)
                        wh_g = wh_lookup_t14.get(first_key)
                        if wh_g and wh_g.get("daily_draw") is not None:
                            st.caption(f"🧮 سحب المستودع اليومي للكود (آخر 30 يوم): نون {wh_g.get('draw_noon', 0):g} + أمازون {wh_g.get('draw_amz', 0):g} + تريندول {wh_g.get('draw_ty', 0):g} "
                                       f"= {wh_g['daily_draw']:g} قطعة/يوم — شارة النفاد وموعد الطلب محسوبين على المجموع")
                        render_china_for_sku(first_key, first_row)  # 🚢 كمية/شحنات الصين القادمة (مربوطة بالكود)
                        if noon_chs and b["sku"]:
                            render_noon_rivals(b["sku"], (live_map_t14.get(b["sku"].upper()) or {}).get("price"))
                            render_ads_reco_chip(b["sku"])

                    # ══ نون | أمازون | تريندول جنب بعض — كل قسم في صف مستقل عشان الأقسام تتحاذى (السعر جنب السعر) ══
                    r_sku = r_asin = r_ty = None
                    if ty_chs:
                        r_ty = ty_chs["TY"]
                    if noon_chs:
                        r_any = noon_chs.get("FBN") or noon_chs.get("FBB")
                        r_sku = dict(r_any, day_prices={d: list(noon_chs["FBN"]["day_prices"].get(d, [])) if "FBN" in noon_chs else []
                                                        for d in sales_dates})
                        if "FBB" in noon_chs:
                            for d in sales_dates:
                                r_sku["day_prices"][d] = r_sku["day_prices"].get(d, []) + list(noon_chs["FBB"]["day_prices"].get(d, []))
                        # عدد القطع المباعة لكل نوع تنفيذ في فترة العرض — بيستخدمه بلوك الربح لتطبيق عمولة/توصيل كل نوع
                        r_sku["_mix"] = {}
                        for _k in ("FBN", "FBB"):
                            if _k in noon_chs:
                                _idx = next((_i for _i, _d in enumerate(sales_dates)
                                             if float(noon_chs[_k]["day_counts"].get(_d, 0) or 0) > 0), None)
                                r_sku["_mix"][_k] = {"idx": _idx,
                                                     "units": sum(float(noon_chs[_k]["day_counts"].get(_d, 0) or 0) for _d in sales_dates)}
                    if amz_chs:
                        r_any = amz_chs.get("FSAB") or amz_chs.get("FBA")
                        r_asin = dict(r_any, day_prices={d: list(amz_chs["FSAB"]["day_prices"].get(d, [])) if "FSAB" in amz_chs else []
                                                         for d in sales_dates})
                        if "FBA" in amz_chs:
                            for d in sales_dates:
                                r_asin["day_prices"][d] = r_asin["day_prices"].get(d, []) + list(amz_chs["FBA"]["day_prices"].get(d, []))
                    def _cols3():
                        if _vw == "all":
                            return st.columns([5, 5, 4])
                        _one = st.container()      # متجر واحد: عرض كامل بدل 3 أعمدة
                        return _one, _one, _one
                    _t14_bn[0] += 1
                    with st.container(border=True, key=f"t14blk_{_t14_bn[0]}"):
                        # صف 0: عناوين المتاجر (تريندول مكانه محجوز هنا)
                        cn, ca, ct = _cols3()
                        with cn:
                            if _vw in ("all", "noon"):
                                st.markdown('<span class="wh-badge" style="background:#064e3b;color:#6ee7b7;">' + plat_badge("noon", 16) + 'نون | Noon</span>', unsafe_allow_html=True)
                                if noon_chs:
                                    _t14_yday_line([("تخزين FBN", y_n), ("عادي FBB", y_b)], "noon")
                                if not noon_chs:
                                    st.caption("— مفيش SKU نون لهذا الكود | no Noon SKU for this code")
                        with ca:
                            if _vw in ("all", "amz"):
                                st.markdown('<span class="wh-badge" style="background:#7c2d12;color:#fed7aa;">' + plat_badge("amazon", 16) + 'أمازون | Amazon</span>', unsafe_allow_html=True)
                                if amz_chs:
                                    _t14_yday_line([("تخزين FBA", y_f), ("عادي FSAB", y_a)], "amz")
                                if not amz_chs:
                                    if b.get("amz_hint"):
                                        if _asin_sales(b["amz_hint"])[0] > 0:
                                            st.caption("— ASIN هذا الكود ظاهر مع SKU آخر لنفس الكود | shown with another SKU of this code")
                                        else:
                                            st.caption("— مفيش مبيعات أمازون أمس | no Amazon sales yesterday")
                                    else:
                                        st.caption("— مفيش ASIN أمازون مربوط بالكود ده | no Amazon ASIN linked to this code")
                        with ct:
                            if _vw in ("all", "ty"):
                                st.markdown('<span class="wh-badge" style="background:#4c1d95;color:#ddd6fe;">' + plat_badge("trendyol", 16) + 'تريندول | Trendyol</span>', unsafe_allow_html=True)
                                if r_ty:
                                    _t14_yday_line([("", y_t)], "ty")
                                if not r_ty:
                                    if get_trendyol_bundle()["rows"]:
                                        st.caption("— مفيش مبيعات تريندول أمس | no Trendyol sales yesterday")
                                    else:
                                        _t14_trendyol_placeholder()
                        # قنوات المتجر المعروضة — بنحسبها هنا عشان كبسولات التغطية/الجدولة تطلع فوق جنب الهوية
                        _shown_a = {_ach: bool(amz_chs and _ach in amz_chs and (q_t14 or _yv(amz_chs, _ach) > 0)) for _ach in ("FSAB", "FBA")}
                        _sn = _t14_noon_shown(noon_chs)
                        _sa = _shown_a
                        _lvl = {"FBN": True, "FBB": not _sn["FBN"], "FSAB": True, "FBA": not _sa["FSAB"]}
                        def _top_cov(chlist, src, shown):
                            _sink = []
                            _COV_SINK[0] = _sink
                            try:
                                for _c in chlist:
                                    if src and shown.get(_c):
                                        _t14_channel_part(src[_c], _c, _lvl[_c], part="cov")
                            finally:
                                _COV_SINK[0] = None
                            if _sink:
                                st.markdown("".join(_sink), unsafe_allow_html=True)
                        # صف 1: الهوية (SKU / ASIN) + كبسولات التغطية والجدولة جنبها
                        cn, ca, ct = _cols3()
                        with cn:
                            if r_sku: _t14_noon_part(r_sku, "ident")
                            _top_cov(["FBN", "FBB"], noon_chs, _sn)
                        with ca:
                            if r_asin:
                                _t14_amz_part(r_asin, "ident")
                            elif _vw in ("all", "amz") and not amz_chs and b.get("amz_hint"):      # ASIN مربوط بس مباعش أمس — في نفس سطر SKU/Barcode
                                st.markdown(f"**ASIN:** {asin_link_html(b['amz_hint'])}", unsafe_allow_html=True)
                        with ca:
                            _top_cov(["FSAB", "FBA"], amz_chs, _sa)
                        with ct:
                            if r_ty:
                                _t14_ty_part(r_ty, "ident")
                            elif _vw in ("all", "ty") and b.get("ty_hint") and get_trendyol_bundle()["rows"]:   # Barcode مربوط بس مباعش أمس
                                st.markdown(
                                    f'**Barcode:** <span style="font-family:monospace;font-weight:700;color:#e2e8f0;background:#1e293b;'
                                    f'border:1px solid #334155;border-radius:6px;padding:2px 10px;user-select:all;-webkit-user-select:all;">{b["ty_hint"]}</span>',
                                    unsafe_allow_html=True)
                                _nm_h = get_ty_products()["name"].get(b["ty_hint"]) or get_trendyol_bundle()["name"].get(b["ty_hint"])
                                if _nm_h:
                                    st.caption(_nm_h[:70])
                        # صف 2: صندوق السعر والربح — صف 3: تفاصيل الحساب — صف 4: تفاصيل حكم الإعلان — صف 5: أداء الإعلانات
                        _has_ads_n = bool(r_sku and ads_map_t14.get(r_sku["sku_up"]))
                        _has_ads_a = bool(r_asin and ads_map_amz_t14.get((r_asin.get("asin") or "").upper()))
                        _has_ads_t = bool(r_ty and get_ads_by_barcode_trendyol().get(r_ty["sku_up"]))
                        # صندوق السعر والربح وتحته «الحساب | الحكم | إعلانات» جنب بعض — في صف واحد لكل متجر (من غير فراغ بينهم)
                        def _exp_row(fn, r, has_ads, extra=()):
                            _pl = ["calc"] + (["verdict", "ads"] if has_ads else []) + list(extra)
                            for _ec, _pt in zip(st.columns(len(_pl)), _pl):
                                with _ec:
                                    fn(r, _pt)
                        cn, ca, ct = _cols3()
                        with cn:
                            if r_sku:
                                _t14_noon_part(r_sku, "box"); _exp_row(_t14_noon_part, r_sku, _has_ads_n)
                        with ca:
                            if r_asin:
                                _t14_amz_part(r_asin, "box"); _exp_row(_t14_amz_part, r_asin, _has_ads_a)
                        with ct:
                            if r_ty:
                                _t14_ty_part(r_ty, "box"); _exp_row(_t14_ty_part, r_ty, _has_ads_t)
                        # قنوات المبيعات: (FBN | FSAB) ثم (FBB | FBA) — كل جزء (مبيعات / مخزون / تغطية) في صف مستقل عشان يتحاذوا.
                        # أمازون: القناة اللي مفيهاش مبيع أمس مبتتعرضش (مفيش داعي لأيام كلها صفر).
                        _shown_a = {}
                        for _ach in ("FSAB", "FBA"):
                            _shown_a[_ach] = bool(amz_chs and _ach in amz_chs and (q_t14 or _yv(amz_chs, _ach) > 0))
                        # كل قنوات المتاجر جنب بعض في صف واحد: FBN | FBB (نون) | FSAB | FBA (أمازون) | تريندول (محجوز)
                        _sn = _t14_noon_shown(noon_chs)
                        _sa = _shown_a
                        if any(_sn.values()) or any(_sa.values()) or r_ty:
                            _lvl = {"FBN": True, "FBB": not _sn["FBN"], "FSAB": True, "FBA": not _sa["FSAB"]}
                            _src = {"FBN": noon_chs, "FBB": noon_chs, "FSAB": amz_chs, "FBA": amz_chs}
                            _chlist = {"all": ["FBN", "FBB", "FSAB", "FBA"], "noon": ["FBN", "FBB"], "amz": ["FSAB", "FBA"], "ty": []}[_vw]
                            _ty_on = _vw in ("all", "ty")
                            _wid = [2.5] * len(_chlist) + ([4] if _ty_on else [])
                            for _pt in ("sales", "stock"):
                                _cols = st.columns(_wid)
                                for _col, _ch in zip(_cols[:len(_chlist)], _chlist):
                                    with _col:
                                        if (_sn.get(_ch) if _ch in _sn else _sa.get(_ch)):
                                            _t14_channel_part(_src[_ch][_ch], _ch, _lvl[_ch], part=_pt)
                                if _ty_on:
                                    with _cols[-1]:
                                        if r_ty: _t14_channel_part(r_ty, "TY", True, part=_pt)
                    st.divider()
                # حفظ المرحلين في session_state بعد اكتمال العرض
                if _vw in ("all", "noon") and not q_t14:      # بيتحسب من بلوكات نون بس — فمنكتبش فوقه وإحنا بنعرض أمازون/تريندول أو بحث
                    for b in blocks_all_t14[len(blocks_show_t14):]:     # البلوكات اللي ما اترسمتش: نحسبها من غير رسم
                        _nc = sku_rows_t14.get(b["sku"]) if b["sku"] else None
                        if not _nc:
                            continue
                        _snh = _t14_noon_shown(_nc)
                        for _chh in ("FBN", "FBB"):
                            if _snh[_chh]:
                                _te = _t14_transfer_entry(_nc[_chh])
                                if _te:
                                    _new_transferred.append(_te)
                    st.session_state["transferred_skus_t14"] = _new_transferred
                if len(blocks_all_t14) > len(blocks_show_t14):
                    _rest = len(blocks_all_t14) - len(blocks_show_t14)
                    st.info(f"معروض {len(blocks_show_t14)} من {len(blocks_all_t14)} بلوك (مرتبين بمبيعات أمس) | Showing {len(blocks_show_t14)} of {len(blocks_all_t14)}")
                    cm1, cm2 = st.columns(2)
                    with cm1:
                        if st.button(f"⬇️ كمان 30 بلوك | 30 more (باقي {_rest})", key="t14_more", use_container_width=True):
                            st.session_state["t14_limit"] = _lim_t14 + 30
                            st.rerun()
                    with cm2:
                        if st.button(f"📜 اعرض الكل ({_rest} بلوك) — أبطأ | Show all (slower)", key="t14_all", use_container_width=True):
                            st.session_state["t14_limit"] = len(blocks_all_t14)
                            st.rerun()


if st.session_state["nav_page"] == "tab_dash":
    with tab_dash:
        if _tab_gate("tab_dash", "📊 داشبورد المبيعات | Sales Dashboard"):
            st.header("📊 داشبورد المبيعات | Sales Dashboard")
            st.caption("تحليلات موسّعة على بيانات المبيعات مع صور المنتجات — منفصلة تمامًا عن تاب المبيعات الأصلي ولا تؤثر عليه | Extended sales analytics with product images — fully separate from, and does not affect, the original Sales tab")

            if not inv_map and not amz_inv_map_asin and not get_trendyol_bundle()["rows"]:
                st.info("ارفع ملف المخزون (نون أو أمازون) أو مبيعات تريندول أولاً | Upload Inventory (Noon/Amazon) or Trendyol sales first")
            else:
                _DSUB = ["🔷 الكل (FBN+FBB) | All Combined", "🅽 مبيعات نون FBN | Noon FBN Sales", "🅱 مبيعات نون FBB | Noon FBB Sales",
                         "🟠 امازون الكل | Amazon All", "🅽 امازون عادي (FSAB) | Amazon Normal", "📦 امازون تخزين FBA | Amazon FBA Storage",
                         "🟣 تريندول | Trendyol", "🔀 تحليل الأكواد عبر المتاجر | Cross-store"]
                # راديو بدل st.tabs: بيرسم القسم المختار بس (أسرع بكتير) | radio renders only the chosen section
                _dsel = st.radio("القسم | Section", _DSUB, horizontal=True, key="dash_sub", label_visibility="collapsed")
                _di = _DSUB.index(_dsel)
                if _di == 0:
                    # ── الكل مع بعض: كل الطلبات (FBN + FBB) من غير أي فلترة على نوع
                    #    التنفيذ — عشان في النهاية هو متجر نون واحد ومبيعات واحدة | All
                    #    orders combined (FBN + FBB) with no fulfillment-type filter —
                    #    it's one Noon store and one combined sales figure.
                    if not inv_map:
                        st.info("ارفع ملف مخزون نون أولاً من تاب «رفع ملف» | Upload Noon Inventory first")
                    else:
                        _render_sales_dashboard_body(
                            build_daily_orders_counts, build_daily_orders_prices,
                            build_daily_orders_family_stats, "all")
                if _di == 1:
                    # ── FBN بس: مستبعد منه أي طلب مصنّف FBB (Fulfillment=FBP)،
                    #    عشان FBN + FBB = تاب "الكل" بالظبط من غير أي تكرار |
                    #    FBN only: excludes anything classified as FBB
                    #    (Fulfillment=FBP), so FBN + FBB always adds up exactly to
                    #    "All Combined" with no double-counting.
                    if not inv_map:
                        st.info("ارفع ملف مخزون نون أولاً من تاب «رفع ملف» | Upload Noon Inventory first")
                    else:
                        _render_sales_dashboard_body(
                            build_daily_orders_counts_fbn, build_daily_orders_prices_fbn,
                            build_daily_orders_family_stats_fbn, "fbn")
                if _di == 2:
                    if not inv_map:
                        st.info("ارفع ملف مخزون نون أولاً من تاب «رفع ملف» | Upload Noon Inventory first")
                    else:
                        _render_sales_dashboard_body(
                            build_daily_orders_counts_fbb, build_daily_orders_prices_fbb,
                            build_daily_orders_family_stats_fbb, "fbb")
                # ══ تحليلات أمازون — منفصلة تمامًا عن تحليلات نون فوق، بتقرا من شيت
                #    DailyOrdersAmazon ومخزون StockAmazon بس (مش inv_map بتاع نون خالص)،
                #    ومقسّمة نفس فكرة الكل/عادي/FBA زي تاب مبيعات امازون بالظبط. الأرقام
                #    مجمّعة على ASIN (مش MSKU) عشان تتربط صح بمخزون أمازون | Amazon
                #    analytics — fully separate from the Noon analytics above, reads
                #    DailyOrdersAmazon and StockAmazon stock only (never Noon's inv_map),
                #    split the same way (All/Normal/FBA) as the Amazon Sales tab itself.
                #    Grouped by ASIN (not MSKU) so it joins correctly against Amazon stock.
                if _di == 3:
                    if not amz_inv_map_asin:
                        st.info("ارفع ملف مخزون أمازون (StockAmazon) أولاً | Upload Amazon Stock (StockAmazon) first")
                    else:
                        _render_sales_dashboard_body(
                            build_daily_orders_counts_amazon_all_asin, build_daily_orders_prices_amazon_all_asin,
                            build_daily_orders_family_stats_amazon_all, "amz_all", channel="amazon")
                if _di == 4:
                    if not amz_inv_map_asin:
                        st.info("ارفع ملف مخزون أمازون (StockAmazon) أولاً | Upload Amazon Stock (StockAmazon) first")
                    else:
                        _render_sales_dashboard_body(
                            build_daily_orders_counts_amazon_normal_asin, build_daily_orders_prices_amazon_normal_asin,
                            build_daily_orders_family_stats_amazon_normal, "amz_normal", channel="amazon")
                if _di == 5:
                    if not amz_inv_map_asin:
                        st.info("ارفع ملف مخزون أمازون (StockAmazon) أولاً | Upload Amazon Stock (StockAmazon) first")
                    else:
                        _render_sales_dashboard_body(
                            build_daily_orders_counts_amazon_fba_asin, build_daily_orders_prices_amazon_fba_asin,
                            build_daily_orders_family_stats_amazon_fba, "amz_fba", channel="amazon")
                if _di == 6:
                    plat_header("trendyol", "داشبورد تريندول | Trendyol Dashboard")
                    if not get_trendyol_bundle()["rows"]:
                        st.info("ارفع ملف طلبات تريندول من تاب «رفع ملف» أولاً | Upload Trendyol orders from the Sales tab first")
                    else:
                        _render_sales_dashboard_body(
                            build_daily_orders_counts_trendyol, build_daily_orders_prices_trendyol,
                            build_daily_orders_family_stats_trendyol, "ty_all", channel="trendyol")
                if _di == 7:
                    render_cross_store_tab()

        # ══ TAB 15 — تحليل الجدولة ══
        # ══ TAB 16 — مخزون بدون بيع ══
if st.session_state["nav_page"] == "tab16":
    with tab16:
        if _tab_gate("tab16", "📦 مخزون بدون بيع | No Sales"):
            st.subheader("📦 مخزون بدون بيع | Stock With No Sales")
            st.caption("SKUs موجودة في المخزون لكن ما بيعت في الفترة المحددة — نون وأمازون كل واحد في تابه، مستقلين تمامًا عن بعض | SKUs in inventory with no sales in the selected period — Noon and Amazon each in their own sub-tab, fully independent")

            def render_no_sale_list(rows, period_label, dl_key, channel="noon"):
                if not rows:
                    st.success(f"✅ لا يوجد {'SKUs' if channel=='noon' else 'ASINs'} بدون مبيعات في {period_label} | No {'SKUs' if channel=='noon' else 'ASINs'} without sales in {period_label}")
                    return
                df_ns = pd.DataFrame([{
                    ("ASIN" if channel=="amazon" else "SKU"): (r["asin"] if channel=="amazon" else r["sku"]),
                    **({"SKU": r["sku"]} if channel=="amazon" else {}),
                    "مخزون | Stock": r["stock"],
                    "مبيع شهري | Monthly Sales": r["sales_month"],
                } for r in rows])
                c1, c2 = st.columns(2)
                with c1: dl_btn(df_ns, dl_key, key=f"dlbtn_{dl_key}")
                with c2: st.warning(f"⚠️ {len(rows)} {'SKU' if channel=='noon' else 'ASIN'} بدون مبيعات | {'SKUs' if channel=='noon' else 'ASINs'} without sales")
                _pgk = f"pg_ns_{dl_key}"
                _pg_items = list(rows)
                for r in page_slice(_pg_items, _pgk):
                    c_img, c_info = st.columns([1, 6])
                    with c_img:
                        show_img(r["img"], 60)
                    with c_info:
                        if channel == "amazon":
                            st.markdown(f"**ASIN:** `{r['asin']}` &nbsp;|&nbsp; **SKU:** `{r['sku']}`", unsafe_allow_html=True)
                        else:
                            st.markdown(f"**SKU:** `{r['sku']}`", unsafe_allow_html=True)
                        st.markdown(
                            f"📦 **مخزون:** {r['stock']} &nbsp;|&nbsp; 📈 **شهري:** {r['sales_month']}",
                        )
                        if channel == "noon":
                            sched_ns = get_latest_schedule_info(r["sku"])
                            if sched_ns:
                                arrival_ns = (sched_ns["parsed"] + timedelta(days=int(load_settings().get("schedule_delay_days","3") or 3))).date() if sched_ns.get("parsed") else None
                                st.caption(f"📅 ASN {sched_ns['asn']} بتاريخ {sched_ns['date']}" + (f" — وصول: {arrival_ns}" if arrival_ns else ""))
                        else:
                            render_recent_expired_note(r["sku"])
                        for note in get_unavailable_ordered_note(r["sku"]):
                            st.markdown(big_note_html(note), unsafe_allow_html=True)
                    st.divider()

                page_more(_pg_items, _pgk)
            channel_tab_noon, channel_tab_amz, channel_tab_ty = st.tabs([
                "🅽 نون | Noon",
                "🅰️ أمازون | Amazon",
                "🟣 تريندول | Trendyol",
            ])

            # ══════════ نون ══════════
            with channel_tab_noon:
                if not inv_map:
                    st.info("ارفع ملف المخزون أولاً من تاب «رفع ملف» | Upload Inventory first")
                else:
                    today_t16 = datetime.now().date()

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

                    sub1, sub2, sub3 = st.tabs([
                        f"📅 بدون مبيع أمس ({len(no_sale_1d)})",
                        f"📅 بدون مبيع آخر 3 أيام ({len(no_sale_3d)})",
                        f"📅 بدون مبيع آخر أسبوع ({len(no_sale_7d)})",
                    ])
                    with sub1:
                        render_no_sale_list(no_sale_1d, "أمس", "no_sale_1d", channel="noon")
                    with sub2:
                        render_no_sale_list(no_sale_3d, "آخر 3 أيام", "no_sale_3d", channel="noon")
                    with sub3:
                        render_no_sale_list(no_sale_7d, "آخر أسبوع", "no_sale_7d", channel="noon")

            # ══════════ أمازون (تخزين FBA بس — مستقل تمامًا عن نون) ══════════
            with channel_tab_amz:
                st.caption("بيقرا بس من تخزين FBA (تخزين عادي/FSAB مستبعد) — نفس فلتر مراجعة مخزون/مبيعات أمازون | Reads FBA storage only (Normal/FSAB excluded) — same filter as the Amazon review tabs")
                if not amz_inv_map_asin:
                    st.info("ارفع ملف مخزون أمازون (StockAmazon) أولاً | Upload Amazon Stock (StockAmazon) first")
                else:
                    today_t16_amz = datetime.now().date()

                    dates_1d_amz  = [today_t16_amz - timedelta(days=1)]
                    dates_3d_amz  = [today_t16_amz - timedelta(days=i) for i in range(1, 4)]
                    dates_7d_amz  = [today_t16_amz - timedelta(days=i) for i in range(1, 8)]
                    dates_30d_amz = [today_t16_amz - timedelta(days=i) for i in range(1, 31)]

                    all_dates_t16_amz = list({d for d in dates_1d_amz + dates_3d_amz + dates_7d_amz})
                    counts_t16_amz         = build_daily_orders_counts_amazon_fba_asin(all_dates_t16_amz)
                    monthly_counts_t16_amz = build_daily_orders_counts_amazon_fba_asin(dates_30d_amz)

                    links_map_asin_t16 = get_links_map_asin()
                    links_map_sku_t16  = get_links_map()

                    def asin_sold_in_amz(asin_up, dates_list):
                        dc = counts_t16_amz.get(asin_up, {})
                        return sum(dc.get(d, 0) for d in dates_list) > 0

                    no_sale_1d_amz, no_sale_3d_amz, no_sale_7d_amz = [], [], []
                    for asin_up, info in amz_inv_map_asin.items():
                        stock    = info.get("stock", 0)
                        mskus    = info.get("skus", [])
                        sku_disp = ", ".join(mskus) if mskus else asin_up
                        mc = monthly_counts_t16_amz.get(asin_up, {})
                        sales_month = sum(mc.values()) if mc else 0
                        img = links_map_asin_t16.get(asin_up, "") or (links_map_sku_t16.get(mskus[0].upper(), "") if mskus else "")
                        row_t16_amz = {
                            "sku": sku_disp, "sku_up": asin_up, "asin": info.get("asin", asin_up),
                            "stock": stock, "sales_month": sales_month, "img": img,
                        }
                        if not asin_sold_in_amz(asin_up, dates_1d_amz):
                            no_sale_1d_amz.append(row_t16_amz)
                        if not asin_sold_in_amz(asin_up, dates_3d_amz):
                            no_sale_3d_amz.append(row_t16_amz)
                        if not asin_sold_in_amz(asin_up, dates_7d_amz):
                            no_sale_7d_amz.append(row_t16_amz)
                    no_sale_1d_amz.sort(key=lambda x: -x["stock"])
                    no_sale_3d_amz.sort(key=lambda x: -x["stock"])
                    no_sale_7d_amz.sort(key=lambda x: -x["stock"])

                    sub1a, sub2a, sub3a = st.tabs([
                        f"📅 بدون مبيع أمس ({len(no_sale_1d_amz)})",
                        f"📅 بدون مبيع آخر 3 أيام ({len(no_sale_3d_amz)})",
                        f"📅 بدون مبيع آخر أسبوع ({len(no_sale_7d_amz)})",
                    ])
                    with sub1a:
                        render_no_sale_list(no_sale_1d_amz, "أمس", "no_sale_1d_amz", channel="amazon")
                    with sub2a:
                        render_no_sale_list(no_sale_3d_amz, "آخر 3 أيام", "no_sale_3d_amz", channel="amazon")
                    with sub3a:
                        render_no_sale_list(no_sale_7d_amz, "آخر أسبوع", "no_sale_7d_amz", channel="amazon")

            # ══════════ تريندول ══════════
            with channel_tab_ty:
                plat_header("trendyol")
                st.caption("منتجات (Barcode) اتباعت في آخر 60 يوم من تريندول ومفيهاش مبيعات في الفترة — مخزون المستودع من كود التكويد | Trendyol barcodes sold in the last 60 days but not in the period — warehouse stock via the Tacweed code")
                if not get_trendyol_bundle()["rows"]:
                    st.info("ارفع ملف طلبات تريندول من تاب «رفع ملف» أولاً | Upload Trendyol orders first")
                else:
                    _ns1, _ns3, _ns7 = (compute_no_sale_rows_trendyol(1), compute_no_sale_rows_trendyol(3), compute_no_sale_rows_trendyol(7))
                    _t1, _t3, _t7 = st.tabs([f"📅 بدون مبيع أمس ({len(_ns1)})", f"📅 بدون مبيع آخر 3 أيام ({len(_ns3)})", f"📅 بدون مبيع آخر أسبوع ({len(_ns7)})"])
                    with _t1:
                        render_no_sale_list_trendyol(_ns1, "أمس", "no_sale_1d_ty")
                    with _t3:
                        render_no_sale_list_trendyol(_ns3, "آخر 3 أيام", "no_sale_3d_ty")
                    with _t7:
                        render_no_sale_list_trendyol(_ns7, "آخر أسبوع", "no_sale_7d_ty")

# ══ TAB CHINA — شحنات الصين (تاب جديدة منفصلة) ══
if st.session_state["nav_page"] == "tab_china":
    with tab_china:
        if _tab_gate("tab_china", "🚢 شحنات الصين | China Shipments"):
            render_china_tab()

if st.session_state["nav_page"] == "tab_contracts":
    with tab_contracts:
        if _tab_gate("tab_contracts", "📑 العقود | Contracts"):
            render_contracts_tab()

if st.session_state["nav_page"] == "tab_upload":
    with tab_upload:
        if _tab_gate("tab_upload", "📤 رفع ملف | Upload Files"):
            render_upload_tab()

if st.session_state["nav_page"] == "tab_ads":
    with tab_ads:
        if _tab_gate("tab_ads", "📢 الإعلانات | Ads"):
            st.header("📢 الإعلانات | Ads")
            st.caption("منقولة برة داشبورد المبيعات وبتتعرض مرة واحدة بس — بيانات الإعلانات (Views/Clicks/Orders/Spends/Revenue) جايه من شيت الإعلانات نفسه، واللي مفيهوش تصنيف FBN/FBB أصلاً | Moved out of the sales dashboard and shown once — ad data (Views/Clicks/Orders/Spends/Revenue) comes from the Advertisements sheet, which has no FBN/FBB split to begin with")
            _noon_ads_subtab, _amz_ads_subtab, _ty_ads_subtab = st.tabs(["🅽 إعلانات نون | Noon Ads", "📦 إعلانات أمازون | Amazon Ads", "🟣 إعلانات تريندول | Trendyol Ads"])
            with _noon_ads_subtab:
                if not inv_map:
                    st.info("ارفع ملف المخزون أولاً من تاب «رفع ملف» | Upload Inventory first")
                else:
                    _render_ads_performance_tab()
            with _amz_ads_subtab:
                st.caption("بيانات إعلانات أمازون جايه من شيت AdvertisementsAmz، ومفتاحها ASIN بدل SKU — نفس منطق وتصميم تاب إعلانات نون بالظبط | Amazon ad data comes from the AdvertisementsAmz sheet, keyed by ASIN instead of SKU — identical logic and layout to the Noon Ads tab")
                _render_ads_performance_tab_amazon()
            with _ty_ads_subtab:
                _render_ads_performance_tab_trendyol()
