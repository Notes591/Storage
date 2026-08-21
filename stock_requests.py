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

# ══ قفل القايمة الجانبية تلقائياً على الموبايل لما تدوس على أي حاجة برا | Auto-close
#    the sidebar on mobile when tapping outside it, instead of only via the toggle
#    button ══
# ملحوظة: القايمة عندنا بتاخد جزء من عرض الشاشة بس وهي مفتوحة على الموبايل (مش
# تغطية كاملة)، يعني باقي الشاشة (زي فورم كلمة السر) بيفضل ظاهر وقابل للمس جنبها.
# النسخة القديمة كانت بتسمع لأي ضغطة على الصفحة وتقفل القايمة برمجياً، لكن من غير
# ما توقف الضغطة الأصلية من إنها توصل للعنصر اللي تحتها — فكان بيحصل: القايمة
# تتقفل وفي نفس اللحظة الصفحة تعيد ترتيب نفسها، فإصبعك يبقى فوق عنصر تاني
# (زي زرار "دخول")، وده اللي كان بيظهر كـ"نطة" وتغطية لفورم تسجيل الدخول.
# الحل: طبقة شفافة (overlay) تغطي الشاشة كاملة وقت ما القايمة مفتوحة، وهي اللي
# بتاخد الضغطة الأولى بالكامل ومتسيبهاش توصل لأي حاجة تحتها — تقفل القايمة بس،
# وأي ضغطة بعد كده بتوصل طبيعي وسليم لللي تحتها | Note: our sidebar only takes
# part of the screen width when open on mobile (not a full overlay), so the rest
# of the screen (like the password form) stays visible and tappable right next to
# it. The old version listened for any tap on the page and closed the sidebar
# programmatically, but never stopped that same tap from reaching whatever was
# underneath it — so the sidebar would close, the page would reflow at that exact
# moment, and your finger would end up over a different element (like the
# "Unlock" button), which showed up as a "jump" covering the login form. Fix: a
# transparent overlay that covers the full screen while the sidebar is open. It
# fully absorbs the first tap so nothing underneath ever receives it — that tap
# only closes the sidebar, and every tap after that lands normally and safely on
# whatever is underneath.
st.components.v1.html("""
<script>
(function() {
    const doc = window.parent.document;
    if (doc.__sidebarOverlayBound) return;
    doc.__sidebarOverlayBound = true;

    function isMobile() { return window.parent.innerWidth < 768; }

    function getSidebar() { return doc.querySelector('[data-testid="stSidebar"]'); }

    function isSidebarOpen(sidebar) {
        if (!sidebar) return false;
        const expanded = sidebar.getAttribute('aria-expanded');
        return (expanded === null) ? sidebar.offsetWidth > 0 : expanded === 'true';
    }

    function getOpenToggle() {
        return doc.querySelector('[data-testid="stSidebarCollapseButton"] button')
            || doc.querySelector('[data-testid="stSidebarCollapseButton"]')
            || doc.querySelector('[data-testid="baseButton-headerNoPadding"]');
    }

    function ensureOverlay(sidebar) {
        let overlay = doc.getElementById('__sidebarCloseOverlay');
        if (overlay) return overlay;
        overlay = doc.createElement('div');
        overlay.id = '__sidebarCloseOverlay';
        overlay.style.position = 'fixed';
        overlay.style.top = '0';
        overlay.style.left = '0';
        overlay.style.right = '0';
        overlay.style.bottom = '0';
        overlay.style.background = 'transparent';
        // نحطها فوق باقي المحتوى بس تحت القايمة الجانبية نفسها، عشان تمنع أي لمسة
        // من الوصول لأي عنصر تحتها من غير ما تغطي القايمة ذاتها | Sit above the
        // rest of the content but below the sidebar itself, so it blocks any tap
        // from reaching anything underneath without covering the sidebar
        const sidebarZ = parseInt(window.parent.getComputedStyle(sidebar).zIndex, 10) || 999990;
        overlay.style.zIndex = String(sidebarZ - 1);

        function closeOnTap(e) {
            e.preventDefault();
            e.stopPropagation();
            const toggle = getOpenToggle();
            if (toggle) toggle.click();
        }
        overlay.addEventListener('touchstart', closeOnTap, { passive: false, capture: true });
        overlay.addEventListener('click', closeOnTap, true);
        doc.body.appendChild(overlay);
        return overlay;
    }

    function removeOverlay() {
        const overlay = doc.getElementById('__sidebarCloseOverlay');
        if (overlay) overlay.remove();
    }

    function sync() {
        const sidebar = getSidebar();
        if (!sidebar || !isMobile()) { removeOverlay(); return; }
        if (isSidebarOpen(sidebar)) {
            ensureOverlay(sidebar);
        } else {
            removeOverlay();
        }
    }

    const observer = new MutationObserver(sync);
    function attachObserver() {
        const sidebar = getSidebar();
        if (sidebar) {
            observer.observe(sidebar, { attributes: true, attributeFilter: ['aria-expanded', 'style', 'class'] });
        }
    }
    attachObserver();
    // القايمة الجانبية ممكن تتحمّل متأخرة شوية عن السكريبت ده، فبنعيد المحاولة
    // لحد ما تظهر | The sidebar may render slightly after this script runs, so
    // retry until it shows up
    let tries = 0;
    const retryId = setInterval(function() {
        tries++;
        if (getSidebar() || tries > 20) { clearInterval(retryId); attachObserver(); sync(); }
    }, 250);

    window.parent.addEventListener('resize', sync);
    sync();
})();
</script>
""", height=0)

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

def build_daily_orders_counts_fbb(dates):
    """نفس منطق build_daily_orders_counts بالظبط، لكن بيقتصر بس على صفوف FBB
    (Fulfillment Model = Fulfilled by Partner / FBP فقط) | Same logic as
    build_daily_orders_counts exactly, restricted to FBB rows
    (Fulfillment Model = FBP) only."""
    data = get_cached(daily_orders_sheet)
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

def build_daily_orders_prices_fbb(dates):
    """نفس منطق build_daily_orders_prices بالظبط، لكن بيقتصر بس على صفوف FBB
    (Fulfillment Model = Fulfilled by Partner / FBP فقط) | Same logic as
    build_daily_orders_prices exactly, restricted to FBB rows
    (Fulfillment Model = FBP) only."""
    data = get_cached(daily_orders_sheet)
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

def build_daily_orders_family_stats_fbb(dates, live_map=None):
    """نفس منطق build_daily_orders_family_stats بالظبط، لكن بيقتصر بس على صفوف FBB
    (Fulfillment Model = Fulfilled by Partner / FBP فقط) —
    نسخة منفصلة تمامًا زي build_daily_orders_counts_fbb/build_daily_orders_prices_fbb
    فوق، من غير أي تعديل على الدالة الأصلية | Same logic as
    build_daily_orders_family_stats exactly, restricted to FBB rows only. A fully
    separate copy — mirrors build_daily_orders_counts_fbb / build_daily_orders_prices_fbb
    above — the original function is left completely untouched."""
    data = get_cached(daily_orders_sheet)
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
def build_daily_orders_counts_fbn(dates):
    """نفس منطق build_daily_orders_counts بالظبط، لكن بيستبعد صفوف FBB (عكس
    build_daily_orders_counts_fbb) | Same logic as build_daily_orders_counts
    exactly, but excludes FBB rows (inverse of build_daily_orders_counts_fbb)."""
    data = get_cached(daily_orders_sheet)
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

def build_daily_orders_prices_fbn(dates):
    """نفس منطق build_daily_orders_prices بالظبط، لكن بيستبعد صفوف FBB (عكس
    build_daily_orders_prices_fbb) | Same logic as build_daily_orders_prices
    exactly, but excludes FBB rows (inverse of build_daily_orders_prices_fbb)."""
    data = get_cached(daily_orders_sheet)
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

def build_daily_orders_family_stats_fbn(dates, live_map=None):
    """نفس منطق build_daily_orders_family_stats بالظبط، لكن بيستبعد صفوف FBB
    (عكس build_daily_orders_family_stats_fbb) | Same logic as
    build_daily_orders_family_stats exactly, but excludes FBB rows (inverse of
    build_daily_orders_family_stats_fbb)."""
    data = get_cached(daily_orders_sheet)
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
    —  مخزونها انتهى بالكامل وخرجت من ملف المخزون. تظهر بنفس تفاصيل تابي المراجعة."""
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
st.title("🏢 عالم الرشاقة للتجارة | Fitness World Trading")

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

NAV_ITEMS = [
    ("tab14",    "🛒", "المبيعات | Sales"),
    ("tab_dash", "📊", "داشبورد المبيعات | Dashboard"),
    ("tab9",     "📦", "المخزون | Inventory"),
    ("tab16",    "🗂️", "مخزون بدون بيع | No Sales"),
    ("tab_ads",  "📢", "الإعلانات | Ads"),
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

tab14    = st.container(key="navpanel_tab14")
tab_dash = st.container(key="navpanel_tab_dash")
tab9     = st.container(key="navpanel_tab9")
tab16    = st.container(key="navpanel_tab16")
tab_ads  = st.container(key="navpanel_tab_ads")

for _key, _i, _l in NAV_ITEMS:
    _disp = "block" if st.session_state["nav_page"] == _key else "none"
    st.markdown(f'<style>.st-key-navpanel_{_key} {{ display: {_disp} !important; }}</style>', unsafe_allow_html=True)

with tab9:
    if _tab_gate("tab9", "📊 المخزون | Inventory"):
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
    # ══ TAB 14 — المبيعات ══
with tab14:
    if _tab_gate("tab14", "🛒 المبيعات | Sales"):
        _fbn_subtab_t14, _fbb_subtab_t14 = st.tabs(["🅽 مبيعات نون FBN | Noon FBN Sales", "🅱 مبيعات نون FBB | Noon FBB Sales"])
        with _fbn_subtab_t14:
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
                multi_counts_t14 = build_daily_orders_counts_fbn(sales_dates)
                prices_map_t14   = build_daily_orders_prices_fbn(sales_dates)

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

                # ══ إجماليات اليومية في الأعلى — بتتحسب من كل صفوف الأوردرز الخام
                #    (multi_counts_t14)، مش بس الـ SKUs الموجودة في ملف المخزون
                #    المرفوع، عشان الإجمالي يعكس العدد الحقيقي دايمًا | Daily totals
                #    are computed from the raw daily-orders counts, not only SKUs
                #    present in the uploaded inventory file, so the total always
                #    reflects the true order count ══
                totals_per_day = {d: sum(day_counts.get(d, 0) for day_counts in multi_counts_t14.values()) for d in sales_dates}
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
                live_map_t14 = get_live_map()
                xdock_threshold_t14 = int(load_settings().get("xdock_low_stock_threshold","10") or 10)

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

                        # ══ صافي سعر البيع بعد العمولة والتوصيل والضريبة | Net price after commission,
                        #    delivery fees, and VAT ══
                        # سعر البيع الأساسي بييجي من عمود sale_price في تاب LIVE أولاً (وده اللي بتُحسب عليه
                        # الخصومات والعمولة والضريبة والإعلانات)، ولو مش موجود بيرجع لسعر الطلبات القديم
                        # (بقى اسمه "سعر العرض") كبديل.
                        com_info_t14 = com_map_t14.get(r["sku_up"])
                        if com_info_t14:
                            offer_price_t14 = get_latest_sku_price(r, sales_dates)
                            latest_price_t14, price_from_live_t14 = get_base_price(r["sku_up"], live_map_t14, offer_price_t14)
                            if latest_price_t14 is not None:
                                net_fees_t14, net_tax_t14 = compute_net_price_after_fees(latest_price_t14, com_info_t14)

                                # ── فلوس الإعلان دي مفديه ولا لأ؟ | Did the ad spend pay off overall? ──
                                # مقارنة إجمالية: صافي الربح الكلي من الطلبات اللي جابها الإعلان مقابل
                                # إجمالي اللي اتصرف على الإعلان — مش مقارنة لكل طلب لوحده
                                ad_insight_t14 = ""
                                if ads_entries_t14:
                                    if total_orders_t14 <= 0:
                                        ad_insight_t14 = (
                                            '<div dir="rtl" style="margin-top:8px;padding:7px 11px;background:#4c051655;'
                                            'border:1px solid #dc2626;border-radius:7px;">'
                                            f'<span style="color:#f87171;font-size:13px;font-weight:800;">🚨 مفدتش لحد دلوقتي: '
                                            f'اتصرف {total_spends_t14:,.2f} ريال على الإعلان ده ولسه ما جابش أي طلبات فعلية — يستاهل تراجع التفاصيل فوق 👆</span>'
                                            '</div>')
                                    else:
                                        total_net_from_ads_t14 = total_orders_t14 * net_tax_t14
                                        net_result_t14 = total_net_from_ads_t14 - total_spends_t14
                                        if net_result_t14 >= 0:
                                            ad_insight_t14 = (
                                                '<div dir="rtl" style="margin-top:8px;padding:7px 11px;background:#052e1655;'
                                                'border:1px solid #16a34a;border-radius:7px;">'
                                                f'<span style="color:#4ade80;font-size:13px;font-weight:800;">🎯 الاعلان مربح: '
                                                f'عدد طلبات الاعلان {total_orders_t14:,.0f} طلب بصافي ربح إجمالي {total_net_from_ads_t14:,.2f} ريال مقابل '
                                                f'{total_spends_t14:,.2f} ريال مدفوع — حقق <u>{net_result_t14:,.2f} ريال</u>  👌</span>'
                                                '</div>')
                                        else:
                                            ad_insight_t14 = (
                                                '<div dir="rtl" style="margin-top:8px;padding:7px 11px;background:#4c051655;'
                                                'border:1px solid #dc2626;border-radius:7px;">'
                                                f'<span style="color:#f87171;font-size:13px;font-weight:800;">🚨 الاعلان غير مربح: '
                                                f'مدفوع {total_spends_t14:,.2f} ريال، لكن صافي الربح من {total_orders_t14:,.0f} طلب بس {total_net_from_ads_t14:,.2f} ريال — '
                                                f'خسران <u>{abs(net_result_t14):,.2f} ريال إجمالي</u>. افتح تفاصيل الحملة فوق 👆 وتراجعها</span>'
                                                '</div>')

                                offer_price_line_t14 = ""
                                if offer_price_t14 is not None and (not price_from_live_t14 or round(offer_price_t14, 2) != round(latest_price_t14, 2)):
                                    offer_price_line_t14 = f' &nbsp;|&nbsp; 🏷️ سعر العرض (معلومة فقط، غير مستخدم في الحسابات): <b>{offer_price_t14:,.2f}</b> ريال'
                                price_label_t14 = "سعر البيع الأساسي" if price_from_live_t14 else "سعر البيع (سعر العرض)"
                                st.markdown(
                                    f'<div style="background:#1e293b;border:1px solid #334155;border-radius:8px;'
                                    f'padding:8px 14px;margin:4px 0;">'
                                    f'<span style="color:#e2e8f0;font-size:13px;">💵 {price_label_t14}: <b>{latest_price_t14:,.2f}</b> ريال'
                                    f'{offer_price_line_t14}'
                                    f'&nbsp;|&nbsp; 🚚 توصيل: <b>{com_info_t14["delivery"]:,.0f}</b> '
                                    f'&nbsp;|&nbsp; 🏷️ عمولة: <b>{com_info_t14["commission_pct"]:,.0f}%</b></span><br>'
                                    f'<span style="color:#4ade80;font-size:14px;font-weight:bold;">💳 الصافي بعد خصم العمولة والتوصيل: {net_fees_t14:,.2f} ريال</span><br>'
                                    f'<span style="color:#fbbf24;font-size:14px;font-weight:bold;">🧾 الصافي بعد خصم 15% ضريبة: {net_tax_t14:,.2f} ريال</span>'
                                    f'{ad_insight_t14}'
                                    f'</div>',
                                    unsafe_allow_html=True)
                            else:
                                st.caption("ℹ️ فيه عمولة وتوصيل مسجلين لكن مفيش سعر بيع حديث (لا في LIVE ولا في الطلبات) لحساب الصافي منهم | Commission & delivery are set but no price found (neither in LIVE nor in orders) to calculate the net")

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

        with _fbb_subtab_t14:
            st.subheader("🛒 مبيعات نون FBB | Noon FBB Sales")
            st.caption("نفس تاب المبيعات بالظبط، لكن مقتصر على الطلبات اللي Fulfillment Model بتاعها Fulfilled by Partner (FBP) في شيت DailyOrders (بغض النظر عن الـ Status) | Same as the Sales tab exactly, restricted to orders whose Fulfillment Model is Fulfilled by Partner (FBP) in the DailyOrders sheet (regardless of Status)")
            render_tacweed_upload("sales_fbb")
            render_warehouse_stock_upload("sales_fbb")

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
                multi_counts_t14 = build_daily_orders_counts_fbb(sales_dates)
                prices_map_t14   = build_daily_orders_prices_fbb(sales_dates)

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

                # ── SKUs باعت طلبات FBB لكن مش موجودة في ملف المخزون المرفوع —
                #    بتتضاف كمان عشان كل مبيعات FBB تظهر في القائمة، مش بس اللي
                #    عندها مخزون مرفوع. المخزون بيتحط 0 (غير معروف) ومعاها علامة
                #    توضيحية | SKUs with FBB orders that aren't in the uploaded
                #    inventory file — added too so every FBB sale shows up, not
                #    only SKUs with uploaded stock. Stock shown as 0 (unknown)
                #    with a note explaining why.
                inv_skus_seen_fbb = set(inv_map.keys())
                for sku_up, day_counts in multi_counts_t14.items():
                    if sku_up in inv_skus_seen_fbb:
                        continue
                    total_recent_orphan = sum(day_counts.get(d, 0) for d in sales_dates)
                    if total_recent_orphan <= 0:
                        continue
                    day_prices_orphan = prices_map_t14.get(sku_up, {d: [] for d in sales_dates})
                    sales_tab_rows.append({
                        "sku": sku_up, "sku_up": sku_up,
                        "stock": 0, "sales_month": 0, "img": "",
                        "day_counts": day_counts, "day_prices": day_prices_orphan,
                        "total_recent": total_recent_orphan,
                        "effective_avg": 0,
                        "days_to_stockout": 9999,
                        "not_in_inventory": True,
                    })

                # ترتيب: الأكتر مبيعاً أمس أولاً
                sales_tab_rows.sort(key=lambda r: -r["day_counts"].get(sales_dates[0], 0) if sales_dates else 0)

                # ══ إجماليات اليومية في الأعلى — بتتحسب من كل صفوف الأوردرز الخام
                #    (multi_counts_t14)، مش بس الـ SKUs الموجودة في ملف المخزون
                #    المرفوع، عشان الإجمالي يعكس العدد الحقيقي دايمًا | Daily totals
                #    are computed from the raw daily-orders counts, not only SKUs
                #    present in the uploaded inventory file, so the total always
                #    reflects the true order count ══
                totals_per_day = {d: sum(day_counts.get(d, 0) for day_counts in multi_counts_t14.values()) for d in sales_dates}
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

                srch_t14 = st.text_input("🔍 بحث SKU | Search SKU", key="srch_t14_fbb", placeholder="اكتب SKU...")
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
                    with c1: dl_btn(df_t14, "sales_daily_fbb", key="dlbtn_t14_fbb")
                    with c2: st.info(f"📦 SKUs: {len(sales_tab_rows)} | 📅 {sales_display_days} يوم")

                # ══ قائمة SKUs المرحلة من المبيعات (محتاج جدولة فقط) ══
                # تحديث المرحلين بعد بناء الصفوف الكاملة
                _new_transferred = []

                # ══ خريطة SKUs المجدولة خلال آخر 4 أيام (لعرض ASN + الكمية لو فعلاً ليها جدولة) ══
                recent_sched_map_t14 = get_recent_schedule_rows(days_back=4)
                pending_approval_skus_t14 = get_pending_approval_skus()
                ads_map_t14 = get_ads_map()
                com_map_t14 = get_com_map()
                live_map_t14 = get_live_map()
                xdock_threshold_t14 = int(load_settings().get("xdock_low_stock_threshold","10") or 10)

                st.divider()
                for r in sales_tab_rows:
                    c_img, c_info = st.columns([1, 7])
                    with c_img:
                        show_img(r["img"], 70)
                    with c_info:
                        st.markdown(f"**SKU:** {sku_link_html(r['sku'])}", unsafe_allow_html=True)
                        if r.get("not_in_inventory"):
                            st.markdown(
                                '<span style="background:#78350f;color:#fde68a;padding:2px 8px;'
                                'border-radius:6px;font-size:12px;">⚠️ مش موجود في ملف المخزون المرفوع | Not in uploaded inventory</span>',
                                unsafe_allow_html=True)
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

                        # ══ صافي سعر البيع بعد العمولة والتوصيل والضريبة | Net price after commission,
                        #    delivery fees, and VAT ══
                        # سعر البيع الأساسي بييجي من عمود sale_price في تاب LIVE أولاً (وده اللي بتُحسب عليه
                        # الخصومات والعمولة والضريبة والإعلانات)، ولو مش موجود بيرجع لسعر الطلبات القديم
                        # (بقى اسمه "سعر العرض") كبديل.
                        com_info_t14 = com_map_t14.get(r["sku_up"])
                        if com_info_t14:
                            offer_price_t14 = get_latest_sku_price(r, sales_dates)
                            latest_price_t14, price_from_live_t14 = get_base_price(r["sku_up"], live_map_t14, offer_price_t14)
                            if latest_price_t14 is not None:
                                net_fees_t14, net_tax_t14 = compute_net_price_after_fees(latest_price_t14, com_info_t14)

                                # ── فلوس الإعلان دي مفديه ولا لأ؟ | Did the ad spend pay off overall? ──
                                # مقارنة إجمالية: صافي الربح الكلي من الطلبات اللي جابها الإعلان مقابل
                                # إجمالي اللي اتصرف على الإعلان — مش مقارنة لكل طلب لوحده
                                ad_insight_t14 = ""
                                if ads_entries_t14:
                                    if total_orders_t14 <= 0:
                                        ad_insight_t14 = (
                                            '<div dir="rtl" style="margin-top:8px;padding:7px 11px;background:#4c051655;'
                                            'border:1px solid #dc2626;border-radius:7px;">'
                                            f'<span style="color:#f87171;font-size:13px;font-weight:800;">🚨 مفدتش لحد دلوقتي: '
                                            f'اتصرف {total_spends_t14:,.2f} ريال على الإعلان ده ولسه ما جابش أي طلبات فعلية — يستاهل تراجع التفاصيل فوق 👆</span>'
                                            '</div>')
                                    else:
                                        total_net_from_ads_t14 = total_orders_t14 * net_tax_t14
                                        net_result_t14 = total_net_from_ads_t14 - total_spends_t14
                                        if net_result_t14 >= 0:
                                            ad_insight_t14 = (
                                                '<div dir="rtl" style="margin-top:8px;padding:7px 11px;background:#052e1655;'
                                                'border:1px solid #16a34a;border-radius:7px;">'
                                                f'<span style="color:#4ade80;font-size:13px;font-weight:800;">🎯 الاعلان مربح: '
                                                f'عدد طلبات الاعلان {total_orders_t14:,.0f} طلب بصافي ربح إجمالي {total_net_from_ads_t14:,.2f} ريال مقابل '
                                                f'{total_spends_t14:,.2f} ريال مدفوع — حقق <u>{net_result_t14:,.2f} ريال</u>  👌</span>'
                                                '</div>')
                                        else:
                                            ad_insight_t14 = (
                                                '<div dir="rtl" style="margin-top:8px;padding:7px 11px;background:#4c051655;'
                                                'border:1px solid #dc2626;border-radius:7px;">'
                                                f'<span style="color:#f87171;font-size:13px;font-weight:800;">🚨 الاعلان غير مربح: '
                                                f'مدفوع {total_spends_t14:,.2f} ريال، لكن صافي الربح من {total_orders_t14:,.0f} طلب بس {total_net_from_ads_t14:,.2f} ريال — '
                                                f'خسران <u>{abs(net_result_t14):,.2f} ريال إجمالي</u>. افتح تفاصيل الحملة فوق 👆 وتراجعها</span>'
                                                '</div>')

                                offer_price_line_t14 = ""
                                if offer_price_t14 is not None and (not price_from_live_t14 or round(offer_price_t14, 2) != round(latest_price_t14, 2)):
                                    offer_price_line_t14 = f' &nbsp;|&nbsp; 🏷️ سعر العرض (معلومة فقط، غير مستخدم في الحسابات): <b>{offer_price_t14:,.2f}</b> ريال'
                                price_label_t14 = "سعر البيع الأساسي" if price_from_live_t14 else "سعر البيع (سعر العرض)"
                                st.markdown(
                                    f'<div style="background:#1e293b;border:1px solid #334155;border-radius:8px;'
                                    f'padding:8px 14px;margin:4px 0;">'
                                    f'<span style="color:#e2e8f0;font-size:13px;">💵 {price_label_t14}: <b>{latest_price_t14:,.2f}</b> ريال'
                                    f'{offer_price_line_t14}'
                                    f'&nbsp;|&nbsp; 🚚 توصيل: <b>{com_info_t14["delivery"]:,.0f}</b> '
                                    f'&nbsp;|&nbsp; 🏷️ عمولة: <b>{com_info_t14["commission_pct"]:,.0f}%</b></span><br>'
                                    f'<span style="color:#4ade80;font-size:14px;font-weight:bold;">💳 الصافي بعد خصم العمولة والتوصيل: {net_fees_t14:,.2f} ريال</span><br>'
                                    f'<span style="color:#fbbf24;font-size:14px;font-weight:bold;">🧾 الصافي بعد خصم 15% ضريبة: {net_tax_t14:,.2f} ريال</span>'
                                    f'{ad_insight_t14}'
                                    f'</div>',
                                    unsafe_allow_html=True)
                            else:
                                st.caption("ℹ️ فيه عمولة وتوصيل مسجلين لكن مفيش سعر بيع حديث (لا في LIVE ولا في الطلبات) لحساب الصافي منهم | Commission & delivery are set but no price found (neither in LIVE nor in orders) to calculate the net")

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
    # ══ TAB DASHBOARD — داشبورد تحليلات المبيعات (تاب جديدة منفصلة) ══
def _render_sales_dashboard_body(counts_fn, prices_fn, family_stats_fn, key_suffix):
    """يعرض كل محتوى داشبورد المبيعات (KPIs + تنبيهات سريعة + تحليل إعلانات + رسوم
    بيانية + أعلى/أبطأ الأصناف + تحليل SKU فردي) — بيتفعّل مرتين: مرة بأرقام FBN
    (counts_fn/prices_fn/family_stats_fn = النسخة الأصلية غير المفلترة) ومرة بأرقام
    FBB (النسخة المفلترة _fbb) — نفس الكود بالظبط في الحالتين، غير بس مصدر البيانات
    و key_suffix (عشان مفاتيح عناصر Streamlit متتكررش) | Renders the entire sales
    dashboard content — called twice: once with the original (FBN) data-builder
    functions, once with the _fbb-filtered ones. Identical code both times; only the
    data source and the widget key_suffix differ."""
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
    live_map_dash  = get_live_map()

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
    for sku_up_td, info_td in inv_map.items():
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
    #    اللي عندها مخزون مرفوع | SKUs that sold but aren't in the uploaded
    #    inventory file — added too so the dashboard totals (All/FBN/FBB)
    #    reflect the true order count, not only SKUs with uploaded stock.
    _inv_skus_seen_td = set(inv_map.keys())
    _orphan_skus_td = (set(cur_counts_td.keys()) | set(prev_counts_td.keys())) - _inv_skus_seen_td
    for sku_up_td in _orphan_skus_td:
        cur_t_td  = _td_total(cur_counts_td, sku_up_td, cur_dates_td)
        prev_t_td = _td_total(prev_counts_td, sku_up_td, prev_dates_td)
        if cur_t_td <= 0 and prev_t_td <= 0:
            continue
        cur_rev_td  = _td_revenue_total(cur_prices_td, sku_up_td, cur_dates_td, live_map_dash)
        prev_rev_td = _td_revenue_total(prev_prices_td, sku_up_td, prev_dates_td, live_map_dash)
        rows_td.append({
            "sku_up": sku_up_td, "sku": sku_up_td, "img": "",
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
    xdock_threshold_dash = int(load_settings().get("xdock_low_stock_threshold", "10") or 10)
    xdock_low_rows_td = []
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
            st.markdown(f"{sku_link_html(r['sku'])}", unsafe_allow_html=True)
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
            for r in attention_rows_td:
                _render_alert_sku_row(
                    r,
                    lines=[f"📦 مخزون: {r['stock']:,} — ⏳ نفاد خلال {r['days_to_so']} يوم", r["status"]],
                    badges_html=_extra_context_badges(r, include_schedule=True))
        else:
            st.caption("لا توجد أصناف معرضة لنفاد المخزون حالياً")

    with st.expander(f"🟡 عرض منتجات بدون مبيعات في الفترة ({zero_skus_td:,}) | Show no-sale SKUs"):
        st.caption("ℹ️ الأصناف اللي مخزونها صفر مستبعدة من القائمة دي — طبيعي متبعش لو مفيش مخزون أصلاً | SKUs with zero stock are excluded — no stock naturally means no sales")
        if zero_rows_td:
            df_al2 = pd.DataFrame([{
                "SKU": r["sku"], "المخزون | Stock": r["stock"], "مبيعات الفترة | Period Sales": r["cur"],
            } for r in zero_rows_td])
            dl_btn(df_al2, "alert_no_sales", key=f"dl_alert_nosale_td_{key_suffix}")
            for r in zero_rows_td:
                _render_alert_sku_row(
                    r,
                    lines=[f"📦 مخزون: {r['stock']:,} — لا يوجد مبيعات خلال {analysis_days_td} يوم رغم توفر المخزون"],
                    badges_html=_extra_context_badges(r, include_schedule=False))
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
            for r in decline_rows_td:
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
        else:
            st.caption("لا توجد أصناف انخفضت مبيعاتها بنسبة 20%+ حالياً")

    with st.expander(f"🟢 عرض منتجات ارتفعت مبيعاتها ({len(rise_rows_td):,}) | Show rising SKUs"):
        if rise_rows_td:
            df_al4 = pd.DataFrame([{
                "SKU": r["sku"], "الفترة الحالية | Current": r["cur"], "الفترة السابقة | Previous": r["prev"],
                "الارتفاع | Rise %": round((r["cur"] - r["prev"]) / r["prev"] * 100, 1),
            } for r in rise_rows_td])
            dl_btn(df_al4, "alert_rising", key=f"dl_alert_rise_td_{key_suffix}")
            for r in rise_rows_td:
                rise_pct = (r["cur"] - r["prev"]) / r["prev"] * 100
                _render_alert_sku_row(
                    r,
                    lines=[f"📈 {r['cur']:,} مقابل {r['prev']:,} (+{rise_pct:.0f}%)", f"📦 مخزون حالي: {r['stock']:,}"],
                    badges_html=_extra_context_badges(r, include_schedule=False))
        else:
            st.caption("لا توجد أصناف ارتفعت مبيعاتها بنسبة 20%+ حالياً")

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
            for r in xdock_low_rows_td:
                if r["has_other_stock"]:
                    other_badge = f'<span style="color:#4ade80;font-size:12px;">📦 عنده مخزون FBN (Inventory): {r["other_stock"]:,} — أقل إلحاحاً</span>'
                else:
                    other_badge = '<span style="color:#f87171;font-size:12px;font-weight:700;">🚫 لا يوجد مخزون FBN — الاعتماد على Xdock بس</span>'
                _render_alert_sku_row(
                    r,
                    lines=[f"🟣 مخزون Xdock: {r['stock_xdock_net']:,} — 📈 مبيع شهري: {r['sales_month']:,}"
                           + (f" — 💵 {r['price']:,.2f} ريال" if r["price"] is not None else "")],
                    badges_html=other_badge)
        else:
            st.caption(f"لا توجد أصناف مخزون Xdock عندها {xdock_threshold_dash} قطعة أو أقل حالياً")

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

    # ── ربحية الإعلانات لكل SKU (نفس منطق تاب المبيعات: صافي الربح الكلي من طلبات الإعلان
    #    مقابل إجمالي المصروف عليه) — عشان تظهر في التنبيهات السريعة تحت ──
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
    for sku_up_ad, ads_entries_ad in ads_map_dash.items():
        com_info_ad = com_map_dash.get(sku_up_ad)
        if not ads_entries_ad or not com_info_ad:
            continue
        latest_price_ad = _ads_latest_price_for_sku(sku_up_ad)
        if latest_price_ad is None:
            continue
        _, net_tax_ad = compute_net_price_after_fees(latest_price_ad, com_info_ad)
        total_spends_ad = sum(a["spends"] for a in ads_entries_ad)
        total_orders_ad = sum(a["orders"] for a in ads_entries_ad)
        total_net_ad = total_orders_ad * net_tax_ad
        result_ad = total_net_ad - total_spends_ad
        inv_info_ad = inv_map.get(sku_up_ad, {})
        entry_ad = {"sku_up": sku_up_ad, "sku": inv_info_ad.get("sku", sku_up_ad),
                    "img": inv_info_ad.get("img", ""),
                    "spends": total_spends_ad, "orders": total_orders_ad,
                    "net_total": total_net_ad, "result": result_ad}
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
        if r["orders"] <= 0:
            return (f'<span style="color:#f87171;font-size:12px;font-weight:700;">🚨 مفدتش لحد دلوقتي: '
                    f'اتصرف {r["spends"]:,.2f} ريال ولسه ما جابش أي طلبات فعلية</span>')
        if r["result"] >= 0:
            return (f'<span style="color:#4ade80;font-size:12px;font-weight:700;">🎯 الاعلان مربح: '
                    f'عدد طلبات الاعلان {r["orders"]:,.0f} طلب بصافي ربح إجمالي {r["net_total"]:,.2f} ريال مقابل '
                    f'{r["spends"]:,.2f} ريال مدفوع — حقق {r["result"]:,.2f} ريال 👌</span>')
        return (f'<span style="color:#f87171;font-size:12px;font-weight:700;">🚨 الاعلان غير مربح: '
                f'مدفوع {r["spends"]:,.2f} ريال، لكن صافي الربح من {r["orders"]:,.0f} طلب بس {r["net_total"]:,.2f} ريال — '
                f'خسران {abs(r["result"]):,.2f} ريال إجمالي</span>')

    with st.expander(f"🎯 عرض منتجات ربحانة من الإعلانات ({len(ads_profit_rows_ap):,}) | Show profitable-ads SKUs"):
        if ads_profit_rows_ap:
            df_al5 = pd.DataFrame([{
                "SKU": r["sku"], "طلبات الإعلان | Ad Orders": r["orders"],
                "المصروف | Spends": round(r["spends"], 2), "صافي الربح | Net Total": round(r["net_total"], 2),
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
                "المصروف | Spends": round(r["spends"], 2), "صافي الربح | Net Total": round(r["net_total"], 2),
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
        for sku_up_r in shown_r:
            inv_info_r = inv_map.get(sku_up_r, {})
            ci_r, cinfo_r = st.columns([1, 6])
            with ci_r:
                show_img(inv_info_r.get("img", ""), 45)
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

with tab_dash:
    if _tab_gate("tab_dash", "📊 داشبورد المبيعات | Sales Dashboard"):
        st.header("📊 داشبورد المبيعات | Sales Dashboard")
        st.caption("تحليلات موسّعة على بيانات المبيعات مع صور المنتجات — منفصلة تمامًا عن تاب المبيعات الأصلي ولا تؤثر عليه | Extended sales analytics with product images — fully separate from, and does not affect, the original Sales tab")

        if not inv_map:
            st.info("ارفع ملف المخزون أولاً من تاب المخزون | Upload Inventory first")
        else:
            _dash_all_subtab_td, _dash_fbn_subtab_td, _dash_fbb_subtab_td = st.tabs(
                ["🔷 الكل (FBN+FBB) | All Combined",
                 "🅽 مبيعات نون FBN | Noon FBN Sales",
                 "🅱 مبيعات نون FBB | Noon FBB Sales"])
            with _dash_all_subtab_td:
                # ── الكل مع بعض: كل الطلبات (FBN + FBB) من غير أي فلترة على نوع
                #    التنفيذ — عشان في النهاية هو متجر نون واحد ومبيعات واحدة | All
                #    orders combined (FBN + FBB) with no fulfillment-type filter —
                #    it's one Noon store and one combined sales figure.
                _render_sales_dashboard_body(
                    build_daily_orders_counts, build_daily_orders_prices,
                    build_daily_orders_family_stats, "all")
            with _dash_fbn_subtab_td:
                # ── FBN بس: مستبعد منه أي طلب مصنّف FBB (Fulfillment=FBP)،
                #    عشان FBN + FBB = تاب "الكل" بالظبط من غير أي تكرار |
                #    FBN only: excludes anything classified as FBB
                #    (Fulfillment=FBP), so FBN + FBB always adds up exactly to
                #    "All Combined" with no double-counting.
                _render_sales_dashboard_body(
                    build_daily_orders_counts_fbn, build_daily_orders_prices_fbn,
                    build_daily_orders_family_stats_fbn, "fbn")
            with _dash_fbb_subtab_td:
                _render_sales_dashboard_body(
                    build_daily_orders_counts_fbb, build_daily_orders_prices_fbb,
                    build_daily_orders_family_stats_fbb, "fbb")

    # ══ TAB 15 — تحليل الجدولة ══
    # ══ TAB 16 — مخزون بدون بيع ══
with tab16:
    if _tab_gate("tab16", "📦 مخزون بدون بيع | No Sales"):
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

with tab_ads:
    if _tab_gate("tab_ads", "📢 الإعلانات | Ads"):
        st.header("📢 الإعلانات | Ads")
        st.caption("منقولة برة داشبورد المبيعات وبتتعرض مرة واحدة بس — بيانات الإعلانات (Views/Clicks/Orders/Spends/Revenue) جايه من شيت الإعلانات نفسه، واللي مفيهوش تصنيف FBN/FBB أصلاً | Moved out of the sales dashboard and shown once — ad data (Views/Clicks/Orders/Spends/Revenue) comes from the Advertisements sheet, which has no FBN/FBB split to begin with")
        if not inv_map:
            st.info("ارفع ملف المخزون أولاً من تاب المخزون | Upload Inventory first")
        else:
            _render_ads_performance_tab()
