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
    "Scheduled":         ["ASN","SKU","Quantity","Schedule Date","Image URL","Date Added","Notes","Flag"],
    "CancelledSchedule": ["ASN","SKU","Quantity","Schedule Date","Image URL","Date Added","Cancel Reason","Date Cancelled"],
    "Rescheduled":       ["ASN","SKU","Quantity","Old Schedule Date","Image URL","Date Added","Reschedule Reason","Date Moved"],
    "Expired":           ["ASN","SKU","Quantity","Schedule Date","Image URL","Date Added","Date Expired"],
    "Inventory":         ["SKU","Warehouse","Stock","Monthly Sales","Image URL","Date Uploaded"],
    "DailyOrders":       ["SKU","Order Timestamp","Status","Price","Quantity","Date Uploaded"],
    "Settings":          ["Key","Value"],
    "Check":             ["ASN","SKU","Quantity","Schedule Date","Image URL","Date Added","Notes","Flag"],
    "CancelNotifications": ["ASN","SKUs","Schedule Date","Reason","Timestamp"],
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
cancelled_sheet   = sheets["CancelledSchedule"]
reschedule_sheet  = sheets["Rescheduled"]
expired_sheet     = sheets["Expired"]
inventory_sheet      = sheets["Inventory"]
daily_orders_sheet   = sheets["DailyOrders"]
settings_sheet       = sheets["Settings"]
cancel_notif_sheet   = sheets["CancelNotifications"]

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
    data = get_cached(scheduled_sheet, force=True)
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

def compute_stock_sales_rows(target_date, display_dates=None):
    """يحسب لكل SKU ظهر في أوردرز اليوم المحدد نفس مخرجات استعلامي مراجعة مخزون / مراجعة مبيعات.
    display_dates (اختياري): قائمة تواريخ إضافية تتعرض جنب كل SKU (مثلاً أمس/أول أمس/أول أول أمس)."""
    daily_qty = build_daily_orders_map(target_date)
    display_dates = display_dates or [target_date]
    multi_counts = build_daily_orders_counts(display_dates)
    rows = []
    for sku_up, qty in daily_qty.items():
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
        data = get_cached(sheets[sheet_key])
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
        f'<span style="background:{color};color:white;border-radius:6px;padding:3px 10px;font-size:12px;">'
        f'📅 مجدول بتاريخ {entry["date"]} (ASN {entry["asn"]}) [{entry["source_label"]}] — '
        f'خلال آخر 4 أيام، لسه في فترة الوصول — لا تطلبه تاني | '
        f'Scheduled within the last 4 days — still within arrival window, don\'t re-order</span>'
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
            st.markdown(f"📦 **المخزون | Stock:** {r['stock']} &nbsp;|&nbsp; 📈 **مبيع شهري | Monthly:** {r['sales_month']}")
            st.markdown("🛒 " + render_day_counts_md(r["day_counts"], day_dates, day_labels))
            st.markdown(recent_schedule_badge_html(r["sched"]), unsafe_allow_html=True)
            for note in get_unavailable_ordered_note(r["sku"]):
                st.caption(note)
        st.divider()

def get_latest_schedule_info(sku):
    """يدوّر على SKU في الجدولة والتشييك ويرجع أقرب جدولة (تاريخ) أو None."""
    sku_up = sku.strip().upper()
    candidates = []
    for sheet_key in ("Scheduled","Check"):
        data = get_cached(sheets[sheet_key])
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
        stock_ok     = days_to_so >= cov_days_now if eff_avg > 0 else False
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
    "📊 المخزون | Inventory",
    "🔴 مراجعة المخزون | Stock Review",
    "🗂️ منتهية | Expired",
    "⚙️ الإعدادات | Settings",
    "📈 مراجعة المبيعات | Sales Review",
    "🛒 المبيعات | Sales",
    "🗓️ تحليل الجدولة | Schedule Analysis",
    "📦 مخزون بدون بيع | No Sales",
])
(tab1,tab2,tab3,tab4,tab5,tab_check,tab6,tab7,tab8,tab9,tab10,tab11,tab12,tab13,tab14,tab15,tab16) = tabs

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
                existing = get_cached(scheduled_sheet, force=True)
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
                            to_add.append([asn,sku,qty,ds,img,dn])
                            ex_pairs.add(pair)
                safe_batch_append(scheduled_sheet,to_add)
                msg = f"✅ أُضيف | Added: {len(to_add)}"
                if skipped: msg += f" | ⚠️ مكرر | Duplicates: {skipped}"
                st.success(msg); st.rerun()
        except Exception as e:
            st.error(f"❌ {e}")

    st.divider()
    st.subheader("📋 الجدولة الحالية | Current Schedule")
    data_sch = get_cached(scheduled_sheet)
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
                        sch_d = get_cached(scheduled_sheet, force=True)
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
                        sch_data = get_cached(scheduled_sheet, force=True)
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
                        sch_data = get_cached(scheduled_sheet, force=True)
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
                    to_add = [[r[0],r[1],r[2],r[3],lm.get(r[1].strip().upper(),r[4]),dn,"تم تشييكه | Checked",""] for r in skus_]
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
                        to_add = [[new_asn, sku, qty, new_date, links_map2.get(sku.upper(), img), dn] for sku,qty,img in edited_skus]
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
    data_sc8 = get_cached(scheduled_sheet)
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
    all_review_rows = compute_stock_sales_rows(d1, day_dates)
    stock_review_rows = [r for r in all_review_rows if r["stock_alert"]]

    # إضافة المرحلين من تاب المبيعات (محتاج جدولة فقط)
    transferred_from_sales = st.session_state.get("transferred_skus_t14", [])
    existing_skus_in_review = {r["sku_up"] for r in stock_review_rows}
    for tr in transferred_from_sales:
        if tr["sku_up"] not in existing_skus_in_review:
            avg_tr = tr.get("effective_avg", 0)
            suggested_tr = round(avg_tr * 18) if avg_tr > 0 else 0
            stock_review_rows.append({
                "sku": tr["sku"], "sku_up": tr["sku_up"],
                "stock": tr["stock"], "sales_month": tr["sales_month"],
                "img": tr["img"], "stock_alert": True, "sales_alert": False,
                "suggested_qty": suggested_tr,
                "days_to_stockout": tr.get("days_to_stockout", 0),
                "days_to_stockout_today": tr.get("days_to_stockout", 0),
                "qty": tr["day_counts"].get(d1, 0) if tr.get("day_counts") else 0,
                "day_counts": tr.get("day_counts", {d: 0 for d in day_dates}),
                "_transferred_from_sales": True,
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

        for r in stock_review_rows:
            c_img,c_info = st.columns([1,6])
            with c_img: show_img(r["img"],70)
            with c_info:
                st.markdown(f"**SKU:** `{r['sku']}`")
                if r.get("_transferred_from_sales"):
                    st.markdown('<span style="background:#7c3aed;color:white;border-radius:6px;padding:2px 10px;font-size:11px;">📌 مرحّل من تاب المبيعات — محتاج جدولة | Transferred from Sales tab — needs scheduling</span>', unsafe_allow_html=True)
                st.markdown(f"📦 **المخزون | Stock:** {r['stock']} &nbsp;|&nbsp; 📈 **مبيع شهري | Monthly:** {r['sales_month']}")
                st.markdown("🛒 " + render_day_counts_md(r["day_counts"], day_dates, day_labels))
                st.markdown(f"💡 **اقتراح الكمية | Suggested Qty:** **{r['suggested_qty']}** &nbsp;|&nbsp; ⏳ **نفاد خلال | Days to stockout:** {r['days_to_stockout']} يوم")
                if r["sales_alert"]:
                    st.warning("📈 مبيعات أعلى من المعتاد كمان | Also selling faster than usual")
                badge_text, badge_color, _ = schedule_coverage_badge(r["sku"], r["days_to_stockout"], delay_days)
                recent_sched_r = get_recent_schedule_rows(days_back=4).get(r["sku_up"])
                show_normal_badge = not (recent_sched_r and "محتاج جدولة" in badge_text)
                if show_normal_badge:
                    st.markdown(f'<span style="background:{badge_color};color:white;border-radius:6px;padding:3px 10px;font-size:12px;">{badge_text}</span>', unsafe_allow_html=True)
                if recent_sched_r:
                    st.markdown(recent_schedule_badge_html(recent_sched_r), unsafe_allow_html=True)
                render_recent_expired_note(r["sku"])
                for note in get_unavailable_ordered_note(r["sku"]):
                    st.caption(note)
            st.divider()

    # ══ نحسب الأول قايمة "منتهي بالكامل" عشان نستبعدها من سكشن "مجدولة مؤخراً" ══
    missing_rows_t10 = compute_missing_inventory_rows(day_dates)

    # ══ SKUs مجدولة خلال آخر 4 أيام ومش ظاهرة أصلاً في القايمتين اللي فوق وتحت ══
    exclude_skus_t10 = {r["sku_up"] for r in stock_review_rows} | {r["sku_up"] for r in missing_rows_t10}
    recent_scheduled_rows_t10 = compute_recent_scheduled_rows(exclude_skus_t10, day_dates, days_back=4)
    render_recent_scheduled_section(recent_scheduled_rows_t10, day_dates, day_labels, "recent_scheduled_t10")

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
        recent_sched_map_t10 = get_recent_schedule_rows(days_back=4)
        for r in missing_rows_t10:
            c_img,c_info = st.columns([1,6])
            with c_img: show_img(r["img"],70)
            with c_info:
                st.markdown(f"**SKU:** `{r['sku']}`")
                st.error("⛔ مخزونه انتهى — مش موجود في ملف المخزون | Stock ran out — not found in inventory file")
                st.markdown("🛒 " + render_day_counts_md(r["day_counts"], day_dates, day_labels))
                st.markdown(f"📈 **مبيع شهري تقديري (بناءً على آخر 3 أيام) | Estimated Monthly Sales (based on last 3 days):** **{r['est_monthly_sales']}**")
                badge_text, badge_color, _ = schedule_coverage_badge(r["sku"], 0, delay_days)
                recent_sched_miss = recent_sched_map_t10.get(r["sku_up"])
                show_normal_badge_miss = not (recent_sched_miss and "محتاج جدولة" in badge_text)
                if show_normal_badge_miss:
                    st.markdown(f'<span style="background:{badge_color};color:white;border-radius:6px;padding:3px 10px;font-size:12px;">{badge_text}</span>', unsafe_allow_html=True)
                if recent_sched_miss:
                    st.markdown(recent_schedule_badge_html(recent_sched_miss), unsafe_allow_html=True)
                render_recent_expired_note(r["sku"])
                for note in get_unavailable_ordered_note(r["sku"]):
                    st.caption(note)
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
    all_review_rows2 = compute_stock_sales_rows(e1, day_dates2)
    valid_days_set = {1,2,3,4,5,6,7,8,10}
    sales_review_rows = [r for r in all_review_rows2
        if r["days_to_stockout_today"] in valid_days_set
        and r["sales_month"] > 0
        and r["sales_alert"]
        and not r["stock_alert"]]
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

        for r in sales_review_rows:
            c_img,c_info = st.columns([1,6])
            with c_img: show_img(r["img"],70)
            with c_info:
                st.markdown(f"**SKU:** `{r['sku']}`")
                st.markdown(f"📦 **المخزون | Stock:** {r['stock']} &nbsp;|&nbsp; 📈 **مبيع شهري | Monthly:** {r['sales_month']}")
                st.markdown("🛒 " + render_day_counts_md(r["day_counts"], day_dates2, day_labels2))
                st.markdown(f"⚡ **نفاد خلال بيع اليوم | Days to stockout (today's rate):** {r['days_to_stockout_today']} يوم")
                badge_text, badge_color, _ = schedule_coverage_badge(r["sku"], r["days_to_stockout"], delay_days2)
                recent_sched_r2 = get_recent_schedule_rows(days_back=4).get(r["sku_up"])
                show_normal_badge2 = not (recent_sched_r2 and "محتاج جدولة" in badge_text)
                if show_normal_badge2:
                    st.markdown(f'<span style="background:{badge_color};color:white;border-radius:6px;padding:3px 10px;font-size:12px;">{badge_text}</span>', unsafe_allow_html=True)
                if recent_sched_r2:
                    st.markdown(recent_schedule_badge_html(recent_sched_r2), unsafe_allow_html=True)
                render_recent_expired_note(r["sku"])
                for note in get_unavailable_ordered_note(r["sku"]):
                    st.caption(note)
            st.divider()

    # ══ نحسب الأول قايمة "منتهي بالكامل" عشان نستبعدها من سكشن "مجدولة مؤخراً" ══
    missing_rows_t13 = compute_missing_inventory_rows(day_dates2)

    # ══ SKUs مجدولة خلال آخر 4 أيام ومش ظاهرة أصلاً في القايمتين اللي فوق وتحت ══
    exclude_skus_t13 = {r["sku_up"] for r in sales_review_rows} | {r["sku_up"] for r in missing_rows_t13}
    recent_scheduled_rows_t13 = compute_recent_scheduled_rows(exclude_skus_t13, day_dates2, days_back=4)
    render_recent_scheduled_section(recent_scheduled_rows_t13, day_dates2, day_labels2, "recent_scheduled_t13")

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
        recent_sched_map_t13 = get_recent_schedule_rows(days_back=4)
        for r in missing_rows_t13:
            c_img,c_info = st.columns([1,6])
            with c_img: show_img(r["img"],70)
            with c_info:
                st.markdown(f"**SKU:** `{r['sku']}`")
                st.error("⛔ مخزونه انتهى — مش موجود في ملف المخزون | Stock ran out — not found in inventory file")
                st.markdown("🛒 " + render_day_counts_md(r["day_counts"], day_dates2, day_labels2))
                st.markdown(f"📈 **مبيع شهري تقديري (بناءً على آخر 3 أيام) | Estimated Monthly Sales (based on last 3 days):** **{r['est_monthly_sales']}**")
                badge_text, badge_color, _ = schedule_coverage_badge(r["sku"], 0, delay_days2)
                recent_sched_miss2 = recent_sched_map_t13.get(r["sku_up"])
                show_normal_badge_miss2 = not (recent_sched_miss2 and "محتاج جدولة" in badge_text)
                if show_normal_badge_miss2:
                    st.markdown(f'<span style="background:{badge_color};color:white;border-radius:6px;padding:3px 10px;font-size:12px;">{badge_text}</span>', unsafe_allow_html=True)
                if recent_sched_miss2:
                    st.markdown(recent_schedule_badge_html(recent_sched_miss2), unsafe_allow_html=True)
                render_recent_expired_note(r["sku"])
                for note in get_unavailable_ordered_note(r["sku"]):
                    st.caption(note)
            st.divider()

# ══ TAB 14 — المبيعات ══
with tab14:
    st.subheader("🛒 المبيعات اليومية | Daily Sales")
    st.caption("كل SKU عنده مخزون — مبيعاته اليومية من أمس للوراء بجانب مخزونه ومبيعاته الشهرية وحالة التغطية | All SKUs with inventory — daily sales, stock, monthly sales, and coverage status")

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

        st.divider()
        for r in sales_tab_rows:
            c_img, c_info = st.columns([1, 7])
            with c_img:
                show_img(r["img"], 70)
            with c_info:
                st.markdown(f"**SKU:** `{r['sku']}`")

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
                stock_self_ok = r["days_to_stockout"] >= coverage_days_t14 if r["effective_avg"] > 0 else False
                un_notes = get_unavailable_ordered_note(r["sku"])

                if stock_self_ok and not sched_t14:
                    cov_badge_text  = f"✅ مخزون كافٍ ({r['days_to_stockout']} يوم) — لا يحتاج جدولة الآن | Stock sufficient"
                    cov_badge_color = "#15803d"
                elif stock_self_ok and sched_t14:
                    sched_src_t14 = "تشييك" if sched_t14.get("source") == "Check" else "مجدول"
                    arrival_t14 = (sched_t14["parsed"] + timedelta(days=delay_days_t14)).date() if sched_t14.get("parsed") else None
                    cov_badge_text  = (f"✅ مخزون كافٍ ({r['days_to_stockout']} يوم) + ASN {sched_t14['asn']} بتاريخ {sched_t14['date']}"
                                       + (f" — وصول: {arrival_t14}" if arrival_t14 else "") + f" [{sched_src_t14}]")
                    cov_badge_color = "#15803d"
                else:
                    cov_badge_text  = badge_text_t14
                    cov_badge_color = badge_color_t14

                st.markdown(
                    f'<span style="background:{cov_badge_color};color:white;border-radius:6px;padding:3px 10px;font-size:12px;">{cov_badge_text}</span>',
                    unsafe_allow_html=True)

                # ══ ترحيل لتاب مخزون بدون بيع إذا كانت الحالة "محتاج جدولة" فقط بدون أي جدولة وبدون تفاصيل أخرى ══
                is_needs_sched_only = (
                    not stock_self_ok
                    and badge_text_t14 and "محتاج جدولة" in badge_text_t14
                    and not sched_t14
                    and not un_notes
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
                        st.caption(note)
                render_recent_expired_note(r["sku"])
            st.divider()
        # حفظ المرحلين في session_state بعد اكتمال العرض
        st.session_state["transferred_skus_t14"] = _new_transferred

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
                    sdata = get_cached(sheets[sheet_key])
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
                    st.caption(note)

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
                        st.caption(note)
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
