# -*- coding: utf-8 -*-
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import time
import pandas as pd
import io
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
    "Unavailable":       ["SKU","Quantity","Image URL","Date Added","Date Marked Unavailable","Times Marked Unavailable"],
    "Ordered":           ["SKU","Quantity","Image URL","Date Added","Order Count","Notes"],
    "Scheduled":         ["ASN","SKU","Quantity","Schedule Date","Image URL","Date Added","Notes","Flag"],
    "CancelledSchedule": ["ASN","SKU","Quantity","Schedule Date","Image URL","Date Added","Cancel Reason","Date Cancelled"],
    "Rescheduled":       ["ASN","SKU","Quantity","Old Schedule Date","Image URL","Date Added","Reschedule Reason","Date Moved"],
    "Expired":           ["ASN","SKU","Quantity","Schedule Date","Image URL","Date Added","Date Expired"],
    "Inventory":         ["SKU","Warehouse","Stock","Monthly Sales","Image URL","Date Uploaded"],
    "Settings":          ["Key","Value"],
    "Check":             ["ASN","SKU","Quantity","Schedule Date","Image URL","Date Added","Notes","Flag"],
    "CancelNotifications": ["ASN","SKUs","Schedule Date","Reason","Timestamp"],
}

def get_or_create_worksheet(tab, headers, retries=5, delay=2):
    for attempt in range(retries):
        try:
            return ss.worksheet(tab)
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

sheets = {}
for tab, headers in TABS_CONFIG.items():
    sheets[tab] = get_or_create_worksheet(tab, headers)

def sync_header(sheet, expected_headers, retries=3, delay=1):
    """لو الشيت قديم وناقصه أعمدة جديدة من TABS_CONFIG، يضيفها لصف الهيدر من غير ما يلمس البيانات"""
    for attempt in range(retries):
        try:
            current = sheet.row_values(1)
            if len(current) >= len(expected_headers):
                return
            missing = expected_headers[len(current):]
            for offset, col_name in enumerate(missing):
                sheet.update_cell(1, len(current) + offset + 1, col_name)
            return
        except gspread.exceptions.APIError:
            time.sleep(delay * (2 ** attempt))
        except Exception:
            return

for tab, headers in TABS_CONFIG.items():
    sync_header(sheets[tab], headers)

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
settings_sheet       = sheets["Settings"]
cancel_notif_sheet   = sheets["CancelNotifications"]

# ══ كاش ══
def get_cached(sheet, force=False):
    key = f"cache_{sheet.title}"
    if force or key not in st.session_state:
        st.session_state[key] = sheet.get_all_values()
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

def get_system_delay_days():
    """عدد الأيام اللي المخزون بياخدها لحد ما ينزل ع السيستم بعد وصول الجدولة"""
    val = load_settings().get("system_update_delay_days","2")
    try:
        d = int(float(val))
        return d if d >= 0 else 2
    except:
        return 2

# ══ links map ══
@st.cache_data(ttl=300)
def get_links_map():
    data = links_ws.get_all_values()
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

def upsert_unavailable(sku, qty, img, date_added):
    """لو الـ SKU موجود بالفعل في غير متوفر، حدّث صفه (الكمية + التاريخ + العداد) بدل ما يتضاف صف جديد"""
    dn = now_str()
    sku_clean = sku.strip()
    data = get_cached(unavailable_sheet, force=True)
    if len(data) > 1:
        for ri, r in enumerate(data[1:], start=2):
            while len(r) < 6: r.append("")
            if r[0].strip().upper() == sku_clean.upper():
                prev_count = _to_int(r[5]) if r[5] else 1
                new_count = prev_count + 1
                new_img = img if img else r[2]
                safe_update_row(unavailable_sheet, ri, [sku_clean, qty, new_img, r[3], dn, str(new_count)])
                return
    safe_append(unavailable_sheet, [sku_clean, qty, img, date_added, dn, "1"])

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def file_timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def rows_to_df(rows, headers):
    """يبني DataFrame من صفوف الشيت حتى لو أطوالها مش متطابقة مع الهيدر (صفوف قديمة قبل إضافة عمود جديد مثلاً)"""
    n = len(headers)
    fixed = []
    for r in rows:
        r = list(r)
        if len(r) < n:
            r = r + [""] * (n - len(r))
        elif len(r) > n:
            r = r[:n]
        fixed.append(r)
    return pd.DataFrame(fixed, columns=headers)

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

def dl_btn(df, prefix, label="⬇️ Excel | Download"):
    st.download_button(label, data=to_excel(df),
        file_name=f"{prefix}_{file_timestamp()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)

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
    expired_rows, keep = [], []
    for i, row in enumerate(data[1:], start=2):
        while len(row) < 6: row.append("")
        d = parse_excel_date(row[3])
        if d and today > d.date():
            expired_rows.append(row[:6] + [now_str()])
        else:
            keep.append(i)
    if expired_rows:
        safe_batch_append(expired_sheet, expired_rows)
        del_idx = sorted([x for x in range(2,len(data[1:])+2) if x not in keep], reverse=True)
        if del_idx:
            try:
                scheduled_sheet.delete_rows(del_idx[-1], del_idx[0] - del_idx[-1] + 1)
                clear_cache(scheduled_sheet)
            except Exception:
                for idx in del_idx:
                    safe_delete(scheduled_sheet, idx)

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
if "expired_checked" not in st.session_state:
    check_expired_scheduled()
    st.session_state["expired_checked"] = True

# تحميل الإشعارات من Sheets عند كل تشغيل (تدوم بعد الإغلاق)
if "cancel_notifs_loaded" not in st.session_state:
    st.session_state["check_cancel_notifications"] = load_cancel_notifications()
    st.session_state["cancel_notifs_loaded"] = True
elif "check_cancel_notifications" not in st.session_state:
    st.session_state["check_cancel_notifications"] = []

excluded_wh = get_excluded_warehouses()
inv_map     = build_inv_map(excluded_wh)
system_delay_days = get_system_delay_days()

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

# ══ خرائط حالة الـ SKU (للمخزون المنخفض) ══
def build_sku_schedule_map():
    """لكل SKU: أقرب جدولة قادمة (تاريخ + كمية + ASN) لسه ما اتشيكتش ولا اتكنسلتش"""
    data = get_cached(scheduled_sheet)
    m = {}
    if len(data) <= 1:
        return m
    for r in data[1:]:
        while len(r) < 6: r.append("")
        asn, sku, qty, sdate = r[0].strip(), r[1].strip().upper(), r[2], r[3]
        if not sku:
            continue
        d = parse_excel_date(sdate)
        entry = {"asn":asn,"qty":_to_int(qty),"date":sdate,"date_obj":d}
        if sku not in m:
            m[sku] = []
        m[sku].append(entry)
    # رتب كل واحد حسب أقرب تاريخ
    for sku in m:
        m[sku].sort(key=lambda e: e["date_obj"] if e["date_obj"] else datetime(2099,1,1))
    return m

def build_sku_check_map():
    """SKUs قيد التشييك (محولة من الجدولة) — لسه مش نهائية"""
    data = get_cached(sheets["Check"])
    m = {}
    if len(data) <= 1:
        return m
    for r in data[1:]:
        while len(r) < 6: r.append("")
        asn, sku, qty, sdate = r[0].strip(), r[1].strip().upper(), r[2], r[3]
        if not sku:
            continue
        d = parse_excel_date(sdate)
        entry = {"asn":asn,"qty":_to_int(qty),"date":sdate,"date_obj":d}
        m.setdefault(sku, []).append(entry)
    for sku in m:
        m[sku].sort(key=lambda e: e["date_obj"] if e["date_obj"] else datetime(2099,1,1))
    return m

def build_sku_ordered_set():
    data = get_cached(ordered_sheet)
    s = set()
    if len(data) <= 1:
        return s
    for r in data[1:]:
        if r and r[0].strip():
            s.add(r[0].strip().upper())
    return s

def build_sku_unavailable_map():
    """SKU -> {date, count} — تاريخ آخر مرة اتسجل غير متوفر وعدد المرات"""
    data = get_cached(unavailable_sheet)
    m = {}
    if len(data) <= 1:
        return m
    for r in data[1:]:
        while len(r) < 6: r.append("")
        sku = r[0].strip()
        if sku:
            m[sku.upper()] = {"date": r[4], "count": _to_int(r[5]) if r[5] else 1}
    return m

# للتوافق الخلفي مع أي استخدام كـ set
def build_sku_unavailable_set():
    return set(build_sku_unavailable_map().keys())


ordinal_map = {1:"الثانية|Second",2:"الثالثة|Third",3:"الرابعة|Fourth",4:"الخامسة|Fifth"}

# ══════════════════════════════════════════════
st.title("📦 Stock Requests | طلبات المخزون")

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
    "🔴 مخزون منخفض | Low Stock",
    "🗂️ منتهية | Expired",
    "⚙️ الإعدادات | Settings",
])
(tab1,tab2,tab3,tab4,tab5,tab_check,tab6,tab7,tab8,tab9,tab10,tab11,tab12) = tabs

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
                for r in rows:
                    upsert_unavailable(r[0], r[1], r[2] if len(r)>2 else "", r[3] if len(r)>3 else "")
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
        for i, row in enumerate(rows, start=2):
            while len(row) < 5: row.append("")
            sku,qty,img,date_added,fname = row[0],row[1],row[2],row[3],row[4]
            c_img,c_info,c_act = st.columns([1,4,3])
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
                            st.rerun()
                with cb:
                    if st.button("❌ غير\nمتوفر\nUnavailable", key=f"unavail_{i}"):
                        upsert_unavailable(sku, qty, img, date_added)
                        safe_delete(requests_sheet,i)
                        st.rerun()
                with cc:
                    if st.button("🛒 طلب\nOrder", key=f"order_{i}"):
                        dn = now_str()
                        prev = ordered_skus.get(sku.upper(),0)
                        if prev > 0:
                            o_rows = get_cached(ordered_sheet)
                            for oi, or_ in enumerate(o_rows[1:], start=2):
                                if or_[0].strip().upper() == sku.upper():
                                    new_count = prev + 1
                                    ordn = ordinal_map.get(prev, f"{prev+1}")
                                    note = f"تم الطلب للمرة {ordn} | Ordered {ordn} time"
                                    safe_update_row(ordered_sheet, oi, [or_[0],qty,or_[2],dn,str(new_count),note])
                                    break
                        else:
                            safe_append(ordered_sheet,[sku,qty,img,dn,"1",""])
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
        df_ap = rows_to_df(rows_ap, data_ap[0])
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
        df_un = rows_to_df(rows_un, data_un[0])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_un,"unavailable")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_un", use_container_width=True):
                st.session_state["confirm_clear_un"] = True
        confirm_clear("clear_un", unavailable_sheet, "غير المتوفر | Unavailable")
        st.write(f"**عرض | Showing: {len(filtered)} / {len(rows_un)}**")
        for ri, row in filtered:
            while len(row)<6: row.append("")
            sku,qty,img,da,dm,cnt = row[0],row[1],row[2],row[3],row[4],row[5]
            c_img,c_info,c_act = st.columns([1,4,2])
            with c_img: show_img(img,70)
            with c_info:
                cnt_badge = ""
                try:
                    cnt_n = int(float(cnt)) if cnt else 1
                except:
                    cnt_n = 1
                if cnt_n > 1:
                    cnt_badge = f' &nbsp;<span style="background:#7c2d12;color:#fdba74;border-radius:6px;padding:2px 10px;font-size:12px;font-weight:bold;">🔁 تكرر {cnt_n} مرات | Repeated {cnt_n}x</span>'
                st.markdown(f"**SKU:** `{sku}`" + cnt_badge, unsafe_allow_html=True)
                show_sku_inv(sku)
                st.markdown(f"**Qty طلب | Requested:** {qty}")
                st.caption(f"📅 Requested | طُلب: {da} | ❌ Unavailable | غير متوفر: {dm}")
            with c_act:
                ca,cb = st.columns(2)
                with ca:
                    with st.popover("↩️ رجّع للموافقة\nReturn to Approved"):
                        st.caption("لو المخزون بقى متوفر | If stock is now available")
                        nq = st.text_input("الكمية المعتمدة | Approved Qty", value=qty, key=f"ret_un_qty_{ri}")
                        if st.button("✅ أرسل للموافقة | Send to Approved", key=f"ret_un_conf_{ri}"):
                            safe_append(approved_sheet,[sku,qty,nq,img,da,now_str()])
                            safe_delete(unavailable_sheet,ri)
                            st.success("✅ تم الإرجاع للموافقة | Returned to Approved")
                            st.rerun()
                with cb:
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
        df_ord = rows_to_df(rows_ord, data_ord[0])
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
            c_img,c_info,c_act = st.columns([1,4,2])
            with c_img: show_img(img,70)
            with c_info:
                st.markdown(f"**SKU:** `{sku}`")
                show_sku_inv(sku)
                st.markdown(f"**Quantity | الكمية:** {qty}")
                if note:
                    st.warning(f"🔁 {note}")
                st.caption(f"📅 Date | التاريخ: {da} | 🔢 Order Count | عدد الطلبات: {cnt}")
            with c_act:
                ca,cb = st.columns(2)
                with ca:
                    with st.popover("↩️ رجّع\nReturn"):
                        nq = st.text_input("الكمية المعدّلة | Adjusted Qty", value=qty, key=f"ret_qty_{ri}")
                        if st.button("✅ أرسل للموافقة | Send to Approved", key=f"ret_conf_{ri}"):
                            safe_append(approved_sheet,[sku,qty,nq,img,da,now_str()])
                            safe_delete(ordered_sheet,ri)
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

        df_sch = rows_to_df(rows_sch, data_sch[0])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_sch,"scheduled")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_sc", use_container_width=True):
                st.session_state["confirm_clear_sc"] = True
        confirm_clear("clear_sc", scheduled_sheet, "الجدولة | Schedule")

        srch_asn = st.text_input("🔍 بحث ASN | Search by ASN", key="srch_asn", placeholder="اكتب رقم ASN...")
        srch_sku_sch = st.text_input("🔍 بحث SKU | Search by SKU", key="srch_sku_sch", placeholder="اكتب SKU...")
        today = datetime.now().date()
        st.write(f"**إجمالي ASN | Total ASNs: {len(asn_groups)}**")

        for asn, group in asn_groups.items():
            if srch_asn and srch_asn.strip().upper() not in asn.upper():
                continue
            if srch_sku_sch:
                skus_upper_in_group = {r[1].strip().upper() for r in group["skus"]}
                if not any(srch_sku_sch.strip().upper() in s for s in skus_upper_in_group):
                    continue
            sdate   = group["date"]
            pd_date = parse_excel_date(sdate)
            is_exp  = pd_date and today > pd_date.date()
            skus_   = group["skus"]
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

        df_chk = rows_to_df(rows_chk, data_chk[0])
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
        df_can = rows_to_df(rows_can, data_can[0])
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

        df_res = rows_to_df(rows_res, data_res[0])
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
        df_inv_dl = rows_to_df(raw_inv[1:], raw_inv[0])
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

# ══ TAB 10 — مخزون منخفض ══
with tab10:
    st.subheader("🔴 مخزون منخفض | Low Stock")
    st.caption("المخزون الإجمالي أقل من 50% من المبيع الشهري | Total stock < 50% of monthly sales")
    st.caption(f"⏱️ فرضية تأخر تحديث المخزون على السيستم بعد وصول الجدولة: **{system_delay_days} يوم** — عدّلها من تاب الإعدادات | Assumed system stock-update delay after schedule arrival: **{system_delay_days} day(s)** — adjustable in Settings tab")

    today = datetime.now().date()
    sched_map  = build_sku_schedule_map()
    check_map  = build_sku_check_map()
    ordered_set = build_sku_ordered_set()
    unavail_map = build_sku_unavailable_map()

    low_stock = []
    for sku_key,info in inv_map.items():
        total=info["total_stock"]; sales=info["sales"]
        if sales>0 and total<sales*0.5:
            pct=round(total/sales*100,1)
            low_stock.append((info["sku"],total,sales,pct,info["img"]))
    low_stock.sort(key=lambda x:x[3])

    if not inv_map:
        st.info("ارفع ملف المخزون أولاً | Upload Inventory first")
    elif not low_stock:
        st.success("✅ كل المخزون كافي | All stock levels sufficient (≥ 50% of sales)")
    else:
        # ══ تجهيز بيانات كل SKU مع الحالة والتحليل ══
        enriched = []
        for sku,total,sales,pct,img in low_stock:
            sku_up = sku.strip().upper()
            daily_sales = (sales / 30.0) if sales > 0 else 0
            coverage_days = (total / daily_sales) if daily_sales > 0 else None

            # أقرب جدولة قادمة (من تاب الجدولة أو التشييك)
            sched_entries = sched_map.get(sku_up, []) + check_map.get(sku_up, [])
            sched_entries = [e for e in sched_entries if e["date_obj"]]
            sched_entries.sort(key=lambda e: e["date_obj"])
            next_sched = sched_entries[0] if sched_entries else None

            is_ordered  = sku_up in ordered_set
            unavail_info = unavail_map.get(sku_up)
            unavail_date = unavail_info["date"] if unavail_info else None
            unavail_count = unavail_info["count"] if unavail_info else None
            is_unavail  = unavail_info is not None

            # ══ تحديد الحالة الأساسية ══
            gap_days = None
            schedule_covers = None
            if next_sched:
                days_until_sched = (next_sched["date_obj"].date() - today).days
                required_coverage_days = max(days_until_sched, 0) + system_delay_days
                if coverage_days is not None:
                    schedule_covers = coverage_days >= required_coverage_days
                    gap_days = round(required_coverage_days - coverage_days, 1)
                status = "scheduled_ok" if schedule_covers else "scheduled_gap"
            elif is_unavail:
                status = "unavailable"
            elif is_ordered:
                status = "ordered"
            else:
                status = "action_needed"

            enriched.append({
                "sku":sku,"total":total,"sales":sales,"pct":pct,"img":img,
                "daily_sales":round(daily_sales,2),"coverage_days":coverage_days,
                "next_sched":next_sched,"status":status,
                "schedule_covers":schedule_covers,"gap_days":gap_days,
                "is_ordered":is_ordered,"is_unavail":is_unavail,"unavail_date":unavail_date,"unavail_count":unavail_count,
            })

        STATUS_META = {
            "action_needed": {"label":"🆘 محتاج إجراء",          "full":"🆘 محتاج إجراء | Needs Action",          "color":"#ef4444"},
            "scheduled_gap": {"label":"⚠️ مجدول بفجوة",          "full":"⚠️ مجدول لكن الجدولة مش هتكفي | Scheduled but Gap","color":"#f97316"},
            "scheduled_ok":  {"label":"📅 مجدول وكافي",          "full":"📅 مجدول وكافي | Scheduled & Covered",   "color":"#3b82f6"},
            "ordered":       {"label":"🛒 تم الطلب",             "full":"🛒 تم الطلب | Ordered",                  "color":"#a855f7"},
            "unavailable":   {"label":"❌ غير متوفر",            "full":"❌ غير متوفر | Unavailable",             "color":"#6b7280"},
        }

        # ══ تنزيل + ملخص عام ══
        c1,c2 = st.columns(2)
        with c1:
            df_low = pd.DataFrame([
                {"SKU":e["sku"],"Total Stock":e["total"],"Monthly Sales":e["sales"],"Stock %":e["pct"],
                 "Daily Sales":e["daily_sales"],"Coverage Days":e["coverage_days"],
                 "Status":STATUS_META[e["status"]]["full"],
                 "Next Schedule Date": e["next_sched"]["date"] if e["next_sched"] else "",
                 "Gap Days": e["gap_days"] if e["gap_days"] is not None else "",
                 "Unavailable Since": e["unavail_date"] or "",
                 "Image URL":e["img"]}
                for e in enriched
            ])
            dl_btn(df_low,"low_stock")
        with c2:
            st.error(f"🔴 SKUs منخفضة | Low Stock SKUs: {len(low_stock)}")

        # ══ تجميع كل SKU تحت حالته ══
        status_groups = {k: [] for k in STATUS_META}
        for e in enriched:
            status_groups[e["status"]].append(e)

        def render_low_stock_card(e):
            sku, total, sales, pct, img = e["sku"], e["total"], e["sales"], e["pct"], e["img"]
            if pct<20:   sev_color="#ef4444"; sev_label="⛔ حرج جداً | Critical"
            elif pct<35: sev_color="#f97316"; sev_label="🔴 منخفض جداً | Very Low"
            else:        sev_color="#eab308"; sev_label="🟡 منخفض | Low"

            c_img,c_info = st.columns([1,6])
            with c_img: show_img(img,70)
            with c_info:
                st.markdown(f"**SKU:** `{sku}`")
                show_sku_inv(sku)
                st.progress(min(pct/100,1.0))
                st.markdown(f"{sev_label} — **{pct}%** &nbsp;|&nbsp; مخزون | Stock: **{total}** / مبيع شهري | Monthly Sales: **{sales}**")
                if e["daily_sales"]:
                    cov_txt = f"{e['coverage_days']:.1f} يوم | days" if e["coverage_days"] is not None else "—"
                    st.caption(f"📈 مبيع يومي تقريبي | Approx Daily Sales: **{e['daily_sales']}** &nbsp;|&nbsp; ⏳ تغطية المخزون الحالي | Current Stock Covers: **{cov_txt}**")

                if e["status"] == "scheduled_ok":
                    ns = e["next_sched"]
                    st.success(f"📅 أقرب جدولة | Next Schedule: **{ns['date']}** (ASN: `{ns['asn']}`, Qty: {ns['qty']}) — هتوصل وتتفعل قبل ما المخزون يخلص | Will arrive & activate before stock runs out ✅")
                elif e["status"] == "scheduled_gap":
                    ns = e["next_sched"]
                    st.warning(
                        f"📅 أقرب جدولة | Next Schedule: **{ns['date']}** (ASN: `{ns['asn']}`, Qty: {ns['qty']}) — "
                        f"⚠️ فيه فجوة محتملة قبل ما المخزون الجديد يتفعل على السيستم بحوالي **{e['gap_days']} يوم** "
                        f"(محسوبة باعتبار {system_delay_days} يوم تأخر تحديث) | "
                        f"Possible stockout gap before new stock reflects on system: ~**{e['gap_days']} day(s)**")
                elif e["status"] == "ordered":
                    st.info("🛒 موجود في تاب «تم الطلب» — مفيش جدولة بتاريخ محدد لسه | Currently in 'Ordered' tab — no dated schedule yet")
                elif e["status"] == "unavailable":
                    repeat_txt = f" &nbsp;| 🔁 تكرر {e['unavail_count']} مرات | Repeated {e['unavail_count']}x" if e["unavail_count"] and e["unavail_count"] > 1 else ""
                    st.error(f"❌ مسجل في تاب «غير متوفر» | Marked as 'Unavailable' — 📅 بتاريخ | Since: **{e['unavail_date'] or '—'}**{repeat_txt}")
                else:
                    st.error("🆘 مفيش جدولة ولا طلب ولا حتى مسجل غير متوفر — يحتاج متابعة فورية | No schedule, no order, not marked unavailable — needs immediate attention")
            st.divider()

        # ══ تابات فرعية لكل حالة ══
        sub_labels = [f'{STATUS_META[k]["label"]} ({len(status_groups[k])})' for k in STATUS_META]
        sub_tabs = st.tabs(sub_labels)

        for (status_key, meta), sub_tab in zip(STATUS_META.items(), sub_tabs):
            with sub_tab:
                group_items = status_groups[status_key]
                if not group_items:
                    st.success("✅ لا يوجد عناصر في هذه الحالة | No items in this status")
                    continue
                srch_key = f"srch_low_{status_key}"
                srch_low = st.text_input("🔍 بحث SKU | Search SKU", key=srch_key, placeholder="اكتب SKU...")
                st.caption(f"📦 العدد | Count: {len(group_items)}")
                st.divider()
                for e in group_items:
                    if srch_low and srch_low.strip().upper() not in e["sku"].upper():
                        continue
                    render_low_stock_card(e)



# ══ TAB 11 — منتهية الصلاحية ══
with tab11:
    st.subheader("🗂️ الجدولة منتهية الصلاحية | Expired Schedule")
    data_ex = get_cached(expired_sheet)
    if len(data_ex) <= 1:
        st.info("لا يوجد منتهي | No expired items.")
    else:
        rows_ex = data_ex[1:]
        df_ex = rows_to_df(rows_ex, data_ex[0])
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
    st.markdown("### ⏱️ تأخر تحديث المخزون على السيستم | System Stock Update Delay")
    st.caption("بعد ما الجدولة توصل، المخزون مش بينزل ع السيستم فورًا — حدد كام يوم بياخد عشان نحسب تغطية المخزون المنخفض بدقة | After a schedule arrives, stock doesn't reflect on the system instantly — set how many days it takes so Low Stock coverage is calculated accurately")
    current_delay = system_delay_days
    new_delay = st.number_input("عدد الأيام | Number of days", min_value=0, max_value=14, value=current_delay, step=1, key="delay_days_input")
    if st.button("💾 حفظ مدة التأخير | Save Delay", type="primary", key="save_delay_btn"):
        save_setting("system_update_delay_days", str(int(new_delay)))
        st.success("✅ تم الحفظ | Saved"); st.rerun()
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
