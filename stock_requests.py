# -*- coding: utf-8 -*-
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import time
import pandas as pd
import io
import json
import gspread.exceptions

# ══════════════════════════════════════════════
# إعدادات الصفحة
# ══════════════════════════════════════════════
st.set_page_config(page_title="📦 Stock Requests | طلبات المخزون", page_icon="📦", layout="wide")

# ══════════════════════════════════════════════
# اتصال جوجل شيت
# ══════════════════════════════════════════════
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds_dict = st.secrets["gcp_service_account"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)
ss = client.open("Complaints")

# ══════════════════════════════════════════════
# الأوراق
# ══════════════════════════════════════════════
TABS_CONFIG = {
    "Requests":    ["SKU", "Quantity", "Image URL", "Date Added", "File Name"],
    "Approved":    ["SKU", "Quantity Requested", "Quantity Approved", "Image URL", "Date Added", "Date Approved"],
    "Unavailable": ["SKU", "Quantity", "Image URL", "Date Added", "Date Marked Unavailable"],
    "Scheduled":   ["ASN", "SKU", "Quantity", "Schedule Date", "Image URL", "Date Added"],
    "Expired":     ["ASN", "SKU", "Quantity", "Schedule Date", "Image URL", "Date Added", "Date Expired"],
    "Inventory":   ["SKU", "Warehouse", "Current Stock", "Monthly Sales", "Image URL", "Date Uploaded"],
    "Settings":    ["Key", "Value"],
}

sheets = {}
for tab, headers in TABS_CONFIG.items():
    try:
        ws = ss.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab, rows="3000", cols="12")
        ws.append_row(headers)
    sheets[tab] = ws

try:
    links_ws = ss.worksheet("links n")
except gspread.exceptions.WorksheetNotFound:
    links_ws = ss.add_worksheet(title="links n", rows="2000", cols="2")
    links_ws.append_row(["SKU", "Image URL"])

requests_sheet    = sheets["Requests"]
approved_sheet    = sheets["Approved"]
unavailable_sheet = sheets["Unavailable"]
scheduled_sheet   = sheets["Scheduled"]
expired_sheet     = sheets["Expired"]
inventory_sheet   = sheets["Inventory"]
settings_sheet    = sheets["Settings"]

# ══════════════════════════════════════════════
# كاش
# ══════════════════════════════════════════════
def get_cached(sheet, force=False):
    key = f"cache_{sheet.title}"
    if force or key not in st.session_state:
        st.session_state[key] = sheet.get_all_values()
    return st.session_state[key]

def clear_cache(sheet):
    key = f"cache_{sheet.title}"
    if key in st.session_state:
        del st.session_state[key]

# ══════════════════════════════════════════════
# الإعدادات (محفوظة في جوجل شيت)
# ══════════════════════════════════════════════
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
    s = load_settings()
    val = s.get("excluded_warehouses", "")
    if not val:
        return set()
    return {w.strip().upper() for w in val.split(",") if w.strip()}

# ══════════════════════════════════════════════
# links map
# ══════════════════════════════════════════════
@st.cache_data(ttl=300)
def get_links_map():
    data = links_ws.get_all_values()
    mapping = {}
    for row in data[1:]:
        if len(row) >= 2 and row[0].strip():
            mapping[row[0].strip().upper()] = row[1].strip()
    return mapping

# ══════════════════════════════════════════════
# بناء inv_map مع مراعاة المستودعات المستثناة
# ══════════════════════════════════════════════
def build_inv_map(excluded_wh=None):
    """
    SKU → {
        total_stock: int,
        sales: int,
        img: str,
        date: str,
        warehouses: {wh_name: stock}
    }
    """
    if excluded_wh is None:
        excluded_wh = set()
    inv_data = get_cached(inventory_sheet)
    inv_map = {}
    if len(inv_data) > 1:
        for r in inv_data[1:]:
            while len(r) < 6: r.append("")
            sku, wh, stock, sales, img, date_up = r[0], r[1], r[2], r[3], r[4], r[5]
            if not sku: continue
            sku_up = sku.upper()
            wh_up  = wh.upper()
            if sku_up not in inv_map:
                inv_map[sku_up] = {
                    "sku": sku,
                    "total_stock": 0,
                    "sales": 0,
                    "img": img,
                    "date": date_up,
                    "warehouses": {}
                }
            inv_map[sku_up]["warehouses"][wh] = inv_map[sku_up]["warehouses"].get(wh, 0) + _to_int(stock)
            if wh_up not in excluded_wh:
                inv_map[sku_up]["total_stock"] += _to_int(stock)
            try:
                inv_map[sku_up]["sales"] = int(float(sales))
            except: pass
            if not inv_map[sku_up]["img"]:
                inv_map[sku_up]["img"] = img
    return inv_map

def _to_int(v):
    try: return int(float(v))
    except: return 0

# ══════════════════════════════════════════════
# دوال مساعدة
# ══════════════════════════════════════════════
def safe_append(sheet, row, retries=4, delay=1):
    for _ in range(retries):
        try:
            sheet.append_row(row, value_input_option="USER_ENTERED")
            clear_cache(sheet)
            return True
        except Exception:
            time.sleep(delay)
    return False

def safe_delete(sheet, row_idx, retries=4, delay=1):
    for _ in range(retries):
        try:
            sheet.delete_rows(row_idx)
            clear_cache(sheet)
            return True
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

def safe_batch_append(sheet, rows_data, retries=4, delay=1):
    for _ in range(retries):
        try:
            sheet.append_rows(rows_data, value_input_option="USER_ENTERED")
            clear_cache(sheet)
            return True
        except Exception:
            time.sleep(delay)
    return False

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def file_timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def to_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()

def make_empty_template(columns):
    return to_excel(pd.DataFrame(columns=columns))

def parse_excel_date(val):
    try:
        if isinstance(val, (int, float)):
            return datetime(1899, 12, 30) + timedelta(days=int(val))
        return datetime.strptime(str(val).strip()[:10], "%Y-%m-%d")
    except:
        return None

def check_expired_scheduled():
    data = get_cached(scheduled_sheet, force=True)
    if len(data) <= 1:
        return
    today = datetime.now().date()
    expired_rows = []
    keep = []
    for i, row in enumerate(data[1:], start=2):
        while len(row) < 6: row.append("")
        sched_date = parse_excel_date(row[3])
        if sched_date and (today > sched_date.date() + timedelta(days=1)):
            expired_rows.append(row + [now_str()])
        else:
            keep.append(i)
    if expired_rows:
        safe_batch_append(expired_sheet, expired_rows)
        del_indices = sorted([x for x in range(2, len(data[1:])+2) if x not in keep], reverse=True)
        for idx in del_indices:
            safe_delete(scheduled_sheet, idx)

# ══════════════════════════════════════════════
# CSS
# ══════════════════════════════════════════════
st.markdown("""
<style>
    .stTabs [data-baseweb="tab-list"] { gap: 6px; flex-wrap: wrap; }
    .stTabs [data-baseweb="tab"] {
        background: #1e293b; color: white; border-radius: 8px;
        padding: 7px 14px; font-weight: bold; font-size: 12px;
    }
    .stTabs [aria-selected="true"] { background: #3b82f6 !important; }
    .inv-card {
        background: #0f172a; border-radius: 10px;
        padding: 10px 14px; margin-bottom: 6px;
        border-left: 4px solid #3b82f6;
    }
    .wh-badge {
        display: inline-block; background: #1e3a5f;
        border-radius: 6px; padding: 2px 8px;
        margin: 2px; font-size: 12px; color: #93c5fd;
    }
    .alert-card {
        border-left: 5px solid #ef4444;
        background: #2d1515; border-radius: 8px;
        padding: 10px 14px; margin-bottom: 6px;
    }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════
# Init
# ══════════════════════════════════════════════
if "expired_checked" not in st.session_state:
    check_expired_scheduled()
    st.session_state["expired_checked"] = True

excluded_wh = get_excluded_warehouses()

# ══════════════════════════════════════════════
# دوال عرض مشتركة
# ══════════════════════════════════════════════
def show_img(img, width=75):
    if img and str(img).startswith("http"):
        st.image(img, width=width)
    else:
        st.markdown("🖼️")

def show_sku_inv(sku, inv_map, excluded_wh):
    """يعرض فوق الصورة: المبيع والمخزون ومستودعات بتفصيل"""
    info = inv_map.get(sku.upper())
    if not info:
        return
    total  = info["total_stock"]
    sales  = info["sales"]
    wh_map = info["warehouses"]
    # بيانات المستودعات (مع تمييز المستثناة)
    wh_parts = []
    for wh, stk in wh_map.items():
        if wh.upper() in excluded_wh:
            wh_parts.append(f"~~{wh}: {stk}~~")
        else:
            wh_parts.append(f"**{wh}:** {stk}")
    st.markdown(
        f"📈 **مبيع شهري | Monthly Sales:** {sales} &nbsp;|&nbsp; "
        f"📦 **مخزون | Stock:** {total}"
    )
    st.caption("🏭 " + " | ".join(wh_parts))

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

def dl_btn(df, prefix):
    st.download_button(
        "⬇️ Excel",
        data=to_excel(df),
        file_name=f"{prefix}_{file_timestamp()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# ══════════════════════════════════════════════
st.title("📦 Stock Requests | طلبات المخزون")

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "📋 الطلبات | Requests",
    "✅ الموافقة | Approved",
    "❌ غير متوفر | Unavailable",
    "📅 الجدولة | Scheduled",
    "⚠️ تنبيهات | Alerts",
    "📊 المخزون | Inventory",
    "🗂️ منتهية | Expired",
    "⚙️ الإعدادات | Settings",
])

inv_map = build_inv_map(excluded_wh)

# ══════════════════════════════════════════════
# TAB 1 — الطلبات
# ══════════════════════════════════════════════
with tab1:
    st.subheader("➕ إضافة طلبات | Add Requests")
    links_map = get_links_map()

    col_method, col_tmpl = st.columns([3, 1])
    with col_tmpl:
        st.download_button("⬇️ Template فارغ | Empty Template",
            data=make_empty_template(["SKU", "Quantity"]),
            file_name=f"request_template_{file_timestamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with col_method:
        method = st.radio("طريقة الإضافة | Method:", ["📂 رفع ملف | Upload", "✏️ لصق | Paste"], horizontal=True)

    added_rows = []
    file_name_label = ""

    if "Upload" in method:
        uploaded = st.file_uploader("ارفع Excel أو CSV | Upload Excel or CSV", type=["xlsx","xls","csv"])
        if uploaded:
            file_name_label = uploaded.name
            try:
                df_up = pd.read_csv(uploaded, dtype=str).fillna("") if uploaded.name.endswith(".csv") \
                    else pd.read_excel(uploaded, dtype=str).fillna("")
                sku_col = qty_col = None
                for c in df_up.columns:
                    cl = c.strip().lower()
                    if cl in ("sku","item","product","item nr","item_nr"): sku_col = c
                    if cl in ("quantity","qty","كمية","الكمية","amount"): qty_col = c
                if not sku_col: sku_col = df_up.columns[0]
                if not qty_col and len(df_up.columns) > 1: qty_col = df_up.columns[1]
                st.info(f"📊 {len(df_up)} صف | rows")
                st.dataframe(df_up[[c for c in [sku_col, qty_col] if c]], use_container_width=True, height=150)
                for _, row in df_up.iterrows():
                    sku = str(row[sku_col]).strip()
                    qty = str(row[qty_col]).strip() if qty_col else ""
                    img = links_map.get(sku.upper(), "")
                    if sku and sku.lower() != "nan":
                        added_rows.append((sku, qty, img))
            except Exception as e:
                st.error(f"❌ {e}")
    else:
        pasted = st.text_area("الصق هنا | Paste here (SKU,Qty):", height=110, placeholder="SKU001,5\nSKU002,3")
        file_name_label = "Manual Entry"
        if pasted.strip():
            for line in pasted.strip().splitlines():
                parts = [p.strip() for p in line.split(",")]
                sku = parts[0] if parts else ""
                qty = parts[1] if len(parts) > 1 else ""
                img = links_map.get(sku.upper(), "")
                if sku: added_rows.append((sku, qty, img))
            if added_rows: st.success(f"✅ {len(added_rows)} صف | rows ready")

    if added_rows:
        if st.button("📤 إضافة | Add", type="primary"):
            date_now = now_str()
            rows_to_add = [[sku, qty, img, date_now, file_name_label] for sku, qty, img in added_rows]
            if safe_batch_append(requests_sheet, rows_to_add):
                st.success(f"✅ أُضيف {len(rows_to_add)} صف")
                st.rerun()

    st.divider()
    st.subheader("📋 الطلبات الحالية | Current Requests")
    data = get_cached(requests_sheet)

    if len(data) <= 1:
        st.info("لا توجد طلبات | No requests yet.")
    else:
        rows = data[1:]
        df_req = pd.DataFrame(rows, columns=data[0])

        c1, c2, c3, c4 = st.columns(4)
        with c1: dl_btn(df_req, "requests")
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
            st.warning("⚠️ موافقة على كل الطلبات؟")
            cy, cn = st.columns(2)
            if cy.button("✅ نعم", key="yes_app_all"):
                date_now = now_str()
                to_add = []
                for row in rows:
                    while len(row) < 5: row.append("")
                    to_add.append([row[0], row[1], row[1], row[2], row[3], date_now])
                safe_batch_append(approved_sheet, to_add)
                safe_delete_all(requests_sheet)
                st.session_state["confirm_approve_all"] = False
                st.rerun()
            if cn.button("❌ لا", key="no_app_all"):
                st.session_state["confirm_approve_all"] = False
                st.rerun()

        if st.session_state.get("confirm_reject_all"):
            st.warning("⚠️ رفض كل الطلبات كغير متوفرة؟")
            cy, cn = st.columns(2)
            if cy.button("✅ نعم", key="yes_rej_all"):
                date_now = now_str()
                to_add = []
                for row in rows:
                    while len(row) < 5: row.append("")
                    to_add.append([row[0], row[1], row[2], row[3], date_now])
                safe_batch_append(unavailable_sheet, to_add)
                safe_delete_all(requests_sheet)
                st.session_state["confirm_reject_all"] = False
                st.rerun()
            if cn.button("❌ لا", key="no_rej_all"):
                st.session_state["confirm_reject_all"] = False
                st.rerun()

        confirm_clear("clear_req", requests_sheet, "الطلبات | Requests")

        st.write(f"**الإجمالي | Total: {len(rows)}**")
        for i, row in enumerate(rows, start=2):
            while len(row) < 5: row.append("")
            sku, qty, img, date_added, fname = row[0], row[1], row[2], row[3], row[4]
            col_img, col_info, col_actions = st.columns([1, 4, 3])
            with col_img:
                show_img(img, 75)
            with col_info:
                st.markdown(f"**SKU:** `{sku}`")
                show_sku_inv(sku, inv_map, excluded_wh)
                st.markdown(f"**طلب | Requested Qty:** {qty}")
                st.caption(f"📅 {date_added} | 📁 {fname}")
            with col_actions:
                ca, cb, cc = st.columns(3)
                with ca:
                    with st.popover("✅ وافق | Approve"):
                        new_qty = st.text_input("Approved Qty", value=qty, key=f"aqty_{i}")
                        if st.button("✅ تأكيد", key=f"aconf_{i}"):
                            safe_append(approved_sheet, [sku, qty, new_qty, img, date_added, now_str()])
                            safe_delete(requests_sheet, i)
                            st.rerun()
                with cb:
                    if st.button("❌ غير متوفر\nUnavailable", key=f"unavail_{i}"):
                        safe_append(unavailable_sheet, [sku, qty, img, date_added, now_str()])
                        safe_delete(requests_sheet, i)
                        st.rerun()
                with cc:
                    if st.button("🗑️ حذف\nDelete", key=f"del_req_{i}"):
                        safe_delete(requests_sheet, i)
                        st.rerun()
            st.divider()

# ══════════════════════════════════════════════
# TAB 2 — الموافقة
# ══════════════════════════════════════════════
with tab2:
    st.subheader("✅ الطلبات الموافق عليها | Approved")
    data_ap = get_cached(approved_sheet)

    if len(data_ap) <= 1:
        st.info("لا توجد موافقات | No approvals yet.")
    else:
        rows_ap = data_ap[1:]
        search_ap = st.text_input("🔍 بحث SKU | Search SKU", key="srch_ap", placeholder="اكتب SKU...")
        filtered_ap = [r for r in rows_ap if not search_ap or search_ap.strip().upper() in r[0].upper()]

        df_ap = pd.DataFrame(rows_ap, columns=data_ap[0])
        c1, c2 = st.columns(2)
        with c1: dl_btn(df_ap, "approved")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_ap", use_container_width=True):
                st.session_state["confirm_clear_ap"] = True
        confirm_clear("clear_ap", approved_sheet, "الموافقة | Approved")

        st.write(f"**عرض | Showing: {len(filtered_ap)} / {len(rows_ap)}**")
        for row in filtered_ap:
            real_i = rows_ap.index(row) + 2
            while len(row) < 6: row.append("")
            sku, qty_req, qty_app, img, date_add, date_app = row[0], row[1], row[2], row[3], row[4], row[5]
            col_img, col_info, col_del = st.columns([1, 5, 1])
            with col_img:
                show_img(img, 70)
            with col_info:
                st.markdown(f"**SKU:** `{sku}`")
                show_sku_inv(sku, inv_map, excluded_wh)
                if qty_app and qty_app != qty_req:
                    st.markdown(f"**طلبت | Req:** {qty_req} → **وافقوا | App:** ⚠️ **{qty_app}**")
                else:
                    st.markdown(f"**Quantity:** {qty_app}")
                st.caption(f"📅 Requested: {date_add} | ✅ Approved: {date_app}")
            with col_del:
                if st.button("🗑️", key=f"del_ap_{real_i}"):
                    safe_delete(approved_sheet, real_i)
                    st.rerun()
            st.divider()

# ══════════════════════════════════════════════
# TAB 3 — غير متوفر
# ══════════════════════════════════════════════
with tab3:
    st.subheader("❌ غير متوفر | Unavailable")
    data_un = get_cached(unavailable_sheet)

    if len(data_un) <= 1:
        st.info("لا يوجد | Nothing unavailable yet.")
    else:
        rows_un = data_un[1:]
        search_un = st.text_input("🔍 بحث SKU | Search SKU", key="srch_un", placeholder="اكتب SKU...")
        filtered_un = [r for r in rows_un if not search_un or search_un.strip().upper() in r[0].upper()]

        df_un = pd.DataFrame(rows_un, columns=data_un[0])
        c1, c2 = st.columns(2)
        with c1: dl_btn(df_un, "unavailable")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_un", use_container_width=True):
                st.session_state["confirm_clear_un"] = True
        confirm_clear("clear_un", unavailable_sheet, "غير المتوفر | Unavailable")

        st.write(f"**عرض | Showing: {len(filtered_un)} / {len(rows_un)}**")
        for row in filtered_un:
            real_i = rows_un.index(row) + 2
            while len(row) < 5: row.append("")
            sku, qty, img, date_add, date_marked = row[0], row[1], row[2], row[3], row[4]
            col_img, col_info, col_del = st.columns([1, 5, 1])
            with col_img:
                show_img(img, 70)
            with col_info:
                st.markdown(f"**SKU:** `{sku}`")
                show_sku_inv(sku, inv_map, excluded_wh)
                st.markdown(f"**Qty طلب | Requested:** {qty}")
                st.caption(f"📅 Requested: {date_add} | ❌ Unavailable: {date_marked}")
            with col_del:
                if st.button("🗑️", key=f"del_un_{real_i}"):
                    safe_delete(unavailable_sheet, real_i)
                    st.rerun()
            st.divider()

# ══════════════════════════════════════════════
# TAB 4 — الجدولة
# ══════════════════════════════════════════════
with tab4:
    st.subheader("📅 الجدولة | Scheduled Items")
    links_map = get_links_map()

    col_tmpl2, _ = st.columns([1, 3])
    with col_tmpl2:
        st.download_button("⬇️ Template الجدولة | Schedule Template",
            data=make_empty_template(["ASN", "SKU", "qty", "تاريخ الجدولة"]),
            file_name=f"schedule_template_{file_timestamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

    uploaded_sched = st.file_uploader("ارفع ملف الجدولة | Upload Schedule File",
        type=["xlsx","xls","csv"], key="sched_upload")

    if uploaded_sched:
        try:
            df_sched = pd.read_csv(uploaded_sched, dtype=str).fillna("") if uploaded_sched.name.endswith(".csv") \
                else pd.read_excel(uploaded_sched, dtype=str).fillna("")
            col_map = {}
            for c in df_sched.columns:
                cl = c.strip().lower()
                if cl == "asn": col_map["asn"] = c
                if cl in ("sku","item nr","item_nr"): col_map["sku"] = c
                if cl in ("qty","quantity","كمية"): col_map["qty"] = c
                if "جدول" in cl or "schedule" in cl or "date" in cl: col_map["date"] = c
            asn_col  = col_map.get("asn",  df_sched.columns[0] if len(df_sched.columns) > 0 else None)
            sku_col  = col_map.get("sku",  df_sched.columns[1] if len(df_sched.columns) > 1 else None)
            qty_col  = col_map.get("qty",  df_sched.columns[2] if len(df_sched.columns) > 2 else None)
            date_col = col_map.get("date", df_sched.columns[3] if len(df_sched.columns) > 3 else None)

            st.info(f"📊 {len(df_sched)} صف | rows")
            st.dataframe(df_sched, use_container_width=True, height=150)

            if st.button("📤 إضافة الجدولة | Add Schedule", type="primary"):
                existing = get_cached(scheduled_sheet, force=True)
                existing_asns = {r[0].strip() for r in existing[1:] if r} if len(existing) > 1 else set()
                date_now = now_str()
                to_add = []
                skipped = 0
                for _, row in df_sched.iterrows():
                    asn = str(row[asn_col]).strip() if asn_col else ""
                    sku = str(row[sku_col]).strip() if sku_col else ""
                    qty = str(row[qty_col]).strip() if qty_col else ""
                    date_val = str(row[date_col]).strip() if date_col else ""
                    img = links_map.get(sku.upper(), "")
                    parsed = parse_excel_date(date_val)
                    date_str = parsed.strftime("%Y-%m-%d") if parsed else date_val
                    if asn and asn.lower() != "nan":
                        if asn in existing_asns:
                            skipped += 1
                        else:
                            to_add.append([asn, sku, qty, date_str, img, date_now])
                            existing_asns.add(asn)
                if to_add:
                    safe_batch_append(scheduled_sheet, to_add)
                msg = f"✅ أُضيف | Added: {len(to_add)}"
                if skipped: msg += f" | ⚠️ مكرر | Duplicates: {skipped}"
                st.success(msg)
                st.rerun()
        except Exception as e:
            st.error(f"❌ {e}")

    st.divider()
    st.subheader("📋 الجدولة الحالية | Current Schedule")
    data_sc = get_cached(scheduled_sheet)

    if len(data_sc) <= 1:
        st.info("لا توجد جدولة | No scheduled items.")
    else:
        rows_sc = data_sc[1:]
        df_sc = pd.DataFrame(rows_sc, columns=data_sc[0])
        c1, c2 = st.columns(2)
        with c1: dl_btn(df_sc, "scheduled")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_sc", use_container_width=True):
                st.session_state["confirm_clear_sc"] = True
        confirm_clear("clear_sc", scheduled_sheet, "الجدولة | Schedule")

        today = datetime.now().date()
        st.write(f"**الإجمالي | Total: {len(rows_sc)}**")

        for i, row in enumerate(rows_sc, start=2):
            while len(row) < 6: row.append("")
            asn, sku, qty, sched_date, img, date_added = row[0], row[1], row[2], row[3], row[4], row[5]
            parsed_date = parse_excel_date(sched_date)
            is_expired = parsed_date and (today > parsed_date.date() + timedelta(days=1))

            inv_info = inv_map.get(sku.upper(), {})
            monthly_sales = inv_info.get("sales", 0)
            is_alert = False
            try:
                if monthly_sales > 0 and _to_int(qty) > monthly_sales:
                    is_alert = True
            except: pass

            # شريط أحمر للتنبيه
            border = "#ef4444" if is_alert else ("#f59e0b" if is_expired else "#3b82f6")
            bg     = "#2d1515" if is_alert else ("#2d2000" if is_expired else "#0f172a")

            st.markdown(f'<div style="border-left:5px solid {border}; background:{bg}; border-radius:8px; padding:2px 0 2px 8px; margin-bottom:2px;"></div>', unsafe_allow_html=True)

            col_img, col_info, col_del = st.columns([1, 5, 1])
            with col_img:
                show_img(img, 70)
            with col_info:
                st.markdown(f"**ASN:** `{asn}` | **SKU:** `{sku}`")
                show_sku_inv(sku, inv_map, excluded_wh)
                st.markdown(f"**Qty جدولة | Scheduled:** {qty}")
                if is_alert:
                    st.markdown(f"🔴 **تنبيه | Alert:** الكمية ({qty}) > المبيع الشهري ({monthly_sales})")
                status = "⚠️ منتهي | Expired" if is_expired else "✅ ساري | Active"
                st.markdown(f"📅 **تاريخ الجدولة | Schedule Date: {sched_date}** — {status}")
            with col_del:
                if st.button("🗑️", key=f"del_sc_{i}"):
                    safe_delete(scheduled_sheet, i)
                    st.rerun()
            st.divider()

# ══════════════════════════════════════════════
# TAB 5 — تنبيهات
# ══════════════════════════════════════════════
with tab5:
    st.subheader("⚠️ تنبيهات الجدولة | Schedule Alerts")
    st.caption("SKU الكمية المجدولة أعلى من المبيع الشهري | Scheduled qty > Monthly sales")

    data_sc2 = get_cached(scheduled_sheet)
    alerts = []
    if len(data_sc2) > 1:
        for row in data_sc2[1:]:
            while len(row) < 6: row.append("")
            asn, sku, qty, sched_date, img = row[0], row[1], row[2], row[3], row[4]
            info = inv_map.get(sku.upper(), {})
            monthly_sales = info.get("sales", 0)
            total_stock   = info.get("total_stock", 0)
            try:
                if monthly_sales > 0 and _to_int(qty) > monthly_sales:
                    alerts.append((asn, sku, qty, monthly_sales, total_stock, sched_date, img, info))
            except: pass

    if not inv_map:
        st.info("ارفع ملف المخزون أولاً من تاب المخزون | Upload Inventory file first")
    elif not alerts:
        st.success("✅ لا توجد تنبيهات | No alerts")
    else:
        df_alerts = pd.DataFrame(
            [(a[0],a[1],a[2],a[3],a[4],a[5]) for a in alerts],
            columns=["ASN","SKU","Scheduled Qty","Monthly Sales","Total Stock","Schedule Date"]
        )
        c1, c2 = st.columns(2)
        with c1: dl_btn(df_alerts, "alerts")
        with c2: st.error(f"⚠️ عدد التنبيهات | Alerts: {len(alerts)}")

        for asn, sku, qty, monthly_sales, stock, sched_date, img, info in alerts:
            col_img, col_info = st.columns([1, 6])
            with col_img:
                show_img(img, 70)
            with col_info:
                st.markdown(f"**ASN:** `{asn}` | **SKU:** `{sku}`")
                show_sku_inv(sku, inv_map, excluded_wh)
                st.markdown(f"🔴 **الكمية المجدولة | Scheduled:** {qty} > **المبيع الشهري | Monthly Sales:** {monthly_sales}")
                st.caption(f"📅 تاريخ الجدولة | Schedule Date: {sched_date}")
            st.divider()

# ══════════════════════════════════════════════
# TAB 6 — المخزون
# ══════════════════════════════════════════════
with tab6:
    st.subheader("📊 المخزون والمبيع الشهري | Inventory & Monthly Sales")
    links_map = get_links_map()

    col_t1, _ = st.columns([1, 3])
    with col_t1:
        st.download_button("⬇️ Template المخزون | Inventory Template",
            data=make_empty_template(["warehouse_code","CURRENT STOCK.QTY","sku","مبيع شهر جدول.QTY"]),
            file_name=f"inventory_template_{file_timestamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

    uploaded_inv = st.file_uploader("ارفع ملف المخزون | Upload Inventory File",
        type=["xlsx","xls","xlsm","csv"], key="inv_upload")

    if uploaded_inv:
        try:
            df_inv = pd.read_csv(uploaded_inv, dtype=str).fillna("") if uploaded_inv.name.endswith(".csv") \
                else pd.read_excel(uploaded_inv, dtype=str).fillna("")

            wh_col = sku_col2 = stock_col = sales_col = None
            for c in df_inv.columns:
                cl = c.strip().lower()
                if "warehouse" in cl: wh_col = c
                if cl in ("sku","item nr","item_nr"): sku_col2 = c
                if "current stock" in cl or cl == "stock": stock_col = c
                if "مبيع" in cl or "sales" in cl: sales_col = c
                if "qty" in cl and sales_col is None: sales_col = c

            if not sku_col2:  sku_col2  = df_inv.columns[2] if len(df_inv.columns) > 2 else df_inv.columns[0]
            if not wh_col:    wh_col    = df_inv.columns[0] if len(df_inv.columns) > 0 else None
            if not stock_col: stock_col = df_inv.columns[1] if len(df_inv.columns) > 1 else None
            if not sales_col: sales_col = df_inv.columns[3] if len(df_inv.columns) > 3 else None

            st.info(f"📊 {len(df_inv)} صف | Warehouse:`{wh_col}` | SKU:`{sku_col2}` | Stock:`{stock_col}` | Sales:`{sales_col}`")
            st.dataframe(df_inv.head(10), use_container_width=True, height=180)

            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("📤 إضافة للموجود | Append", type="primary", use_container_width=True):
                    date_now = now_str()
                    to_add = []
                    for _, row in df_inv.iterrows():
                        wh    = str(row[wh_col]).strip()    if wh_col    else ""
                        sku   = str(row[sku_col2]).strip()  if sku_col2  else ""
                        stock = str(row[stock_col]).strip() if stock_col else ""
                        sales = str(row[sales_col]).strip() if sales_col else ""
                        img   = links_map.get(sku.upper(), "")
                        if sku and sku.lower() != "nan":
                            to_add.append([sku, wh, stock, sales, img, date_now])
                    if safe_batch_append(inventory_sheet, to_add):
                        st.success(f"✅ أُضيف {len(to_add)} صف")
                        st.rerun()
            with col_b:
                if st.button("🔄 استبدال الكل | Replace All", type="secondary", use_container_width=True):
                    st.session_state["confirm_replace_inv"] = True

            if st.session_state.get("confirm_replace_inv"):
                st.warning("⚠️ هيمسح الكل ويرفع الجديد؟")
                cy, cn = st.columns(2)
                if cy.button("✅ نعم", key="yes_replace_inv"):
                    safe_delete_all(inventory_sheet)
                    date_now = now_str()
                    to_add = []
                    for _, row in df_inv.iterrows():
                        wh    = str(row[wh_col]).strip()    if wh_col    else ""
                        sku   = str(row[sku_col2]).strip()  if sku_col2  else ""
                        stock = str(row[stock_col]).strip() if stock_col else ""
                        sales = str(row[sales_col]).strip() if sales_col else ""
                        img   = links_map.get(sku.upper(), "")
                        if sku and sku.lower() != "nan":
                            to_add.append([sku, wh, stock, sales, img, date_now])
                    safe_batch_append(inventory_sheet, to_add)
                    st.session_state["confirm_replace_inv"] = False
                    st.success(f"✅ تم الاستبدال — {len(to_add)} صف")
                    st.rerun()
                if cn.button("❌ لا", key="no_replace_inv"):
                    st.session_state["confirm_replace_inv"] = False
                    st.rerun()
        except Exception as e:
            st.error(f"❌ {e}")

    st.divider()
    st.subheader("📋 بيانات المخزون الحالية | Current Inventory")

    if not inv_map:
        st.info("لم يُرفع ملف مخزون بعد | No inventory uploaded yet.")
    else:
        search_inv = st.text_input("🔍 بحث SKU | Search SKU", key="srch_inv", placeholder="اكتب SKU...")
        df_inv_all = pd.DataFrame(
            get_cached(inventory_sheet)[1:],
            columns=get_cached(inventory_sheet)[0]
        )
        c1, c2 = st.columns(2)
        with c1: dl_btn(df_inv_all, "inventory")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_inv", use_container_width=True):
                st.session_state["confirm_clear_inv"] = True
        confirm_clear("clear_inv", inventory_sheet, "المخزون | Inventory")

        if excluded_wh:
            st.info(f"⚙️ مستودعات مستثناة من الحساب | Excluded: {', '.join(excluded_wh)}")

        filtered_inv = {k: v for k, v in inv_map.items()
                        if not search_inv or search_inv.strip().upper() in k.upper()}

        st.write(f"**SKUs: {len(filtered_inv)}**")

        for sku_key, info in filtered_inv.items():
            col_img, col_info = st.columns([1, 6])
            with col_img:
                show_img(info["img"], 70)
            with col_info:
                st.markdown(f"**SKU:** `{info['sku']}`")
                st.markdown(
                    f"📦 **إجمالي المخزون (المحسوب) | Stock:** **{info['total_stock']}** &nbsp;|&nbsp; "
                    f"📈 **مبيع شهري | Monthly Sales:** **{info['sales']}**"
                )
                # تفصيل المستودعات
                wh_badges = []
                for wh, stk in info["warehouses"].items():
                    color = "#6b2222" if wh.upper() in excluded_wh else "#1e3a5f"
                    txt_color = "#fca5a5" if wh.upper() in excluded_wh else "#93c5fd"
                    strike = "text-decoration:line-through;" if wh.upper() in excluded_wh else ""
                    wh_badges.append(
                        f'<span style="display:inline-block;background:{color};border-radius:6px;'
                        f'padding:2px 8px;margin:2px;font-size:12px;color:{txt_color};{strike}">'
                        f'{wh}: {stk}</span>'
                    )
                st.markdown("🏭 " + " ".join(wh_badges), unsafe_allow_html=True)
                st.caption(f"📅 {info['date']}")
            st.divider()

# ══════════════════════════════════════════════
# TAB 7 — منتهية الصلاحية
# ══════════════════════════════════════════════
with tab7:
    st.subheader("🗂️ الجدولة منتهية الصلاحية | Expired Schedule")
    data_ex = get_cached(expired_sheet)

    if len(data_ex) <= 1:
        st.info("لا يوجد منتهي | No expired items.")
    else:
        rows_ex = data_ex[1:]
        df_ex = pd.DataFrame(rows_ex, columns=data_ex[0])
        c1, c2 = st.columns(2)
        with c1: dl_btn(df_ex, "expired")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_ex", use_container_width=True):
                st.session_state["confirm_clear_ex"] = True
        confirm_clear("clear_ex", expired_sheet, "المنتهية | Expired")

        st.write(f"**الإجمالي | Total: {len(rows_ex)}**")
        for i, row in enumerate(rows_ex, start=2):
            while len(row) < 7: row.append("")
            asn, sku, qty, sched_date, img, date_added, date_exp = row[0], row[1], row[2], row[3], row[4], row[5], row[6]
            col_img, col_info, col_del = st.columns([1, 5, 1])
            with col_img:
                show_img(img, 70)
            with col_info:
                st.markdown(f"**ASN:** `{asn}` | **SKU:** `{sku}`")
                show_sku_inv(sku, inv_map, excluded_wh)
                st.markdown(f"**Quantity:** {qty}")
                st.caption(f"📅 Schedule: {sched_date} | 🗂️ Expired: {date_exp}")
            with col_del:
                if st.button("🗑️", key=f"del_ex_{i}"):
                    safe_delete(expired_sheet, i)
                    st.rerun()
            st.divider()

# ══════════════════════════════════════════════
# TAB 8 — الإعدادات
# ══════════════════════════════════════════════
with tab8:
    st.subheader("⚙️ الإعدادات | Settings")
    st.caption("الإعدادات محفوظة في جوجل شيت وتبقى حتى لو قفلت الموقع | Settings are saved in Google Sheets")

    current_settings = load_settings()

    st.markdown("### 🏭 المستودعات المستثناة من حساب المخزون | Excluded Warehouses")
    st.caption("اكتب أسماء المستودعات مفصولة بفاصلة — هتتلغى من حساب الإجمالي | Comma-separated warehouse names to exclude from total stock")

    # نجيب كل المستودعات الموجودة
    all_warehouses = sorted(set(
        r[1].strip() for r in get_cached(inventory_sheet)[1:] if len(r) > 1 and r[1].strip()
    ))

    current_excluded_str = current_settings.get("excluded_warehouses", "")
    current_excluded_list = [w.strip() for w in current_excluded_str.split(",") if w.strip()]

    if all_warehouses:
        st.write("**المستودعات المتاحة | Available Warehouses:**")
        selected_excluded = st.multiselect(
            "اختر المستودعات المستثناة | Select excluded warehouses:",
            options=all_warehouses,
            default=[w for w in current_excluded_list if w in all_warehouses],
            key="wh_multiselect"
        )
    else:
        st.info("ارفع ملف المخزون أولاً لتظهر المستودعات | Upload inventory file to see warehouses")
        selected_excluded = []
        manual_excluded = st.text_input(
            "أو اكتب يدوياً | Or type manually (comma-separated):",
            value=current_excluded_str,
            key="wh_manual"
        )
        if manual_excluded != current_excluded_str:
            selected_excluded = [w.strip() for w in manual_excluded.split(",") if w.strip()]

    if st.button("💾 حفظ الإعدادات | Save Settings", type="primary"):
        new_val = ",".join(selected_excluded) if selected_excluded else ""
        save_setting("excluded_warehouses", new_val)
        st.success("✅ تم الحفظ | Settings saved!")
        st.rerun()

    st.divider()

    # عرض الإعدادات الحالية
    st.markdown("### 📋 الإعدادات الحالية | Current Settings")
    if excluded_wh:
        st.warning(f"🚫 مستودعات مستثناة الآن | Currently excluded: **{', '.join(excluded_wh)}**")
    else:
        st.success("✅ لا توجد مستودعات مستثناة — كل المخزون محسوب | All warehouses included")
