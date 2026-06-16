# -*- coding: utf-8 -*-
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import time
import pandas as pd
import io
import gspread.exceptions

# ====== إعدادات الصفحة ======
st.set_page_config(page_title="📦 Stock Requests | طلبات المخزون", page_icon="📦", layout="wide")

# ====== الاتصال بجوجل شيت ======
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds_dict = st.secrets["gcp_service_account"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# ====== فتح شيت Complaints (نفس شيت Complain.py) ======
COMPLAINTS_SHEET = "Complaints"
ss = client.open(COMPLAINTS_SHEET)

# ====== ورقة links (SKU → Image URL) ======
try:
    links_ws = ss.worksheet("links n")
except gspread.exceptions.WorksheetNotFound:
    links_ws = ss.add_worksheet(title="links n", rows="2000", cols="3")
    links_ws.append_row(["SKU", "Image URL"])

# ====== أوراق Stock Requests ======
STOCK_TABS = ["Requests", "Approved", "Unavailable"]
STOCK_HEADERS = {
    "Requests":    ["SKU", "Quantity", "Image URL", "Date Added", "File Name"],
    "Approved":    ["SKU", "Quantity Requested", "Quantity Approved", "Image URL", "Date Added", "Date Approved"],
    "Unavailable": ["SKU", "Quantity", "Image URL", "Date Added", "Date Marked Unavailable"],
}

sheets = {}
for tab in STOCK_TABS:
    try:
        ws = ss.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab, rows="2000", cols="10")
        ws.append_row(STOCK_HEADERS[tab])
    sheets[tab] = ws

requests_sheet    = sheets["Requests"]
approved_sheet    = sheets["Approved"]
unavailable_sheet = sheets["Unavailable"]

# ====== كاش ======
def get_cached(sheet, force=False):
    key = f"cache_{sheet.title}"
    if force or key not in st.session_state:
        st.session_state[key] = sheet.get_all_values()
    return st.session_state[key]

def clear_cache(sheet):
    key = f"cache_{sheet.title}"
    if key in st.session_state:
        del st.session_state[key]

# ====== جلب خريطة SKU → Image URL من ورقة links ======
@st.cache_data(ttl=300)
def get_links_map():
    data = links_ws.get_all_values()
    mapping = {}
    for row in data[1:]:
        if len(row) >= 2 and row[0].strip():
            mapping[row[0].strip().upper()] = row[1].strip()
    return mapping

# ====== دوال مساعدة ======
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

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def file_timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def to_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()

def make_empty_template():
    """إنشاء ملف Excel فارغ جاهز للرفع"""
    df = pd.DataFrame(columns=["SKU", "Quantity"])
    return to_excel(df)

# ====== CSS ======
st.markdown("""
<style>
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] {
        background: #1e293b; color: white; border-radius: 8px;
        padding: 8px 20px; font-weight: bold;
    }
    .stTabs [aria-selected="true"] { background: #3b82f6 !important; }
    div[data-testid="stHorizontalBlock"] { align-items: center; }
</style>
""", unsafe_allow_html=True)

st.title("📦 Stock Requests | طلبات المخزون")
st.caption("أضف SKU والكمية المطلوبة — الصورة تُجلب أوتوماتيك من ورقة links")

tab1, tab2, tab3 = st.tabs(["📋 الطلبات | Requests", "✅ الموافقة | Approved", "❌ غير متوفر | Unavailable"])

# ================================================
# TAB 1 — الطلبات
# ================================================
with tab1:
    st.subheader("➕ إضافة طلبات | Add Requests")

    # زر تحميل Template فارغ
    col_up, col_dl = st.columns([3, 1])
    with col_dl:
        st.download_button(
            label="⬇️ تحميل Template فارغ",
            data=make_empty_template(),
            file_name=f"stock_template_{file_timestamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col_up:
        method = st.radio(
            "طريقة الإضافة | Add Method:",
            ["📂 رفع ملف Excel/CSV | Upload File", "✏️ لصق بيانات | Paste Data"],
            horizontal=True
        )

    added_rows = []
    file_name_label = ""
    links_map = get_links_map()

    if "Upload" in method:
        uploaded = st.file_uploader("ارفع ملف Excel أو CSV | Upload Excel or CSV", type=["xlsx", "xls", "csv"])
        if uploaded:
            file_name_label = uploaded.name
            try:
                if uploaded.name.endswith(".csv"):
                    df_up = pd.read_csv(uploaded, dtype=str).fillna("")
                else:
                    df_up = pd.read_excel(uploaded, dtype=str).fillna("")

                sku_col = qty_col = None
                for c in df_up.columns:
                    cl = c.strip().lower()
                    if cl in ("sku", "item", "product", "item nr", "item_nr"):
                        sku_col = c
                    if cl in ("quantity", "qty", "كمية", "الكمية", "amount"):
                        qty_col = c

                if not sku_col:
                    sku_col = df_up.columns[0]
                if not qty_col and len(df_up.columns) > 1:
                    qty_col = df_up.columns[1]

                st.info(f"📊 {len(df_up)} صف | rows — SKU: `{sku_col}` | Quantity: `{qty_col}`")
                st.dataframe(df_up[[c for c in [sku_col, qty_col] if c]], use_container_width=True, height=180)

                for _, row in df_up.iterrows():
                    sku = str(row[sku_col]).strip() if sku_col else ""
                    qty = str(row[qty_col]).strip() if qty_col else ""
                    img = links_map.get(sku.upper(), "")
                    if sku and sku.lower() != "nan":
                        added_rows.append((sku, qty, img))

            except Exception as e:
                st.error(f"❌ خطأ في قراءة الملف | File error: {e}")

    else:
        st.info("الصق SKU والكمية — كل صف في سطر، مفصول بفاصلة | Paste SKU,Quantity one per line")
        pasted = st.text_area(
            "الصق هنا | Paste here:",
            height=150,
            placeholder="SKU001 , 5\nSKU002 , 3\nSKU003 , 10"
        )
        file_name_label = "Manual Entry"
        if pasted.strip():
            for line in pasted.strip().splitlines():
                parts = [p.strip() for p in line.split(",")]
                sku = parts[0] if len(parts) > 0 else ""
                qty = parts[1] if len(parts) > 1 else ""
                img = links_map.get(sku.upper(), "")
                if sku:
                    added_rows.append((sku, qty, img))
            if added_rows:
                st.success(f"✅ {len(added_rows)} صف جاهز | rows ready")

    if added_rows:
        if st.button("📤 إضافة إلى الطلبات | Add to Requests", type="primary"):
            date_now = now_str()
            count = 0
            with st.spinner("جاري الإضافة... | Adding..."):
                for sku, qty, img in added_rows:
                    if safe_append(requests_sheet, [sku, qty, img, date_now, file_name_label]):
                        count += 1
            if count:
                st.success(f"✅ تمت إضافة {count} صف | rows added — {date_now}")
                st.rerun()
            else:
                st.error("❌ فشلت الإضافة | Add failed")

    st.divider()

    # ====== عرض الطلبات الحالية ======
    st.subheader("📋 الطلبات الحالية | Current Requests")
    data = get_cached(requests_sheet)

    if len(data) <= 1:
        st.info("لا توجد طلبات حالياً | No requests yet.")
    else:
        rows = data[1:]

        c_dl, c_clear = st.columns(2)
        df_req = pd.DataFrame(rows, columns=data[0])
        with c_dl:
            st.download_button(
                "⬇️ تحميل Excel | Download Excel",
                data=to_excel(df_req),
                file_name=f"requests_{file_timestamp()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with c_clear:
            if st.button("🗑️ مسح كل الطلبات | Clear All", type="secondary", use_container_width=True):
                st.session_state["confirm_clear"] = True

        if st.session_state.get("confirm_clear"):
            st.warning("⚠️ هتحذف كل الطلبات، متأكد؟ | Delete all requests, sure?")
            cy, cn = st.columns(2)
            if cy.button("✅ نعم | Yes"):
                try:
                    requests_sheet.delete_rows(2, len(rows) + 1)
                    clear_cache(requests_sheet)
                    st.session_state["confirm_clear"] = False
                    st.success("✅ تم المسح | Cleared")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ {e}")
            if cn.button("❌ لا | No"):
                st.session_state["confirm_clear"] = False
                st.rerun()

        st.write(f"**الإجمالي | Total: {len(rows)}**")

        for i, row in enumerate(rows, start=2):
            while len(row) < 5:
                row.append("")
            sku, qty, img, date_added, fname = row[0], row[1], row[2], row[3], row[4]

            col_img, col_info, col_actions = st.columns([1, 4, 3])

            with col_img:
                if img and img.startswith("http"):
                    st.image(img, width=80)
                else:
                    st.markdown("🖼️")

            with col_info:
                st.markdown(f"**SKU:** `{sku}`")
                st.markdown(f"**Quantity | الكمية:** {qty}")
                st.caption(f"📅 {date_added} | 📁 {fname}")

            with col_actions:
                c1, c2, c3 = st.columns(3)

                with c1:
                    with st.popover("✅ وافق\nApprove"):
                        new_qty = st.text_input("الكمية الموافقة | Approved Qty", value=qty, key=f"aqty_{i}")
                        if st.button("✅ تأكيد | Confirm", key=f"aconf_{i}"):
                            if safe_append(approved_sheet, [sku, qty, new_qty, img, date_added, now_str()]):
                                if safe_delete(requests_sheet, i):
                                    st.success("✅ تمت الموافقة | Approved")
                                    st.rerun()

                with c2:
                    if st.button("❌ غير متوفر\nUnavailable", key=f"unavail_{i}"):
                        if safe_append(unavailable_sheet, [sku, qty, img, date_added, now_str()]):
                            if safe_delete(requests_sheet, i):
                                st.rerun()

                with c3:
                    if st.button("🗑️ حذف\nDelete", key=f"del_req_{i}"):
                        if safe_delete(requests_sheet, i):
                            st.rerun()

            st.divider()

# ================================================
# TAB 2 — الموافقة
# ================================================
with tab2:
    st.subheader("✅ الطلبات الموافق عليها | Approved Requests")
    data_ap = get_cached(approved_sheet)

    if len(data_ap) <= 1:
        st.info("لا توجد موافقات حتى الآن | No approvals yet.")
    else:
        rows_ap = data_ap[1:]
        df_ap = pd.DataFrame(rows_ap, columns=data_ap[0])
        st.download_button(
            "⬇️ تحميل Excel | Download Excel",
            data=to_excel(df_ap),
            file_name=f"approved_{file_timestamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.write(f"**الإجمالي | Total: {len(rows_ap)}**")

        for i, row in enumerate(rows_ap, start=2):
            while len(row) < 6:
                row.append("")
            sku, qty_req, qty_app, img, date_add, date_app = row[0], row[1], row[2], row[3], row[4], row[5]

            col_img, col_info, col_del = st.columns([1, 5, 1])

            with col_img:
                if img and img.startswith("http"):
                    st.image(img, width=70)
                else:
                    st.markdown("🖼️")

            with col_info:
                st.markdown(f"**SKU:** `{sku}`")
                if qty_app and qty_app != qty_req:
                    st.markdown(f"**طلبت | Requested:** {qty_req} ← **وافقوا على | Approved:** ⚠️ {qty_app}")
                else:
                    st.markdown(f"**الكمية | Quantity:** {qty_app}")
                st.caption(f"📅 طُلب | Requested: {date_add} | ✅ وُفِق | Approved: {date_app}")

            with col_del:
                if st.button("🗑️", key=f"del_ap_{i}", help="حذف | Delete"):
                    if safe_delete(approved_sheet, i):
                        st.rerun()

            st.divider()

# ================================================
# TAB 3 — غير متوفر
# ================================================
with tab3:
    st.subheader("❌ غير متوفر في المستودع | Unavailable")
    data_un = get_cached(unavailable_sheet)

    if len(data_un) <= 1:
        st.info("لا يوجد شيء غير متوفر | Nothing unavailable yet.")
    else:
        rows_un = data_un[1:]
        df_un = pd.DataFrame(rows_un, columns=data_un[0])
        st.download_button(
            "⬇️ تحميل Excel | Download Excel",
            data=to_excel(df_un),
            file_name=f"unavailable_{file_timestamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.write(f"**الإجمالي | Total: {len(rows_un)}**")

        for i, row in enumerate(rows_un, start=2):
            while len(row) < 5:
                row.append("")
            sku, qty, img, date_add, date_marked = row[0], row[1], row[2], row[3], row[4]

            col_img, col_info, col_del = st.columns([1, 5, 1])

            with col_img:
                if img and img.startswith("http"):
                    st.image(img, width=70)
                else:
                    st.markdown("🖼️")

            with col_info:
                st.markdown(f"**SKU:** `{sku}` | **Quantity | الكمية:** {qty}")
                st.caption(f"📅 طُلب | Requested: {date_add} | ❌ غير متوفر | Unavailable: {date_marked}")

            with col_del:
                if st.button("🗑️", key=f"del_un_{i}", help="حذف | Delete"):
                    if safe_delete(unavailable_sheet, i):
                        st.rerun()

            st.divider()
