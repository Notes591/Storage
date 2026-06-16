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
st.set_page_config(page_title="📦 طلبات المخزون", page_icon="📦", layout="wide")

# ====== الاتصال بجوجل شيت ======
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds_dict = st.secrets["gcp_service_account"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# ====== فتح أو إنشاء الملف ======
SHEET_NAME = "Complaints"

try:
    ss = client.open(SHEET_NAME)
except gspread.exceptions.SpreadsheetNotFound:
    ss = client.create(SHEET_NAME)
    # شارك الملف مع الإيميل اللي في secrets لو موجود
    try:
        share_email = st.secrets.get("share_email", None)
        if share_email:
            ss.share(share_email, perm_type='user', role='writer')
    except Exception:
        pass

# ====== أوراق الشيت ======
SHEET_TABS = ["Requests", "Approved", "Unavailable"]
HEADERS = {
    "Requests":    ["SKU", "Quantity", "Image URL", "Date Added", "File Name"],
    "Approved":    ["SKU", "Quantity Requested", "Quantity Approved", "Image URL", "Date Added", "Date Approved"],
    "Unavailable": ["SKU", "Quantity", "Image URL", "Date Added", "Date Marked Unavailable"],
}

sheets = {}
for tab in SHEET_TABS:
    try:
        ws = ss.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab, rows="2000", cols="10")
        ws.append_row(HEADERS[tab])
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

def safe_update_cell(sheet, row_idx, col_idx, value, retries=4, delay=1):
    for _ in range(retries):
        try:
            sheet.update_cell(row_idx, col_idx, value)
            clear_cache(sheet)
            return True
        except Exception:
            time.sleep(delay)
    return False

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def to_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()

# ====== CSS ======
st.markdown("""
<style>
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] { 
        background: #1e293b; color: white; border-radius: 8px;
        padding: 8px 20px; font-weight: bold;
    }
    .stTabs [aria-selected="true"] { background: #3b82f6 !important; }
    .sku-card {
        background: #1e293b; border-radius: 10px;
        padding: 14px 18px; margin-bottom: 10px;
        border-left: 4px solid #3b82f6;
    }
    .approved-card { border-left-color: #22c55e !important; }
    .unavail-card  { border-left-color: #ef4444 !important; }
    .img-thumb { border-radius: 8px; max-height: 80px; }
</style>
""", unsafe_allow_html=True)

st.title("📦 نظام طلبات المخزون")
st.caption("أضف طلبات البضاعة المطلوبة — المستودع يوافق أو يرفض")

# ====== التابات ======
tab1, tab2, tab3 = st.tabs(["📋 الطلبات", "✅ الموافقة", "❌ غير متوفر"])

# ================================================
# TAB 1 — الطلبات
# ================================================
with tab1:
    st.subheader("➕ إضافة طلبات جديدة")

    method = st.radio("طريقة الإضافة:", ["رفع ملف Excel/CSV", "لصق بيانات مباشرة"], horizontal=True)

    added_rows = []
    file_name_label = ""

    if method == "رفع ملف Excel/CSV":
        uploaded = st.file_uploader("ارفع ملف Excel أو CSV", type=["xlsx", "xls", "csv"])
        if uploaded:
            file_name_label = uploaded.name
            try:
                if uploaded.name.endswith(".csv"):
                    df_up = pd.read_csv(uploaded, dtype=str).fillna("")
                else:
                    df_up = pd.read_excel(uploaded, dtype=str).fillna("")

                # نحاول نطابق عمود SKU وعمود Quantity
                cols = [c.strip().lower() for c in df_up.columns]
                sku_col = qty_col = img_col = None
                for c in df_up.columns:
                    cl = c.strip().lower()
                    if cl in ("sku", "item", "product", "item nr", "item_nr"):
                        sku_col = c
                    if cl in ("quantity", "qty", "كمية", "الكمية", "amount"):
                        qty_col = c
                    if cl in ("image", "image url", "img", "صورة", "link"):
                        img_col = c

                if not sku_col:
                    sku_col = df_up.columns[0]
                if not qty_col and len(df_up.columns) > 1:
                    qty_col = df_up.columns[1]
                if not img_col and len(df_up.columns) > 2:
                    img_col = df_up.columns[2]

                st.write(f"📊 {len(df_up)} صف — عمود SKU: `{sku_col}` | Quantity: `{qty_col}`")
                st.dataframe(df_up[[c for c in [sku_col, qty_col, img_col] if c]], use_container_width=True, height=200)

                for _, row in df_up.iterrows():
                    sku = str(row[sku_col]).strip() if sku_col else ""
                    qty = str(row[qty_col]).strip() if qty_col else ""
                    img = str(row[img_col]).strip() if img_col else ""
                    if sku:
                        added_rows.append((sku, qty, img))

            except Exception as e:
                st.error(f"❌ خطأ في قراءة الملف: {e}")

    else:
        st.info("الصق البيانات بالشكل: **SKU , Quantity , Image URL** (كل صف في سطر)")
        pasted = st.text_area("الصق هنا:", height=150, placeholder="SKU001 , 5 , https://...\nSKU002 , 3 , https://...")
        file_name_label = "Manual Entry"
        if pasted.strip():
            for line in pasted.strip().splitlines():
                parts = [p.strip() for p in line.split(",")]
                sku = parts[0] if len(parts) > 0 else ""
                qty = parts[1] if len(parts) > 1 else ""
                img = parts[2] if len(parts) > 2 else ""
                if sku:
                    added_rows.append((sku, qty, img))
            if added_rows:
                st.write(f"✅ {len(added_rows)} صف جاهز للإضافة")

    if added_rows:
        if st.button("📤 إضافة إلى الطلبات", type="primary"):
            date_now = now_str()
            success_count = 0
            with st.spinner("جاري الإضافة..."):
                for sku, qty, img in added_rows:
                    if safe_append(requests_sheet, [sku, qty, img, date_now, file_name_label]):
                        success_count += 1
            if success_count:
                st.success(f"✅ تمت إضافة {success_count} صف بتاريخ {date_now}")
                st.rerun()
            else:
                st.error("❌ فشلت الإضافة، حاول تاني")

    st.divider()

    # ====== عرض الطلبات ======
    st.subheader("📋 الطلبات الحالية")
    data = get_cached(requests_sheet)

    if len(data) <= 1:
        st.info("لا توجد طلبات حالياً.")
    else:
        headers = data[0]
        rows    = data[1:]

        # زر تحميل Excel
        df_req = pd.DataFrame(rows, columns=headers)
        st.download_button(
            "⬇️ تحميل Excel",
            data=to_excel(df_req),
            file_name=f"requests_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # زر مسح الملف كله
        if st.button("🗑️ مسح كل الطلبات (لليوم الجديد)", type="secondary"):
            st.session_state["confirm_clear"] = True

        if st.session_state.get("confirm_clear"):
            st.warning("⚠️ هتحذف كل الطلبات، متأكد؟")
            col_yes, col_no = st.columns(2)
            if col_yes.button("✅ نعم، امسح"):
                # نمسح كل الصفوف بعد الهيدر
                try:
                    requests_sheet.delete_rows(2, len(rows) + 1)
                    clear_cache(requests_sheet)
                    st.session_state["confirm_clear"] = False
                    st.success("✅ تم مسح كل الطلبات")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ {e}")
            if col_no.button("❌ لا، رجوع"):
                st.session_state["confirm_clear"] = False
                st.rerun()

        st.write(f"**إجمالي الطلبات: {len(rows)}**")

        for i, row in enumerate(rows, start=2):
            while len(row) < 5:
                row.append("")
            sku, qty, img, date_added, fname = row[0], row[1], row[2], row[3], row[4]

            with st.container():
                col_img, col_info, col_actions = st.columns([1, 4, 3])

                with col_img:
                    if img and img.startswith("http"):
                        st.image(img, width=80)
                    else:
                        st.write("🖼️")

                with col_info:
                    st.markdown(f"**SKU:** `{sku}`")
                    st.markdown(f"**Quantity:** {qty}")
                    st.caption(f"📅 {date_added} | 📁 {fname}")

                with col_actions:
                    c1, c2, c3 = st.columns(3)

                    # موافقة
                    with c1:
                        with st.popover("✅ وافق"):
                            new_qty = st.text_input("الكمية الموافق عليها", value=qty, key=f"aqty_{i}")
                            if st.button("تأكيد الموافقة", key=f"aconf_{i}"):
                                if safe_append(approved_sheet, [sku, qty, new_qty, img, date_added, now_str()]):
                                    if safe_delete(requests_sheet, i):
                                        st.success("✅ تمت الموافقة")
                                        st.rerun()

                    # غير متوفر
                    with c2:
                        if st.button("❌ مش موجود", key=f"unavail_{i}"):
                            if safe_append(unavailable_sheet, [sku, qty, img, date_added, now_str()]):
                                if safe_delete(requests_sheet, i):
                                    st.rerun()

                    # حذف
                    with c3:
                        if st.button("🗑️", key=f"del_req_{i}", help="حذف الطلب"):
                            if safe_delete(requests_sheet, i):
                                st.rerun()

                st.divider()

# ================================================
# TAB 2 — الموافقة
# ================================================
with tab2:
    st.subheader("✅ الطلبات الموافق عليها")
    data_ap = get_cached(approved_sheet)

    if len(data_ap) <= 1:
        st.info("لا توجد موافقات حتى الآن.")
    else:
        headers_ap = data_ap[0]
        rows_ap    = data_ap[1:]

        df_ap = pd.DataFrame(rows_ap, columns=headers_ap)
        st.download_button(
            "⬇️ تحميل Excel",
            data=to_excel(df_ap),
            file_name=f"approved_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.write(f"**إجمالي: {len(rows_ap)}**")

        for i, row in enumerate(rows_ap, start=2):
            while len(row) < 6:
                row.append("")
            sku, qty_req, qty_app, img, date_add, date_app = row[0], row[1], row[2], row[3], row[4], row[5]

            col_img, col_info, col_del = st.columns([1, 5, 1])

            with col_img:
                if img and img.startswith("http"):
                    st.image(img, width=70)
                else:
                    st.write("🖼️")

            with col_info:
                st.markdown(f"**SKU:** `{sku}`")
                qty_text = f"**Quantity طلبت:** {qty_req}"
                if qty_app and qty_app != qty_req:
                    qty_text += f" ← **وافقوا على:** {qty_app} ⚠️"
                else:
                    qty_text += f" ← **موافق:** {qty_app}"
                st.markdown(qty_text)
                st.caption(f"📅 طُلب: {date_add} | ✅ وُفِق: {date_app}")

            with col_del:
                if st.button("🗑️", key=f"del_ap_{i}", help="حذف"):
                    if safe_delete(approved_sheet, i):
                        st.rerun()

            st.divider()

# ================================================
# TAB 3 — غير متوفر
# ================================================
with tab3:
    st.subheader("❌ غير متوفر في المستودع")
    data_un = get_cached(unavailable_sheet)

    if len(data_un) <= 1:
        st.info("لا يوجد شيء غير متوفر حتى الآن.")
    else:
        headers_un = data_un[0]
        rows_un    = data_un[1:]

        df_un = pd.DataFrame(rows_un, columns=headers_un)
        st.download_button(
            "⬇️ تحميل Excel",
            data=to_excel(df_un),
            file_name=f"unavailable_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.write(f"**إجمالي: {len(rows_un)}**")

        for i, row in enumerate(rows_un, start=2):
            while len(row) < 5:
                row.append("")
            sku, qty, img, date_add, date_marked = row[0], row[1], row[2], row[3], row[4]

            col_img, col_info, col_del = st.columns([1, 5, 1])

            with col_img:
                if img and img.startswith("http"):
                    st.image(img, width=70)
                else:
                    st.write("🖼️")

            with col_info:
                st.markdown(f"**SKU:** `{sku}` | **Quantity:** {qty}")
                st.caption(f"📅 طُلب: {date_add} | ❌ رُفض: {date_marked}")

            with col_del:
                if st.button("🗑️", key=f"del_un_{i}", help="حذف"):
                    if safe_delete(unavailable_sheet, i):
                        st.rerun()

            st.divider()
