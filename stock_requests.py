# -*- coding: utf-8 -*-
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
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

# ====== فتح شيت Complaints ======
ss = client.open("Complaints")

# ====== ورقة links n ======
try:
    links_ws = ss.worksheet("links n")
except gspread.exceptions.WorksheetNotFound:
    links_ws = ss.add_worksheet(title="links n", rows="2000", cols="2")
    links_ws.append_row(["SKU", "Image URL"])

# ====== أوراق النظام ======
TABS_CONFIG = {
    "Requests":    ["SKU", "Quantity", "Image URL", "Date Added", "File Name"],
    "Approved":    ["SKU", "Quantity Requested", "Quantity Approved", "Image URL", "Date Added", "Date Approved"],
    "Unavailable": ["SKU", "Quantity", "Image URL", "Date Added", "Date Marked Unavailable"],
    "Scheduled":   ["ASN", "SKU", "Quantity", "Schedule Date", "Image URL", "Date Added"],
    "Expired":     ["ASN", "SKU", "Quantity", "Schedule Date", "Image URL", "Date Added", "Date Expired"],
}

sheets = {}
for tab, headers in TABS_CONFIG.items():
    try:
        ws = ss.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab, rows="2000", cols="12")
        ws.append_row(headers)
    sheets[tab] = ws

requests_sheet    = sheets["Requests"]
approved_sheet    = sheets["Approved"]
unavailable_sheet = sheets["Unavailable"]
scheduled_sheet   = sheets["Scheduled"]
expired_sheet     = sheets["Expired"]

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

# ====== خريطة SKU → Image URL ======
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

def safe_delete_all(sheet):
    try:
        data = sheet.get_all_values()
        if len(data) > 1:
            sheet.delete_rows(2, len(data))
        clear_cache(sheet)
        return True
    except Exception:
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
    """تحويل تاريخ Excel الرقمي أو النصي لـ datetime"""
    try:
        if isinstance(val, (int, float)):
            return datetime(1899, 12, 30) + timedelta(days=int(val))
        return datetime.strptime(str(val).strip(), "%Y-%m-%d")
    except Exception:
        try:
            return datetime.strptime(str(val).strip()[:10], "%Y-%m-%d")
        except Exception:
            return None

# ====== فحص الجدولة المنتهية (بعد التاريخ بيوم) ======
def check_expired_scheduled():
    data = get_cached(scheduled_sheet, force=True)
    if len(data) <= 1:
        return
    today = datetime.now().date()
    expired_indices = []
    for i, row in enumerate(data[1:], start=2):
        while len(row) < 6:
            row.append("")
        sched_date = parse_excel_date(row[3])
        if sched_date and (today > sched_date.date() + timedelta(days=1)):
            expired_indices.append((i, row))

    # نبدأ من الأكبر عشان الحذف ما يؤثرش على الـ index
    for i, row in sorted(expired_indices, reverse=True):
        safe_append(expired_sheet, row + [now_str()])
        safe_delete(scheduled_sheet, i)

# ====== CSS ======
st.markdown("""
<style>
    .stTabs [data-baseweb="tab-list"] { gap: 8px; flex-wrap: wrap; }
    .stTabs [data-baseweb="tab"] {
        background: #1e293b; color: white; border-radius: 8px;
        padding: 8px 16px; font-weight: bold; font-size: 13px;
    }
    .stTabs [aria-selected="true"] { background: #3b82f6 !important; }
    .block-container { padding-top: 1rem; }
</style>
""", unsafe_allow_html=True)

# ====== فحص المنتهية عند كل تحميل ======
if "expired_checked" not in st.session_state:
    check_expired_scheduled()
    st.session_state["expired_checked"] = True

st.title("📦 Stock Requests | طلبات المخزون")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 الطلبات | Requests",
    "✅ الموافقة | Approved",
    "❌ غير متوفر | Unavailable",
    "📅 الجدولة | Scheduled",
    "🗂️ منتهية الصلاحية | Expired",
])

# ════════════════════════════════════════
# TAB 1 — الطلبات | Requests
# ════════════════════════════════════════
with tab1:
    st.subheader("➕ إضافة طلبات | Add Requests")

    col_method, col_tmpl = st.columns([3, 1])
    with col_tmpl:
        st.download_button(
            "⬇️ Template فارغ | Empty Template",
            data=make_empty_template(["SKU", "Quantity"]),
            file_name=f"request_template_{file_timestamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_method:
        method = st.radio(
            "طريقة الإضافة | Add Method:",
            ["📂 رفع ملف | Upload File", "✏️ لصق | Paste"],
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
                df_up = pd.read_csv(uploaded, dtype=str).fillna("") if uploaded.name.endswith(".csv") \
                    else pd.read_excel(uploaded, dtype=str).fillna("")
                sku_col = qty_col = None
                for c in df_up.columns:
                    cl = c.strip().lower()
                    if cl in ("sku", "item", "product", "item nr", "item_nr"):
                        sku_col = c
                    if cl in ("quantity", "qty", "كمية", "الكمية", "amount"):
                        qty_col = c
                if not sku_col: sku_col = df_up.columns[0]
                if not qty_col and len(df_up.columns) > 1: qty_col = df_up.columns[1]

                st.info(f"📊 {len(df_up)} صف | rows")
                st.dataframe(df_up[[c for c in [sku_col, qty_col] if c]], use_container_width=True, height=180)
                for _, row in df_up.iterrows():
                    sku = str(row[sku_col]).strip()
                    qty = str(row[qty_col]).strip() if qty_col else ""
                    img = links_map.get(sku.upper(), "")
                    if sku and sku.lower() != "nan":
                        added_rows.append((sku, qty, img))
            except Exception as e:
                st.error(f"❌ خطأ | Error: {e}")
    else:
        st.info("الصق SKU,Quantity في كل سطر | Paste SKU,Quantity per line")
        pasted = st.text_area("الصق هنا | Paste here:", height=130, placeholder="SKU001 , 5\nSKU002 , 3")
        file_name_label = "Manual Entry"
        if pasted.strip():
            for line in pasted.strip().splitlines():
                parts = [p.strip() for p in line.split(",")]
                sku = parts[0] if parts else ""
                qty = parts[1] if len(parts) > 1 else ""
                img = links_map.get(sku.upper(), "")
                if sku:
                    added_rows.append((sku, qty, img))
            if added_rows:
                st.success(f"✅ {len(added_rows)} صف جاهز | rows ready")

    if added_rows:
        if st.button("📤 إضافة | Add to Requests", type="primary"):
            date_now = now_str()
            count = 0
            with st.spinner("جاري الإضافة... | Adding..."):
                for sku, qty, img in added_rows:
                    if safe_append(requests_sheet, [sku, qty, img, date_now, file_name_label]):
                        count += 1
            st.success(f"✅ تمت إضافة {count} صف | rows added")
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
        with c1:
            st.download_button("⬇️ تحميل Excel | Download", data=to_excel(df_req),
                file_name=f"requests_{file_timestamp()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with c2:
            if st.button("✅ موافقة على الكل | Approve All", use_container_width=True):
                st.session_state["confirm_approve_all"] = True
        with c3:
            if st.button("❌ رفض الكل | Reject All", use_container_width=True):
                st.session_state["confirm_reject_all"] = True
        with c4:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", use_container_width=True):
                st.session_state["confirm_clear_req"] = True

        # تأكيد موافقة الكل
        if st.session_state.get("confirm_approve_all"):
            st.warning("⚠️ موافقة على كل الطلبات؟ | Approve all requests?")
            cy, cn = st.columns(2)
            if cy.button("✅ نعم | Yes", key="yes_approve_all"):
                date_now = now_str()
                for row in rows:
                    while len(row) < 5: row.append("")
                    safe_append(approved_sheet, [row[0], row[1], row[1], row[2], row[3], date_now])
                safe_delete_all(requests_sheet)
                st.session_state["confirm_approve_all"] = False
                st.success("✅ تمت الموافقة على الكل | All approved")
                st.rerun()
            if cn.button("❌ لا | No", key="no_approve_all"):
                st.session_state["confirm_approve_all"] = False
                st.rerun()

        # تأكيد رفض الكل
        if st.session_state.get("confirm_reject_all"):
            st.warning("⚠️ رفض كل الطلبات كغير متوفرة؟ | Reject all as unavailable?")
            cy, cn = st.columns(2)
            if cy.button("✅ نعم | Yes", key="yes_reject_all"):
                date_now = now_str()
                for row in rows:
                    while len(row) < 5: row.append("")
                    safe_append(unavailable_sheet, [row[0], row[1], row[2], row[3], date_now])
                safe_delete_all(requests_sheet)
                st.session_state["confirm_reject_all"] = False
                st.success("❌ تم رفض الكل | All rejected")
                st.rerun()
            if cn.button("❌ لا | No", key="no_reject_all"):
                st.session_state["confirm_reject_all"] = False
                st.rerun()

        # تأكيد مسح الكل
        if st.session_state.get("confirm_clear_req"):
            st.warning("⚠️ مسح كل الطلبات نهائياً؟ | Delete all requests permanently?")
            cy, cn = st.columns(2)
            if cy.button("✅ نعم | Yes", key="yes_clear_req"):
                safe_delete_all(requests_sheet)
                st.session_state["confirm_clear_req"] = False
                st.success("✅ تم المسح | Cleared")
                st.rerun()
            if cn.button("❌ لا | No", key="no_clear_req"):
                st.session_state["confirm_clear_req"] = False
                st.rerun()

        st.write(f"**الإجمالي | Total: {len(rows)}**")

        for i, row in enumerate(rows, start=2):
            while len(row) < 5: row.append("")
            sku, qty, img, date_added, fname = row[0], row[1], row[2], row[3], row[4]

            col_img, col_info, col_actions = st.columns([1, 4, 3])
            with col_img:
                if img and img.startswith("http"):
                    st.image(img, width=75)
                else:
                    st.markdown("🖼️")
            with col_info:
                st.markdown(f"**SKU:** `{sku}`")
                st.markdown(f"**Quantity | الكمية:** {qty}")
                st.caption(f"📅 {date_added} | 📁 {fname}")
            with col_actions:
                ca, cb, cc = st.columns(3)
                with ca:
                    with st.popover("✅ وافق | Approve"):
                        new_qty = st.text_input("الكمية الموافقة | Approved Qty", value=qty, key=f"aqty_{i}")
                        if st.button("✅ تأكيد | Confirm", key=f"aconf_{i}"):
                            if safe_append(approved_sheet, [sku, qty, new_qty, img, date_added, now_str()]):
                                safe_delete(requests_sheet, i)
                                st.rerun()
                with cb:
                    if st.button("❌ غير متوفر\nUnavailable", key=f"unavail_{i}"):
                        if safe_append(unavailable_sheet, [sku, qty, img, date_added, now_str()]):
                            safe_delete(requests_sheet, i)
                            st.rerun()
                with cc:
                    if st.button("🗑️ حذف\nDelete", key=f"del_req_{i}"):
                        safe_delete(requests_sheet, i)
                        st.rerun()
            st.divider()

# ════════════════════════════════════════
# TAB 2 — الموافقة | Approved
# ════════════════════════════════════════
with tab2:
    st.subheader("✅ الطلبات الموافق عليها | Approved Requests")
    data_ap = get_cached(approved_sheet)

    if len(data_ap) <= 1:
        st.info("لا توجد موافقات | No approvals yet.")
    else:
        rows_ap = data_ap[1:]

        # بحث
        search_ap = st.text_input("🔍 بحث بـ SKU | Search by SKU", key="search_approved", placeholder="اكتب SKU...")
        filtered_ap = [r for r in rows_ap if not search_ap or search_ap.strip().upper() in r[0].upper()] if search_ap else rows_ap

        df_ap = pd.DataFrame(rows_ap, columns=data_ap[0])
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇️ تحميل Excel | Download", data=to_excel(df_ap),
                file_name=f"approved_{file_timestamp()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="clear_approved", use_container_width=True):
                st.session_state["confirm_clear_ap"] = True

        if st.session_state.get("confirm_clear_ap"):
            st.warning("⚠️ مسح كل الموافقات؟ | Clear all approvals?")
            cy, cn = st.columns(2)
            if cy.button("✅ نعم | Yes", key="yes_clear_ap"):
                safe_delete_all(approved_sheet)
                st.session_state["confirm_clear_ap"] = False
                st.success("✅ تم المسح | Cleared")
                st.rerun()
            if cn.button("❌ لا | No", key="no_clear_ap"):
                st.session_state["confirm_clear_ap"] = False
                st.rerun()

        st.write(f"**عرض | Showing: {len(filtered_ap)} / {len(rows_ap)}**")

        for i, row in enumerate(filtered_ap):
            real_i = rows_ap.index(row) + 2
            while len(row) < 6: row.append("")
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
                    st.markdown(f"**طلبت | Requested:** {qty_req} → **وافقوا | Approved:** ⚠️ {qty_app}")
                else:
                    st.markdown(f"**Quantity | الكمية:** {qty_app}")
                st.caption(f"📅 طُلب | Requested: {date_add} | ✅ وُفِق | Approved: {date_app}")
            with col_del:
                if st.button("🗑️", key=f"del_ap_{real_i}", help="حذف | Delete"):
                    safe_delete(approved_sheet, real_i)
                    st.rerun()
            st.divider()

# ════════════════════════════════════════
# TAB 3 — غير متوفر | Unavailable
# ════════════════════════════════════════
with tab3:
    st.subheader("❌ غير متوفر في المستودع | Unavailable")
    data_un = get_cached(unavailable_sheet)

    if len(data_un) <= 1:
        st.info("لا يوجد شيء غير متوفر | Nothing unavailable yet.")
    else:
        rows_un = data_un[1:]

        # بحث
        search_un = st.text_input("🔍 بحث بـ SKU | Search by SKU", key="search_unavailable", placeholder="اكتب SKU...")
        filtered_un = [r for r in rows_un if not search_un or search_un.strip().upper() in r[0].upper()] if search_un else rows_un

        df_un = pd.DataFrame(rows_un, columns=data_un[0])
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇️ تحميل Excel | Download", data=to_excel(df_un),
                file_name=f"unavailable_{file_timestamp()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="clear_unavail", use_container_width=True):
                st.session_state["confirm_clear_un"] = True

        if st.session_state.get("confirm_clear_un"):
            st.warning("⚠️ مسح كل غير المتوفر؟ | Clear all unavailable?")
            cy, cn = st.columns(2)
            if cy.button("✅ نعم | Yes", key="yes_clear_un"):
                safe_delete_all(unavailable_sheet)
                st.session_state["confirm_clear_un"] = False
                st.success("✅ تم المسح | Cleared")
                st.rerun()
            if cn.button("❌ لا | No", key="no_clear_un"):
                st.session_state["confirm_clear_un"] = False
                st.rerun()

        st.write(f"**عرض | Showing: {len(filtered_un)} / {len(rows_un)}**")

        for i, row in enumerate(filtered_un):
            real_i = rows_un.index(row) + 2
            while len(row) < 5: row.append("")
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
                if st.button("🗑️", key=f"del_un_{real_i}", help="حذف | Delete"):
                    safe_delete(unavailable_sheet, real_i)
                    st.rerun()
            st.divider()

# ════════════════════════════════════════
# TAB 4 — الجدولة | Scheduled
# ════════════════════════════════════════
with tab4:
    st.subheader("📅 إضافة جدولة | Add Scheduled Items")
    links_map = get_links_map()

    col_tmpl2, _ = st.columns([1, 3])
    with col_tmpl2:
        st.download_button(
            "⬇️ Template الجدولة | Schedule Template",
            data=make_empty_template(["ASN", "SKU", "qty", "تاريخ الجدولة"]),
            file_name=f"schedule_template_{file_timestamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    uploaded_sched = st.file_uploader(
        "ارفع ملف الجدولة | Upload Schedule File",
        type=["xlsx", "xls", "csv"],
        key="sched_upload"
    )

    if uploaded_sched:
        try:
            df_sched = pd.read_csv(uploaded_sched, dtype=str).fillna("") if uploaded_sched.name.endswith(".csv") \
                else pd.read_excel(uploaded_sched, dtype=str).fillna("")

            # تطابق الأعمدة
            col_map = {}
            for c in df_sched.columns:
                cl = c.strip().lower()
                if cl == "asn": col_map["asn"] = c
                if cl in ("sku", "item nr", "item_nr"): col_map["sku"] = c
                if cl in ("qty", "quantity", "كمية"): col_map["qty"] = c
                if "جدول" in cl or "schedule" in cl or "date" in cl: col_map["date"] = c

            asn_col  = col_map.get("asn",  df_sched.columns[0] if len(df_sched.columns) > 0 else None)
            sku_col  = col_map.get("sku",  df_sched.columns[1] if len(df_sched.columns) > 1 else None)
            qty_col  = col_map.get("qty",  df_sched.columns[2] if len(df_sched.columns) > 2 else None)
            date_col = col_map.get("date", df_sched.columns[3] if len(df_sched.columns) > 3 else None)

            st.info(f"📊 {len(df_sched)} صف | rows — ASN: `{asn_col}` | SKU: `{sku_col}` | Qty: `{qty_col}` | Date: `{date_col}`")
            st.dataframe(df_sched, use_container_width=True, height=180)

            if st.button("📤 إضافة الجدولة | Add Schedule", type="primary"):
                # جلب البيانات الموجودة لمنع التكرار
                existing_data = get_cached(scheduled_sheet, force=True)
                existing_asns = {r[0].strip() for r in existing_data[1:] if r} if len(existing_data) > 1 else set()

                date_now = now_str()
                added = skipped = 0
                with st.spinner("جاري الإضافة... | Adding..."):
                    for _, row in df_sched.iterrows():
                        asn  = str(row[asn_col]).strip()  if asn_col  else ""
                        sku  = str(row[sku_col]).strip()  if sku_col  else ""
                        qty  = str(row[qty_col]).strip()  if qty_col  else ""
                        date_val = str(row[date_col]).strip() if date_col else ""
                        img  = links_map.get(sku.upper(), "")

                        # تحويل التاريخ
                        parsed = parse_excel_date(date_val)
                        date_str = parsed.strftime("%Y-%m-%d") if parsed else date_val

                        if asn and asn.lower() != "nan":
                            if asn in existing_asns:
                                skipped += 1
                            else:
                                if safe_append(scheduled_sheet, [asn, sku, qty, date_str, img, date_now]):
                                    existing_asns.add(asn)
                                    added += 1

                msg = f"✅ أُضيف | Added: {added}"
                if skipped: msg += f" | ⚠️ مكرر تجاهلناه | Duplicates skipped: {skipped}"
                st.success(msg)
                st.rerun()

        except Exception as e:
            st.error(f"❌ خطأ | Error: {e}")

    st.divider()
    st.subheader("📋 الجدولة الحالية | Current Schedule")
    data_sc = get_cached(scheduled_sheet)

    if len(data_sc) <= 1:
        st.info("لا توجد جدولة | No scheduled items.")
    else:
        rows_sc = data_sc[1:]
        df_sc = pd.DataFrame(rows_sc, columns=data_sc[0])
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇️ تحميل Excel | Download", data=to_excel(df_sc),
                file_name=f"scheduled_{file_timestamp()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="clear_sched", use_container_width=True):
                st.session_state["confirm_clear_sc"] = True

        if st.session_state.get("confirm_clear_sc"):
            st.warning("⚠️ مسح كل الجدولة؟ | Clear all scheduled?")
            cy, cn = st.columns(2)
            if cy.button("✅ نعم | Yes", key="yes_clear_sc"):
                safe_delete_all(scheduled_sheet)
                st.session_state["confirm_clear_sc"] = False
                st.success("✅ تم المسح | Cleared")
                st.rerun()
            if cn.button("❌ لا | No", key="no_clear_sc"):
                st.session_state["confirm_clear_sc"] = False
                st.rerun()

        today = datetime.now().date()
        st.write(f"**الإجمالي | Total: {len(rows_sc)}**")

        for i, row in enumerate(rows_sc, start=2):
            while len(row) < 6: row.append("")
            asn, sku, qty, sched_date, img, date_added = row[0], row[1], row[2], row[3], row[4], row[5]

            parsed_date = parse_excel_date(sched_date)
            is_expired = parsed_date and (today > parsed_date.date() + timedelta(days=1))
            border_color = "#ef4444" if is_expired else "#3b82f6"

            col_img, col_info, col_del = st.columns([1, 5, 1])
            with col_img:
                if img and img.startswith("http"):
                    st.image(img, width=70)
                else:
                    st.markdown("🖼️")
            with col_info:
                st.markdown(f"**ASN:** `{asn}` | **SKU:** `{sku}`")
                st.markdown(f"**Quantity | الكمية:** {qty}")
                status = "⚠️ منتهي | Expired" if is_expired else "✅ ساري | Active"
                st.markdown(f"📅 **تاريخ الجدولة | Schedule Date: {sched_date}** — {status}")
            with col_del:
                if st.button("🗑️", key=f"del_sc_{i}", help="حذف | Delete"):
                    safe_delete(scheduled_sheet, i)
                    st.rerun()
            st.divider()

# ════════════════════════════════════════
# TAB 5 — منتهية الصلاحية | Expired
# ════════════════════════════════════════
with tab5:
    st.subheader("🗂️ الجدولة منتهية الصلاحية | Expired Scheduled")
    data_ex = get_cached(expired_sheet)

    if len(data_ex) <= 1:
        st.info("لا يوجد منتهي الصلاحية | No expired items.")
    else:
        rows_ex = data_ex[1:]
        df_ex = pd.DataFrame(rows_ex, columns=data_ex[0])
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇️ تحميل Excel | Download", data=to_excel(df_ex),
                file_name=f"expired_{file_timestamp()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="clear_exp", use_container_width=True):
                st.session_state["confirm_clear_ex"] = True

        if st.session_state.get("confirm_clear_ex"):
            st.warning("⚠️ مسح كل منتهيات الصلاحية؟ | Clear all expired?")
            cy, cn = st.columns(2)
            if cy.button("✅ نعم | Yes", key="yes_clear_ex"):
                safe_delete_all(expired_sheet)
                st.session_state["confirm_clear_ex"] = False
                st.success("✅ تم المسح | Cleared")
                st.rerun()
            if cn.button("❌ لا | No", key="no_clear_ex"):
                st.session_state["confirm_clear_ex"] = False
                st.rerun()

        st.write(f"**الإجمالي | Total: {len(rows_ex)}**")

        for i, row in enumerate(rows_ex, start=2):
            while len(row) < 7: row.append("")
            asn, sku, qty, sched_date, img, date_added, date_exp = row[0], row[1], row[2], row[3], row[4], row[5], row[6]

            col_img, col_info, col_del = st.columns([1, 5, 1])
            with col_img:
                if img and img.startswith("http"):
                    st.image(img, width=70)
                else:
                    st.markdown("🖼️")
            with col_info:
                st.markdown(f"**ASN:** `{asn}` | **SKU:** `{sku}`")
                st.markdown(f"**Quantity | الكمية:** {qty}")
                st.caption(f"📅 تاريخ الجدولة | Schedule: {sched_date} | 🗂️ انتهى | Expired: {date_exp}")
            with col_del:
                if st.button("🗑️", key=f"del_ex_{i}", help="حذف | Delete"):
                    safe_delete(expired_sheet, i)
                    st.rerun()
            st.divider()
