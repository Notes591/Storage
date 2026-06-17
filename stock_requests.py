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
ss = client.open("Complaints")

# ══ الأوراق ══
TABS_CONFIG = {
    "Requests":          ["SKU","Quantity","Image URL","Date Added","File Name"],
    "Approved":          ["SKU","Quantity Requested","Quantity Approved","Image URL","Date Added","Date Approved"],
    "Unavailable":       ["SKU","Quantity","Image URL","Date Added","Date Marked Unavailable"],
    "Ordered":           ["SKU","Quantity","Image URL","Date Added","Order Count","Notes"],
    "Scheduled":         ["ASN","SKU","Quantity","Schedule Date","Image URL","Date Added"],
    "CancelledSchedule": ["ASN","SKU","Quantity","Schedule Date","Image URL","Date Added","Cancel Reason","Date Cancelled"],
    "Rescheduled":       ["ASN","SKU","Quantity","Old Schedule Date","Image URL","Date Added","Reschedule Reason","Date Moved"],
    "Expired":           ["ASN","SKU","Quantity","Schedule Date","Image URL","Date Added","Date Expired"],
    "Inventory":         ["SKU","Warehouse","Stock","Monthly Sales","Image URL","Date Uploaded"],
    "Settings":          ["Key","Value"],
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
    links_ws.append_row(["SKU","Image URL"])

requests_sheet    = sheets["Requests"]
approved_sheet    = sheets["Approved"]
unavailable_sheet = sheets["Unavailable"]
ordered_sheet     = sheets["Ordered"]
scheduled_sheet   = sheets["Scheduled"]
cancelled_sheet   = sheets["CancelledSchedule"]
reschedule_sheet  = sheets["Rescheduled"]
expired_sheet     = sheets["Expired"]
inventory_sheet   = sheets["Inventory"]
settings_sheet    = sheets["Settings"]

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
    if not rows_data:
        return True
    for _ in range(retries):
        try:
            sheet.append_rows(rows_data, value_input_option="USER_ENTERED")
            clear_cache(sheet)
            return True
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
    try:
        if isinstance(val,(int,float)):
            return datetime(1899,12,30)+timedelta(days=int(val))
        return datetime.strptime(str(val).strip()[:10],"%Y-%m-%d")
    except:
        return None

def dl_btn(df, prefix, label="⬇️ Excel | Download"):
    st.download_button(label, data=to_excel(df),
        file_name=f"{prefix}_{file_timestamp()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)

def check_expired_scheduled():
    data = get_cached(scheduled_sheet, force=True)
    if len(data) <= 1:
        return
    today = datetime.now().date()
    expired_rows, keep = [], []
    for i, row in enumerate(data[1:], start=2):
        while len(row) < 6: row.append("")
        d = parse_excel_date(row[3])
        if d and today > d.date() + timedelta(days=1):
            expired_rows.append(row + [now_str()])
        else:
            keep.append(i)
    if expired_rows:
        safe_batch_append(expired_sheet, expired_rows)
        del_idx = sorted([x for x in range(2,len(data[1:])+2) if x not in keep], reverse=True)
        for idx in del_idx:
            safe_delete(scheduled_sheet, idx)

# ══ CSS ══
st.markdown("""
<style>
.stTabs [data-baseweb="tab-list"]{gap:5px;flex-wrap:wrap;}
.stTabs [data-baseweb="tab"]{background:#1e293b;color:white;border-radius:8px;padding:6px 12px;font-weight:bold;font-size:11px;}
.stTabs [aria-selected="true"]{background:#3b82f6!important;}
.wh-badge{display:inline-block;border-radius:6px;padding:2px 9px;margin:2px;font-size:12px;}
</style>
""", unsafe_allow_html=True)

# ══ Init ══
if "expired_checked" not in st.session_state:
    check_expired_scheduled()
    st.session_state["expired_checked"] = True

excluded_wh = get_excluded_warehouses()
inv_map     = build_inv_map(excluded_wh)

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

ordinal_map = {1:"الثانية|Second",2:"الثالثة|Third",3:"الرابعة|Fourth",4:"الخامسة|Fifth"}

# ══════════════════════════════════════════════
st.title("📦 Stock Requests | طلبات المخزون")

tabs = st.tabs([
    "📋 الطلبات | Requests",
    "✅ الموافقة | Approved",
    "❌ غير متوفر | Unavailable",
    "🛒 تم الطلب | Ordered",
    "📅 الجدولة | Scheduled",
    "🚫 جدولة ملغية | Cancelled",
    "🔄 تعديل موعد | Rescheduled",
    "⚠️ تنبيهات | Alerts",
    "📊 المخزون | Inventory",
    "🔴 مخزون منخفض | Low Stock",
    "🗂️ منتهية | Expired",
    "⚙️ الإعدادات | Settings",
])
(tab1,tab2,tab3,tab4,tab5,tab6,tab7,tab8,tab9,tab10,tab11,tab12) = tabs

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
                dn = now_str()
                safe_batch_append(unavailable_sheet, [[r[0],r[1],r[2] if len(r)>2 else "",r[3] if len(r)>3 else "",dn] for r in rows])
                safe_delete_all(requests_sheet)
                st.session_state["confirm_reject_all"] = False
                st.rerun()
            if cn.button("❌ لا | No", key="no_rej_all"):
                st.session_state["confirm_reject_all"] = False
                st.rerun()

        confirm_clear("clear_req", requests_sheet, "الطلبات | Requests")

        # بناء قائمة SKUs الموجودة في Ordered
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
                        safe_append(unavailable_sheet,[sku,qty,img,date_added,now_str()])
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
        filtered = [r for r in rows_ap if not srch or srch.strip().upper() in r[0].upper()]
        df_ap = pd.DataFrame(rows_ap, columns=data_ap[0])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_ap,"approved")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_ap", use_container_width=True):
                st.session_state["confirm_clear_ap"] = True
        confirm_clear("clear_ap", approved_sheet, "الموافقة | Approved")
        st.write(f"**عرض | Showing: {len(filtered)} / {len(rows_ap)}**")
        for row in filtered:
            ri = rows_ap.index(row)+2
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
        filtered = [r for r in rows_un if not srch or srch.strip().upper() in r[0].upper()]
        df_un = pd.DataFrame(rows_un, columns=data_un[0])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_un,"unavailable")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_un", use_container_width=True):
                st.session_state["confirm_clear_un"] = True
        confirm_clear("clear_un", unavailable_sheet, "غير المتوفر | Unavailable")
        st.write(f"**عرض | Showing: {len(filtered)} / {len(rows_un)}**")
        for row in filtered:
            ri = rows_un.index(row)+2
            while len(row)<5: row.append("")
            sku,qty,img,da,dm = row[0],row[1],row[2],row[3],row[4]
            c_img,c_info,c_del = st.columns([1,5,1])
            with c_img: show_img(img,70)
            with c_info:
                st.markdown(f"**SKU:** `{sku}`")
                show_sku_inv(sku)
                st.markdown(f"**Qty طلب | Requested:** {qty}")
                st.caption(f"📅 Requested | طُلب: {da} | ❌ Unavailable | غير متوفر: {dm}")
            with c_del:
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
        filtered = [r for r in rows_ord if not srch or srch.strip().upper() in r[0].upper()]
        df_ord = pd.DataFrame(rows_ord, columns=data_ord[0])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_ord,"ordered")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_ord", use_container_width=True):
                st.session_state["confirm_clear_ord"] = True
        confirm_clear("clear_ord", ordered_sheet, "تم الطلب | Ordered")
        st.write(f"**عرض | Showing: {len(filtered)} / {len(rows_ord)}**")
        for row in filtered:
            ri = rows_ord.index(row)+2
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
                    ds   = pd_.strftime("%Y-%m-%d") if pd_ else dval
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

        # ترتيب من الأقرب للأبعد
        def sort_key(r):
            d = parse_excel_date(r[3] if len(r)>3 else "")
            return d if d else datetime(2099,1,1)
        rows_sch_sorted = sorted(rows_sch, key=sort_key)

        # تجميع SKUs تحت كل ASN
        asn_groups = {}
        for r in rows_sch_sorted:
            while len(r)<6: r.append("")
            asn = r[0].strip()
            if asn not in asn_groups:
                asn_groups[asn] = {"date":r[3],"skus":[]}
            asn_groups[asn]["skus"].append(r)

        df_sch = pd.DataFrame(rows_sch, columns=data_sch[0])
        c1,c2 = st.columns(2)
        with c1: dl_btn(df_sch,"scheduled")
        with c2:
            if st.button("🗑️ مسح الكل | Clear All", type="secondary", key="btn_clear_sc", use_container_width=True):
                st.session_state["confirm_clear_sc"] = True
        confirm_clear("clear_sc", scheduled_sheet, "الجدولة | Schedule")

        srch_asn = st.text_input("🔍 بحث ASN | Search by ASN", key="srch_asn", placeholder="اكتب رقم ASN...")
        today = datetime.now().date()
        st.write(f"**إجمالي ASN | Total ASNs: {len(asn_groups)}**")

        for asn, group in asn_groups.items():
            if srch_asn and srch_asn.strip().upper() not in asn.upper():
                continue
            sdate   = group["date"]
            pd_date = parse_excel_date(sdate)
            is_exp  = pd_date and today > pd_date.date() + timedelta(days=1)
            skus_   = group["skus"]
            has_alert = any(
                inv_map.get(r[1].strip().upper(),{}).get("sales",0) > 0 and
                _to_int(r[2]) > inv_map.get(r[1].strip().upper(),{}).get("sales",0)
                for r in skus_)
            border = "#ef4444" if has_alert else "#f59e0b" if is_exp else "#3b82f6"
            bg     = "#2d1515" if has_alert else "#2d2000" if is_exp else "#0f172a"

            st.markdown(
                f'<div style="border-left:5px solid {border};background:{bg};border-radius:10px;padding:8px 14px;margin-bottom:4px;">'
                f'<b>ASN:</b> {asn} &nbsp;|&nbsp; 📅 <b>تاريخ الجدولة | Schedule Date:</b> <b>{sdate}</b></div>',
                unsafe_allow_html=True)

            for r in skus_:
                while len(r)<6: r.append("")
                sku,qty,img = r[1].strip(),r[2],r[4]
                info    = inv_map.get(sku.upper(),{})
                monthly = info.get("sales",0)
                is_al   = monthly>0 and _to_int(qty)>monthly
                c_img2,c_info2 = st.columns([1,6])
                with c_img2: show_img(img,60)
                with c_info2:
                    st.markdown(f"&nbsp;&nbsp;**SKU:** `{sku}` | **Qty:** {qty}")
                    show_sku_inv(sku)
                    if is_al:
                        st.markdown(f"&nbsp;&nbsp;🔴 **تنبيه | Alert:** الكمية ({qty}) > المبيع ({monthly})")

            ca,cb,cc = st.columns(3)
            with ca:
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
            with cb:
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
            with cc:
                status = "⚠️ منتهي | Expired" if is_exp else "✅ ساري | Active"
                st.markdown(f"&nbsp;{status}")
            st.divider()

