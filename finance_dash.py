# -*- coding: utf-8 -*-
"""
finance_dash.py — لوحة القسم المالي (غرفة القيادة) من بيانات ميزان المراجعة TB وجرد INV في ملف Finance.
بتتفتح من finance_tb.render كأول تاب. كل الأرقام بتتحسب من التابات (مفيش أرقام ثابتة).

ربط الحسابات (غيّره من هنا لو أكواد المحاسب اختلفت):
"""
import re

MAP = {
    "sales_prefix": ("401",),                 # حسابات المبيعات (أكواد تبدأ بـ)
    "revenue_prefix": ("4",),                 # كل الإيرادات (عشان نلاقي المرتجعات والخصومات بالاسم)
    "returns_words": ("مرتجع", "مردود"),      # اسم الحساب فيه كلمة من دول = مرتجعات
    "discount_words": ("خصم",),               # = خصومات
    "commission_prefix": ("50132",),          # عمولات المنصات
    "expense_prefix": ("5",),                 # كل المصاريف
    "purchases_main01": ("صافي المشتريات",),  # المشتريات (التصنيف الرئيسي في الميزان) — مش جوه المصروفات
    "cash_main02": ("الاموال الجاهزة",),      # بنوك وصناديق وبطاقات
    "supplier_prefix": ("221", "222"),        # موردين (الرصيد الدائن) ودفعات مقدمة (الرصيد المدين) — نفس الحسابات
    "platform_prefix": ("122",),              # أرصدة المنصات (نون/أمازون/تريندول...)
}
# تجميع المصاريف (الأول اللي يطابق): (اسم، أكواد تبدأ بـ)
EXPENSE_GROUPS = [
    ("عمولات منصات البيع", ("50132",)),
    ("الرواتب والأجور والحوافز", ("501010", "50125")),
    ("الصيانة والإصلاح", ("50122", "50109")),
    ("شؤون الموظفين والتأمينات", ("50116", "50114")),
    ("الإعلانات والتسويق والدعاية", ("50117", "50102")),
    ("عمولة وكلاء بيع", ("50124",)),
    ("شحن وتغليف وتحميل البضاعة", ("50110", "50115", "50118")),
    ("نقل وسيارات وبنوك", ("50107", "50112", "50111")),
]
PRIVATE_PREFIX = ("123", "124", "126", "213")     # C13: حسابات الشركاء والسلف والعهد — بتتشال من أي تفصيل

MONTHS_AR = {"01": "يناير", "02": "فبراير", "03": "مارس", "04": "أبريل", "05": "مايو", "06": "يونيو",
             "07": "يوليو", "08": "أغسطس", "09": "سبتمبر", "10": "أكتوبر", "11": "نوفمبر", "12": "ديسمبر"}

CSS = """
<style>
.fd-card{background:#fff;border:1px solid #e2e8f0;border-radius:16px;padding:16px 18px;box-shadow:0 2px 8px rgba(15,23,42,.05);direction:rtl;color:#0f172a;height:100%}
.fd-lbl{font-size:12.5px;font-weight:700;color:#334155}
.fd-val{font-size:24px;font-weight:800;margin:6px 0;direction:ltr;text-align:right}
.fd-sub{font-size:11.5px;color:#64748b;border-top:1px solid #f1f5f9;padding-top:7px;margin-top:4px}
.fd-banner{background:linear-gradient(135deg,#fff,#f0f9ff);border:1px solid #bfdbfe;border-radius:18px;padding:18px 22px;direction:rtl;color:#0f172a}
.fd-banner h3{margin:4px 0 6px;font-size:17px;color:#0f172a}
.fd-banner p{font-size:13px;color:#334155;line-height:1.8;margin:0}
.fd-warn{background:#fff1f2;border:1px solid #fecdd3;color:#881337;border-radius:12px;padding:10px 14px;font-size:12.5px;direction:rtl;margin:6px 0}
.fd-ok{background:#ecfdf5;border:1px solid #a7f3d0;color:#065f46;border-radius:12px;padding:10px 14px;font-size:12.5px;direction:rtl;margin:6px 0}
.fd-box{border-radius:14px;padding:14px;direction:rtl;font-size:12px;line-height:1.8;height:100%}
.fd-box b.big{display:block;font-size:20px;margin:4px 0;direction:ltr;text-align:right}
.pos{color:#059669}.neg{color:#e11d48}
</style>
"""


def _sar(v):
    return f"{'-' if v < 0 else ''}{abs(round(v)):,.0f}"


def _mname(p):
    y, m = p.split("-")
    return f"{MONTHS_AR.get(m, m)} {y}"


def _card(lbl, val, sub="", cls=""):
    return (f"<div class='fd-card'><div class='fd-lbl'>{lbl}</div><div class='fd-val {cls}'>{val}</div>"
            f"<div class='fd-sub'>{sub}</div></div>")


def _starts(code, prefixes):
    return str(code).startswith(tuple(prefixes))


def _has(name, words):
    n = str(name)
    return any(w in n for w in words)


def build_monthly(pd, tb, year):
    """جدول شهري: مبيعات إجمالية، مرتجعات، خصومات، صافي، عمولات، مصاريف أخرى، فائض، مشتريات."""
    d = tb[(tb["Year"] == year) & (tb["Period"] != "Opening")].copy()
    if d.empty:
        return pd.DataFrame()
    d["AccountCode"] = d["AccountCode"].astype(str)
    d = d[~d["AccountCode"].map(lambda c: _starts(c, PRIVATE_PREFIX))]
    rev = d[d["AccountCode"].map(lambda c: _starts(c, MAP["revenue_prefix"]))]
    ret_m = rev["AccountName"].map(lambda n: _has(n, MAP["returns_words"]))
    dis_m = rev["AccountName"].map(lambda n: _has(n, MAP["discount_words"])) & ~ret_m
    sal_m = rev["AccountCode"].map(lambda c: _starts(c, MAP["sales_prefix"])) & ~ret_m & ~dis_m
    if not sal_m.any():
        sal_m = ~ret_m & ~dis_m
    exp = d[d["AccountCode"].map(lambda c: _starts(c, MAP["expense_prefix"]))]
    com_m = exp["AccountCode"].map(lambda c: _starts(c, MAP["commission_prefix"]))
    pur = d[d["Main01"].isin(MAP["purchases_main01"])].groupby("Period")["Amount"].sum()
    g = lambda df, m: df[m].groupby("Period")["Amount"].sum().abs()
    out = pd.DataFrame({
        "gross": g(rev, sal_m), "returns": g(rev, ret_m), "discounts": g(rev, dis_m),
        "commissions": g(exp, com_m), "purchases": pur,
        "opex_all": exp.groupby("Period")["Amount"].sum().abs(),
    }).fillna(0.0).sort_index()
    out["net"] = out["gross"] - out["returns"] - out["discounts"]
    out["opex"] = (out["opex_all"] - out["commissions"]).clip(lower=0)
    out["surplus"] = out["net"] - out["commissions"] - out["opex"]
    out["name"] = [_mname(p) for p in out.index]
    return out


def balances(pd, tb, year, prefixes=(), main02=()):
    """رصيد كل حساب (افتتاحي + حركة الشهور) للأكواد اللي تبدأ بـ prefixes أو التصنيف الرئيسي 02 = main02."""
    d = tb[(tb["Year"] == year)].copy()
    d["AccountCode"] = d["AccountCode"].astype(str)
    sel = d["AccountCode"].map(lambda c: _starts(c, prefixes)) if prefixes else d["Main02"].isin(main02)
    d = d[sel & ~d["AccountCode"].map(lambda c: _starts(c, PRIVATE_PREFIX))]
    if d.empty:
        return pd.DataFrame(columns=["AccountCode", "AccountName", "Balance"])
    return (d.groupby(["AccountCode", "AccountName"], as_index=False)["Amount"].sum().rename(columns={"Amount": "Balance"}))


def render(st, pd, tb, inv):
    st.markdown(CSS, unsafe_allow_html=True)
    if tb is None or tb.empty:
        st.info("مفيش بيانات ميزان مراجعة لسه.")
        return
    tb = tb.copy()
    tb["Amount"] = pd.to_numeric(tb["Amount"], errors="coerce").fillna(0.0)
    tb["Year"] = tb["Year"].astype(str).str.replace(r"\.0$", "", regex=True)
    years = sorted(tb["Year"].unique(), reverse=True)
    yr = st.selectbox("السنة المالية", years, key="fd_year") if len(years) > 1 else years[0]
    mon = build_monthly(pd, tb, yr)
    if mon.empty or mon["net"].sum() == 0:
        st.warning("مش قادر أحسب المبيعات من الميزان — ربط الحسابات محتاج تعديل (شوف «تشخيص ربط الحسابات» تحت).")
    # الشهور الناقصة: مبيعاتها أقل من 25% من الوسيط
    med = mon["net"].median() if not mon.empty else 0
    mon["partial"] = mon["net"] < 0.25 * med if med else False
    full = mon[~mon["partial"]]

    # ── الأرصدة ──
    cash = balances(pd, tb, yr, main02=MAP["cash_main02"])
    sup = balances(pd, tb, yr, MAP["supplier_prefix"])
    adv = sup[sup["Balance"] > 0]
    plat = balances(pd, tb, yr, MAP["platform_prefix"])
    cash_total = float(cash["Balance"].sum())
    sup_liab = float(-sup.loc[sup["Balance"] < 0, "Balance"].sum())
    adv_total = float(adv["Balance"].sum())
    plat_total = float(plat["Balance"].sum())
    inv_val, inv_qty = 0.0, 0.0
    if inv is not None and not inv.empty:
        iv = inv.copy()
        iv["Year"] = iv["Year"].astype(str).str.replace(r"\.0$", "", regex=True)
        iv = iv[iv["Year"] == yr] if (iv["Year"] == yr).any() else iv
        for c in ("Qty", "Value"):
            iv[c] = pd.to_numeric(iv[c], errors="coerce").fillna(0.0)
        inv_val, inv_qty = float(iv["Value"].sum()), float(iv["Qty"].sum())
    else:
        iv = None

    # ── 1) البانر ──
    if not full.empty:
        peak = full["net"].idxmax()
        last = full.index[-1]
        comm_pct = full["commissions"].sum() / full["net"].sum() * 100 if full["net"].sum() else 0
        head = f"صافي المبيعات {full['net'].sum() / 1e6:,.2f} مليون ر.س في {len(full)} شهر مكتمل، والفائض التشغيلي {full['surplus'].sum() / 1e6:,.2f} مليون"
        desc = (f"ذروة المبيعات كانت في <b>{mon.loc[peak, 'name']}</b> ({mon.loc[peak, 'net'] / 1e6:,.2f}M)، وآخر شهر مكتمل <b>{mon.loc[last, 'name']}</b> "
                f"({mon.loc[last, 'net'] / 1e6:,.2f}M). عمولات المنصات = <b>{comm_pct:.1f}%</b> من صافي المبيعات. السيولة بالبنوك <b>{cash_total / 1e6:,.2f}M</b> "
                f"والمخزون <b>{inv_val / 1e6:,.2f}M</b>.")
    else:
        head, desc = "لسه مفيش شهور مكتملة", ""
    st.markdown(f"<div class='fd-banner'><span style='font-size:11.5px;font-weight:800;color:#1d4ed8'>🧭 نبض الشركة واتجاهها</span><h3>{head}</h3><p>{desc}</p></div>",
                unsafe_allow_html=True)
    for p in mon.index[mon["partial"]]:
        st.markdown(f"<div class='fd-warn'>⚠️ <b>{mon.loc[p, 'name']}</b> غير مكتمل في الميزان (مبيعاته {_sar(mon.loc[p, 'net'])} فقط) — غالبًا فواتير المبيعات لسه متقيّدتش. "
                    f"مش داخل في الإجماليات.</div>", unsafe_allow_html=True)
    st.write("")

    # ── 2) كروت المؤشرات ──
    mg = (full["surplus"].sum() / full["net"].sum() * 100) if (not full.empty and full["net"].sum()) else 0
    cards = [
        ("صافي المبيعات (الشهور المكتملة)", _sar(full["net"].sum()) if not full.empty else "—", f"{len(full)} شهر • الذروة: {mon.loc[full['net'].idxmax(), 'name'] if not full.empty else '—'}", ""),
        ("الفائض التشغيلي (قبل تكلفة البضاعة)", _sar(full["surplus"].sum()) if not full.empty else "—", f"هامش {mg:.1f}% بعد العمولات والمصاريف", "pos"),
        ("السيولة بالبنوك والصندوق", _sar(cash_total), f"{len(cash)} حساب", ""),
        ("المخزون (جرد INV)", _sar(inv_val), f"{inv_qty:,.0f} وحدة", ""),
        ("دفعات مقدمة وأصول قيد التوريد", _sar(adv_total), "حسابات مدينة عند الموردين", ""),
        ("التزامات الموردين", _sar(sup_liab), f"{int((sup['Balance'] < 0).sum())} مورد دائن", "neg"),
    ]
    for row in (cards[:3], cards[3:]):
        cols = st.columns(3)
        for c, (l, v, s, cl) in zip(cols, row):
            c.markdown(_card(l, v + " ر.س" if v != "—" else v, s, cl), unsafe_allow_html=True)
        st.write("")

    # ── 3) المسار + تشخيص الشهر ──
    st.markdown("#### 📈 مسار المبيعات والفائض والعمولات")
    if not mon.empty:
        ch = mon[["net", "surplus", "commissions"]].rename(columns={"net": "صافي المبيعات", "surplus": "الفائض", "commissions": "العمولات"})
        ch.index = mon["name"]
        st.line_chart(ch, height=260)
        names = list(mon["name"])
        sel = st.radio("اختر الشهر للتشخيص", names, index=len(names) - 1, horizontal=True, key="fd_month")
        p = mon.index[names.index(sel)]
        i = list(mon.index).index(p)
        cur = mon.loc[p]
        prev = mon.iloc[i - 1] if i > 0 else None
        c1, c2 = st.columns([1, 1])
        def chg(a, b):
            return (a - b) / b * 100 if b else 0.0
        with c1:
            sc = chg(cur["net"], prev["net"]) if prev is not None else None
            uc = chg(cur["surplus"], prev["surplus"]) if prev is not None else None
            st.markdown(_card(f"صافي المبيعات — {sel}", _sar(cur["net"]) + " ر.س",
                              "شهر الأساس" if sc is None else f"{'↗' if sc >= 0 else '↘'} {sc:+.1f}% عن الشهر السابق", "pos" if (sc or 0) >= 0 else "neg"), unsafe_allow_html=True)
        with c2:
            st.markdown(_card(f"الفائض التشغيلي — {sel}", _sar(cur["surplus"]) + " ر.س",
                              "" if uc is None else f"{'↗' if uc >= 0 else '↘'} {uc:+.1f}% عن الشهر السابق", "pos" if cur["surplus"] >= 0 else "neg"), unsafe_allow_html=True)
        notes = []
        if cur["net"]:
            notes.append(f"المرتجعات {_sar(cur['returns'])} ر.س ({cur['returns'] / cur['gross'] * 100 if cur['gross'] else 0:.1f}% من الإجمالي)")
            notes.append(f"عمولات المنصات {_sar(cur['commissions'])} ر.س ({cur['commissions'] / cur['net'] * 100:.1f}% من صافي المبيعات)")
        if cur["purchases"]:
            notes.append(f"مشتريات البضاعة {_sar(cur['purchases'])} ر.س")
        notes.append(f"مصاريف تشغيل أخرى {_sar(cur['opex'])} ر.س")
        if prev is not None:
            for k, lbl in (("returns", "المرتجعات"), ("commissions", "العمولات"), ("opex", "المصاريف الأخرى")):
                if prev[k] and abs(cur[k] - prev[k]) / prev[k] > 0.3:
                    notes.append(f"⚠️ {lbl} {'زادت' if cur[k] > prev[k] else 'قلّت'} {abs(chg(cur[k], prev[k])):.0f}% عن الشهر السابق")
        st.markdown("<div class='fd-ok' style='background:#eff6ff;border-color:#bfdbfe;color:#1e3a8a'>" + "<br>".join("• " + n for n in notes) + "</div>", unsafe_allow_html=True)

    # ── 4) أين فلوس وبضاعة الشركة ──
    st.markdown("#### 🗺️ أين توجد أموال وبضائع الشركة اليوم؟")
    b = st.columns(4)
    def box(col, title, total, lines, bg, bd, fg):
        col.markdown(f"<div class='fd-box' style='background:{bg};border:1px solid {bd};color:{fg}'><b>{title}</b><b class='big'>{_sar(total)} ر.س</b>{lines}</div>", unsafe_allow_html=True)
    top = lambda df, n=3: "<br>".join(f"• {r.AccountName[:34]}: <b>{_sar(r.Balance)}</b>" for r in df.reindex(df["Balance"].abs().sort_values(ascending=False).index).head(n).itertuples())
    box(b[0], "1. سيولة نقدية جاهزة", cash_total, top(cash), "#ecfdf5", "#a7f3d0", "#065f46")
    inv_lines = ""
    if iv is not None and not iv.empty:
        w = iv.groupby("Warehouse")["Value"].sum().sort_values(ascending=False).head(3)
        inv_lines = "<br>".join(f"• {k[:30]}: <b>{_sar(v)}</b>" for k, v in w.items()) + f"<br>• الكمية: <b>{inv_qty:,.0f}</b> وحدة"
    box(b[1], "2. بضاعة بالمستودعات", inv_val, inv_lines, "#eff6ff", "#bfdbfe", "#1e3a8a")
    box(b[2], "3. دفعات مقدمة / قيد التوريد", adv_total, top(adv), "#fffbeb", "#fde68a", "#78350f")
    def _pg(n):
        n = str(n)
        return "نون" if "نون" in n else ("أمازون" if ("امازون" in n or "أمازون" in n) else ("تريندول" if ("تريندول" in n or "ترندول" in n) else "أخرى"))
    pl = plat.assign(g=plat["AccountName"].map(_pg)).groupby("g")["Balance"].sum().sort_values()
    pl_lines = "<br>".join(f"• {k}: <b>{_sar(v)}</b>" for k, v in pl.items()) + ("<br><em>(سالب = لسه ماتقيّدتش الفواتير)</em>" if plat_total < 0 else "")
    box(b[3], "4. أرصدة المنصات", plat_total, pl_lines, "#fff1f2", "#fecdd3", "#881337")
    st.write("")

    # ── 5) مقارنة شهرين + المصاريف ──
    st.markdown("#### ⚖️ مقارنة شهر بشهر وأين زاد الصرف؟")
    if len(mon) >= 2:
        names = list(mon["name"])
        ca, cb = st.columns(2)
        a = ca.selectbox("الشهر الحالي", names, index=len(names) - 1, key="fd_ma")
        bm = cb.selectbox("مقابل الشهر", names, index=max(0, len(names) - 2), key="fd_mb")
        A, B = mon.iloc[names.index(a)], mon.iloc[names.index(bm)]
        rows = []
        for k, lbl in (("net", "صافي المبيعات"), ("returns", "المرتجعات"), ("commissions", "عمولات المنصات"), ("opex", "مصاريف أخرى"), ("surplus", "الفائض التشغيلي")):
            d = A[k] - B[k]
            rows.append({"البند": lbl, a: A[k], bm: B[k], "الفرق": d, "التغيّر %": (d / B[k] * 100) if B[k] else 0.0})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True,
                     column_config={c: st.column_config.NumberColumn(c, format="%.0f") for c in (a, bm, "الفرق")} | {"التغيّر %": st.column_config.NumberColumn("التغيّر %", format="%.1f")})
    d = tb[(tb["Year"] == yr) & (tb["Period"] != "Opening")].copy()
    d["AccountCode"] = d["AccountCode"].astype(str)
    d = d[d["AccountCode"].map(lambda c: _starts(c, MAP["expense_prefix"])) & ~d["AccountCode"].map(lambda c: _starts(c, PRIVATE_PREFIX))]
    if not d.empty:
        def grp(c):
            for nm, pf in EXPENSE_GROUPS:
                if _starts(c, pf):
                    return nm
            return "مصاريف أخرى"
        d["g"] = d["AccountCode"].map(grp)
        pv = d.pivot_table(index="g", columns="Period", values="Amount", aggfunc="sum", fill_value=0.0).abs()
        tot = pv.sum(axis=1)
        peak = pv.idxmax(axis=1).map(_mname)
        ex = pd.DataFrame({"بند المصروف": pv.index, "الإجمالي": tot.values, "الحصة %": (tot / tot.sum() * 100).values, "شهر الذروة": peak.values,
                           "قيمة الذروة": pv.max(axis=1).values}).sort_values("الإجمالي", ascending=False)
        st.dataframe(ex, use_container_width=True, hide_index=True,
                     column_config={"الإجمالي": st.column_config.NumberColumn(format="%.0f"), "قيمة الذروة": st.column_config.NumberColumn(format="%.0f"),
                                    "الحصة %": st.column_config.NumberColumn(format="%.1f")})

    # ── 6) الموردين والبنوك ──
    st.markdown("#### 🏦 البنوك والموردين")
    t1, t2, t3 = st.tabs(["💵 البنوك والصناديق", "🚢 الموردين والدفعات", "📦 شواذ الجرد"])
    with t1:
        if cash.empty:
            st.info("مفيش حسابات نقدية مطابقة لأكواد الربط.")
        else:
            c = cash.sort_values("Balance", ascending=False).rename(columns={"AccountCode": "الكود", "AccountName": "الحساب", "Balance": "الرصيد"})
            c["الحصة %"] = c["الرصيد"] / c["الرصيد"].sum() * 100 if c["الرصيد"].sum() else 0.0
            st.dataframe(c, use_container_width=True, hide_index=True, column_config={"الرصيد": st.column_config.NumberColumn(format="%.0f"), "الحصة %": st.column_config.NumberColumn(format="%.1f")})
    with t2:
        s1, s2 = st.columns(2)
        sd = sup[sup["Balance"] < 0].sort_values("Balance").head(15).rename(columns={"AccountCode": "الكود", "AccountName": "المورد", "Balance": "المستحق"})
        ad = adv.sort_values("Balance", ascending=False).head(15).rename(columns={"AccountCode": "الكود", "AccountName": "الجهة", "Balance": "الدفعة المقدمة"})
        s1.markdown("**التزامات الموردين (أكبر 15)**")
        s1.dataframe(sd, use_container_width=True, hide_index=True, column_config={"المستحق": st.column_config.NumberColumn(format="%.0f")})
        s2.markdown("**دفعات مقدمة / قيد التوريد (أكبر 15)**")
        s2.dataframe(ad, use_container_width=True, hide_index=True, column_config={"الدفعة المقدمة": st.column_config.NumberColumn(format="%.0f")})
    with t3:
        if iv is None or iv.empty:
            st.info("مفيش جرد.")
        else:
            neg = iv[iv["Qty"] < 0]
            zc = iv[(iv["Qty"] > 0) & (iv["Value"] == 0)]
            m = st.columns(3)
            m[0].metric("أصناف بكمية سالبة", f"{len(neg):,}", f"{neg['Qty'].sum():,.0f} وحدة")
            m[1].metric("أصناف بتكلفة صفر", f"{len(zc):,}", f"{zc['Qty'].sum():,.0f} وحدة")
            m[2].metric("إجمالي أصناف معيبة", f"{len(neg) + len(zc):,}")
            f = st.radio("عرض", ["سالبة", "تكلفة صفر"], horizontal=True, key="fd_inv_f")
            x = (neg if f == "سالبة" else zc).rename(columns={"ItemCode": "الكود", "ItemName": "الصنف", "Warehouse": "المستودع", "Qty": "الكمية", "Price": "التكلفة", "Value": "القيمة"})
            st.dataframe(x[["الكود", "الصنف", "المستودع", "الكمية", "التكلفة", "القيمة"]].sort_values("الكمية", ascending=(f == "سالبة")).head(300),
                         use_container_width=True, hide_index=True)

    # ── 7) تشخيص الربط ──
    with st.expander("🔧 تشخيص ربط الحسابات (لو رقم مش مظبوط)", expanded=False):
        st.caption("الأرقام فوق بتتحسب من أكواد حسابات مربوطة في أول ملف finance_dash.py (MAP). لو أرقام المبيعات أو العمولات مش مطابقة لملف المحاسب، ابعتلي الجدول ده.")
        dd = tb[(tb["Year"] == yr) & (tb["Period"] != "Opening")].copy()
        dd["AccountCode"] = dd["AccountCode"].astype(str)
        top20 = dd.groupby(["AccountCode", "AccountName", "Main01"], as_index=False)["Amount"].sum()
        top20 = top20.reindex(top20["Amount"].abs().sort_values(ascending=False).index).head(40)
        st.dataframe(top20.rename(columns={"AccountCode": "الكود", "AccountName": "الحساب", "Main01": "الرئيسي", "Amount": "مجموع الحركة"}),
                     use_container_width=True, hide_index=True)
        if not mon.empty:
            st.dataframe(mon.drop(columns=["name"]).round(0), use_container_width=True)
