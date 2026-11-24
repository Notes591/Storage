# -*- coding: utf-8 -*-
"""
PG_Verify.py — بيقارن حسابات SQL (acct_sql.py) بالحسابات القديمة (calc.py) على بياناتك الحقيقية في Neon.
ممنوع يكتب أي حاجة: قراءة بس. شغّله من نفس مجلد stock_requests.py/calc.py/acct_sql.py:
    py PG_Verify.py            (بيقرا الرابط من pg_url.txt جنبه، أو متغير PG_URL)
    set PG_ONLY=noon           (اختياري: noon / ty / amz)
كل فحص بيطبع PASS لو الأرقام متطابقة (فرق أقل من 0.01) أو FAIL ومعاه أكبر فرق. أي FAIL ابعتهولي.
"""
import os, re, sys, time, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd
import psycopg2
import calc
import acct_sql as A

HERE = os.path.dirname(os.path.abspath(__file__))


def connect():
    loc = os.environ.get("PG_LOCAL", "")          # للاختبار المحلي: host:port:user:db
    if loc:
        h, p, u, d = loc.split(":")
        return psycopg2.connect(host=h, port=int(p), user=u, dbname=d)
    u = os.environ.get("PG_URL", "").strip()
    if not u:
        try:
            u = open(os.path.join(HERE, "pg_url.txt"), encoding="utf-8").read().strip()
        except Exception:
            u = ""
    u = re.sub(r"[&?]channel_binding=[^&]*", "", u.strip().strip('"'))
    m = re.search(r"postgres(?:ql)?://([^:@/\s]+):([^@\s]+)@([^:/?&\s'\"]+)(?::(\d+))?/([A-Za-z0-9_\-]+)", u)
    if not m:
        sys.exit("مفيش رابط Neon (pg_url.txt أو PG_URL)")
    return psycopg2.connect(user=m.group(1), password=m.group(2), host=m.group(3), port=int(m.group(4) or 5432),
                            dbname=m.group(5), sslmode="require", connect_timeout=20)


CONN = connect()
CONN.set_session(readonly=True, autocommit=True)


def q(sql, params=None):
    cur = CONN.cursor()
    if not getattr(q, "fn", False):
        try:
            CONN.set_session(readonly=False, autocommit=True)
            cur.execute(A.NOON_ACCT_NUM_SQL)
            CONN.set_session(readonly=True, autocommit=True)
        except Exception as e:
            print("تحذير: مفيش صلاحية إنشاء دالة acct_num:", e)
        q.fn = True
    cur.execute(sql, params)
    cols = [d[0] for d in cur.description]
    return cols, cur.fetchall()


def raw_tab(title):
    cur = CONN.cursor()
    cur.execute("SELECT table_name, headers, col_names FROM sheet_copy._meta WHERE tab_title = %s", (title,))
    m = cur.fetchone()
    if not m:
        return None
    t, hdr, cols = m
    cur.execute("SELECT " + ", ".join('"%s"' % c for c in cols) + ' FROM sheet_copy."%s" ORDER BY _row' % t)
    return [list(hdr)] + [["" if v is None else v for v in r] for r in cur.fetchall()]


RES = []


def report(name, ok, detail=""):
    RES.append(ok)
    print(("PASS " if ok else "FAIL ") + name + (("  — " + detail) if detail else ""), flush=True)


def cmp_df(name, old, new, key, tol=0.01):
    """مقارنة جدولين على مفتاح: نفس الصفوف، النصوص متطابقة، الأرقام بفرق < tol."""
    if (old is None or old.empty) and (new is None or new.empty):
        return report(name, True, "فاضيين الاتنين")
    if old is None or new is None or old.empty != new.empty:
        return report(name, False, f"واحد فاضي: القديم {0 if old is None else len(old)} صف · SQL {0 if new is None else len(new)} صف")
    old, new = old.reset_index(drop=True), new.reset_index(drop=True)
    if list(old.columns) != list(new.columns):
        return report(name, False, f"الأعمدة مختلفة: {list(old.columns)} ≠ {list(new.columns)}")
    if len(old) != len(new):
        return report(name, False, f"عدد الصفوف: القديم {len(old)} · SQL {len(new)}")
    m = old.merge(new, on=key, how="outer", suffixes=("_a", "_b"), indicator=True)
    miss = int((m["_merge"] != "both").sum())
    worst, bad_txt = 0.0, []
    for c in old.columns:
        if c in key:
            continue
        a_, b_ = m[c + "_a"], m[c + "_b"]
        if pd.api.types.is_numeric_dtype(old[c]) and pd.api.types.is_numeric_dtype(new[c]):
            d = (a_.astype(float).fillna(0) - b_.astype(float).fillna(0)).abs().max()
            worst = max(worst, float(d) if d == d else 0.0)
        else:
            nb = int((a_.astype(str).replace("None", "") != b_.astype(str).replace("None", "")).sum())
            if nb:
                bad_txt.append(f"{c}:{nb}")
    ok = miss == 0 and worst < tol and not bad_txt
    report(name, ok, f"{len(old)} صف · أكبر فرق رقمي {worst:.6f}" + (f" · صفوف ناقصة {miss}" if miss else "") + (f" · نصوص مختلفة {bad_txt}" if bad_txt else ""))


def noon():
    t0 = time.time()
    fin_raw, ord_raw = raw_tab("Noon_Finance_API"), raw_tab("Noon_Orders_API")
    if not fin_raw:
        return print("مفيش Noon_Finance_API في PostgreSQL")
    fin = calc.acct_rows_to_df(fin_raw)
    orders = calc.acct_rows_to_df(ord_raw) if ord_raw else None
    print(f"نون: دفتر {len(fin):,} صف · أوردرات {0 if orders is None else len(orders):,} صف · تحميل {time.time() - t0:.0f} ث", flush=True)
    dts = pd.to_datetime(fin["Transaction Date"], errors="coerce").dt.date
    dmin, dmax = dts.min(), dts.max()
    ranges = [(dmin, dmax), (max(dmin, dmax - datetime.timedelta(days=60)), dmax)]
    for (d1, d2) in ranges:
        tag = f"{d1}→{d2}"
        S = calc.acct_noon_summary(fin, d1, d2)
        p = {"stores": None, "d_from": d1, "d_to": d2}
        # الدفعات / الأصناف / اليومي
        c, r = q(A.NOON_PAY_SQL, p)
        pay = pd.DataFrame(r, columns=c)
        old = S["payments"].reset_index(drop=True).assign(**{"Total": lambda x: x["Total"].astype(float)})
        old = old.assign(**{"التاريخ": old["التاريخ"].astype(str)}); pay = pay.assign(**{"التاريخ": pay["التاريخ"].astype(str), "Total": pay["Total"].astype(float)})
        cmp_df(f"نون الدفعات {tag}", old.drop(columns=["Title"], errors="ignore").sort_values(["Reference Nr", "Total"]).reset_index(drop=True),
               pay.drop(columns=["Title"], errors="ignore").sort_values(["Reference Nr", "Total"]).reset_index(drop=True), ["Reference Nr", "Total", "التاريخ"])
        c, r = q(A.NOON_SKU_SQL, p)
        sk = pd.DataFrame(r, columns=c)
        for col in sk.columns:
            if col not in ("SKUs", "Partner SKU"):
                sk[col] = pd.to_numeric(sk[col]).astype(float)
        o2 = S["skus"].reset_index(drop=True)
        for col in ("مبيع", "مرتجع"):
            o2[col] = o2[col].astype(float)
        cmp_df(f"نون حسب الصنف {tag}", o2, sk, ["SKUs"])
        c, r = q(A.NOON_DAILY_SQL, p)
        dl = pd.DataFrame(r, columns=c)
        for col in ("Net Proceeds", "Total"):
            dl[col] = dl[col].astype(float)
        od = S["daily"].reset_index(drop=True).assign(**{"التاريخ": lambda x: x["التاريخ"].astype(str)}); dl["التاريخ"] = dl["التاريخ"].astype(str)
        cmp_df(f"نون اليومي {tag}", od, dl, ["التاريخ"])
        # الربحية
        cf = lambda s_, p_: (len(s_) % 50 + 1.0) if len(s_) % 3 else None
        for vat in (1.0, 1.15):
            pf_old = calc.acct_noon_profit(fin, d1, d2, cf, None, vat=vat, alloc_order=True, order_items=orders)
            c, r = q(A.NOON_PROFIT_G_SQL, p)
            g = pd.DataFrame(r, columns=c)
            for col in g.columns:
                if col not in ("SKU", "المنتج", "Partner SKU"):
                    g[col] = pd.to_numeric(g[col]).astype(float).round(6)
            c, r = q(A.NOON_PROFIT_O_SQL, p)
            o_tot = o_sub = im = None
            if r:
                od_ = pd.DataFrame(r, columns=c)
                o_tot = pd.Series(od_["tot"].astype(float).values, index=od_["on_"].values); o_tot.index.name = "_on"
                o_sub = pd.Series(od_["sub"].astype(float).values, index=od_["on_"].values); o_sub.index.name = "_on"
                c, r = q(A.NOON_PROFIT_IM_SQL, p)
                im = pd.DataFrame(r, columns=c).rename(columns={"on_": "_on"})
                im = im[im["_on"].isin(o_tot.index)].reset_index(drop=True)
            def oi(miss):
                if not miss or orders is None:
                    return None
                c2, r2 = q("SELECT order_nr, sku, item_nr FROM sheet_copy.noon_orders_api WHERE btrim(order_nr) = ANY(%s)", (list(miss),))
                return pd.DataFrame(r2, columns=c2)
            pf_new = calc.acct_noon_profit_finish((g, o_tot, o_sub, im), cf, None, vat, True, oi)
            cmp_df(f"نون ربحية المنتجات {tag} vat={vat}", pf_old.drop(columns=["تكلفة ناقصة"]), pf_new.drop(columns=["تكلفة ناقصة"]), ["SKU"])
        # العمولة والتوصيل + آخر أوردر
        if orders is not None:
            for chan in ("all", "FBN", "FBB"):
                for vat in (1.0, 1.15):
                    try:
                        old = calc.acct_noon_com(fin, d1, d2, orders, chan, vat)
                        g, unk, inc = A.noon_com(q, None, d1, d2, chan)
                        new = calc.acct_noon_com_finish(g, vat, unk, inc)
                        cmp_df(f"نون العمولة والتوصيل {tag} {chan} vat={vat}", old, new, ["SKU"])
                        report(f"   عدّادات (غير معروف/ناقص) {chan}", old.attrs.get("unknown") == new.attrs.get("unknown") and old.attrs.get("incomplete") == new.attrs.get("incomplete"),
                               f"قديم {old.attrs.get('unknown')}/{old.attrs.get('incomplete')} · SQL {new.attrs.get('unknown')}/{new.attrs.get('incomplete')}")
                    except Exception as e:
                        report(f"نون العمولة والتوصيل {tag} {chan}", False, "خطأ: " + str(e)[:200])
                try:
                    old = calc.acct_noon_last(fin, d1, d2, orders, chan, 1.15)
                    rows, today = A.noon_last(q, None, d1, d2, chan)
                    new = calc.acct_noon_last_finish(calc.acct_noon_last_from_sql(rows), today, 1.15)
                    cmp_df(f"نون آخر أوردر {tag} {chan}", old, new, ["SKU"])
                except Exception as e:
                    report(f"نون آخر أوردر {tag} {chan}", False, "خطأ: " + str(e)[:200])
        break_ = None
    if orders is not None:
        try:
            today = datetime.date.today()
            old = calc.acct_noon_pending(fin, orders, today)
            new = A.noon_pending(q, None, today)
            cols = ["order_nr", "item_nr", "sku", "delivered_at", "أيام من التسليم", "الحالة", "Store"]
            o_ = old[cols].reset_index(drop=True) if not old.empty else old
            cmp_df("نون ⏳ لسه ماظهرش", o_.assign(**{"أيام من التسليم": o_["أيام من التسليم"].astype(float)}) if not o_.empty else o_,
                   new[cols].assign(**{"أيام من التسليم": new["أيام من التسليم"].astype(float)}) if not new.empty else new, ["item_nr", "order_nr"])
        except Exception as e:
            report("نون ⏳ لسه ماظهرش", False, "خطأ: " + str(e)[:200])


if __name__ == "__main__":
    only = os.environ.get("PG_ONLY", "").lower()
    if only in ("", "noon"):
        noon()
    print("\n" + ("✅ كل الفحوصات PASS" if all(RES) else f"❌ فيه {RES.count(False)} فحص FAIL من {len(RES)}"))
