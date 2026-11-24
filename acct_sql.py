# -*- coding: utf-8 -*-
"""acct_sql.py — استعلامات SQL لحسابات المتاجر (PostgreSQL، schema sheet_copy).
من غير Streamlit، عشان نفس الاستعلامات تتستخدم في التطبيق (stock_requests.py) وفي أداة المقارنة PG_Verify.py."""

_N_TYPED = r"(CASE WHEN {c} ~ '^-?[0-9]+(\.[0-9]+)?$' THEN {c}::double precision ELSE sheet_copy.acct_num({c}) END)"


def _nt(c):
    return _N_TYPED.format(c=c)


NOON_ACCT_NUM_SQL = r"""
CREATE OR REPLACE FUNCTION sheet_copy.acct_num(x text) RETURNS double precision
LANGUAGE plpgsql IMMUTABLE AS $f$
DECLARE t text; neg boolean; parts text[]; v double precision;
BEGIN
  IF x IS NULL THEN RETURN 0; END IF;
  t := btrim(x);
  IF t = '' THEN RETURN 0; END IF;
  IF t ~ '^-?[0-9]+(\.[0-9]+)?$' THEN RETURN t::double precision; END IF;
  t := translate(t, '٠١٢٣٤٥٦٧٨٩٫٬', '0123456789.,');
  t := replace(replace(replace(replace(t, '−', '-'), '–', '-'), chr(160), ''), ' ', '');
  neg := left(t, 1) = '(' AND right(t, 1) = ')';
  t := btrim(t, '()');
  IF position(',' in t) > 0 AND position('.' in t) > 0 THEN
    IF strpos(reverse(t), '.') < strpos(reverse(t), ',') THEN
      t := replace(t, ',', '');
    ELSE
      t := replace(replace(t, '.', ''), ',', '.');
    END IF;
  ELSIF position(',' in t) > 0 THEN
    parts := string_to_array(t, ',');
    IF length(parts[array_length(parts, 1)]) = 3 AND array_length(parts, 1) > 1 AND ltrim(parts[1], '-') <> '0' THEN
      t := replace(t, ',', '');
    ELSE
      t := replace(t, ',', '.');
    END IF;
  END IF;
  BEGIN
    v := t::double precision;
  EXCEPTION WHEN others THEN
    RETURN 0;
  END;
  IF neg THEN RETURN -v; END IF;
  RETURN v;
END $f$;
"""

# الدفتر بأنواعه الصحيحة (CTE) — بنقراه مباشرة من جدول التاب (من غير View عشان إعادة بناء الجدول ما تتعطلش)
NOON_FIN_CTE = ("""
WITH f AS (
  SELECT store, COALESCE(reference_nr, '') AS ref, COALESCE(item_nr, '') AS item, transaction_type AS tt,
         CASE WHEN transaction_date ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}' THEN left(transaction_date, 10)::date END AS d,
         """ + _nt("net_proceeds") + """ AS net,
         """ + _nt("referral_fee_including_vat") + """ AS referral,
         """ + _nt("fullfilment_logistics_fees_including_vat") + """ AS fulfil,
         """ + _nt("shipping_credits_including_vat") + """ AS shipcr,
         """ + _nt("other_order_fees_including_vat") + """ AS other_fees,
         """ + _nt("order_subsidies_including_vat") + """ AS subsidy,
         """ + _nt("non_order_fees_including_vat") + """ AS nof,
         """ + _nt("non_order_subsidies_including_vat") + """ AS nos,
         """ + _nt("total") + """ AS total
  FROM sheet_copy.noon_finance_api
  WHERE (%(stores)s::text[] IS NULL OR store = ANY(%(stores)s))
)
""")

_EARN = "('order','order_update','statement_fee')"

# نسخة من الدفتر فيها أعمدة إضافية (SKU/العنوان/ترتيب الصف) للعروض التفصيلية فقط — الاستعلامات التانية بتفضل خفيفة
NOON_FIN_CTE_X = NOON_FIN_CTE.replace("COALESCE(item_nr, '') AS item,",
                                      "COALESCE(item_nr, '') AS item, COALESCE(skus, '') AS skus, COALESCE(partner_skus, '') AS pskus, COALESCE(title, '') AS title, _row,", 1)
assert "AS pskus" in NOON_FIN_CTE_X

NOON_PAY_SQL = NOON_FIN_CTE_X + """
SELECT d AS "التاريخ", ref AS "Reference Nr", title AS "Title", tt AS "Transaction Type", total AS "Total"
FROM f WHERE tt IN ('payment','balance_transfer') AND d >= %(d_from)s AND d <= %(d_to)s
ORDER BY d DESC, _row DESC
"""
NOON_SKU_SQL = NOON_FIN_CTE_X + f"""
SELECT skus AS "SKUs", (array_agg(pskus ORDER BY _row))[1] AS "Partner SKU",
  COUNT(*) FILTER (WHERE net > 0) AS "مبيع",
  COUNT(*) FILTER (WHERE tt = 'order_update' AND net < 0) AS "مرتجع",
  SUM(net) AS "Net Proceeds", SUM(referral) AS "Referral Fee including VAT",
  SUM(fulfil) AS "Fullfilment & Logistics Fees including VAT", SUM(other_fees) AS "Other Order Fees including VAT",
  SUM(subsidy) AS "Order Subsidies including VAT", SUM(total) AS "Total"
FROM f
WHERE tt IN {_EARN} AND btrim(item) <> '' AND btrim(skus) <> '' AND d >= %(d_from)s AND d <= %(d_to)s
GROUP BY skus ORDER BY SUM(total) DESC, skus
"""
NOON_UNALLOC_SQL = NOON_FIN_CTE + f"""
SELECT COALESCE(SUM(total), 0) FROM f
WHERE tt IN {_EARN} AND btrim(item) = '' AND tt <> 'statement_fee' AND d >= %(d_from)s AND d <= %(d_to)s
"""
NOON_FIN_CTE_O = NOON_FIN_CTE_X.replace("_row,", "_row, COALESCE(order_nr, '') AS onr,", 1)
assert "AS onr" in NOON_FIN_CTE_O

# ربحية المنتجات — المرحلة الأولى (تجميع الدفتر) = calc.acct_noon_profit_raw
NOON_PROFIT_G_SQL = NOON_FIN_CTE_X + """
SELECT skus AS "SKU", (array_agg(title ORDER BY _row))[1] AS "المنتج", (array_agg(pskus ORDER BY _row))[1] AS "Partner SKU",
  COUNT(*) FILTER (WHERE net > 0) AS "مبيع",
  COUNT(*) FILTER (WHERE tt = 'order_update' AND net < 0) AS "مرتجع",
  SUM(net) AS "صافي المبيعات",
  COALESCE(SUM(net) FILTER (WHERE net > 0), 0) AS "_بيع",
  SUM(referral + fulfil + other_fees + nof) AS "الرسوم",
  SUM(subsidy + nos) AS "دعم",
  SUM(total) AS "الصافي بعد الرسوم"
FROM f
WHERE tt IN ('order','order_update') AND btrim(item) <> '' AND btrim(skus) <> '' AND d >= %(d_from)s AND d <= %(d_to)s
GROUP BY skus ORDER BY skus COLLATE "C"
"""
NOON_PROFIT_O_SQL = NOON_FIN_CTE_O + """
SELECT btrim(onr) AS on_, SUM(total) AS tot, SUM(subsidy) AS sub
FROM f
WHERE tt IN ('order','order_update') AND btrim(skus) = '' AND btrim(onr) <> '' AND d >= %(d_from)s AND d <= %(d_to)s
GROUP BY btrim(onr)
"""
NOON_PROFIT_IM_SQL = NOON_FIN_CTE_O + """
SELECT DISTINCT ON (btrim(onr), item) btrim(onr) AS on_, skus AS "SKUs"
FROM f
WHERE tt IN ('order','order_update') AND btrim(item) <> '' AND btrim(skus) <> '' AND btrim(onr) <> '' AND d >= %(d_from)s AND d <= %(d_to)s
ORDER BY btrim(onr), item, _row
"""
NOON_OUTSIDE_SQL = NOON_FIN_CTE + f"""
SELECT COALESCE(SUM(total) FILTER (WHERE tt IN {_EARN}), 0), COALESCE(SUM(total) FILTER (WHERE tt = 'statement_fee'), 0)
FROM f WHERE d >= %(d_from)s AND d <= %(d_to)s
"""
NOON_DAILY_SQL = NOON_FIN_CTE + f"""
SELECT d AS "التاريخ", SUM(net) AS "Net Proceeds", SUM(total) AS "Total"
FROM f WHERE tt IN {_EARN} AND d >= %(d_from)s AND d <= %(d_to)s
GROUP BY d ORDER BY d
"""

NOON_KPI_SQL = NOON_FIN_CTE + f"""
SELECT
  COALESCE(SUM(net) FILTER (WHERE tt IN {_EARN}), 0) AS net,
  COALESCE(SUM(referral) FILTER (WHERE tt IN {_EARN}), 0) AS referral,
  COALESCE(SUM(fulfil) FILTER (WHERE tt IN {_EARN}), 0) AS fulfil,
  COALESCE(SUM(other_fees) FILTER (WHERE tt IN {_EARN}), 0) AS other_fees,
  COALESCE(SUM(subsidy) FILTER (WHERE tt IN {_EARN}), 0) AS subsidy,
  COALESCE(SUM(nof) FILTER (WHERE tt = 'statement_fee'), 0) AS ads,
  COALESCE(SUM(nof) FILTER (WHERE tt IN ('order','order_update')), 0) AS nonorder_fees_other,
  COALESCE(SUM(nos) FILTER (WHERE tt IN {_EARN}), 0) AS nonorder_sub,
  COALESCE(SUM(total) FILTER (WHERE tt IN {_EARN}), 0) AS earned,
  COALESCE(-SUM(total) FILTER (WHERE tt = 'payment'), 0) AS paid,
  COALESCE(-SUM(total) FILTER (WHERE tt = 'balance_transfer'), 0) AS transfers,
  COALESCE(SUM(total), 0) AS balance,
  COUNT(*) FILTER (WHERE tt = 'order' AND net > 0) AS orders_n,
  COUNT(*) FILTER (WHERE tt = 'order_update' AND net < 0) AS returns_n,
  COALESCE(SUM(net) FILTER (WHERE tt = 'order_update' AND net < 0), 0) AS returns_amt,
  COUNT(*) AS n
FROM f
WHERE d >= %(d_from)s AND d <= %(d_to)s
"""

NOON_STM_SQL = NOON_FIN_CTE + f"""
SELECT ref AS "Reference Nr", MAX(d) AS "تاريخ الكشف", COUNT(*) AS "صفوف",
  SUM(net) AS "Net Proceeds", SUM(referral) AS "Referral Fee including VAT", SUM(fulfil) AS "Fullfilment & Logistics Fees including VAT",
  SUM(shipcr) AS "Shipping Credits including VAT", SUM(other_fees) AS "Other Order Fees including VAT",
  SUM(subsidy) AS "Order Subsidies including VAT", SUM(nof) AS "Non-Order Fees including VAT",
  SUM(nos) AS "Non-Order Subsidies including VAT", SUM(total) AS "Total"
FROM f
WHERE tt IN {_EARN} AND ref LIKE 'PS-%%' AND d >= %(d_from)s AND d <= %(d_to)s
GROUP BY ref
ORDER BY MAX(d) DESC NULLS LAST
"""

# تفقيط الكشف: body = صفوف الأوردرات/الكشف، money = الدفعات وتحويلات الرصيد. picks = كشوف محددة (أو NULL = كل الفترة)
NOON_BRK_SQL = NOON_FIN_CTE + f""",
b AS (
  SELECT * FROM f WHERE tt IN {_EARN} AND (
    CASE WHEN %(picks)s::text[] IS NULL THEN d >= %(d_from)s AND d <= %(d_to)s ELSE ref = ANY(%(picks)s) END)
),
lim AS (SELECT MAX(d) AS dmax FROM b),
m AS (
  SELECT * FROM f WHERE tt IN ('payment','balance_transfer') AND (
    CASE WHEN %(picks)s::text[] IS NULL THEN d >= %(d_from)s AND d <= %(d_to)s
         ELSE (SELECT dmax FROM lim) IS NULL OR d <= (SELECT dmax FROM lim) END)
)
SELECT
  (SELECT COALESCE(SUM(net), 0) FROM b) AS net,
  (SELECT COALESCE(SUM(nos), 0) FROM b) AS nos,
  (SELECT COALESCE(SUM(subsidy) FILTER (WHERE NOT (btrim(item) = '' OR (ABS(subsidy - (-fulfil) * 0.2) < 0.02 AND fulfil < 0 AND subsidy > 0))), 0) FROM b) AS sub_ref,
  (SELECT COALESCE(SUM(subsidy) FILTER (WHERE (btrim(item) = '' OR (ABS(subsidy - (-fulfil) * 0.2) < 0.02 AND fulfil < 0 AND subsidy > 0))), 0) FROM b) AS sub_dlv,
  (SELECT COALESCE(SUM(other_fees), 0) FROM b) AS oth,
  (SELECT COALESCE(SUM(nof) FILTER (WHERE tt = 'statement_fee'), 0) FROM b) AS ads,
  (SELECT COALESCE(SUM(referral), 0) FROM b) AS referral,
  (SELECT COALESCE(SUM(fulfil) FILTER (WHERE btrim(item) = ''), 0) FROM b) AS ship_direct,
  (SELECT COALESCE(SUM(fulfil) FILTER (WHERE btrim(item) <> ''), 0) FROM b) AS ship_fbn,
  (SELECT COALESCE(SUM(nof) FILTER (WHERE tt <> 'statement_fee'), 0) FROM b) AS warranty,
  (SELECT COALESCE(SUM(total) FILTER (WHERE tt = 'payment'), 0) FROM m) AS pay,
  (SELECT COALESCE(SUM(total) FILTER (WHERE tt = 'balance_transfer'), 0) FROM m) AS bt
"""

# مدخلات التسوية بشكل مصغّر: كل كشف (متجر + مرجع) صف واحد، والدفعات/التحويلات صفوف صفوف، + أحدث تاريخ
NOON_RECON_STM_SQL = NOON_FIN_CTE + f"""
SELECT store AS "Store", ref AS "Reference Nr", MAX(d) AS "Transaction Date", 'order' AS "Transaction Type", SUM(total) AS "Total"
FROM f WHERE tt IN {_EARN} AND ref LIKE 'PS-%%' GROUP BY store, ref ORDER BY store, ref
"""
NOON_RECON_PAY_SQL = NOON_FIN_CTE + """
SELECT store AS "Store", ref AS "Reference Nr", d AS "Transaction Date", tt AS "Transaction Type", total AS "Total"
FROM f WHERE tt IN ('payment','balance_transfer') ORDER BY store, d NULLS LAST
"""
NOON_META_SQL = """
SELECT (SELECT array_agg(DISTINCT store ORDER BY store) FROM sheet_copy.noon_finance_api WHERE COALESCE(store, '') <> '') AS stores,
       MIN(CASE WHEN transaction_date ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}' THEN left(transaction_date, 10)::date END) AS dmin,
       MAX(CASE WHEN transaction_date ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}' THEN left(transaction_date, 10)::date END) AS dmax
FROM sheet_copy.noon_finance_api
"""


# ══════════════════════════════════════════════════════════════════════════════════════════
# نون: آخر أوردر / العمولة والتوصيل / ⏳ لسه ماظهرش — كلها محسوبة جوه PostgreSQL
# نفس منطق calc._noon_items و calc.acct_noon_com / acct_noon_last / acct_noon_pending بالظبط (أداة PG_Verify.py بتقارنهم على بياناتك).
# q(sql, params) = دالة بتنفّذ الاستعلام وترجّع (أسماء الأعمدة، الصفوف).
# ══════════════════════════════════════════════════════════════════════════════════════════
def _z(c):
    return "COALESCE(" + _nt(c) + ", 0.0)"


def _ts(col):
    """نص → timestamp (بدون وقت-زون) لو بيبدأ بتاريخ ISO، وإلا NULL (زي pd.to_datetime بالنسبة للصيغة اللي في الدفتر)."""
    return ("(CASE WHEN {c} ~ '^[0-9]{{4}}-[0-9]{{2}}-[0-9]{{2}}' THEN replace(left({c}, 19), 'T', ' ')::timestamp END)").format(c=col)


# بنقفل nested loop للاستعلامات دي: من غير إحصائيات (ANALYZE) البلانر بيغلط في عدد الصفوف ويلف على 100 ألف صف لكل صف
NOLOOP = "SET enable_nestloop = off;\n"

NOON_ITEMS_CTE = NOLOOP + """
WITH l AS (
  SELECT _row, btrim(COALESCE(store, '')) AS st,
         upper(btrim(split_part(COALESCE(skus, ''), ',', 1))) AS sku,
         btrim(COALESCE(order_nr, '')) AS on_, btrim(COALESCE(item_nr, '')) AS item,
         COALESCE(partner_skus, '') AS psku, COALESCE(title, '') AS prod,
         left(COALESCE(transaction_date, ''), 10) AS tdate, left(COALESCE(order_date, ''), 10) AS odate,
         """ + _z("net_proceeds") + """ AS net, """ + _z("referral_fee_including_vat") + """ AS ref,
         """ + _z("fullfilment_logistics_fees_including_vat") + """ AS ful, """ + _z("shipping_credits_including_vat") + """ AS shc,
         """ + _z("other_order_fees_including_vat") + """ AS oth, """ + _z("order_subsidies_including_vat") + """ AS sub
  FROM sheet_copy.noon_finance_api
  WHERE lower(btrim(COALESCE(transaction_type, ''))) IN ('order', 'order_update')
    AND (%(stores)s::text[] IS NULL OR store = ANY(%(stores)s))
),
f AS (SELECT *, (ful + shc + oth + sub) AS fee, st || '|' || on_ AS key FROM l),
ofee AS (SELECT key, SUM(fee) AS v FROM f WHERE sku = '' AND fee < 0 AND on_ <> '' GROUP BY key),
it0 AS (SELECT * FROM f WHERE sku <> '' AND net > 0 AND ref < 0),
nfin AS (SELECT key, COUNT(*) AS n FROM it0 GROUP BY key),
o AS (SELECT btrim(COALESCE(store, '')) || '|' || btrim(COALESCE(order_nr, '')) AS key, btrim(COALESCE(item_nr, '')) AS item,
             btrim(COALESCE(sku, '')) <> '' AS ws, lower(btrim(COALESCE(is_fulfilled_by_noon, ''))) = 'true' AS fb, _row
      FROM sheet_copy.noon_orders_api),
cnto AS (SELECT key, COUNT(*) FILTER (WHERE ws) AS n FROM o GROUP BY key),
imap AS (SELECT item, (array_agg(fb ORDER BY _row))[1] AS fb FROM o GROUP BY item),
omap AS (SELECT key, bool_and(fb) AS fb FROM o GROUP BY key HAVING bool_and(fb) = bool_or(fb)),
it AS (
  SELECT i._row, i.sku, i.psku, i.prod, i.net, -i.ref AS ref, (i.fee < 0) AS own, i.on_ AS ord, i.key, i.item, i.tdate, i.odate,
         (CASE WHEN i.fee < 0 THEN i.fee
               WHEN COALESCE(ofee.v, 0) <> 0 THEN ofee.v / GREATEST(COALESCE(CASE WHEN cnto.n > 0 THEN cnto.n END, nfin.n), 1)::double precision END) AS dlv0,
         COALESCE(imap.fb, omap.fb) AS fbv
  FROM it0 i
  LEFT JOIN ofee ON ofee.key = i.key
  LEFT JOIN nfin ON nfin.key = i.key
  LEFT JOIN cnto ON cnto.key = i.key
  LEFT JOIN imap ON imap.item = i.item
  LEFT JOIN omap ON omap.key = i.key
),
itm AS (
  SELECT *, -dlv0 AS dlv, (dlv0 IS NOT NULL) AS cp,
         (CASE WHEN fbv IS NULL THEN NULL WHEN fbv THEN 'FBN' ELSE 'FBB' END) AS model
  FROM it
),
d AS (
  SELECT * FROM itm
  WHERE (CASE WHEN tdate ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$' THEN tdate::date END) >= %(d_from)s
    AND (CASE WHEN tdate ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$' THEN tdate::date END) <= %(d_to)s
)
"""

NOON_COM_COUNTS_SQL = NOON_ITEMS_CTE + """
SELECT COUNT(*) FILTER (WHERE NOT cp) AS incomplete,
       COUNT(*) FILTER (WHERE cp AND model IS NULL) AS unknown
FROM d
"""

# channel_sql: '' = الكل · 'FBN' · 'FBB' (بيتحط جوه الاستعلام كنص ثابت — مش من المستخدم)
NOON_COM_SQL = NOON_ITEMS_CTE + """,
c AS (SELECT * FROM d WHERE cp {chan}),
sb AS (SELECT on_, sku, SUM(sub) AS v FROM f WHERE sku <> '' GROUP BY on_, sku),
cc AS (SELECT c.*, COUNT(*) OVER (PARTITION BY c.ord, c.sku) AS nsk FROM c)
SELECT cc.sku AS "SKU", (array_agg(cc.psku ORDER BY cc._row))[1] AS "Partner SKU", (array_agg(cc.prod ORDER BY cc._row))[1] AS "المنتج",
       COUNT(*)::double precision AS "قطع", SUM(cc.net) AS "مبيعات", SUM(cc.ref) AS "عمولة", SUM(cc.dlv) AS "توصيل", 0.0 AS "تانية",
       SUM(COALESCE(sb.v, 0) / cc.nsk) AS "_s", SUM(CASE WHEN COALESCE(sb.v, 0) / cc.nsk <> 0 THEN 1.0 ELSE 0.0 END) AS "_h"
FROM cc LEFT JOIN sb ON sb.on_ = cc.ord AND sb.sku = cc.sku
GROUP BY cc.sku
ORDER BY cc.sku COLLATE "C"
"""

NOON_LAST_SQL = NOON_ITEMS_CTE + """,
c AS (SELECT * FROM d WHERE TRUE {chan}),
pick AS (
  SELECT DISTINCT ON (sku) * FROM c
  ORDER BY sku, cp DESC, odate COLLATE "C" DESC, ord COLLATE "C" DESC, _row DESC
),
nsk AS (SELECT ord, sku, COUNT(*) AS n FROM c GROUP BY ord, sku),
nord AS (SELECT ord, COUNT(*) AS n FROM c WHERE NOT own GROUP BY ord),
isub AS (SELECT on_, sku, SUM(sub) AS v FROM f WHERE sku <> '' AND sub <> 0 GROUP BY on_, sku),
ifl AS (SELECT on_, sku, SUM(ful + shc + oth) AS v FROM f WHERE sku <> '' AND item <> '' GROUP BY on_, sku),
og AS (SELECT on_, SUM(sub) AS cr, SUM(ful + shc + oth) AS fl FROM f WHERE sku = '' AND on_ <> '' GROUP BY on_)
SELECT p.sku AS "SKU", p.psku AS "Partner SKU", p.prod AS "المنتج", p.ord AS "order", p.odate AS "date", p.model AS "model",
       p.net AS "net", p.ref AS "ref", p.dlv AS "dlv", p.own AS "own", p.cp AS "_cp",
       COALESCE(isub.v, 0) AS "isub_v", COALESCE(ifl.v, 0) AS "ifl_v", nsk.n AS "n_sku",
       COALESCE(og.cr, 0) AS "ocr_v", COALESCE(og.fl, 0) AS "ogfl_v", COALESCE(nord.n, 1) AS "n_ord"
FROM pick p
LEFT JOIN nsk ON nsk.ord = p.ord AND nsk.sku = p.sku
LEFT JOIN nord ON nord.ord = p.ord
LEFT JOIN isub ON isub.on_ = p.ord AND isub.sku = p.sku
LEFT JOIN ifl ON ifl.on_ = p.ord AND ifl.sku = p.sku
LEFT JOIN og ON og.on_ = p.ord
"""

NOON_TODAY_SQL = ("SELECT MAX(" + _ts("transaction_date") + ") FROM sheet_copy.noon_finance_api "
                  "WHERE (%(stores)s::text[] IS NULL OR store = ANY(%(stores)s))")

NOON_PENDING_SQL = NOLOOP + """
WITH seen AS (
  SELECT DISTINCT btrim(COALESCE(item_nr, '')) AS item FROM sheet_copy.noon_finance_api
  WHERE (%(stores)s::text[] IS NULL OR store = ANY(%(stores)s))
),
st AS (SELECT MIN(""" + _ts("transaction_date") + """) AS s FROM sheet_copy.noon_finance_api
       WHERE (%(stores)s::text[] IS NULL OR store = ANY(%(stores)s))),
o AS (
  SELECT order_nr, item_nr, sku, delivered_at, store AS "Store",
         """ + _ts("replace(COALESCE(delivered_at, ''), ' UTC', '')") + """ AS dd
  FROM sheet_copy.noon_orders_api
  WHERE lower(btrim(COALESCE(item_status, ''))) = 'delivered'
    AND (%(stores)s::text[] IS NULL OR store = ANY(%(stores)s))
    AND btrim(COALESCE(item_nr, '')) NOT IN (SELECT item FROM seen)
)
SELECT o.order_nr, o.item_nr, o.sku, o.delivered_at, o."Store",
       floor(extract(epoch FROM (%(today)s::timestamp - o.dd)) / 86400)::int AS days,
       (SELECT s FROM st) AS ledger_start
FROM o
WHERE (SELECT s FROM st) IS NULL OR o.dd >= (SELECT s FROM st)
ORDER BY o.dd NULLS LAST, o.order_nr
"""


def _stores(stores):
    return list(stores) if stores else None


def noon_com(q, stores, d_from, d_to, channel="all"):
    """(g, unknown, incomplete) — g جاهز لـ calc.acct_noon_com_finish."""
    import pandas as pd
    prm = {"stores": _stores(stores), "d_from": d_from, "d_to": d_to}
    chan = "" if channel == "all" else ("AND model = 'FBN'" if channel == "FBN" else "AND model = 'FBB'")
    c, r = q(NOON_COM_COUNTS_SQL, prm)
    incomplete, unknown = int(r[0][0] or 0), int(r[0][1] or 0)
    c, r = q(NOON_COM_SQL.replace("{chan}", chan), prm)
    g = pd.DataFrame(r, columns=c)
    for col in g.columns:
        if col not in ("SKU", "Partner SKU", "المنتج"):
            g[col] = pd.to_numeric(g[col]).astype(float)
    return g, (unknown if channel != "all" else 0), incomplete


def noon_last(q, stores, d_from, d_to, channel="all"):
    """(rows, today) — جاهز لـ calc.acct_noon_last_finish."""
    import pandas as pd
    prm = {"stores": _stores(stores), "d_from": d_from, "d_to": d_to}
    chan = "" if channel == "all" else ("AND model = 'FBN'" if channel == "FBN" else "AND model = 'FBB'")
    c, r = q(NOON_LAST_SQL.replace("{chan}", chan), prm)
    rows = pd.DataFrame(r, columns=c)
    for col in ("net", "ref", "dlv", "isub_v", "ifl_v", "ocr_v", "ogfl_v"):
        rows[col] = pd.to_numeric(rows[col]).astype(float)
    c, r = q(NOON_TODAY_SQL, {"stores": _stores(stores)})
    return rows, (pd.Timestamp(r[0][0]) if r and r[0][0] is not None else pd.NaT)


def noon_pending(q, stores, today=None):
    """نفس خرج calc.acct_noon_pending (الأعمدة اللي التطبيق بيعرضها) — attrs['ledger_start']."""
    import pandas as pd
    from datetime import datetime
    today = today or datetime.now().date()
    c, r = q(NOON_PENDING_SQL, {"stores": _stores(stores), "today": str(today)})
    o = pd.DataFrame(r, columns=c)
    start = None
    if len(o):
        start = o["ledger_start"].iloc[0]
    else:
        c2, r2 = q("SELECT MIN(" + _ts("transaction_date") + ") FROM sheet_copy.noon_finance_api WHERE (%(stores)s::text[] IS NULL OR store = ANY(%(stores)s))",
                   {"stores": _stores(stores)})
        start = r2[0][0] if r2 else None
    o = o.drop(columns=["ledger_start"]).rename(columns={"days": "أيام من التسليم"})
    o["الحالة"] = o["أيام من التسليم"].map(lambda d: "⚠️ متأخر" if d > 7 else "عادي (أقل من أسبوع)")
    o.attrs["ledger_start"] = start.date() if start is not None and pd.notna(start) else None
    return o
