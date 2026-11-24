# -*- coding: utf-8 -*-
"""
pg_sheet.py — «شيت» فوق PostgreSQL بنفس واجهة gspread اللي التطبيق بيستخدمها.
بيشتغل بس لما PG_MODE = "pg": التطبيق بيبدّل ss (الـ Spreadsheet) بـ PGSpreadsheet، فكل الكود التاني بيفضل زي ما هو.
  • كل تاب = جدول في schema sheet_copy (نفس جداول PG_Copy_All والبوتات) + صف في sheet_copy._meta فيه العناوين الأصلية.
  • تابات الصين (sheet_only) بتتحوّل للشيت الحقيقي (real_ss) — الموظف هناك بيكتب فيه.
  • كل كتابة = transaction واحدة بـ advisory lock على التاب: بتقرا الحالة، بتعدّلها، وبتحفظها (مفيش صفوف بتضيع لو جلستين كتبوا مع بعض).
"""
import csv
import io
import json
import re
import zlib

SCHEMA = "sheet_copy"

try:
    import gspread
    WorksheetNotFound = gspread.exceptions.WorksheetNotFound
except Exception:                                    # بيئة من غير gspread (اختبار)
    class WorksheetNotFound(Exception):
        pass


# ───────────────────────── أدوات ─────────────────────────
def _s(v):
    """قيمة → نص زي ما الشيت بيعرضها (get_all_values)."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _col_idx(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _col_letters(i):
    s = ""
    while i > 0:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def parse_a1(rng):
    """'A2:C9' / 'B3' / 'A:C' / "'Tab'!A1:B2" → (r1, c1, r2, c2) — None = من غير حد."""
    rng = str(rng).split("!")[-1].replace("$", "").strip()
    m = re.match(r"^([A-Za-z]*)(\d*)(?::([A-Za-z]*)(\d*))?$", rng)
    if not m:
        raise ValueError("A1 range مش مفهوم: %r" % rng)
    c1, r1, c2, r2 = m.group(1), m.group(2), m.group(3), m.group(4)
    if c2 is None and r2 is None:                    # خلية واحدة
        c2, r2 = c1, r1
    return (int(r1) if r1 else None, _col_idx(c1) if c1 else None,
            int(r2) if r2 else None, _col_idx(c2) if c2 else None)


def snake_names(headers):
    used, out = set(), []
    for h in headers:
        s_ = re.sub(r"[^0-9a-zA-Z]+", "_", str(h)).strip("_").lower()
        if not s_:
            s_ = "c_%d" % (len(used) + 1)
        if s_[0].isdigit():
            s_ = "c_" + s_
        s_ = s_[:60]
        base, i = s_, 2
        while s_ in used:
            s_, i = "%s_%d" % (base, i), i + 1
        used.add(s_)
        out.append(s_)
    return out


class _Cell:
    def __init__(self, row, col, value):
        self.row, self.col, self.value = row, col, value


# ───────────────────────── الحالة (عناوين + صفوف) ─────────────────────────
class _State:
    def __init__(self, hdr=None, rows=None, table=None, exists=False):
        self.hdr = list(hdr or [])
        self.rows = [list(r) for r in (rows or [])]
        self.table = table
        self.exists = exists
        self.changed = False

    def ensure_cols(self, n):
        while len(self.hdr) < n:
            self.hdr.append("")
            for r in self.rows:
                r.append("")
            self.changed = True

    def ensure_rows(self, n):                       # n = عدد صفوف البيانات (من غير الهيدر)
        while len(self.rows) < n:
            self.rows.append([""] * len(self.hdr))
            self.changed = True

    def set(self, r, c, v):
        v = _s(v)
        self.ensure_cols(c)
        if r == 1:
            if self.hdr[c - 1] != v:
                self.hdr[c - 1] = v
                self.changed = True
        else:
            self.ensure_rows(r - 1)
            if self.rows[r - 2][c - 1] != v:
                self.rows[r - 2][c - 1] = v
                self.changed = True

    def get(self, r, c):
        if r == 1:
            return self.hdr[c - 1] if c - 1 < len(self.hdr) else ""
        if r - 2 < len(self.rows) and c - 1 < len(self.hdr):
            return self.rows[r - 2][c - 1]
        return ""

    def put_block(self, r0, c0, values):
        for i, row in enumerate(values):
            for j, v in enumerate(row):
                self.set(r0 + i, c0 + j, v)

    def all_values(self):
        out = [list(self.hdr)] + [list(r) for r in self.rows] if (self.hdr or self.rows) else []
        while out and not any(str(c).strip() for c in out[-1]):
            out.pop()
        return out

    def trim(self):
        while self.rows and not any(str(c).strip() for c in self.rows[-1]):
            self.rows.pop()


# ───────────────────────── Worksheet ─────────────────────────
class PGWorksheet:
    def __init__(self, ssn, title):
        self._ss = ssn
        self.title = title
        self.spreadsheet = ssn
        self.id = zlib.crc32(title.encode("utf-8")) % 1000000000
        self._cnt = (0, 0)

    # --- داخلي ---
    def _tx(self, fn):
        conn = self._ss._conn()
        try:
            cur = conn.cursor()
            cur.execute("SELECT pg_advisory_xact_lock(%s)", (zlib.crc32(self.title.encode("utf-8")),))
            st = self._ss._load(cur, self.title)
            res = fn(st)
            if st.changed:
                st.trim()
                self._ss._store(cur, self.title, st)
            conn.commit()
            self._cnt = (0, 0)
            return res
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _read(self):
        conn = self._ss._conn()
        try:
            return self._ss._load(conn.cursor(), self.title)
        finally:
            conn.close()

    # --- خصائص ---
    @property
    def row_count(self):
        import time
        if self._cnt[0] and time.time() - self._cnt[1] < 30:
            return self._cnt[0]
        st = self._read()
        n = len(st.rows) + 1 + 100
        self._cnt = (n, time.time())
        return n

    @property
    def col_count(self):
        return max(len(self._read().hdr), 1) + 5

    @property
    def url(self):
        return "pg://%s" % self.title

    # --- قراءة ---
    def get_all_values(self, *a, **k):
        return self._read().all_values()

    def get(self, range_name=None, *a, **k):
        vals = self._read().all_values()
        if not range_name:
            return vals
        r1, c1, r2, c2 = parse_a1(range_name)
        r1, c1 = r1 or 1, c1 or 1
        r2 = r2 or len(vals)
        out = []
        for r in range(r1, min(r2, len(vals)) + 1):
            row = vals[r - 1]
            part = row[c1 - 1:(c2 or len(row))]
            out.append(part)
        while out and not any(str(c).strip() for c in out[-1]):
            out.pop()
        return out

    def row_values(self, i, *a, **k):
        vals = self._read().all_values()
        row = list(vals[i - 1]) if 0 < i <= len(vals) else []
        while row and str(row[-1]) == "":
            row.pop()
        return row

    def col_values(self, i, *a, **k):
        vals = self._read().all_values()
        out = [(r[i - 1] if i - 1 < len(r) else "") for r in vals]
        while out and out[-1] == "":
            out.pop()
        return out

    def cell(self, row, col, *a, **k):
        return _Cell(row, col, self._read().get(row, col))

    def acell(self, a1, *a, **k):
        r, c, _, _ = parse_a1(a1)
        return self.cell(r, c)

    # --- كتابة ---
    def append_row(self, values, *a, **k):
        return self.append_rows([values])

    def append_rows(self, rows, *a, **k):
        rows = [[_s(v) for v in r] for r in rows]
        if not rows:
            return

        def f(st):
            for vals in rows:
                if not st.hdr and not st.rows:        # تاب فاضي: أول صف = الهيدر (زي الشيت)
                    st.hdr = list(vals)
                    st.changed = True
                    continue
                st.ensure_cols(len(vals))
                st.rows.append(list(vals) + [""] * (len(st.hdr) - len(vals)))
                st.changed = True
        return self._tx(f)

    def append_cols(self, values, *a, **k):
        def f(st):
            for col in values:
                c = len(st.hdr) + 1
                for i, v in enumerate(col):
                    st.set(i + 1, c, v)
        return self._tx(f)

    def update(self, *args, **kw):
        values, rng = kw.get("values"), kw.get("range_name")
        for x in args:                                # update("A1", [[..]]) أو update([[..]], "A1")
            if isinstance(x, str):
                rng = x
            elif values is None:
                values = x
        if values is None:
            return
        r1, c1, _, _ = parse_a1(rng or "A1")

        def f(st):
            st.put_block(r1 or 1, c1 or 1, values)
        return self._tx(f)

    def update_cell(self, row, col, value, *a, **k):
        return self._tx(lambda st: st.set(row, col, value))

    def update_acell(self, a1, value, *a, **k):
        r, c, _, _ = parse_a1(a1)
        return self.update_cell(r, c, value)

    def batch_update(self, data, *a, **k):
        def f(st):
            for d in data:
                r1, c1, _, _ = parse_a1(d["range"])
                st.put_block(r1 or 1, c1 or 1, d["values"])
        return self._tx(f)

    def delete_rows(self, start, end=None):
        end = end or start

        def f(st):
            allr = [st.hdr] + st.rows
            del allr[start - 1:end]
            st.hdr, st.rows = (list(allr[0]), allr[1:]) if allr else ([], [])
            st.changed = True
        return self._tx(f)

    def delete_row(self, i):
        return self.delete_rows(i, i)

    def clear(self):
        def f(st):
            st.hdr, st.rows = [], []
            st.changed = True
        return self._tx(f)

    def batch_clear(self, ranges):
        def f(st):
            for rng in ranges:
                r1, c1, r2, c2 = parse_a1(rng)
                r1, c1 = r1 or 1, c1 or 1
                r2 = r2 or (len(st.rows) + 1)
                c2 = c2 or len(st.hdr)
                for r in range(r1, min(r2, len(st.rows) + 1) + 1):
                    for c in range(c1, min(c2, len(st.hdr)) + 1):
                        st.set(r, c, "")
        return self._tx(f)

    def update_title(self, new_title):
        self._ss._rename(self.title, new_title)
        self.title = new_title

    # تنسيق/حجم: مالهاش معنى في PostgreSQL
    def add_rows(self, *a, **k):
        return None

    def add_cols(self, *a, **k):
        return None

    def resize(self, *a, **k):
        return None

    def format(self, *a, **k):
        return None

    def freeze(self, *a, **k):
        return None

    def clear_basic_filter(self, *a, **k):
        return None


# ───────────────────────── Spreadsheet ─────────────────────────
class PGSpreadsheet:
    """connect = دالة بترجّع اتصال psycopg2 جديد · real_ss = الـ Spreadsheet الحقيقي (للصين) · sheet_only(title) = تاب ده يفضل في الشيت."""

    def __init__(self, connect, real_ss=None, sheet_only=lambda t: False):
        self._connect = connect
        self.real = real_ss
        self._sheet_only = sheet_only
        self.id = getattr(real_ss, "id", "pg")
        self._ready = False

    def _conn(self):
        c = self._connect()
        if not self._ready:
            cur = c.cursor()
            cur.execute("CREATE SCHEMA IF NOT EXISTS %s" % SCHEMA)
            cur.execute("CREATE TABLE IF NOT EXISTS %s._meta (table_name TEXT PRIMARY KEY, tab_title TEXT, headers JSONB, col_names JSONB, "
                        "copied_at TIMESTAMPTZ DEFAULT now())" % SCHEMA)
            c.commit()
            self._ready = True
        return c

    # --- تحميل/حفظ ---
    def _load(self, cur, title):
        cur.execute("SELECT table_name, headers, col_names FROM %s._meta WHERE tab_title = %%s" % SCHEMA, (title,))
        m = cur.fetchone()
        if not m:
            return _State()
        table, hdr, cols = m
        rows = []
        if cols:
            cur.execute("SELECT " + ", ".join('"%s"' % c for c in cols) + ' FROM %s."%s" ORDER BY _row' % (SCHEMA, table))
            rows = [["" if v is None else v for v in r] for r in cur.fetchall()]
        return _State(hdr, rows, table, True)

    def _new_table_name(self, cur, title):
        base = re.sub(r"[^0-9a-zA-Z]+", "_", title).strip("_").lower()[:55] or ("t_%d" % (zlib.crc32(title.encode("utf-8")) % 1000000))
        if base[0].isdigit():
            base = "t_" + base
        cur.execute("SELECT 1 FROM %s._meta WHERE table_name = %%s" % SCHEMA, (base,))
        return base if not cur.fetchone() else "%s_%d" % (base[:50], zlib.crc32(title.encode("utf-8")) % 1000000)

    def _store(self, cur, title, st):
        table = st.table or self._new_table_name(cur, title)
        hdr = [str(h) for h in st.hdr]
        w = len(hdr)
        while w and not hdr[w - 1]:
            w -= 1
        hdr = hdr[:w]
        cols = snake_names(hdr)
        cur.execute('DROP TABLE IF EXISTS %s."%s"' % (SCHEMA, table))
        cur.execute('CREATE TABLE %s."%s" (' % (SCHEMA, table) + ", ".join('"%s" TEXT' % c for c in cols) + (", " if cols else "") + "_row BIGSERIAL)")
        cur.execute("INSERT INTO %s._meta (table_name, tab_title, headers, col_names, copied_at) VALUES (%%s,%%s,%%s,%%s,now()) "
                    "ON CONFLICT (table_name) DO UPDATE SET tab_title=EXCLUDED.tab_title, headers=EXCLUDED.headers, col_names=EXCLUDED.col_names, "
                    "copied_at=now()" % SCHEMA, (table, title, json.dumps(hdr, ensure_ascii=False), json.dumps(cols)))
        st.table, st.exists = table, True
        if cols and st.rows:
            buf = io.StringIO()
            wr = csv.writer(buf)
            for r in st.rows:
                wr.writerow((list(r) + [""] * w)[:w])
            buf.seek(0)
            cur.copy_expert('COPY %s."%s" (' % (SCHEMA, table) + ", ".join('"%s"' % c for c in cols) + ") FROM STDIN WITH CSV", buf)

    def _rename(self, old, new):
        conn = self._conn()
        try:
            cur = conn.cursor()
            cur.execute("UPDATE %s._meta SET tab_title = %%s WHERE tab_title = %%s" % SCHEMA, (new, old))
            conn.commit()
        finally:
            conn.close()

    def _titles(self):
        conn = self._conn()
        try:
            cur = conn.cursor()
            cur.execute("SELECT tab_title FROM %s._meta ORDER BY tab_title" % SCHEMA)
            return [r[0] for r in cur.fetchall() if r[0] and not self._sheet_only(r[0])]
        finally:
            conn.close()

    # --- واجهة gspread ---
    def worksheets(self):
        out = [PGWorksheet(self, t) for t in self._titles()]
        if self.real is not None:
            try:
                out += [w for w in self.real.worksheets() if self._sheet_only(w.title)]
            except Exception:
                pass
        return out

    def worksheet(self, title):
        if self._sheet_only(title):
            if self.real is None:
                raise WorksheetNotFound(title)
            return self.real.worksheet(title)
        if title in self._titles():
            return PGWorksheet(self, title)
        raise WorksheetNotFound(title)

    def add_worksheet(self, title, rows=None, cols=None, *a, **k):
        if self._sheet_only(title) and self.real is not None:
            return self.real.add_worksheet(title=title, rows=rows, cols=cols)
        ws = PGWorksheet(self, title)
        conn = self._conn()
        try:
            cur = conn.cursor()
            cur.execute("SELECT pg_advisory_xact_lock(%s)", (zlib.crc32(title.encode("utf-8")),))
            st = self._load(cur, title)
            if not st.exists:
                self._store(cur, title, st)
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
        return ws

    def del_worksheet(self, ws):
        title = getattr(ws, "title", ws)
        if self._sheet_only(title) and self.real is not None:
            return self.real.del_worksheet(ws)
        conn = self._conn()
        try:
            cur = conn.cursor()
            cur.execute("SELECT table_name FROM %s._meta WHERE tab_title = %%s" % SCHEMA, (title,))
            m = cur.fetchone()
            if m:
                cur.execute('DROP TABLE IF EXISTS %s."%s"' % (SCHEMA, m[0]))
                cur.execute("DELETE FROM %s._meta WHERE table_name = %%s" % SCHEMA, (m[0],))
            conn.commit()
        finally:
            conn.close()

    def values_batch_get(self, ranges, *a, **k):
        out = []
        for rg in ranges:
            m = re.match(r"^'((?:[^']|'')*)'(?:!(.*))?$", rg)
            title = m.group(1).replace("''", "'") if m else rg
            if self._sheet_only(title) and self.real is not None:
                out.append(self.real.values_batch_get([rg])["valueRanges"][0])
                continue
            vals = PGWorksheet(self, title).get_all_values()
            out.append({"range": "'%s'!A1:%s%d" % (title, _col_letters(max((len(r) for r in vals), default=1)), max(len(vals), 1)),
                        "majorDimension": "ROWS", "values": vals})
        return {"spreadsheetId": self.id, "valueRanges": out}

    def batch_update(self, body, *a, **k):
        if self.real is not None:
            return self.real.batch_update(body)       # تنسيق شيت الصين (الباقي مالوش معنى في PostgreSQL)
        return {}

    def fetch_sheet_metadata(self, *a, **k):
        return self.real.fetch_sheet_metadata(*a, **k) if self.real is not None else {"sheets": []}

    @property
    def url(self):
        return getattr(self.real, "url", "pg://")
