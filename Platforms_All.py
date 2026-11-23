# -*- coding: utf-8 -*-
"""
Platforms_Server.py — برنامج واحد لبوتات نون + أمازون + تريندول.

  • كل بوت في تاب (بنفس إعداداته وجدولته وتفاصيله بالظبط) — كود البوتات التلاتة جوه الملف ده (آخره) وبيتحمّل كل بوت في مساحة أسماء منفصلة
  • قفل عام: بوت واحد بس بيشغّل كروم في نفس الوقت. لو بوت تاني جه ميعاده بيستنى دوره (مفيش تعارض بين نوافذ كروم)
  • فتح كروم نفسه (تجهيز chromedriver) بيتم تحت قفل واحد
  • سجل مشترك تحت لكل اللي بيحصل (ملوّن حسب البوت) — بتسحب الفاصل بالماوس تكبّره/تصغّره، أو زرار «تكبير السجل»
  • النافذة بتتكبّر وتتصغّر براحتك، وكل تاب فيه شريط تمرير لو الشاشة صغيرة
"""
import os
import sys
import time
import queue
import threading
import traceback
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, simpledialog, filedialog   # simpledialog/filedialog يستخدمهم كود البوتات المدمج (لازم يتحزّموا في الـ EXE)

# ───── مجلد البرنامج (للإعدادات/التنزيلات) — بيتحدد قبل ما نستورد البوتات ─────
BASE_DIR = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
os.environ["PLATFORMS_BASE_DIR"] = BASE_DIR
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

# مكتبات البوتات (عشان PyInstaller يلاقيها لما الكود المدمج يتنفّذ)
try:
    import collections  # noqa: F401
except Exception:
    pass
try:
    import csv  # noqa: F401
except Exception:
    pass
try:
    import datetime  # noqa: F401
except Exception:
    pass
try:
    import email  # noqa: F401
except Exception:
    pass
try:
    import glob  # noqa: F401
except Exception:
    pass
try:
    import gspread  # noqa: F401
except Exception:
    pass
try:
    import gspread.utils  # noqa: F401
except Exception:
    pass
try:
    import imaplib  # noqa: F401
except Exception:
    pass
try:
    import json  # noqa: F401
except Exception:
    pass
try:
    import oauth2client.service_account  # noqa: F401
except Exception:
    pass
try:
    import openpyxl  # noqa: F401
except Exception:
    pass
try:
    import openpyxl.styles  # noqa: F401
except Exception:
    pass
try:
    import os  # noqa: F401
except Exception:
    pass
try:
    import pythoncom  # noqa: F401
except Exception:
    pass
try:
    import random  # noqa: F401
except Exception:
    pass
try:
    import re  # noqa: F401
except Exception:
    pass
try:
    import selenium  # noqa: F401
except Exception:
    pass
try:
    import selenium.common.exceptions  # noqa: F401
except Exception:
    pass
try:
    import selenium.webdriver.chrome.service  # noqa: F401
except Exception:
    pass
try:
    import selenium.webdriver.common.action_chains  # noqa: F401
except Exception:
    pass
try:
    import selenium.webdriver.common.by  # noqa: F401
except Exception:
    pass
try:
    import selenium.webdriver.common.keys  # noqa: F401
except Exception:
    pass
try:
    import selenium.webdriver.support  # noqa: F401
except Exception:
    pass
try:
    import selenium.webdriver.support.ui  # noqa: F401
except Exception:
    pass
try:
    import shutil  # noqa: F401
except Exception:
    pass
try:
    import sys  # noqa: F401
except Exception:
    pass
try:
    import tempfile  # noqa: F401
except Exception:
    pass
try:
    import threading  # noqa: F401
except Exception:
    pass
try:
    import time  # noqa: F401
except Exception:
    pass
try:
    import traceback  # noqa: F401
except Exception:
    pass
try:
    import urllib.parse  # noqa: F401
except Exception:
    pass
try:
    import urllib.request  # noqa: F401
except Exception:
    pass
try:
    import webdriver_manager.chrome  # noqa: F401
except Exception:
    pass
try:
    import win32com.client  # noqa: F401
except Exception:
    pass
try:
    import requests  # noqa: F401
    import jwt  # noqa: F401
    import cryptography  # noqa: F401
    import oauth2client.service_account  # noqa: F401
    import concurrent.futures  # noqa: F401
except Exception:
    pass
try:
    import zipfile  # noqa: F401
except Exception:
    pass

BG, CARD, TEXT, MUTED = "#0f1923", "#1a2535", "#e8edf3", "#7a8fa6"
GREEN, RED, ENTRY_BG, ACCENT = "#2ecc71", "#e74c3c", "#243044", "#f5a623"

# (المفتاح, العنوان, اسم الملف, لون السجل)
BOTS = [
    ("noon",   "🟡 Noon",     "NOONServer",      "#f5d000"),
    ("amazon", "🟠 Amazon",   "Amazon_Server",   "#ff9900"),
    ("ty",     "🔶 Trendyol", "Trendyol_Server", "#f27a1a"),
    ("noonapi", "🔵 Noon API", "NoonAPI",       "#4aa3ff"),
    ("tyapi",  "🟣 Trendyol API", "TrendyolAPI", "#b266ff"),
    ("noonacct", "🏦 حسابات نون", "NoonAccounts", "#2ecc71"),
    ("tyacct", "🏦 حسابات تريندول", "TrendyolAccounts", "#c77dff"),
    ("amzacct", "🏦 حسابات أمازون", "AmazonAccounts", "#ffb347"),
    ("nttrack", "🟢 منافسين نون", "NoonTracker", "#3ddc97"),
]

RUN_GATE = threading.Lock()                 # بوت واحد بس بيشغّل كروم في نفس الوقت
CHROME_START_LOCK = threading.RLock()       # فتح كروم/تجهيز chromedriver مرة واحدة في كل وقت
_STATE = {"running": "", "waiting": []}     # للعرض في شريط الحالة
LOG_HIST = []                                # (وقت، بوت، سطر) — بيتستخدم لملخص السلسلة
LOG_Q = queue.Queue()                       # (bot_key, text) من كل الـ Threads → الواجهة
MODS = {}                                   # key → module



# ═══════════════ تحميل البوتات المدمجة (كل بوت في موديول منفصل → مفيش تعارض أسماء) ═══════════════
def _load_embedded(name):
    import types
    mod = types.ModuleType(name)
    mod.__file__ = os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__)
    sys.modules[name] = mod
    exec(compile(_EMBEDDED[name], f"<{name}>", "exec"), mod.__dict__)
    return mod

# ═══════════════ ربط البوتات: قفل التشغيل + السجل المشترك ═══════════════
def _gated(fn, key, title, log_arg_index=None):
    """بيغلّف نقطة دخول التشغيل (do_run / do_run_reports): لو بوت تاني شغّال بيستنى دوره بدل ما يفتح كروم فوقيه."""
    def wrapper(*args, **kwargs):
        log = None
        if log_arg_index is not None and len(args) > log_arg_index:
            log = args[log_arg_index]
        if not RUN_GATE.acquire(blocking=False):
            who = _STATE["running"] or "بوت تاني"
            LOG_Q.put((key, f"⏳ {title}: مستني دور {who} يخلص (عشان كروم ما يتعارضش)...\n"))
            _STATE["waiting"].append(title)
            try:
                RUN_GATE.acquire()
            finally:
                try:
                    _STATE["waiting"].remove(title)
                except ValueError:
                    pass
        _STATE["running"] = title
        LOG_Q.put((key, f"▶ {title}: بدأ التشغيل\n"))
        try:
            return fn(*args, **kwargs)
        finally:
            _STATE["running"] = ""
            LOG_Q.put((key, f"■ {title}: خلّص\n"))
            RUN_GATE.release()
    wrapper.__name__ = getattr(fn, "__name__", "gated")
    return wrapper


def _locked(fn):
    def wrapper(*args, **kwargs):
        with CHROME_START_LOCK:
            return fn(*args, **kwargs)
    wrapper.__name__ = getattr(fn, "__name__", "locked")
    return wrapper


def _mirror_log(mod, key):
    """كل log_write في البوت بيتكتب في سجله + السجل المشترك."""
    orig = mod.log_write

    def lw(log, text):
        try:
            orig(log, text)
        finally:
            try:
                LOG_Q.put((key, str(text)))
            except Exception:
                pass
    mod.log_write = lw


def hook_module(key, title, mod):
    if key in ("noonapi", "tyapi", "noonacct", "tyacct", "amzacct"):
        # أدوات الـ API: مش بتفتح كروم → من غير قفل؛ سجلها بيروح للسجل المشترك
        mod.LOG_HOOK = lambda t: LOG_Q.put((key, t))
        return
    _mirror_log(mod, key)
    if key == "noon":
        mod.do_run_reports = _gated(mod.do_run_reports, key, title, log_arg_index=3)
    else:
        mod.do_run = _gated(mod.do_run, key, title, log_arg_index=1)
    if hasattr(mod, "start_driver"):
        mod.start_driver = _locked(mod.start_driver)
    if hasattr(mod, "EXT_GATE"):
        mod.EXT_GATE = RUN_GATE


# ═══════════════ تاب قابل للتمرير ═══════════════
class ScrollTab(tk.Frame):
    """إطار فيه شريط تمرير عمودي — محتوى البوت بيتبني جوه self.inner."""
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        self.canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self.canvas, bg=BG)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self._on_canvas)

    def _on_canvas(self, e):
        # المحتوى بيملا العرض، وبيفضل على الأقل بطول الشاشة (عشان اللي بـ expand يتمدد)
        self.canvas.itemconfigure(self._win, width=e.width)
        self.canvas.itemconfigure(self._win, height=max(e.height, self.inner.winfo_reqheight()))


def _install_wheel(root):
    """عجلة الماوس بتمرّر التاب اللي تحت المؤشر (إلا لو المؤشر فوق قايمة/نص ليهم تمريرهم)."""
    def target_canvas(w):
        while w is not None:
            if isinstance(w, ScrollTab):
                return w.canvas
            w = getattr(w, "master", None)
        return None

    def on_wheel(e):
        w = root.winfo_containing(e.x_root, e.y_root)
        if w is None or isinstance(w, (tk.Text, tk.Listbox, ttk.Treeview, ttk.Combobox)):
            return
        c = target_canvas(w)
        if c is None:
            return
        first, last = c.yview()
        if first <= 0.0 and last >= 1.0:
            return
        c.yview_scroll(int(-e.delta / 120) if e.delta else (-1 if e.num == 4 else 1), "units")
    root.bind_all("<MouseWheel>", on_wheel)
    root.bind_all("<Button-4>", on_wheel)
    root.bind_all("<Button-5>", on_wheel)


# ═══════════════ السجل المشترك ═══════════════
class UnifiedLog(tk.Frame):
    MAX_LINES = 6000

    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        self.filter = "all"
        self.autoscroll = tk.BooleanVar(value=True)
        self.font_size = 10
        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=8, pady=(6, 2))
        tk.Label(bar, text="📋 كل اللي بيحصل", bg=BG, fg=MUTED, font=("Segoe UI", 10, "bold")).pack(side="left")
        self.fbtns = {}
        for k, t in [("all", "الكل")] + [(b[0], b[1]) for b in BOTS]:
            b = tk.Button(bar, text=t, relief="flat", bg=CARD, fg=TEXT, font=("Segoe UI", 9), padx=8, cursor="hand2",
                          command=lambda k=k: self.set_filter(k))
            b.pack(side="left", padx=(8 if k == "all" else 2, 0))
            self.fbtns[k] = b
        for txt, cmd in (("A−", lambda: self.zoom(-1)), ("A+", lambda: self.zoom(+1)), ("🧹 مسح", self.clear)):
            tk.Button(bar, text=txt, relief="flat", bg=CARD, fg=TEXT, font=("Segoe UI", 9), padx=8, cursor="hand2",
                      command=cmd).pack(side="right", padx=2)
        tk.Checkbutton(bar, text="تمرير تلقائي", variable=self.autoscroll, bg=BG, fg=MUTED, selectcolor=ENTRY_BG,
                       activebackground=BG, activeforeground=TEXT, font=("Segoe UI", 9)).pack(side="right", padx=6)
        self.text = scrolledtext.ScrolledText(self, bg=ENTRY_BG, fg=TEXT, font=("Consolas", self.font_size), relief="flat",
                                              insertbackground=ACCENT, height=8, wrap="word")
        self.text.pack(fill="both", expand=True, padx=8, pady=(0, 6))
        for k, _t, _f, col in BOTS:
            self.text.tag_configure("tag_" + k, foreground=col)
        self.text.tag_configure("sys", foreground=MUTED)
        self.text.bind("<Control-MouseWheel>", self._ctrl_wheel)
        self._refresh_btns()

    def _ctrl_wheel(self, e):
        self.zoom(+1 if e.delta > 0 else -1)
        return "break"

    def zoom(self, d):
        self.font_size = max(7, min(22, self.font_size + d))
        self.text.configure(font=("Consolas", self.font_size))

    def clear(self):
        self.text.delete("1.0", "end")

    def set_filter(self, k):
        self.filter = k
        for bk, _t, _f, _c in BOTS:
            self.text.tag_configure("tag_" + bk, elide=(k != "all" and k != bk))
        self._refresh_btns()
        if self.autoscroll.get():
            self.text.see("end")

    def _refresh_btns(self):
        for k, b in self.fbtns.items():
            b.configure(bg=ACCENT if k == self.filter else CARD, fg="#0f1923" if k == self.filter else TEXT)

    def add(self, key, text):
        keys = {b[0] for b in BOTS}
        label = {b[0]: b[1].split(" ", 1)[1] for b in BOTS}.get(key, key)
        tag = "tag_" + key if key in keys else "sys"
        stamp = time.strftime("%H:%M:%S")
        for line in str(text).replace("\r", "").split("\n"):
            if line.strip():
                LOG_HIST.append((time.time(), key, line.strip()))
                if len(LOG_HIST) > 30000:
                    del LOG_HIST[:5000]
                self.text.insert("end", f"{stamp} [{label}] {line}\n", (tag,))
        n = int(self.text.index("end-1c").split(".")[0])
        if n > self.MAX_LINES:
            self.text.delete("1.0", f"{n - self.MAX_LINES}.0")
        if self.autoscroll.get():
            self.text.see("end")


# ═══════════════ البرنامج ═══════════════
# ═══════════════ ▶ تشغيل نون بالترتيب (سلسلة) ═══════════════
INSTS = {}   # key → كائن الواجهة (App) للبوتات اللي ليها App
CHAIN_SETTINGS = os.path.join(BASE_DIR, "noon_chain_settings.json")


class _LogSink:
    """بديل لصندوق السجل: log_write(log, text) بتنادي after/insert/see/update → بنوجّه النص للسجل المشترك."""
    def __init__(self, key):
        self.key = key

    def after(self, _ms, fn):
        try:
            fn()
        except Exception:
            pass

    def insert(self, _idx, text):
        pass     # النص بيروح للسجل المشترك من _mirror_log

    def see(self, *_a):
        pass

    def update(self):
        pass


def build_chain_ui(parent, root):
    steps = [("noon", "1) 🟡 Noon (إعلانات + تكويد + باقي التقارير)"),
             ("noonapi", "2) 🔵 Noon API (أوردرات + مخزون + عمولات)"),
             ("noonacct", "3) 🏦 حسابات نون (مالية + أوردرات → الشيت)"),
             ("tyapi", "4) 🟣 Trendyol API (أوردرات + منتجات + مالية خام)"),
             ("tyacct", "5) 🏦 حسابات تريندول (مالية → الشيت)"),
             ("amzacct", "6) 🏦 حسابات أمازون (مالية + شحنات + مخزون → الشيت)")]
    var = {k: tk.BooleanVar(value=True) for k, _ in steps}
    auto = tk.BooleanVar(value=False)
    times = tk.StringVar(value="07:00")
    stop_ev = threading.Event()
    state = {"running": False, "last_key": ""}
    lbls = {}

    wrap = tk.Frame(parent, bg=BG)
    wrap.pack(fill="both", expand=True, padx=14, pady=12)
    tk.Label(wrap, text="▶ تشغيل بالترتيب (نون + تريندول)", bg=BG, fg=ACCENT, font=("Segoe UI", 16, "bold")).pack(anchor="w")
    tk.Label(wrap, text="بيشغّل الخطوات واحدة ورا التانية. لو خطوة فشلت بيكمّل للي بعدها ويكتب ده في السجل.",
             bg=BG, fg=MUTED, font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 10))
    card = tk.Frame(wrap, bg=CARD, padx=16, pady=12)
    card.pack(fill="x")
    for k, t in steps:
        row = tk.Frame(card, bg=CARD)
        row.pack(fill="x", pady=3)
        tk.Checkbutton(row, text=t, variable=var[k], bg=CARD, fg=TEXT, selectcolor=ENTRY_BG, activebackground=CARD,
                       activeforeground=TEXT, font=("Segoe UI", 11), command=lambda: save()).pack(side="left")
        lbls[k] = tk.Label(row, text="—", bg=CARD, fg=MUTED, font=("Segoe UI", 10, "bold"))
        lbls[k].pack(side="right")
        if k == "noonapi":
            prow = tk.Frame(card, bg=CARD)
            prow.pack(fill="x", pady=(0, 3))
            tk.Label(prow, text="        ↳ مكان الكتابة (كل البوتات):", bg=CARD, fg=MUTED, font=("Segoe UI", 10)).pack(side="left")
            _PGL = ["📄 Google Sheet فقط", "🐘 PostgreSQL فقط", "📄+🐘 الاتنين مع بعض"]
            _pm = {"sheet": _PGL[0], "pg": _PGL[1], "both": _PGL[2]}
            try:
                _cur = open(os.path.join(BASE_DIR, "noon_use_pg.txt"), encoding="utf-8").read().strip()
            except Exception:
                _cur = "sheet"
            pgv = tk.StringVar(value=_pm.get(_cur, _PGL[0]))
            _cb = ttk.Combobox(prow, textvariable=pgv, values=_PGL, state="readonly", width=24)
            _cb.pack(side="left", padx=8)

            def _pg_pick(_e=None):
                inst = INSTS.get("noonapi")
                if inst is not None and hasattr(inst, "pg_mode"):
                    inst.pg_mode.set(pgv.get())
                else:
                    try:
                        open(os.path.join(BASE_DIR, "noon_use_pg.txt"), "w", encoding="utf-8").write({v: k2 for k2, v in _pm.items()}[pgv.get()])
                    except Exception:
                        pass
            _cb.bind("<<ComboboxSelected>>", _pg_pick)

            def _pg_sync():
                try:
                    inst = INSTS.get("noonapi")
                    if inst is not None and hasattr(inst, "pg_mode") and pgv.get() != inst.pg_mode.get():
                        pgv.set(inst.pg_mode.get())
                except Exception:
                    pass
                root.after(2000, _pg_sync)
            root.after(2000, _pg_sync)
            srow2 = tk.Frame(card, bg=CARD)
            srow2.pack(fill="x", pady=(0, 3))
            stv = tk.BooleanVar(value=True)
            tk.Checkbutton(srow2, text="        ↳ 🏬 سحب رسوم التخزين الشهرية (Noon_Storage_API)", variable=stv, bg=CARD, fg=TEXT,
                           selectcolor=ENTRY_BG, activebackground=CARD, activeforeground=TEXT, font=("Segoe UI", 10),
                           command=lambda: _st_pick()).pack(side="left")

            def _st_pick():
                inst = INSTS.get("noonapi")
                if inst is not None and hasattr(inst, "storage"):
                    inst.storage.set(bool(stv.get()))

            def _st_sync():
                try:
                    inst = INSTS.get("noonapi")
                    if inst is not None and hasattr(inst, "storage") and bool(stv.get()) != bool(inst.storage.get()):
                        stv.set(bool(inst.storage.get()))
                except Exception:
                    pass
                root.after(2000, _st_sync)
            root.after(2000, _st_sync)

    srow = tk.Frame(wrap, bg=BG)
    srow.pack(fill="x", pady=10)
    tk.Checkbutton(srow, text="⏰ تشغيل تلقائي كل يوم الساعة (HH:MM — لو أكتر من موعد افصلهم بفاصلة):", variable=auto, bg=BG, fg=TEXT,
                   selectcolor=ENTRY_BG, activebackground=BG, activeforeground=TEXT, font=("Segoe UI", 10),
                   command=lambda: on_auto()).pack(side="left")
    tk.Entry(srow, textvariable=times, width=16, bg=ENTRY_BG, fg=TEXT, insertbackground=TEXT, relief="flat").pack(side="left", padx=8)
    info = tk.Label(wrap, text="", bg=BG, fg=MUTED, font=("Segoe UI", 9), justify="left")
    info.pack(anchor="w")
    brow = tk.Frame(wrap, bg=BG)
    brow.pack(fill="x", pady=10)
    run_btn = tk.Button(brow, text="▶ شغّل دلوقتي", bg=GREEN, fg="#0f1923", font=("Segoe UI", 12, "bold"), relief="flat",
                        padx=20, pady=8, cursor="hand2", command=lambda: start(False))
    run_btn.pack(side="left")
    tk.Button(brow, text="⏹ إيقاف", bg=RED, fg="white", font=("Segoe UI", 11, "bold"), relief="flat", padx=14, pady=8,
              cursor="hand2", command=lambda: stop_ev.set()).pack(side="left", padx=8)

    tk.Label(wrap, text="📝 ملخص آخر تشغيل (انسخه من هنا أو من زرار النسخ):", bg=BG, fg=MUTED,
             font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8, 2))
    sum_box = scrolledtext.ScrolledText(wrap, height=14, bg=ENTRY_BG, fg=TEXT, font=("Consolas", 10), relief="flat", wrap="word")
    sum_box.pack(fill="both", expand=True)
    sum_box.insert("1.0", "لسه مفيش تشغيل.")

    def copy_summary():
        try:
            root.clipboard_clear()
            root.clipboard_append(sum_box.get("1.0", "end").strip())
            root.update()
        except Exception:
            pass
    tk.Button(brow, text="📋 انسخ الملخص", bg=ACCENT, fg="#0f1923", font=("Segoe UI", 11, "bold"), relief="flat", padx=14,
              pady=8, cursor="hand2", command=copy_summary).pack(side="left", padx=8)

    def save():
        try:
            json.dump({"steps": {k: bool(v.get()) for k, v in var.items()}, "auto": bool(auto.get()), "times": times.get()},
                      open(CHAIN_SETTINGS, "w", encoding="utf-8"), ensure_ascii=False)
        except Exception:
            pass

    def say(m):
        LOG_Q.put(("noon", f"[السلسلة] {m}\n"))

    def setl(k, txt, col):
        root.after(0, lambda: lbls[k].configure(text=txt, fg=col))

    def fixed():
        out = []
        for t in str(times.get()).replace("،", ",").replace(" ", ",").split(","):
            t = t.strip()
            try:
                if t:
                    out.append(datetime.datetime.strptime(t, "%H:%M").strftime("%H:%M"))
            except ValueError:
                pass
        return out

    def on_auto():
        if auto.get():
            # نقفل الجدولات القديمة عشان مايشتغلوش مرتين
            for k in ("noonapi", "noonacct", "tyapi", "tyacct", "amzacct"):
                a = INSTS.get(k)
                try:
                    if a is not None:
                        a.auto.set(False)
                        if hasattr(a, "cat_auto"):
                            pass
                        a.save_settings()
                except Exception:
                    pass
            try:
                MODS["noon"].stop_scheduler()
                if hasattr(MODS["noon"], "sched_status"):
                    MODS["noon"].sched_status.config(text="⏸ الجدولة متوقفة (السلسلة هي اللي بتشغّل)", fg=MUTED)
                    MODS["noon"].sched_toggle_btn.config(text="▶ تشغيل الجدولة", bg=GREEN)
            except Exception:
                pass
            say("⏰ التشغيل التلقائي للسلسلة اتفعّل — الجدولات القديمة (Noon / Noon API / حسابات نون / Trendyol API / حسابات تريندول) اتقفلت")
        save()

    def wait_app(k, label, fn):
        """يشغّل fn على خيط الواجهة ويستنى لحد ما app.running يرجع False."""
        a = INSTS.get(k)
        if a is None:
            raise RuntimeError("التاب مش محمّل")
        root.after(0, fn)
        t0 = time.time()
        started = False
        while time.time() - t0 < 30:
            if getattr(a, "running", False):
                started = True
                break
            time.sleep(0.5)
        if not started:
            say(f"⚠️ {label}: ماشتغلش (غالباً مفيش متاجر أو شغّال بالفعل) — اتخطّى")
            return False
        while getattr(a, "running", False):
            if stop_ev.is_set():
                try:
                    a.stop.set()
                except Exception:
                    pass
            time.sleep(1)
        return True

    def step_noon():
        m = MODS["noon"]
        try:
            email = m.email_entry.get().strip()
        except Exception:
            email = m.settings.get("email", "")
        stores = m.settings.get("stores", [])
        if not email or not stores:
            say("⚠️ Noon: مفيش إيميل أو متاجر محفوظة — اتخطّى")
            return False
        try:
            otp = int(m.otp_entry.get().strip() or 30)
        except Exception:
            otp = int(m.settings.get("otp_wait", 30))
        m.do_run_reports(email, stores, otp, _LogSink("noon"))
        return True

    def run_chain(auto_run=False):
        state["running"] = True
        stop_ev.clear()
        root.after(0, lambda: run_btn.configure(state="disabled"))
        for k, _ in steps:
            setl(k, "—", MUTED)
        say("▶ بدأت السلسلة" + (" (تلقائي)" if auto_run else ""))
        t_start = time.time()
        results = {}

        def build_summary():
            key_words = ("❌", "⚠", "خطأ", "فشل", "Error", "error", "429", "limit", "Traceback", "اتخطّى", "تخطّى")
            out = [f"ملخص التشغيل — {time.strftime('%Y-%m-%d %H:%M')}", "=" * 40]
            for k, t in steps:
                if k not in results:
                    continue
                st_txt, mins = results[k]
                out.append(f"\n{t}\n  الحالة: {st_txt} ({mins:.1f} دقيقة)")
                lines = [(ts, ln) for ts, kk, ln in LOG_HIST if kk == k and ts >= t_start]
                lines = [(ts, ln) for ts, ln in lines if not ln.startswith("[السلسلة]")]
                errs, seen = [], set()
                for ts, ln in lines:
                    if any(w in ln for w in key_words) and ln not in seen:
                        seen.add(ln)
                        errs.append(ln)
                if errs:
                    out.append("  مشاكل:")
                    out += ["   - " + e[:220] for e in errs[-8:]]
                tail = [ln for _ts, ln in lines[-10:]]
                if tail:
                    out.append("  آخر الأحداث:")
                    out += ["   · " + x[:220] for x in tail]
                if not lines:
                    out.append("  (مفيش سجل للخطوة دي)")
            chain_lines = [ln for ts, kk, ln in LOG_HIST if ts >= t_start and ln.startswith("[السلسلة]")]
            if chain_lines:
                out.append("\nأحداث السلسلة:")
                out += ["  " + x[:220] for x in chain_lines[-15:]]
            return "\n".join(out)

        def show_summary():
            txt = build_summary()
            try:
                with open(os.path.join(BASE_DIR, "chain_summary.txt"), "w", encoding="utf-8") as f:
                    f.write(txt)
            except Exception:
                pass

            def put():
                sum_box.delete("1.0", "end")
                sum_box.insert("1.0", txt)
                copy_summary()
            root.after(0, put)

        def run_step(k, t):
            setl(k, "⏳ شغّال...", ACCENT)
            say(f"▶ {t}")
            t0 = time.time()
            try:
                if k == "noon":
                    ok = step_noon()
                elif k == "noonapi":
                    ok = wait_app("noonapi", "Noon API", lambda: INSTS["noonapi"].go(auto=True))
                elif k == "noonacct":
                    ok = wait_app("noonacct", "حسابات نون", lambda: INSTS["noonacct"].go("daily", auto=True, sheet=True))
                elif k == "tyapi":
                    ok = wait_app("tyapi", "Trendyol API", lambda: INSTS["tyapi"].go(auto=True))
                elif k == "tyacct":
                    ok = wait_app("tyacct", "حسابات تريندول", lambda: INSTS["tyacct"].go("daily", auto=True, sheet=True))
                else:
                    ok = wait_app("amzacct", "حسابات أمازون", lambda: INSTS["amzacct"].go("daily", auto=True, sheet=True))
                mins = (time.time() - t0) / 60
                results[k] = ("✅ خلّص" if ok else "⚠️ اتخطّى", mins)
                setl(k, (f"✅ خلّص ({mins:.1f} د)" if ok else "⚠️ اتخطّى"), GREEN if ok else ACCENT)
                say(f"{'■' if ok else '⚠️'} {t}: {'خلّص' if ok else 'اتخطّى'} ({mins:.1f} دقيقة)")
            except Exception as e:
                results[k] = ("❌ فشل", (time.time() - t0) / 60)
                setl(k, "❌ فشل", RED)
                say(f"❌ {t}: {type(e).__name__}: {str(e)[:200]} — هكمّل للي بعده")

        # المنصات بتشتغل بالتوازي. جوه المنصة الواحدة الترتيب ثابت: Noon API قبل حسابات نون (بتاخد أوردراتها)، وTrendyol API قبل حسابات تريندول
        GROUPS = [["noon"], ["noonapi", "noonacct"], ["tyapi", "tyacct"], ["amzacct"]]
        tt = dict(steps)

        def run_group(keys):
            for k in keys:
                if not var[k].get():
                    continue
                if stop_ev.is_set():
                    say("⏹ السلسلة اتوقفت")
                    break
                run_step(k, tt[k])

        try:
            ths = [threading.Thread(target=run_group, args=(g,), daemon=True) for g in GROUPS if any(var[k].get() for k in g)]
            for th in ths:
                th.start()
            for th in ths:
                th.join()
            say("✅ السلسلة خلصت")
            try:
                show_summary()
            except Exception as e:
                say(f"⚠️ تعذر بناء الملخص: {e}")
        finally:
            state["running"] = False
            root.after(0, lambda: run_btn.configure(state="normal"))

    def start(auto_run):
        if state["running"]:
            return
        save()
        threading.Thread(target=run_chain, args=(auto_run,), daemon=True).start()

    def tick():
        try:
            ft = fixed()
            info.configure(text=("المواعيد: " + " ، ".join(ft)) if (auto.get() and ft) else "")
            now = time.strftime("%H:%M")
            key = time.strftime("%Y-%m-%d ") + now
            if auto.get() and now in ft and key != state["last_key"] and not state["running"]:
                state["last_key"] = key
                start(True)
        except Exception:
            pass
        root.after(20000, tick)

    try:
        d = json.load(open(CHAIN_SETTINGS, encoding="utf-8"))
        for k, v in (d.get("steps") or {}).items():
            if k in var:
                var[k].set(bool(v))
        auto.set(bool(d.get("auto", False)))
        times.set(str(d.get("times", "07:00")))
    except Exception:
        pass
    times.trace_add("write", lambda *a: save())
    root.after(8000, lambda: (on_auto() if auto.get() else None))
    root.after(20000, tick)



def main():
    try:                                   # ويندوز: واجهة حادة على الشاشات عالية الدقة
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    root.title("Platforms Server — Noon · Amazon · Trendyol · Noon API")
    root.geometry("1100x820")
    root.minsize(720, 480)
    root.configure(bg=BG)
    try:
        base = getattr(sys, "_MEIPASS", BASE_DIR)
        for ico in ("Platforms.ico", "Noon.ico"):
            p = os.path.join(base, ico)
            if os.path.exists(p):
                root.iconbitmap(default=p)   # default = كل النوافذ المنبثقة كمان بنفس الأيقونة (من غير ريشة)
                break
    except Exception:
        pass

    style = ttk.Style()
    try:
        style.theme_use("default")
    except Exception:
        pass
    style.configure("Main.TNotebook", background=BG, borderwidth=0)
    style.configure("Main.TNotebook.Tab", background=CARD, foreground=TEXT, padding=[18, 8], font=("Segoe UI", 11, "bold"))
    style.map("Main.TNotebook.Tab", background=[("selected", ACCENT)], foreground=[("selected", "#0f1923")])

    # شريط الحالة
    status = tk.Frame(root, bg=CARD)
    status.pack(fill="x")
    status_lbl = tk.Label(status, text="🟢 فاضي", bg=CARD, fg=GREEN, font=("Segoe UI", 10, "bold"), anchor="w")
    status_lbl.pack(side="left", padx=12, pady=5)
    chrome_note = tk.Label(status, text="🔒 قفل كروم عام: بوت واحد بس بيشغّل كروم في الوقت", bg=CARD, fg=MUTED, font=("Segoe UI", 9))
    chrome_note.pack(side="right", padx=12)

    paned = tk.PanedWindow(root, orient="vertical", bg=BG, sashwidth=8, sashrelief="raised", bd=0)
    paned.pack(fill="both", expand=True)

    nb = ttk.Notebook(paned, style="Main.TNotebook")
    ulog = UnifiedLog(paned)
    paned.add(nb, minsize=200, stretch="always")
    paned.add(ulog, minsize=90, stretch="never")

    # زرار تكبير/استرجاع السجل
    state = {"max": False, "sash": None}

    def toggle_log():
        try:
            if not state["max"]:
                state["sash"] = paned.sash_coord(0)[1]
                paned.sash_place(0, 0, 160)
                state["max"] = True
                big_btn.configure(text="⤡ استرجاع")
            else:
                paned.sash_place(0, 0, state["sash"] or int(root.winfo_height() * 0.62))
                state["max"] = False
                big_btn.configure(text="⤢ تكبير السجل")
        except Exception:
            pass
    big_btn = tk.Button(status, text="⤢ تكبير السجل", relief="flat", bg=ENTRY_BG, fg=TEXT, font=("Segoe UI", 9, "bold"),
                        padx=10, cursor="hand2", command=toggle_log)
    big_btn.pack(side="right", padx=6, pady=3)

    _install_wheel(root)

    shutdowns = []
    for key, title, modname, _col in BOTS:
        tab = ScrollTab(nb)
        nb.add(tab, text=f"  {title}  ")
        try:
            mod = _load_embedded(modname)
            MODS[key] = mod
            hook_module(key, title.split(" ", 1)[1], mod)
            sd = mod.build_ui(tab.inner)
            if callable(sd):
                shutdowns.append(sd)
            try:
                _AppCls = mod.__dict__.get("App")
                if _AppCls is not None:
                    for _w in tab.inner.winfo_children():
                        if isinstance(_w, _AppCls):
                            INSTS[key] = _w
                            break
            except Exception:
                pass
            LOG_Q.put((key, "✅ البوت جاهز\n"))
        except Exception as e:
            tk.Label(tab.inner, text=f"❌ تعذّر تحميل {title}\n{e}", bg=BG, fg=RED, font=("Segoe UI", 12, "bold"),
                     justify="left").pack(padx=20, pady=20, anchor="w")
            tk.Label(tab.inner, text=traceback.format_exc()[-1500:], bg=BG, fg=MUTED, font=("Consolas", 9),
                     justify="left", anchor="w").pack(padx=20, anchor="w")
            LOG_Q.put((key, f"❌ تعذّر تحميل البوت: {e}\n"))

    # تاب السلسلة (أول تاب)
    try:
        _ctab = ScrollTab(nb)
        nb.insert(0, _ctab, text="  ▶ بالترتيب  ")
        build_chain_ui(_ctab.inner, root)
        nb.select(_ctab)
    except Exception as e:
        LOG_Q.put(("noon", f"❌ تعذّر بناء تاب السلسلة: {e}\n"))

    # دورة تحديث: السجل المشترك + شريط الحالة
    def pump():
        try:
            for _ in range(400):
                k, t = LOG_Q.get_nowait()
                ulog.add(k, t)
        except queue.Empty:
            pass
        except Exception:
            pass
        run, wait = _STATE["running"], _STATE["waiting"]
        if run:
            txt = f"▶ شغّال دلوقتي: {run}" + (f"   |   ⏳ مستني: {' ، '.join(wait)}" if wait else "")
            status_lbl.configure(text=txt, fg=ACCENT)
        else:
            status_lbl.configure(text="🟢 فاضي", fg=GREEN)
        root.after(250, pump)
    root.after(250, pump)

    def on_close():
        if RUN_GATE.locked():
            if not messagebox.askyesno("تأكيد", f"فيه تشغيل شغّال دلوقتي ({_STATE['running'] or '...'}).\nتقفل البرنامج برضه؟"):
                return
        for sd in shutdowns:
            try:
                sd()
            except Exception:
                pass
        root.destroy()
        os._exit(0)          # يقفل أي Thread شغّال في الخلفية

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()




# ═══════════════════════════════════════════════════════════════════════════
# كود البوتات الثلاثة (مدمج) — تقدر تعدّل جواه عادي
# ═══════════════════════════════════════════════════════════════════════════
_EMBEDDED = {}

_EMBEDDED['NOONServer'] = r'''import time
import re
import imaplib
import email as email_lib
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk, simpledialog, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
import os
import sys
import glob
import csv
import json
import threading
import datetime
import openpyxl
import urllib.request
import urllib.parse
import shutil
import tempfile
from openpyxl.styles import Font, PatternFill, Alignment
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ===== Lock Check =====
lock_folder = r"C:\Lock noon"
LOCK_OK = os.path.exists(lock_folder)      # من غيره بوت نون بيتعطّل (زي ما كان بيقفل البرنامج)

# ===== Paths =====
DOWNLOADS_DIR      = os.path.join(os.path.expanduser("~"), "Downloads")
SCRIPT_DIR         = os.environ.get("PLATFORMS_BASE_DIR") or os.path.dirname(os.path.abspath(__file__))
SETTINGS_PATH      = os.path.join(os.path.expanduser("~"), "noon_settings.json")

def resource_path(relative_path):
    """
    بيرجع المسار الصحيح لملف مرفق (زي الأيقونة) سواء البرنامج شغال من
    السورس (.py) أو متجمّع كملف .exe واحد (PyInstaller --onefile). في حالة
    الـ exe، الملفات المرفقة بـ --add-data بتتفك مؤقتًا وقت التشغيل في
    مجلد sys._MEIPASS.
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = SCRIPT_DIR
    return os.path.join(base_path, relative_path)

# مسار ثابت لمجلد ملف اعتماد Google (Service Account JSON) على السيرفر
# المشترك، واسم الـ Google Sheet ثابت — القيمتين دول مش بيدخلهم المستخدم
# ولا بيظهروا في الواجهة، بياخدهم البرنامج تلقائيًا.
GSHEET_CREDS_DIR   = r"\\WIN-181U7G4HFM5\Share For Markting\NOONIMG\Copy"
GSHEET_NAME        = "Complaints"

# مسار ثابت لملف إكسيل "تكويد نون" اللي بنسحب منه بيانات شيتات Tacweed /
# تكلفة البيع / WarehouseStock — القيمة دي مش بتدخلها المستخدم ولا بتظهر
# في الواجهة، بياخدها البرنامج تلقائيًا زي GSHEET_CREDS_DIR.
TACWEED_XLSM_PATH  = r"\\WIN-181U7G4HFM5\Share For All\تكويد المتاجر\تكويد نون - 25-05-2023.xlsm"
TACWEED_SHEET_NOON = "Noon"
TACWEED_SHEET_JARD = "الجرد"

def _resolve_gsheet_creds_path():
    """
    بيدور على أول ملف .json موجود جوه GSHEET_CREDS_DIR ويرجع مساره الكامل،
    أو "" لو مفيش أي ملف .json في المجلد أو المجلد مش متاح.
    """
    try:
        matches = glob.glob(os.path.join(GSHEET_CREDS_DIR, "*.json"))
        if matches:
            return matches[0]
    except Exception:
        pass
    return ""


# ================= Settings =================
def load_settings():
    default = {
        "email": "",
        "otp_wait": 30,
        "gmail_app_password": "",
        "stores": [],        # [{"name": "...", "url": "..."}, ...]
        "stores2": [],      # المتاجر الثانية directship
        "schedules": [],      # ["08:00", "13:00", ...]
        "gsheet_creds_path": "",     # مسار ملف JSON بتاع Google Service Account
        "gsheet_name": "Complaints", # اسم الـ Google Sheet (نفس اللي في برنامج الطلبات)
        "skip_inventory": True,      # الجرد FBN بييجي من تاب Noon API → البوت يتخطّى تنزيله
        "skip_sales": True,          # المبيعات بتيجي من تاب Noon API
        "skip_catalog": True,        # الكتالوج (LIVE) بيجي من تاب Noon API
    }
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, encoding="utf-8") as f:
                data = json.load(f)
                for k, v in default.items():
                    if k not in data:
                        data[k] = v
                return data
        except Exception:
            pass
    return default

def save_settings(settings):
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)



# ================= Driver =================
def start_driver(log=None):
    """
    كل تشغيلة (Run) بيتفتح فيها Chrome ببروفايل مؤقت جديد كليًا (مفيش أي هيستوري
    أو كوكيز محفوظة من تشغيلة سابقة)، فكل تشغيلة جديدة لازم تاخد OTP من الأول.
    لكن جوه نفس التشغيلة، نفس المتصفح/الجلسة بيتستخدم لكل المتاجر بالترتيب —
    فالـ OTP بيتطلب مرة واحدة بس في أول متجر، وباقي المتاجر بتكمل على نفس الجلسة.

    ⚠️ تم الاستغناء عن webdriver_manager بالكامل — Selenium Manager المدمج
    في مكتبة selenium (4.6+) هو اللي بيحدد نسخة Chrome المثبتة ويحمّل
    chromedriver المطابق ليها تلقائيًا، فمفيش مشكلة اختلاف نسخ ولا ملفات lock.

    بترجع (driver, profile_dir) — عشان بعد driver.quit() نقدر نمسح
    الفولدر المؤقت ده من على الجهاز.
    """
    profile_dir = tempfile.mkdtemp(prefix="NoonChromeTemp_")

    opts = webdriver.ChromeOptions()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--remote-debugging-port=0")
    opts.add_argument("--disable-features=FedCm,FedCmIdpSigninStatus,IdentityConsentDialog,AutofillServerCommunication")
    # ── لو الشباك اتقفل تحت (minimized) أو مش عليه فوكس، كروم بيبطّئ/يوقف
    # التايمرات والتابات اللي في الخلفية (background throttling) — وده بيوقف
    # الـ polling بتاع صفحات زي تقرير المبيعات وتصدير الكتالوج. الأسطر دي
    # بتمنع كروم من عمل الإبطاء ده عشان الصفحة تفضل شغالة حتى لو الشباك تحت ──
    opts.add_argument("--disable-background-timer-throttling")
    opts.add_argument("--disable-backgrounding-occluded-windows")
    opts.add_argument("--disable-renderer-backgrounding")
    opts.add_argument(f"--user-data-dir={profile_dir}")
    opts.add_argument("--profile-directory=NoonProfile")
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    opts.add_experimental_option("prefs", {"credentials_enable_service": False})
    # ── لو ظهر أي Native Dialog (زي تنبيه "مغادرة الصفحة؟" beforeunload) —
    # المتصفح يقبله تلقائيًا بدل ما يوقف أي أمر Selenium (زي driver.get) لحد الأبد ──
    opts.set_capability("unhandledPromptBehavior", "accept")

    try:
        driver = webdriver.Chrome(options=opts)
    except Exception as e:
        if log:
            log_write(log, f"⚠️ فشل تشغيل Chrome: {str(e)[:300]}\n")
        shutil.rmtree(profile_dir, ignore_errors=True)
        raise

    # ── Timeout واضح لأي تنقل/سكريبت — عشان أي تعليق (hang) يتحول لخطأ سريع
    # نقدر نتعافى منه (بدل ما البرنامج يفضل واقف من غير أي رسالة) ──
    try:
        driver.set_page_load_timeout(45)
        driver.set_script_timeout(30)
    except Exception:
        pass
    return driver, profile_dir


def close_driver(driver, profile_dir):
    """يقفل الدرايفر ويمسح الفولدر المؤقت بتاعه (مفيش أي أثر يفضل على الجهاز)."""
    try:
        driver.quit()
    except Exception:
        pass
    if profile_dir:
        shutil.rmtree(profile_dir, ignore_errors=True)


def ensure_driver_alive(driver, driver_profile_dir, log):
    """
    بيتأكد إن جلسة Selenium لسه شغالة، ولو ماتت (مثلاً بعد Exception من chromedriver
    زي اللي بيظهر أحيانًا بعد عمليات طويلة زي تحميل الكتالوج) بيقفل القديم ويفتح
    Chrome جديد كليًا تلقائيًا بدل ما البرنامج يقف أو يحتاج تدخل يدوي.
    بيرجع (driver, driver_profile_dir) — ممكن يكونوا نفسهم أو جدد لو حصل تعافي.
    """
    try:
        if driver is not None and _is_driver_alive(driver):
            return driver, driver_profile_dir
    except Exception:
        pass

    log_write(log, "🔄 الجلسة القديمة ماتت — جاري فتح Chrome جديد تلقائيًا...\n")
    try:
        close_driver(driver, driver_profile_dir)
    except Exception:
        pass
    new_driver, new_profile_dir = start_driver(log)
    return new_driver, new_profile_dir


# ================= إغلاق نافذة "تسجيل الدخول بجوجل" =================
def dismiss_google_one_tap(driver):
    """
    صفحة نون بتظهر أحيانًا نافذة "تواصل مع هذا المستخدم / استخدم حسابًا مختلفًا"
    (Google One Tap / FedCM) فوق الصفحة، وده بيمنع الأتمتة من الوصول لخانة الـ OTP
    أو زرار Continue. الدالة دي بتحاول تشيل النافذة دي من الصفحة.
    """
    try:
        driver.execute_script("""
            document.querySelectorAll(
                "iframe[src*='accounts.google.com'], #credential_picker_container, " +
                "#credential_picker_iframe, .cred-picker-container, div[id^='credential_picker']"
            ).forEach(function(el){ el.remove(); });
        """)
    except Exception:
        pass
    try:
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
    except Exception:
        pass


# ================= OTP من الجيميل (Headless عبر IMAP) =================
def _extract_otp_code(html_raw, plain_text):
    """
    بيدور على كود الـ 6 أرقام بالترتيب ده (الأدق الأول):
    1) جوه الـ div المميز اللي نون بتحط فيه الكود (font-weight:700 وخط كبير) — ده شكله الثابت في إيميلات نون.
    2) لو مفيش، سطر مستقل في نسخة النص العادي (text/plain) مكون من 6 أرقام بس.
    3) آخر حل: أي 6 أرقام متجاورة في نص الإيميل.
    """
    # 1) الـ div المميز في تصميم إيميل نون
    if html_raw:
        m = re.search(r'font-weight:\s*700[^>]*>\s*(\d{6})\s*<', html_raw)
        if m:
            return m.group(1)
        # نسخة احتياطية لو ترتيب الـ style اتغير شوية
        m = re.search(r'letter-spacing[^>]*>\s*(\d{6})\s*<', html_raw)
        if m:
            return m.group(1)

    # 2) سطر مستقل من 6 أرقام في النص العادي
    if plain_text:
        for line in plain_text.splitlines():
            line = line.strip()
            if re.fullmatch(r"\d{6}", line):
                return line

    # 3) آخر حل — أي 6 أرقام متجاورة في أي نص متاح
    for text in (plain_text, re.sub(r"<[^<]+?>", " ", html_raw) if html_raw else ""):
        if text:
            m = re.search(r"\b(\d{6})\b", text)
            if m:
                return m.group(1)

    return None


def fetch_otp_from_gmail_imap(gmail_email, gmail_app_password, log, wait_before=5, timeout=45, poll_interval=3):
    """
    يجيب كود الـ OTP من الجيميل مباشرة عن طريق IMAP — من غير فتح أي متصفح أو تاب خالص (headless).
    ده أسرع وأثبت من فتح جيميل بالمتصفح لأنه مش بيتأثر بنوافذ جوجل (One Tap/FedCM) ولا بشكل صفحة الجيميل.

    ⚠️ مهم: لازم "App Password" من حساب جوجل (مش الباسورد العادي بتاع الحساب).
    App Password بيتعمل من: myaccount.google.com/apppasswords
    (لازم يكون مفعّل عندك "تحقق بخطوتين / 2-Step Verification" الأول عشان الخيار ده يظهر).
    """
    if not gmail_email or not gmail_app_password:
        log_write(log, "⚠️ مفيش إيميل/App Password للجيميل محفوظين في الإعدادات\n")
        return None

    log_write(log, f"📧 جاري انتظار وصول كود الـ OTP بالإيميل ({wait_before} ثانية)...\n")
    time.sleep(wait_before)

    otp_code = None
    imap = None
    try:
        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        imap.login(gmail_email, gmail_app_password)
    except Exception as e:
        log_write(log, f"⚠️ فشل تسجيل الدخول على IMAP: {str(e)}\n")
        log_write(log, "💡 تأكد إنك مستخدم App Password صحيح من myaccount.google.com/apppasswords\n")
        return None

    deadline = time.time() + timeout
    try:
        while time.time() < deadline and not otp_code:
            imap.select("INBOX")
            status, data = imap.search(None, '(FROM "verify@noon.com")')
            if status == "OK" and data and data[0]:
                ids = data[0].split()
                latest_id = ids[-1]
                status2, msg_data = imap.fetch(latest_id, "(RFC822)")
                if status2 == "OK" and msg_data and msg_data[0]:
                    msg = email_lib.message_from_bytes(msg_data[0][1])
                    plain_text = ""
                    html_raw = ""

                    if msg.is_multipart():
                        for part in msg.walk():
                            ctype = part.get_content_type()
                            try:
                                if ctype == "text/plain" and not plain_text:
                                    plain_text = part.get_payload(decode=True).decode(errors="ignore")
                                elif ctype == "text/html" and not html_raw:
                                    html_raw = part.get_payload(decode=True).decode(errors="ignore")
                            except Exception:
                                pass
                    else:
                        try:
                            payload = msg.get_payload(decode=True).decode(errors="ignore")
                        except Exception:
                            payload = str(msg.get_payload())
                        if msg.get_content_type() == "text/html":
                            html_raw = payload
                        else:
                            plain_text = payload

                    otp_code = _extract_otp_code(html_raw, plain_text)
                    if otp_code:
                        log_write(log, f"✅ تم جلب كود التحقق من الجيميل (IMAP): {otp_code}\n")
                        break

            if not otp_code:
                time.sleep(poll_interval)
    except Exception as e:
        log_write(log, f"⚠️ خطأ أثناء قراءة الجيميل عبر IMAP: {str(e)}\n")
    finally:
        try:
            imap.logout()
        except Exception:
            pass

    if not otp_code:
        log_write(log, "⚠️ ما لقيتش كود الـ OTP في الوقت المحدد\n")

    return otp_code


# ================= Login =================
def _noon_logged_in(driver):
    """الجلسة شغالة فعلاً؟ (مش واقفين في صفحة اللوجين)"""
    url = driver.current_url
    return "noon.partners" in url and "login" not in url


def _is_driver_alive(driver):
    """
    بيتأكد إن الـ session بتاعة Selenium لسه شغالة فعليًا (مش بس المتصفح فاتح).
    أحيانًا بعد عمليات طويلة (تحميل كتالوج، polling، أوامر CDP) الاتصال بين
    Selenium والمتصفح بيقع حتى لو نافذة Chrome نفسها لسه مفتوحة — وده بيظهر
    كخطأ WebDriverException (Stacktrace من chromedriver) عند أي أمر جديد.
    """
    try:
        _ = driver.current_url
        return True
    except Exception:
        return False


def _safe_get(driver, url, log, description="", retries=2):
    """
    زي driver.get() العادي بس بيحاول أكتر من مرة (وبيدي فرصة للصفحة/المتصفح
    يستقر) قبل ما يرفع الخطأ فعليًا. ده بيغطي حالات كتير بتحصل بعد عملية طويلة
    زي تصدير المبيعات (تنزيل ملف، نافذة تصدير، إلخ) واللي بتخلي أول محاولة
    navigation بعدها تفشل لوحدها بس تنجح لو اتعادت.
    """
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            driver.get(url)
            return True
        except Exception as e:
            last_err = e
            if log:
                log_write(log, f"⚠️ محاولة {attempt}/{retries} فشلت في فتح {description or url}: {str(e)[:150]}\n")
            time.sleep(2)
    if last_err:
        raise last_err
    return False


def _wait_noon_login_success(driver, timeout=20):
    """يستنى لحد ما نطلع من صفحة اللوجين فعلاً (يعني الدخول نجح)، أو يخلص الوقت."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _noon_logged_in(driver):
            return True
        time.sleep(1)
    return False


def login_to_noon(driver, email, orders_url, log, otp_wait):
    """
    الـ OTP بيتطلب مرة واحدة بس في أول التشغيلة (Run) — أول متجر بيدخل بالإيميل + كود
    الـ OTP من الجيميل. باقي المتاجر في نفس التشغيلة (نفس المتصفح/نفس الجلسة) بتستخدم
    الجلسة الحية دي مباشرة من غير ما تطلب كود جديد.

    مفيش أي اعتماد على Session محفوظة عبر تشغيلات مختلفة: البروفايل مؤقت وبيتمسح
    بالكامل بعد كل تشغيلة (close_driver)، فالتشغيلة الجاية هتحتاج OTP جديد من الأول.
    التعليق ده بيتطابق دلوقتي مع سلوك الكود فعليًا.

    يرجع True لو الدخول نجح فعلاً، وFalse لو فشل — عشان الكود اللي بعده
    (استخراج الطلبات) ميتنفذش غير لما الدخول ينجح فعلاً.
    """
    # امنع تحميل ودجت "تسجيل الدخول بجوجل" (Google One Tap/FedCM) من الأساس
    # عشان ميظهرش فوق الصفحة ويوقف الأتمتة — إحنا مش محتاجينه أصلاً لأننا بنستخدم إيميل + OTP.
    try:
        driver.execute_cdp_cmd("Network.enable", {})
        driver.execute_cdp_cmd("Network.setBlockedURLs", {"urls": ["*accounts.google.com*"]})
    except Exception:
        pass

    # ── أول حاجة: جرب روح على رابط المتجر مباشرة. لو الجلسة شغالة من نفس
    # التشغيلة (يعني مش أول متجر) هيدخل على طول من غير أي داعي للمرور
    # بصفحة https://login.noon.partners/ar خالص. المرور بصفحة اللوجين دي
    # وهي مش لازمة هو اللي كان بيسبب التعليق/الكراش في المتاجر اللي بعد الأول. ──
    try:
        _safe_get(driver, orders_url, log, "صفحة المتجر")
        time.sleep(2)
        dismiss_google_one_tap(driver)
        if _noon_logged_in(driver):
            log_write(log, "✅ الجلسة شغالة من نفس التشغيلة — دخلنا على المتجر مباشرة من غير المرور بصفحة تسجيل الدخول\n")
            return True
    except Exception as e:
        log_write(log, f"⚠️ فشلت محاولة الدخول المباشر على رابط المتجر ({str(e)[:150]}) — هنجرب من صفحة تسجيل الدخول\n")

    try:
        _safe_get(driver, "https://login.noon.partners/ar", log, "صفحة الدخول")
        wait = WebDriverWait(driver, 30)
        time.sleep(3)
        dismiss_google_one_tap(driver)

        # ── لو الجلسة شغالة فعلاً (متجر مش أول متجر في نفس التشغيلة) — منعملش OTP تاني ──
        if _noon_logged_in(driver):
            log_write(log, "✅ الجلسة شغالة من نفس التشغيلة — من غير الحاجة لكود OTP تاني\n")
            _safe_get(driver, orders_url, log, "صفحة الطلبات")
            return True

        email_input = wait.until(EC.presence_of_element_located((By.NAME, "channelIdentifier")))
        email_input.clear()
        email_input.send_keys(email)

        dismiss_google_one_tap(driver)
        continue_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//button[@type='submit']//span[text()='مواصلة']/parent::button"
        )))
        continue_btn.click()
        time.sleep(3)
        dismiss_google_one_tap(driver)

        # ── ممكن يدخل من غير OTP خالص (باسكي/Session) — نتأكد قبل ما نجيب كود جديد ──
        if _noon_logged_in(driver):
            log_write(log, "✅ الدخول نجح من غير الحاجة لكود OTP\n")
            _safe_get(driver, orders_url, log, "صفحة الطلبات")
            return True

        # ── لو نون نفسها عارضة "لقد قمت بتسجيل الدخول بالفعل. تابع هنا" ──
        try:
            continue_card = WebDriverWait(driver, 5).until(EC.presence_of_element_located((
                By.XPATH,
                "//div[contains(@class,'cursor-pointer')]"
                "[.//p[contains(text(),'لقد قمت بتسجيل الدخول بالفعل')]]"
            )))
            log_write(log, "✅ نون لقت جلسة سابقة (لقد قمت بتسجيل الدخول بالفعل) — جاري الضغط على الحساب للمتابعة\n")
            driver.execute_script("arguments[0].click();", continue_card)
            time.sleep(3)
            if _wait_noon_login_success(driver, timeout=15):
                log_write(log, "✅ تم الدخول من غير الحاجة لكود OTP\n")
                _safe_get(driver, orders_url, log, "صفحة الطلبات")
                return True
        except Exception:
            pass  # النص ده مش موجود — نكمل لخطة الـ OTP العادية تحت

        # ── دلوقتي فعلاً محتاجين OTP (مرة واحدة بس أول التشغيلة) — نجيبه من الجيميل ──
        log_write(log, "📨 جاري جلب كود الـ OTP من الجيميل تلقائيًا...\n")

        otp_code = fetch_otp_from_gmail_imap(
            gmail_email=settings.get("email", email),
            gmail_app_password=settings.get("gmail_app_password", ""),
            log=log
        )

        dismiss_google_one_tap(driver)

        if otp_code:
            try:
                otp_input = wait.until(EC.presence_of_element_located((
                    By.CSS_SELECTOR, "input[data-input-otp='true']"
                )))
                dismiss_google_one_tap(driver)
                driver.execute_script("arguments[0].scrollIntoView(true);", otp_input)
                driver.execute_script("arguments[0].click();", otp_input)
                otp_input.send_keys(otp_code)
                time.sleep(1)
                dismiss_google_one_tap(driver)

                submit_btn = wait.until(EC.element_to_be_clickable((
                    By.XPATH, "//button[@type='submit']"
                )))
                driver.execute_script("arguments[0].click();", submit_btn)
            except Exception as e:
                log_write(log, f"⚠️ اتجاب الكود بس فشل إدخاله في الصفحة: {str(e)}\n")
                otp_code = None  # يخليها تروح لخطة الإدخال اليدوي تحت

        if not otp_code:
            log_write(log, f"⏳ أدخل OTP يدويًا... لديك {otp_wait} ثانية\n")
            for remaining in range(otp_wait, 0, -5):
                log_write(log, f"   ⏱ باقي {remaining} ثانية...\n")
                time.sleep(5)

        # ── 4) الكود اللي بعد كده (استخراج الطلبات) منفذش غير لو الدخول نجح فعلاً ──
        if _wait_noon_login_success(driver, timeout=20):
            log_write(log, "✅ تم تسجيل الدخول بنجاح — جاري الانتقال للطلبات...\n")
            _safe_get(driver, orders_url, log, "صفحة الطلبات")
            return True
        else:
            log_write(log, "❌ الدخول لم ينجح (لسه واقف في صفحة اللوجين) — هيتم تخطي المتجر ده\n")
            return False

    except Exception as e:
        err_msg = str(e)
        if len(err_msg) > 300:
            err_msg = err_msg[:300] + " ...(تم اختصار الرسالة)"
        if _is_driver_alive(driver):
            log_write(log, f"⚠️ خطأ أثناء تسجيل الدخول: {err_msg}\n")
        else:
            log_write(log, f"💥 انقطع الاتصال بجلسة المتصفح (WebDriver) — هيتم إعادة فتح Chrome تلقائيًا: {err_msg}\n")
        return False


# ================= تقرير الجرد (Inventory) =================
def _download_inventory_report_once(driver, store_code, log):
    """
    يفتح صفحة الجرد بتاعة المتجر (fbn.noon.partners) ويضغط زرار "تنزيل" وينتظر الملف.
    ⚠️ ملحوظة: الزرار هنا مبني على افتراض عام (أي عنصر قابل للنقر فيه نص "تنزيل")
    لأننا معندناش الـ HTML الدقيق لزرار الصفحة دي. لو مانفعش زي ما هو، ابعتلي
    الـ HTML بتاعه (زي ما بعتّ HTML نافذة تصدير المبيعات بالظبط) وهظبطه فورًا.
    """
    if not store_code:
        log_write(log, "⚠️ مفيش كود متجر — تخطي الجرد\n")
        return None

    url = f"https://fbn.noon.partners/ar-sa/inventory?mp=noon&project=PRJ{store_code}&page=1"
    log_write(log, f"\n📦 جرد PRJ{store_code}...\n")
    wait = WebDriverWait(driver, 30)

    try:
        before = set(glob.glob(os.path.join(DOWNLOADS_DIR, "*.csv")))
        driver.get(url)
        time.sleep(4)

        download_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//button[contains(concat(' ', normalize-space(@class), ' '), ' ant-btn ')]"
            "[.//span[normalize-space(text())='تنزيل']]"
        )))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", download_btn)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", download_btn)
        log_write(log, "✅ تم الضغط على تنزيل\n⏳ جاري انتظار تحميل الملف...\n")

        csv_path  = None
        prev_size = -1
        stable    = 0
        for _ in range(60):
            time.sleep(1)
            after = set(glob.glob(os.path.join(DOWNLOADS_DIR, "*.csv")))
            new_files = after - before
            if new_files:
                csv_path = max(new_files, key=os.path.getctime)
                try:
                    curr = os.path.getsize(csv_path)
                    stable = (stable + 1) if curr == prev_size and curr > 0 else 0
                    if stable >= 2:
                        break
                    prev_size = curr
                except Exception:
                    pass

        if not csv_path:
            log_write(log, "❌ لم يتم العثور على ملف الجرد\n")
            return None

        log_write(log, f"✅ تم تحميل: {os.path.basename(csv_path)}\n")
        return csv_path

    except Exception as e:
        log_write(log, f"❌ فشل تنزيل الجرد: {str(e)}\n")
        return None


def download_inventory_report(driver, store_code, log, max_retries=None):
    """
    غلاف حوالين _download_inventory_report_once: لو التنزيل فشل (رجّع None)،
    بيحاول تاني تلقائيًا وباستمرار لحد ما ينجح، من غير حد أقصى للمحاولات.
    """
    attempt = 0
    while True:
        attempt += 1
        result = _download_inventory_report_once(driver, store_code, log)
        if result:
            return result
        log_write(log, f"🔁 إعادة محاولة تنزيل الجرد (محاولة {attempt})...\n")
        time.sleep(2)


def _ensure_window_visible(driver):
    """
    لو المستخدم صغّر (minimize) شباك كروم، الصفحة بتبقى document.hidden=True،
    وساعتها كروم مش بيحسب مقاسات/ظهور العناصر بشكل طبيعي — فأوامر Selenium
    زي wait.until(EC.element_to_be_clickable/visibility_of_element_located)
    بتفضل تفشل بـ TimeoutException (رسالتها بتطلع فاضية تقريبًا: "Message:")
    لحد ما حد يرجع يفتح الشباك يدويًا. الدالة دي بتتأكد كل شوية إن الشباك
    مش مخفي، ولو لقيته مصغّر بترجّعه (restore) تلقائيًا من غير تدخل من
    المستخدم، عشان الصفحة تكمل شغل عادي حتى لو حد قفل الشباك تحت وسابه.
    بترجع بهدوء (من غير استثناء) لو حصل أي خطأ، عشان ما توقفش أي حاجة.
    """
    try:
        is_hidden = driver.execute_script("return document.hidden === true;")
    except Exception:
        return
    if is_hidden:
        try:
            driver.maximize_window()
        except Exception:
            pass


# ================= تصدير المبيعات (Sales Export) =================
def _get_sales_export_link(driver):
    """
    يرجع (href, filename) لأحدث رابط تنزيل ظاهر لملف تصدير المبيعات، أو (None, None).
    بيدور على أي <a download> رابطه بيروح على storage.googleapis.com ويحتوي على
    'sales_export' (زي الرابط اللي بيتبعت فعليًا من نون) — بدل ما يعتمد على
    كلاس المودال (ant-modal-container) اللي ممكن يتغير مع تحديثات الموقع.
    اسم الملف بيتقرا من الرابط نفسه (بعد فك الـ URL-encoding) بدل عنصر منفصل،
    لأن الزرار في النسخة الجديدة أيقونة فقط من غير نص اسم ملف جنبه.
    """
    try:
        a_el = driver.find_element(
            By.XPATH,
            "//a[@download and contains(@href,'storage.googleapis.com') and contains(@href,'sales_export')]"
        )
        href = a_el.get_attribute("href")
        raw_name = href.split("?")[0].rstrip("/").split("/")[-1]
        fname = urllib.parse.unquote(raw_name)
        return href, fname
    except Exception:
        return None, None


def _download_sales_export_report_once(driver, store_code, log):
    """
    يفتح صفحة تقرير المبيعات، يضغط "تصدير" لفتح المودال، يضغط "تصدير جديد"،
    ويستنى لحد ما الملف يجهز (بيتغير اسم الملف الظاهر عن اللي كان قبل الضغط)،
    وبعدين ينزل الملف مباشرة من الرابط الموقّع (GCS) من غير ما يعدي على مجلد
    التنزيلات، ويحفظه هناك بنفس اسمه الأصلي.
    """
    if not store_code:
        log_write(log, "⚠️ مفيش كود متجر — تخطي تقرير المبيعات\n")
        return None

    url = f"https://reports.noon.partners/ar/sales/?project=PRJ{store_code}"
    log_write(log, f"\n💰 تقرير مبيعات PRJ{store_code}...\n")
    wait = WebDriverWait(driver, 30)

    try:
        driver.get(url)
        time.sleep(4)
        _ensure_window_visible(driver)

        export_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//button[.//span[normalize-space(text())='تصدير']]"
        )))
        driver.execute_script("arguments[0].click();", export_btn)
        wait.until(EC.visibility_of_element_located((
            By.XPATH, "//div[contains(@class,'ant-modal-title') and contains(text(),'ملف التصدير')]"
        )))
        time.sleep(1.5)
        log_write(log, "✅ فتح نافذة التصدير\n")

        before_href, before_name = _get_sales_export_link(driver)

        new_export_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//div[contains(@class,'ant-modal-container')]"
            "//button[contains(@class,'ant-btn-primary') and not(contains(@class,'ant-btn-dangerous'))]"
            "[.//span[normalize-space(text())='تصدير جديد']]"
        )))
        driver.execute_script("arguments[0].click();", new_export_btn)
        log_write(log, "⏳ جاري إنشاء تصدير جديد... (بياخد حوالي دقيقة)\n")

        href, fname = None, None
        MAX_ATTEMPTS  = 50   # 50 × 3.5s ≈ 3 دقائق
        REFRESH_EVERY = 5    # كل 5 محاولات (~17s) نعمل refresh ونفتح المودال تاني

        for attempt in range(1, MAX_ATTEMPTS + 1):
            time.sleep(3.5)
            _ensure_window_visible(driver)

            if attempt % REFRESH_EVERY == 0:
                try:
                    driver.refresh()
                    time.sleep(3)
                    export_btn2 = wait.until(EC.element_to_be_clickable((
                        By.XPATH, "//button[.//span[normalize-space(text())='تصدير']]"
                    )))
                    driver.execute_script("arguments[0].click();", export_btn2)
                    wait.until(EC.visibility_of_element_located((
                        By.XPATH, "//div[contains(@class,'ant-modal-title') and contains(text(),'ملف التصدير')]"
                    )))
                    time.sleep(1.5)
                except Exception:
                    pass

            href, fname = _get_sales_export_link(driver)
            if href and fname and fname != before_name:
                log_write(log, f"   ✅ الملف جاهز: {fname}\n")
                break
            if attempt % REFRESH_EVERY == 0:
                log_write(log, f"   🔄 لسه بيجهز... ({int(attempt*3.5)}s)\n")

        if not href or fname == before_name:
            log_write(log, "❌ لم يجهز ملف تصدير المبيعات في الوقت المحدد\n")
            return None

        # اسم الملف بيجيله وقت زي "12:24:56" واللي فيه ":" — وده حرف ممنوع في
        # أسماء ملفات Windows (بيدي [Errno 22] Invalid argument)، فبنستبدل أي
        # حرف ممنوع بـ "-" قبل الحفظ.
        safe_fname = re.sub(r'[\\/:*?"<>|]', "-", fname)
        local_path = os.path.join(DOWNLOADS_DIR, safe_fname)
        try:
            urllib.request.urlretrieve(href, local_path)
            log_write(log, f"✅ تم تنزيل: {safe_fname}\n")
            return local_path
        except Exception as e:
            log_write(log, f"❌ فشل تنزيل الملف: {str(e)}\n")
            return None

    except Exception as e:
        log_write(log, f"❌ خطأ تقرير المبيعات: {str(e)}\n")
        return None


def download_sales_export_report(driver, store_code, log, max_retries=None):
    """
    غلاف حوالين _download_sales_export_report_once: لو التنزيل فشل (رجّع None)،
    بيحاول تاني تلقائيًا وباستمرار لحد ما ينجح، من غير حد أقصى للمحاولات.
    """
    attempt = 0
    while True:
        attempt += 1
        _ensure_window_visible(driver)
        result = _download_sales_export_report_once(driver, store_code, log)
        if result:
            return result
        log_write(log, f"🔁 إعادة محاولة تنزيل تقرير المبيعات (محاولة {attempt})...\n")
        time.sleep(2)


# ================= تصدير الكتالوج (Catalog Export) =================
"""
بيفتح صفحة exports بتاعة الكتالوج لكل متجر، يضغط "إضافة تصدير"، يختار
النوع "تصدير الكتالوج" ومدى الظهور "نشط"، يضغط "إنشاء"، يستنى لحد ما أحدث
صف في القايمة يبقى "مكتمل"، وبعدين يضغط زرار "تحميل" بتاعه وينتظر الملف
في مجلد التنزيلات (بنفس أسلوب تحميل تقرير الجرد).

⚠️ ملحوظة: الـ selectors هنا مبنية على الـ HTML اللي بعته بالظبط، لكن بما
إن الـ id بتاع عناصر الـ ant-select (زي _r_4_) عشوائي وبيتغير كل مرة،
استخدمنا نص اللابل ("نوع") ونص الأوبشن ("تصدير الكتالوج" / "نشط") بدل
الـ id. لو حصل أي تغيير في تصميم الصفحة ومشتغلتش، ابعتلي الـ HTML الجديد
وهظبطها فورًا.
"""

def _select_combobox_option(driver, wait, label_text, option_text, log=None):
    """
    بيدوس على حقل اختيار (combobox) داخل مودال "إضافة تصدير" بالاعتماد على
    نص اللابل بتاعه (زي "نوع" أو "مدى الظهور")، وبعدين بيدوس على الأوبشن
    اللي نصه option_text من القايمة اللي بتفتح. بيرجع True لو نجح، False
    لو فشل (وبيسجل السبب في اللوج لو log مبعوت).

    ⚠️ الواجهة الجديدة (بعد تغيير نون) مبنية بمكتبة Base UI مش Ant Design
    زي الأول، فالـ combobox بقى <button role="combobox"> والأوبشنز بتفتح
    في قايمة منفصلة (popup) — استخدمنا role="option" كأساس، ولو مالقيناش
    بنرجع لأي عنصر نصه مطابق تمامًا كـ fallback.
    """
    try:
        trigger = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            f"//label[normalize-space(text())='{label_text}']"
            "/following::button[@role='combobox'][1]"
        )))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", trigger)
        time.sleep(0.2)
        driver.execute_script("arguments[0].click();", trigger)
        time.sleep(0.5)

        try:
            option = wait.until(EC.element_to_be_clickable((
                By.XPATH,
                f"//*[@role='option'][normalize-space(.)='{option_text}']"
            )))
        except TimeoutException:
            # fallback: أي عنصر ظاهر نصه مطابق تمامًا للأوبشن المطلوب
            option = wait.until(EC.element_to_be_clickable((
                By.XPATH,
                f"//*[normalize-space(text())='{option_text}']"
            )))
        driver.execute_script("arguments[0].click();", option)
        time.sleep(0.5)
        return True
    except Exception as e:
        if log:
            log_write(log, f"⚠️ فشل اختيار '{option_text}' في حقل '{label_text}': {str(e)[:150]}\n")
        return False


def _catalog_latest_export_status(driver):
    """
    يرجع (status_text, download_button_or_None) لأحدث صف تصدير في قايمة
    الكتالوج (بيفترض إن أحدث تصدير بيظهر في أول صف بالقايمة، زي أغلب
    الجداول دي في نون). لو مفيش أي تاج حالة لسه، بيرجع (None, None).

    ⚠️ الواجهة الجديدة: التاج بقى <span data-slot="badge"> بدل ما كان
    ant-design statusTag، وزرار التحميل بقى نصه "تحميل" مكتوب مباشرة
    جوه الـ <button> (من غير <span>) مع أيقونة svg قبله.
    """
    try:
        status_el = driver.find_element(
            By.XPATH, "(//span[@data-slot='badge'])[1]"
        )
    except Exception:
        return None, None

    status_text = status_el.text.strip()
    download_btn = None
    try:
        row = status_el.find_element(
            By.XPATH,
            "./ancestor::tr[1] | ./ancestor::*[contains(@class,'row') or contains(@class,'Row')][1]"
        )
        download_btn = row.find_element(
            By.XPATH, ".//button[contains(normalize-space(.), 'تحميل')]"
        )
    except Exception:
        pass
    return status_text, download_btn


_ARABIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")


def _catalog_find_recent_completed_export(driver, max_age_days=1):
    """
    بيدوّر في قايمة التصديرات كلها على أحدث تصدير حالته "مكتمل" وتاريخه النهارده
    أو امبارح (max_age_days=1) — زي "2026-10-17 08:06:03" — وعنده زرار "تحميل".
    بيرجع (download_button, date_text) أو (None, None) لو مفيش.
    القايمة بتبقى مرتبة من الأحدث للأقدم، فأول واحد مطابق هو الأحدث.
    """
    today = datetime.date.today()
    oldest_ok = today - datetime.timedelta(days=max_age_days)
    try:
        badges = driver.find_elements(By.XPATH, "//span[@data-slot='badge']")
    except Exception:
        return None, None

    for badge in badges:
        try:
            if "مكتمل" not in badge.text:
                continue
            try:
                row = badge.find_element(By.XPATH, "./ancestor::tr[1]")
            except Exception:
                row = badge.find_element(
                    By.XPATH,
                    "./ancestor::*[contains(@class,'row') or contains(@class,'Row')][1]"
                )
            row_text = row.text.translate(_ARABIC_DIGITS)
            m = re.search(r"(\d{4})-(\d{2})-(\d{2})(?:[ T]+(\d{2}:\d{2}:\d{2}))?", row_text)
            if not m:
                continue
            file_date = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            if file_date < oldest_ok:
                continue
            btn = row.find_element(
                By.XPATH, ".//button[contains(normalize-space(.), 'تحميل')]"
            )
            return btn, m.group(0)
        except Exception:
            continue
    return None, None


def _download_catalog_export_report_once(driver, store_code, log):
    if not store_code:
        log_write(log, "⚠️ مفيش كود متجر — تخطي تصدير الكتالوج\n")
        return None

    url = f"https://noon-catalog.noon.partners/ar/exports?project=PRJ{store_code}"
    log_write(log, f"\n📚 تصدير الكتالوج PRJ{store_code}...\n")
    wait = WebDriverWait(driver, 30)

    try:
        driver.get(url)
        time.sleep(4)
        _ensure_window_visible(driver)

        # ── فتح نافذة "إضافة تصدير" ──
        add_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//button[contains(normalize-space(.), 'إضافة تصدير')]"
        )))
        driver.execute_script("arguments[0].click();", add_btn)
        wait.until(EC.visibility_of_element_located((
            By.XPATH, "//h2[@data-slot='dialog-title' and normalize-space(text())='إضافة تصدير']"
        )))
        time.sleep(1)

        # ── اختيار النوع "تصدير الكتالوج" ──
        # الواجهة اتغيرت بالكامل (بقت مبنية بمكتبة Base UI بدل Ant Design)،
        # فبقينا بندوس على الـ combobox مباشرة عن طريق نص اللابل بتاعه.
        if not _select_combobox_option(driver, wait, "نوع", "تصدير الكتالوج", log):
            log_write(log, "❌ فشل اختيار النوع 'تصدير الكتالوج'\n")
            return None

        # ── حقل "مدى الظهور" بيظهر بس بعد اختيار النوع (حقل ديناميكي) ──
        try:
            wait.until(EC.visibility_of_element_located((
                By.XPATH, "//label[normalize-space(text())='مدى الظهور']"
            )))
        except TimeoutException:
            pass
        time.sleep(0.3)

        if not _select_combobox_option(driver, wait, "مدى الظهور", "نشط", log):
            log_write(log, "❌ فشل اختيار مدى الظهور 'نشط'\n")
            return None

        # ── زرار إنشاء ──
        create_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//button[@data-variant='solid' and normalize-space(text())='إنشاء']"
        )))
        driver.execute_script("arguments[0].click();", create_btn)
        log_write(log, "⏳ جاري إنشاء تصدير الكتالوج... (بياخد شوية وقت)\n")
        time.sleep(3)

        # ── الانتظار لحد ما يبقى "مكتمل" ──
        download_btn = None
        MAX_ATTEMPTS = 60   # 60 × 5s = 5 دقايق
        for attempt in range(1, MAX_ATTEMPTS + 1):
            time.sleep(5)
            _ensure_window_visible(driver)
            status_text, btn = _catalog_latest_export_status(driver)
            if status_text and "مكتمل" in status_text and btn is not None:
                log_write(log, "✅ تصدير الكتالوج جاهز\n")
                download_btn = btn
                break
            if status_text and attempt % 6 == 0:
                log_write(log, f"   🔄 الحالة الحالية: {status_text} — لسه بيجهز... ({attempt*5}s)\n")
            if attempt % 6 == 0:
                try:
                    driver.refresh()
                    time.sleep(3)
                except Exception:
                    pass

        if not download_btn:
            # ── بعد الـ 5 دقايق الأولى: بدل ما نعمل تصدير جديد من الأول،
            # نشوف لو فيه تصدير سابق "مكتمل" بتاريخ النهارده أو امبارح ونستخدمه ──
            log_write(log, "⚠️ التصدير الجديد لم يجهز خلال 5 دقايق — جاري البحث عن تصدير سابق مكتمل بتاريخ اليوم أو أمس...\n")
            try:
                driver.refresh()
                time.sleep(4)
                _ensure_window_visible(driver)
            except Exception:
                pass
            old_btn, old_date = _catalog_find_recent_completed_export(driver)
            if old_btn is not None:
                log_write(log, f"♻️ تم العثور على تصدير مكتمل سابق ({old_date}) — هيتم استخدامه\n")
                download_btn = old_btn
            else:
                log_write(log, "❌ لم يجهز تصدير الكتالوج في الوقت المحدد ولا يوجد تصدير سابق مكتمل بتاريخ حديث\n")
                return None

        # ── تحميل الملف (بيدخل مجلد التنزيلات العادي، زي تقرير الجرد) ──
        before = (set(glob.glob(os.path.join(DOWNLOADS_DIR, "*.csv")))
                  | set(glob.glob(os.path.join(DOWNLOADS_DIR, "*.xlsx"))))

        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", download_btn)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", download_btn)
        log_write(log, "✅ تم الضغط على تحميل\n⏳ جاري انتظار تحميل الملف...\n")

        file_path = None
        prev_size = -1
        stable    = 0
        for _ in range(60):
            time.sleep(1)
            after = (set(glob.glob(os.path.join(DOWNLOADS_DIR, "*.csv")))
                     | set(glob.glob(os.path.join(DOWNLOADS_DIR, "*.xlsx"))))
            new_files = after - before
            if new_files:
                file_path = max(new_files, key=os.path.getctime)
                try:
                    curr = os.path.getsize(file_path)
                    stable = (stable + 1) if curr == prev_size and curr > 0 else 0
                    if stable >= 2:
                        break
                    prev_size = curr
                except Exception:
                    pass

        if not file_path:
            log_write(log, "❌ لم يتم العثور على ملف تصدير الكتالوج\n")
            return None

        log_write(log, f"✅ تم تحميل: {os.path.basename(file_path)}\n")
        return file_path

    except Exception as e:
        log_write(log, f"❌ خطأ تصدير الكتالوج: {str(e)}\n")
        return None


def download_catalog_export_report(driver, store_code, log, max_retries=None):
    """
    غلاف حوالين _download_catalog_export_report_once: لو التنزيل فشل (رجّع None)،
    بيحاول تاني تلقائيًا وباستمرار لحد ما ينجح، من غير حد أقصى للمحاولات
    (زي الجرد والمبيعات بالظبط).
    """
    attempt = 0
    while True:
        attempt += 1
        _ensure_window_visible(driver)
        result = _download_catalog_export_report_once(driver, store_code, log)
        if result:
            return result
        log_write(log, f"🔁 إعادة محاولة تصدير الكتالوج (محاولة {attempt})...\n")
        time.sleep(2)


# ================= تجميع ملفات CSV في إكسيل واحد =================
def _read_csv_rows(path):
    try:
        with open(path, encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            rows = list(reader)
            fieldnames = list(reader.fieldnames or [])
        return rows, fieldnames
    except Exception:
        return [], []


def combine_csvs_to_excel(csv_paths, out_path, sheet_title, log):
    """
    يجمع كذا ملف CSV من نفس النوع في ملف إكسيل واحد، مع عمود "المصدر" يوضح
    اسم الملف الأصلي بتاع كل صف — عشان تعرف كل صف جاي من أنهي متجر/ملف.
    """
    if not csv_paths:
        log_write(log, f"⚠️ مفيش ملفات لتجميعها ({sheet_title})\n")
        return None

    all_fieldnames = []
    all_rows = []  # (source_file, row_dict)

    for p in csv_paths:
        rows, fieldnames = _read_csv_rows(p)
        for f in fieldnames:
            if f not in all_fieldnames:
                all_fieldnames.append(f)
        for r in rows:
            all_rows.append((os.path.basename(p), r))

    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_title[:31]  # أقصى طول لاسم شيت في إكسيل هو 31 حرف

        header_row = ["المصدر"] + all_fieldnames
        ws.append(header_row)

        hfill = PatternFill(start_color="F5A623", end_color="F5A623", fill_type="solid")
        hfont = Font(bold=True, color="000000", size=11)
        for col in range(1, len(header_row) + 1):
            c = ws.cell(row=1, column=col)
            c.fill = hfill
            c.font = hfont
            c.alignment = Alignment(horizontal="center", vertical="center")

        for source, row in all_rows:
            ws.append([source] + [row.get(f, "") for f in all_fieldnames])

        ws.column_dimensions["A"].width = 32
        wb.save(out_path)
        log_write(log, f"✅ تم إنشاء: {os.path.basename(out_path)} ({len(all_rows)} صف)\n")
        return out_path
    except Exception as e:
        log_write(log, f"❌ فشل حفظ {sheet_title}: {str(e)[:200]}\n💡 تأكد إن الملف مش مفتوح في Excel\n")
        return None


# ================= مزامنة Google Sheets (Inventory + DailyOrders) =================
"""
بتربط نفس ملفين الإكسيل المجمعين (جرد_مجمع_*.xlsx و مبيعات_مجمعة_*.xlsx) بنفس
الـ Google Sheet اللي بيستخدمه برنامج الطلبات (stock_requests) — شيت "Inventory"
وشيت "DailyOrders" — بنفس أسماء الأعمدة وترتيبها المتفق عليه هناك:

  Inventory  : SKU | Warehouse | Stock | Monthly Sales | Image URL | Date Uploaded
  DailyOrders: SKU | Order Timestamp | Status | Price | Quantity | Date Uploaded

ملاحظات مهمة:
- الـ "SKU" المستخدم هنا هو عمود sku (الكود الداخلي زي ZB287E2DFD92585AFAD6CZ-1)
  — مش عمود partner_sku (كود المنتج بتاع الـ seller نفسه زي N3222353A).
  لو ده غلط وعايز تستخدم عمود تاني، قولّي وأظبطه فورًا.
- "Monthly Sales" بيتحسب هنا تلقائيًا = عدد صفوف كل SKU في ملف المبيعات المجمع
  بعد ما نفلتر ونسيب بس الصفوف اللي fulfillment_model بتاعها
  "Fulfilled by Noon (FBN)" أو "Supermall" (أي نوع تاني زي FBP بيتجاهل تمامًا).
- روابط الصور (Image URL) بتتجاب أوتوماتيك من شيت "links n" (SKU → صورة)
  بنفس الطريقة اللي برنامج الطلبات شغال بيها.
- المزامنة بتمسح كل البيانات القديمة في الشيتين (مع الحفاظ على صف الهيدر)
  وبعدين تضيف الجديد كله دفعة واحدة.
"""

GSHEET_SCOPE = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

INVENTORY_HEADERS    = ["SKU", "Warehouse", "Stock", "Monthly Sales", "Image URL", "Date Uploaded"]
DAILY_ORDERS_HEADERS = ["SKU", "Order Timestamp", "Status", "Price", "Quantity", "Fulfillment Model", "Date Uploaded", "family"]

# شيتات مستخرجة من ملف إكسيل "تكويد نون" (TACWEED_XLSM_PATH):
TACWEED_HEADERS         = ["SKU", "Code01", "Total Cost", "Date Uploaded"]
WAREHOUSE_STOCK_HEADERS = ["Code", "Quantity", "Item Name", "Date Uploaded"]

# بس الأوردرز اللي fulfillment_model بتاعها من النوعين دول هي اللي بتتحسب/تتضاف —
# أي نوع تاني (زي Fulfilled by Partner (FBP)) بيتم تجاهله بالكامل.
ALLOWED_FULFILLMENT_MODELS = {"Fulfilled by Noon (FBN)", "Supermall", "Fulfilled by Partner (FBP)"}


def _gsheet_client(creds_path):
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, GSHEET_SCOPE)
    return gspread.authorize(creds)


def _gsheet_get_or_create_ws(ss, title, headers):
    try:
        ws = ss.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title=title, rows="2000", cols=str(max(len(headers), 6)))
        ws.append_row(headers)
    return ws


def _gsheet_clear_data_keep_header(ws):
    """بيمسح كل الصفوف ماعدا صف الهيدر الأول (صف 1)."""
    values = ws.get_all_values()
    if len(values) > 1:
        ws.delete_rows(2, len(values))


def _tw_num(v):
    """رقم أو None (فاضي / خطأ زي #VALUE! / مش رقم)."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace(",", "")
    if t == "" or t.startswith("#"):
        return None
    try:
        return float(t)
    except ValueError:
        return None


def _tw_retry(fn, log, tries=5):
    for a in range(tries):
        try:
            return fn()
        except Exception as e:
            if a == tries - 1:
                raise
            log_write(log, f"⚠️ محاولة {a + 1} فشلت: {str(e)[:100]}\n")
            time.sleep(65 if "429" in str(e) else 3 * (a + 1))


def _tw_col_letter(n):
    t = ""
    while n:
        n, r = divmod(n - 1, 26)
        t = chr(65 + r) + t
    return t


def _tacweed_upsert(ws, rows, log, chunk=20000):
    """تحديث تاب Tacweed بالـ SKU من غير مسح: الكود/التكلفة المتغيّرة بتتحدّث، والـ SKU الجديدة بتتضاف، والباقي زي ما هو.
    تكلفة فاضية أو خطأ (#VALUE!) في الملف = 0 مؤقتًا لو مفيش تكلفة قبلها (المنتج يتحسب بصفر زي المحاسب ويتبلّغ عنه)،
    ولو فيه تكلفة قديمة في الشيت بتتساب زي ما هي. أول ما التكلفة الحقيقية توصل في الملف بتحل مكان الصفر.
    rows: [[sku, code01, total_cost, date_uploaded, company(اختياري)], ...]. بيرجع (محدّث، مضاف، [SKU اتحطت صفر]).
    عمود Company (الشركة/المجموعة): لو مش موجود في التاب بيتضاف في آخر الهيدر، وبيتحدّث من الملف (الفاضي مش بيمسح القديم)."""
    fresh = {}
    for rw in rows:
        sku, code, cost, dtu = rw[:4]
        comp = str(rw[4]).strip() if len(rw) > 4 and rw[4] is not None else ""
        k = str(sku).strip().upper()
        if not k:
            continue
        c = _tw_num(cost)
        if k not in fresh or (fresh[k][1] is None and c is not None):
            fresh[k] = [("" if code is None else code), c, dtu, comp]
        elif comp and not fresh[k][3]:
            fresh[k][3] = comp
    probe = _tw_retry(lambda: ws.row_values(1), log)
    hdr = [str(h).strip().lower() for h in probe]
    cn = "cost" if "cost" in hdr else ("total cost" if "total cost" in hdr else None)
    if "sku" not in hdr or not cn:
        raise RuntimeError("تاب Tacweed مفيهوش عمودين SKU و Cost — مفيش تحديث اتعمل")
    i_sku, i_cost = hdr.index("sku"), hdr.index(cn)
    if "company" not in hdr and any(v[3] for v in fresh.values()):
        _tw_retry(lambda: ws.update(range_name=f"{_tw_col_letter(len(hdr) + 1)}1", values=[["Company"]], value_input_option="RAW"), log)
        hdr.append("company")
    i_comp = hdr.index("company") if "company" in hdr else None
    i_code = hdr.index("code01") if "code01" in hdr else None
    i_date = hdr.index("date uploaded") if "date uploaded" in hdr else None
    width = len(hdr)
    last = _tw_col_letter(width)
    vals = _tw_retry(lambda: ws.get(f"A1:{last}", value_render_option="UNFORMATTED_VALUE"), log)
    nrows = len(vals)
    col = {}
    for name, i in (("cost", i_cost), ("code", i_code), ("date", i_date), ("comp", i_comp)):
        if i is not None:
            col[name] = [(r[i] if i < len(r) else "") for r in vals]
    seen, changed, zero_list = set(), {}, []
    for r in range(2, nrows + 1):
        row = vals[r - 1]
        k = str(row[i_sku]).strip().upper() if i_sku < len(row) else ""
        if not k or k not in fresh or k in seen:
            continue
        seen.add(k)
        code, newc, dtu, comp = fresh[k]
        oldc = _tw_num(col["cost"][r - 1])
        target = newc if newc is not None else (oldc if oldc is not None else 0.0)
        if newc is None and target == 0.0:
            zero_list.append(k)
        ch = oldc is None or abs(oldc - target) > 0.004
        if i_code is not None and str(code).strip() != "" and str(col["code"][r - 1]).strip() != str(code).strip():
            col["code"][r - 1] = code
            ch = True
        if i_comp is not None and comp and str(col["comp"][r - 1]).strip() != comp:
            col["comp"][r - 1] = comp
            ch = True
        if oldc is None or abs(oldc - target) > 0.004:
            col["cost"][r - 1] = target
        if ch:
            if i_date is not None:
                col["date"][r - 1] = dtu
            changed[r] = True
    if changed:
        lo, hi = min(changed), max(changed)
        for name, i in (("cost", i_cost), ("code", i_code), ("date", i_date), ("comp", i_comp)):
            if i is None:
                continue
            L = _tw_col_letter(i + 1)
            for a in range(lo, hi + 1, chunk):
                b = min(a + chunk - 1, hi)
                part = [[x] for x in col[name][a - 1:b]]
                _tw_retry(lambda p=part, a=a, b=b, L=L: ws.update(range_name=f"{L}{a}:{L}{b}", values=p, value_input_option="RAW"), log)
                time.sleep(1.5)
    new_rows = []
    for k, (code, newc, dtu, comp) in fresh.items():
        if k in seen:
            continue
        row = [""] * width
        row[i_sku] = k
        row[i_cost] = newc if newc is not None else 0.0
        if newc is None:
            zero_list.append(k)
        if i_code is not None:
            row[i_code] = code
        if i_date is not None:
            row[i_date] = dtu
        if i_comp is not None:
            row[i_comp] = comp
        new_rows.append(row)
    if new_rows:
        _gsheet_append_in_batches(ws, new_rows, log)
    return len(changed), len(new_rows), zero_list


def _gsheet_links_map(ss):
    """SKU (upper) -> Image URL من شيت 'links n' — لو مش موجود بيرجع dict فاضي."""
    m = {}
    try:
        ws = ss.worksheet("links n")
        for row in ws.get_all_values()[1:]:
            if len(row) >= 2 and row[0].strip():
                m[row[0].strip().upper()] = row[1].strip()
    except Exception:
        pass
    return m


def _read_xlsx_as_dicts(path):
    """بيرجع list of dict لكل صف في أول Sheet من ملف إكسيل، مفاتيح الـ dict = أسماء الأعمدة (زي ما هي)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    headers = [str(h).strip() if h is not None else "" for h in next(rows_iter)]
    out = []
    for r in rows_iter:
        d = {}
        for i, h in enumerate(headers):
            if not h:
                continue
            d[h] = r[i] if i < len(r) else None
        out.append(d)
    wb.close()
    return out


def _read_named_sheet_as_dicts(path, sheet_name):
    """
    زي _read_xlsx_as_dicts بالظبط، بس بتاخد اسم شيت محدد جوه ملف الإكسيل
    بدل أول شيت تلقائيًا (مطلوبة لملف تكويد نون اللي فيه أكتر من شيت زي
    "Noon" و "الجرد"). بترجع None لو الشيت مش موجود جوه الملف.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return None
    ws = wb[sheet_name]
    rows_iter = ws.iter_rows(values_only=True)
    headers = [str(h).strip() if h is not None else "" for h in next(rows_iter)]
    out = []
    for r in rows_iter:
        d = {}
        for i, h in enumerate(headers):
            if not h:
                continue
            d[h] = r[i] if i < len(r) else None
        out.append(d)
    wb.close()
    return out


def _gsheet_append_in_batches(ws, rows, log, batch_size=2000):
    """بيضيف صفوف كتير على دفعات — عشان نتفادى أي حد أقصى لحجم الطلب مع Google Sheets API،
    وبيوسّع (resize) الشيت الأول لو عدد الصفوف المطلوب أكبر من مساحته الحالية."""
    if not rows:
        return
    needed_rows = len(rows) + 10
    try:
        if ws.row_count < needed_rows:
            ws.resize(rows=needed_rows)
    except Exception as e:
        log_write(log, f"⚠️ تعذر توسيع شيت {ws.title}: {str(e)[:150]}\n")

    for i in range(0, len(rows), batch_size):
        chunk = rows[i:i + batch_size]
        for attempt in range(5):
            try:
                ws.append_rows(chunk, value_input_option="USER_ENTERED")
                break
            except gspread.exceptions.APIError as e:
                if "429" in str(e) or "quota" in str(e).lower():
                    time.sleep(2 * (2 ** attempt))
                else:
                    time.sleep(1.5)
            except Exception:
                time.sleep(1.5)
        log_write(log, f"   ⬆️ {ws.title}: {min(i + batch_size, len(rows))}/{len(rows)} صف\n")


def sync_reports_to_google_sheets(inv_xlsx_path, sales_xlsx_path, settings, log):
    """
    بتاخد ملفي الإكسيل المجمعين (الجرد + المبيعات) وتحدّث بيهم شيتي
    Inventory و DailyOrders في Google Sheets — مسح كل القديم وإضافة الجديد.
    """
    creds_path = _resolve_gsheet_creds_path()
    sheet_name = GSHEET_NAME

    if not creds_path:
        log_write(log, f"⚠️ مزامنة Google Sheets متخطاة — مفيش ملف .json اعتماد (Service Account) موجود في: {GSHEET_CREDS_DIR}\n")
        return
    if not os.path.exists(creds_path):
        log_write(log, f"❌ ملف اعتماد Google غير موجود: {creds_path}\n")
        return
    if not inv_xlsx_path and not sales_xlsx_path:
        log_write(log, "⚠️ مفيش ملفات جرد/مبيعات لمزامنتها مع Google Sheets\n")
        return

    log_write(log, f"\n🔗 جاري الاتصال بـ Google Sheets ({sheet_name})...\n")
    try:
        client = _gsheet_client(creds_path)
        ss = client.open(sheet_name)
    except Exception as e:
        log_write(log, f"❌ فشل الاتصال بـ Google Sheets: {str(e)[:300]}\n")
        return

    links_map = _gsheet_links_map(ss)
    date_uploaded = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # ── 1) قراءة ملف المبيعات وحساب "Monthly Sales" لكل SKU + صفوف DailyOrders ──
    sales_counts = {}
    daily_orders_rows = []
    if sales_xlsx_path and os.path.exists(sales_xlsx_path):
        try:
            log_write(log, "📖 قراءة ملف المبيعات المجمع...\n")
            sales_rows = _read_xlsx_as_dicts(sales_xlsx_path)
            skipped_fm = 0
            for r in sales_rows:
                fm = str(r.get("fulfillment_model") or "").strip()
                if fm not in ALLOWED_FULFILLMENT_MODELS:
                    skipped_fm += 1
                    continue
                sku = str(r.get("sku") or "").strip()
                if not sku:
                    continue
                sales_counts[sku] = sales_counts.get(sku, 0) + 1
                ts     = str(r.get("order_timestamp") or "").strip()
                status = str(r.get("status") or "").strip()
                price  = r.get("offer_price")
                price  = "" if price is None else price
                family = str(r.get("family") or "").strip()
                daily_orders_rows.append([sku, ts, status, price, 1, fm, date_uploaded, family])
            log_write(log, f"✅ {len(daily_orders_rows)} صف مبيعات (FBN/Supermall بس) جاهزين، {len(sales_counts)} SKU مختلف — اتجاهل {skipped_fm} صف من fulfillment models تانية\n")
        except Exception as e:
            log_write(log, f"❌ فشل قراءة ملف المبيعات: {str(e)[:200]}\n")

    # ── 2) قراءة ملف الجرد وتجميع الكمية لكل (SKU, مخزن) ──
    inv_agg = {}  # (sku, warehouse) -> qty
    if inv_xlsx_path and os.path.exists(inv_xlsx_path):
        try:
            log_write(log, "📖 قراءة ملف الجرد المجمع...\n")
            inv_rows = _read_xlsx_as_dicts(inv_xlsx_path)
            for r in inv_rows:
                sku = str(r.get("sku") or "").strip()
                wh  = str(r.get("warehouse_code") or "").strip()
                if not sku:
                    continue
                try:
                    qty = int(float(r.get("qty") or 0))
                except Exception:
                    qty = 0
                key = (sku, wh)
                inv_agg[key] = inv_agg.get(key, 0) + qty
            log_write(log, f"✅ {len(inv_agg)} صف جرد (SKU × مخزن) جاهزين\n")
        except Exception as e:
            log_write(log, f"❌ فشل قراءة ملف الجرد: {str(e)[:200]}\n")

    inventory_rows = []
    for (sku, wh), qty in inv_agg.items():
        img = links_map.get(sku.upper(), "")
        monthly_sales = sales_counts.get(sku, 0)
        inventory_rows.append([sku, wh, qty, monthly_sales, img, date_uploaded])

    # ── 3) رفع Inventory (مسح القديم + إضافة الجديد) ──
    if inventory_rows:
        try:
            inv_ws = _gsheet_get_or_create_ws(ss, "Inventory", INVENTORY_HEADERS)
            log_write(log, f"🗑️ مسح بيانات Inventory القديمة...\n")
            _gsheet_clear_data_keep_header(inv_ws)
            log_write(log, f"⬆️ رفع {len(inventory_rows)} صف جرد جديد لشيت Inventory...\n")
            _gsheet_append_in_batches(inv_ws, inventory_rows, log)
            log_write(log, f"✅ تم تحديث شيت Inventory بالكامل\n")
        except Exception as e:
            log_write(log, f"❌ فشل رفع الجرد لـ Google Sheets: {str(e)[:250]}\n")

    # ── 4) رفع DailyOrders (مسح القديم + إضافة الجديد) ──
    if daily_orders_rows:
        try:
            do_ws = _gsheet_get_or_create_ws(ss, "DailyOrders", DAILY_ORDERS_HEADERS)
            log_write(log, f"🗑️ مسح بيانات DailyOrders القديمة...\n")
            _gsheet_clear_data_keep_header(do_ws)
            log_write(log, f"⬆️ رفع {len(daily_orders_rows)} صف مبيعات جديد لشيت DailyOrders...\n")
            _gsheet_append_in_batches(do_ws, daily_orders_rows, log)
            log_write(log, f"✅ تم تحديث شيت DailyOrders بالكامل\n")
        except Exception as e:
            log_write(log, f"❌ فشل رفع المبيعات لـ Google Sheets: {str(e)[:250]}\n")

    log_write(log, "🔗 اكتملت مزامنة Google Sheets\n")


# ================= مزامنة شيت LIVE (تصدير الكتالوج لكل المتاجر) =================
def _read_report_rows_generic(path):
    """يقرا أي ملف تقرير CSV أو XLSX ويرجع (rows_as_dicts, fieldnames)."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        rows = _read_xlsx_as_dicts(path)
        fieldnames = list(rows[0].keys()) if rows else []
        return rows, fieldnames
    return _read_csv_rows(path)


def sync_catalog_to_google_sheets(catalog_files, settings, log):
    """
    catalog_files: [(store_name, file_path), ...] بنفس ترتيب المتاجر.
    بيرفع ملفات تصدير الكتالوج بتاعة كل المتاجر في شيت واحد اسمه "LIVE" —
    كل متجر بيتحط تحت التاني في نفس الشيت (Stacked)، مع عمود "Store" في
    الأول يوضح مصدر كل صف. بيمسح بيانات LIVE القديمة كل تشغيلة ويحط
    الجديد بدالها بالكامل، بنفس أسلوب مزامنة Inventory/DailyOrders.
    """
    creds_path = _resolve_gsheet_creds_path()
    sheet_name = GSHEET_NAME

    if not creds_path:
        log_write(log, f"⚠️ مزامنة LIVE (الكتالوج) متخطاة — مفيش ملف .json اعتماد (Service Account) موجود في: {GSHEET_CREDS_DIR}\n")
        return
    if not os.path.exists(creds_path):
        log_write(log, f"❌ ملف اعتماد Google غير موجود: {creds_path}\n")
        return
    if not catalog_files:
        log_write(log, "⚠️ مفيش ملفات كتالوج لمزامنتها مع شيت LIVE\n")
        return

    log_write(log, f"\n🔗 جاري رفع تصدير الكتالوج لشيت LIVE ({sheet_name})...\n")
    try:
        client = _gsheet_client(creds_path)
        ss = client.open(sheet_name)
    except Exception as e:
        log_write(log, f"❌ فشل الاتصال بـ Google Sheets: {str(e)[:300]}\n")
        return

    all_fieldnames = []
    all_rows = []  # (store_name, row_dict)
    for store_name, path in catalog_files:
        try:
            rows, fieldnames = _read_report_rows_generic(path)
        except Exception as e:
            log_write(log, f"❌ فشل قراءة ملف كتالوج {store_name}: {str(e)[:150]}\n")
            continue
        for f in fieldnames:
            if f not in all_fieldnames:
                all_fieldnames.append(f)
        for r in rows:
            all_rows.append((store_name, r))

    if not all_rows:
        log_write(log, "⚠️ ملفات الكتالوج فاضية — مفيش صفوف لرفعها لشيت LIVE\n")
        return

    date_uploaded = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    headers = ["Store"] + all_fieldnames + ["Date Uploaded"]
    sheet_rows = [
        [store_name] + [row.get(f, "") for f in all_fieldnames] + [date_uploaded]
        for store_name, row in all_rows
    ]

    try:
        live_ws = _gsheet_get_or_create_ws(ss, "LIVE", headers)
        # لو أعمدة تصدير الكتالوج اتغيرت عن آخر مرة، بنحدّث صف الهيدر معاها
        current_header = live_ws.row_values(1)
        if current_header != headers:
            live_ws.update("A1", [headers])
        log_write(log, "🗑️ مسح بيانات LIVE القديمة...\n")
        _gsheet_clear_data_keep_header(live_ws)
        log_write(log, f"⬆️ رفع {len(sheet_rows)} صف كتالوج (كل المتاجر) لشيت LIVE...\n")
        _gsheet_append_in_batches(live_ws, sheet_rows, log)
        log_write(log, "✅ تم تحديث شيت LIVE بالكامل\n")
    except Exception as e:
        log_write(log, f"❌ فشل رفع الكتالوج لشيت LIVE: {str(e)[:250]}\n")


# ================= إعلانات نون (Ad Manager) — يومي (أمس) + شهري =================
"""
لكل متجر: يفتح admanager.noon.partners على تاب الحملات بمدى تواريخ (أمس / من أول الشهر لحد أمس)،
يضغط «تصدير» وينتظر ملف campaigns-export-*.xlsx في مجلد التنزيلات، وبعدين بيرفع صفوف تاب «SKUs»
(أداء كل SKU في كل حملة) لشيتين في Google Sheets: AdsDaily (أمس بس) و AdsMonthly (من أول الشهر لحد أمس).
كل تشغيلة بتمسح الشيتين وتكتب الجديد بدالهم (مفيش تاريخ بيتراكم).
⚠️ كود الـ advertiser (AC_...) اختياري: لو المتجر فيه مفتاح "advertiser_code" بيتحط في الرابط، وإلا
بيتفتح الرابط من غيره (نون غالباً بتختاره لوحدها). ⚠️ لو الرابط من غيره ما اشتغلش مع متجر، اكتب كوده
في noon_settings.json جوه المتجر: "advertiser_code": "AC_XXXX".
"""

ADS_DAILY_SHEET   = "AdsDaily"
ADS_MONTHLY_SHEET = "AdsMonthly"
ADS_FIXED_HEADERS = ["Store", "Period", "Start Date", "End Date"]
ADS_TAIL_HEADERS  = ["Ad Type", "Date Uploaded"]


def _ads_periods():
    """(يومي, شهري) = [(اسم, بداية, نهاية)]: أمس، ومن أول شهر أمس لحد أمس."""
    y = datetime.date.today() - datetime.timedelta(days=1)
    return [("daily", y, y), ("monthly", y.replace(day=1), y)]


def _ads_url(store_code, advertiser_code, start, end):
    url = f"https://admanager.noon.partners/ar-sa/home?project=PRJ{store_code}"
    if advertiser_code:
        url += f"&advertiserCode={advertiser_code}"
    return (url + f"&mpCode=noon&tab=campaigns&campaignsStart={start.isoformat()}&campaignsEnd={end.isoformat()}")


def _download_ads_export_once(driver, store_code, advertiser_code, start, end, log):
    """يفتح صفحة الحملات للمدى ده ويضغط «تصدير» وينتظر ملف xlsx جديد. يرجّع المسار أو None."""
    url = _ads_url(store_code, advertiser_code, start, end)
    wait = WebDriverWait(driver, 30)
    before = set(glob.glob(os.path.join(DOWNLOADS_DIR, "*.xlsx")))
    try:
        driver.get(url)
        time.sleep(5)
        _ensure_window_visible(driver)
        btn_xpath = "//button[.//span[normalize-space(text())='تصدير']]"
        export_btn = wait.until(EC.element_to_be_clickable((By.XPATH, btn_xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", export_btn)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", export_btn)
        log_write(log, "   ✅ ضغطت تصدير — مستني الملف...\n")

        path, prev_size, stable = None, -1, 0
        for i in range(60):
            time.sleep(1)
            new = [p for p in set(glob.glob(os.path.join(DOWNLOADS_DIR, "*.xlsx"))) - before
                   if not p.lower().endswith(".crdownload")]
            if new:
                path = max(new, key=os.path.getctime)
                try:
                    cur = os.path.getsize(path)
                    stable = (stable + 1) if cur == prev_size and cur > 0 else 0
                    if stable >= 2:
                        return path
                    prev_size = cur
                except Exception:
                    pass
            elif i == 20:   # لو «تصدير» فتحت قايمة/نافذة بدل التنزيل المباشر: ندوس على زرار تصدير/تنزيل جواها
                try:
                    inner = driver.find_elements(By.XPATH,
                        "//div[contains(@class,'ant-modal') or contains(@class,'ant-dropdown') or contains(@class,'ant-popover')]"
                        "//*[self::button or @role='menuitem' or self::li][normalize-space(.)='تصدير' or normalize-space(.)='تنزيل' or normalize-space(.)='تحميل']")
                    for el in inner:
                        if el.is_displayed():
                            driver.execute_script("arguments[0].click();", el)
                            break
                except Exception:
                    pass
        return path
    except TimeoutException:
        cur = driver.current_url
        why = "الصفحة طلبت تسجيل دخول" if "login" in cur.lower() or "auth" in cur.lower() else "زرار «تصدير» ما ظهرش"
        log_write(log, f"   ❌ {why} ({cur[:90]})\n")
        return None
    except Exception as e:
        log_write(log, f"   ❌ خطأ تصدير الإعلانات: {str(e)[:200]}\n")
        return None


def download_ads_exports(driver, store, log, tries=2):
    """يرجّع [(period, start, end, path)] لتصدير أمس + الشهر لمتجر واحد. (محاولات محدودة عشان ما نعلقش المتجر.)"""
    code = store.get("store_code", "")
    ac = (store.get("advertiser_code") or "").strip()
    out = []
    for period, s, e in _ads_periods():
        log_write(log, f"\n📣 إعلانات PRJ{code} — {'أمس' if period == 'daily' else 'الشهر'} ({s} → {e})...\n")
        path = None
        for attempt in range(1, tries + 1):
            path = _download_ads_export_once(driver, code, ac, s, e, log)
            if path:
                log_write(log, f"   ✅ تم تحميل: {os.path.basename(path)}\n")
                break
            log_write(log, f"   🔁 إعادة محاولة ({attempt}/{tries})...\n")
        if path:
            out.append((period, s, e, path))
    return out


def ads_export_period(path):
    """الفترة اللي الملف فعلاً بيغطيها، من ورقة Read me ("Performance data is for 01 Nov 2026 to 04 Nov 2026").
    يرجّع (بداية, نهاية) كـ date، أو None لو ما اتقرتش."""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            for row in wb["Read me"].iter_rows(min_row=1, max_row=6, values_only=True):
                for v in row:
                    m = re.search(r"for\s+(\d{1,2}\s+\w{3}\s+\d{4})\s+to\s+(\d{1,2}\s+\w{3}\s+\d{4})", str(v or ""))
                    if m:
                        d = [datetime.datetime.strptime(x, "%d %b %Y").date() for x in m.groups()]
                        return d[0], d[1]
        finally:
            wb.close()
    except Exception:
        pass
    return None


def read_ads_export(path):
    """بيقرا تاب SKUs من ملف تصدير الإعلانات (الهيدر في الصف 2) ويضيف Ad Type من تاب Campaigns.
    يرجّع (أسماء_الأعمدة, [dict]) — من غير عمود Operation."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        def table(name):
            if name not in wb.sheetnames:
                return [], []
            rows = list(wb[name].iter_rows(values_only=True))
            if len(rows) < 2:
                return [], []
            hdr = [str(h).strip() if h is not None else "" for h in rows[1]]
            data = []
            for r in rows[2:]:
                if not any(v not in (None, "") for v in r):
                    continue
                data.append({h: (r[i] if i < len(r) else None) for i, h in enumerate(hdr) if h})
            return [h for h in hdr if h], data
        ch, camp = table("Campaigns")
        sh, skus = table("SKUs")
    finally:
        wb.close()
    types = {c.get("Campaign ID"): c.get("Ad Type") for c in camp}
    for r in skus:
        r["Ad Type"] = types.get(r.get("Campaign ID"), "")
    return [h for h in sh if h != "Operation"], skus


def _ads_cell(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return "" if v is None else v


def sync_ads_to_google_sheets(ads_files, settings, log):
    """ads_files = [(store_name, period, start, end, path)]. كل شيت (AdsDaily / AdsMonthly) بيتمسح بالكامل
    وبيتكتب فيه الجديد بدل القديم — بيحصل بس لو فيه صفوف جديدة (عشان فشل التنزيل ما يفضّيش الشيت)."""
    if not ads_files:
        log_write(log, "⚠️ مفيش ملفات إعلانات لرفعها\n")
        return
    creds_path = _resolve_gsheet_creds_path()
    if not creds_path or not os.path.exists(creds_path):
        log_write(log, f"⚠️ رفع الإعلانات متخطّى — مفيش ملف اعتماد Google في: {GSHEET_CREDS_DIR}\n")
        return
    try:
        ss = _gsheet_client(creds_path).open(GSHEET_NAME)
    except Exception as e:
        log_write(log, f"❌ فشل الاتصال بـ Google Sheets (إعلانات): {str(e)[:250]}\n")
        return

    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for period, sheet_title in (("daily", ADS_DAILY_SHEET), ("monthly", ADS_MONTHLY_SHEET)):
        new_rows, cols = [], []
        for store_name, p, s, e, path in ads_files:
            if p != period:
                continue
            got = ads_export_period(path)
            if got and got != (s, e):    # نون مطبّقتش المدى المطلوب — ما نحطش بيانات فترة غلط
                log_write(log, f"⚠️ {store_name} ({p}): الملف لفترة {got[0]} → {got[1]} مش {s} → {e} — اتخطّى\n")
                continue
            try:
                hdr, rows = read_ads_export(path)
            except Exception as ex:
                log_write(log, f"❌ فشل قراءة إعلانات {store_name} ({p}): {str(ex)[:150]}\n")
                continue
            for h in hdr:
                if h not in cols:
                    cols.append(h)
            for r in rows:
                new_rows.append((store_name, s, e, r))
        if not new_rows:
            log_write(log, f"⚠️ {sheet_title}: مفيش صفوف إعلانات جديدة\n")
            continue
        try:
            ws = _gsheet_get_or_create_ws(ss, sheet_title, ADS_FIXED_HEADERS + cols + ADS_TAIL_HEADERS)
            headers = ADS_FIXED_HEADERS + cols + ADS_TAIL_HEADERS
            added = [[sn, "Daily" if period == "daily" else "Monthly", s.isoformat(), e.isoformat()]
                     + [_ads_cell(r.get(h)) for h in cols] + [_ads_cell(r.get("Ad Type")), stamp]
                     for sn, s, e, r in new_rows]
            out = [headers] + added          # مسح كامل واستبدال بالجديد (من غير تراكم تاريخ)
            ws.clear()
            if ws.row_count < len(out) + 10:
                ws.resize(rows=len(out) + 10, cols=max(ws.col_count, len(headers)))
            elif ws.col_count < len(headers):
                ws.resize(rows=ws.row_count, cols=len(headers))
            for i in range(0, len(out), 2000):
                ws.update(range_name=f"A{i + 1}", values=out[i:i + 2000], value_input_option="RAW")
            log_write(log, f"✅ {sheet_title}: اتضاف {len(added)} صف (إجمالي {len(out) - 1})\n")
        except Exception as ex:
            log_write(log, f"❌ فشل رفع {sheet_title}: {str(ex)[:250]}\n")


# ================= تاريخ الإعلانات (خفيف) — AdsHistory + AdsCampaignHist =================
"""
بيحفظ تاريخ يومي صغير عشان التطبيق يحسب الترند والمقارنة قبل/بعد وتوصيات الإعلانات:
- AdsHistory: لكل (متجر + يوم + SKU) مجموع الحملات — بس الصفوف اللي فيها نشاط (مفيش صفوف أصفار).
- AdsCampaignHist: لكل (متجر + يوم + حملة): الميزانية اليومية والحالة والأداء — بس الحملات اللي فيها نشاط.
مبيتمسحش: بيستبدل بس (المتجر + اليوم) اللي اتنزل دلوقتي، وبيقلّم أي صف أقدم من ADS_HISTORY_KEEP_DAYS يوم.
بيشتغل من ملفات «يومي» بس (أمس) وبعد التأكد إن فترة الملف هي اليوم المطلوب.
"""

ADS_HISTORY_SHEET      = "AdsHistory"
ADS_CAMP_HISTORY_SHEET = "AdsCampaignHist"
ADS_HISTORY_KEEP_DAYS  = 120
ADS_HISTORY_HEADERS = ["Store", "Date", "SKU", "Campaigns", "Views", "Clicks", "Orders", "ATC", "Spends", "Revenue", "Date Uploaded"]
ADS_CAMP_HISTORY_HEADERS = ["Store", "Date", "Campaign ID", "Campaign Name", "Ad Type", "Serving Status", "Status",
                            "Daily Budget", "Bidding Strategy", "Top of Search Bid %", "Views", "Clicks", "Orders",
                            "ATC", "Spends", "Revenue", "Date Uploaded"]


def _ads_num(v):
    try:
        return float(str(v).replace(",", "").strip()) if v not in (None, "") else 0.0
    except Exception:
        return 0.0


def read_ads_history_rows(path, store_name, day):
    """من ملف يومي واحد: (صفوف AdsHistory, صفوف AdsCampaignHist) — الصفوف الصفرية بتتشال."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        def table(name):
            if name not in wb.sheetnames:
                return []
            rows = list(wb[name].iter_rows(values_only=True))
            if len(rows) < 3:
                return []
            hdr = [str(h).strip() if h is not None else "" for h in rows[1]]
            return [{h: (r[i] if i < len(r) else None) for i, h in enumerate(hdr) if h}
                    for r in rows[2:] if any(v not in (None, "") for v in r)]
        skus, camps = table("SKUs"), table("Campaigns")
    finally:
        wb.close()
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    d = day.isoformat()
    agg = {}
    for r in skus:
        sku = str(r.get("SKU") or "").strip().upper()
        if not sku:
            continue
        a = agg.setdefault(sku, {"c": set(), "Views": 0.0, "Clicks": 0.0, "Orders": 0.0, "ATC": 0.0, "Spends": 0.0, "Revenue": 0.0})
        a["c"].add(r.get("Campaign ID"))
        for k in ("Views", "Clicks", "Orders", "ATC", "Spends", "Revenue"):
            a[k] += _ads_num(r.get(k))
    sku_rows = []
    for sku, a in sorted(agg.items()):
        if a["Views"] == 0 and a["Clicks"] == 0 and a["Orders"] == 0 and a["Spends"] == 0:
            continue
        sku_rows.append([store_name, d, sku, len(a["c"]), a["Views"], a["Clicks"], a["Orders"], a["ATC"],
                         round(a["Spends"], 2), round(a["Revenue"], 2), stamp])
    camp_rows = []
    for c in camps:
        if _ads_num(c.get("Views")) == 0 and _ads_num(c.get("Spends")) == 0:
            continue
        camp_rows.append([store_name, d, c.get("Campaign ID") or "", c.get("Campaign Name") or "", c.get("Ad Type") or "",
                          c.get("Serving Status") or "", c.get("Status") or "", _ads_cell(c.get("Daily Budget")),
                          c.get("Bidding Strategy") or "", _ads_cell(c.get("Top of Search Bid %")),
                          _ads_num(c.get("Views")), _ads_num(c.get("Clicks")), _ads_num(c.get("Orders")), _ads_num(c.get("ATC")),
                          round(_ads_num(c.get("Spends")), 2), round(_ads_num(c.get("Revenue")), 2), stamp])
    return sku_rows, camp_rows


def _upsert_history(ss, title, headers, new_rows, keys, log):
    """يستبدل صفوف (Store, Date) الجديدة، يقلّم الأقدم من ADS_HISTORY_KEEP_DAYS، ويكتب الشيت."""
    ws = _gsheet_get_or_create_ws(ss, title, headers)
    existing = ws.get_all_values()
    old_hdr = [h for h in (existing[0] if existing else []) if h]
    cutoff = (datetime.date.today() - datetime.timedelta(days=ADS_HISTORY_KEEP_DAYS)).isoformat()
    kept = []
    for row in existing[1:]:
        d = {old_hdr[i]: (row[i] if i < len(row) else "") for i in range(len(old_hdr))}
        if (d.get("Store"), d.get("Date")) in keys or (d.get("Date") or "9999") < cutoff:
            continue
        kept.append([d.get(h, "") for h in headers])
    out = [headers] + kept + new_rows
    ws.clear()
    if ws.row_count < len(out) + 50:
        ws.resize(rows=len(out) + 50, cols=max(ws.col_count, len(headers)))
    for i in range(0, len(out), 2000):
        ws.update(range_name=f"A{i + 1}", values=out[i:i + 2000], value_input_option="RAW")
    log_write(log, f"✅ {title}: +{len(new_rows)} صف (إجمالي {len(out) - 1})\n")


def sync_ads_history_to_google_sheets(ads_files, log):
    """ads_files = [(store_name, period, start, end, path)] — بياخد اليومي بس."""
    daily = [(sn, s, p) for sn, per, s, e, p in ads_files if per == "daily"]
    if not daily:
        return
    creds_path = _resolve_gsheet_creds_path()
    if not creds_path or not os.path.exists(creds_path):
        return
    try:
        ss = _gsheet_client(creds_path).open(GSHEET_NAME)
    except Exception as e:
        log_write(log, f"❌ فشل الاتصال بـ Google Sheets (تاريخ الإعلانات): {str(e)[:200]}\n")
        return
    sku_rows, camp_rows, keys = [], [], set()
    for sn, day, path in daily:
        got = ads_export_period(path)
        if got and got != (day, day):
            log_write(log, f"⚠️ {sn}: ملف التاريخ لفترة {got[0]} → {got[1]} مش {day} — اتخطّى\n")
            continue
        try:
            a, b = read_ads_history_rows(path, sn, day)
        except Exception as ex:
            log_write(log, f"❌ فشل قراءة تاريخ إعلانات {sn}: {str(ex)[:150]}\n")
            continue
        sku_rows += a
        camp_rows += b
        keys.add((sn, day.isoformat()))
    if not keys:
        return
    try:
        _upsert_history(ss, ADS_HISTORY_SHEET, ADS_HISTORY_HEADERS, sku_rows, keys, log)
        _upsert_history(ss, ADS_CAMP_HISTORY_SHEET, ADS_CAMP_HISTORY_HEADERS, camp_rows, keys, log)
    except Exception as ex:
        log_write(log, f"❌ فشل رفع تاريخ الإعلانات: {str(ex)[:250]}\n")


# ================= مزامنة شيتات Tacweed / تكلفة البيع / WarehouseStock =================
def sync_tacweed_to_google_sheets(settings, log, only_jard=False):
    """
    بتفتح ملف إكسيل "تكويد نون" من المسار الثابت (TACWEED_XLSM_PATH) وتحدّث
    بيه 3 شيتات في نفس الـ Google Sheet (GSHEET_NAME) — بنفس أسلوب مسح
    القديم (مع الحفاظ على صف الهيدر) وإضافة الجديد كله دفعة واحدة:

      Tacweed      : SKU | Code01 | Total Cost | Date Uploaded — من شيت "Noon" (SKU NooN + الكود 01 + اجمالي التكلفة)
      WarehouseStock: Code | Quantity | Item Name | Date Uploaded — من شيت "الجرد" (رمز الباركود + الكمية + اسم المادة)
    """
    creds_path = _resolve_gsheet_creds_path()
    sheet_name = GSHEET_NAME

    if not creds_path:
        log_write(log, f"⚠️ مزامنة تكويد نون متخطاة — مفيش ملف .json اعتماد (Service Account) موجود في: {GSHEET_CREDS_DIR}\n")
        return
    if not os.path.exists(creds_path):
        log_write(log, f"❌ ملف اعتماد Google غير موجود: {creds_path}\n")
        return
    if not os.path.exists(TACWEED_XLSM_PATH):
        log_write(log, f"⚠️ مزامنة تكويد نون متخطاة — ملف التكويد غير موجود في: {TACWEED_XLSM_PATH}\n")
        return

    log_write(log, f"\n🔗 جاري قراءة ملف تكويد نون ({os.path.basename(TACWEED_XLSM_PATH)})...\n")
    try:
        noon_rows = _read_named_sheet_as_dicts(TACWEED_XLSM_PATH, TACWEED_SHEET_NOON)
    except Exception as e:
        noon_rows = None
        log_write(log, f"❌ فشل قراءة شيت \"{TACWEED_SHEET_NOON}\" من ملف تكويد نون: {str(e)[:200]}\n")
    if noon_rows is None:
        log_write(log, f"⚠️ شيت \"{TACWEED_SHEET_NOON}\" غير موجود جوه ملف تكويد نون\n")
    if only_jard:      # تحديث الجرد (WarehouseStock) بس — من غير Tacweed
        noon_rows = []

    try:
        jard_rows = _read_named_sheet_as_dicts(TACWEED_XLSM_PATH, TACWEED_SHEET_JARD)
    except Exception as e:
        jard_rows = None
        log_write(log, f"❌ فشل قراءة شيت \"{TACWEED_SHEET_JARD}\" من ملف تكويد نون: {str(e)[:200]}\n")
    if jard_rows is None:
        log_write(log, f"⚠️ شيت \"{TACWEED_SHEET_JARD}\" غير موجود جوه ملف تكويد نون\n")

    if not noon_rows and not jard_rows:
        return

    date_uploaded = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # ── من شيت "Noon": SKU NooN + الكود 01 + اجمالي التكلفة  →  Tacweed ──
    tacweed_rows = []
    if noon_rows:
        for r in noon_rows:
            sku = str(r.get("SKU NooN") or "").strip()
            if not sku:
                continue
            code01 = r.get("الكود 01")
            code01 = "" if code01 is None else code01
            total_cost = r.get("اجمالي التكلفة")
            total_cost = "" if total_cost is None else total_cost
            _comp = r.get("الشركة")
            _comp = "" if _comp is None else str(_comp).strip()
            tacweed_rows.append([sku, code01, total_cost, date_uploaded, _comp])
        log_write(log, f"✅ {len(tacweed_rows)} صف جاهزين من شيت \"{TACWEED_SHEET_NOON}\"\n")

    # ── من شيت "الجرد": رمز الباركود + الكمية + اسم المادة  →  WarehouseStock ──
    warehouse_rows = []
    if jard_rows:
        for r in jard_rows:
            code = str(r.get("رمز الباركود") or "").strip()
            if not code:
                continue
            qty = r.get("الكمية")
            qty = "" if qty is None else qty
            item_name = r.get("اسم المادة")
            item_name = "" if item_name is None else item_name
            warehouse_rows.append([code, qty, item_name, date_uploaded])
        log_write(log, f"✅ {len(warehouse_rows)} صف جاهزين من شيت \"{TACWEED_SHEET_JARD}\"\n")
        try:   # تشخيص: أعمدة الشيت + كام صف كميته صفر/فاضية + كام كود متكرر
            _hdrs = list(jard_rows[0].keys()) if jard_rows else []
            log_write(log, f"🔎 أعمدة «{TACWEED_SHEET_JARD}»: {' | '.join(str(h) for h in _hdrs)[:300]}\n")
            _zero = sum(1 for w in warehouse_rows if w[1] in ("", None, 0, "0"))
            _codes = [w[0] for w in warehouse_rows]
            log_write(log, f"🔎 صفوف كميتها صفر/فاضية: {_zero} من {len(warehouse_rows)} | أكواد متكررة: {len(_codes) - len(set(_codes))}\n")
        except Exception:
            pass

    try:
        client = _gsheet_client(creds_path)
        ss = client.open(sheet_name)
    except Exception as e:
        log_write(log, f"❌ فشل الاتصال بـ Google Sheets: {str(e)[:300]}\n")
        return

    # ── تحديث Tacweed بالـ SKU (من غير مسح) ──
    if tacweed_rows:
        try:
            tw_ws = _gsheet_get_or_create_ws(ss, "Tacweed", TACWEED_HEADERS)
            log_write(log, f"⬆️ تحديث Tacweed بالـ SKU ({len(tacweed_rows)} صف في الملف) — من غير مسح...\n")
            upd, add, zl = _tacweed_upsert(tw_ws, tacweed_rows, log)
            log_write(log, f"✅ Tacweed: {upd} صف اتحدّث · {add} SKU اتضاف\n")
            if zl:
                log_write(log, f"⚠️ {len(zl)} منتج تكلفتهم فاضية/خطأ في ملف التكويد وبيتحسبوا بصفر لحد ما المحاسب يكتب التكلفة: {', '.join(zl[:8])}{' ...' if len(zl) > 8 else ''}\n")
        except Exception as e:
            log_write(log, f"❌ فشل تحديث Tacweed (مفيش حاجة اتمسحت): {str(e)[:250]}\n")

    # ── رفع WarehouseStock (مسح القديم + إضافة الجديد) ──
    if warehouse_rows:
        try:
            wh_ws = _gsheet_get_or_create_ws(ss, "WarehouseStock", WAREHOUSE_STOCK_HEADERS)
            log_write(log, "🗑️ مسح بيانات WarehouseStock القديمة...\n")
            _gsheet_clear_data_keep_header(wh_ws)
            log_write(log, f"⬆️ رفع {len(warehouse_rows)} صف جديد لشيت WarehouseStock...\n")
            _gsheet_append_in_batches(wh_ws, warehouse_rows, log)
            log_write(log, "✅ تم تحديث شيت WarehouseStock بالكامل\n")
        except Exception as e:
            log_write(log, f"❌ فشل رفع البيانات لشيت WarehouseStock: {str(e)[:250]}\n")

    log_write(log, "🔗 اكتملت مزامنة ملف تكويد نون\n")


# ================= تشغيل تقارير الجرد والمبيعات لكل المتاجر =================
# ================= تحديث ملف الجرد (إكسيل) مرة واحدة عند بداية البرنامج =================
EXCEL_REFRESH_DIR  = r"\\WIN-181U7G4HFM5\Share For All\تكويد المتاجر"
EXCEL_REFRESH_NAME = "جرد المواد - رئيسي + جملة 005"
EXCEL_REFRESH_WAIT = 120          # ثانية استنى بعد الرفريش قبل الحفظ (حد أدنى)
EXCEL_TACWEED_WAIT = 60           # ثانية استنى بعد تحديث ملف «تكويد نون» قبل الحفظ
EXCEL_EXTRA_WAIT   = 300          # لو الاستعلام لسه شغّال بعد الدقيقتين: أقصى انتظار إضافي (ثانية)
EXCEL_HARD_TIMEOUT = 600          # لو إكسيل علّق خالص: بنقفله بالقوة بعد كده (ثانية) وبوت نون يكمّل
_EXCEL_DONE = threading.Event()   # بيتفعّل لما التحديث يخلص (أو يفشل) — التشغيلة العادية بتستناه


def _find_refresh_file():
    try:
        for fn in os.listdir(EXCEL_REFRESH_DIR):
            base, ext = os.path.splitext(fn)
            if base.strip().lower() == EXCEL_REFRESH_NAME.lower() and ext.lower() in (".xlsx", ".xlsm", ".xlsb", ".xls") \
                    and not fn.startswith("~$"):
                return os.path.join(EXCEL_REFRESH_DIR, fn)
    except Exception:
        pass
    return ""


def _refresh_one_excel(log, path, wait, update_links=0):
    """بيفتح ملف «جرد المواد - رئيسي + جملة 005» في إكسيل فعلي، يعمل Refresh All، يستنى دقيقتين، يحفظ ويقفل.
    بيتنفّذ في أول كل تشغيلة لبوت نون. أي فشل بيتسجل في اللوج والبرنامج يكمّل عادي.
    إكسيل بيتقفل في كل الأحوال: Close ثم Quit، ولو فضل شغّال بنقفله بالـ PID (نسخة إكسيل بتاعتنا بس)."""
    app = None
    wb = None
    pid = None
    watchdog = None
    try:
        log_write(log, f"\n📊 فتح في إكسيل: {os.path.basename(path)}\n")
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx("Excel.Application")     # نسخة إكسيل مستقلة (ما تلمسش أي إكسيل مفتوح عندك)
        app.Visible = True
        app.DisplayAlerts = False
        try:
            import win32process
            pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
        except Exception:
            pid = None
        wb = app.Workbooks.Open(path, UpdateLinks=update_links)
        if wb.ReadOnly:
            log_write(log, "⚠️ الملف مفتوح عند حد تاني (قراءة فقط) — مش هقدر أحفظ، اتخطّيت\n")
            return
        # حارس: لو إكسيل علّق خالص (نافذة مخفية/استعلام واقف) بنقفله بالقوة بعد EXCEL_HARD_TIMEOUT عشان بوت نون ما يفضلش واقف
        if pid:
            def _killer(p=pid):
                try:
                    import subprocess
                    subprocess.run(["taskkill", "/PID", str(p), "/F"], capture_output=True,
                                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                    log_write(log, "⛔ إكسيل علّق أكتر من اللازم — اتقفل بالقوة\n")
                except Exception:
                    pass
            watchdog = threading.Timer(EXCEL_HARD_TIMEOUT, _killer)
            watchdog.daemon = True
            watchdog.start()
        log_write(log, "🔄 Refresh All...\n")
        wb.RefreshAll()          # الرفريش بيشتغل في الخلفية — مفيش أي نداء بيقفل إكسيل وإحنا مستنيين
        log_write(log, f"⏳ استنى {int(wait)} ثانية...\n")
        time.sleep(wait)

        def _refreshing():
            try:
                for c in wb.Connections:
                    for attr in ("OLEDBConnection", "ODBCConnection"):
                        try:
                            if getattr(c, attr).Refreshing:
                                return True
                        except Exception:
                            pass
            except Exception:
                pass
            return False
        t0 = time.time()
        while _refreshing() and time.time() - t0 < EXCEL_EXTRA_WAIT:
            time.sleep(3)
        if _refreshing():
            log_write(log, "⚠️ الاستعلام لسه شغّال بعد الانتظار الإضافي — هحفظ اللي وصل\n")
        wb.Save()
        log_write(log, "💾 اتحفظ ✅\n")
    except Exception as e:
        log_write(log, f"❌ تعذّر تحديث ملف الجرد: {str(e)[:250]}\n")
    finally:
        try:
            if watchdog is not None:
                watchdog.cancel()
        except Exception:
            pass
        # 1) قفل الملف
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception as e:
            log_write(log, f"⚠️ تعذّر قفل الملف: {str(e)[:150]}\n")
        # 2) قفل إكسيل
        try:
            if app is not None:
                app.DisplayAlerts = False
                app.Quit()
        except Exception as e:
            log_write(log, f"⚠️ تعذّر قفل إكسيل: {str(e)[:150]}\n")
        wb = None
        app = None
        try:
            import gc
            gc.collect()
        except Exception:
            pass
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass
        # 3) لو العملية لسه حيّة بعد ثواني: نقفلها بالـ PID (نسخة إكسيل بتاعتنا بس)
        if pid:
            try:
                time.sleep(4)
                import subprocess
                chk = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/NH"], capture_output=True, text=True,
                                     creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).stdout
                if "EXCEL" in chk.upper():
                    subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True,
                                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                    log_write(log, "🔒 إكسيل ما قفلش لوحده — قفلته بالقوة\n")
                else:
                    log_write(log, "🔒 إكسيل اتقفل\n")
            except Exception:
                pass
        _EXCEL_DONE.set()


def refresh_inventory_excel(log):
    """أول كل تشغيلة: (1) ملف «جرد المواد - رئيسي + جملة 005» (2) ملف «تكويد نون» — بنفس الطريقة (فتح → Refresh → حفظ → قفل).
    التكويد بنحدّث روابطه الخارجية (UpdateLinks=3) عشان شيت «الجرد» اللي البوت بيقراه يبقى بآخر أرقام ومتحفوظة في الملف."""
    p1 = _find_refresh_file()
    if p1:
        _refresh_one_excel(log, p1, EXCEL_REFRESH_WAIT, 0)
    else:
        log_write(log, f"⚠️ ملف الجرد مش موجود: {EXCEL_REFRESH_DIR}\\{EXCEL_REFRESH_NAME}.* — اتخطّيت\n")
    if os.path.exists(TACWEED_XLSM_PATH):
        _refresh_one_excel(log, TACWEED_XLSM_PATH, EXCEL_TACWEED_WAIT, 3)
    else:
        log_write(log, f"⚠️ ملف التكويد مش موجود: {TACWEED_XLSM_PATH} — اتخطّيت\n")


def do_run_reports(email, stores, otp_wait, log):
    """
    غلاف حماية خارجي — أي خطأ غير متوقع يتسجل في اللوج بدل ما الـ Thread يموت بصمت.
    """
    try:
        _run_reports_inner(email, stores, otp_wait, log)
    except Exception as e:
        log_write(log, f"\n💥 خطأ غير متوقع أوقف تشغيلة التقارير بالكامل: {str(e)[:400]}\n")


def _run_reports_inner(email, stores, otp_wait, log):
    """
    stores = [{"name": "...", "url": "...", "store_code": "..."}, ...] — نفس قايمة
    "المتاجر الأولى". لكل متجر: يدخل بنفس جلسة Chrome، يحمّل ملف الجرد، وبعدين
    ملف تصدير المبيعات. في الآخر بيجمع كل ملفات الجرد في إكسيل واحد، وكل ملفات
    المبيعات في إكسيل تاني، ويحفظهم جنب الكود.
    """
    if not email or not stores:
        log_write(log, "⚠️ لازم يكون في إيميل ومتجر واحد على الأقل\n")
        return

    _cfg = load_settings()
    skip_inv, skip_sales, skip_cat = (bool(_cfg.get("skip_inventory")), bool(_cfg.get("skip_sales")), bool(_cfg.get("skip_catalog")))
    _sk = [n for n, f in (("الجرد", skip_inv), ("المبيعات", skip_sales), ("الكتالوج", skip_cat)) if f]
    if _sk:
        log_write(log, f"⏭️ متخطّى في التشغيلة دي (الـ API بيجيبهم): {' ، '.join(_sk)}\n")

    # أول ما التشغيلة تبدأ: تحديث ملف الجرد في إكسيل (Refresh → استنى → حفظ → قفل) وبعدين نكمّل عادي
    try:
        refresh_inventory_excel(log)
    except Exception as e:
        log_write(log, f"❌ خطأ في تحديث ملف الجرد: {str(e)[:200]}\n")

    # ── قبل ما نبدأ ندخل على المتاجر خالص: نرفع بيانات ملف تكويد نون
    # (Tacweed / WarehouseStock) الأول ──
    try:
        sync_tacweed_to_google_sheets(load_settings(), log)
    except Exception as e:
        log_write(log, f"❌ خطأ غير متوقع في مزامنة ملف تكويد نون: {str(e)[:250]}\n")

    driver, driver_profile_dir = start_driver(log)

    inventory_files = []
    sales_files     = []
    catalog_files   = []   # [(store_name, file_path), ...] لشيت LIVE
    ads_files       = []   # [(store_name, period, start, end, path), ...] لشيتي AdsDaily / AdsMonthly

    for idx, store in enumerate(stores, 1):
        store_name = store.get("name", f"متجر {idx}")
        store_code = store.get("store_code", "")
        store_url  = store.get("url", "")

        if not store_code:
            log_write(log, f"⚠️ {store_name}: مفيش كود متجر (store_code) محفوظ — تخطي\n")
            continue

        log_write(log, f"\n{'='*40}\n🏪 [{idx}] {store_name} (PRJ{store_code})\n{'='*40}\n")

        login_url = store_url or f"https://fbn.noon.partners/ar-sa/inventory?mp=noon&project=PRJ{store_code}&page=1"

        # ── أي فشل في الدخول (حتى لو الفحص السطحي بيقول إن الجلسة "حية") —
        # بيعتبر مؤشر إن حاجة اتعطلت في المتصفح بعد المتجر اللي قبله، فبنقفل
        # Chrome بالكامل ونفتحه من جديد كليًا قبل ما نحاول تاني. ده الفرق
        # الأساسي عن قبل: مكناش بنعيد الفتح غير لو _is_driver_alive بترجع False،
        # وده كان بيرجع True غلط في حالات كتير فتفشل باقي المتاجر كلها ورا بعض. ──
        login_ok = False
        max_attempts = 2
        for attempt in range(1, max_attempts + 1):
            driver, driver_profile_dir = ensure_driver_alive(driver, driver_profile_dir, log)
            login_ok = login_to_noon(driver, email, login_url, log, otp_wait)
            if login_ok:
                break
            if attempt < max_attempts:
                log_write(log, f"🔄 محاولة {attempt}/{max_attempts} فشلت لـ {store_name} — جاري إغلاق Chrome وفتحه من جديد كليًا والمحاولة تاني (هيتطلب OTP جديد)...\n")
                close_driver(driver, driver_profile_dir)
                driver, driver_profile_dir = start_driver(log)

        if not login_ok:
            log_write(log, f"⚠️ {store_name}: تم تخطيه بسبب فشل تسجيل الدخول (بعد {max_attempts} محاولات)\n")
            continue

        if not skip_inv:
            try:
                inv_path = download_inventory_report(driver, store_code, log)
                if inv_path:
                    inventory_files.append(inv_path)
            except Exception as e:
                log_write(log, f"❌ {store_name}: خطأ في تنزيل الجرد — {str(e)[:200]}\n")

        if not skip_sales:
            try:
                sales_path = download_sales_export_report(driver, store_code, log)
                if sales_path:
                    sales_files.append(sales_path)
            except Exception as e:
                log_write(log, f"❌ {store_name}: خطأ في تنزيل المبيعات — {str(e)[:200]}\n")

        if not skip_cat:
            try:
                catalog_path = download_catalog_export_report(driver, store_code, log)
                if catalog_path:
                    catalog_files.append((store_name, catalog_path))
            except Exception as e:
                log_write(log, f"❌ {store_name}: خطأ في تصدير الكتالوج — {str(e)[:200]}\n")

        try:
            for period, a_s, a_e, a_path in download_ads_exports(driver, store, log):
                ads_files.append((store_name, period, a_s, a_e, a_path))
        except Exception as e:
            log_write(log, f"❌ {store_name}: خطأ في تصدير الإعلانات — {str(e)[:200]}\n")

    if driver:
        close_driver(driver, driver_profile_dir)

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    inv_out_path = sales_out_path = None
    if not skip_inv:
        inv_out_path = combine_csvs_to_excel(
            inventory_files, os.path.join(SCRIPT_DIR, f"جرد_مجمع_{ts}.xlsx"), "Inventory", log
        )
    if not skip_sales:
        sales_out_path = combine_csvs_to_excel(
            sales_files, os.path.join(SCRIPT_DIR, f"مبيعات_مجمعة_{ts}.xlsx"), "Sales", log
        )

    settings_now = load_settings()
    if not (skip_inv and skip_sales):
        try:
            sync_reports_to_google_sheets(inv_out_path, sales_out_path, settings_now, log)
        except Exception as e:
            log_write(log, f"❌ خطأ غير متوقع في مزامنة Google Sheets: {str(e)[:250]}\n")

    if not skip_cat:
        try:
            sync_catalog_to_google_sheets(catalog_files, settings_now, log)
        except Exception as e:
            log_write(log, f"تم\n")

    try:
        sync_ads_history_to_google_sheets(ads_files, log)
    except Exception as e:
        log_write(log, f"❌ خطأ في تاريخ الإعلانات: {str(e)[:250]}\n")
    try:
        sync_ads_to_google_sheets(ads_files, settings_now, log)
    except Exception as e:
        log_write(log, f"❌ خطأ غير متوقع في رفع الإعلانات: {str(e)[:250]}\n")

    log_write(log, "\n🎉 اكتملت!\n")


# ================= Log helper (thread-safe) =================
def log_write(log, text):
    log.after(0, lambda: (log.insert(tk.END, text), log.see(tk.END), log.update()))


# ================= Scheduler =================
scheduler_running = False
scheduler_thread  = None

def scheduler_loop(email_getter, stores_getter, otp_getter, log, schedules_getter):
    global scheduler_running
    notified = set()          # مواعيد اتشغلت في الدقيقة الحالية بس
    last_run_date = {}        # آخر تاريخ (يوم) كل موعد اشتغل فيه فعليًا -> يمنع التكرار في نفس اليوم ويسمح بالتشغيل تاني في اليوم اللي بعده
    while scheduler_running:
        now_dt   = datetime.datetime.now()
        now      = now_dt.strftime("%H:%M")
        today    = now_dt.strftime("%Y-%m-%d")
        schedules = schedules_getter()

        # لو الدقيقة اتغيرت، امسح كل اللي كان متسجل كـ"اشتغل في الدقيقة دي"
        # (قبل كده الكود كان بيمسح فقط قيمة "now" الجديدة، مش القديمة، فالوقت
        # القديم زي "08:50" كان فاضل في notified للأبد ومبيتمسحش، فكان بيمنع
        # التشغيل تاني في اليوم اللي بعده على نفس الميعاد)
        if notified and now not in notified:
            notified.clear()

        if now in schedules and now not in notified and last_run_date.get(now) != today:
            notified.add(now)
            last_run_date[now] = today
            log_write(log, f"\n🕐 جدولة تلقائية الساعة {now}\n")
            threading.Thread(
                target=do_run_reports,
                args=(email_getter(), stores_getter(), otp_getter(), log),
                daemon=True
            ).start()
        time.sleep(15)

def start_scheduler(email_getter, stores_getter, otp_getter, log, schedules_getter):
    global scheduler_running, scheduler_thread
    scheduler_running = True
    scheduler_thread = threading.Thread(
        target=scheduler_loop,
        args=(email_getter, stores_getter, otp_getter, log, schedules_getter),
        daemon=True
    )
    scheduler_thread.start()

def stop_scheduler():
    global scheduler_running
    scheduler_running = False


# ================= UI =================
def build_ui(root):
    """واجهة بوت نون — بتتبني جوه إطار (root) بيدّيه البرنامج الموحّد."""
    if not LOCK_OK:
        tk.Label(root, text="🔒 بوت نون متعطّل: مجلد C:\\Lock noon مش موجود", bg="#0f1923", fg="#e74c3c",
                 font=("Segoe UI", 14, "bold")).pack(expand=True)
        return None
    global ACCENT, BG, BLUE, CARD, ENTRY_BG, GREEN, MUTED, RED, TEXT, add_frame, add_schedule, add_store, bottom, btn_row, card1, delete_schedule, delete_store, email_entry, gmail_pass_entry, gsheet_card, header, log_box, log_frame, mk_btn, move_store_down, move_store_up, nb, otp_entry, refresh_sched_lb, refresh_stores_lb, refresh_stores_lb_detail, run_now, save_email_settings, sb, sched_btn_row, sched_card, sched_lb, sched_lb_frame, sched_status, sched_toggle_btn, settings, store_code_entry, store_name_entry, store_url_entry, stores_card, stores_lb, stores_list_frame, style, style_entry, style_label, tab_gsheet, tab_main, tab_schedule, tab_stores, time_add_frame, time_entry, toggle_scheduler
    settings = load_settings()

    root.configure(bg="#0f1923")

    BG       = "#0f1923"
    CARD     = "#1a2535"
    ACCENT   = "#f5a623"
    TEXT     = "#e8edf3"
    MUTED    = "#7a8fa6"
    GREEN    = "#2ecc71"
    RED      = "#e74c3c"
    BLUE     = "#3498db"
    ENTRY_BG = "#243044"

    def style_label(parent, text, size=10, color=MUTED):
        return tk.Label(parent, text=text, bg=CARD, fg=color,
                        font=("Segoe UI", size), anchor="w")

    def style_entry(parent, width=40):
        e = tk.Entry(parent, width=width, bg=ENTRY_BG, fg=TEXT,
                     insertbackground=ACCENT, relief="flat",
                     font=("Segoe UI", 10))
        e.configure(highlightthickness=1, highlightbackground=MUTED, highlightcolor=ACCENT)
        return e

    def mk_btn(parent, text, cmd, color=ACCENT, fg="#0f1923", size=10, padx=12, pady=6):
        return tk.Button(parent, text=text, command=cmd, bg=color, fg=fg,
                         font=("Segoe UI", size, "bold"), relief="flat",
                         padx=padx, pady=pady, cursor="hand2",
                         activebackground=color)

    # ── Header ──
    header = tk.Frame(root, bg=ACCENT, height=50)
    header.pack(fill="x")
    tk.Label(header, text="Tool",
             bg=ACCENT, fg="#0f1923", font=("Segoe UI", 14, "bold")).pack(side="left", padx=15, pady=10)

    # ── Notebook tabs ──
    nb = ttk.Notebook(root, style="Noon.TNotebook")
    nb.pack(fill="both", expand=True, padx=10, pady=8)

    style = ttk.Style()
    style.theme_use("default")
    style.configure("Noon.TNotebook", background=BG, borderwidth=0)
    style.configure("Noon.TNotebook.Tab", background=CARD, foreground=TEXT,
                    padding=[12, 6], font=("Segoe UI", 10))
    style.map("Noon.TNotebook.Tab", background=[("selected", ACCENT)], foreground=[("selected", "#0f1923")])

    tab_main     = tk.Frame(nb, bg=BG)
    tab_stores   = tk.Frame(nb, bg=BG)
    tab_schedule = tk.Frame(nb, bg=BG)
    tab_gsheet   = tk.Frame(nb, bg=BG)
    nb.add(tab_main,     text="⚙️  الإعدادات")
    nb.add(tab_stores,   text="🏪  المتاجر الأولى")
    nb.add(tab_schedule, text="🕐  الجدولة")
    nb.add(tab_gsheet,   text="")

    # ════════════════════════════════════════
    # TAB 1 — الإعدادات
    # ════════════════════════════════════════
    card1 = tk.Frame(tab_main, bg=CARD, padx=20, pady=15)
    card1.pack(fill="x", padx=10, pady=10)

    style_label(card1, "📧  الإيميل / Email").grid(row=0, column=0, sticky="w", pady=(5,2))
    email_entry = style_entry(card1, width=55)
    email_entry.insert(0, settings.get("email", ""))
    email_entry.grid(row=1, column=0, pady=(0,10), sticky="w")

    style_label(card1, "⏱  وقت انتظار OTP (ثانية)").grid(row=2, column=0, sticky="w", pady=(5,2))
    otp_entry = style_entry(card1, width=10)
    otp_entry.insert(0, str(settings.get("otp_wait", 30)))
    otp_entry.grid(row=3, column=0, pady=(0,10), sticky="w")

    style_label(card1, "").grid(row=4, column=0, sticky="w", pady=(5,2))
    gmail_pass_entry = style_entry(card1, width=30)
    gmail_pass_entry.configure(show="•")
    gmail_pass_entry.insert(0, settings.get("gmail_app_password", ""))
    gmail_pass_entry.grid(row=5, column=0, pady=(0,10), sticky="w")

    tk.Label(card1, text="",
             bg=CARD, fg=MUTED, font=("Segoe UI", 8), wraplength=700, justify="left"
    ).grid(row=6, column=0, sticky="w", pady=(0,5))

    tk.Label(card1, text="",
             bg=CARD, fg=MUTED, font=("Segoe UI", 8), wraplength=700, justify="left"
    ).grid(row=7, column=0, sticky="w", pady=(5,5))

    # ── الحاجات اللي الـ API بيجيبها: البوت يتخطّاها ──
    _skip_fr = tk.Frame(card1, bg=CARD)
    _skip_fr.grid(row=8, column=0, sticky="w", pady=(6, 0))
    tk.Label(_skip_fr, text="⏭️ تخطّي في البوت (الـ API بيجيبهم من تاب Noon API):", bg=CARD, fg=MUTED,
             font=("Segoe UI", 9, "bold")).pack(anchor="w")
    _skip_vars = {}
    for _k, _t in (("skip_inventory", "الجرد (FBN)"), ("skip_sales", "المبيعات"), ("skip_catalog", "الكتالوج (LIVE)")):
        _skip_vars[_k] = tk.BooleanVar(value=bool(settings.get(_k, True)))
        tk.Checkbutton(_skip_fr, text=_t, variable=_skip_vars[_k], bg=CARD, fg=TEXT, selectcolor=ENTRY_BG,
                       activebackground=CARD, activeforeground=TEXT, font=("Segoe UI", 10),
                       command=lambda k=_k: (settings.__setitem__(k, _skip_vars[k].get()), save_settings(settings))
                       ).pack(side="left", padx=(0, 14))

    def save_email_settings():
        settings["email"]              = email_entry.get().strip()
        settings["otp_wait"]           = int(otp_entry.get().strip() or 30)
        settings["gmail_app_password"] = gmail_pass_entry.get()
        for _k, _v in _skip_vars.items():
            settings[_k] = _v.get()
        save_settings(settings)
        messagebox.showinfo("✅", "تم حفظ الإعدادات")

    mk_btn(card1, "💾  حفظ الإعدادات", save_email_settings, color=BLUE, fg="white"
    ).grid(row=9, column=0, sticky="w", pady=(8,0))

    # ════════════════════════════════════════
    # TAB 2 — المتاجر
    # ════════════════════════════════════════
    stores_card = tk.Frame(tab_stores, bg=CARD, padx=15, pady=12)
    stores_card.pack(fill="both", expand=True, padx=10, pady=10)

    tk.Label(stores_card, text="",
             bg=CARD, fg=ACCENT, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0,8))

    stores_list_frame = tk.Frame(stores_card, bg=CARD)
    stores_list_frame.pack(fill="both", expand=True)

    # Listbox للمتاجر
    stores_lb = tk.Listbox(stores_list_frame, bg=ENTRY_BG, fg=TEXT,
                            font=("Segoe UI", 10), selectbackground=ACCENT,
                            selectforeground="#0f1923", relief="flat",
                            height=8, activestyle="none")
    stores_lb.pack(side="left", fill="both", expand=True)

    sb = tk.Scrollbar(stores_list_frame, orient="vertical", command=stores_lb.yview)
    sb.pack(side="right", fill="y")
    stores_lb.config(yscrollcommand=sb.set)

    # تعبئة القائمة من الإعدادات
    def refresh_stores_lb():
        stores_lb.delete(0, tk.END)
        for s in settings["stores"]:
            code = s.get("store_code", "")
            code_txt = f" | PRJ{code}" if code else ""
            stores_lb.insert(tk.END, f"  {s['name']}{code_txt}  —  {s['url'][:45]}")

    refresh_stores_lb()

    # إضافة متجر
    add_frame = tk.Frame(stores_card, bg=CARD)
    add_frame.pack(fill="x", pady=(10,0))

    tk.Label(add_frame, text="الاسم:", bg=CARD, fg=MUTED, font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w")
    store_name_entry = style_entry(add_frame, width=14)
    store_name_entry.grid(row=0, column=1, padx=(3,8))

    tk.Label(add_frame, text="الرابط:", bg=CARD, fg=MUTED, font=("Segoe UI", 9)).grid(row=0, column=2, sticky="w")
    store_url_entry = style_entry(add_frame, width=30)
    store_url_entry.grid(row=0, column=3, padx=(3,8))

    tk.Label(add_frame, text="كود المتجر:", bg=CARD, fg=MUTED, font=("Segoe UI", 9)).grid(row=0, column=4, sticky="w")
    store_code_entry = style_entry(add_frame, width=10)
    store_code_entry.grid(row=0, column=5, padx=(3,0))

    def refresh_stores_lb_detail():
        stores_lb.delete(0, tk.END)
        for s in settings["stores"]:
            code = s.get("store_code", "")
            code_txt = f" | PRJ{code}" if code else ""
            stores_lb.insert(tk.END, f"  {s['name']}{code_txt}  —  {s['url'][:45]}")

    def add_store():
        name = store_name_entry.get().strip()
        url  = store_url_entry.get().strip()
        code = store_code_entry.get().strip()
        if not name or not url:
            messagebox.showwarning("تنبيه", "ادخل اسم ورابط")
            return
        settings["stores"].append({"name": name, "url": url, "store_code": code})
        save_settings(settings)
        refresh_stores_lb_detail()
        store_name_entry.delete(0, tk.END)
        store_url_entry.delete(0, tk.END)
        store_code_entry.delete(0, tk.END)

    def delete_store():
        sel = stores_lb.curselection()
        if not sel:
            return
        if messagebox.askyesno("تأكيد", "حذف المتجر المحدد؟"):
            del settings["stores"][sel[0]]
            save_settings(settings)
            refresh_stores_lb()

    def move_store_up():
        sel = stores_lb.curselection()
        if not sel or sel[0] == 0:
            return
        i = sel[0]
        settings["stores"][i-1], settings["stores"][i] = settings["stores"][i], settings["stores"][i-1]
        save_settings(settings)
        refresh_stores_lb()
        stores_lb.selection_set(i-1)

    def move_store_down():
        sel = stores_lb.curselection()
        if not sel or sel[0] >= len(settings["stores"]) - 1:
            return
        i = sel[0]
        settings["stores"][i+1], settings["stores"][i] = settings["stores"][i], settings["stores"][i+1]
        save_settings(settings)
        refresh_stores_lb()
        stores_lb.selection_set(i+1)

    btn_row = tk.Frame(stores_card, bg=CARD)
    btn_row.pack(fill="x", pady=(6,0))
    mk_btn(btn_row, "➕ إضافة",   add_store,     color=GREEN,  fg="#0f1923", padx=10, pady=5).pack(side="left", padx=2)
    mk_btn(btn_row, "🗑 حذف",     delete_store,  color=RED,    fg="white",   padx=10, pady=5).pack(side="left", padx=2)
    mk_btn(btn_row, "⬆ لأعلى",   move_store_up, color=BLUE,   fg="white",   padx=10, pady=5).pack(side="left", padx=2)
    mk_btn(btn_row, "⬇ لأسفل",   move_store_down, color=BLUE, fg="white",   padx=10, pady=5).pack(side="left", padx=2)

    # ════════════════════════════════════════
    # TAB 3 — الجدولة
    # ════════════════════════════════════════
    sched_card = tk.Frame(tab_schedule, bg=CARD, padx=15, pady=12)
    sched_card.pack(fill="both", expand=True, padx=10, pady=10)

    tk.Label(sched_card, text="🕐 مواعيد التشغيل التلقائي (HH:MM — 24 ساعة)",
             bg=CARD, fg=ACCENT, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0,8))

    sched_lb_frame = tk.Frame(sched_card, bg=CARD)
    sched_lb_frame.pack(fill="x")

    sched_lb = tk.Listbox(sched_lb_frame, bg=ENTRY_BG, fg=TEXT,
                           font=("Segoe UI", 11), selectbackground=ACCENT,
                           selectforeground="#0f1923", relief="flat",
                           height=6, activestyle="none")
    sched_lb.pack(side="left", fill="x", expand=True)

    def refresh_sched_lb():
        sched_lb.delete(0, tk.END)
        for t in sorted(settings["schedules"]):
            sched_lb.insert(tk.END, f"  ⏰  {t}")

    refresh_sched_lb()

    time_add_frame = tk.Frame(sched_card, bg=CARD)
    time_add_frame.pack(fill="x", pady=(10,0))

    tk.Label(time_add_frame, text="الوقت (HH:MM):", bg=CARD, fg=MUTED,
             font=("Segoe UI", 9)).pack(side="left")
    time_entry = style_entry(time_add_frame, width=8)
    time_entry.pack(side="left", padx=(4,10))

    def add_schedule():
        t = time_entry.get().strip()
        try:
            datetime.datetime.strptime(t, "%H:%M")
        except ValueError:
            messagebox.showwarning("")
            return
        if t not in settings["schedules"]:
            settings["schedules"].append(t)
            save_settings(settings)
            refresh_sched_lb()
        time_entry.delete(0, tk.END)

    def delete_schedule():
        sel = sched_lb.curselection()
        if not sel:
            return
        t = sorted(settings["schedules"])[sel[0]]
        settings["schedules"].remove(t)
        save_settings(settings)
        refresh_sched_lb()

    sched_btn_row = tk.Frame(sched_card, bg=CARD)
    sched_btn_row.pack(fill="x", pady=(6,0))
    mk_btn(sched_btn_row, "➕ إضافة موعد",  add_schedule,    color=GREEN, fg="#0f1923", padx=10, pady=5).pack(side="left", padx=2)
    mk_btn(sched_btn_row, "🗑 حذف",         delete_schedule, color=RED,   fg="white",   padx=10, pady=5).pack(side="left", padx=2)

    # Scheduler status
    sched_status = tk.Label(sched_card, text="⏸ الجدولة متوقفة", bg=CARD, fg=MUTED,
                             font=("Segoe UI", 10))
    sched_status.pack(anchor="w", pady=(12,4))

    def toggle_scheduler():
        global scheduler_running
        if not scheduler_running:
            if not settings["schedules"]:
                messagebox.showwarning("")
                return
            start_scheduler(
                lambda: email_entry.get().strip(),
                lambda: settings["stores"],
                lambda: int(otp_entry.get().strip() or 30),
                log_box,
                lambda: settings["schedules"]
            )
            sched_status.config(text="▶ الجدولة شغالة", fg=GREEN)
            sched_toggle_btn.config(text="⏹ إيقاف الجدولة", bg=RED)
        else:
            stop_scheduler()
            sched_status.config(text="⏸ الجدولة متوقفة", fg=MUTED)
            sched_toggle_btn.config(text="▶ تشغيل الجدولة", bg=GREEN)

    sched_toggle_btn = mk_btn(sched_card, "▶ تشغيل الجدولة", toggle_scheduler, color=GREEN, fg="#0f1923", padx=15, pady=7)
    sched_toggle_btn.pack(anchor="w")

    tk.Label(sched_card, text="",
             bg=CARD, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", pady=(8,0))

    # ════════════════════════════════════════
    # TAB 4 — Google Sheets
    # ════════════════════════════════════════
    gsheet_card = tk.Frame(tab_gsheet, bg=CARD, padx=20, pady=15)
    gsheet_card.pack(fill="x", padx=10, pady=10)

    tk.Label(gsheet_card, text="", bg=CARD, fg=ACCENT,
             font=("Segoe UI", 11, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0,10))

    tk.Label(gsheet_card,
             text=""
                  "",
             bg=CARD, fg=MUTED, font=("Segoe UI", 9), wraplength=750, justify="left"
    ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0,12))

    # مسار ملف اعتماد Google (JSON) واسم الـ Google Sheet ثابتين في الكود
    # (GSHEET_CREDS_DIR / GSHEET_NAME) ومش بيظهروا للمستخدم في الواجهة.


    # ════════════════════════════════════════
    # Bottom — Buttons + Log
    # ════════════════════════════════════════
    bottom = tk.Frame(root, bg=BG)
    bottom.pack(fill="x", padx=10, pady=(0,6))

    def run_now():
        email  = email_entry.get().strip()
        stores = settings.get("stores", [])
        if not email:
            messagebox.showwarning("")
            return
        if not stores:
            messagebox.showwarning("تنبيه", "")
            return
        otp = int(otp_entry.get().strip() or 30)
        log_box.delete("1.0", tk.END)
        threading.Thread(target=do_run_reports, args=(email, stores, otp, log_box), daemon=True).start()

    mk_btn(bottom, "", run_now, color=ACCENT, fg="#0f1923", size=12, padx=20, pady=8).pack(side="left", padx=(0,8))

    # Log
    log_frame = tk.Frame(root, bg=BG, padx=10)
    log_frame.pack(fill="both", expand=True, pady=(0,10))
    tk.Label(log_frame, text="📋 السجل", bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w")
    log_box = scrolledtext.ScrolledText(log_frame, width=95, height=10,
                                         bg=ENTRY_BG, fg=TEXT, font=("Consolas", 9),
                                         relief="flat", insertbackground=ACCENT)
    log_box.pack(fill="both", expand=True)

    # تشغيل الجدولة تلقائيًا عند فتح البرنامج (لو فيه مواعيد محفوظة) — عشان
    # لو البرنامج اتحط في مجلد بدء التشغيل (startup) وفتح لوحده بعد ريستارت
    # الجهاز، الجدولة تشتغل من غير ما حد يدوس زرار "▶ تشغيل الجدولة" يدويًا.
    if settings.get("schedules"):
        root.after(1000, toggle_scheduler)




    return stop_scheduler
'''

_EMBEDDED['Amazon_Server'] = r'''# -*- coding: utf-8 -*-
"""
Amazon Server — بوت تنزيل تقرير شحنات العملاء (SHIPMENT_SALES) من Seller Central السعودية
ورفعه لشيت DailyOrdersAmazon في Google Sheets.

نفس فكرة برنامج Noon: واجهة Tkinter + Selenium + جدولة + سجل.
الخطوات: دخول (إيميل -> باسورد -> OTP يدوي) -> اختيار الحساب -> صفحة التقرير -> «آخر 30 يومًا»
-> «طلب تنزيل صيغة .csv» -> انتظار ظهور «تنزيل» -> تنزيل الملف -> رفعه للشيت.
"""
import os
import re
import sys
import csv
import json
import time
import glob
import openpyxl
import zipfile
import datetime
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from gspread.utils import rowcol_to_a1
from oauth2client.service_account import ServiceAccountCredentials

# ===== Paths =====
SCRIPT_DIR    = os.environ.get("PLATFORMS_BASE_DIR") or os.path.dirname(os.path.abspath(sys.argv[0] if getattr(sys, "frozen", False) else __file__))
SETTINGS_PATH = os.path.join(os.path.expanduser("~"), "amazon_settings.json")
# بروفايل كروم ثابت: الجلسة وخيار "تذكّر هذا المتصفح" بيتحفظوا فيه، فالـ OTP بيتطلب نادرًا بعد أول مرة
CHROME_PROFILE_DIR = os.path.join(os.path.expanduser("~"), "AmazonChromeProfile")
# مجلد تنزيل مخصص للبوت (عشان نعرف الملف الجديد بسهولة)
AMZ_DOWNLOADS_DIR  = os.path.join(SCRIPT_DIR, "amazon_downloads")

# نفس المسارات الثابتة بتاعة برنامج Noon
GSHEET_CREDS_DIR = r"\\WIN-181U7G4HFM5\Share For Markting\NOONIMG\Copy"
GSHEET_NAME      = "Complaints"
GSHEET_SCOPE     = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# ===== روابط أمازون =====
WELCOME_URL = "https://sellercentral.amazon.sa/reportcentral/WelcomePage"
REPORT_URL  = "https://sellercentral.amazon.sa/reportcentral/SHIPMENT_SALES/1"
STOCK_URL   = "https://sellercentral.amazon.sa/reportcentral/AFNInventoryReport/2"

# ===== شيت DailyOrdersAmazon (نفس الأعمدة اللي التطبيق بيقراها بالاسم) =====
ORDERS_SHEET = "DailyOrdersAmazon"
ORDERS_HEADERS = ["تاريخ شحنة العميل", "رقم تخزين سلعة التاجر MSKU",
                  "رقم تخزين السلعة لشبكة نظام الشحن (FNSKU)", "ASIN",
                  "حاوية كاملة الحمولة", "الكمية", "رقم الطلب من أمازون",
                  "العملة", "مبلغ المنتج"]
# كلمات بنلاقي بيها عمود التقرير لو اسمه مختلف شوية عن اسم الشيت (بالترتيب نفسه)
ORDERS_KEYWORDS = [
    ["تاريخ شحنة العميل", "تاريخ الشحن", "shipment date", "ship date"],
    ["msku", "تخزين سلعة التاجر", "merchant sku", "seller sku"],
    ["fnsku", "شبكة نظام الشحن"],
    ["asin"],
    ["حاوية كاملة الحمولة", "fulfillment center", "fulfilment center", "fulfillment", "=fc"],
    ["الكمية", "quantity"],
    ["رقم الطلب", "order id", "amazon order"],
    ["العملة", "currency"],
    ["مبلغ المنتج", "product amount", "item price", "product sales"],
]

# ===== شيت StockAmazon (مخزون أمازون FBA) — بنسيب بس SELLABLE وكمية أكبر من صفر =====
STOCK_SHEET = "StockAmazon"
STOCK_HEADERS = ["seller-sku", "fulfillment-channel-sku", "asin", "condition-type",
                 "Warehouse-Condition-code", "Quantity Available"]

# ===== ملف تكويد أمازون (ASIN + الكود 01 + اجمالي التكلفة) → شيت TacweedAmazon =====
TACWEED_DIR        = r"\\WIN-181U7G4HFM5\Share For All\تكويد المتاجر"
TACWEED_AMZ_FILE   = "امازون 2023-06-02.xlsm"          # لو اسمه اتغيّر بيدوّر على أحدث ملف فيه كلمة «امازون»
TACWEED_AMZ_SHEET  = "Amazon"
TACWEED_AMZ_SHEET_NAME = "TacweedAmazon"
TACWEED_AMZ_HEADERS = ["ASIN", "Code01", "Cost", "Date Uploaded"]

DAYS_CHOICES = {"اليوم (حتى الآن)": "0", "أمس": "1", "آخر 3 أيام": "3", "آخر 7 أيام": "7",
                "آخر 14 يومًا": "14", "آخر 30 يومًا": "30"}


# ================= Settings =================
def load_settings():
    default = {
        "email": "",
        "password": "",
        "account_name": "",        # اختياري: اسم الحساب لو عندك أكتر من حساب في صفحة الاختيار
        "email2": "",              # الحساب 2 (متجر المخزون والتقارير التالية)
        "password2": "",
        "account_name2": "",
        "otp_wait": 120,           # أقصى انتظار لكتابة كود OTP (ثانية)
        "days": "30",
        "upload": True,            # رفع للجوجل شيت بعد التنزيل
        "clear_before_upload": False,  # شحنات: false = إضافة تحت آخر صف بدون تكرار (الافتراضي)، true = مسح وكتابة
        "sync_tacweed": True,      # أول خطوة: تحديث TacweedAmazon من ملف التكويد (قبل فتح أمازون)
        "run_orders": True,        # تشغيل تقرير شحنات العملاء
        "run_orders2": True,       # تقرير شحنات الحساب 2 (بيتضاف دايمًا على DailyOrdersAmazon بدون مسح)
        "run_stock": True,         # تشغيل تقرير مخزون أمازون (StockAmazon — دايمًا مسح وكتابة)
        "schedules": [],
    }
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, encoding="utf-8") as f:
                data = json.load(f)
            for k, v in default.items():
                data.setdefault(k, v)
            return data
        except Exception:
            pass
    return default


def save_settings(settings):
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


# ================= Log =================
LOG_WIDGET = None
ROOT = None


def log_write(log, text):
    try:
        if log is None:
            print(text, end="")
            return
        log.after(0, lambda: (log.insert(tk.END, text), log.see(tk.END)))
    except Exception:
        print(text, end="")


# ================= Driver =================
_CACHED_CHROMEDRIVER_PATH = None


def _clear_stale_wdm_lock():
    try:
        wdm_dir = os.path.join(os.path.expanduser("~"), ".wdm")
        if os.path.isdir(wdm_dir):
            for fname in os.listdir(wdm_dir):
                if fname.startswith(".wdm-lock"):
                    try:
                        os.remove(os.path.join(wdm_dir, fname))
                    except Exception:
                        pass
    except Exception:
        pass


def _get_chromedriver_path(log=None):
    global _CACHED_CHROMEDRIVER_PATH
    if _CACHED_CHROMEDRIVER_PATH and os.path.exists(_CACHED_CHROMEDRIVER_PATH):
        return _CACHED_CHROMEDRIVER_PATH
    _clear_stale_wdm_lock()
    try:
        _CACHED_CHROMEDRIVER_PATH = ChromeDriverManager().install()
    except Exception as e:
        log_write(log, f"⚠️ فشل تجهيز chromedriver: {str(e)[:200]}\n")
        _clear_stale_wdm_lock()
        _CACHED_CHROMEDRIVER_PATH = ChromeDriverManager().install()
    return _CACHED_CHROMEDRIVER_PATH


_NO_PASSKEY_JS = """
try {
  if (navigator.credentials) {
    navigator.credentials.get = () => Promise.reject(new DOMException('Blocked', 'NotAllowedError'));
    navigator.credentials.create = () => Promise.reject(new DOMException('Blocked', 'NotAllowedError'));
  }
  if (window.PublicKeyCredential) {
    PublicKeyCredential.isUserVerifyingPlatformAuthenticatorAvailable = () => Promise.resolve(false);
    PublicKeyCredential.isConditionalMediationAvailable = () => Promise.resolve(false);
  }
} catch (e) {}
"""


def start_driver(log=None, account=1):
    profile_dir = CHROME_PROFILE_DIR + ("" if account == 1 else str(account))   # بروفايل منفصل لكل حساب
    os.makedirs(profile_dir, exist_ok=True)
    os.makedirs(AMZ_DOWNLOADS_DIR, exist_ok=True)
    opts = webdriver.ChromeOptions()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--remote-debugging-port=0")
    opts.add_argument("--disable-background-timer-throttling")
    opts.add_argument("--disable-backgrounding-occluded-windows")
    opts.add_argument("--disable-renderer-backgrounding")
    opts.add_argument(f"--user-data-dir={profile_dir}")
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    opts.add_experimental_option("prefs", {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        "download.default_directory": AMZ_DOWNLOADS_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    })
    opts.set_capability("unhandledPromptBehavior", "accept")
    driver = webdriver.Chrome(service=Service(_get_chromedriver_path(log)), options=opts)
    # منع نافذة الـ Passkey ("No passkeys available") اللي أمازون بتفتحها على بعض الأجهزة —
    # بنخلّي الطلب يترفض فأمازون ترجع لصفحة الباسورد العادية
    try:
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": _NO_PASSKEY_JS})
    except Exception:
        pass
    try:
        driver.set_page_load_timeout(60)
        driver.set_script_timeout(30)
    except Exception:
        pass
    return driver


def close_driver(driver):
    try:
        driver.quit()
    except Exception:
        pass


def _is_driver_alive(driver):
    try:
        _ = driver.current_url
        return True
    except Exception:
        return False


def _debug_dump(driver, log, tag):
    """لو حصل خطأ: بنحفظ لقطة شاشة + HTML جنب البرنامج عشان نظبط الـ selectors لو احتاج."""
    try:
        ts = datetime.datetime.now().strftime("%H%M%S")
        base = os.path.join(SCRIPT_DIR, f"amazon_debug_{tag}_{ts}")
        driver.save_screenshot(base + ".png")
        with open(base + ".html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        log_write(log, f"🧾 اتحفظت لقطة الشاشة والـ HTML: {base}.png / .html\n")
    except Exception:
        pass


# ================= JS helpers (بتدخل جوه الـ Shadow DOM بتاع عناصر Amazon) =================
JS_LIB = r"""
const norm = s => (s || '').replace(/[\u200e\u200f\u202a-\u202e\u00a0]/g, ' ').replace(/\s+/g, ' ').trim();
function deepAll(root, out) {
  out = out || [];
  root.querySelectorAll('*').forEach(e => { out.push(e); if (e.shadowRoot) deepAll(e.shadowRoot, out); });
  return out;
}
const vis = e => !!(e.getClientRects && e.getClientRects().length);
const clickable = e => (e.closest && e.closest('button,a,kat-button,kat-link,[role=button]')) || e;
// كل العناصر المرئية اللي نصها = أحد النصوص (exact) أو بيحتويها (partial)، الأقصر نصًا الأول
function findText(texts, exact) {
  const res = [];
  deepAll(document).forEach(e => {
    if (['SCRIPT', 'STYLE', 'TEMPLATE'].includes(e.tagName) || !vis(e)) return;
    const t = norm(e.textContent);
    if (!t) return;
    const ok = exact ? texts.includes(t) : texts.some(x => t.includes(x));
    if (ok) res.push([t.length, e]);
  });
  res.sort((a, b) => a[0] - b[0]);
  return res.map(x => x[1]);
}
"""


def _js(driver, body, *args):
    return driver.execute_script(JS_LIB + "\n" + body, *args)


def _displayed(driver, css):
    try:
        for el in driver.find_elements(By.CSS_SELECTOR, css):
            if el.is_displayed():
                return el
    except Exception:
        pass
    return None


def _type_into(el, text):
    el.click()
    try:
        el.clear()
    except Exception:
        pass
    el.send_keys(Keys.CONTROL, "a")
    el.send_keys(Keys.DELETE)
    el.send_keys(text)


# ================= تسجيل الدخول =================
SELECT_ACCOUNT_TEXTS = ["Select account", "Select Account", "تحديد الحساب", "اختيار الحساب", "حدد الحساب"]


def ask_otp_dialog(driver, timeout):
    """بيظهر نافذة (فوق كل الشبابيك) تطلب كود OTP، وبيرجّع الكود. لو كتبت الكود بنفسك في المتصفح النافذة بتتقفل لوحدها."""
    if ROOT is None:
        return ""
    out, ev, holder = {"v": ""}, threading.Event(), {}

    def _build():
        w = tk.Toplevel(ROOT)
        holder["w"] = w
        w.title("كود التحقق | OTP")
        w.configure(bg="#1a2535")
        w.attributes("-topmost", True)
        w.geometry("430x190")
        tk.Label(w, text="اكتب كود OTP اللي وصلك على الموبايل\n(أو اكتبه بنفسك في المتصفح وهتتقفل النافذة لوحدها)",
                 bg="#1a2535", fg="#e8edf3", font=("Segoe UI", 10), justify="center").pack(pady=(14, 8))
        e = tk.Entry(w, font=("Segoe UI", 16), justify="center", width=12)
        e.pack(pady=4)

        def ok(_=None):
            out["v"] = re.sub(r"\D", "", e.get())
            ev.set()
            w.destroy()

        tk.Button(w, text="تأكيد | OK", command=ok, bg="#ff9900", font=("Segoe UI", 10, "bold"), relief="flat", padx=16, pady=4).pack(pady=8)
        e.bind("<Return>", ok)
        w.protocol("WM_DELETE_WINDOW", lambda: (ev.set(), w.destroy()))
        w.lift()
        w.focus_force()
        e.focus_set()
        try:
            w.bell()
        except Exception:
            pass

    ROOT.after(0, _build)
    t0 = time.time()
    while time.time() - t0 < timeout and not ev.is_set():
        if _displayed(driver, "#auth-mfa-otpcode") is None:      # الصفحة عدّت من شاشة OTP
            break
        time.sleep(1)

    def _close():
        w = holder.get("w")
        try:
            if w is not None and w.winfo_exists():
                w.destroy()
        except Exception:
            pass

    ROOT.after(0, _close)
    return out["v"]


def _click_select_account(driver, account_name, log):
    """بيضغط زرار Select account (أو اللي جوه صف الحساب المطلوب لو اتحدد اسم)."""
    clicked = _js(driver, """
        const texts = arguments[0], acc = norm(arguments[1]);
        const cands = findText(texts, true);
        if (!cands.length) return 0;
        let pick = cands[0];
        if (acc) {
          for (const c of cands) {
            let p = c, n = 0;
            while (p && n < 6) {
              if (norm(p.textContent).includes(acc)) { pick = c; n = 99; break; }
              p = p.parentNode || p.host; n++;
            }
            if (n === 99) break;
          }
        }
        clickable(pick).click();
        return cands.length;
    """, SELECT_ACCOUNT_TEXTS, account_name or "")
    if clicked:
        log_write(log, "🔘 ضغطت Select account\n")
    return bool(clicked)


def _auth_step(driver, ctx, log):
    """بينفذ خطوة واحدة حسب الشاشة الحالية. بيرجّع: email / password / otp / select / captcha / none."""
    s = ctx["settings"]

    el = _displayed(driver, "#ap_email")
    if el is not None:
        if ctx["email_n"] >= 3:
            raise RuntimeError("الإيميل اتكتب 3 مرات والصفحة لسه بتطلبه — راجع الإيميل في الإعدادات")
        ctx["email_n"] += 1
        log_write(log, "✉️ كتابة الإيميل...\n")
        _type_into(el, s["email"])
        btn = _displayed(driver, "#continue")
        (btn.click() if btn else el.send_keys(Keys.ENTER))
        time.sleep(2.5)
        return "email"

    el = _displayed(driver, "#ap_password")
    if el is not None:
        if ctx["pass_n"] >= 3:
            raise RuntimeError("الباسورد اتكتب 3 مرات والصفحة لسه بتطلبه — راجع الباسورد في الإعدادات (وقف عشان الحساب ما يتقفلش)")
        ctx["pass_n"] += 1
        log_write(log, "🔑 كتابة الباسورد...\n")
        _type_into(el, s["password"])
        rm = _displayed(driver, "input[name=rememberMe]")
        try:
            if rm is not None and not rm.is_selected():
                rm.click()
        except Exception:
            pass
        btn = _displayed(driver, "#signInSubmit")
        (btn.click() if btn else el.send_keys(Keys.ENTER))
        time.sleep(2.5)
        return "password"

    el = _displayed(driver, "#auth-mfa-otpcode")
    if el is not None:
        if ctx["otp_n"] >= 3:
            raise RuntimeError("كود OTP اتكتب 3 مرات ومتقبلش")
        ctx["otp_n"] += 1
        rem = _displayed(driver, "#auth-mfa-remember-device")
        try:
            if rem is not None and not rem.is_selected():
                rem.click()      # «لا تطلب OTP على هذا المتصفح» — عشان الأكواد تقل بعد كده
        except Exception:
            pass
        log_write(log, "📱 مطلوب كود OTP — اكتبه في النافذة اللي ظهرت...\n")
        code = ask_otp_dialog(driver, s.get("otp_wait", 120))
        if code:
            el = _displayed(driver, "#auth-mfa-otpcode") or el
            _type_into(el, code)
            btn = _displayed(driver, "#auth-signin-button")
            (btn.click() if btn else el.send_keys(Keys.ENTER))
            time.sleep(3)
        else:
            log_write(log, "⏳ مفيش كود في النافذة — مستني تكتبه بنفسك في المتصفح...\n")
            t0 = time.time()
            while time.time() - t0 < s.get("otp_wait", 120) and _displayed(driver, "#auth-mfa-otpcode"):
                time.sleep(2)
        return "otp"

    if _displayed(driver, "#auth-captcha-guess") or _displayed(driver, "#captchacharacters"):
        log_write(log, "🧩 ظهر كابتشا — حلها بنفسك في المتصفح، مستنيك...\n")
        time.sleep(5)
        return "captcha"

    if _click_select_account(driver, s.get("account_name", ""), log):
        time.sleep(3)
        return "select"
    return "none"


def _is_signed_in_url(url):
    return ("sellercentral.amazon.sa" in url and "/ap/" not in url and "signin" not in url.lower())


def login_to_amazon(driver, settings, log):
    """بيدخل Seller Central ويعدّي شاشات (الإيميل/الباسورد/OTP/اختيار الحساب). True لو نجح."""
    if not settings.get("email") or not settings.get("password"):
        log_write(log, "❌ اكتب الإيميل والباسورد في الإعدادات الأول\n")
        return False
    ctx = {"settings": settings, "email_n": 0, "pass_n": 0, "otp_n": 0}
    log_write(log, "🌐 فتح Seller Central...\n")
    driver.get(WELCOME_URL)
    time.sleep(4)
    t0 = time.time()
    quiet = 0
    while time.time() - t0 < 300:
        try:
            state = _auth_step(driver, ctx, log)
        except RuntimeError:
            raise
        except Exception as e:           # عنصر اتغيّر/الصفحة بتتحمّل — نعيد المحاولة
            log_write(log, f"…الصفحة بتتحدّث ({type(e).__name__}) — إعادة محاولة\n")
            time.sleep(2)
            continue
        if state == "none":
            if _is_signed_in_url(driver.current_url):
                quiet += 1
                if quiet >= 2:        # مفيش شاشة دخول مرتين ورا بعض
                    log_write(log, "✅ تم تسجيل الدخول\n")
                    return True
            time.sleep(2)
        else:
            quiet = 0
    log_write(log, "❌ انتهى وقت تسجيل الدخول\n")
    _debug_dump(driver, log, "login")
    return False


# ================= تقرير شحنات العملاء =================
def _wait_report_ready(driver, ctx, log, timeout=150):
    """بيستنى ظهور قايمة «تاريخ الحدث». لو ظهرت شاشة دخول/اختيار حساب بيعدّيها."""
    t0 = time.time()
    while time.time() - t0 < timeout:
        ok = _js(driver, "const d = document.querySelector('kat-dropdown'); return !!(d && d.shadowRoot && vis(d));")
        if ok:
            return True
        try:
            state = _auth_step(driver, ctx, log)
        except RuntimeError:
            raise
        except Exception:
            time.sleep(2)
            continue
        if state == "none":
            # ممكن يكون تاب «عرض على الإنترنت» هو المفتوح — نفتح تاب «تنزيل»
            _js(driver, """
                const t = findText(['تنزيل', 'Download'], true).find(e => e.closest('[role=tab],kat-tab,kat-tabs,.tab'));
                if (t) clickable(t).click();
            """)
            time.sleep(2)
    return False


def _set_days(driver, days, log):
    """بيختار «آخر N يومًا» من قايمة تاريخ الحدث (kat-dropdown) ويتأكد."""
    days = str(days)
    for attempt in range(3):
        _js(driver, """
            const dd = document.querySelector('#daily-time-picker-style kat-dropdown') || document.querySelector('kat-dropdown');
            const h = dd.shadowRoot.querySelector('.select-header'); if (h) h.click();
        """)
        time.sleep(0.8)
        _js(driver, """
            const dd = document.querySelector('#daily-time-picker-style kat-dropdown') || document.querySelector('kat-dropdown');
            const o = dd.querySelector('kat-option[value="' + arguments[0] + '"]');
            if (o) o.click();
        """, days)
        time.sleep(1)
        cur = _js(driver, """
            const dd = document.querySelector('#daily-time-picker-style kat-dropdown') || document.querySelector('kat-dropdown');
            return String(dd.value || dd.getAttribute('value') || '');
        """)
        if cur == days:
            log_write(log, f"✅ المدة: {days} يوم\n")
            return True
        # بديل: ضبط القيمة مباشرة + إطلاق حدث التغيير
        _js(driver, """
            const dd = document.querySelector('#daily-time-picker-style kat-dropdown') || document.querySelector('kat-dropdown');
            dd.value = arguments[0];
            dd.dispatchEvent(new CustomEvent('change', {detail: {value: arguments[0]}, bubbles: true, composed: true}));
        """, days)
        time.sleep(1)
    cur = _js(driver, "const dd = document.querySelector('kat-dropdown'); return String(dd.value || '');")
    if cur == days:
        return True
    log_write(log, f"❌ مقدرتش أختار المدة ({days}) — القيمة الحالية: {cur}\n")
    return False


_ROWS_JS = r"""
// صفوف جدول التقارير المطلوبة: [{t: نص الصف, d: فيه «تنزيل» ولا لا}]
function reportRows() {
  const rows = deepAll(document).filter(e => e.tagName === 'KAT-TABLE-ROW' && !e.closest('kat-table-head') && vis(e));
  return rows.filter(r => !/نوع التقرير|Report Type/i.test(norm(r.textContent)));
}
function rowDownload(r) {
  return deepAll(r).concat([r]).find(e => ['تنزيل', 'Download'].includes(norm(e.textContent)) && vis(e));
}
"""


def _snapshot_rows(driver):
    return _js(driver, _ROWS_JS + """
        return reportRows().map(r => ({t: norm(r.textContent), d: !!rowDownload(r)}));
    """) or []


def _click_row_download(driver, row_text):
    return _js(driver, _ROWS_JS + """
        const r = reportRows().find(x => norm(x.textContent) === arguments[0]);
        if (!r) return false;
        const b = rowDownload(r); if (!b) return false;
        clickable(b).click(); return true;
    """, row_text)


def _wait_new_file(before, timeout=180):
    """بيستنى ظهور ملف جديد مكتمل في مجلد التنزيل ويرجّع مساره."""
    t0 = time.time()
    while time.time() - t0 < timeout:
        now = set(glob.glob(os.path.join(AMZ_DOWNLOADS_DIR, "*")))
        new = [p for p in now - before if not p.lower().endswith((".crdownload", ".tmp"))]
        if new:
            p = max(new, key=os.path.getmtime)
            s1 = os.path.getsize(p)
            time.sleep(1.5)
            if os.path.exists(p) and os.path.getsize(p) == s1 and s1 > 0:
                return p
        time.sleep(1)
    return None


def _unzip_if_needed(path, log):
    if path.lower().endswith(".zip"):
        try:
            with zipfile.ZipFile(path) as z:
                names = [n for n in z.namelist() if n.lower().endswith((".csv", ".txt"))]
                if names:
                    z.extract(names[0], AMZ_DOWNLOADS_DIR)
                    return os.path.join(AMZ_DOWNLOADS_DIR, names[0])
        except Exception as e:
            log_write(log, f"⚠️ فك الضغط فشل: {str(e)[:150]}\n")
    return path


def download_shipment_report_once(driver, settings, ctx, log):
    log_write(log, "\n📦 تقرير شحنات العملاء (SHIPMENT_SALES)...\n")
    driver.get(REPORT_URL)
    time.sleep(4)
    if not _wait_report_ready(driver, ctx, log):
        log_write(log, "❌ صفحة التقرير ما ظهرتش\n")
        _debug_dump(driver, log, "report")
        return None
    time.sleep(1.5)
    if not _set_days(driver, settings.get("days", "30"), log):
        _debug_dump(driver, log, "days")
        return None

    before = _snapshot_rows(driver)
    before_texts = [r["t"] for r in before]
    ok = _js(driver, """
        // 1) الزرار بقيمته value=CSV (بيشتغل بأي لغة للصفحة)  2) أو بنصه عربي/إنجليزي
        let b = deepAll(document).find(e => ['BUTTON', 'KAT-BUTTON'].includes(e.tagName) && vis(e) &&
                 String(e.getAttribute('value') || '').toUpperCase() === 'CSV');
        if (!b) {
          const t = findText(['طلب تنزيل صيغة .csv', 'Request .csv Download', 'Request .CSV Download'], false)[0];
          if (t) b = clickable(t);
        }
        if (!b) return false;
        b.click(); return true;
    """)
    if not ok:
        log_write(log, "❌ زرار «طلب تنزيل صيغة .csv / Request .csv Download» مش لاقيه\n")
        _debug_dump(driver, log, "csvbtn")
        return None
    log_write(log, "⏳ تم طلب التقرير — مستني يجهز...\n")

    target = None
    t0 = time.time()
    attempt = 0
    while time.time() - t0 < 420:          # 7 دقايق
        attempt += 1
        time.sleep(5)
        if attempt % 6 == 0:                # كل ~30 ثانية: تحديث الصفحة عشان الجدول يتحدّث
            try:
                driver.refresh()
                time.sleep(4)
                _wait_report_ready(driver, ctx, log, timeout=60)
            except Exception:
                pass
        rows = _snapshot_rows(driver)
        pool = list(before_texts)
        new_ready = None
        for r in rows:
            if r["t"] in pool:
                pool.remove(r["t"])
                continue
            if r["d"]:
                new_ready = r
                break
        if new_ready:
            target = new_ready
            break
        # لو الأمازون ما أضافش صف جديد (نفس الطلب موجود بالفعل) ومفيش جديد بعد دقيقتين — نستخدم أعلى صف جاهز
        if time.time() - t0 > 120 and rows and rows[0]["d"] and len(rows) == len(before):
            log_write(log, "ℹ️ مفيش صف جديد — هنزّل أحدث تقرير موجود\n")
            target = rows[0]
            break
    if not target:
        log_write(log, "❌ التقرير ما جهزش في الوقت المحدد\n")
        _debug_dump(driver, log, "wait")
        return None

    log_write(log, f"✅ التقرير جاهز: {target['t'][:90]}\n")
    files_before = set(glob.glob(os.path.join(AMZ_DOWNLOADS_DIR, "*")))
    if not _click_row_download(driver, target["t"]):
        log_write(log, "❌ زرار «تنزيل» مش لاقيه\n")
        _debug_dump(driver, log, "dl")
        return None
    path = _wait_new_file(files_before)
    if not path:
        log_write(log, "❌ الملف ما نزلش\n")
        _debug_dump(driver, log, "nofile")
        return None
    path = _unzip_if_needed(path, log)
    log_write(log, f"✅ تم تنزيل: {os.path.basename(path)}\n")
    return path


def download_shipment_report(driver, settings, ctx, log, tries=3):
    for i in range(1, tries + 1):
        try:
            p = download_shipment_report_once(driver, settings, ctx, log)
        except RuntimeError as e:
            log_write(log, f"❌ {e}\n")
            return None
        except Exception as e:
            log_write(log, f"❌ خطأ في التنزيل: {str(e)[:200]}\n")
            p = None
        if p:
            return p
        if i < tries:
            log_write(log, f"🔁 محاولة {i + 1}/{tries}...\n")
            time.sleep(3)
    return None


def _row_end_date(text):
    """تاريخ نهاية نطاق التقرير من نص الصف (d/m/yy - d/m/yy) — بيرجّع datetime.date أو None."""
    ds = re.findall(r"(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{2,4})", text)
    if not ds:
        return None
    d, m, y = ds[1] if len(ds) >= 2 else ds[0]
    y = int(y) + (2000 if int(y) < 100 else 0)
    try:
        return datetime.date(y, int(m), int(d))
    except ValueError:
        return None


def download_stock_report_once(driver, settings, ctx, log):
    log_write(log, "\n📊 تقرير مخزون أمازون (AFNInventoryReport)...\n")
    driver.get(STOCK_URL)
    time.sleep(4)
    rows = []
    t0 = time.time()
    while time.time() - t0 < 150:
        rows = [r for r in _snapshot_rows(driver) if _row_end_date(r["t"])]
        if rows:
            break
        try:
            state = _auth_step(driver, ctx, log)
        except RuntimeError:
            raise
        except Exception:
            state = "err"
        time.sleep(2 if state in ("none", "err") else 0.5)
    if not rows:
        log_write(log, "❌ جدول التقارير ما ظهرش\n")
        _debug_dump(driver, log, "stock_table")
        return None
    ready = [r for r in rows if r["d"]]
    if not ready:
        log_write(log, "❌ مفيش تقرير جاهز للتنزيل في الجدول\n")
        _debug_dump(driver, log, "stock_none")
        return None
    best = max(ready, key=lambda r: _row_end_date(r["t"]))      # أحدث تاريخ نهاية (والأول في الجدول لو تساوا)
    log_write(log, f"✅ أحدث تقرير جاهز: {best['t'][:90]}\n")
    files_before = set(glob.glob(os.path.join(AMZ_DOWNLOADS_DIR, "*")))
    if not _click_row_download(driver, best["t"]):
        log_write(log, "❌ زرار «تنزيل» مش لاقيه\n")
        _debug_dump(driver, log, "stock_dl")
        return None
    path = _wait_new_file(files_before)
    if not path:
        log_write(log, "❌ الملف ما نزلش\n")
        _debug_dump(driver, log, "stock_nofile")
        return None
    path = _unzip_if_needed(path, log)
    log_write(log, f"✅ تم تنزيل: {os.path.basename(path)}\n")
    return path


def download_stock_report(driver, settings, ctx, log, tries=3):
    for i in range(1, tries + 1):
        try:
            p = download_stock_report_once(driver, settings, ctx, log)
        except RuntimeError as e:
            log_write(log, f"❌ {e}\n")
            return None
        except Exception as e:
            log_write(log, f"❌ خطأ في تنزيل المخزون: {str(e)[:200]}\n")
            p = None
        if p:
            return p
        if i < tries:
            log_write(log, f"🔁 محاولة {i + 1}/{tries}...\n")
            time.sleep(3)
    return None


# ================= قراءة الملف وتجهيزه للشيت =================
def _read_text(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-16", "cp1256", "latin-1"):
        try:
            return raw.decode(enc)
        except Exception:
            continue
    return raw.decode("utf-8", errors="replace")


def read_report_rows(path, header_kw=None):
    """بيرجّع (headers, rows) من CSV/TXT — بيكتشف الفاصل وصف العناوين تلقائي."""
    text = _read_text(path)
    lines = [l for l in text.splitlines() if l.strip()]
    if not lines:
        return [], []
    sample = "\n".join(lines[:10])
    delim = max([",", "\t", ";"], key=lambda d: sample.count(d))
    table = list(csv.reader(lines, delimiter=delim))
    flat_kw = header_kw or [re.sub(r"[-_\s]+", " ", k.lower()) for kws in ORDERS_KEYWORDS for k in kws if not k.startswith("=")]
    hdr_i = 0
    for i, row in enumerate(table[:15]):
        hits = sum(1 for c in row if any(k in re.sub(r"[-_\s]+", " ", str(c).lower()) for k in flat_kw))
        if hits >= 3:
            hdr_i = i
            break
    headers = [str(c).strip() for c in table[hdr_i]]
    rows = [r for r in table[hdr_i + 1:] if any(str(c).strip() for c in r)]
    return headers, rows


def map_orders_columns(src_headers):
    """لكل عمود في ORDERS_HEADERS بيرجّع رقم العمود في الملف (أو None)."""
    nz = lambda x: re.sub(r"[-_\s]+", " ", str(x).lower()).strip()
    low = [nz(h) for h in src_headers]
    used = set()
    res = []
    for target, kws in zip(ORDERS_HEADERS, ORDERS_KEYWORDS):
        idx = None
        for i, h in enumerate(src_headers):           # 1) تطابق تام مع اسم الشيت
            if i not in used and h.strip() == target:
                idx = i
                break
        if idx is None:                                # 2) كلمة مفتاحية (الأطول أولًا)
            for kw in sorted(kws, key=len, reverse=True):
                for i, h in enumerate(low):
                    if i not in used and ((h == nz(kw[1:])) if kw.startswith('=') else (nz(kw) in h)):
                        idx = i
                        break
                if idx is not None:
                    break
        if idx is not None:
            used.add(idx)
        res.append(idx)
    return res


def build_orders_rows(path, log):
    headers, rows = read_report_rows(path)
    if not headers or not rows:
        log_write(log, "❌ الملف فاضي أو شكله غير مفهوم\n")
        return []
    idx = map_orders_columns(headers)
    missing = [ORDERS_HEADERS[i] for i, v in enumerate(idx) if v is None]
    if missing:
        log_write(log, "⚠️ أعمدة مش لاقيها في الملف (هتتسيب فاضية): " + " | ".join(missing) + "\n")
        log_write(log, "   أعمدة الملف: " + " | ".join(headers) + "\n")
    out = []
    for r in rows:
        out.append([(r[i].strip() if i is not None and i < len(r) else "") for i in idx])
    # لازم يبقى فيه MSKU/ASIN على الأقل عشان الصف يبقى مفيد
    out = [r for r in out if r[1] or r[3]]
    log_write(log, f"✅ {len(out)} صف شحنات جاهزين للرفع\n")
    return out


def build_stock_rows(path, log):
    """مخزون أمازون: بنسيب Warehouse-Condition-code = SELLABLE بس، وQuantity Available أكبر من صفر."""
    nz = lambda x: re.sub(r"[-_\s]+", " ", str(x).lower()).strip()
    headers, rows = read_report_rows(path, header_kw=[nz(h) for h in STOCK_HEADERS])
    if not headers or not rows:
        log_write(log, "❌ ملف المخزون فاضي أو شكله غير مفهوم\n")
        return []
    low = [nz(h) for h in headers]
    idx = []
    for h in STOCK_HEADERS:
        idx.append(low.index(nz(h)) if nz(h) in low else None)
    missing = [STOCK_HEADERS[i] for i, v in enumerate(idx) if v is None]
    if missing:
        log_write(log, "❌ أعمدة ناقصة في ملف المخزون: " + " | ".join(missing) + "\n   أعمدة الملف: " + " | ".join(headers) + "\n")
        return []
    ci, qi = idx[4], idx[5]
    out, skipped = [], 0
    for r in rows:
        get = lambda i: (r[i].strip() if i < len(r) else "")
        if get(ci).upper() != "SELLABLE":
            skipped += 1
            continue
        try:
            q = float(get(qi).replace(",", ""))
        except ValueError:
            skipped += 1
            continue
        if q <= 0:
            skipped += 1
            continue
        out.append([get(i) for i in idx])
    log_write(log, f"✅ مخزون: {len(out)} صف (SELLABLE وكمية > 0) — اتجاهل {skipped} صف\n")
    return out


# ================= Google Sheets =================
def _resolve_gsheet_creds_path():
    try:
        m = glob.glob(os.path.join(GSHEET_CREDS_DIR, "*.json"))
        if m:
            return m[0]
    except Exception:
        pass
    return ""


def _gsheet_append_in_batches(ws, rows, log, batch_size=1000):
    """بيكتب الصفوف تحت آخر صف موجود على دفعات، وأي فشل بيطلع كخطأ واضح (مش بيتبلع). بعد الكتابة بيتأكد من العدد."""
    if not rows:
        return
    start = len(ws.col_values(1)) + 1               # أول صف فاضي (بعد الهيدر / آخر بيانات)
    need = start + len(rows) - 1
    if ws.row_count < need + 5:
        ws.add_rows(need + 5 - ws.row_count)         # لو الشيت وصل حد الخلايا هيطلع الخطأ هنا بوضوح
    for i in range(0, len(rows), batch_size):
        chunk = rows[i:i + batch_size]
        last = None
        for attempt in range(5):
            try:
                ws.update(values=chunk, range_name=f"A{start + i}", value_input_option="USER_ENTERED")
                last = None
                break
            except gspread.exceptions.APIError as e:
                last = e
                if "429" in str(e) or "quota" in str(e).lower():
                    time.sleep(2 * (2 ** attempt))
                else:
                    break                              # خطأ حقيقي (مش ضغط) — مفيش فايدة من التكرار
        if last is not None:
            raise RuntimeError(f"فشلت كتابة الدفعة {i + 1}–{i + len(chunk)}: {str(last)[:300]}")
        log_write(log, f"   ⬆️ {min(i + batch_size, len(rows))}/{len(rows)} صف\n")
    got = len(ws.col_values(1)) - (start - 1)
    if got < len(rows):
        raise RuntimeError(f"بعد الرفع لقيت {got} صف بس من {len(rows)} في الشيت")


def align_rows_to_sheet_header(sheet_header, rows, headers=None):
    """بيرتّب الصفوف على حسب ترتيب هيدر الشيت الحالي (بالاسم)، وأي عمود في الشيت مش بتاعنا بيتسيب فاضي.
    بيرجّع (header, rows_aligned) — والهيدر ممكن يتزوّد بأعمدة ناقصة."""
    headers = headers or ORDERS_HEADERS
    header = list(sheet_header)
    pos = {}
    for t in headers:
        if t in header:
            pos[t] = header.index(t)
        else:
            header.append(t)
            pos[t] = len(header) - 1
    out = []
    for r in rows:
        row = [""] * len(header)
        for ti, t in enumerate(headers):
            row[pos[t]] = r[ti]
        out.append(row)
    return header, out


def _norm_cell(v):
    t = str(v).strip()
    try:
        return f"{float(t.replace(',', '')):.2f}"
    except ValueError:
        return t.lower()


def sync_rows_to_google_sheets(sheet_name, headers, rows, clear, log, dedupe_cols=None, batch_size=1000):
    """clear=True: مسح القديم وكتابة الجديد. clear=False: إضافة تحت آخر صف، وبيتجاهل الصفوف الموجودة بالفعل (حسب dedupe_cols)."""
    if not rows:
        log_write(log, f"⚠️ مفيش صفوف للرفع — شيت {sheet_name} ما اتلمسش\n")
        return False
    creds_path = _resolve_gsheet_creds_path()
    if not creds_path:
        log_write(log, f"⚠️ الرفع متخطى — مفيش ملف .json (Service Account) في: {GSHEET_CREDS_DIR}\n")
        return False
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, GSHEET_SCOPE)
        ss = gspread.authorize(creds).open(GSHEET_NAME)
    except Exception as e:
        log_write(log, f"❌ فشل الاتصال بـ Google Sheets: {str(e)[:300]}\n")
        return False
    try:
        try:
            ws = ss.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            ws = ss.add_worksheet(title=sheet_name, rows="2000", cols=str(len(headers)))
            ws.update(values=[headers], range_name="A1")
        sheet_header = [h.strip() for h in ws.row_values(1)]
        header, aligned = align_rows_to_sheet_header(sheet_header if any(sheet_header) else headers, rows, headers)
        if header != sheet_header:
            ws.update(values=[header], range_name="A1")
        ncols = len(header)
        pos = {t: header.index(t) for t in headers}
        if clear:
            log_write(log, f"🗑️ مسح بيانات {sheet_name} القديمة...\n")
            ws.batch_clear([f"A2:{rowcol_to_a1(max(ws.row_count, 2), ncols)}"])
        elif dedupe_cols:
            from collections import Counter
            kcols = [pos[headers[c]] for c in dedupe_cols]
            have = Counter(tuple(_norm_cell(r[c] if c < len(r) else "") for c in kcols) for r in ws.get_all_values()[1:])
            before, fresh = len(aligned), []
            for r in aligned:
                k = tuple(_norm_cell(r[c]) for c in kcols)
                if have[k] > 0:
                    have[k] -= 1          # موجود بالفعل
                else:
                    fresh.append(r)
            aligned = fresh
            log_write(log, f"➕ إضافة: {len(aligned)} صف جديد من {before} (الباقي موجود قبل كده)\n")
        if not aligned:
            log_write(log, f"✅ شيت {sheet_name} محدّث بالفعل — مفيش صفوف جديدة\n")
            return True
        log_write(log, f"⬆️ رفع {len(aligned)} صف لشيت {sheet_name}...\n")
        _gsheet_append_in_batches(ws, aligned, log, batch_size)
        log_write(log, f"✅ تم تحديث شيت {sheet_name}\n")
        return True
    except Exception as e:
        log_write(log, f"❌ فشل الرفع لـ Google Sheets: {str(e)[:300]}\n")
        return False


def sync_orders_to_google_sheets(rows, settings, log, force_append=False):
    # المفتاح: رقم الطلب + MSKU + ASIN + الكمية + المبلغ (عشان نمنع التكرار في وضع الإضافة)
    return sync_rows_to_google_sheets(ORDERS_SHEET, ORDERS_HEADERS, rows, (False if force_append else bool(settings.get("clear_before_upload", False))),
                                      log, dedupe_cols=(6, 1, 3, 5, 8))


def sync_stock_to_google_sheets(rows, settings, log):
    return sync_rows_to_google_sheets(STOCK_SHEET, STOCK_HEADERS, rows, True, log)


# ================= تكويد أمازون (أول خطوة، قبل فتح الموقع) =================
def _find_tacweed_amazon_file():
    f = os.path.join(TACWEED_DIR, TACWEED_AMZ_FILE)
    if os.path.exists(f):
        return f
    try:
        c = [x for x in glob.glob(os.path.join(TACWEED_DIR, "*امازون*.xls*")) if not os.path.basename(x).startswith("~$")]
        if c:
            return max(c, key=os.path.getmtime)
    except Exception:
        pass
    return ""


def sync_tacweed_amazon(settings, log):
    """بتقرا شيت «Amazon» من ملف تكويد أمازون: SKU Amazon (= ASIN) + الكود 01 + اجمالي التكلفة، وتحدّث بيهم TacweedAmazon (مسح وكتابة)."""
    path = _find_tacweed_amazon_file()
    if not path:
        log_write(log, f"⚠️ تكويد أمازون متخطى — ملف التكويد مش موجود في: {TACWEED_DIR}\n")
        return False
    log_write(log, f"\n📗 قراءة ملف تكويد أمازون ({os.path.basename(path)})... ممكن ياخد دقيقة\n")
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        if TACWEED_AMZ_SHEET not in wb.sheetnames:
            wb.close()
            log_write(log, f"⚠️ شيت \"{TACWEED_AMZ_SHEET}\" مش موجود جوه الملف\n")
            return False
        it = wb[TACWEED_AMZ_SHEET].iter_rows(values_only=True)
        hdr = [str(h).strip() if h is not None else "" for h in next(it)]
        need = {"SKU Amazon": None, "الكود 01": None, "اجمالي التكلفة": None}
        for k in need:
            if k not in hdr:
                wb.close()
                log_write(log, f"❌ العمود «{k}» مش موجود في شيت Amazon. أعمدة الشيت: {' | '.join(h for h in hdr if h)}\n")
                return False
            need[k] = hdr.index(k)
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows, bad = [], 0
        for r in it:
            g = lambda i: (r[i] if i < len(r) and r[i] is not None else "")
            asin = str(g(need["SKU Amazon"])).strip()
            if not asin:
                continue
            code, cost = g(need["الكود 01"]), g(need["اجمالي التكلفة"])
            if isinstance(cost, str) and cost.strip().startswith("#"):      # أخطاء إكسيل (#VALUE! / #N/A ...) → فاضي
                cost = ""
                bad += 1
            if isinstance(code, str) and code.strip().startswith("#"):
                code = ""
            rows.append([asin, code, cost, stamp])
        wb.close()
    except Exception as e:
        log_write(log, f"❌ فشل قراءة ملف تكويد أمازون: {str(e)[:250]}\n")
        return False
    log_write(log, f"✅ {len(rows)} صف جاهزين من شيت Amazon" + (f" — {bad} صف تكلفته خطأ في الإكسيل (اتسابت فاضية)" if bad else "") + "\n")
    return sync_rows_to_google_sheets(TACWEED_AMZ_SHEET_NAME, TACWEED_AMZ_HEADERS, rows, True, log, batch_size=5000)


# ================= التشغيل =================
_RUN_LOCK = threading.Lock()


_LAST_DRIVER = None


def _account_settings(settings, n):
    """نسخة من الإعدادات فيها إيميل/باسورد/اسم الحساب الخاص برقم الحساب n (1 أو 2)."""
    st = dict(settings)
    if n == 2:
        st["email"], st["password"], st["account_name"] = settings.get("email2", ""), settings.get("password2", ""), settings.get("account_name2", "")
    return st


def _run_job(settings, log, title, download_fn, process_fn, account=1):
    """كل مهمة بتفتح Chrome جديد (بنفس البروفايل المحفوظ = نفس الجلسة)، تدخل، تنزّل، تقفل، وبعدين ترفع للشيت."""
    global _LAST_DRIVER
    driver = None
    ok = False
    try:
        settings = _account_settings(settings, account)
        log_write(log, f"\n{'─' * 36}\n▶ {title}  (الحساب {account}: {settings.get('email') or '—'})\n{'─' * 36}\n")
        if not settings.get("email") or not settings.get("password"):
            log_write(log, f"❌ اكتب إيميل وباسورد الحساب {account} في الإعدادات\n")
            return False
        driver = start_driver(log, account)
        if not login_to_amazon(driver, settings, log):
            return False
        ctx = {"settings": settings, "email_n": 0, "pass_n": 0, "otp_n": 0}
        path = download_fn(driver, settings, ctx, log)
        if not path:
            log_write(log, "⚠️ مفيش ملف اتنزل\n")
            return False
        close_driver(driver)
        driver = None
        if settings.get("upload", True):
            process_fn(path)
        else:
            log_write(log, f"ℹ️ الرفع متوقف في الإعدادات — الملف عندك: {path}\n")
        ok = True
        return True
    except Exception as e:
        import traceback
        log_write(log, f"\n💥 خطأ غير متوقع: {e}\n{traceback.format_exc()[-900:]}\n")
        if driver is not None:
            _debug_dump(driver, log, "fatal")
        return False
    finally:
        if driver is not None:
            if ok:
                close_driver(driver)
            else:
                _LAST_DRIVER = driver
                log_write(log, "ℹ️ سبت Chrome مفتوح عشان تشوف الصفحة — هيتقفل تلقائي في التشغيلة الجاية\n")


def do_run(settings, log):
    """تشغيلة كاملة: شحنات الحساب 1 ثم شحنات الحساب 2 (الاتنين إضافة بدون تكرار على DailyOrdersAmazon) ثم مخزون الحساب 2 (مسح وكتابة على StockAmazon). Chrome جديد لكل تقرير."""
    global _LAST_DRIVER
    if not _RUN_LOCK.acquire(blocking=False):
        log_write(log, "⚠️ فيه تشغيلة شغالة حاليًا — استنى تخلص\n")
        return
    try:
        if _LAST_DRIVER is not None:           # نقفل متصفح التشغيلة الفاشلة السابقة
            close_driver(_LAST_DRIVER)
            _LAST_DRIVER = None
        log_write(log, f"\n{'=' * 40}\n🚀 بدء التشغيل {datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n{'=' * 40}\n")
        if settings.get("sync_tacweed", True):
            try:
                sync_tacweed_amazon(settings, log)
            except Exception as e:
                log_write(log, f"❌ خطأ في تكويد أمازون: {str(e)[:250]}\n")
        if settings.get("run_orders", True):
            _run_job(settings, log, "تقرير شحنات العملاء → DailyOrdersAmazon", download_shipment_report,
                     lambda path: sync_orders_to_google_sheets(build_orders_rows(path, log), settings, log), account=1)
        if settings.get("run_orders2", True) and _LAST_DRIVER is None:
            _run_job(settings, log, "شحنات الحساب 2 → DailyOrdersAmazon (إضافة)", download_shipment_report,
                     lambda path: sync_orders_to_google_sheets(build_orders_rows(path, log), settings, log, force_append=True), account=2)
        if settings.get("run_stock", True) and _LAST_DRIVER is None:
            _run_job(settings, log, "مخزون أمازون → StockAmazon", download_stock_report,
                     lambda path: sync_stock_to_google_sheets(build_stock_rows(path, log), settings, log), account=2)
        log_write(log, "\n🎉 انتهت التشغيلة\n")
    finally:
        _RUN_LOCK.release()


# ================= Scheduler =================
scheduler_running = False


def scheduler_loop(settings_getter, log):
    last_run = {}
    while scheduler_running:
        now_dt = datetime.datetime.now()
        now, today = now_dt.strftime("%H:%M"), now_dt.strftime("%Y-%m-%d")
        st = settings_getter()
        if now in st.get("schedules", []) and last_run.get(now) != today:
            last_run[now] = today
            log_write(log, f"\n🕐 جدولة تلقائية الساعة {now}\n")
            threading.Thread(target=do_run, args=(dict(st), log), daemon=True).start()
        time.sleep(15)


def start_scheduler(settings_getter, log):
    global scheduler_running
    scheduler_running = True
    threading.Thread(target=scheduler_loop, args=(settings_getter, log), daemon=True).start()


def stop_scheduler():
    global scheduler_running
    scheduler_running = False


# ================= UI =================
def build_ui(root):
    """واجهة البوت — بتتبني جوه إطار (root) بيدّيه البرنامج الموحّد."""
    global ROOT
    settings = load_settings()
    ROOT = root
    BG, CARD, ACCENT = "#0f1923", "#1a2535", "#ff9900"
    TEXT, MUTED, GREEN, RED, BLUE, ENTRY_BG = "#e8edf3", "#7a8fa6", "#2ecc71", "#e74c3c", "#3498db", "#243044"
    root.configure(bg=BG)

    def lbl(parent, text):
        return tk.Label(parent, text=text, bg=CARD, fg=MUTED, font=("Segoe UI", 10), anchor="w")

    def entry(parent, width=45, show=None):
        e = tk.Entry(parent, width=width, bg=ENTRY_BG, fg=TEXT, insertbackground=ACCENT, relief="flat",
                     font=("Segoe UI", 10), show=show)
        e.configure(highlightthickness=1, highlightbackground=MUTED, highlightcolor=ACCENT)
        return e

    def btn(parent, text, cmd, color=ACCENT, fg="#0f1923", size=10, padx=12, pady=6):
        return tk.Button(parent, text=text, command=cmd, bg=color, fg=fg, font=("Segoe UI", size, "bold"),
                         relief="flat", padx=padx, pady=pady, cursor="hand2", activebackground=color)

    header = tk.Frame(root, bg=ACCENT, height=50)
    header.pack(fill="x")
    tk.Label(header, text="Amazon Tool", bg=ACCENT, fg="#0f1923", font=("Segoe UI", 14, "bold")).pack(side="left", padx=15, pady=10)

    style = ttk.Style()
    style.theme_use("default")
    style.configure("Amz.TNotebook", background=BG, borderwidth=0)
    style.configure("Amz.TNotebook.Tab", background=CARD, foreground=TEXT, padding=[12, 6], font=("Segoe UI", 10))
    style.map("Amz.TNotebook.Tab", background=[("selected", ACCENT)], foreground=[("selected", "#0f1923")])
    nb = ttk.Notebook(root, style="Amz.TNotebook")
    nb.pack(fill="both", expand=True, padx=10, pady=8)
    tab_main = tk.Frame(nb, bg=BG)
    tab_sched = tk.Frame(nb, bg=BG)
    nb.add(tab_main, text="⚙️  الإعدادات")
    nb.add(tab_sched, text="🕐  الجدولة")

    # ── الإعدادات ──
    c = tk.Frame(tab_main, bg=CARD, padx=20, pady=15)
    c.pack(fill="x", padx=10, pady=10)
    def acct_block(col, title, k_email, k_pass, k_name):
        lbl(c, title).grid(row=0, column=col, sticky="w", padx=(0, 14))
        e1 = entry(c, width=32)
        e1.insert(0, settings.get(k_email, ""))
        e1.grid(row=1, column=col, sticky="w", pady=(0, 8), padx=(0, 14))
        lbl(c, "🔑  الباسورد").grid(row=2, column=col, sticky="w")
        e2 = entry(c, width=32, show="•")
        e2.insert(0, settings.get(k_pass, ""))
        e2.grid(row=3, column=col, sticky="w", pady=(0, 8), padx=(0, 14))
        lbl(c, "🏢  اسم الحساب في صفحة الاختيار (اختياري)").grid(row=4, column=col, sticky="w")
        e3 = entry(c, width=32)
        e3.insert(0, settings.get(k_name, ""))
        e3.grid(row=5, column=col, sticky="w", pady=(0, 8), padx=(0, 14))
        return e1, e2, e3

    email_e, pass_e, acc_e = acct_block(0, "📧  الحساب 1 — شحنات العملاء", "email", "password", "account_name")
    email2_e, pass2_e, acc2_e = acct_block(1, "📧  الحساب 2 — مخزون أمازون", "email2", "password2", "account_name2")
    lbl(c, "📅  مدة التقرير").grid(row=6, column=0, sticky="w")
    days_cb = ttk.Combobox(c, values=list(DAYS_CHOICES.keys()), state="readonly", width=25)
    inv = {v: k for k, v in DAYS_CHOICES.items()}
    days_cb.set(inv.get(str(settings.get("days", "30")), "آخر 30 يومًا"))
    days_cb.grid(row=7, column=0, sticky="w", pady=(0, 8))
    lbl(c, "⏱  أقصى انتظار لكتابة OTP (ثانية)").grid(row=8, column=0, sticky="w")
    otp_e = entry(c, width=10)
    otp_e.insert(0, str(settings.get("otp_wait", 120)))
    otp_e.grid(row=9, column=0, sticky="w", pady=(0, 8))
    upload_v = tk.BooleanVar(value=bool(settings.get("upload", True)))
    clear_v = tk.BooleanVar(value=bool(settings.get("clear_before_upload", False)))
    tw_v = tk.BooleanVar(value=bool(settings.get("sync_tacweed", True)))
    orders_v = tk.BooleanVar(value=bool(settings.get("run_orders", True)))
    orders2_v = tk.BooleanVar(value=bool(settings.get("run_orders2", True)))
    stock_v = tk.BooleanVar(value=bool(settings.get("run_stock", True)))
    for r, (txt, var) in enumerate([("رفع الملفات لـ Google Sheets بعد التنزيل", upload_v),
                                    ("أول خطوة: تحديث TacweedAmazon من ملف تكويد أمازون (مسح وكتابة)", tw_v),
                                    ("حساب 1: شحنات العملاء → DailyOrdersAmazon", orders_v),
                                    ("حساب 2: شحنات العملاء → DailyOrdersAmazon (إضافة دايمًا بدون مسح)", orders2_v),
                                    ("مسح DailyOrdersAmazon قبل رفع الحساب 1 (لو مقفول: إضافة تحت آخر صف بدون تكرار)", clear_v),
                                    ("حساب 2: مخزون أمازون → StockAmazon (SELLABLE وكمية > 0، مسح وكتابة)", stock_v)]):
        tk.Checkbutton(c, text=txt, variable=var, bg=CARD, fg=TEXT, selectcolor=ENTRY_BG, activebackground=CARD,
                       activeforeground=TEXT, font=("Segoe UI", 10)).grid(row=10 + r, column=0, sticky="w")

    def save_all(show=True):
        settings["email"] = email_e.get().strip()
        settings["password"] = pass_e.get()
        settings["account_name"] = acc_e.get().strip()
        settings["email2"] = email2_e.get().strip()
        settings["password2"] = pass2_e.get()
        settings["account_name2"] = acc2_e.get().strip()
        settings["days"] = DAYS_CHOICES.get(days_cb.get(), "30")
        try:
            settings["otp_wait"] = int(otp_e.get().strip() or 120)
        except ValueError:
            settings["otp_wait"] = 120
        settings["upload"] = bool(upload_v.get())
        settings["clear_before_upload"] = bool(clear_v.get())
        settings["sync_tacweed"] = bool(tw_v.get())
        settings["run_orders"] = bool(orders_v.get())
        settings["run_orders2"] = bool(orders2_v.get())
        settings["run_stock"] = bool(stock_v.get())
        save_settings(settings)
        if show:
            messagebox.showinfo("✅", "تم حفظ الإعدادات")

    # الخيارات بتتحفظ فورًا أول ما تتغيّر (من غير ما تدوس «حفظ») — عشان تفضل زي ما سيبتها بعد قفل/فتح البرنامج
    for _v in (upload_v, clear_v, tw_v, orders_v, orders2_v, stock_v):
        _v.trace_add("write", lambda *a: save_all(False))
    days_cb.bind("<<ComboboxSelected>>", lambda e: save_all(False))

    btn(c, "💾  حفظ الإعدادات", save_all, color=BLUE, fg="white").grid(row=16, column=0, sticky="w", pady=(10, 0))

    # ── الجدولة ──
    sc = tk.Frame(tab_sched, bg=CARD, padx=20, pady=15)
    sc.pack(fill="x", padx=10, pady=10)
    sched_lb = tk.Listbox(sc, bg=ENTRY_BG, fg=TEXT, font=("Segoe UI", 10), selectbackground=ACCENT,
                          selectforeground="#0f1923", relief="flat", height=6, activestyle="none")
    sched_lb.pack(fill="x")

    def refresh_sched():
        sched_lb.delete(0, tk.END)
        for t in sorted(settings["schedules"]):
            sched_lb.insert(tk.END, f"  ⏰  {t}")

    refresh_sched()
    row = tk.Frame(sc, bg=CARD)
    row.pack(fill="x", pady=(10, 0))
    tk.Label(row, text="الوقت (HH:MM):", bg=CARD, fg=MUTED, font=("Segoe UI", 9)).pack(side="left")
    time_e = entry(row, width=8)
    time_e.pack(side="left", padx=(4, 10))

    def add_sched():
        t = time_e.get().strip()
        try:
            datetime.datetime.strptime(t, "%H:%M")
        except ValueError:
            messagebox.showwarning("تنبيه", "اكتب الوقت بصيغة HH:MM مثل 08:30")
            return
        if t not in settings["schedules"]:
            settings["schedules"].append(t)
            save_settings(settings)
            refresh_sched()
        time_e.delete(0, tk.END)

    def del_sched():
        sel = sched_lb.curselection()
        if sel:
            settings["schedules"].remove(sorted(settings["schedules"])[sel[0]])
            save_settings(settings)
            refresh_sched()

    btn(row, "➕ إضافة موعد", add_sched, color=GREEN, padx=10, pady=5).pack(side="left", padx=2)
    btn(row, "🗑 حذف", del_sched, color=RED, fg="white", padx=10, pady=5).pack(side="left", padx=2)
    status = tk.Label(sc, text="⏸ الجدولة متوقفة", bg=CARD, fg=MUTED, font=("Segoe UI", 10))
    status.pack(anchor="w", pady=(12, 4))

    def toggle():
        if not scheduler_running:
            if not settings["schedules"]:
                messagebox.showwarning("تنبيه", "ضيف موعد واحد على الأقل")
                return
            start_scheduler(lambda: settings, log_box)
            status.config(text="▶ الجدولة شغالة", fg=GREEN)
            tbtn.config(text="⏹ إيقاف الجدولة", bg=RED)
        else:
            stop_scheduler()
            status.config(text="⏸ الجدولة متوقفة", fg=MUTED)
            tbtn.config(text="▶ تشغيل الجدولة", bg=GREEN)

    tbtn = btn(sc, "▶ تشغيل الجدولة", toggle, color=GREEN, padx=15, pady=7)
    tbtn.pack(anchor="w")

    # ── تشغيل + سجل ──
    bottom = tk.Frame(root, bg=BG)
    bottom.pack(fill="x", padx=10, pady=(0, 6))

    def run_now():
        save_all(show=False)
        need = []
        if settings.get("run_orders", True) and not (settings["email"] and settings["password"]):
            need.append("الحساب 1")
        if (settings.get("run_stock", True) or settings.get("run_orders2", True)) and not (settings["email2"] and settings["password2"]):
            need.append("الحساب 2")
        if need:
            messagebox.showwarning("تنبيه", "اكتب الإيميل والباسورد لـ: " + " و ".join(need))
            return
        log_box.delete("1.0", tk.END)
        threading.Thread(target=do_run, args=(dict(settings), log_box), daemon=True).start()

    btn(bottom, "▶  تشغيل الآن", run_now, size=12, padx=20, pady=8).pack(side="left")
    lf = tk.Frame(root, bg=BG, padx=10)
    lf.pack(fill="both", expand=True, pady=(0, 10))
    tk.Label(lf, text="📋 السجل", bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w")
    log_box = scrolledtext.ScrolledText(lf, width=90, height=12, bg=ENTRY_BG, fg=TEXT, font=("Consolas", 9),
                                        relief="flat", insertbackground=ACCENT)
    log_box.pack(fill="both", expand=True)

    if settings.get("schedules"):
        root.after(1000, toggle)
    def _shutdown():
        try:
            stop_scheduler()
        except Exception:
            pass

    return _shutdown
'''

_EMBEDDED['Trendyol_Server'] = r'''# -*- coding: utf-8 -*-
"""
Trendyol Server — بوت تنزيل ملف الطلبات (Order-based) من لوحة تريندول (partner.trendyol.com/sa)
ورفعه لشيت DailyOrdersTrendyol في Google Sheets. (3 متاجر = 3 إيميلات، كل متجر ببروفايل كروم منفصل)

الخطوات لكل متجر: دخول (إيميل -> باسورد -> صورة التحقق بإيديك لو ظهرت) -> صفحة الطلبات -> التواريخ -> Filter
-> Download Excel -> Order-based (Standard view) -> انتظار جاهزية أحدث ملف في Download history -> تنزيل -> قراءة.
بعد ما كل المتاجر تخلص: مسح DailyOrdersTrendyol وكتابة طلبات المتاجر كلها مرة واحدة.
TrendyolProducts: رفع يدوي لملف Products من زرار في الواجهة (مسح وكتابة). AdvertisementsTrendyol: يدوي من التطبيق.
"""
import os
import re
import sys
import csv
import json
import time
import glob
import openpyxl
import datetime
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk, filedialog

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from gspread.utils import rowcol_to_a1
from oauth2client.service_account import ServiceAccountCredentials

# ===== Paths =====
SCRIPT_DIR    = os.environ.get("PLATFORMS_BASE_DIR") or os.path.dirname(os.path.abspath(sys.argv[0] if getattr(sys, "frozen", False) else __file__))
SETTINGS_PATH = os.path.join(os.path.expanduser("~"), "trendyol_settings.json")
# بروفايل كروم ثابت لكل متجر: الجلسة بتتحفظ فيه، فصورة التحقق بتظهر نادرًا بعد أول مرة
CHROME_PROFILE_DIR = os.path.join(os.path.expanduser("~"), "TrendyolChromeProfile")
TY_DOWNLOADS_DIR   = os.path.join(SCRIPT_DIR, "trendyol_downloads")

# نفس المسارات الثابتة بتاعة برنامج Noon
GSHEET_CREDS_DIR = r"\\WIN-181U7G4HFM5\Share For Markting\NOONIMG\Copy"
GSHEET_NAME      = "Complaints"
GSHEET_SCOPE     = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# ===== تريندول =====
LOGIN_URL  = "https://partner.trendyol.com/sa/auth/login"
ORDERS_URL = "https://partner.trendyol.com/sa/orders/shipment-packages/all"
N_STORES   = 3

# ===== الشيتات (نفس أعمدة التطبيق بالظبط TABS_CONFIG) =====
ORDERS_SHEET   = "DailyOrdersTrendyol"
ORDERS_HEADERS = ["Row Key", "Barcode", "Order Date", "Order Number", "Order Status", "Quantity",
                  "Net Unit Price", "Invoice Amount", "Commission Rate", "Image URL", "Product Name", "Date Uploaded"]
PRODUCTS_SHEET   = "TrendyolProducts"
PRODUCTS_HEADERS = ["Content ID", "Barcode", "Product Name", "Image URL", "Date Uploaded"]

DAYS_CHOICES = {"آخر 7 أيام": "7", "آخر 14 يومًا": "14", "آخر 30 يومًا": "30", "آخر 45 يومًا": "45", "آخر 60 يومًا": "60"}


def now_str():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ================= Settings =================
def load_settings():
    default = {
        "stores": [{"name": f"متجر {i + 1}", "email": "", "password": "", "enabled": True} for i in range(N_STORES)],
        "login_wait": 300,         # أقصى انتظار لتسجيل الدخول/صورة التحقق (ثانية)
        "days": "30",
        "upload": True,            # رفع للجوجل شيت بعد التنزيل
        "keep_open": True,         # سيب المتصفحات مفتوحة بين التشغيلات (عشان الجلسة وما تدخلش اليوزر كل مرة)
        "run_orders": False,   # احتياطي: الـ API (تاب Trendyol API) هو الأساسي — فعّله لو الـ API عطل
        "schedules": [],
    }
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, encoding="utf-8") as f:
                data = json.load(f)
            for k, v in default.items():
                data.setdefault(k, v)
            st = list(data.get("stores") or [])
            while len(st) < N_STORES:
                st.append(dict(default["stores"][len(st)]))
            data["stores"] = st[:N_STORES]
            return data
        except Exception:
            pass
    return default


def save_settings(settings):
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


# ================= Log =================
LOG_WIDGET = None
ROOT = None


def log_write(log, text):
    try:
        if log is None:
            print(text, end="")
            return
        log.after(0, lambda: (log.insert(tk.END, text), log.see(tk.END)))
    except Exception:
        print(text, end="")


# ================= Driver =================
_CACHED_CHROMEDRIVER_PATH = None


def _clear_stale_wdm_lock():
    try:
        wdm_dir = os.path.join(os.path.expanduser("~"), ".wdm")
        if os.path.isdir(wdm_dir):
            for fname in os.listdir(wdm_dir):
                if fname.startswith(".wdm-lock"):
                    try:
                        os.remove(os.path.join(wdm_dir, fname))
                    except Exception:
                        pass
    except Exception:
        pass


def _get_chromedriver_path(log=None):
    global _CACHED_CHROMEDRIVER_PATH
    if _CACHED_CHROMEDRIVER_PATH and os.path.exists(_CACHED_CHROMEDRIVER_PATH):
        return _CACHED_CHROMEDRIVER_PATH
    _clear_stale_wdm_lock()
    try:
        _CACHED_CHROMEDRIVER_PATH = ChromeDriverManager().install()
    except Exception as e:
        log_write(log, f"⚠️ فشل تجهيز chromedriver: {str(e)[:200]}\n")
        _clear_stale_wdm_lock()
        _CACHED_CHROMEDRIVER_PATH = ChromeDriverManager().install()
    return _CACHED_CHROMEDRIVER_PATH


def start_driver(log=None, account=1):
    profile_dir = CHROME_PROFILE_DIR + str(account)   # بروفايل منفصل لكل متجر
    os.makedirs(profile_dir, exist_ok=True)
    os.makedirs(TY_DOWNLOADS_DIR, exist_ok=True)
    opts = webdriver.ChromeOptions()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--remote-debugging-port=0")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--disable-background-timer-throttling")
    opts.add_argument("--disable-backgrounding-occluded-windows")
    opts.add_argument("--disable-renderer-backgrounding")
    opts.add_argument(f"--user-data-dir={profile_dir}")
    opts.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    opts.add_experimental_option("prefs", {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        "download.default_directory": TY_DOWNLOADS_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    })
    opts.set_capability("unhandledPromptBehavior", "accept")
    driver = webdriver.Chrome(service=Service(_get_chromedriver_path(log)), options=opts)
    try:
        driver.set_page_load_timeout(60)
        driver.set_script_timeout(30)
    except Exception:
        pass
    return driver


def close_driver(driver):
    try:
        driver.quit()
    except Exception:
        pass


def _is_driver_alive(driver):
    try:
        _ = driver.current_url
        return True
    except Exception:
        return False


def _debug_dump(driver, log, tag):
    """لو حصل خطأ: بنحفظ لقطة شاشة + HTML جنب البرنامج عشان نظبط الـ selectors لو احتاج."""
    try:
        ts = datetime.datetime.now().strftime("%H%M%S")
        base = os.path.join(SCRIPT_DIR, f"trendyol_debug_{tag}_{ts}")
        driver.save_screenshot(base + ".png")
        with open(base + ".html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        log_write(log, f"🧾 اتحفظت لقطة الشاشة والـ HTML: {base}.png / .html\n")
    except Exception:
        pass


# ================= JS helpers =================
JS_LIB = r"""
const norm = s => (s || '').replace(/[‎‏‪-‮ ]/g, ' ').replace(/\s+/g, ' ').trim();
const vis = e => !!(e.getClientRects && e.getClientRects().length);
function deepAll(root, out) {
  out = out || [];
  root.querySelectorAll('*').forEach(e => { out.push(e); if (e.shadowRoot) deepAll(e.shadowRoot, out); });
  return out;
}
const clickable = e => (e.closest && e.closest('button,a,[role=button]')) || e;
// أقصر عنصر مرئي نصه = أحد النصوص (exact) — بنستبعد اللي جوه لوحة التنزيل لو excludePanel
function findText(texts, excludePanel) {
  const res = [];
  document.querySelectorAll('p,span,div,button,a,li').forEach(e => {
    if (!vis(e)) return;
    if (excludePanel && e.closest('[data-testid="report-downloader-panel"]')) return;
    const t = norm(e.textContent);
    if (t && texts.includes(t) && !Array.from(e.children).some(c => norm(c.textContent) === t)) res.push([t.length, e]);   // الأعمق فقط
  });
  res.sort((a, b) => a[0] - b[0]);
  return res.map(x => x[1]);
}
"""


def _js(driver, body, *args):
    return driver.execute_script(JS_LIB + "\n" + body, *args)


def _displayed(driver, css):
    try:
        for el in driver.find_elements(By.CSS_SELECTOR, css):
            if el.is_displayed():
                return el
    except Exception:
        pass
    return None


def _type_into(el, text):
    el.click()
    try:
        el.clear()
    except Exception:
        pass
    el.send_keys(Keys.CONTROL, "a")
    el.send_keys(Keys.DELETE)
    el.send_keys(text)


# ================= اللغة (لازم English) =================
# باقي الخطوات بتدوّر على نصوص إنجليزي (Download Excel / Order-based / أسماء الشهور في التقويم / أعمدة الملف)،
# فأول ما نفتح الموقع بنتأكد إن اللغة English، ولو عربي بنحوّلها. عنصر اللغة في الهيدر: <span class="current-language-text">
_JS_LANG_CUR = r"""
    const el = Array.from(document.querySelectorAll('.current-language-text')).find(vis) || document.querySelector('.current-language-text');
    return el ? norm(el.textContent) : '';
"""
_JS_LANG_OPEN = r"""
    const cur = Array.from(document.querySelectorAll('.current-language-text')).find(vis) || document.querySelector('.current-language-text');
    if (!cur) return 'nospan';
    const t = clickable(cur);
    ['mouseover', 'mouseenter', 'mousedown', 'mouseup'].forEach(n => t.dispatchEvent(new MouseEvent(n, {bubbles: true, cancelable: true, view: window})));
    t.click();
    return 'clicked';
"""
_JS_LANG_PICK = r"""
    const cands = findText(['English'], false).filter(e => !e.closest('.current-language-text'));
    if (!cands.length) return 'nooption';
    const o = clickable(cands[0]);
    ['mouseover', 'mouseenter', 'mousedown', 'mouseup'].forEach(n => o.dispatchEvent(new MouseEvent(n, {bubbles: true, cancelable: true, view: window})));
    o.click();
    return 'picked';
"""


def _current_lang(driver):
    """نص اللغة الحالية اللي في الهيدر (مثلًا «العربية» أو «English») — فاضي لو العنصر مش ظاهر."""
    try:
        return _js(driver, _JS_LANG_CUR) or ""
    except Exception:
        return ""


def _is_english(txt):
    return "english" in (txt or "").lower()


def ensure_english(driver, log):
    """بيتأكد إن لغة الموقع English ولو لأ بيحوّلها. بيرجّع:
    "ok" = كانت English | "switched" = اتحوّلت (الصفحة ممكن تكون اتحمّلت تاني) |
    "nospan" = عنصر اللغة مش موجود في الصفحة (بنكمل زي ما هي) | "failed" = فشل التحويل."""
    cur, t0 = "", time.time()
    while time.time() - t0 < 20:                       # استنى ظهور عنصر اللغة
        cur = _current_lang(driver)
        if cur:
            break
        time.sleep(1)
    if not cur:
        log_write(log, "⚠️ عنصر اللغة مش ظاهر في الصفحة — هكمل من غير ما أغيّر اللغة\n")
        return "nospan"
    if _is_english(cur):
        return "ok"
    log_write(log, f"🌐 لغة الموقع «{cur}» — بحوّلها لـ English...\n")
    for attempt in range(1, 4):
        try:
            _js(driver, _JS_LANG_OPEN)
            time.sleep(1)
            r = _js(driver, _JS_LANG_PICK)
        except Exception as e:
            r = f"{type(e).__name__}"
        if r != "picked":
            log_write(log, f"…محاولة {attempt}: مش لاقي خيار English في القايمة ({r})\n")
            try:
                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            except Exception:
                pass
            time.sleep(1)
            continue
        for _ in range(20):                            # الصفحة ممكن تعمل reload بعد اختيار اللغة
            time.sleep(1)
            if _is_english(_current_lang(driver)):
                log_write(log, "✅ اللغة بقت English\n")
                time.sleep(1.5)
                return "switched"
        log_write(log, f"…محاولة {attempt}: اخترت English بس اللغة لسه ما اتغيّرتش\n")
    _debug_dump(driver, log, "lang")
    log_write(log, "❌ ما قدرتش أحوّل لغة الموقع لـ English\n")
    return "failed"


# ================= تسجيل الدخول =================
class FatalLoginError(RuntimeError):
    """إيميل/باسورد غلط (الصفحة لسه بتطلبه بعد عدة محاولات) — مفيش فايدة من إعادة المحاولة وممكن تقفل الحساب."""


def _on_auth_page(url):
    u = (url or "").lower()
    return "/auth/" in u or "/login" in u


def _is_signed_in_url(url):
    return "partner.trendyol.com" in (url or "") and not _on_auth_page(url)


def _submit_login(driver, pass_el_or_email_el):
    for css in ('button[type="submit"]', 'button[data-testid*="login"]', 'button[data-testid*="submit"]'):
        b = _displayed(driver, css)
        if b is not None:
            try:
                b.click()
                return
            except Exception:
                pass
    pass_el_or_email_el.send_keys(Keys.ENTER)


def _login_step(driver, ctx, log):
    """خطوة واحدة على صفحة الدخول. بيكتب الإيميل/الباسورد بس لو الخانة فاضية (عشان ما يمسحش اللي بتكتبه بإيدك).
    بيرجّع: filled / waiting / none."""
    s = ctx["settings"]
    email_el = _displayed(driver, 'input[name="email"]')
    pass_el = _displayed(driver, 'input[type="password"]')
    did = False
    if email_el is not None and not (email_el.get_attribute("value") or "").strip():
        log_write(log, "✉️ كتابة الإيميل...\n")
        _type_into(email_el, s["email"])
        did = True
    if pass_el is not None and not (pass_el.get_attribute("value") or "").strip():
        log_write(log, "🔑 كتابة الباسورد...\n")
        _type_into(pass_el, s["password"])
        did = True
    if did:
        if ctx["submit_n"] >= 5:
            raise FatalLoginError("اتحاول تسجيل الدخول 5 مرات والصفحة لسه بتطلبه — راجع الإيميل والباسورد (وقف عشان الحساب ما يتقفلش)")
        ctx["submit_n"] += 1
        time.sleep(0.5)
        _submit_login(driver, pass_el if pass_el is not None else email_el)
        time.sleep(3)
        return "filled"
    if email_el is not None or pass_el is not None:
        if time.time() - ctx.get("last_hint", 0) > 25:
            ctx["last_hint"] = time.time()
            log_write(log, "🧩 الصفحة مستنية تحقق (صورة/كود) — حلّه واضغط دخول بنفسك في المتصفح، مستنيك...\n")
        time.sleep(3)
        return "waiting"
    return "none"


def login_to_trendyol(driver, settings, log):
    """بيفتح لوحة تريندول ويعدّي الدخول. لو الجلسة محفوظة بيدخل على طول. صورة التحقق بتتحل بإيدك."""
    if not settings.get("email") or not settings.get("password"):
        log_write(log, "❌ اكتب الإيميل والباسورد للمتجر في الإعدادات الأول\n")
        return False
    ctx = {"settings": settings, "submit_n": 0}
    log_write(log, "🌐 فتح لوحة تريندول...\n")
    driver.get(ORDERS_URL)
    time.sleep(5)
    t0 = time.time()
    quiet = 0
    limit = int(settings.get("login_wait", 300))
    while time.time() - t0 < limit:
        url = driver.current_url
        if _is_signed_in_url(url):
            quiet += 1
            if quiet >= 2:
                log_write(log, "✅ داخل على لوحة تريندول\n")
                ensure_english(driver, log)            # أول ما ندخل: نتأكد إن اللغة English
                return True
            time.sleep(2)
            continue
        quiet = 0
        try:
            state = _login_step(driver, ctx, log)
        except RuntimeError:
            raise
        except Exception as e:
            log_write(log, f"…الصفحة بتتحدّث ({type(e).__name__}) — إعادة محاولة\n")
            time.sleep(2)
            continue
        if state == "none":
            time.sleep(2)
    log_write(log, "❌ انتهى وقت تسجيل الدخول\n")
    _debug_dump(driver, log, "login")
    return False


# ================= تنزيل ملف الطلبات =================
FILTER_BTN = 'button[data-testid="shipment-packages-filter-submit-button"]'
PANEL = '[data-testid="report-downloader-panel"]'
ORDER_ACTION = 'button[data-testid="report-downloader-action-order"]'
HIST_BTN = 'button[data-testid^="download-excel-file-button-"]'



# ================= قفل النوافذ/الطبقات المنبثقة =================
_CLOSE_TEXTS = ["Skip tour", "Close", "Got it", "Okay", "OK", "Skip", "Later", "Not now", "Dismiss", "No thanks", "Maybe later",
                "Kapat", "Tamam", "Anladım", "Sonra", "Atla", "Şimdi değil", "إغلاق", "تخطي", "لاحقا", "حسنا", "×", "✕", "X"]

def dismiss_overlays(driver, log=None):
    """بيقفل أي popup / tour / banner / طبقة معتمة بتغطي الصفحة وبتمنع الضغط. آمن: مفيش حاجة بتتعمل لو مفيش طبقة."""
    try:
        n = _js(driver, r"""
            const texts = arguments[0];
            let closed = 0;
            // 0) الجولة التعريفية (Tour): زرار "Skip tour" أو OK جنب "1/1"
            const skip = findText(['Skip tour', 'Skip Tour', 'Skip'], false)[0];
            if (skip) { clickable(skip).click(); closed++; }
            else {
              const ok = findText(['OK', 'Ok', 'Got it'], false).find(e => {
                const box = e.closest('div');
                let up = e; for (let i = 0; i < 5 && up; i++, up = up.parentElement) {
                  if (/\b\d+\s*\/\s*\d+\b/.test(norm(up.textContent)) && /tour|menu|new/i.test(norm(up.textContent))) return true;
                }
                return false;
              });
              if (ok) { clickable(ok).click(); closed++; }
            }
            // رسالة الخطأ (toast) — زرار X بتاعها
            document.querySelectorAll('[class*=toast],[class*=Toast],[class*=notification],[role=alert]').forEach(t => {
              if (!vis(t) || !/error/i.test(norm(t.textContent))) return;
              const x = Array.from(t.querySelectorAll('button,[role=button],svg,span,i')).filter(vis).pop();
              if (x) { clickable(x).click(); closed++; }
            });
            const inAllowed = e => e.closest('[data-testid="report-downloader-panel"]') || e.closest('.calendar-content');
            // 1) أزرار قفل جوه نوافذ/dialogs/popovers
            const boxes = Array.from(document.querySelectorAll('[role=dialog],[aria-modal=true],[class*=modal],[class*=Modal],[class*=popup],[class*=Popup],[class*=popover],[class*=tour],[class*=Tour],[class*=onboarding],[class*=banner],[class*=Banner],[class*=toast]'))
                                .filter(b => vis(b) && !inAllowed(b));
            for (const b of boxes) {
              const btns = Array.from(b.querySelectorAll('button,[role=button],a,span,div')).filter(vis);
              const hit = btns.find(x => {
                const t = norm(x.textContent), a = ((x.getAttribute('aria-label') || '') + ' ' + (x.getAttribute('data-testid') || '') + ' ' + (x.className && x.className.toString ? x.className.toString() : '')).toLowerCase();
                return (t && t.length <= 20 && texts.includes(t)) || /(^|[-_ ])close([-_ ]|$)|dismiss|kapat/.test(a);
              });
              if (hit) { clickable(hit).click(); closed++; }
            }
            // 2) طبقة معتمة (backdrop) ثابتة بتغطي معظم الشاشة وفاضية من المحتوى — نخفيها
            const W = innerWidth, H = innerHeight;
            document.querySelectorAll('div,section').forEach(e => {
              if (!vis(e) || inAllowed(e) || e.querySelector('[data-testid="report-downloader-panel"],.calendar-content,button,input,table')) return;
              const cs = getComputedStyle(e);
              if (cs.position !== 'fixed' || cs.pointerEvents === 'none') return;
              const r = e.getBoundingClientRect();
              if (r.width >= W * 0.9 && r.height >= H * 0.9 && !norm(e.textContent)) { e.style.display = 'none'; closed++; }
            });
            return closed;
        """, _CLOSE_TEXTS)
        if n:
            try:
                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            except Exception:
                pass
            if log is not None:
                log_write(log, f"🧹 قفلت {n} نافذة/طبقة منبثقة\n")
            time.sleep(0.6)
        return n or 0
    except Exception:
        return 0


def _wait_orders_page(driver, ctx, log, timeout=120):
    """بيستنى زرار Filter. لو الجلسة اتقفلت وظهرت صفحة الدخول بيعدّيها."""
    t0 = time.time()
    while time.time() - t0 < timeout:
        if _displayed(driver, FILTER_BTN) is not None:
            return True
        if _on_auth_page(driver.current_url):
            try:
                _login_step(driver, ctx, log)
            except RuntimeError:
                raise
            except Exception:
                time.sleep(2)
        else:
            time.sleep(2)
    return False


_PICK_WHY = ""


def _pick_day(driver, iso):
    global _PICK_WHY
    """بيدوس على يوم في تقويم تريندول المفتوح. بيقرا الشهر المعروض من العنوان (October 2026) وبيتنقل بالأسهم
    لحد ما يوصل لشهر اليوم المطلوب، وبعدين يدوس اليوم (bl-button.day وid = منتصف الليل بالميلي ثانية)."""
    for _ in range(16):
        r = _js(driver, """
            const realClick = el => { const i = el.shadowRoot && el.shadowRoot.querySelector('button'); (i || el).click(); };   // الضغط على الزرار الداخلي (بيطلع للـ host)
            const [y, m, d] = arguments[0].split('-').map(Number);
            const ms = String(new Date(y, m - 1, d).getTime());
            const days = deepAll(document).filter(x => x.classList && x.classList.contains('day') && /^\\d{10,}$/.test(x.id) && vis(x));
            if (!days.length) return 'nopicker';
            const root = days[0].closest('.calendar-content,.calendar,[id^="popover"]') || days[0].getRootNode();
            const MN = ['january','february','march','april','may','june','july','august','september','october','november','december'];
            const mt = norm(root.textContent).toLowerCase().match(new RegExp('(' + MN.join('|') + ')\\\\s*(\\\\d{4})'));
            let diff = 0;
            if (mt) diff = (y * 12 + (m - 1)) - (Number(mt[2]) * 12 + MN.indexOf(mt[1]));
            if (diff === 0) {
              const c = days.find(x => x.id === ms && !x.hasAttribute('disabled') && !String(x.className).includes('disabled'));
              if (c) { realClick(c); return 'clicked'; }
              if (mt) return 'missing';
            }
            const btns = Array.from(root.querySelectorAll('bl-button,button')).filter(x => vis(x) && !x.closest('.day') && !String(x.className).includes('header-text'));
            if (!btns.length) return 'nonav';
            const goPrev = mt ? diff < 0 : Number(ms) < Math.min(...days.map(x => Number(x.id)));
            const lab = x => (x.getAttribute('aria-label') || x.getAttribute('title') || '').toLowerCase();
            let b = btns.find(x => (x.getAttribute('icon') || '') === (goPrev ? 'arrow_left' : 'arrow_right')) || btns.find(x => goPrev ? /prev|back|left/.test(lab(x)) : /next|forward|right/.test(lab(x)));
            if (!b) b = goPrev ? btns[0] : btns[btns.length - 1];
            realClick(b); return 'nav';
        """, iso)
        _PICK_WHY = r
        if r == "clicked":
            time.sleep(0.7)
            return True
        if r in ("nopicker", "nonav", "missing"):
            return False
        time.sleep(0.6)
    return False


def _fmt_d(d):
    return d.strftime("%d/%m/%Y")


def _set_dates(driver, days, log):
    """كل خانة تاريخ ليها تقويم لوحدها: بنفتح خانة البداية ونختار اليوم، وبعدين خانة النهاية ونختار اليوم."""
    end = datetime.date.today()
    start = end - datetime.timedelta(days=int(days))
    want = (_fmt_d(start), _fmt_d(end))
    js_read = "return [(document.querySelector('#originShipmentStartDate')||{}).value||'', (document.querySelector('#originShipmentEndDate')||{}).value||''];"
    cur = _js(driver, js_read)
    if tuple(cur) == want:
        log_write(log, f"📅 المدى جاهز: {want[0]} → {want[1]}\n")
        return True
    for attempt in range(2):
        dismiss_overlays(driver, log)
        for fid, d in (("#originShipmentStartDate", start), ("#originShipmentEndDate", end)):
            el = _displayed(driver, fid)
            if el is None:
                continue
            try:
                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            except Exception:
                pass
            time.sleep(0.4)
            try:
                el.click()
            except Exception:
                _js(driver, "const e = document.querySelector(arguments[0]); if (e) e.click();", fid)
            time.sleep(1)
            if not _pick_day(driver, d.isoformat()):
                log_write(log, f"…مقدرتش أختار {_fmt_d(d)} من التقويم ({_PICK_WHY})\n")
            time.sleep(0.5)
        try:
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        except Exception:
            pass
        time.sleep(0.5)
        cur = _js(driver, js_read)
        if tuple(cur) == want:
            log_write(log, f"📅 اتظبط المدى: {want[0]} → {want[1]}\n")
            return True
    log_write(log, f"⚠️ مقدرتش أظبط التواريخ (الحالي: {cur[0] or '—'} → {cur[1] or '—'}) — هكمل بالمدى اللي في الصفحة\n")
    try:   # نحفظ شكل التقويم عشان نظبط الـ selectors
        el = _displayed(driver, "#originShipmentStartDate")
        if el is not None:
            el.click()
            time.sleep(1)
        _debug_dump(driver, log, "datepicker")
        html = _js(driver, """
            function ser(n) {
              if (n.nodeType === 3) return n.textContent.replace(/\\s+/g, ' ');
              if (n.nodeType !== 1) return '';
              let a = Array.from(n.attributes).map(x => x.name + '="' + x.value + '"').join(' ');
              let inner = Array.from(n.childNodes).map(ser).join('');
              if (n.shadowRoot) inner = '<#shadow>' + Array.from(n.shadowRoot.childNodes).map(ser).join('') + '</#shadow>' + inner;
              return '<' + n.tagName.toLowerCase() + (a ? ' ' + a : '') + '>' + inner + '</' + n.tagName.toLowerCase() + '>';
            }
            const hosts = deepAll(document).filter(e => e.classList && e.classList.contains('calendar-content'));
            return hosts.slice(0, 2).map(ser).join('\\n\\n') || 'NO calendar-content FOUND';
        """)
        fp = os.path.join(SCRIPT_DIR, "trendyol_debug_calendar.txt")
        with open(fp, "w", encoding="utf-8") as f:
            f.write(str(html))
        log_write(log, f"🧾 اتحفظ شكل التقويم: {fp}\n")
    except Exception:
        pass
    return False


def _open_download_panel(driver):
    """بيفتح قايمة Download Excel لو مش مفتوحة. بيرجّع True لو اللوحة ظاهرة."""
    if _displayed(driver, PANEL) is not None:
        return True
    _js(driver, """
        const t = findText(['Download Excel'], true)[0];
        if (t) clickable(t).click();
    """)
    for _ in range(10):
        time.sleep(0.5)
        if _displayed(driver, PANEL) is not None:
            return True
    return False


def _history(driver):
    """قايمة سجل التنزيل (الأحدث الأول): [{id, t, ready}]. ready = الزرار مش disabled."""
    return _js(driver, """
        return Array.from(document.querySelectorAll('button[data-testid^="download-excel-file-button-"]')).map(b => ({
          id: b.getAttribute('data-testid'),
          t: norm(b.textContent),
          ready: !b.disabled && !b.hasAttribute('disabled')
        }));
    """) or []


def _wait_new_file(before, timeout=180):
    """بيستنى ظهور ملف جديد مكتمل في مجلد التنزيل ويرجّع مساره."""
    t0 = time.time()
    while time.time() - t0 < timeout:
        now = set(glob.glob(os.path.join(TY_DOWNLOADS_DIR, "*")))
        new = [p for p in now - before if not p.lower().endswith((".crdownload", ".tmp"))]
        if new:
            p = max(new, key=os.path.getmtime)
            s1 = os.path.getsize(p)
            time.sleep(1.5)
            if os.path.exists(p) and os.path.getsize(p) == s1 and s1 > 0:
                return p
        time.sleep(1)
    return None


def download_orders_once(driver, settings, ctx, log):
    log_write(log, "\n📦 ملف الطلبات (Order-based)...\n")
    driver.get(ORDERS_URL)
    time.sleep(4)
    if not _wait_orders_page(driver, ctx, log):
        log_write(log, "❌ صفحة الطلبات ما ظهرتش (زرار Filter)\n")
        _debug_dump(driver, log, "orders")
        return None
    time.sleep(1.5)
    dismiss_overlays(driver, log)
    lang = ensure_english(driver, log)                # لازم English قبل التواريخ وقايمة Download Excel
    if lang == "failed":
        log_write(log, "❌ لغة الموقع مش English — الخطوات الجاية بتعتمد على النصوص الإنجليزي، فمش هكمل\n")
        return None
    if lang == "switched":                            # الصفحة اتحمّلت تاني بعد تغيير اللغة
        if not _wait_orders_page(driver, ctx, log, timeout=90):
            log_write(log, "❌ صفحة الطلبات ما ظهرتش بعد تغيير اللغة\n")
            _debug_dump(driver, log, "afterlang")
            return None
        time.sleep(1.5)
    dismiss_overlays(driver, log)
    _set_dates(driver, settings.get("days", "30"), log)
    dismiss_overlays(driver, log)
    _js(driver, "const b = document.querySelector(arguments[0]); if (b) b.click();", FILTER_BTN)
    log_write(log, "🔎 ضغطت Filter\n")
    time.sleep(4)

    dismiss_overlays(driver, log)
    if not _open_download_panel(driver):
        log_write(log, "❌ قايمة «Download Excel» ما فتحتش\n")
        _debug_dump(driver, log, "panel")
        return None
    before_ids = {h["id"] for h in _history(driver)}
    ok = _js(driver, "const b = document.querySelector(arguments[0]); if (!b) return false; b.click(); return true;", ORDER_ACTION)
    if not ok:
        log_write(log, "❌ زرار «Order-based (Standard view)» مش لاقيه\n")
        _debug_dump(driver, log, "orderbtn")
        return None
    log_write(log, "⏳ تم طلب الملف — مستني يجهز...\n")

    target = None
    t0 = time.time()
    n = 0
    while time.time() - t0 < 420:          # 7 دقايق
        time.sleep(5)
        n += 1
        dismiss_overlays(driver)
        if not _open_download_panel(driver):
            continue
        hist = _history(driver)
        new = [h for h in hist if h["id"] not in before_ids]
        ready = [h for h in new if h["ready"]]
        if ready:
            target = ready[0]
            break
        # لو ما ظهرش عنصر جديد بعد دقيقتين (نفس الدقيقة/اتدمج) — نستخدم أحدث ملف جاهز
        if not new and time.time() - t0 > 120 and hist and hist[0]["ready"]:
            log_write(log, "ℹ️ مفيش عنصر جديد في السجل — هنزّل أحدث ملف جاهز\n")
            target = hist[0]
            break
        if n % 6 == 0:                      # كل ~30 ثانية: نقفل ونفتح القايمة عشان السجل يتحدّث
            try:
                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            except Exception:
                pass
            time.sleep(1)
    if not target:
        log_write(log, "❌ الملف ما جهزش في الوقت المحدد\n")
        _debug_dump(driver, log, "wait")
        return None

    log_write(log, f"✅ الملف جاهز: {target['t'][:60]}\n")
    files_before = set(glob.glob(os.path.join(TY_DOWNLOADS_DIR, "*")))
    clicked = _js(driver, """
        const b = Array.from(document.querySelectorAll('button[data-testid^="download-excel-file-button-"]'))
                       .find(x => x.getAttribute('data-testid') === arguments[0]);
        if (!b) return false; b.click(); return true;
    """, target["id"])
    if not clicked:
        log_write(log, "❌ زرار تنزيل الملف مش لاقيه\n")
        _debug_dump(driver, log, "dl")
        return None
    path = _wait_new_file(files_before)
    if not path:
        log_write(log, "❌ الملف ما نزلش\n")
        _debug_dump(driver, log, "nofile")
        return None
    log_write(log, f"💾 اتنزل: {os.path.basename(path)}\n")
    return path


def download_orders(driver, settings, ctx, log, tries=3):
    for i in range(tries):
        try:
            p = download_orders_once(driver, settings, ctx, log)
        except RuntimeError:
            raise
        except Exception as e:
            log_write(log, f"⚠️ خطأ أثناء التنزيل: {type(e).__name__}: {str(e)[:150]}\n")
            p = None
        if p:
            return p
        if i < tries - 1:
            log_write(log, f"🔁 محاولة {i + 2}/{tries}...\n")
            time.sleep(3)
    return None


# ================= قراءة الملفات (نفس منطق التطبيق بالظبط) =================
def _ty_find_col(hdr, *names):
    low = [str(h).strip().lower() for h in hdr]
    for n in names:
        if n in low:
            return low.index(n)
    return None


def parse_orders_xlsx(path):
    """بيقرأ ملف طلبات تريندول ويطلّع الأعمدة المطلوبة بس بترتيب ORDERS_HEADERS (مفيش بيانات عملاء).
    بيرجّع (rows, عدد الصفوف المتخطّية)."""
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    hdr = None
    for r in it:
        if r and any(str(c).strip().lower() == "barcode" for c in r if c is not None) and \
           any(str(c).strip().lower() == "order date" for c in r if c is not None):
            hdr = list(r)
            break
    if hdr is None:
        raise ValueError("مش لاقي صف العناوين (Barcode / Order Date) في الملف")
    c_bc = _ty_find_col(hdr, "barcode")
    c_dt = _ty_find_col(hdr, "order date")
    c_ord = _ty_find_col(hdr, "order number")
    c_pkg = _ty_find_col(hdr, "package no")
    c_st = _ty_find_col(hdr, "order status")
    c_q = _ty_find_col(hdr, "quantity")
    c_inv = _ty_find_col(hdr, "invoice amount")
    c_cm = _ty_find_col(hdr, "commission rate")
    c_nm = _ty_find_col(hdr, "product name")
    c_im = next((i for i, h in enumerate(hdr) if h is not None and str(h).strip().lower().startswith("product image")), None)
    if None in (c_bc, c_dt, c_ord, c_st, c_q, c_inv):
        raise ValueError("الملف ناقص أعمدة مطلوبة (Barcode / Order Date / Order Number / Order Status / Quantity / Invoice Amount)")
    now_s = now_str()
    out, skipped, seen = [], 0, {}

    def cell(r, i):
        return r[i] if (i is not None and i < len(r)) else None

    for r in it:
        bc = str(cell(r, c_bc) or "").strip()
        dv = cell(r, c_dt)
        if not bc or dv is None:
            skipped += 1
            continue
        if isinstance(dv, datetime.datetime):
            dt = dv
        else:
            dt = None
            for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y %H:%M:%S", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
                try:
                    dt = datetime.datetime.strptime(str(dv).strip(), fmt)
                    break
                except Exception:
                    pass
        if dt is None:
            skipped += 1
            continue
        try:
            qty = max(int(float(str(cell(r, c_q) or 1).replace(",", ""))), 1)
        except Exception:
            qty = 1
        try:
            inv = float(str(cell(r, c_inv) or 0).replace(",", ""))
        except Exception:
            inv = 0.0
        order = str(cell(r, c_ord) or "").strip()
        pkg = str(cell(r, c_pkg) or "").strip() if c_pkg is not None else ""
        base_key = f"{order}|{pkg}|{bc}"
        seen[base_key] = seen.get(base_key, 0) + 1
        key = base_key if seen[base_key] == 1 else f"{base_key}#{seen[base_key]}"
        img = ""
        if c_im is not None:
            m = re.search(r'https?://[^"\')\s]+', str(cell(r, c_im) or ""))
            img = m.group(0) if m else ""
        out.append([key, bc, dt.strftime("%Y-%m-%d %H:%M:%S"), order, str(cell(r, c_st) or "").strip(),
                    str(qty), f"{inv / qty:.2f}", f"{inv:.2f}", str(cell(r, c_cm) or "").strip(),
                    img, str(cell(r, c_nm) or "").strip()[:120], now_s])
    wb.close()
    return out, skipped


_TY_CID_RE = re.compile(r"-p-(\d+)")


def parse_products_xlsx(path):
    """ملف Products: Content ID (من رابط Trendyol.com Link بعد -p-) / Barcode / اسم المنتج / صورة 1."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Products"] if "Products" in wb.sheetnames else wb.worksheets[0]
    allrows = list(ws.iter_rows(values_only=True))
    wb.close()
    hi = None
    for i, r in enumerate(allrows[:10]):
        low = [str(c).strip().lower() for c in r if c is not None]
        if "barcode" in low and "trendyol.com link" in low:
            hi = i
            break
    if hi is None:
        raise ValueError("مش لاقي عمودي Barcode و Trendyol.com Link")
    idx = {}
    for j, c in enumerate(allrows[hi]):
        k = str(c).strip().lower() if c is not None else ""
        if k and k not in idx:
            idx[k] = j

    def g(r, name):
        j = idx.get(name)
        v = r[j] if (j is not None and j < len(r)) else ""
        return "" if v is None else str(v).strip()

    out, seen, ts = [], set(), now_str()
    for r in allrows[hi + 1:]:
        bc = g(r, "barcode")
        m = _TY_CID_RE.search(g(r, "trendyol.com link"))
        if not bc or not m or (m.group(1), bc) in seen:
            continue
        seen.add((m.group(1), bc))
        out.append([m.group(1), bc, g(r, "product name")[:120], g(r, "image 1"), ts])
    return out


# ================= Google Sheets =================
def _resolve_gsheet_creds_path():
    try:
        m = glob.glob(os.path.join(GSHEET_CREDS_DIR, "*.json"))
        if m:
            return m[0]
    except Exception:
        pass
    return ""


def _gsheet_append_in_batches(ws, rows, log, batch_size=1000, vio="USER_ENTERED"):
    """بيكتب الصفوف تحت آخر صف موجود على دفعات، وأي فشل بيطلع كخطأ واضح (مش بيتبلع). بعد الكتابة بيتأكد من العدد."""
    if not rows:
        return
    start = len(ws.col_values(1)) + 1               # أول صف فاضي (بعد الهيدر / آخر بيانات)
    need = start + len(rows) - 1
    if ws.row_count < need + 5:
        ws.add_rows(need + 5 - ws.row_count)         # لو الشيت وصل حد الخلايا هيطلع الخطأ هنا بوضوح
    for i in range(0, len(rows), batch_size):
        chunk = rows[i:i + batch_size]
        last = None
        for attempt in range(5):
            try:
                ws.update(values=chunk, range_name=f"A{start + i}", value_input_option=vio)
                last = None
                break
            except gspread.exceptions.APIError as e:
                last = e
                if "429" in str(e) or "quota" in str(e).lower():
                    time.sleep(2 * (2 ** attempt))
                else:
                    break                              # خطأ حقيقي (مش ضغط) — مفيش فايدة من التكرار
        if last is not None:
            raise RuntimeError(f"فشلت كتابة الدفعة {i + 1}–{i + len(chunk)}: {str(last)[:300]}")
        log_write(log, f"   ⬆️ {min(i + batch_size, len(rows))}/{len(rows)} صف\n")
    got = len(ws.col_values(1)) - (start - 1)
    if got < len(rows):
        raise RuntimeError(f"بعد الرفع لقيت {got} صف بس من {len(rows)} في الشيت")


def align_rows_to_sheet_header(sheet_header, rows, headers=None):
    """بيرتّب الصفوف على حسب ترتيب هيدر الشيت الحالي (بالاسم)، وأي عمود في الشيت مش بتاعنا بيتسيب فاضي.
    بيرجّع (header, rows_aligned) — والهيدر ممكن يتزوّد بأعمدة ناقصة."""
    headers = headers or ORDERS_HEADERS
    header = list(sheet_header)
    pos = {}
    for t in headers:
        if t in header:
            pos[t] = header.index(t)
        else:
            header.append(t)
            pos[t] = len(header) - 1
    out = []
    for r in rows:
        row = [""] * len(header)
        for ti, t in enumerate(headers):
            row[pos[t]] = r[ti]
        out.append(row)
    return header, out


def _norm_cell(v):
    t = str(v).strip()
    try:
        return f"{float(t.replace(',', '')):.2f}"
    except ValueError:
        return t.lower()


def sync_rows_to_google_sheets(sheet_name, headers, rows, clear, log, dedupe_cols=None, batch_size=1000, vio="RAW"):
    """clear=True: مسح القديم وكتابة الجديد. clear=False: إضافة تحت آخر صف، وبيتجاهل الصفوف الموجودة بالفعل (حسب dedupe_cols)."""
    if not rows:
        log_write(log, f"⚠️ مفيش صفوف للرفع — شيت {sheet_name} ما اتلمسش\n")
        return False
    creds_path = _resolve_gsheet_creds_path()
    if not creds_path:
        log_write(log, f"⚠️ الرفع متخطى — مفيش ملف .json (Service Account) في: {GSHEET_CREDS_DIR}\n")
        return False
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, GSHEET_SCOPE)
        ss = gspread.authorize(creds).open(GSHEET_NAME)
    except Exception as e:
        log_write(log, f"❌ فشل الاتصال بـ Google Sheets: {str(e)[:300]}\n")
        return False
    try:
        try:
            ws = ss.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            ws = ss.add_worksheet(title=sheet_name, rows="2000", cols=str(len(headers)))
            ws.update(values=[headers], range_name="A1")
        sheet_header = [h.strip() for h in ws.row_values(1)]
        header, aligned = align_rows_to_sheet_header(sheet_header if any(sheet_header) else headers, rows, headers)
        if header != sheet_header:
            ws.update(values=[header], range_name="A1")
        ncols = len(header)
        pos = {t: header.index(t) for t in headers}
        if clear:
            log_write(log, f"🗑️ مسح بيانات {sheet_name} القديمة...\n")
            ws.batch_clear([f"A2:{rowcol_to_a1(max(ws.row_count, 2), ncols)}"])
        elif dedupe_cols:
            from collections import Counter
            kcols = [pos[headers[c]] for c in dedupe_cols]
            have = Counter(tuple(_norm_cell(r[c] if c < len(r) else "") for c in kcols) for r in ws.get_all_values()[1:])
            before, fresh = len(aligned), []
            for r in aligned:
                k = tuple(_norm_cell(r[c]) for c in kcols)
                if have[k] > 0:
                    have[k] -= 1          # موجود بالفعل
                else:
                    fresh.append(r)
            aligned = fresh
            log_write(log, f"➕ إضافة: {len(aligned)} صف جديد من {before} (الباقي موجود قبل كده)\n")
        if not aligned:
            log_write(log, f"✅ شيت {sheet_name} محدّث بالفعل — مفيش صفوف جديدة\n")
            return True
        log_write(log, f"⬆️ رفع {len(aligned)} صف لشيت {sheet_name}...\n")
        _gsheet_append_in_batches(ws, aligned, log, batch_size, vio)
        log_write(log, f"✅ تم تحديث شيت {sheet_name}\n")
        return True
    except Exception as e:
        log_write(log, f"❌ فشل الرفع لـ Google Sheets: {str(e)[:300]}\n")
        return False

def sync_orders_to_google_sheets(rows, log, clear):
    # clear=True: مسح وكتابة (كل المتاجر نجحت). clear=False: إضافة بدون تكرار حسب Row Key (لو متجر فشل — عشان ما نمسحش بياناته القديمة)
    return sync_rows_to_google_sheets(ORDERS_SHEET, ORDERS_HEADERS, rows, clear, log, dedupe_cols=(0,), batch_size=2000, vio="RAW")


def sync_products_to_google_sheets(rows, log):
    return sync_rows_to_google_sheets(PRODUCTS_SHEET, PRODUCTS_HEADERS, rows, True, log, batch_size=3000, vio="RAW")


# ================= التشغيل =================
_RUN_LOCK = threading.Lock()
_DRIVERS = {}          # رقم المتجر -> متصفح مفتوح (بيفضل مفتوح بين التشغيلات عشان الجلسة والتحقق بالصورة)


def _store_settings(settings, n):
    """نسخة من الإعدادات فيها إيميل/باسورد/اسم المتجر رقم n (1..3)."""
    st = dict(settings)
    s = settings["stores"][n - 1]
    st["email"], st["password"], st["store_name"] = s.get("email", ""), s.get("password", ""), s.get("name", f"متجر {n}")
    return st


def _get_driver(store, log):
    """بيرجّع المتصفح المفتوح للمتجر لو لسه شغال، وإلا بيفتح واحد جديد (بروفايل المتجر المحفوظ)."""
    d = _DRIVERS.get(store)
    if d is not None and _is_driver_alive(d):
        log_write(log, "♻️ بستخدم المتصفح المفتوح للمتجر\n")
        return d
    if d is not None:
        close_driver(d)
        _DRIVERS.pop(store, None)
    d = start_driver(log, store)
    _DRIVERS[store] = d
    return d


def close_all_drivers():
    for k in list(_DRIVERS):
        close_driver(_DRIVERS.pop(k))


def _run_job(settings, log, title, download_fn, process_fn, store=1):
    """بيدخل (أو يكمّل في المتصفح المفتوح)، ينزّل، وبعدين يقرا الملف. بيرجّع ناتج process_fn أو None لو فشل."""
    try:
        settings = _store_settings(settings, store)
        log_write(log, f"\n{'─' * 36}\n▶ {title}  ({settings['store_name']}: {settings.get('email') or '—'})\n{'─' * 36}\n")
        if not settings.get("email") or not settings.get("password"):
            log_write(log, f"❌ اكتب إيميل وباسورد {settings['store_name']} في الإعدادات\n")
            return None
        driver = _get_driver(store, log)
        if not login_to_trendyol(driver, settings, log):
            return None
        ctx = {"settings": settings, "submit_n": 0}
        path = download_fn(driver, settings, ctx, log)
        if not path:
            log_write(log, "⚠️ مفيش ملف اتنزل\n")
            return None
        if not settings.get("keep_open", True):
            close_driver(_DRIVERS.pop(store, driver))
        return process_fn(path)
    except FatalLoginError:
        raise
    except Exception as e:
        import traceback
        log_write(log, f"\n💥 خطأ غير متوقع: {e}\n{traceback.format_exc()[-900:]}\n")
        d = _DRIVERS.get(store)
        if d is not None and _is_driver_alive(d):
            _debug_dump(d, log, "fatal")
        return None


def _process_orders(path, log, store_name):
    rows, skipped = parse_orders_xlsx(path)
    if rows:
        dts = [r[2][:10] for r in rows]
        log_write(log, f"✅ {store_name}: {len(rows):,} سطر | {len({r[3] for r in rows}):,} طلب | من {min(dts)} إلى {max(dts)}"
                       + (f" | اتخطى {skipped} سطر ناقص" if skipped else "") + "\n")
    else:
        log_write(log, f"ℹ️ {store_name}: الملف مفيهوش طلبات في المدى ده\n")
    return rows


_STOP = threading.Event()
RETRY_WAIT = 60          # ثانية بين كل محاولة وأخرى


def _wait_or_stop(sec):
    for _ in range(int(sec)):
        if _STOP.is_set():
            return True
        time.sleep(1)
    return _STOP.is_set()


def do_run(settings, log):
    """تشغيلة كاملة: لكل متجر مفعّل (بالترتيب، متصفح منفصل بيفضل مفتوح): تنزيل الطلبات ثم رفعها فورًا.
    أول متجر بيمسح DailyOrdersTrendyol ويكتب، والمتاجر اللي بعده بتضيف تحت الموجود (بدون تكرار حسب Row Key).
    أي متجر بيفشل بيتعاد عليه لحد ما ينجح (وبعدين ننتقل للي بعده). الإيقاف بزرار «إيقاف»."""
    if not _RUN_LOCK.acquire(blocking=False):
        log_write(log, "⚠️ فيه تشغيلة شغالة حاليًا — استنى تخلص\n")
        return
    _STOP.clear()
    try:
        log_write(log, f"\n{'=' * 40}\n🚀 بدء التشغيل {datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n{'=' * 40}\n")
        if settings.get("run_orders", False):
            cleared, skipped = False, []
            for n in range(1, N_STORES + 1):
                s = settings["stores"][n - 1]
                if not s.get("enabled", True) or not s.get("email"):
                    continue
                name = s.get("name") or f"متجر {n}"
                if not s.get("password"):
                    log_write(log, f"❌ {name}: مفيش باسورد — اتخطى\n")
                    skipped.append(name)
                    continue
                attempt = 0
                while not _STOP.is_set():
                    attempt += 1
                    if attempt > 1:
                        log_write(log, f"\n🔁 {name}: محاولة رقم {attempt}\n")
                    try:
                        rows = _run_job(settings, log, f"ملف الطلبات → {name}", download_orders,
                                        lambda path, nm=name: _process_orders(path, log, nm), store=n)
                    except FatalLoginError as e:
                        log_write(log, f"❌ {name}: {e}\n   اتخطى المتجر ده (صحّح الإيميل/الباسورد وشغّل تاني)\n")
                        skipped.append(name)
                        break
                    done = rows is not None
                    if done and rows and settings.get("upload", True):
                        ok = sync_orders_to_google_sheets(rows, log, clear=not cleared)
                        if ok and not cleared:
                            cleared = True
                        done = bool(ok)
                    elif done and not settings.get("upload", True):
                        log_write(log, "ℹ️ الرفع متوقف في الإعدادات\n")
                    if done:
                        break
                    log_write(log, f"⏳ {name}: فشل — هعيد المحاولة بعد {RETRY_WAIT} ثانية (زرار «إيقاف» لو عايز توقف)\n")
                    if _wait_or_stop(RETRY_WAIT):
                        break
                if _STOP.is_set():
                    log_write(log, "⏹ اتوقف التشغيل\n")
                    break
            if skipped:
                log_write(log, f"\n⚠️ متاجر اتخطت: {' | '.join(skipped)}\n")
        log_write(log, "\n🎉 انتهت التشغيلة\n")
    finally:
        _RUN_LOCK.release()


# ================= رفريش خفيف للمتصفحات المفتوحة =================
KEEPALIVE_MIN, KEEPALIVE_MAX = 20 * 60, 30 * 60    # ثانية: رفريش كل 20 لـ 30 دقيقة (عشوائي)
_KA_STOP = threading.Event()
_KA_STARTED = False
EXT_GATE = None     # البرنامج الموحّد بيحط هنا القفل العام (لو بوت تاني شغّال منعملش رفريش)


def _keepalive_loop(log):
    import random
    while not _KA_STOP.is_set():
        if _KA_STOP.wait(random.randint(KEEPALIVE_MIN, KEEPALIVE_MAX)):
            break
        if EXT_GATE is not None and EXT_GATE.locked():
            continue
        # لو فيه تشغيلة شغالة منعملش حاجة (المتصفح بيتستخدم)
        if not _RUN_LOCK.acquire(blocking=False):
            continue
        try:
            for k, d in list(_DRIVERS.items()):
                try:
                    if not _is_driver_alive(d):
                        continue
                    # بنرفريش بس لو الصفحة جوه اللوحة (مش في صفحة الدخول/التحقق عشان منقطعش حد بيحل الصورة)
                    if not _is_signed_in_url(d.current_url):
                        continue
                    d.refresh()
                    log_write(log, f"🔄 رفريش للمتصفح (متجر {k}) {datetime.datetime.now():%H:%M}\n")
                except Exception:
                    pass
        finally:
            _RUN_LOCK.release()


def start_keepalive(log):
    global _KA_STARTED
    if _KA_STARTED:
        return
    _KA_STARTED = True
    threading.Thread(target=_keepalive_loop, args=(log,), daemon=True).start()


def upload_products_file(path, log):
    """رفع يدوي لملف Products → TrendyolProducts (مسح وكتابة)."""
    try:
        log_write(log, f"\n📦 قراءة ملف Products: {os.path.basename(path)}\n")
        rows = parse_products_xlsx(path)
        log_write(log, f"✅ {len(rows):,} منتج (Content ID + Barcode)\n")
        sync_products_to_google_sheets(rows, log)
    except Exception as e:
        log_write(log, f"❌ تعذر رفع ملف Products: {e}\n")


# ================= Scheduler =================
scheduler_running = False


def scheduler_loop(settings_getter, log):
    last_run = {}
    while scheduler_running:
        now_dt = datetime.datetime.now()
        now, today = now_dt.strftime("%H:%M"), now_dt.strftime("%Y-%m-%d")
        st = settings_getter()
        if now in st.get("schedules", []) and last_run.get(now) != today:
            last_run[now] = today
            log_write(log, f"\n🕐 جدولة تلقائية الساعة {now}\n")
            threading.Thread(target=do_run, args=(dict(st), log), daemon=True).start()
        time.sleep(15)


def start_scheduler(settings_getter, log):
    global scheduler_running
    scheduler_running = True
    threading.Thread(target=scheduler_loop, args=(settings_getter, log), daemon=True).start()


def stop_scheduler():
    global scheduler_running
    scheduler_running = False

# ================= UI =================
def build_ui(root):
    """واجهة البوت — بتتبني جوه إطار (root) بيدّيه البرنامج الموحّد."""
    global ROOT
    settings = load_settings()
    ROOT = root
    BG, CARD, ACCENT = "#0f1923", "#1a2535", "#f27a1a"
    TEXT, MUTED, GREEN, RED, BLUE, ENTRY_BG = "#e8edf3", "#7a8fa6", "#2ecc71", "#e74c3c", "#3498db", "#243044"
    root.configure(bg=BG)

    def lbl(parent, text):
        return tk.Label(parent, text=text, bg=CARD, fg=MUTED, font=("Segoe UI", 10), anchor="w")

    def entry(parent, width=45, show=None):
        e = tk.Entry(parent, width=width, bg=ENTRY_BG, fg=TEXT, insertbackground=ACCENT, relief="flat",
                     font=("Segoe UI", 10), show=show)
        e.configure(highlightthickness=1, highlightbackground=MUTED, highlightcolor=ACCENT)
        return e

    def btn(parent, text, cmd, color=ACCENT, fg="#0f1923", size=10, padx=12, pady=6):
        return tk.Button(parent, text=text, command=cmd, bg=color, fg=fg, font=("Segoe UI", size, "bold"),
                         relief="flat", padx=padx, pady=pady, cursor="hand2", activebackground=color)

    header = tk.Frame(root, bg=ACCENT, height=50)
    header.pack(fill="x")
    tk.Label(header, text="Trendyol Tool", bg=ACCENT, fg="#0f1923", font=("Segoe UI", 14, "bold")).pack(side="left", padx=15, pady=10)

    style = ttk.Style()
    style.theme_use("default")
    style.configure("Ty.TNotebook", background=BG, borderwidth=0)
    style.configure("Ty.TNotebook.Tab", background=CARD, foreground=TEXT, padding=[12, 6], font=("Segoe UI", 10))
    style.map("Ty.TNotebook.Tab", background=[("selected", ACCENT)], foreground=[("selected", "#0f1923")])
    nb = ttk.Notebook(root, style="Ty.TNotebook")
    nb.pack(fill="both", expand=True, padx=10, pady=8)
    tab_main = tk.Frame(nb, bg=BG)
    tab_sched = tk.Frame(nb, bg=BG)
    nb.add(tab_main, text="⚙️  الإعدادات")
    nb.add(tab_sched, text="🕐  الجدولة")

    # ── الإعدادات ──
    c = tk.Frame(tab_main, bg=CARD, padx=20, pady=15)
    c.pack(fill="x", padx=10, pady=10)
    for col, t in enumerate(["مفعّل", "اسم المتجر", "📧  الإيميل", "🔑  الباسورد"]):
        lbl(c, t).grid(row=0, column=col, sticky="w", padx=(0, 10))
    store_w = []
    for i in range(N_STORES):
        s = settings["stores"][i]
        en_v = tk.BooleanVar(value=bool(s.get("enabled", True)))
        tk.Checkbutton(c, variable=en_v, bg=CARD, selectcolor=ENTRY_BG, activebackground=CARD).grid(row=i + 1, column=0, sticky="w")
        e_name, e_mail, e_pass = entry(c, width=14), entry(c, width=34), entry(c, width=22, show="•")
        e_name.insert(0, s.get("name", f"متجر {i + 1}"))
        e_mail.insert(0, s.get("email", ""))
        e_pass.insert(0, s.get("password", ""))
        e_name.grid(row=i + 1, column=1, sticky="w", pady=4, padx=(0, 10))
        e_mail.grid(row=i + 1, column=2, sticky="w", pady=4, padx=(0, 10))
        e_pass.grid(row=i + 1, column=3, sticky="w", pady=4)
        store_w.append((en_v, e_name, e_mail, e_pass))

    lbl(c, "📅  مدة الطلبات").grid(row=5, column=1, columnspan=2, sticky="w", pady=(12, 0))
    days_cb = ttk.Combobox(c, values=list(DAYS_CHOICES.keys()), state="readonly", width=25)
    inv = {v: k for k, v in DAYS_CHOICES.items()}
    days_cb.set(inv.get(str(settings.get("days", "30")), "آخر 30 يومًا"))
    days_cb.grid(row=6, column=1, columnspan=2, sticky="w", pady=(0, 8))
    lbl(c, "⏱  أقصى انتظار للدخول / صورة التحقق (ثانية)").grid(row=7, column=1, columnspan=3, sticky="w")
    wait_e = entry(c, width=10)
    wait_e.insert(0, str(settings.get("login_wait", 300)))
    wait_e.grid(row=8, column=1, sticky="w", pady=(0, 8))
    upload_v = tk.BooleanVar(value=bool(settings.get("upload", True)))
    orders_v = tk.BooleanVar(value=bool(settings.get("run_orders", False)))
    keep_v = tk.BooleanVar(value=bool(settings.get("keep_open", True)))
    for r, (txt, var) in enumerate([("تنزيل ملف الطلبات لكل متجر → DailyOrdersTrendyol (احتياطي لو الـ API عطل — الأساسي دلوقتي Trendyol API)", orders_v),
                                    ("سيب متصفح كل متجر مفتوح بين التشغيلات (عشان الجلسة وصورة التحقق)", keep_v),
                                    ("رفع الملفات لـ Google Sheets بعد التنزيل", upload_v)]):
        tk.Checkbutton(c, text=txt, variable=var, bg=CARD, fg=TEXT, selectcolor=ENTRY_BG, activebackground=CARD,
                       activeforeground=TEXT, font=("Segoe UI", 10)).grid(row=9 + r, column=0, columnspan=4, sticky="w")

    def save_all(show=True):
        for i, (en_v, e_name, e_mail, e_pass) in enumerate(store_w):
            settings["stores"][i] = {"name": e_name.get().strip() or f"متجر {i + 1}", "email": e_mail.get().strip(),
                                     "password": e_pass.get(), "enabled": bool(en_v.get())}
        settings["days"] = DAYS_CHOICES.get(days_cb.get(), "30")
        try:
            settings["login_wait"] = int(wait_e.get().strip() or 300)
        except ValueError:
            settings["login_wait"] = 300
        settings["upload"] = bool(upload_v.get())
        settings["run_orders"] = bool(orders_v.get())
        settings["keep_open"] = bool(keep_v.get())
        save_settings(settings)
        if show:
            messagebox.showinfo("✅", "تم حفظ الإعدادات")

    # الخيارات بتتحفظ فورًا أول ما تتغيّر (من غير ما تدوس «حفظ»)
    for _v in [upload_v, orders_v, keep_v] + [w[0] for w in store_w]:
        _v.trace_add("write", lambda *a: save_all(False))
    days_cb.bind("<<ComboboxSelected>>", lambda e: save_all(False))

    brow = tk.Frame(c, bg=CARD)
    brow.grid(row=13, column=0, columnspan=4, sticky="w", pady=(10, 0))
    btn(brow, "💾  حفظ الإعدادات", save_all, color=BLUE, fg="white").pack(side="left")

    def pick_products():
        p = filedialog.askopenfilename(title="اختار ملف Products", filetypes=[("Excel", "*.xlsx *.xlsm"), ("All", "*.*")])
        if p:
            if not messagebox.askyesno("تأكيد", "هيتم مسح TrendyolProducts وكتابة الملف الجديد مكانه. متأكد؟"):
                return
            threading.Thread(target=upload_products_file, args=(p, log_box), daemon=True).start()

    btn(brow, "📦  رفع ملف Products (يدوي — مسح وكتابة)", pick_products, color=GREEN).pack(side="left", padx=10)
    tk.Label(c, text="الإعلانات (AdvertisementsTrendyol): رفع يدوي من التطبيق.", bg=CARD, fg=MUTED,
             font=("Segoe UI", 9)).grid(row=14, column=0, columnspan=4, sticky="w", pady=(8, 0))

    # ── الجدولة ──
    sc = tk.Frame(tab_sched, bg=CARD, padx=20, pady=15)
    sc.pack(fill="x", padx=10, pady=10)
    sched_lb = tk.Listbox(sc, bg=ENTRY_BG, fg=TEXT, font=("Segoe UI", 10), selectbackground=ACCENT,
                          selectforeground="#0f1923", relief="flat", height=6, activestyle="none")
    sched_lb.pack(fill="x")

    def refresh_sched():
        sched_lb.delete(0, tk.END)
        for t in sorted(settings["schedules"]):
            sched_lb.insert(tk.END, f"  ⏰  {t}")

    refresh_sched()
    row = tk.Frame(sc, bg=CARD)
    row.pack(fill="x", pady=(10, 0))
    tk.Label(row, text="الوقت (HH:MM):", bg=CARD, fg=MUTED, font=("Segoe UI", 9)).pack(side="left")
    time_e = entry(row, width=8)
    time_e.pack(side="left", padx=(4, 10))

    def add_sched():
        t = time_e.get().strip()
        try:
            datetime.datetime.strptime(t, "%H:%M")
        except ValueError:
            messagebox.showwarning("تنبيه", "اكتب الوقت بصيغة HH:MM مثل 08:30")
            return
        if t not in settings["schedules"]:
            settings["schedules"].append(t)
            save_settings(settings)
            refresh_sched()
        time_e.delete(0, tk.END)

    def del_sched():
        sel = sched_lb.curselection()
        if sel:
            settings["schedules"].remove(sorted(settings["schedules"])[sel[0]])
            save_settings(settings)
            refresh_sched()

    btn(row, "➕ إضافة موعد", add_sched, color=GREEN, padx=10, pady=5).pack(side="left", padx=2)
    btn(row, "🗑 حذف", del_sched, color=RED, fg="white", padx=10, pady=5).pack(side="left", padx=2)
    status = tk.Label(sc, text="⏸ الجدولة متوقفة", bg=CARD, fg=MUTED, font=("Segoe UI", 10))
    status.pack(anchor="w", pady=(12, 4))

    def toggle():
        if not scheduler_running:
            if not settings["schedules"]:
                messagebox.showwarning("تنبيه", "ضيف موعد واحد على الأقل")
                return
            start_scheduler(lambda: settings, log_box)
            status.config(text="▶ الجدولة شغالة", fg=GREEN)
            tbtn.config(text="⏹ إيقاف الجدولة", bg=RED)
        else:
            stop_scheduler()
            status.config(text="⏸ الجدولة متوقفة", fg=MUTED)
            tbtn.config(text="▶ تشغيل الجدولة", bg=GREEN)

    tbtn = btn(sc, "▶ تشغيل الجدولة", toggle, color=GREEN, padx=15, pady=7)
    tbtn.pack(anchor="w")
    tk.Label(sc, text="ملحوظة: لو صورة التحقق ظهرت وقت الجدولة والشاشة مش قدامك، المتجر هيتعدّى وهيتسجّل فشله في السجل.",
             bg=CARD, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(8, 0))

    # ── تشغيل + سجل ──
    bottom = tk.Frame(root, bg=BG)
    bottom.pack(fill="x", padx=10, pady=(0, 6))

    def run_now():
        save_all(show=False)
        act = [s for s in settings["stores"] if s.get("enabled") and s.get("email") and s.get("password")]
        if settings.get("run_orders", False) and not act:
            messagebox.showwarning("تنبيه", "فعّل متجر واحد على الأقل واكتب الإيميل والباسورد")
            return
        log_box.delete("1.0", tk.END)
        threading.Thread(target=do_run, args=(dict(settings), log_box), daemon=True).start()

    btn(bottom, "▶  تشغيل الآن", run_now, size=12, padx=20, pady=8).pack(side="left")
    btn(bottom, "⏹  إيقاف", lambda: (_STOP.set(), log_write(log_box, "\n⏹ طلب إيقاف — هيقف بعد الخطوة الحالية\n")),
        color=RED, fg="white", size=12, padx=20, pady=8).pack(side="left", padx=8)
    lf = tk.Frame(root, bg=BG, padx=10)
    lf.pack(fill="both", expand=True, pady=(0, 10))
    tk.Label(lf, text="📋 السجل", bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w")
    log_box = scrolledtext.ScrolledText(lf, width=90, height=12, bg=ENTRY_BG, fg=TEXT, font=("Consolas", 9),
                                        relief="flat", insertbackground=ACCENT)
    log_box.pack(fill="both", expand=True)

    def _shutdown():
        try:
            stop_scheduler()
            _KA_STOP.set()
            close_all_drivers()
        except Exception:
            pass

    start_keepalive(log_box)
    if settings.get("schedules"):
        root.after(1000, toggle)
    return _shutdown
'''


_EMBEDDED['NoonAPI'] = r'''# -*- coding: utf-8 -*-
"""
Noon_API_Orders.py — تجربة: طلبات نون من الـ API → تاب جديد في Google Sheet (DailyOrders_API)
الشيت القديم (DailyOrders) مابيتلمسش خالص. قراءة فقط من نون.

المتطلبات:  py -m pip install requests pyjwt cryptography gspread oauth2client
ملفات مفاتيح نون: بتتختار من الواجهة وبتتحفظ مساراتها بس في noon_stores.json (المفاتيح نفسها ماتتنسخش ولا بتتبعت لأي حد).
"""
import csv, io, json, os, glob, time, uuid, zipfile, threading, datetime as dt
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox

import jwt
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials

BASE = "https://noon-api-gateway.noon.partners"
UA = {"User-Agent": "NoonApiClient/1.0"}
APP_DIR = os.environ.get("PLATFORMS_BASE_DIR") or os.path.dirname(os.path.abspath(__file__))
LOG_HOOK = None   # بيتحط من البرنامج الموحد → السجل المشترك
SETTINGS_FILE = os.path.join(APP_DIR, "noon_api_settings.json")
STORES_FILE = os.path.join(APP_DIR, "noon_stores.json")

GSHEET_CREDS_DIR = r"\\WIN-181U7G4HFM5\Share For Markting\NOONIMG\Copy"
GSHEET_NAME = "Complaints"
TEST_TAB = "DailyOrders_API"
STOCK_TAB = "Inventory_API_Raw"
LIVE_TAB = "LIVE_API"
HIST_TAB = "PriceHistory_API"
HIST_HEADERS = ["Date", "Store", "SKU", "Price"]
FIN_TAB = "Finance_API_Raw"
COM_TAB = "COM_API"
COM_HEADERS = ["SKU", "مصاريف توصيل", "العمولة", "توصيل FBN", "عمولة FBN", "توصيل FBB", "عمولة FBB", "تاريخ آخر بيع FBN", "تاريخ آخر بيع FBB", "Date Uploaded"]
INV_TAB = "Inventory_API"
IMG_TAB = "Images_API"          # كاش صور وأسماء المنتجات من نون API (جوه شيتنا)
IMG_HEADERS = ["SKU", "Image URL", "Title", "Status", "Checked At", "Family", "Product Type", "Brand", "Category"]
IMG_TIME_BUDGET = 480           # أقصى وقت (ثواني) لجلب الصور في التشغيلة الواحدة
IMG_MAX_PER_RUN = 300           # أقصى عدد SKU جديد بنسأل عنه في التشغيلة الواحدة
INV_HEADERS = ["SKU", "Warehouse", "Stock", "Monthly Sales", "Image URL", "Date Uploaded"]
GSHEET_SCOPE = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

HEADERS = ["SKU", "Order Timestamp", "Status", "Price", "Quantity", "Fulfillment Model", "Date Uploaded", "family",
           "Store", "Order Timestamp (UTC)", "order_nr", "item_nr", "warehouse_code", "partner_sku"]
EXCLUDE_STATUS = {"cancelled", "canceled", "undeliverable"}


# ---------------- مفاتيح نون ----------------
def read_key_file(path):
    with open(path, encoding="utf-8-sig") as fh:
        raw = json.load(fh)

    def find(o):
        if isinstance(o, dict):
            if ("key_id" in o or "keyId" in o) and ("private_key" in o or "privateKey" in o):
                return o
            for v in o.values():
                r = find(v)
                if r:
                    return r
        if isinstance(o, list):
            for v in o:
                r = find(v)
                if r:
                    return r
        return None
    k = find(raw)
    if not k:
        raise ValueError("مالقيتش key_id / private_key في الملف")
    k = dict(k)
    for a, b in (("key_id", "keyId"), ("private_key", "privateKey"), ("project_code", "projectCode")):
        if not k.get(a) and k.get(b):
            k[a] = k[b]
    if not k.get("project_code") and k.get("default_project_code"):
        k["project_code"] = k["default_project_code"]
    miss = [x for x in ("key_id", "private_key", "project_code") if not k.get(x)]
    if miss:
        raise ValueError("ناقص: " + ", ".join(miss))
    return k


class Store:
    def __init__(self, path):
        self.path = path
        self.name = os.path.splitext(os.path.basename(path))[0]
        self.key = read_key_file(path)
        self.sess = None
        self.lock = threading.Lock()


def login(st, log):
    s = requests.Session()
    s.headers.update(UA)
    try:   # اتصالات أكتر مع بعض (الافتراضي 10 بس وده كان بيخنق الطلبات المتوازية)
        from requests.adapters import HTTPAdapter
        s.mount('https://', HTTPAdapter(pool_connections=48, pool_maxsize=48))
    except Exception:
        pass
    tok = jwt.encode({"sub": st.key["key_id"], "iat": int(time.time()), "jti": str(uuid.uuid4())},
                     st.key["private_key"], algorithm="RS256")
    r = s.post(f"{BASE}/identity/public/v1/api/login",
               json={"token": tok, "default_project_code": st.key["project_code"]}, timeout=30)
    if r.status_code != 200:
        log(f"[{st.name}] ❌ تسجيل الدخول فشل: {r.status_code}")
        return None
    return s


def api_post(st, path, body, log):
    for attempt in (1, 2):
        with st.lock:
            if st.sess is None:
                st.sess = login(st, log)
            s = st.sess
        if s is None:
            return None
        r = None
        for t in (1, 2, 3):     # نون أحيانًا بتتأخر في الرد — نعيد المحاولة بدل ما نوقّف التشغيلة كلها
            try:
                r = s.post(BASE + path, json=body, headers={"Accept": "application/json"}, timeout=60)
                break
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
                log(f"[{st.name}] ⚠️ نون ماردتش في الوقت ({type(e).__name__}) — محاولة {t}/3")
                if t == 3:
                    log(f"[{st.name}] ❌ {path}: فشل بعد 3 محاولات")
                    return None
                time.sleep(5 * t)
        if r.status_code == 401 and attempt == 1:
            with st.lock:
                st.sess = None
            continue
        return r
    return None


EXPORT_LOG = os.path.join(APP_DIR, "noon_exports_log.json")
_elock = threading.Lock()


HANDOFF_DIR = os.path.join(APP_DIR, "noon_orders_handoff")


ACC_ORDERS_START = dt.date(2026, 9, 1)   # لازم يطابق DEFAULT_START في «حسابات نون»
ACC_ORDERS_BACK = 60                      # لازم يطابق ORD_BACKFILL هناك


def acct_orders_need_from(st):
    """لو «حسابات نون» لسه مسحبتش أوردرات المتجر ده من بداية الدفتر (− 60 يوم) بترجع التاريخ ده، وإلا None (يكفي آخر 30 يوم)."""
    want = ACC_ORDERS_START - dt.timedelta(days=ACC_ORDERS_BACK)
    try:
        s = json.load(open(os.path.join(APP_DIR, "noon_accounts", "state.json"), encoding="utf-8")).get(st.name, {})
        if s.get("orders_from") == ACC_ORDERS_START.isoformat() and s.get("orders_backfill_done"):
            return None
    except Exception:
        pass
    return want


def save_orders_handoff(st, d_from, d_to, rows, keep_wider=False):
    """بنسيب أوردرات الـ API اللي سحبناها هنا لخطوة «حسابات نون» تاخدها بدل ما تسحبها تاني من نون."""
    if not rows:
        return
    try:
        fp = os.path.join(HANDOFF_DIR, f"{st.name}.json")
        if keep_wider and os.path.exists(fp):
            old = json.load(open(fp, encoding="utf-8"))
            if old.get("from", "9999") <= str(d_from) and time.time() - float(old.get("ts", 0)) < 1800:
                return
        os.makedirs(HANDOFF_DIR, exist_ok=True)
        tmp = os.path.join(HANDOFF_DIR, f"{st.name}.json.tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"from": str(d_from), "to": str(d_to), "ts": time.time(), "rows": rows}, f, ensure_ascii=False)
        os.replace(tmp, os.path.join(HANDOFF_DIR, f"{st.name}.json"))
    except Exception:
        pass


def log_export(st, code, ec, params=None):
    with _elock:
        try:
            d = json.load(open(EXPORT_LOG, encoding="utf-8"))
        except Exception:
            d = []
        d.append({"store": st.name, "code": code, "ec": ec, "ts": time.time(), "country": (params or {}).get("country", "")})
        d = d[-200:]
        try:
            json.dump(d, open(EXPORT_LOG, "w", encoding="utf-8"), ensure_ascii=False)
        except Exception:
            pass


def recent_exports(st, code, hours=36):
    """أكواد التصديرات اللي الأداة طلبتها قبل كده (الأحدث الأول)، خلال آخر X ساعة."""
    try:
        d = json.load(open(EXPORT_LOG, encoding="utf-8"))
    except Exception:
        return []
    out = [x["ec"] for x in d if x.get("store") == st.name and x.get("code") == code
           and time.time() - x.get("ts", 0) < hours * 3600]
    return list(reversed(out))


def export_state(st, ec, log):
    """(حالة, رابط) لتصدير موجود — من غير ما يطلب جديد."""
    rs = api_post(st, "/impex/v1/export/status", {"export_code": ec}, log)
    if rs is None:
        return "", None
    try:
        sd = rs.json()
    except Exception:
        return "", None
    return str(sd.get("export_status", "")).upper(), sd.get("download_url")


def ready_export(st, code, params, log, max_age_h=48):
    """آخر تصدير جاهز (اللي الأداة طلبته قبل كده، لنفس المتجر والنوع والدولة) خلال آخر max_age_h ساعة.
    بيرجع (export_code, عمره بالساعات) أو None لو مفيش واحد جاهز."""
    try:
        d = json.load(open(EXPORT_LOG, encoding="utf-8"))
    except Exception:
        return None
    ctry = (params or {}).get("country", "")
    for x in reversed(d):
        if x.get("store") != st.name or x.get("code") != code:
            continue
        age = time.time() - x.get("ts", 0)
        if age > max_age_h * 3600:
            break
        if ctry and x.get("country") and x.get("country") != ctry:
            continue
        stt, url = export_state(st, x["ec"], log)
        if url:
            return x["ec"], age / 3600.0
    return None


def request_only(st, code, params, log):
    """يطلب تصدير جديد من غير ما يستنّاه (للمرة الجاية). لو في واحد شغال عند نون يتجاهل."""
    r = api_post(st, "/impex/v1/export/create", {"export_category_code": code, "params": params}, log)
    if r is None:
        return
    try:
        d = r.json()
    except Exception:
        d = {}
    if r.status_code == 200 and d.get("export_code"):
        log_export(st, code, d["export_code"], params)
        log(f"[{st.name}] 🔄 {code}: طلبت تصدير جديد ({d['export_code']}) للمرة الجاية — مش هستنّاه")
    elif r.status_code == 400 and "already running" in (r.text or "").lower():
        log(f"[{st.name}] 🔄 {code}: في تصدير جديد شغال بالفعل عند نون — تمام")
    else:
        log(f"[{st.name}] ⚠️ {code}: ماقدرتش أطلب تصدير جديد للمرة الجاية ({r.status_code}) {(r.text or '')[:200]}")


def run_export(st, code, params, log, stop, max_wait=40, prefer_ready=False):
    """يطلب تقرير، يستنى، ينزّله، ويرجع صفوف dict (من الذاكرة — من غير ما يحفظ ملف).
    prefer_ready=True (للتقارير البطيئة زي مخزون FBN): ينزّل آخر تصدير جاهز فورًا، وبعدين يطلب واحد جديد من غير ما يستنّاه."""
    if prefer_ready:
        got = ready_export(st, code, params, log)
        if got:
            log(f"[{st.name}] ⚡ {code}: باخد آخر تصدير جاهز ({got[0]}) عمره {got[1]:.1f} ساعة بدل ما أستنى")
            rows = fetch_export(st, code, got[0], log, stop)
            request_only(st, code, params, log)
            if rows:
                return rows
    for wait_i in range(max_wait):
        r = api_post(st, "/impex/v1/export/create", {"export_category_code": code, "params": params}, log)
        if r is None:
            return None
        try:
            d = r.json()
        except Exception:
            d = {}
        if r.status_code == 400 and "already running" in (r.text or "").lower() and not stop.is_set():
            if wait_i == max_wait - 1:
                log(f"[{st.name}] ⌛ {code}: لسه في تصدير شغال — هكمل من غيره")
                return None
            if wait_i == 0:
                log(f"[{st.name}] ⏳ {code}: في تصدير تاني من نفس النوع شغال عند نون — هستنى وأعيد المحاولة...")
                for old in recent_exports(st, code, 48):
                    stt, url = export_state(st, old, log)
                    if url or (stt and not any(w in stt for w in ("FAIL", "ERROR", "CANCEL"))):
                        log(f"[{st.name}] ↪️ ده غالباً تصدير قديم من عندنا ({old}) — هاخده بدل ما أطلب جديد")
                        return fetch_export(st, code, old, log, stop)
            time.sleep(20)
            continue
        if r.status_code == 400 and "ratelimit" in (r.text or "").lower() and not stop.is_set() and wait_i < 6:
            log(f"[{st.name}] ⏳ {code}: نون قالت Ratelimit — هستنى دقيقة وأعيد ({wait_i + 1}/6)")
            for _ in range(60):
                if stop.is_set():
                    break
                time.sleep(1)
            continue
        break
    if r.status_code != 200 or not d.get("export_code"):
        log(f"[{st.name}] ❌ طلب {code}: {r.status_code} {(r.text or '')[:300]}")
        return None
    ec = d["export_code"]
    log_export(st, code, ec, params)
    log(f"[{st.name}] ⏳ {code} اتطلب ({ec})")
    return fetch_export(st, code, ec, log, stop)


def fetch_export(st, code, ec, log, stop):
    """ينتظر export_code موجود ويحمّله (من غير ما يطلب تصدير جديد)."""
    url = None
    for _ in range(180):
        if stop.is_set():
            return None
        time.sleep(5)
        rs = api_post(st, "/impex/v1/export/status", {"export_code": ec}, log)
        if rs is None:
            continue
        try:
            sd = rs.json()
        except Exception:
            continue
        if sd.get("download_url"):
            url = sd["download_url"]
            break
        if any(w in str(sd.get("export_status")).upper() for w in ("FAIL", "ERROR", "CANCEL")):
            log(f"[{st.name}] ❌ {code}: {sd}")
            return None
    if not url:
        log(f"[{st.name}] ⌛ {code} ماجهزش في الوقت المحدد")
        return None
    rr = requests.get(url, timeout=600)
    if rr.status_code in (401, 403) and st.sess is not None:
        rr = st.sess.get(url, timeout=600)
    if rr.status_code != 200:
        log(f"[{st.name}] ❌ فشل التحميل {rr.status_code}")
        return None
    data = rr.content
    if data[:2] == b"PK":
        z = zipfile.ZipFile(io.BytesIO(data))
        names = [n for n in z.namelist() if n.lower().endswith(".csv")]
        if not names:
            log(f"[{st.name}] ⚠️ الملف مضغوط ومفيهوش CSV")
            return None
        data = z.read(names[0])
    rows = list(csv.DictReader(io.StringIO(data.decode("utf-8-sig", errors="replace"))))
    log(f"[{st.name}] ✅ {code}: {len(rows)} صف")
    return rows


CACHE_FILE = os.path.join(APP_DIR, "noon_catalog_cache2.json")
CACHE_HOURS = 12


def local_catalog(st, folders, log, max_age_days=2):
    """آخر ملف كتالوج منزّل بالفعل (CSV) للمتجر ده — لو عمره النهارده/أمس بنستخدمه من غير ما نطلب من نون."""
    code = str(st.key.get("project_code", "")).upper().replace("PRJ", "").strip()
    cands = []
    for d in folders:
        if d and os.path.isdir(d):
            cands += glob.glob(os.path.join(d, "*catalog*export*.csv"))
    cands = sorted(set(cands), key=lambda f: os.path.getmtime(f), reverse=True)
    for f in cands:
        if time.time() - os.path.getmtime(f) > max_age_days * 86400:
            break
        try:
            with open(f, encoding="utf-8-sig", newline="") as fh:
                rows = list(csv.DictReader(fh))
        except Exception:
            continue
        if not rows or "partner_sku" not in rows[0]:
            continue
        mine = [r for r in rows if str(r.get("id_partner", "")).strip() == code] if "id_partner" in rows[0] else rows
        if mine:
            age_h = int((time.time() - os.path.getmtime(f)) / 3600)
            log(f"[{st.name}] 🗂️ كتالوج جاهز منزّل: {os.path.basename(f)} (عمره {age_h} ساعة، {len(mine)} صف)")
            return mine
    return None


def get_catalog(st, country, log, stop, ready_code="", folders=None, force=False):
    """الكتالوج بيتغير ببطء — بنحفظه 12 ساعة عشان التشغيلات المتكررة تبقى أسرع ومانتصادمش مع تصدير شغال."""
    try:
        cache = json.load(open(CACHE_FILE, encoding="utf-8"))
    except Exception:
        cache = {}
    c = cache.get(st.name)
    if force:
        log(f"[{st.name}] 🔄 تحديث كتالوج طازة (من غير كاش ولا ملفات قديمة)")
        c, folders = None, None
    if c and time.time() - c.get("ts", 0) < CACHE_HOURS * 3600 and c.get("country") == country and c.get("rows"):
        log(f"[{st.name}] 🗂️ الكتالوج من الكاش ({len(c['rows'])} صف، عمره {int((time.time() - c['ts']) / 60)} دقيقة)")
        return c["rows"]
    if folders and not ready_code:
        lc = local_catalog(st, folders, log)
        if lc:
            return lc
    if not ready_code and not force:
        for old in recent_exports(st, "noon_catalog_catalogexport", 36):
            stt, url = export_state(st, old, log)
            if url:
                log(f"[{st.name}] 🗂️ لقيت كتالوج جاهز من تصدير قريب ({old}) — هنزّله بدل ما أطلب جديد")
                ready_code = old
                break
    if ready_code:
        log(f"[{st.name}] 🗂️ باحمّل الكتالوج من الكود الجاهز {ready_code}")
        rows = fetch_export(st, "noon_catalog_catalogexport", ready_code, log, stop)
    else:
        rows = run_export(st, "noon_catalog_catalogexport", {"country": country, "noon_status": "live"}, log, stop,
                          max_wait=(30 if force else 9))
    if rows:
        cache[st.name] = {"ts": time.time(), "country": country, "rows": rows}
        try:
            json.dump(cache, open(CACHE_FILE, "w", encoding="utf-8"), ensure_ascii=False)
        except Exception:
            pass
    return rows


# ---------------- المنطق ----------------
def fnum(x):
    try:
        return float(str(x).replace(",", "").strip())
    except Exception:
        return ""


def our_price(r):
    """سعرنا اللي حطيناه (sale_price) — مش السعر المخفّض بدعم نون. لو العمود مش موجود يرجع لـ active_price."""
    v = fnum(r.get("sale_price"))
    return v if v != "" else fnum(r.get("active_price"))


def build_catalog_map(rows):
    """مفتاح الكود (partner_sku أو sku_child) → (partner_sku, السعر، الفاميلي)."""
    m = {}
    for r in rows or []:
        ps = (r.get("partner_sku") or "").strip()
        info = (ps, our_price(r), (r.get("family_code") or "").strip())
        for k in (ps, (r.get("sku_child") or "").strip()):
            if k:
                m.setdefault(k, info)
    return m


# ---------------- سجل الأسعار اليومي (لقطة سعر) ----------------
# السعر = sale_price (سعرنا، مش خصم نون). بنسجّله بس لما يتغير: (متجر، SKU) → قايمة مرتبة [(تاريخ، سعر)]
# سعر أي طلب = آخر سعر مسجّل وقت الطلب أو قبله (بالتاريخ والساعة)؛ لو مفيش تسجيل قبل الطلب → السعر الحالي.
def read_history(log):
    try:
        vals = read_tab_values(HIST_TAB)
    except Exception as e:
        log(f"ℹ️ {HIST_TAB} لسه مش موجود أو تعذرت قراءته ({str(e)[:80]}) — هبدأ سجل جديد")
        return {}
    hist = {}
    for r in vals[1:]:
        if len(r) < 4:
            continue
        d, store, sku, pr = r[0].strip(), r[1].strip(), r[2].strip(), fnum(r[3])
        if d and sku and pr != "":
            hist.setdefault((store, sku), []).append((d, pr))
    for v in hist.values():
        v.sort()
    log(f"🕘 {HIST_TAB}: {sum(len(v) for v in hist.values())} تسجيل سعر لـ {len(hist)} SKU")
    return hist


def snapshot_prices(hist, store_name, cat_rows, stamp, stats=None):
    """يسجّل السعر (مع التاريخ والساعة) لو اتغير عن آخر تسجيل. بيرجّع عدد التسجيلات الجديدة.
    stats (اختياري): dict بيتملي بإحصائيات التشخيص: صفوف الكتالوج / بسعر / من غير سعر / مكرر / من غير كود."""
    n = 0
    seen = set()
    st_ = stats if stats is not None else {}
    st_.update(rows=len(cat_rows or []), priced=0, noprice=0, nokey=0, dup=0, unchanged=0)
    for r in cat_rows or []:
        # المفتاح = sku_child (نفس كود الـ SKU اللي في الأوردرات)؛ لو مفيش نرجع لـ partner_sku
        key = (r.get("sku_child") or "").strip() or (r.get("partner_sku") or "").strip()
        price = our_price(r)
        if not key:
            st_["nokey"] += 1
            continue
        if price == "":
            st_["noprice"] += 1
            continue
        if key in seen:
            st_["dup"] += 1
            continue
        seen.add(key)
        st_["priced"] += 1
        lst = hist.setdefault((store_name, key), [])
        if lst and lst[-1][1] == price:
            st_["unchanged"] += 1
            continue
        lst.append((stamp, price))
        n += 1
    return n


def log_snapshot(log, store_name, stats, n):
    log(f"🕘 [{store_name}] كتالوج {stats['rows']} صف → بسعر: {stats['priced']} | من غير سعر: {stats['noprice']} | "
        f"من غير كود: {stats['nokey']} | مكرر: {stats['dup']} | ثابت: {stats['unchanged']} | اتسجّل: {n}")


def hist_price(hist, store_name, sku, ts):
    """آخر سعر مسجّل وقت ts (YYYY-MM-DD HH:MM:SS) أو قبله، أو None لو مفيش.
    تسجيلات قديمة بتاريخ بس (YYYY-MM-DD) بتتعامل كأنها من أول اليوم."""
    lst = hist.get((store_name, sku)) if hist else None
    if not lst:
        return None
    best = None
    for d, p in lst:
        if d <= ts:
            best = p
        else:
            break
    return best


def history_rows(hist):
    return [[d, st, sku, p] for (st, sku), lst in sorted(hist.items()) for d, p in lst]


def utc_to_local(s, hours):
    s = (s or "").replace(" UTC", "").strip()
    try:
        d = dt.datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
    except Exception:
        return s, s
    return (d + dt.timedelta(hours=hours)).strftime("%Y-%m-%d %H:%M:%S"), s


def build_rows(store_name, orders, cat, tz_hours, uploaded, hist=None):
    out, skipped, noprice, fromhist = [], 0, 0, 0
    for o in orders or []:
        stt = (o.get("item_status") or "").strip().lower()
        if stt in EXCLUDE_STATUS:
            skipped += 1
            continue
        raw = (o.get("sku") or "").strip()
        base = raw[:-2] if raw.endswith("-1") else raw
        ps, price, fam = cat.get(raw) or cat.get(base) or (base, "", "")
        local, utc = utc_to_local(o.get("order_placed_at"), tz_hours)
        hp = None
        if hist:
            for k in (raw, base, ps):      # كود الأوردر الأول (= sku_child)، وبعدين البدائل
                hp = hist_price(hist, store_name, k, local)
                if hp is not None:
                    break
        if hp is not None:
            price = hp
            fromhist += 1
        if price == "":
            noprice += 1
        fbn = (o.get("is_fulfilled_by_noon") or "").strip().lower() == "true"
        out.append([raw, local, o.get("item_status", ""), price, 1,
                    "Fulfilled by Noon (FBN)" if fbn else "Fulfilled by Partner (FBP)",
                    uploaded, fam, store_name, utc, o.get("order_nr", ""), o.get("item_nr", ""),
                    o.get("warehouse_code", ""), ps])
    build_rows.last_hist = fromhist
    return out, skipped, noprice


def creds_path():
    ms = glob.glob(os.path.join(GSHEET_CREDS_DIR, "*.json"))
    return ms[0] if ms else None


# ═══════════════ PostgreSQL (Neon) — بديل الكتابة/القراءة من Google Sheet (اختياري بالـ checkbox) ═══════════════
PG_MODE = "sheet"                                        # sheet | pg | both
PG_LABELS = {"sheet": "📄 Google Sheet فقط", "pg": "🐘 PostgreSQL فقط", "both": "📄+🐘 الاتنين مع بعض"}
PG_URL_FILE = os.path.join(APP_DIR, "pg_url.txt")        # رابط Neon (بيتحفظ مرة واحدة)
PG_FLAG_FILE = os.path.join(APP_DIR, "noon_use_pg.txt")  # sheet / pg / both
PG_SCHEMA = "sheet_copy"


def pg_flag_load():
    try:
        v = open(PG_FLAG_FILE, encoding="utf-8").read().strip()
        return {"1": "pg"}.get(v, v) if {"1": "pg"}.get(v, v) in PG_LABELS else "sheet"
    except Exception:
        return "sheet"


def pg_flag_save(mode):
    try:
        open(PG_FLAG_FILE, "w", encoding="utf-8").write(mode)
    except Exception:
        pass


def pg_url():
    import re as _re
    u = os.environ.get("PG_URL", "").strip()
    if not u:
        try:
            u = open(PG_URL_FILE, encoding="utf-8").read().strip()
        except Exception:
            u = ""
    u = _re.sub(r"[&?]channel_binding=[^&]*", "", u.strip().strip('"'))
    if u and "?" not in u and "&" in u:
        u = u.replace("&", "?", 1)
    return u


def pg_ask_url():
    """بيسأل عن رابط Neon (الكتابة مخفية) ويحفظه في pg_url.txt جنب البرنامج."""
    from tkinter import simpledialog
    u = simpledialog.askstring("رابط Neon", "الصق رابط الاتصال (postgresql://...) — الكتابة مخفية:", show="*")
    if u and u.strip().startswith("postgres"):
        open(PG_URL_FILE, "w", encoding="utf-8").write(u.strip())
        return True
    return False


def pg_conn():
    import psycopg2, re as _re
    u = pg_url()
    if not u:
        raise RuntimeError("مفيش رابط Neon — اضغط زر 🔑 رابط Neon")
    m = _re.search(r"postgres(?:ql)?://([^:@/\s]+):([^@\s]+)@([^:/?&\s'\"]+)(?::(\d+))?/([A-Za-z0-9_\-]+)", u)
    if not m:
        raise RuntimeError("رابط Neon مش مفهوم — لازم يبدأ بـ postgresql://user:password@host/dbname")
    return psycopg2.connect(user=m.group(1), password=m.group(2), host=m.group(3), port=int(m.group(4) or 5432),
                            dbname=_re.sub(r"(sslmode|channel_binding)$", "", m.group(5)) or "neondb", sslmode="require", connect_timeout=20)


def pg_toggle(var, log):
    """بيتنادى لما اختيار مكان الكتابة يتغيّر (var = نص الاختيار): بيضبط الوضع وبيختبر الاتصال."""
    global PG_MODE
    mode = {v: k for k, v in PG_LABELS.items()}.get(var.get(), "sheet")
    if mode == PG_MODE:
        return
    if mode != "sheet":
        try:
            import psycopg2  # noqa
        except Exception:
            log("⚠️ psycopg2 مش متسطّب — نفّذ: py -m pip install psycopg2-binary")
            mode = "sheet"
        if mode != "sheet" and not pg_url() and not pg_ask_url():
            mode = "sheet"
        if mode != "sheet":
            try:
                c = pg_conn()
                c.cursor().execute("SELECT 1")
                c.close()
            except Exception as e:
                log(f"⚠️ فشل الاتصال بـ Neon: {str(e)[:150]}")
                mode = "sheet"
    PG_MODE = mode
    pg_flag_save(mode)
    if var.get() != PG_LABELS[mode]:
        var.set(PG_LABELS[mode])
    log(f"🐘 كل البوتات هتكتب في: {PG_LABELS[mode]}" + (" (وبيقرا من PostgreSQL)" if mode == "pg" else ""))


def pg_side(log, fn):
    """بينفّذ كتابة PostgreSQL حسب الوضع. بيرجع True لو الشيت يتخطّى (وضع PostgreSQL فقط)."""
    if PG_MODE == "sheet":
        return False
    try:
        fn()
    except Exception as e:
        if PG_MODE == "pg":
            raise
        log(f"⚠️ PostgreSQL فشل: {str(e)[:150]} — كمّلت على الشيت")
    return PG_MODE == "pg"


def _pg_snake(name, used):
    import re as _re
    s = _re.sub(r"[^0-9a-zA-Z]+", "_", str(name)).strip("_").lower()
    if not s:
        s = f"c_{len(used) + 1}"
    if s[0].isdigit():
        s = "c_" + s
    s = s[:60]
    base, i = s, 2
    while s in used:
        s, i = f"{base}_{i}", i + 1
    used.add(s)
    return s


def _pg_s(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def pg_read_tab(tab):
    """نفس شكل get_all_values (الصف الأول هيدر) من sheet_copy، أو None لو التاب مش موجود."""
    c = pg_conn()
    try:
        cur = c.cursor()
        cur.execute("SELECT to_regclass(%s)", (f"{PG_SCHEMA}._meta",))
        if cur.fetchone()[0] is None:
            return None
        cur.execute(f"SELECT table_name, headers, col_names FROM {PG_SCHEMA}._meta WHERE tab_title=%s", (tab,))
        m = cur.fetchone()
        if not m:
            return None
        table, hdr, cols = m
        cur.execute("SELECT " + ", ".join(f'"{x}"' for x in cols) + f' FROM {PG_SCHEMA}."{table}" ORDER BY _row')
        return [list(hdr)] + [["" if v is None else v for v in r] for r in cur.fetchall()]
    finally:
        c.close()


def read_tab_values(tab):
    """قراءة تاب: من PostgreSQL لو الخيار مفعّل (ولو التاب مش هناك → من الشيت)، وإلا من الشيت."""
    if PG_MODE == "pg":
        try:
            v = pg_read_tab(tab)
            if v is not None:
                return v
        except Exception:
            pass
    return open_ss().worksheet(tab).get_all_values()


def pg_write_tab(tab, headers, rows, log, keys=None, orphan_set=None, scope=None, since=None):
    """بيكتب التاب في PostgreSQL (schema sheet_copy — نفس شكل PG_Copy_All عشان التطبيق يقراه).
    keys → بناء وتحديث بالمفتاح (القديم بيفضل). من غيره → استبدال كامل. كله في transaction واحدة."""
    import io as _io, csv as _csv
    headers = [str(h).strip() for h in headers]
    used = set()
    cols = [_pg_snake(h, used) for h in headers]
    table = _pg_snake(tab, set())
    c = pg_conn()
    try:
        cur = c.cursor()
        cur.execute(f"CREATE SCHEMA IF NOT EXISTS {PG_SCHEMA}")
        cur.execute(f"CREATE TABLE IF NOT EXISTS {PG_SCHEMA}._meta (table_name TEXT PRIMARY KEY, tab_title TEXT, headers JSONB, col_names JSONB, copied_at TIMESTAMPTZ DEFAULT now())")
        cur.execute(f"SELECT headers, col_names FROM {PG_SCHEMA}._meta WHERE table_name=%s", (table,))
        m = cur.fetchone()
        old = []                                   # الصفوف القديمة كـ dict بالهيدر
        if m:
            ohdr, ocols = m
            cur.execute("SELECT " + ", ".join(f'"{x}"' for x in ocols) + f' FROM {PG_SCHEMA}."{table}" ORDER BY _row')
            old = [dict(zip(ohdr, ["" if v is None else v for v in r])) for r in cur.fetchall()]
        if not rows and m:
            log(f"⚠️ {tab}: مفيش بيانات — القاعدة ما اتلمستش")
            return
        new = [{h: _pg_s(v) for h, v in zip(headers, r)} for r in rows]
        if keys and old:
            kf = lambda d: "|".join(str(d.get(k, "")).strip() for k in keys)
            idx = {}
            for i, d in enumerate(old):
                k = kf(d)
                if k.strip("|") and k not in idx:
                    idx[k] = i
            final = [dict(d) for d in old]
            adds = ups = same = orph = 0
            seen = set()
            for d in new:
                k = kf(d)
                if k in idx:
                    seen.add(k)
                    cur_d = final[idx[k]]
                    if any(str(cur_d.get(h, "")).strip() != d[h].strip() for h in headers):
                        cur_d.update(d)
                        ups += 1
                    else:
                        same += 1
                else:
                    final.append(d)
                    adds += 1
            if orphan_set and scope:
                sc_col, sc_vals = scope
                for k, i in idx.items():
                    d = final[i]
                    if k in seen or str(d.get(sc_col, "")).strip() not in sc_vals:
                        continue
                    if since and str(d.get(since[0], "")).strip()[:10] < since[1]:
                        continue
                    if any(str(d.get(cn, "")).strip().lower() != str(x).strip().lower() for cn, x in orphan_set.items()):
                        d.update(orphan_set)
                        orph += 1
            body = final
            msg = f"{adds} جديد · {ups} اتحدّث" + (f" · {orph} اتعلّموا" if orph else "") + f" · {same} زي ما هو"
        else:
            body = new
            msg = f"اتبدّل بالكامل"
        cur.execute(f'DROP TABLE IF EXISTS {PG_SCHEMA}."{table}"')
        cur.execute(f'CREATE TABLE {PG_SCHEMA}."{table}" (' + ", ".join(f'"{x}" TEXT' for x in cols) + ", _row BIGSERIAL)")
        cur.execute(f"INSERT INTO {PG_SCHEMA}._meta (table_name, tab_title, headers, col_names, copied_at) VALUES (%s,%s,%s,%s,now()) "
                    f"ON CONFLICT (table_name) DO UPDATE SET tab_title=EXCLUDED.tab_title, headers=EXCLUDED.headers, col_names=EXCLUDED.col_names, copied_at=now()",
                    (table, tab, json.dumps(headers, ensure_ascii=False), json.dumps(cols)))
        if body:
            buf = _io.StringIO()
            w = _csv.writer(buf)
            for d in body:
                w.writerow([d.get(h, "") for h in headers])
            buf.seek(0)
            cur.copy_expert(f'COPY {PG_SCHEMA}."{table}" (' + ", ".join(f'"{x}"' for x in cols) + ") FROM STDIN WITH CSV", buf)
        c.commit()
        log(f"🐘 {tab}: {msg} — الإجمالي في PostgreSQL {len(body)}")
    except Exception:
        c.rollback()
        raise
    finally:
        c.close()


STORAGE_TAB = "Noon_Storage_API"
STORAGE_REPORT = "fbn_finance_monthlystoragechargereport"
STORAGE_HEADERS = ["Store", "soa_reference_nr", "description_en", "id_partner", "service_month", "pbarcode", "sku", "year_month",
                   "warehouse_code", "country_code", "average_stock_on_hand", "average_pending_qty", "average_net_qty", "longest_side",
                   "median_side", "shortest_side", "weight", "volumetric_weight", "item_volume", "estimated_total_item_volume",
                   "cost_per_cubic_feet", "charged_amount"]


def storage_months(n, tz_hours=3.0):
    """آخر n شهور مكتملة (من الشهر اللي فات رجوعاً) بصيغة YYYY-MM."""
    d = dt.datetime.utcnow() + dt.timedelta(hours=float(tz_hours))
    y, m = d.year, d.month
    out = []
    for _ in range(max(1, int(n))):
        m -= 1
        if m == 0:
            y, m = y - 1, 12
        out.append(f"{y}-{m:02d}")
    return out


def fetch_storage(st, months, country, log, stop):
    """رسوم التخزين الشهرية (FBN Monthly Storage Charge) لكل شهر في months → قايمة صفوف dict (مع اسم المتجر)."""
    import re as _re
    pid = _re.sub(r"\D", "", str(st.key.get("project_code", "")))
    out = []
    for mth in months:
        if stop.is_set():
            break
        rows = run_export(st, STORAGE_REPORT, {"id_partner": int(pid) if pid else "", "country": country, "service_month": mth},
                          log, stop)
        if rows is None:
            log(f"[{st.name}] ⚠️ تقرير التخزين {mth}: مارجعش (يمكن لسه مااتولّدش لشهر {mth})")
            continue
        for r in rows:
            r = dict(r)
            r["Store"] = st.name
            out.append(r)
        log(f"[{st.name}] 🏬 تخزين {mth}: {len(rows)} صف")
    return out


def write_storage(data, months, log):
    """بيبدّل صفوف (المتجر + الشهر) اللي اتسحبت، وبيسيب باقي الشهور القديمة زي ما هي."""
    fetched = {(nm, mth) for nm, rs in data.items() for mth in months}
    new = [[str(r.get(h, "")) for h in STORAGE_HEADERS] for rs in data.values() for r in rs]
    kept = []
    try:
        old = read_tab_values(STORAGE_TAB)
    except Exception:
        old = []
    if len(old) > 1:
        oh = [str(x).strip() for x in old[0]]
        if "Store" in oh and "service_month" in oh:
            si, mi = oh.index("Store"), oh.index("service_month")
            for r in old[1:]:
                r = list(r) + [""] * (len(oh) - len(r))
                if (str(r[si]).strip(), str(r[mi]).strip()) in fetched:
                    continue
                kept.append([str(r[oh.index(h)]) if h in oh else "" for h in STORAGE_HEADERS])
    allrows = kept + new
    if not allrows:
        log("⚠️ تخزين: مفيش صفوف تتكتب")
        return
    write_raw_tab(STORAGE_TAB, STORAGE_HEADERS, allrows, log)
    log(f"✅ رسوم التخزين اتكتبت في تاب {STORAGE_TAB} ({len(new)} جديد · {len(kept)} قديم محفوظ)")


def write_sheet(rows, log):
    """DailyOrders_API: بناء وتحديث بـ (order_nr + item_nr) — الأوردرات القديمة برا نطاق الأيام بتفضل، والحالة/السعر بيتحدّثوا،
    والأوردر اللي اتلغى جوه النطاق بيتعلّم cancelled بدل ما يفضل مبيعات."""
    ix = {h: i for i, h in enumerate(HEADERS)}
    _st = {str(r[ix["Store"]]).strip() for r in rows}
    _dy = sorted(str(r[ix["Order Timestamp"]]).strip()[:10] for r in rows if str(r[ix["Order Timestamp"]]).strip())
    if pg_side(log, lambda: pg_write_tab(TEST_TAB, HEADERS, rows, log, keys=("order_nr", "item_nr", "SKU"), orphan_set={"Status": "cancelled"},
                                         scope=("Store", _st), since=("Order Timestamp", _dy[0]) if _dy else None)):
        return
    cp = creds_path()
    if not cp:
        raise RuntimeError(f"مفيش ملف اعتماد Google في {GSHEET_CREDS_DIR}")
    ss = gspread.authorize(ServiceAccountCredentials.from_json_keyfile_name(cp, GSHEET_SCOPE)).open(GSHEET_NAME)
    stores = {str(r[ix["Store"]]).strip() for r in rows}
    days = sorted(str(r[ix["Order Timestamp"]]).strip()[:10] for r in rows if str(r[ix["Order Timestamp"]]).strip())
    sheet_upsert(ss, TEST_TAB, HEADERS, rows, ("order_nr", "item_nr", "SKU"), log, vio="USER_ENTERED",
                 cmp_cols=["Status", "Price", "Quantity", "Fulfillment Model", "family", "warehouse_code", "partner_sku"],
                 orphan_set={"Status": "cancelled"}, scope=("Store", stores), since=("Order Timestamp", days[0]) if days else None)


def open_ss():
    cp = creds_path()
    if not cp:
        raise RuntimeError(f"مفيش ملف اعتماد Google في {GSHEET_CREDS_DIR}")
    return gspread.authorize(ServiceAccountCredentials.from_json_keyfile_name(cp, GSHEET_SCOPE)).open(GSHEET_NAME)


def live_catalog(log, tab="LIVE"):
    """الأسعار والـ family من تاب LIVE (اللي بوت نون بيملاه) — من غير ما نستنى تصدير من نون."""
    try:
        vals = open_ss().worksheet(tab).get_all_values()
    except Exception as e:
        log(f"⚠️ تعذرت قراءة {tab}: {str(e)[:150]}")
        return []
    if len(vals) < 2:
        return []
    h = [x.strip() for x in vals[0]]
    need = [c for c in ("partner_sku",) if c not in h]
    if "sale_price" not in h and "active_price" not in h:
        need.append("sale_price")
    if need:
        log(f"⚠️ {tab} ناقصه أعمدة: {need} — الأعمدة الموجودة: {h[:12]}")
        return []
    rows = [dict(zip(h, r)) for r in vals[1:]]
    log(f"🗂️ {tab}: {len(rows)} صف (أسعار وfamily)")
    return rows


import re


def _urls(obj, path="", out=None):
    out = [] if out is None else out
    if isinstance(obj, dict):
        for k, v in obj.items():
            _urls(v, f"{path}.{k}", out)
    elif isinstance(obj, list):
        for v in obj:
            _urls(v, path, out)
    elif isinstance(obj, str) and obj.startswith("http"):
        out.append((path.lower(), obj))
    return out


def _strs(obj, path="", out=None):
    out = [] if out is None else out
    if isinstance(obj, dict):
        for k, v in obj.items():
            _strs(v, f"{path}.{k}", out)
    elif isinstance(obj, list):
        for v in obj:
            _strs(v, path, out)
    elif isinstance(obj, str):
        out.append((path.lower(), obj.strip()))
    return out


def pick_image(data):
    """أنسب رابط صورة في رد نون (نفس طريقة AwbPc / noon_api_explorer)."""
    best, best_score = "", -1
    for p, url in _urls(data):
        score = (2 if re.search(r"\.(jpg|jpeg|png|webp|gif)(\?|$)", url, re.I) else 0) + \
                (2 if any(w in p for w in ("image", "img", "media", "picture")) else 0) + \
                (1 if "nooncdn" in url else 0) - (1 if "thumb" in p else 0)
        if score >= 2 and score > best_score:
            best, best_score = url, score
    if best:
        return best
    # أحيانًا نون بيرجّع "مفتاح" الصورة من غير رابط (v1690876543/N12519655A_1) — بنبني الرابط منه
    for p, v in _strs(data):
        if any(w in p for w in ("image", "img", "media")) and re.match(r"^(v\d+/|pzsku/)[\w/.-]+$", v):
            return f"https://f.nooncdn.com/p/{v}" + ("" if re.search(r"\.(jpe?g|png|webp)$", v, re.I) else ".jpg")
    return ""


def pick_title(data):
    found = {}

    def walk(o):
        if isinstance(o, dict):
            for k, v in o.items():
                kl = str(k).lower()
                if isinstance(v, str) and len(v.strip()) > 3 and not v.startswith("http") and \
                        kl in ("title_en", "product_title_en", "name_en", "title", "product_title", "product_name", "name"):
                    found.setdefault(kl, v.strip())
                else:
                    walk(v)
        elif isinstance(o, list):
            for v in o:
                walk(v)
    walk(data)
    for k in ("title_en", "product_title_en", "name_en", "title", "product_title", "product_name", "name"):
        if k in found:
            return found[k][:200]
    return ""


IMG_ALIAS = {}                 # SKU (partner_sku مثل N123..A) → sku_child (Z…-1): content/get بيقبل الـ Z (من غير -1) بس
_IMG_STAT = {}                 # (شكل الكود، المصدر) → عدد النجاحات: بنرتّب المحاولات بالأنجح عشان نقلّل الطلبات الفاشلة
_IMG_MEMO = {}                 # كود الأب → نتيجة (الأولاد -1/-2 بتشارك نفس الصورة فمنسألش تاني)
_IMG_MEMO_LOCK = threading.Lock()
_IMG_KEY_LOCKS = {}


def _parent_of(sku):
    return sku.rsplit("-", 1)[0] if re.match(r"^.+-\d+$", sku) else sku


_REQ_STAT = {"n": 0, "t": 0.0, "429": 0, "err": 0}


def _par_key_for_code(code):
    return _parent_of(code)


def _try_one(st, src, code, log):
    t_ = time.time()
    try:
        return _try_one_raw(st, src, code, log)
    finally:
        _REQ_STAT["n"] += 1
        _REQ_STAT["t"] += time.time() - t_


def _try_one_raw(st, src, code, log):
    path, body = (("/content/v1/product/content/get", {"sku_parent": code}) if src == "content"
                  else (f"/offer/v1/product/{code}", None))
    if body is None:
        with st.lock:
            if st.sess is None:
                st.sess = login(st, log)
            s_ = st.sess
        r = s_.get(BASE + path, headers={"Accept": "application/json"}, timeout=15) if s_ else None
    else:
        r = api_post(st, path, body, log)
    if r is None or r.status_code != 200:
        if r is not None and r.status_code == 429:
            _REQ_STAT["429"] += 1
            time.sleep(3)
        return "", "", f"{st.name}/{src}/{code}: {getattr(r, 'status_code', 'no-login')}"
    data = r.json()
    meta = pick_meta(data)
    if meta:
        IMG_META.setdefault(_par_key_for_code(code), {}).update({k: v for k, v in meta.items() if v})
    return pick_image(data), (meta.get("Title") or pick_title(data)), f"{st.name}/{src}/{code}: no image"


def attr_pairs(data):
    """كل أزواج (كود الخاصية → قيمتها النصية) من رد نون، سواء dict عادي أو list فيها {code/key/name, value}."""
    pairs = {}

    def vs(v):
        if isinstance(v, str):
            return v.strip()
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return str(v)
        if isinstance(v, dict):
            for k in ("en", "value", "name", "label"):
                if isinstance(v.get(k), str) and v[k].strip():
                    return v[k].strip()
            for x in v.values():
                if isinstance(x, str) and x.strip():
                    return x.strip()
        if isinstance(v, list) and v:
            return vs(v[0])
        return ""

    def walk(o):
        if isinstance(o, dict):
            code = next((o[k] for k in ("code", "key", "attribute_code", "name") if isinstance(o.get(k), str) and o.get(k)), None)
            vk = next((k for k in ("value", "values", "value_en", "text") if k in o), None)
            if code and vk:
                v = vs(o[vk])
                if v:
                    pairs.setdefault(code.lower(), v)
            for k, v in o.items():
                if isinstance(v, (str, int, float)) and not isinstance(v, bool):
                    t = vs(v)
                    if t:
                        pairs.setdefault(str(k).lower(), t)
                else:
                    if isinstance(v, dict) and v and all(isinstance(x, (str, int, float)) for x in v.values()):
                        t = vs(v)          # {"en": "...", "ar": "..."} → القيمة الإنجليزية تحت اسم المفتاح
                        if t:
                            pairs.setdefault(str(k).lower(), t)
                    walk(v)
        elif isinstance(o, list):
            for x in o:
                walk(x)
    walk(data)
    return pairs


_META_KEYS = {
    "Title": ("title_en", "product_title_en", "title", "product_title", "name_en", "product_name", "name"),
    "Family": ("family", "family_code", "product_family", "family_en"),
    "Product Type": ("product_type", "product_type_code", "product_subtype", "product_sub_type", "type"),
    "Brand": ("brand", "brand_en", "brand_code", "brand_name"),
    "Category": ("category", "category_code", "category_en", "catalog_category", "product_category", "department", "department_code"),
}


def pick_meta(data):
    pr = attr_pairs(data)
    out = {}
    for col, keys in _META_KEYS.items():
        for k in keys:
            if pr.get(k) and len(pr[k]) > (3 if col == "Title" else 0):
                out[col] = pr[k][:200]
                break
    return out


IMG_FAMILY = {}                # SKU → family من تاب المبيعات (احتياطي لو رد نون ما فيهوش فاميلي)
IMG_META = {}                  # كود الأب → {Title, Family, Product Type, Brand, Category}


def _par_key(sku):
    al = IMG_ALIAS.get(sku, "")
    return _parent_of(al) if al else _parent_of(sku)


def find_image(stores, sku, log, pref=None):
    """أسرع: (1) الأولاد بتوع نفس الأب بيشاركوا نتيجة واحدة، (2) المحاولات بتترتّب بالأنجح (كود كامل/أب × content/offer)،
    (3) المتجر اللي الـ SKU جه منه الأول، وباقي المتاجر بأنجح محاولة بس. بيرجّع (رابط، اسم، ملاحظة)."""
    al = IMG_ALIAS.get(sku, "")
    par = _parent_of(al) if al else _parent_of(sku)
    with _IMG_MEMO_LOCK:
        if par in _IMG_MEMO:
            u, t, n = _IMG_MEMO[par]
            if u:
                return u, t, n + " [نفس الأب]"
        lk = _IMG_KEY_LOCKS.setdefault(par, threading.Lock())
    with lk:                                   # أخوات نفس الأب بيستنوا بعض بدل ما يكرّروا الطلبات
        with _IMG_MEMO_LOCK:
            if par in _IMG_MEMO and _IMG_MEMO[par][0]:
                u, t, n = _IMG_MEMO[par]
                return u, t, n + " [نفس الأب]"
        forms = ([("alias_parent", _parent_of(al)), ("alias_full", al)] if al else []) + [("full", sku)] + \
                ([("parent", _parent_of(sku))] if _parent_of(sku) != sku else [])
        forms = [(f, c) for i, (f, c) in enumerate(forms) if (f, c) not in forms[:i]]
        pairs = [(f, c, src) for f, c in forms for src in ("content", "offer")]
        pairs.sort(key=lambda x: -_IMG_STAT.get((x[0], x[2]), 0))
        first = [pref] if pref is not None else list(stores)
        rest = [x for x in stores if x is not pref] if pref is not None else []
        plan = [(st, f, c, src) for st in first for f, c, src in pairs] + \
               [(st, f, c, src) for st in rest for f, c, src in pairs[:1]]
        title, errs = "", []
        for st, f, c, src in plan:
            try:
                url, t, err = _try_one(st, src, c, log)
            except Exception as e:
                errs.append(f"{st.name}/{src}/{c}: {type(e).__name__}")
                continue
            title = title or t
            if url:
                _IMG_STAT[(f, src)] = _IMG_STAT.get((f, src), 0) + 1
                out = (url, title, f"ok ({st.name}/{src}/{c})")
                with _IMG_MEMO_LOCK:
                    _IMG_MEMO[par] = out
                return out
            errs.append(err)
        return "", title, "not found — " + " | ".join(errs[:2] + errs[-2:] if len(errs) > 4 else errs)


def read_image_cache(log):
    out = {}
    try:
        for r in read_tab_values(IMG_TAB)[1:]:
            r = r + [""] * (len(IMG_HEADERS) - len(r))
            if r[0].strip():
                out[r[0].strip()] = dict(zip(IMG_HEADERS, r))
    except Exception:
        pass
    return out


def update_images(stores, skus, log, write=True, pref=None, alias=None):
    """يكمّل كاش Images_API: SKU جديد (أو من غير صورة ومتجرّبش النهارده) → نسأل نون API. بيرجّع الكاش."""
    if alias:
        IMG_ALIAS.update(alias)
    cache = read_image_cache(log)
    if write and not cache:
        try:   # نعمل التاب فورًا بالهيدر (حتى لو فاضي) عشان يظهر في الشيت
            write_raw_tab(IMG_TAB, IMG_HEADERS, [], log)
            log(f"🆕 اتعمل تاب {IMG_TAB}")
        except Exception as e:
            log(f"⚠️ تعذر إنشاء تاب {IMG_TAB}: {str(e)[:150]}")
    today = dt.date.today().isoformat()
    todo = [k for k in sorted(skus) if k not in cache or
            (not cache[k].get("Image URL") and cache[k].get("Checked At", "")[:10] != today)]
    if len(todo) > IMG_MAX_PER_RUN:
        log(f"🖼️ {len(todo)} SKU محتاجين صور — هاخد {IMG_MAX_PER_RUN} دلوقتي والباقي في التشغيلة الجاية")
        todo = todo[:IMG_MAX_PER_RUN]
    if todo:
        log(f"🖼️ باجيب صور {len(todo)} SKU من نون API (ميزانية وقت {IMG_TIME_BUDGET // 60} دقايق — اللي ما يخلصش يكمل في التشغيلة الجاية)...")
        t_img, done_n, done_keys = time.time(), 0, set()
        lock_img = threading.Lock()

        def one_img(k):
            nonlocal done_n
            if time.time() - t_img > IMG_TIME_BUDGET:
                return None
            res_ = find_image(stores, k, log, (pref or {}).get(k))
            url, title, note = res_
            with lock_img:
                mt = IMG_META.get(_par_key(k), {})
                old_ = cache.get(k, {})
                cache[k] = {"SKU": k, "Image URL": url or old_.get("Image URL", ""),
                            "Title": title or mt.get("Title") or old_.get("Title", ""), "Status": note[:150],
                            "Checked At": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Family": mt.get("Family") or IMG_FAMILY.get(k, "") or old_.get("Family", ""),
                            "Product Type": mt.get("Product Type") or old_.get("Product Type", ""),
                            "Brand": mt.get("Brand") or old_.get("Brand", ""),
                            "Category": mt.get("Category") or old_.get("Category", "")}
                done_n += 1
                done_keys.add(k)
                if done_n <= 8:
                    log(f"   🔬 {k}: {note[:170]}")
                if done_n % 20 == 0 or done_n == len(todo):
                    n_, t_ = _REQ_STAT["n"], _REQ_STAT["t"]
                    log(f"🖼️ {done_n}/{len(todo)} SKU اتفحصوا ({int(time.time() - t_img)} ث) — طلبات: {n_}، متوسط الطلب: "
                        f"{(t_ / n_ if n_ else 0):.1f} ث، 429: {_REQ_STAT['429']}")
                if write and done_n % 100 == 0 and done_n < len(todo):
                    try:
                        write_raw_tab(IMG_TAB, IMG_HEADERS, [[c.get(h, "") for h in IMG_HEADERS] for c in cache.values()], log)
                        log(f"💾 حفظ دوري: {len(cache)} صف في {IMG_TAB}")
                    except Exception as e:
                        log(f"⚠️ الحفظ الدوري: {str(e)[:100]}")
            return res_

        with ThreadPoolExecutor(max_workers=32) as ex:
            list(ex.map(one_img, todo))
        todo = [k for k in todo if k in done_keys]
        got = sum(1 for k in todo if cache[k].get("Image URL"))
        log(f"🖼️ اتلاقت صور {got} من {len(todo)}")
        if write:
            write_raw_tab(IMG_TAB, IMG_HEADERS, [[c.get(h, "") for h in IMG_HEADERS] for c in cache.values()], log)
            log(f"✅ كاش الصور اتكتب في تاب {IMG_TAB}")
    else:
        log("🖼️ كل الـ SKUs ليها صور في الكاش (أو اتجرّبت النهارده)")
    return cache


def links_map(log):
    """SKU → رابط صورة: Images_API (من نون API) بيغلّب، و links n احتياطي بس."""
    m = {}
    try:
        ss = open_ss()
    except Exception as e:
        log(f"⚠️ تعذر فتح الشيت للصور: {str(e)[:100]}")
        return m
    for tab, col in (("links n", 1), (IMG_TAB, 1)):
        try:
            for r in read_tab_values(tab)[1:]:
                if len(r) > col and r[0].strip() and r[col].strip().startswith("http"):
                    m[r[0].strip().upper()] = r[col].strip()
        except Exception as e:
            log(f"⚠️ تعذرت قراءة {tab}: {str(e)[:100]}")
    return m


def tag_finance(fin_data, orders_by_store, log):
    """يضيف لكل صف مالي: Fulfillment (FBN/FBB) + SKU من الأوردر — بالربط على Order Nr مع تقرير الأوردرات."""
    def num(x):
        try:
            return float(str(x).replace(",", "").strip() or 0)
        except Exception:
            return 0.0
    fh = ["Store"]
    for rs in fin_data.values():
        if rs:
            for h in rs[0].keys():
                if h not in fh:
                    fh.append(h)
    base_n = len(fh) - 1
    fh += ["Fulfillment (FBN/FBB)", "SKU من الأوردر"]
    rows, cnt, sums = [], {}, {}
    for nm, rs in fin_data.items():
        m = {}
        for r in (orders_by_store.get(nm) or []):
            e = m.setdefault(str(r.get("order_nr", "")).strip(), {"fbn": set(), "skus": []})
            e["fbn"].add(str(r.get("is_fulfilled_by_noon", "")).strip().lower() == "true")
            sk = str(r.get("sku", "")).strip()
            if sk and sk not in e["skus"]:
                e["skus"].append(sk)
        for r in rs:
            if str(r.get("Transaction Type", "")).strip().lower() == "payment":
                continue
            on = str(r.get("Order Nr", "")).strip()
            e = m.get(on)
            if on in ("", "NA"):
                model, skus = "—", ""
            elif not e:
                model, skus = "غير معروف", ""
            else:
                model = "FBN" if e["fbn"] == {True} else ("FBB" if e["fbn"] == {False} else "مختلط")
                skus = ", ".join(e["skus"])
            cnt[model] = cnt.get(model, 0) + 1
            a = sums.setdefault(model, [0.0, 0.0])
            a[0] += num(r.get("Referral Fee including VAT"))
            a[1] += num(r.get("Fullfilment & Logistics Fees including VAT"))
            rows.append([nm] + [r.get(h, "") for h in fh[1:1 + base_n]] + [model, skus])
    log(f"🔗 الربط بالأوردرات: {cnt}")
    for k, (a, b) in sums.items():
        log(f"   {k}: عمولة {a:,.2f} | توصيل/تنفيذ {b:,.2f}")
    return fh, rows


def build_com(fin_data, orders_by_store, log, uploaded):
    """COM_API: لكل كود وكل نوع تنفيذ، عمولة وتوصيل **آخر عملية بيع مكتملة** (بتاريخ الأوردر) —
    مكتملة = فيها عمولة وتوصيل الاتنين. الأرقام شاملة الضريبة (زي سعر البيع في التطبيق).
    العمولة % = |Referral Fee| ÷ Net Proceeds للصنف ده. التوصيل = رسم الصنف (FBN) أو نصيبه من رسم الأوردر (FBB).
    المرتجعات (Net سالب) متتحسبش كبيع."""
    def num(x):
        try:
            return float(str(x).replace(",", "").strip() or 0)
        except Exception:
            return 0.0
    best = {}      # (sku, "ALL"/"FBN"/"FBB") -> (date, commission_pct, delivery)
    complete = incomplete = 0
    for nm, rs in fin_data.items():
        om, im = {}, {}
        for r in (orders_by_store.get(nm) or []):
            fb = str(r.get("is_fulfilled_by_noon", "")).strip().lower() == "true"
            o = om.setdefault(str(r.get("order_nr", "")).strip(), {"fbn": set(), "items": []})
            o["fbn"].add(fb)
            sk = str(r.get("sku", "")).strip().upper()
            if sk:
                o["items"].append(sk)
            im[str(r.get("item_nr", "")).strip()] = fb
        order_fee = {}   # رسوم على مستوى الأوردر (بدون كود)
        items = []
        for r in rs:
            if str(r.get("Transaction Type", "")).strip().lower() not in ("order", "order_update"):
                continue
            on = str(r.get("Order Nr", "")).strip()
            fee = (num(r.get("Fullfilment & Logistics Fees including VAT")) + num(r.get("Shipping Credits including VAT"))
                   + num(r.get("Other Order Fees including VAT")) + num(r.get("Order Subsidies including VAT")))
            sku = str(r.get("SKUs", "")).split(",")[0].strip().upper()
            if not sku:
                if fee < 0 and on:
                    order_fee[on] = order_fee.get(on, 0.0) + fee
                continue
            net, ref = num(r.get("Net Proceeds")), num(r.get("Referral Fee including VAT"))
            if net > 0 and ref < 0:
                d = str(r.get("Order Date", "") or r.get("Transaction Date", "")).strip()[:10]
                items.append((on, str(r.get("Item Nr", "")).strip(), sku, net, ref, fee, d))
        n_items = {}
        for it in items:
            n_items[it[0]] = n_items.get(it[0], 0) + 1
        for on, itn, sku, net, ref, fee, d in items:
            if fee < 0:
                dlv = fee
            elif order_fee.get(on):
                o = om.get(on)
                dlv = order_fee[on] / max(len(o["items"]) if o and o["items"] else n_items.get(on, 1), 1)
            else:
                incomplete += 1
                continue
            complete += 1
            if itn in im:
                fb = im[itn]
            else:
                o = om.get(on)
                fb = list(o["fbn"])[0] if o and len(o["fbn"]) == 1 else None
            model = "FBN" if fb is True else ("FBB" if fb is False else None)
            val = (d, round(ref / net * -100, 2), round(dlv, 2))
            for key in ((sku, "ALL"),) + (((sku, model),) if model else ()):
                if key not in best or val[0] > best[key][0]:
                    best[key] = val
    rows = []
    for sku in sorted({k[0] for k in best}):
        a, n, b = best.get((sku, "ALL")), best.get((sku, "FBN")), best.get((sku, "FBB"))
        rows.append([sku, a[2], a[1], n[2] if n else "", n[1] if n else "", b[2] if b else "", b[1] if b else "",
                     n[0] if n else "", b[0] if b else "", uploaded])
    log(f"🧮 COM_API: {len(rows)} كود (مبيعات مكتملة {complete}، ناقصة التوصيل {incomplete})")
    return rows


# ═══════════════════════ كتابة الشيت: بناء وتحديث (Upsert) — عمر ما بيمسح الأول ═══════════════════════
def _su_norm(v):
    t = str(v if v is not None else "").strip()
    if t == "":
        return ""
    if t.lower() in ("true", "false"):                  # TRUE (الشيت) = True (بايثون)
        return t.lower()
    try:
        return repr(round(float(t.replace(",", "")), 6))
    except ValueError:
        return t


def _su_retry(fn, log, tag, tries=6):
    for a in range(tries):
        try:
            return fn()
        except Exception as e:
            if a == tries - 1:
                raise
            log(f"⚠️ {tag}: محاولة {a + 1} فشلت: {str(e)[:120]}")
            _m = str(e)
            time.sleep(65 if ("429" in _m or "Quota" in _m or "quota" in _m) else 3 * (a + 1))


def _su_fill(ws, cols, values, log, tag, vio="RAW", chunk=2000):
    """يكتب العنوان ثم الصفوف على دفعات — أي دفعة تفشل بعد 5 محاولات بترمي خطأ (مفيش صمت)."""
    _su_retry(lambda: ws.resize(rows=max(len(values) + 10, 100), cols=max(len(cols), 6)), log, tag)
    _su_retry(lambda: ws.update(range_name="A1", values=[list(cols)], value_input_option=vio), log, tag)
    for i in range(0, len(values), chunk):
        part = values[i:i + chunk]
        _su_retry(lambda p=part: ws.append_rows(p, value_input_option=vio), log, tag)
        log(f"⬆️ {tag}: {min(i + chunk, len(values))}/{len(values)}")


def _su_swap(ss, ws, tab, cols, values, log, vio):
    """يبني التاب كامل في تاب مؤقت، ولما يكتمل بس بيبدّل القديم — القديم عمره ما بيتمسح قبل ما الجديد يتكتب."""
    import gspread
    tmp_title = (tab + "_new")[:100]
    try:
        ss.del_worksheet(ss.worksheet(tmp_title))
    except gspread.exceptions.WorksheetNotFound:
        pass
    tmp = ss.add_worksheet(title=tmp_title, rows="100", cols=str(max(len(cols), 6)))
    try:
        _su_fill(tmp, cols, values, log, tab, vio)
    except Exception as e:
        try:
            ss.del_worksheet(tmp)
        except Exception:
            pass
        if "limit of" in str(e) and "cells" in str(e):
            raise RuntimeError(f"ملف Google Sheet وصل حد الـ 20 مليون خلية — محتاج {len(values) * len(cols):,} خلية لتاب {tab}. امسح تابات/صفوف قديمة وجرّب تاني") from None
        raise
    ss.del_worksheet(ws)
    tmp.update_title(tab)


def sheet_replace_safe(ss, tab, cols, values, log, vio="RAW"):
    """للتابات اللي بتتبني كلها كل مرة (مخزون/كتالوج/كاش): بناء تاب جديد كامل ثم تبديل — مفيش مسح قبل الكتابة."""
    import gspread
    cols = list(cols)
    if not values:
        log(f"⚠️ {tab}: مفيش بيانات — الشيت ما اتلمسش")
        return
    try:
        ws = ss.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab, rows="100", cols=str(max(len(cols), 6)))
        _su_fill(ws, cols, values, log, tab, vio)
        log(f"🆕 {tab}: اتبنى من الصفر — {len(values)} صف")
        return
    _su_swap(ss, ws, tab, cols, values, log, vio)
    log(f"🔁 {tab}: اتبنى كامل واتبدّل — {len(values)} صف")


def sheet_upsert(ss, tab, cols, values, key_cols, log, orphan_set=None, scope=None, since=None, vio="RAW", cmp_cols=None):
    """يضيف الصفوف الجديدة ويحدّث اللي اتغيّر (بالمفتاح key_cols) من غير ما يمسح حاجة.
    cmp_cols: الأعمدة اللي تغيّرها يعتبر الصف اتحدّث (الافتراضي كلها).
    orphan_set/scope/since: صفوف في الشيت ومش في البيانات الجديدة (في متاجر scope=(عمود, قيم) ومن تاريخ since=(عمود, تاريخ))
    بتتعلّم بالقيم دي بدل ما تتمسح."""
    import gspread
    from gspread.utils import rowcol_to_a1
    if not values:
        log(f"⚠️ {tab}: مفيش بيانات — الشيت ما اتلمسش")
        return
    cols = list(cols)
    try:
        ws = ss.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        ws = None
    ex = _su_retry(lambda: ws.get_all_values(), log, tab) if ws is not None else []
    hdr = [str(h).strip() for h in ex[0]] if ex else []
    has_data = len(ex) > 1
    if ws is None:
        ws = ss.add_worksheet(title=tab, rows="100", cols=str(max(len(cols), 6)))
        _su_fill(ws, cols, values, log, tab, vio)
        log(f"🆕 {tab}: اتبنى من الصفر — {len(values)} صف")
        return
    if not hdr or not has_data:
        _su_fill(ws, cols, values, log, tab, vio)
        log(f"🆕 {tab}: كان فاضي — اتكتب {len(values)} صف")
        return
    if not all(k in cols and k in hdr for k in key_cols):     # أول مرة بعد إضافة عمود المفتاح: بناء كامل + تبديل
        _su_swap(ss, ws, tab, cols, values, log, vio)
        log(f"🔁 {tab}: اتبنى من جديد بعمود المفتاح — {len(values)} صف")
        return
    new_cols = [c for c in cols if c not in hdr]
    fh = hdr + new_cols
    if new_cols:
        if ws.col_count < len(fh):
            _su_retry(lambda: ws.resize(cols=len(fh)), log, tab)
        _su_retry(lambda: ws.update(range_name="A1", values=[fh], value_input_option="RAW"), log, tab)
    W = len(fh)
    pos = {c: i for i, c in enumerate(fh)}
    kix = [pos[k] for k in key_cols]
    idx = {}
    for i, row in enumerate(ex[1:], start=2):
        row = list(row) + [""] * (W - len(row))
        k = "|".join(str(row[j]).strip() for j in kix)
        if k.strip("|") and k not in idx:
            idx[k] = (i, row)
    ci = {c: j for j, c in enumerate(cols)}
    newmap = {}
    for v in values:
        newmap["|".join(str(v[ci[k]]).strip() for k in key_cols)] = v
    cc = [c for c in (cmp_cols or cols) if c in ci]
    ups, adds, same, orph, seen = [], [], 0, 0, set()
    _dc, _dx = {}, {}
    for k, v in newmap.items():
        if k in idx:
            i, row = idx[k]
            seen.add(k)
            _dif = [c for c in cc if _su_norm(row[pos[c]]) != _su_norm(v[ci[c]])]
            if _dif:
                for c in _dif:                                   # تشخيص: أنهي أعمدة بتفرّق (عشان نعرف ليه صفوف كتير بتتحدّث)
                    _dc[c] = _dc.get(c, 0) + 1
                    if c not in _dx:
                        _dx[c] = (str(row[pos[c]])[:40], str(v[ci[c]])[:40])
                tgt = list(row)
                for c in cols:
                    tgt[pos[c]] = v[ci[c]]
                ups.append((i, tgt))
            else:
                same += 1
        else:
            adds.append([v[ci[c]] if c in ci else "" for c in fh])
    if orphan_set and scope:
        sc_col, sc_vals = scope
        for k, (i, row) in idx.items():
            if k in seen or str(row[pos[sc_col]]).strip() not in sc_vals:
                continue
            if since and str(row[pos[since[0]]]).strip()[:10] < since[1]:
                continue
            if any(_su_norm(row[pos[c]]) != _su_norm(x) for c, x in orphan_set.items() if c in pos):
                tgt = list(row)
                for c, x in orphan_set.items():
                    if c in pos:
                        tgt[pos[c]] = x
                ups.append((i, tgt))
                orph += 1
    last = "".join(ch for ch in rowcol_to_a1(1, W) if not ch.isdigit())
    for s in range(0, len(ups), 500):
        data = [{"range": f"A{i}:{last}{i}", "values": [r]} for i, r in ups[s:s + 500]]
        _su_retry(lambda d=data: ws.batch_update([dict(x) for x in d], value_input_option=vio), log, tab)
    for s in range(0, len(adds), 2000):
        part = adds[s:s + 2000]
        _su_retry(lambda p=part: ws.append_rows(p, value_input_option=vio), log, tab)
    if _dc and len(ups) - orph > 200:                            # كتير اتحدّث: اطبع أكتر الأعمدة اختلافاً مع مثال (الشيت ← الجديد)
        for c, n in sorted(_dc.items(), key=lambda x: -x[1])[:5]:
            log(f"🔎 {tab}: عمود «{c}» اختلف في {n} صف — مثال: «{_dx[c][0]}» ← «{_dx[c][1]}»")
    log(f"⬆️ {tab}: {len(adds)} جديد · {len(ups) - orph} اتحدّث" + (f" · {orph} اتعلّموا" if orph else "") + f" · {same} زي ما هو — الإجمالي في الشيت {len(idx) + len(adds)}")


def write_raw_tab(tab, headers, rows, log, keys=None, cmp_cols=None):
    """keys → بناء وتحديث بالمفتاح. من غيره → بناء تاب جديد كامل ثم تبديل (عمره ما بيمسح قبل ما الجديد يكتمل)."""
    if pg_side(log, lambda: pg_write_tab(tab, headers, rows, log, keys=keys)):
        return
    cp = creds_path()
    if not cp:
        raise RuntimeError(f"مفيش ملف اعتماد Google في {GSHEET_CREDS_DIR}")
    ss = gspread.authorize(ServiceAccountCredentials.from_json_keyfile_name(cp, GSHEET_SCOPE)).open(GSHEET_NAME)
    if not rows:                                   # لو فاضي: نتأكد إن التاب موجود بعنوانه بس (من غير مسح)
        try:
            ss.worksheet(tab)
        except gspread.exceptions.WorksheetNotFound:
            ws = ss.add_worksheet(title=tab, rows="100", cols=str(max(len(headers), 6)))
            ws.update(range_name="A1", values=[list(headers)], value_input_option="USER_ENTERED")
        return
    if keys:
        sheet_upsert(ss, tab, headers, rows, keys, log, vio="USER_ENTERED", cmp_cols=cmp_cols)
    else:
        sheet_replace_safe(ss, tab, headers, rows, log, vio="USER_ENTERED")
    log(f"✅ اتكتب في تاب {tab} ({len(rows)} صف)")


# ---------------- الواجهة ----------------

class App(tk.Frame):
    def __init__(self, root):
        super().__init__(root, bg="#d9d9d9")
        self.pack(fill="both", expand=True, padx=8, pady=8)
        self.stop = threading.Event()
        self.running = False
        self.next_run = None
        self.stores = []
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="➕ ضيف ملف مفتاح", command=self.add_key).grid(row=0, column=0, padx=3)
        ttk.Button(top, text="🗑️ شيل المحدد", command=self.rm_key).grid(row=0, column=1, padx=3)
        self.lb = tk.Listbox(top, height=5, selectmode="extended")
        self.lb.grid(row=1, column=0, columnspan=8, sticky="ew", pady=4)
        top.columnconfigure(7, weight=1)
        ttk.Label(top, text="آخر كام يوم:").grid(row=2, column=0, sticky="e")
        self.days = tk.IntVar(value=30)
        ttk.Spinbox(top, from_=1, to=90, width=5, textvariable=self.days).grid(row=2, column=1, sticky="w")
        ttk.Label(top, text="فرق التوقيت عن UTC (ساعات):").grid(row=2, column=2, sticky="e")
        self.tz = tk.DoubleVar(value=3)
        ttk.Spinbox(top, from_=-12, to=14, increment=1, width=5, textvariable=self.tz).grid(row=2, column=3, sticky="w")
        ttk.Label(top, text="الدولة:").grid(row=2, column=4, sticky="e")
        self.country = tk.StringVar(value="sa")
        ttk.Entry(top, width=5, textvariable=self.country).grid(row=2, column=5, sticky="w")
        self.write = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text=f"اكتب في تاب {TEST_TAB}", variable=self.write).grid(row=3, column=0, columnspan=3, sticky="w")
        global PG_MODE
        PG_MODE = pg_flag_load() if pg_url() else "sheet"
        self.pg_mode = tk.StringVar(value=PG_LABELS[PG_MODE])
        ttk.Label(top, text="🐘 مكان الكتابة (كل البوتات: نون وتريندول وأمازون وحسابات المتاجر):").grid(row=33, column=0, columnspan=2, sticky="w")
        _pgcb = ttk.Combobox(top, textvariable=self.pg_mode, values=list(PG_LABELS.values()), state="readonly", width=24)
        _pgcb.grid(row=33, column=2, columnspan=2, sticky="w")
        self.pg_mode.trace_add("write", lambda *a: pg_toggle(self.pg_mode, self.log))
        ttk.Button(top, text="🔑 رابط Neon", command=pg_ask_url).grid(row=33, column=4, columnspan=2, sticky="w")
        self.storage = tk.BooleanVar(value=True)
        self.storage_n = tk.IntVar(value=2)
        ttk.Checkbutton(top, text=f"🏬 سحب رسوم التخزين الشهرية (FBN) → تاب {STORAGE_TAB} — آخر", variable=self.storage).grid(row=34, column=0, columnspan=4, sticky="w")
        ttk.Spinbox(top, from_=1, to=12, width=4, textvariable=self.storage_n).grid(row=34, column=4, sticky="w")
        ttk.Label(top, text="شهر").grid(row=34, column=5, sticky="w")
        ttk.Button(top, text="🏬 اسحب التخزين بس (من غير باقي التقرير)", command=self.go_storage).grid(row=35, column=0, columnspan=3, padx=4, pady=3, sticky="w")
        self.stock = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text=f"اسحب المخزون FBN كمان (تاب {STOCK_TAB})", variable=self.stock).grid(row=4, column=0, columnspan=3, sticky="w")
        self.odays = tk.IntVar(value=60)
        ttk.Label(top, text="أيام الأوردرات لربط FBN/FBB بالمالية:").grid(row=6, column=4, columnspan=2, sticky="e")
        ttk.Spinbox(top, from_=30, to=180, width=5, textvariable=self.odays).grid(row=6, column=6, sticky="w")
        self.fin = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text=f"اسحب المعاملات المالية (عمولة/توصيل) خام → تاب {FIN_TAB}", variable=self.fin).grid(row=6, column=0, columnspan=4, sticky="w")
        self.cat_dir = tk.StringVar()
        ttk.Label(top, text="مجلد كتالوج جاهز إضافي (اختياري):").grid(row=5, column=4, columnspan=2, sticky="e")
        ttk.Entry(top, width=22, textvariable=self.cat_dir).grid(row=5, column=6, sticky="w")
        self.cat_api = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text=f"الكتالوج (أسعار/أقسام) من الـ API → تاب {LIVE_TAB}", variable=self.cat_api).grid(row=5, column=0, columnspan=4, sticky="w")
        ttk.Label(top, text="export_code كتالوج جاهز (اختياري):").grid(row=4, column=3, columnspan=2, sticky="e")
        self.cat_code = tk.StringVar()
        ttk.Entry(top, width=18, textvariable=self.cat_code).grid(row=4, column=5, sticky="w")
        self.btn = ttk.Button(top, text="▶ شغّل (كل المتاجر مع بعض)", command=self.go)
        self.btn.grid(row=3, column=3, columnspan=3, sticky="ew", pady=4)
        ttk.Button(top, text="🔎 فحص: هل فيه قايمة تصديرات؟", command=self.probe).grid(row=3, column=6, padx=4)
        self.auto = tk.BooleanVar(value=False)
        self.every = tk.IntVar(value=60)
        ttk.Checkbutton(top, text="⏰ تشغيل تلقائي كل", variable=self.auto, command=self.toggle_auto).grid(row=7, column=0, columnspan=2, sticky="w")
        ttk.Spinbox(top, from_=10, to=1440, increment=10, width=6, textvariable=self.every).grid(row=7, column=2, sticky="w")
        ttk.Label(top, text="دقيقة").grid(row=7, column=3, sticky="w")
        self.next_lbl = ttk.Label(top, text="")
        self.next_lbl.grid(row=7, column=4, columnspan=3, sticky="w")
        self.times = tk.StringVar(value="")
        self._last_fixed = ""
        ttk.Label(top, text="🕐 مواعيد محددة (HH:MM مفصولة بفاصلة):").grid(row=30, column=0, columnspan=2, sticky="e")
        ttk.Entry(top, textvariable=self.times, width=34).grid(row=30, column=2, columnspan=4, sticky="w")
        self.cat_auto = tk.BooleanVar(value=False)
        self.cat_every = tk.IntVar(value=180)
        self.cat_next = None
        self.cat_running = False
        ttk.Checkbutton(top, text="🗂️ تحديث الكتالوج + لقطة السعر تلقائي كل", variable=self.cat_auto, command=self.toggle_cat).grid(row=31, column=0, columnspan=2, sticky="w")
        ttk.Spinbox(top, from_=15, to=1440, increment=15, width=6, textvariable=self.cat_every).grid(row=31, column=2, sticky="w")
        ttk.Label(top, text="دقيقة").grid(row=31, column=3, sticky="w")
        self.cat_lbl = ttk.Label(top, text="")
        self.cat_lbl.grid(row=31, column=4, columnspan=2, sticky="w")
        ttk.Button(top, text="🗂️ حدّث الكتالوج دلوقتي", command=lambda: self.go_cat(False)).grid(row=32, column=0, columnspan=2, sticky="w", pady=2)
        self.out = scrolledtext.ScrolledText(self, font=("Consolas", 10), height=18)
        self.out.pack(fill="both", expand=True, padx=8, pady=8)
        self.load()
        self.load_settings()
        for _v in (self.write, self.stock, self.fin, self.cat_api, self.storage, self.storage_n, self.times, self.cat_every):
            _v.trace_add('write', lambda *a: self.save_settings())
        self.after(5000, self.tick)

    def log(self, m):
        try:
            if LOG_HOOK:
                LOG_HOOK(str(m) + "\n")
        except Exception:
            pass
        try:
            self.after(0, lambda: (self.out.insert("end", m + "\n"), self.out.see("end")))
        except Exception:
            pass

    def load(self):
        try:
            paths = json.load(open(STORES_FILE, encoding="utf-8"))
        except Exception:
            paths = []
        for p in paths:
            self.add_path(p, save=False)

    def save(self):
        json.dump([s.path for s in self.stores], open(STORES_FILE, "w", encoding="utf-8"), ensure_ascii=False)

    def add_path(self, p, save=True):
        try:
            st = Store(p)
        except Exception as e:
            self.log(f"⚠️ {os.path.basename(p)}: {e}")
            return
        if any(s.path == p for s in self.stores):
            return
        self.stores.append(st)
        self.lb.insert("end", f"{st.name}   —   project {st.key['project_code']}")
        if save:
            self.save()

    def add_key(self):
        for p in filedialog.askopenfilenames(title="ملفات مفاتيح نون", filetypes=[("JSON", "*.json"), ("All", "*.*")]):
            self.add_path(p)

    def rm_key(self):
        for i in reversed(self.lb.curselection()):
            self.lb.delete(i)
            del self.stores[i]
        self.save()

    def go_storage(self):
        """يسحب رسوم التخزين الشهرية بس (من غير أوردرات/مخزون/مالية) ويكتبها في Noon_Storage_API."""
        if self.running or self.cat_running:
            self.log("⏳ في تشغيل شغال دلوقتي — جرّب بعد ما يخلص")
            return
        if not self.stores:
            messagebox.showwarning("", "ضيف ملف مفتاح واحد على الأقل")
            return
        self.running = True
        months = storage_months(self.storage_n.get(), self.tz.get())
        country = self.country.get().strip() or "sa"
        self.stop.clear()

        def run():
            try:
                self.log(f"🏬 سحب التخزين فقط — الشهور: {', '.join(months)}")
                data = {}

                def one(st):
                    try:
                        data[st.name] = fetch_storage(st, months, country, self.log, self.stop)
                    except Exception as e:
                        self.log(f"[{st.name}] ⚠️ تخزين: {type(e).__name__}: {str(e)[:150]}")
                with ThreadPoolExecutor(max_workers=len(self.stores)) as ex:
                    list(ex.map(one, self.stores))
                if any(data.values()):
                    write_storage(data, months, self.log)
                else:
                    self.log("⚠️ مفيش رسوم تخزين رجعت من أي متجر")
            except Exception as e:
                self.log(f"❌ سحب التخزين: {type(e).__name__}: {str(e)[:200]}")
            finally:
                self.running = False
        threading.Thread(target=run, daemon=True).start()

    def go(self, auto=False):
        if self.running:
            return
        if self.cat_running:
            self.log("⏳ تحديث الكتالوج شغال — جرّب بعد ما يخلص")
            if auto:
                self.next_run = time.time() + 120
            return
        if not self.stores:
            if not auto:
                messagebox.showwarning("", "ضيف ملف مفتاح واحد على الأقل")
            else:
                self.log("⚠️ مفيش مفاتيح نون — التشغيل التلقائي اتخطّى")
                self.next_run = time.time() + int(self.every.get()) * 60
            return
        self.running = True
        self.stop.clear()
        self.save_settings()
        self.btn.state(["disabled"])
        threading.Thread(target=self.job, daemon=True).start()

    def probe(self):
        if not self.stores:
            messagebox.showwarning("", "ضيف ملف مفتاح الأول")
            return
        st = self.stores[0]

        def run():
            self.log(f"🔎 [{st.name}] باجرب مسارات محتملة لقايمة التصديرات (قراءة بس)...")
            cands = [("POST", "/impex/v1/export/list", {}), ("POST", "/impex/v1/export/history", {}),
                     ("POST", "/impex/v1/exports/list", {}), ("GET", "/impex/v1/export/list", None),
                     ("GET", "/impex/v1/export/history", None), ("GET", "/impex/v1/export/category/list", None)]
            for m, path, body in cands:
                try:
                    with st.lock:
                        if st.sess is None:
                            st.sess = login(st, self.log)
                    s_ = st.sess
                    if s_ is None:
                        return
                    h = {"Accept": "application/json"}
                    r = s_.get(BASE + path, headers=h, timeout=30) if m == "GET" else s_.post(BASE + path, json=body, headers=h, timeout=30)
                    self.log(f"  {m} {path} → {r.status_code}  {(r.text or '')[:300].replace(chr(10), ' ')}")
                except Exception as e:
                    self.log(f"  {m} {path} → خطأ {type(e).__name__}")
            self.log("🔎 خلص الفحص — ابعتلي الأسطر دي.")
        threading.Thread(target=run, daemon=True).start()

    def cat_folders(self):
        return [os.path.join(os.path.expanduser("~"), "Downloads"), APP_DIR, self.cat_dir.get().strip()]

    def build_inventory(self, hdrs, srows, rows_orders, up):
        """يبني Inventory_API بنفس أعمدة Inventory: SKU | Warehouse | Stock | Monthly Sales | Image URL | Date Uploaded."""
        ix = {h: i for i, h in enumerate(hdrs)}
        sku_col = "sku" if "sku" in ix else "psku_canonical"
        need = (sku_col, "warehouse_code", "qty")
        if any(n not in ix for n in need):
            self.log(f"⚠️ أعمدة المخزون مش زي المتوقع، ناقص: {[n for n in need if n not in ix]}")
            return
        agg = {}
        for r in srows:
            if "inventory_type" in ix and str(r[ix["inventory_type"]]).strip().lower() != "saleable":
                continue
            try:
                q = int(float(str(r[ix["qty"]]).replace(",", "") or 0))
            except Exception:
                q = 0
            key = (str(r[ix[sku_col]]).strip(), str(r[ix["warehouse_code"]]).strip())
            agg[key] = agg.get(key, 0) + q
        sales = {}
        for o in rows_orders:
            sales[o[0]] = sales.get(o[0], 0) + 1
        imgs = links_map(self.log)
        out = [[sku, wh, q, sales.get(sku, 0), imgs.get(sku.upper(), ""), up] for (sku, wh), q in sorted(agg.items())]
        self.log(f"📦 Inventory_API: {len(out)} صف (saleable بس)"
                 + ("" if int(self.days.get()) >= 30 else " — ⚠️ Monthly Sales غير دقيق لأن الأيام أقل من 30"))
        if self.write.get() and out:
            write_raw_tab(INV_TAB, INV_HEADERS, out, self.log)
            self.log(f"✅ اتكتب في تاب {INV_TAB} (Inventory القديم ماتلمسش)")

    def write_live(self, api_cat, up):
        lh = ["Store"]
        for rs in api_cat.values():
            for h in rs[0].keys():
                if h not in lh:
                    lh.append(h)
        lh.append("Date Uploaded")
        lrows = [[nm] + [r.get(h, "") for h in lh[1:-1]] + [up] for nm, rs in api_cat.items() for r in rs]
        self.log(f"🗂️ الكتالوج من الـ API: {len(lrows)} صف من {len(api_cat)} متجر")
        if self.write.get() and lrows:
            write_raw_tab(LIVE_TAB, lh, lrows, self.log)
            self.log(f"✅ الكتالوج اتكتب في تاب {LIVE_TAB} (LIVE القديم ماتلمسش)")

    def cat_job(self):
        """مهمة خفيفة بجدولة لوحدها: كتالوج طازة → LIVE_API + لقطة السعر (PriceHistory_API). مفيش أوردرات ولا مالية."""
        try:
            t0 = time.time()
            country = self.country.get().strip() or "sa"
            self.log(f"🗂️ تحديث الكتالوج + لقطة السعر — {len(self.stores)} متجر")
            fresh = {}

            def one(st):
                return st, get_catalog(st, country, self.log, self.stop, "", None, force=True)
            with ThreadPoolExecutor(max_workers=max(1, len(self.stores))) as ex:
                for st, c in ex.map(one, self.stores):
                    if c:
                        fresh[st.name] = c
                    else:
                        self.log(f"⚠️ [{st.name}] الكتالوج ماجاش — اتخطّى المتجر ده")
            if not fresh:
                self.log("⚠️ مفيش كتالوج جه من أي متجر — مفيش تحديث")
                return
            up = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            hist = read_history(self.log)
            today_local = (dt.datetime.utcnow() + dt.timedelta(hours=float(self.tz.get()))).strftime("%Y-%m-%d %H:%M:%S")
            changed = 0
            for nm, c in fresh.items():
                _stt = {}
                _n = snapshot_prices(hist, nm, c, today_local, _stt)
                changed += _n
                log_snapshot(self.log, nm, _stt, _n)
            if self.write.get():
                # LIVE_API بنحدّثه بالمتاجر اللي جت بس — لو متجر فشل بنسيب صفوفه القديمة
                merged = dict(fresh)
                try:
                    old_rows = read_tab_values(LIVE_TAB)
                    if len(old_rows) > 1:
                        oh = [x.strip() for x in old_rows[0]]
                        if "Store" in oh:
                            si = oh.index("Store")
                            for r in old_rows[1:]:
                                nm = r[si] if si < len(r) else ""
                                if nm and nm not in fresh:
                                    merged.setdefault(nm, []).append({k: v for k, v in zip(oh, r) if k not in ("Store", "Date Uploaded")})
                except Exception as e:
                    self.log(f"ℹ️ ماقدرتش أقرا LIVE_API القديم ({str(e)[:80]}) — هكتب المتاجر اللي جت بس")
                self.write_live(merged, up)
                if changed:
                    write_raw_tab(HIST_TAB, HIST_HEADERS, history_rows(hist), self.log)
                    self.log(f"🕘 سجل الأسعار: {changed} تغيير جديد اتسجّل")
                else:
                    self.log("🕘 سجل الأسعار: مفيش تغيير من آخر تسجيل")
            self.log(f"✅ خلص تحديث الكتالوج في {int(time.time() - t0)} ثانية")
        except Exception as e:
            self.log(f"❌ تحديث الكتالوج: {type(e).__name__}: {e}")

    def job(self):
        try:
            t0 = time.time()
            today = dt.datetime.utcnow().date()
            frm = (today - dt.timedelta(days=int(self.days.get()))).isoformat()
            to = today.isoformat()
            country = self.country.get().strip() or "sa"
            self.log(f"📅 من {frm} إلى {to} (UTC) — {len(self.stores)} متجر")

            self.api_cat = {}
            self.fin_data = {}
            self.storage_data = {}
            self.fin_orders = {}
            _fb = {}
            _fb_lock = threading.Lock()

            def fallback():
                with _fb_lock:
                    if "v" not in _fb:
                        _fb["v"] = live_catalog(self.log, LIVE_TAB) or live_catalog(self.log, "LIVE")
                    return _fb["v"]

            def one(st):
                try:
                    return one_inner(st)
                except Exception as e:      # فشل متجر واحد ما يوقّفش الباقي
                    self.log(f"[{st.name}] ❌ {type(e).__name__}: {str(e)[:150]}")
                    return st, None, None, None

            def one_inner(st):
                o = run_export(st, "noon_noonoms_ordersexport", {"from_date": frm, "to_date": to}, self.log, self.stop)
                save_orders_handoff(st, frm, to, o)
                _nf = acct_orders_need_from(st)
                if o is not None and _nf and _nf < dt.date.fromisoformat(frm):   # «حسابات نون» محتاجة أوردرات أقدم (أول سحب) → نسحبها هنا مرة واحدة ونسيبها ليها
                    self.log(f"[{st.name}] 🧾 أوردرات أقدم لخطوة «حسابات نون» (من {_nf}) — مرة واحدة")
                    _w = run_export(st, "noon_noonoms_ordersexport", {"from_date": _nf.isoformat(), "to_date": to}, self.log, self.stop)
                    if _w:
                        save_orders_handoff(st, _nf, to, _w)
                c = None
                if self.cat_api.get():
                    c = get_catalog(st, country, self.log, self.stop, self.cat_code.get().strip(), self.cat_folders())
                    if c:
                        self.api_cat[st.name] = c
                if not c:
                    self.log(f"[{st.name}] ↩️ هاخد الأسعار من الشيت (LIVE) بدل الـ API")
                    c = fallback()
                k = None
                if self.stock.get():
                    for back in (1, 2, 3):
                        cd = (today - dt.timedelta(days=back)).isoformat()
                        self.log(f"[{st.name}] 📦 مخزون FBN بتاريخ إقفال {cd}")
                        k = run_export(st, "fbn_inventory_fbninventorydetailreport",
                                       {"country": country, "close_date": cd}, self.log, self.stop, prefer_ready=True)
                        if k:
                            break
                if self.fin.get():
                    f = run_export(st, "noon_financeweb_transactionviewreportonitemlevel",
                                   {"from_date": frm, "to_date": to}, self.log, self.stop)
                    if f:
                        self.fin_data[st.name] = f
                        od = int(self.odays.get())
                        if od > int(self.days.get()):
                            ofrm = (today - dt.timedelta(days=od)).isoformat()
                            self.log(f"[{st.name}] 🧾 أوردرات أقدم للربط بالمالية (من {ofrm})")
                            self.fin_orders[st.name] = run_export(st, "noon_noonoms_ordersexport",
                                                                  {"from_date": ofrm, "to_date": to}, self.log, self.stop) or o
                            if self.fin_orders[st.name] is not o:
                                save_orders_handoff(st, ofrm, to, self.fin_orders[st.name], keep_wider=True)
                        else:
                            self.fin_orders[st.name] = o
                if self.storage.get():
                    try:
                        _sm = storage_months(self.storage_n.get(), self.tz.get())
                        self.storage_data[st.name] = fetch_storage(st, _sm, country, self.log, self.stop)
                    except Exception as _e:
                        self.log(f"[{st.name}] ⚠️ تخزين: {type(_e).__name__}: {str(_e)[:150]}")
                return st, o, c, k
            with ThreadPoolExecutor(max_workers=len(self.stores)) as ex:
                res = list(ex.map(one, self.stores))
            up = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            hist = read_history(self.log)
            today_local = (dt.datetime.utcnow() + dt.timedelta(hours=float(self.tz.get()))).strftime("%Y-%m-%d %H:%M:%S")
            hist_changed = 0
            rows = []
            all_orders = rows
            pref_store = {}    # SKU → المتجر اللي جه منه (نجرّبه الأول في الصور)
            failed_orders = []
            for st, o, c, k in res:
                if o is None:
                    self.log(f"⚠️ [{st.name}] مفيش أوردرات — اتخطّى")
                    failed_orders.append(st.name)
                    continue
                if c:
                    self.log(f"[{st.name}] 🗂️ الكتالوج: {len(c)} صف")
                if st.name in self.api_cat:   # اللقطة من الكتالوج الحقيقي للمتجر بس (مش من شيت LIVE المشترك)
                    _stt = {}
                    _n = snapshot_prices(hist, st.name, c, today_local, _stt)
                    hist_changed += _n
                    log_snapshot(self.log, st.name, _stt, _n)
                r, sk, npz = build_rows(st.name, o, build_catalog_map(c), float(self.tz.get()), up, hist)
                self.log(f"[{st.name}] 🧾 {len(o)} صف → {len(r)} (استبعاد ملغي/غير مُسلَّم: {sk}) — بدون سعر: {npz} — سعر من السجل: {build_rows.last_hist}")
                for x in r:
                    if x and x[0]:
                        pref_store.setdefault(x[0], st)
                rows += r
            if self.write.get() and hist_changed:
                write_raw_tab(HIST_TAB, HIST_HEADERS, history_rows(hist), self.log)
                self.log(f"🕘 سجل الأسعار: {hist_changed} تغيير جديد اتسجّل في {HIST_TAB}")
            elif hist_changed == 0:
                self.log("🕘 سجل الأسعار: مفيش تغيير في الأسعار من آخر تسجيل")
            if self.api_cat:
                self.write_live(self.api_cat, up)
            if self.fin_data:
                fh = ["Store"]
                for rs in self.fin_data.values():
                    if rs:
                        for h in rs[0].keys():
                            if h not in fh:
                                fh.append(h)
                for nm, rs in self.fin_data.items():
                    if rs:
                        self.log(f"[{nm}] 💰 أعمدة المعاملات المالية ({len(rs)} صف): {list(rs[0].keys())}")
                        self.log(f"[{nm}] 💰 أول صف: {list(rs[0].values())}")
                before = sum(len(rs) for rs in self.fin_data.values())
                fh, frows = tag_finance(self.fin_data, self.fin_orders, self.log)
                crows = build_com(self.fin_data, self.fin_orders, self.log, up)
                if self.write.get() and crows:
                    write_raw_tab(COM_TAB, COM_HEADERS, crows, self.log)
                    self.log(f"✅ اتكتب في تاب {COM_TAB} (COM القديم ماتلمسش)")
                self.log(f"💰 صفوف المالية: {len(frows)} من {before} (استبعدت التحويلات البنكية payment بس)")
                types = {}
                for rs in self.fin_data.values():
                    for r in rs:
                        t = str(r.get("Transaction Type", "")).strip()
                        types[t] = types.get(t, 0) + 1
                self.log(f"💰 أنواع المعاملات: {types}")
                for nm, rs in self.fin_data.items():
                    for r in rs:
                        try:
                            if float(str(r.get("Referral Fee including VAT", 0) or 0).replace(",", "")) != 0:
                                self.log(f"[{nm}] 💰 صف فيه عمولة: {list(r.values())}")
                                break
                        except Exception:
                            pass
                # تاب Finance_API_Raw للمراجعة بس (مفيش حاجة بتقراه) وبياخد ملايين الخلايا — اتقفلت الكتابة فيه
                # عشان حد الـ 20 مليون خلية. الحسابات الفعلية بتروح Noon_Finance_API من بوت حسابات نون.
                self.log(f"ℹ️ كتابة تاب {FIN_TAB} متقفلة (توفير مساحة) — الأرقام بتتسحب وبتتحفظ لحسابات نون")
            elif self.fin.get():
                self.log("⚠️ مفيش معاملات مالية رجعت من أي متجر")
            if self.stock.get():
                hdrs, srows = ["Store"], []
                for st, o, c, k in res:
                    if not k:
                        self.log(f"⚠️ [{st.name}] مفيش مخزون FBN")
                        continue
                    self.log(f"[{st.name}] 📦 أعمدة المخزون: {list(k[0].keys())}")
                    self.log(f"[{st.name}] 📦 أول صف: {list(k[0].values())}")
                    for h in k[0].keys():
                        if h not in hdrs:
                            hdrs.append(h)
                    for r in k:
                        srows.append([st.name] + [r.get(h, "") for h in hdrs[1:]])
                # توحيد الطول لو الأعمدة اختلفت بين المتاجر
                srows = [r + [""] * (len(hdrs) - len(r)) for r in srows]
                if srows:
                    self.build_inventory(hdrs, srows, rows_orders=all_orders, up=up)
                if self.write.get() and srows:
                    write_raw_tab(STOCK_TAB, hdrs, srows, self.log)
                    self.log(f"✅ المخزون اتكتب في تاب {STOCK_TAB} (الخام، من غير ما يلمس Inventory)")
            if self.storage.get() and self.write.get():
                if any(self.storage_data.values()):
                    try:
                        write_storage(self.storage_data, storage_months(self.storage_n.get(), self.tz.get()), self.log)
                    except Exception as _e:
                        self.log(f"⚠️ كتابة التخزين فشلت: {type(_e).__name__}: {str(_e)[:200]}")
                else:
                    self.log("⚠️ مفيش رسوم تخزين رجعت من أي متجر")
            rows.sort(key=lambda x: x[9], reverse=True)
            self.log(f"📊 الإجمالي: {len(rows)} صف")
            by = {}
            for r in rows:
                by[r[1][:10]] = by.get(r[1][:10], 0) + 1
            self.log("📆 عدد الصفوف لكل يوم (آخر 7): " + " | ".join(f"{k}:{v}" for k, v in sorted(by.items())[-7:]))
            if failed_orders:
                self.log(f"⚠️ ماكتبتش {TEST_TAB} لأن أوردرات ({', '.join(failed_orders)}) ماجتش — عشان ما أمسحش بياناتهم القديمة. شغّل تاني.")
            elif self.write.get() and rows:
                write_sheet(rows, self.log)
                self.log(f"✅ اتكتب في تاب {TEST_TAB} (الشيت القديم ماتلمسش)")
            self.log(f"⏱️ خلص في {int(time.time() - t0)} ثانية")
        except Exception as e:
            self.log(f"❌ {type(e).__name__}: {e}")
        finally:
            self.running = False
            if self.auto.get():
                self.next_run = time.time() + max(10, int(self.every.get())) * 60
            self.after(0, lambda: self.btn.state(["!disabled"]))


    # ---------- الإعدادات + الجدولة ----------
    def load_settings(self):
        try:
            s = json.load(open(SETTINGS_FILE, encoding="utf-8"))
            self.days.set(int(s.get("days", 30)))
            self.odays.set(int(s.get("odays", 60)))
            self.tz.set(float(s.get("tz", 3)))
            self.country.set(s.get("country", "sa"))
            self.every.set(int(s.get("every", 60)))
            self.auto.set(bool(s.get("auto", False)))
            for _n in ("write", "stock", "fin", "cat_api", "storage"):
                if _n in s:
                    getattr(self, _n).set(bool(s[_n]))
            self.storage_n.set(int(s.get("storage_n", 2)))
            self.times.set(str(s.get("times", "")))
            self.cat_every.set(int(s.get("cat_every", 180)))
            self.cat_auto.set(bool(s.get("cat_auto", False)))
            if self.cat_auto.get():
                self.cat_next = time.time() + 90          # أول تحديث كتالوج بعد 90 ثانية من فتح البرنامج
                self.log("🗂️ تحديث الكتالوج التلقائي مفعّل (من آخر مرة)")
            if self.auto.get():
                self.next_run = time.time() + 60          # أول تشغيل تلقائي بعد دقيقة من فتح البرنامج
                self.log("⏰ التشغيل التلقائي مفعّل (من آخر مرة) — أول تشغيل بعد دقيقة")
        except Exception:
            pass

    def save_settings(self):
        try:
            json.dump({"days": int(self.days.get()), "odays": int(self.odays.get()), "tz": float(self.tz.get()),
                       "country": self.country.get().strip() or "sa", "every": int(self.every.get()),
                       "auto": bool(self.auto.get()), "write": bool(self.write.get()), "stock": bool(self.stock.get()),
                       "storage": bool(self.storage.get()), "storage_n": int(self.storage_n.get()), "fin": bool(self.fin.get()), "cat_api": bool(self.cat_api.get()), "times": self.times.get(),
                       "cat_every": int(self.cat_every.get()), "cat_auto": bool(self.cat_auto.get())}, open(SETTINGS_FILE, "w", encoding="utf-8"))
        except Exception:
            pass

    def toggle_auto(self):
        if self.auto.get():
            self.next_run = time.time() + int(self.every.get()) * 60
            self.log(f"⏰ تشغيل تلقائي كل {int(self.every.get())} دقيقة (التشغيل الجاي: {time.strftime('%H:%M', time.localtime(self.next_run))})")
        else:
            self.next_run = None
            self.log("⏰ التشغيل التلقائي اتقفل")
        self.save_settings()

    def fixed_times(self):
        out = []
        for t in str(self.times.get()).replace("،", ",").replace(" ", ",").split(","):
            t = t.strip()
            try:
                out.append(dt.datetime.strptime(t, "%H:%M").strftime("%H:%M")) if t else None
            except ValueError:
                pass
        return out

    def tick(self):
        try:
            now = time.strftime("%H:%M")
            key = time.strftime("%Y-%m-%d ") + now
            if now in self.fixed_times() and key != self._last_fixed and not self.running and not self.cat_running:
                self._last_fixed = key
                self.log(f"🕐 موعد محدد {now} — تشغيل تلقائي...")
                self.go(auto=True)
        except Exception:
            pass
        try:
            if self.auto.get() and self.next_run:
                left = self.next_run - time.time()
                self.next_lbl.configure(text=f"التشغيل الجاي: {time.strftime('%H:%M', time.localtime(self.next_run))}")
                if left <= 0 and not self.running:
                    self.log("⏰ تشغيل تلقائي...")
                    self.go(auto=True)
            else:
                self.next_lbl.configure(text="")
        except Exception:
            pass
        try:
            if self.cat_auto.get() and self.cat_next:
                self.cat_lbl.configure(text=f"التحديث الجاي: {time.strftime('%H:%M', time.localtime(self.cat_next))}")
                if self.cat_next - time.time() <= 0 and not self.running and not self.cat_running:
                    self.log("🗂️ موعد تحديث الكتالوج التلقائي...")
                    self.go_cat(auto=True)
            else:
                self.cat_lbl.configure(text="")
        except Exception:
            pass
        self.after(15000, self.tick)

    def toggle_cat(self):
        if self.cat_auto.get():
            self.cat_next = time.time() + 30
            self.log(f"🗂️ تحديث الكتالوج + لقطة السعر تلقائي كل {int(self.cat_every.get())} دقيقة (أول تحديث بعد نص دقيقة)")
        else:
            self.cat_next = None
            self.log("🗂️ تحديث الكتالوج التلقائي اتقفل")
        self.save_settings()

    def go_cat(self, auto=False):
        if self.cat_running or self.running:
            if not auto:
                self.log("⏳ في تشغيل شغال دلوقتي — جرّب بعد ما يخلص")
            return
        if not self.stores:
            if not auto:
                messagebox.showwarning("", "ضيف ملف مفتاح واحد على الأقل")
            else:
                self.log("⚠️ مفيش مفاتيح نون — تحديث الكتالوج اتخطّى")
                self.cat_next = time.time() + max(15, int(self.cat_every.get())) * 60
            return
        self.cat_running = True
        self.stop.clear()

        def run():
            try:
                self.cat_job()
            finally:
                self.cat_running = False
                if self.cat_auto.get():
                    self.cat_next = time.time() + max(15, int(self.cat_every.get())) * 60

        threading.Thread(target=run, daemon=True).start()

    def shutdown(self):
        self.stop.set()


def build_ui(root):
    """بيتنادى من البرنامج الموحد: بيبني الواجهة جوه التاب ويرجّع دالة إيقاف."""
    a = App(root)
    return a.shutdown

'''

_EMBEDDED['TrendyolAPI'] = r'''# -*- coding: utf-8 -*-
"""
Trendyol_API.py — تجربة: بيانات تريندول من الـ API الرسمية → تابات جديدة في Google Sheet (الشيتات القديمة مابتتلمسش)
قراءة فقط من تريندول.

التابات الجديدة:
  DailyOrdersTrendyol_API  : نفس أعمدة DailyOrdersTrendyol + أعمدة زيادة في الآخر
  TrendyolProducts_API     : نفس أعمدة TrendyolProducts + مخزون وأسعار وعمولة
  Finance_TY_Raw           : المبيعات والمرتجعات (Settlements) — عمولة/صافي لكل منتج
  Finance_TY_Other_Raw     : فواتير العمولة والخصومات وأوامر الدفع (otherfinancials) — للبحث عن مصاريف الشحن

المتطلبات:  py -m pip install requests gspread oauth2client
مفاتيح تريندول (Seller ID / API Key / API Secret) بتتكتب في ملف محلي trendyol_stores.json جنب الأداة بس — ماتتبعتش لأي حد.
"""
import glob, json, os, time, threading, datetime as dt
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

import requests
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
except Exception:
    openpyxl = None
import gspread
from oauth2client.service_account import ServiceAccountCredentials

BASE = "https://apigw.trendyol.com"
APP_DIR = os.environ.get("PLATFORMS_BASE_DIR") or os.path.dirname(os.path.abspath(__file__))
STORES_FILE = os.path.join(APP_DIR, "trendyol_stores.json")
LOG_HOOK = None
SETTINGS_FILE = os.path.join(APP_DIR, "trendyol_api_settings.json")

GSHEET_CREDS_DIR = r"\\WIN-181U7G4HFM5\Share For Markting\NOONIMG\Copy"
GSHEET_NAME = "Complaints"
GSHEET_SCOPE = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

ORD_TAB, PRD_TAB, FIN_TAB, OTH_TAB = "DailyOrdersTrendyol_API", "TrendyolProducts_API", "Finance_TY_Raw", "Finance_TY_Other_Raw"
CRG_TAB = "Finance_TY_Cargo_Raw"   # شحن لكل طلب/طرد من فواتير الشحن (cargo-invoice)
CRG_HEADERS = ["Store", "Invoice Id", "Invoice Date", "Order Number", "Parcel Unique Id", "Package Type", "Amount", "Desi"]
ORD_HEADERS = ["Row Key", "Barcode", "Order Date", "Order Number", "Order Status", "Quantity", "Net Unit Price",
               "Invoice Amount", "Commission Rate", "Image URL", "Product Name", "Date Uploaded",
               "Store", "Package Id", "Package Status", "Line Gross", "Seller Discount", "TY Discount", "VAT Rate", "Cargo"]
PRD_HEADERS = ["Content ID", "Barcode", "Product Name", "Image URL", "Date Uploaded",
               "Store", "Stock", "Sale Price", "List Price", "Commission %", "On Sale", "Stock Code", "Product Main Id", "Product URL"]
FIN_HEADERS = ["Store", "Type", "Transaction Date", "Order Date", "Order Number", "Barcode", "Package Id",
               "Commission Rate", "Commission Amount", "Seller Revenue", "Debt", "Credit", "Payment Order Id",
               "Payment Date", "Description", "Receipt Id", "Item Id", "Transaction Type"]
OTHER_TYPES = ("PaymentOrder", "DeductionInvoices", "CreditNote", "CommissionInvoice")
DAY = 86400000
TS_OFFSET_H = 3        # توقيت السعودية (GMT+3)


def num(x):
    try:
        return float(str(x).replace(",", "").strip() or 0)
    except Exception:
        return 0.0


def sv(x):
    """نص من غير كلمة None لو القيمة فاضية."""
    return "" if x is None else str(x)


def ms_to_local(ms, tz=TS_OFFSET_H):
    try:
        return (dt.datetime.utcfromtimestamp(int(ms) / 1000) + dt.timedelta(hours=tz)).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return ""


# ---------------- المتاجر ----------------
class TStore:
    def __init__(self, d):
        self.name = (d.get("name") or d["seller_id"]).strip()
        self.sid = str(d["seller_id"]).strip()
        self.key = d["api_key"].strip()
        self.secret = d["api_secret"].strip()
        self.sf = (d.get("store_front") or "SA").strip().upper()
        self.sess = requests.Session()
        self.sess.auth = (self.key, self.secret)
        self.sess.headers.update({"User-Agent": f"{self.sid} - SelfIntegration", "storeFrontCode": self.sf,
                                  "Accept": "application/json"})

    def as_dict(self):
        return {"name": self.name, "seller_id": self.sid, "api_key": self.key, "api_secret": self.secret, "store_front": self.sf}


def get_json(st, path, params, log, tries=6, extra_headers=None):
    """GET مع إعادة محاولة عند 429/5xx. بيرمي استثناء واضح لو 401/403/404."""
    url = BASE + path.format(sid=st.sid)
    last = ""
    for a in range(tries):
        try:
            r = st.sess.get(url, params=params, headers=extra_headers or {}, timeout=60)
        except Exception as e:
            last = f"{type(e).__name__}: {str(e)[:120]}"
            time.sleep(2 * (a + 1))
            continue
        if r.status_code == 200:
            try:
                return r.json()
            except Exception:
                raise RuntimeError(f"رد مش JSON: {(r.text or '')[:150]}")
        last = f"HTTP {r.status_code}: {(r.text or '')[:200].replace(chr(10), ' ')}"
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(3 * (a + 1))
            continue
        raise RuntimeError(last)
    raise RuntimeError(f"فشل بعد {tries} محاولات — {last}")


# ---------------- الأوردرات ----------------
def _order_window(st, a, b, log, out):
    """يسحب كل صفحات نافذة (أقصى أسبوعين). لو الإجمالي فوق 9500 باكدج بيقسم النافذة نصين."""
    page = 0
    while True:
        j = get_json(st, "/integration/order/sellers/{sid}/v2/orders",
                     {"startDate": a, "endDate": b, "page": page, "size": 200, "supplierId": st.sid,
                      "orderByField": "PackageLastModifiedDate", "orderByDirection": "ASC"}, log)
        total = int(j.get("totalElements") or 0)
        if page == 0 and total > 9500 and (b - a) > 3600 * 1000:
            mid = (a + b) // 2
            _order_window(st, a, mid, log, out)
            _order_window(st, mid + 1, b, log, out)
            return
        out.extend(j.get("content") or [])
        page += 1
        if page >= int(j.get("totalPages") or 0) or page > 49:
            return


def fetch_orders(st, days, log):
    end = int(time.time() * 1000)
    start = end - days * DAY
    out, a = [], start
    while a < end:
        b = min(a + 14 * DAY - 1, end)
        _order_window(st, a, b, log, out)
        a = b + 1
    uniq = {}                                   # نفس الباكدج ممكن يظهر في نافذتين (آخر تعديل) → نحتفظ بآخر نسخة
    for p in out:
        uniq[str(p.get("shipmentPackageId", id(p)))] = p
    out = list(uniq.values())
    log(f"[{st.name}] 🧾 باكدجات: {len(out)}")
    return out


def build_order_rows(st, packages, prod_map, up):
    rows, seen = [], {}
    for p in packages:
        on = str(p.get("orderNumber", "")).strip()
        pid = str(p.get("shipmentPackageId", "")).strip()
        od = ms_to_local(p.get("orderDate"))
        for ln in (p.get("lines") or []):
            bc = str(ln.get("barcode", "")).strip()
            try:
                q = max(int(ln.get("quantity") or 1), 1)
            except Exception:
                q = 1
            gross, sdisc, tdisc = num(ln.get("lineGrossAmount")), num(ln.get("lineSellerDiscount")), num(ln.get("lineTyDiscount"))
            # lineGrossAmount وخصومه بتيجي للوحدة (اتأكدنا من أوردر كميته 2: الرقم 139.03 للوحدة والقديم 278.06 للسطر)
            unit = gross - sdisc                      # صافي الوحدة = الإجمالي − خصم التاجر
            inv = unit * q                            # قيمة السطر كله
            base = f"{on}|{pid}|{bc}"
            seen[base] = seen.get(base, 0) + 1
            key = base if seen[base] == 1 else f"{base}#{seen[base]}"
            pm = prod_map.get(bc) or ("", "", "")
            rows.append([key, bc, od, on, sv(ln.get("orderLineItemStatusName") or p.get("status")), str(q),
                         f"{unit:.2f}", f"{inv:.2f}", sv(ln.get("commission")), pm[2],
                         (str(ln.get("productName") or pm[1]))[:120], up,
                         st.name, pid, sv(p.get("status")), f"{gross:.2f}", f"{sdisc:.2f}", f"{tdisc:.2f}",
                         sv(ln.get("vatRate")), sv(p.get("cargoProviderName"))])
    return rows


# ---------------- المنتجات ----------------
def fetch_products(st, log):
    out, page, token = [], 0, None
    hdr = {"Accept-Language": "en"}
    while True:
        params = {"size": 100, "supplierId": st.sid}
        if token:
            params["nextPageToken"] = token
        else:
            params["page"] = page
        j = get_json(st, "/integration/product/sellers/{sid}/products/approved", params, log, extra_headers=hdr)
        c = j.get("content") or []
        out.extend(c)
        if not c:
            break
        tp = int(j.get("totalPages") or 0)
        if token or (page + 1) * 100 >= 10000:
            token = j.get("nextPageToken")
            if not token:
                break
        else:
            page += 1
            if page >= tp:
                break
    log(f"[{st.name}] 🛍️ منتجات: {len(out)}")
    return out


def build_product_rows(st, contents, up):
    rows, pm = [], {}
    for c in contents:
        img = ((c.get("images") or [{}])[0] or {}).get("url", "")
        title = str(c.get("title", ""))
        for v in (c.get("variants") or []):
            bc = str(v.get("barcode", "")).strip()
            if not bc:
                continue
            pr, stk = v.get("price") or {}, v.get("stock") or {}
            pm[bc] = (str(c.get("contentId", "")), title, img)
            rows.append([str(c.get("contentId", "")), bc, title[:120], img, up, st.name, sv(stk.get("quantity")),
                         sv(pr.get("salePrice")), sv(pr.get("listPrice")), sv(v.get("commission")),
                         sv(v.get("onSale")), sv(v.get("stockCode")), sv(c.get("productMainId")), sv(v.get("productUrl"))])
    return rows, pm


# ---------------- المالية ----------------
def _fin_windows(days):
    end = int(time.time() * 1000)
    a = end - days * DAY
    while a < end:
        b = min(a + 14 * DAY, end)
        yield a, b
        a = b + 1


def fetch_finance(st, days, kind, ttype, log):
    path = "/integration/finance/che/sellers/{sid}/" + ("settlements" if kind == "settlements" else "otherfinancials")
    out = []
    for a, b in _fin_windows(days):
        page = 0
        while True:
            j = get_json(st, path, {"startDate": a, "endDate": b, "transactionType": ttype, "page": page,
                                    "size": 1000, "supplierId": st.sid}, log)
            out.extend(j.get("content") or [])
            page += 1
            if page >= int(j.get("totalPages") or 0):
                break
    uniq = {}
    for i in out:
        uniq[str(i.get("id", id(i)))] = i
    return list(uniq.values())


def build_fin_rows(st, ttype, items):
    return [[st.name, ttype, ms_to_local(i.get("transactionDate")), ms_to_local(i.get("orderDate")),
             str(i.get("orderNumber", "") or ""), str(i.get("barcode", "") or ""), str(i.get("shipmentPackageId", "") or ""),
             i.get("commissionRate", ""), i.get("commissionAmount", ""), i.get("sellerRevenue", ""),
             i.get("debt", ""), i.get("credit", ""), str(i.get("paymentOrderId", "") or ""),
             ms_to_local(i.get("paymentDate")), str(i.get("description", "") or "")[:200], str(i.get("receiptId", "") or ""),
             str(i.get("id", "") or ""), str(i.get("transactionType", "") or "")]
            for i in items]


def is_cargo_invoice(i):
    t = str(i.get("transactionType", "") or "") + " " + str(i.get("description", "") or "")
    t = t.replace("İ", "i").replace("I", "i").replace("ı", "i").lower()
    return "kargo" in t or "cargo" in t or "shipping" in t


def fetch_cargo_items(st, inv_id, log):
    """تفاصيل فاتورة شحن: كل طرد/طلب ومبلغ الشحن بتاعه (Id الفاتورة من DeductionInvoices)."""
    out, page = [], 0
    path = "/integration/finance/che/sellers/{sid}/cargo-invoice/" + str(inv_id) + "/items"
    while True:
        j = get_json(st, path, {"page": page, "size": 500}, log)
        out.extend(j.get("content") or [])
        page += 1
        if page >= int(j.get("totalPages") or 0):
            break
    return out


def build_cargo_rows(st, inv, items):
    return [[st.name, str(inv.get("id", "") or ""), ms_to_local(inv.get("transactionDate")),
             str(i.get("orderNumber", "") or ""), str(i.get("parcelUniqueId", "") or ""),
             str(i.get("shipmentPackageType", "") or ""), i.get("amount", ""), i.get("desi", "")] for i in items]


# ---------------- جوجل شيت ----------------
def creds_path():
    ms = glob.glob(os.path.join(GSHEET_CREDS_DIR, "*.json"))
    return ms[0] if ms else None


# ═══════════════ PostgreSQL (Neon) — نفس اختيار تاب Noon API (ملف noon_use_pg.txt): sheet / pg / both لكل البوتات ═══════════════
import json
PG_MODE = "sheet"
PG_LABELS = {"sheet": "📄 Google Sheet فقط", "pg": "🐘 PostgreSQL فقط", "both": "📄+🐘 الاتنين مع بعض"}
PG_URL_FILE = os.path.join(APP_DIR, "pg_url.txt")
PG_FLAG_FILE = os.path.join(APP_DIR, "noon_use_pg.txt")
PG_SCHEMA = "sheet_copy"


def pg_flag_load():
    try:
        v = open(PG_FLAG_FILE, encoding="utf-8").read().strip()
        return {"1": "pg"}.get(v, v) if {"1": "pg"}.get(v, v) in PG_LABELS else "sheet"
    except Exception:
        return "sheet"


def pg_url():
    import re as _re
    u = os.environ.get("PG_URL", "").strip()
    if not u:
        try:
            u = open(PG_URL_FILE, encoding="utf-8").read().strip()
        except Exception:
            u = ""
    u = _re.sub(r"[&?]channel_binding=[^&]*", "", u.strip().strip('"'))
    if u and "?" not in u and "&" in u:
        u = u.replace("&", "?", 1)
    return u


def pg_conn():
    import psycopg2, re as _re
    u = pg_url()
    if not u:
        raise RuntimeError("مفيش رابط Neon — اضغط زر 🔑 رابط Neon")
    m = _re.search(r"postgres(?:ql)?://([^:@/\s]+):([^@\s]+)@([^:/?&\s'\"]+)(?::(\d+))?/([A-Za-z0-9_\-]+)", u)
    if not m:
        raise RuntimeError("رابط Neon مش مفهوم — لازم يبدأ بـ postgresql://user:password@host/dbname")
    return psycopg2.connect(user=m.group(1), password=m.group(2), host=m.group(3), port=int(m.group(4) or 5432),
                            dbname=_re.sub(r"(sslmode|channel_binding)$", "", m.group(5)) or "neondb", sslmode="require", connect_timeout=20)


def pg_side(log, fn):
    """بينفّذ كتابة PostgreSQL حسب الوضع المختار في تاب Noon API. بيرجع True لو الشيت يتخطّى (وضع PostgreSQL فقط)."""
    global PG_MODE
    PG_MODE = pg_flag_load()
    if PG_MODE == "sheet":
        return False
    try:
        fn()
    except Exception as e:
        if PG_MODE == "pg":
            raise
        log(f"⚠️ PostgreSQL فشل: {str(e)[:150]} — كمّلت على الشيت")
    return PG_MODE == "pg"


def _pg_snake(name, used):
    import re as _re
    s = _re.sub(r"[^0-9a-zA-Z]+", "_", str(name)).strip("_").lower()
    if not s:
        s = f"c_{len(used) + 1}"
    if s[0].isdigit():
        s = "c_" + s
    s = s[:60]
    base, i = s, 2
    while s in used:
        s, i = f"{base}_{i}", i + 1
    used.add(s)
    return s


def _pg_s(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def pg_write_tab(tab, headers, rows, log, keys=None, orphan_set=None, scope=None, since=None):
    """بيكتب التاب في PostgreSQL (schema sheet_copy — نفس شكل PG_Copy_All عشان التطبيق يقراه).
    keys → بناء وتحديث بالمفتاح (القديم بيفضل). من غيره → استبدال كامل. كله في transaction واحدة."""
    import io as _io, csv as _csv
    headers = [str(h).strip() for h in headers]
    used = set()
    cols = [_pg_snake(h, used) for h in headers]
    table = _pg_snake(tab, set())
    c = pg_conn()
    try:
        cur = c.cursor()
        cur.execute(f"CREATE SCHEMA IF NOT EXISTS {PG_SCHEMA}")
        cur.execute(f"CREATE TABLE IF NOT EXISTS {PG_SCHEMA}._meta (table_name TEXT PRIMARY KEY, tab_title TEXT, headers JSONB, col_names JSONB, copied_at TIMESTAMPTZ DEFAULT now())")
        cur.execute(f"SELECT headers, col_names FROM {PG_SCHEMA}._meta WHERE table_name=%s", (table,))
        m = cur.fetchone()
        old = []                                   # الصفوف القديمة كـ dict بالهيدر
        if m:
            ohdr, ocols = m
            cur.execute("SELECT " + ", ".join(f'"{x}"' for x in ocols) + f' FROM {PG_SCHEMA}."{table}" ORDER BY _row')
            old = [dict(zip(ohdr, ["" if v is None else v for v in r])) for r in cur.fetchall()]
        if not rows and m:
            log(f"⚠️ {tab}: مفيش بيانات — القاعدة ما اتلمستش")
            return
        new = [{h: _pg_s(v) for h, v in zip(headers, r)} for r in rows]
        if keys and old:
            kf = lambda d: "|".join(str(d.get(k, "")).strip() for k in keys)
            idx = {}
            for i, d in enumerate(old):
                k = kf(d)
                if k.strip("|") and k not in idx:
                    idx[k] = i
            final = [dict(d) for d in old]
            adds = ups = same = orph = 0
            seen = set()
            for d in new:
                k = kf(d)
                if k in idx:
                    seen.add(k)
                    cur_d = final[idx[k]]
                    if any(str(cur_d.get(h, "")).strip() != d[h].strip() for h in headers):
                        cur_d.update(d)
                        ups += 1
                    else:
                        same += 1
                else:
                    final.append(d)
                    adds += 1
            if orphan_set and scope:
                sc_col, sc_vals = scope
                for k, i in idx.items():
                    d = final[i]
                    if k in seen or str(d.get(sc_col, "")).strip() not in sc_vals:
                        continue
                    if since and str(d.get(since[0], "")).strip()[:10] < since[1]:
                        continue
                    if any(str(d.get(cn, "")).strip().lower() != str(x).strip().lower() for cn, x in orphan_set.items()):
                        d.update(orphan_set)
                        orph += 1
            body = final
            msg = f"{adds} جديد · {ups} اتحدّث" + (f" · {orph} اتعلّموا" if orph else "") + f" · {same} زي ما هو"
        else:
            body = new
            msg = f"اتبدّل بالكامل"
        cur.execute(f'DROP TABLE IF EXISTS {PG_SCHEMA}."{table}"')
        cur.execute(f'CREATE TABLE {PG_SCHEMA}."{table}" (' + ", ".join(f'"{x}" TEXT' for x in cols) + ", _row BIGSERIAL)")
        cur.execute(f"INSERT INTO {PG_SCHEMA}._meta (table_name, tab_title, headers, col_names, copied_at) VALUES (%s,%s,%s,%s,now()) "
                    f"ON CONFLICT (table_name) DO UPDATE SET tab_title=EXCLUDED.tab_title, headers=EXCLUDED.headers, col_names=EXCLUDED.col_names, copied_at=now()",
                    (table, tab, json.dumps(headers, ensure_ascii=False), json.dumps(cols)))
        if body:
            buf = _io.StringIO()
            w = _csv.writer(buf)
            for d in body:
                w.writerow([d.get(h, "") for h in headers])
            buf.seek(0)
            cur.copy_expert(f'COPY {PG_SCHEMA}."{table}" (' + ", ".join(f'"{x}"' for x in cols) + ") FROM STDIN WITH CSV", buf)
        c.commit()
        log(f"🐘 {tab}: {msg} — الإجمالي في PostgreSQL {len(body)}")
    except Exception:
        c.rollback()
        raise
    finally:
        c.close()


STORAGE_TAB = "Noon_Storage_API"
STORAGE_REPORT = "fbn_finance_monthlystoragechargereport"
STORAGE_HEADERS = ["Store", "soa_reference_nr", "description_en", "id_partner", "service_month", "pbarcode", "sku", "year_month",
                   "warehouse_code", "country_code", "average_stock_on_hand", "average_pending_qty", "average_net_qty", "longest_side",
                   "median_side", "shortest_side", "weight", "volumetric_weight", "item_volume", "estimated_total_item_volume",
                   "cost_per_cubic_feet", "charged_amount"]


def pg_push(tab, cols, values, key_cols, log):
    """بناء وتحديث بالمفتاح جوه PostgreSQL نفسها (من غير ما نحمّل التاب كله في الذاكرة): الصفوف اللي مفتاحها اتبعت بتتبدّل والباقي زي ما هو."""
    import io as _io, csv as _csv
    headers = [str(h).strip() for h in cols]
    table = _pg_snake(tab, set())
    c = pg_conn()
    try:
        cur = c.cursor()
        cur.execute(f"CREATE SCHEMA IF NOT EXISTS {PG_SCHEMA}")
        cur.execute(f"CREATE TABLE IF NOT EXISTS {PG_SCHEMA}._meta (table_name TEXT PRIMARY KEY, tab_title TEXT, headers JSONB, col_names JSONB, copied_at TIMESTAMPTZ DEFAULT now())")
        cur.execute(f"SELECT headers, col_names FROM {PG_SCHEMA}._meta WHERE table_name=%s", (table,))
        m = cur.fetchone()
        cur.execute("SELECT to_regclass(%s)", (f'{PG_SCHEMA}."{table}"',))
        if not m or cur.fetchone()[0] is None or not key_cols or any(k not in headers for k in key_cols):
            c.close()
            c = None
            return pg_write_tab(tab, cols, values, log)
        ohdr, ocols = [str(h) for h in m[0]], list(m[1])
        used = set(ocols)
        for h in headers:
            if h not in ohdr:
                cn = _pg_snake(h, used)
                cur.execute(f'ALTER TABLE {PG_SCHEMA}."{table}" ADD COLUMN "{cn}" TEXT')
                ohdr.append(h)
                ocols.append(cn)
        cmap = dict(zip(ohdr, ocols))
        ncols = [cmap[h] for h in headers]
        cur.execute("CREATE TEMP TABLE _pgn (" + ", ".join(f'"{x}" TEXT' for x in ncols) + ", _n BIGSERIAL) ON COMMIT DROP")
        buf = _io.StringIO()
        w = _csv.writer(buf)
        for r in values:
            w.writerow([_pg_s(v) for v in r])
        buf.seek(0)
        cur.copy_expert("COPY _pgn (" + ", ".join(f'"{x}"' for x in ncols) + ") FROM STDIN WITH CSV", buf)
        kc = [cmap[k] for k in key_cols]
        ke = lambda a: "(" + " || '|' || ".join(f"COALESCE(btrim({a}.\"{k}\"), '')" for k in kc) + ")"
        cur.execute(f'DELETE FROM {PG_SCHEMA}."{table}" m USING _pgn n WHERE {ke("m")} = {ke("n")} AND btrim({ke("n")}, \'|\') <> \'\'')
        old_n = cur.rowcount
        cl = ", ".join(f'"{x}"' for x in ncols)
        cur.execute(f'INSERT INTO {PG_SCHEMA}."{table}" ({cl}) SELECT {cl} FROM (SELECT *, row_number() OVER (PARTITION BY {ke("_pgn")} ORDER BY _n DESC) rn FROM _pgn) t '
                    f"WHERE rn = 1 OR btrim({ke('t')}, '|') = ''")
        new_n = cur.rowcount
        cur.execute(f"UPDATE {PG_SCHEMA}._meta SET headers=%s, col_names=%s, copied_at=now() WHERE table_name=%s",
                    (json.dumps(ohdr, ensure_ascii=False), json.dumps(ocols), table))
        c.commit()
        cur.execute(f'SELECT count(*) FROM {PG_SCHEMA}."{table}"')
        total = cur.fetchone()[0]
        log(f"🐘 {tab}: اتكتب {new_n:,} صف ({old_n:,} منهم كانوا موجودين واتبدّلوا) — الإجمالي في PostgreSQL {total:,}")
    finally:
        if c is not None:
            c.close()


# ═══════════════════════ كتابة الشيت: بناء وتحديث (Upsert) — عمر ما بيمسح الأول ═══════════════════════
def _su_norm(v):
    t = str(v if v is not None else "").strip()
    if t == "":
        return ""
    if t.lower() in ("true", "false"):                  # TRUE (الشيت) = True (بايثون)
        return t.lower()
    try:
        return repr(round(float(t.replace(",", "")), 6))
    except ValueError:
        return t


def _su_retry(fn, log, tag, tries=6):
    for a in range(tries):
        try:
            return fn()
        except Exception as e:
            if a == tries - 1:
                raise
            log(f"⚠️ {tag}: محاولة {a + 1} فشلت: {str(e)[:120]}")
            _m = str(e)
            time.sleep(65 if ("429" in _m or "Quota" in _m or "quota" in _m) else 3 * (a + 1))


def _su_fill(ws, cols, values, log, tag, vio="RAW", chunk=2000):
    """يكتب العنوان ثم الصفوف على دفعات — أي دفعة تفشل بعد 5 محاولات بترمي خطأ (مفيش صمت)."""
    _su_retry(lambda: ws.resize(rows=max(len(values) + 10, 100), cols=max(len(cols), 6)), log, tag)
    _su_retry(lambda: ws.update(range_name="A1", values=[list(cols)], value_input_option=vio), log, tag)
    for i in range(0, len(values), chunk):
        part = values[i:i + chunk]
        _su_retry(lambda p=part: ws.append_rows(p, value_input_option=vio), log, tag)
        log(f"⬆️ {tag}: {min(i + chunk, len(values))}/{len(values)}")


def _su_swap(ss, ws, tab, cols, values, log, vio):
    """يبني التاب كامل في تاب مؤقت، ولما يكتمل بس بيبدّل القديم — القديم عمره ما بيتمسح قبل ما الجديد يتكتب."""
    import gspread
    tmp_title = (tab + "_new")[:100]
    try:
        ss.del_worksheet(ss.worksheet(tmp_title))
    except gspread.exceptions.WorksheetNotFound:
        pass
    tmp = ss.add_worksheet(title=tmp_title, rows="100", cols=str(max(len(cols), 6)))
    try:
        _su_fill(tmp, cols, values, log, tab, vio)
    except Exception as e:
        try:
            ss.del_worksheet(tmp)
        except Exception:
            pass
        if "limit of" in str(e) and "cells" in str(e):
            raise RuntimeError(f"ملف Google Sheet وصل حد الـ 20 مليون خلية — محتاج {len(values) * len(cols):,} خلية لتاب {tab}. امسح تابات/صفوف قديمة وجرّب تاني") from None
        raise
    ss.del_worksheet(ws)
    tmp.update_title(tab)


def sheet_replace_safe(ss, tab, cols, values, log, vio="RAW"):
    """للتابات اللي بتتبني كلها كل مرة (مخزون/كتالوج/كاش): بناء تاب جديد كامل ثم تبديل — مفيش مسح قبل الكتابة."""
    import gspread
    cols = list(cols)
    if not values:
        log(f"⚠️ {tab}: مفيش بيانات — الشيت ما اتلمسش")
        return
    if pg_side(log, lambda: pg_write_tab(tab, cols, values, log)):
        return
    try:
        ws = ss.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab, rows="100", cols=str(max(len(cols), 6)))
        _su_fill(ws, cols, values, log, tab, vio)
        log(f"🆕 {tab}: اتبنى من الصفر — {len(values)} صف")
        return
    _su_swap(ss, ws, tab, cols, values, log, vio)
    log(f"🔁 {tab}: اتبنى كامل واتبدّل — {len(values)} صف")


def sheet_upsert(ss, tab, cols, values, key_cols, log, orphan_set=None, scope=None, since=None, vio="RAW", cmp_cols=None):
    """يضيف الصفوف الجديدة ويحدّث اللي اتغيّر (بالمفتاح key_cols) من غير ما يمسح حاجة.
    cmp_cols: الأعمدة اللي تغيّرها يعتبر الصف اتحدّث (الافتراضي كلها).
    orphan_set/scope/since: صفوف في الشيت ومش في البيانات الجديدة (في متاجر scope=(عمود, قيم) ومن تاريخ since=(عمود, تاريخ))
    بتتعلّم بالقيم دي بدل ما تتمسح."""
    import gspread
    from gspread.utils import rowcol_to_a1
    if not values:
        log(f"⚠️ {tab}: مفيش بيانات — الشيت ما اتلمسش")
        return
    if pg_side(log, lambda: (pg_write_tab(tab, cols, values, log, keys=key_cols, orphan_set=orphan_set, scope=scope, since=since)
                             if orphan_set else pg_push(tab, cols, values, key_cols, log))):
        return
    cols = list(cols)
    try:
        ws = ss.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        ws = None
    ex = _su_retry(lambda: ws.get_all_values(), log, tab) if ws is not None else []
    hdr = [str(h).strip() for h in ex[0]] if ex else []
    has_data = len(ex) > 1
    if ws is None:
        ws = ss.add_worksheet(title=tab, rows="100", cols=str(max(len(cols), 6)))
        _su_fill(ws, cols, values, log, tab, vio)
        log(f"🆕 {tab}: اتبنى من الصفر — {len(values)} صف")
        return
    if not hdr or not has_data:
        _su_fill(ws, cols, values, log, tab, vio)
        log(f"🆕 {tab}: كان فاضي — اتكتب {len(values)} صف")
        return
    if not all(k in cols and k in hdr for k in key_cols):     # أول مرة بعد إضافة عمود المفتاح: بناء كامل + تبديل
        _su_swap(ss, ws, tab, cols, values, log, vio)
        log(f"🔁 {tab}: اتبنى من جديد بعمود المفتاح — {len(values)} صف")
        return
    new_cols = [c for c in cols if c not in hdr]
    fh = hdr + new_cols
    if new_cols:
        if ws.col_count < len(fh):
            _su_retry(lambda: ws.resize(cols=len(fh)), log, tab)
        _su_retry(lambda: ws.update(range_name="A1", values=[fh], value_input_option="RAW"), log, tab)
    W = len(fh)
    pos = {c: i for i, c in enumerate(fh)}
    kix = [pos[k] for k in key_cols]
    idx = {}
    for i, row in enumerate(ex[1:], start=2):
        row = list(row) + [""] * (W - len(row))
        k = "|".join(str(row[j]).strip() for j in kix)
        if k.strip("|") and k not in idx:
            idx[k] = (i, row)
    ci = {c: j for j, c in enumerate(cols)}
    newmap = {}
    for v in values:
        newmap["|".join(str(v[ci[k]]).strip() for k in key_cols)] = v
    cc = [c for c in (cmp_cols or cols) if c in ci]
    ups, adds, same, orph, seen = [], [], 0, 0, set()
    _dc, _dx = {}, {}
    for k, v in newmap.items():
        if k in idx:
            i, row = idx[k]
            seen.add(k)
            _dif = [c for c in cc if _su_norm(row[pos[c]]) != _su_norm(v[ci[c]])]
            if _dif:
                for c in _dif:                                   # تشخيص: أنهي أعمدة بتفرّق (عشان نعرف ليه صفوف كتير بتتحدّث)
                    _dc[c] = _dc.get(c, 0) + 1
                    if c not in _dx:
                        _dx[c] = (str(row[pos[c]])[:40], str(v[ci[c]])[:40])
                tgt = list(row)
                for c in cols:
                    tgt[pos[c]] = v[ci[c]]
                ups.append((i, tgt))
            else:
                same += 1
        else:
            adds.append([v[ci[c]] if c in ci else "" for c in fh])
    if orphan_set and scope:
        sc_col, sc_vals = scope
        for k, (i, row) in idx.items():
            if k in seen or str(row[pos[sc_col]]).strip() not in sc_vals:
                continue
            if since and str(row[pos[since[0]]]).strip()[:10] < since[1]:
                continue
            if any(_su_norm(row[pos[c]]) != _su_norm(x) for c, x in orphan_set.items() if c in pos):
                tgt = list(row)
                for c, x in orphan_set.items():
                    if c in pos:
                        tgt[pos[c]] = x
                ups.append((i, tgt))
                orph += 1
    last = "".join(ch for ch in rowcol_to_a1(1, W) if not ch.isdigit())
    for s in range(0, len(ups), 500):
        data = [{"range": f"A{i}:{last}{i}", "values": [r]} for i, r in ups[s:s + 500]]
        _su_retry(lambda d=data: ws.batch_update([dict(x) for x in d], value_input_option=vio), log, tab)
    for s in range(0, len(adds), 2000):
        part = adds[s:s + 2000]
        _su_retry(lambda p=part: ws.append_rows(p, value_input_option=vio), log, tab)
    if _dc and len(ups) - orph > 200:                            # كتير اتحدّث: اطبع أكتر الأعمدة اختلافاً مع مثال (الشيت ← الجديد)
        for c, n in sorted(_dc.items(), key=lambda x: -x[1])[:5]:
            log(f"🔎 {tab}: عمود «{c}» اختلف في {n} صف — مثال: «{_dx[c][0]}» ← «{_dx[c][1]}»")
    log(f"⬆️ {tab}: {len(adds)} جديد · {len(ups) - orph} اتحدّث" + (f" · {orph} اتعلّموا" if orph else "") + f" · {same} زي ما هو — الإجمالي في الشيت {len(idx) + len(adds)}")


def write_raw_tab(tab, headers, rows, log, keys=None, cmp_cols=None):
    """keys → بناء وتحديث بالمفتاح. من غيره → بناء تاب جديد كامل ثم تبديل (عمره ما بيمسح قبل ما الجديد يكتمل)."""
    cp = creds_path()
    if not cp:
        raise RuntimeError(f"مفيش ملف اعتماد Google في {GSHEET_CREDS_DIR}")
    ss = gspread.authorize(ServiceAccountCredentials.from_json_keyfile_name(cp, GSHEET_SCOPE)).open(GSHEET_NAME)
    if not rows:                                   # لو فاضي: نتأكد إن التاب موجود بعنوانه بس (من غير مسح)
        try:
            ss.worksheet(tab)
        except gspread.exceptions.WorksheetNotFound:
            ws = ss.add_worksheet(title=tab, rows="100", cols=str(max(len(headers), 6)))
            ws.update(range_name="A1", values=[list(headers)], value_input_option="USER_ENTERED")
        return
    if keys:
        sheet_upsert(ss, tab, headers, rows, keys, log, vio="USER_ENTERED", cmp_cols=cmp_cols)
    else:
        sheet_replace_safe(ss, tab, headers, rows, log, vio="USER_ENTERED")
    log(f"✅ اتكتب في تاب {tab} ({len(rows)} صف)")


def save_xlsx(path, sheets):
    """sheets: [(اسم الشيت، العناوين، الصفوف)] → ملف إكسيل واحد بشيتات."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title, hdr, rows in sheets:
        ws = wb.create_sheet(title[:31])
        ws.append(hdr)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3A5F")
        for r in rows:
            ws.append(r)
        ws.freeze_panes = "A2"
        if rows:
            ws.auto_filter.ref = ws.dimensions
    wb.save(path)


# ---------------- الواجهة ----------------
class App(tk.Frame):
    def __init__(self, root):
        super().__init__(root, bg="#d9d9d9")
        self.pack(fill="both", expand=True, padx=8, pady=8)
        self.stores = []
        self.running = False
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="➕ ضيف متجر تريندول", command=self.add_store).grid(row=0, column=0, padx=3)
        ttk.Button(top, text="🗑️ شيل المحدد", command=self.rm_store).grid(row=0, column=1, padx=3)
        ttk.Button(top, text="🔌 اختبر الاتصال", command=self.test).grid(row=0, column=2, padx=3)
        self.lb = tk.Listbox(top, height=4, selectmode="extended")
        self.lb.grid(row=1, column=0, columnspan=8, sticky="ew", pady=4)
        top.columnconfigure(7, weight=1)
        ttk.Label(top, text="آخر كام يوم:").grid(row=2, column=0, sticky="e")
        self.days = tk.IntVar(value=30)
        ttk.Spinbox(top, from_=1, to=60, width=5, textvariable=self.days).grid(row=2, column=1, sticky="w")
        self.o_orders = tk.BooleanVar(value=True)
        self.o_prod = tk.BooleanVar(value=True)
        self.o_fin = tk.BooleanVar(value=True)
        self.o_oth = tk.BooleanVar(value=True)
        self.o_crg = tk.BooleanVar(value=True)
        self.write = tk.BooleanVar(value=True)
        self.xl = tk.BooleanVar(value=True)
        self.folder = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Downloads"))
        for i, (v, t) in enumerate([(self.o_orders, f"الأوردرات → {ORD_TAB}"), (self.o_prod, f"المنتجات (مخزون/أسعار/صور) → {PRD_TAB}"),
                                    (self.o_fin, f"المبيعات والمرتجعات (عمولة) → {FIN_TAB}"),
                                    (self.o_oth, f"فواتير العمولة والخصومات → {OTH_TAB}"),
                                    (self.o_crg, f"شحن لكل طلب (من فواتير الشحن) → {CRG_TAB} — محتاج خيار الفواتير شغّال"),
                                    (self.write, "اكتب في جوجل شيت"),
                                    (self.xl, "احفظ كله في ملف إكسيل واحد (شيتات) — عشان ترفعه")]):
            ttk.Checkbutton(top, text=t, variable=v).grid(row=3 + i, column=0, columnspan=6, sticky="w")
        ttk.Label(top, text="مكان الإكسيل:").grid(row=10, column=0, sticky="e")
        ttk.Entry(top, textvariable=self.folder, width=55).grid(row=10, column=1, columnspan=5, sticky="w")
        self.btn = ttk.Button(top, text="▶ شغّل (كل المتاجر مع بعض)", command=self.go)
        self.btn.grid(row=11, column=0, columnspan=3, pady=6, sticky="w")
        self.auto = tk.BooleanVar(value=False)
        self.every = tk.IntVar(value=60)
        self.next_run = None
        self.auto_run = False
        ttk.Checkbutton(top, text="⏰ تشغيل تلقائي كل", variable=self.auto, command=self.toggle_auto).grid(row=12, column=0, columnspan=2, sticky="w")
        ttk.Spinbox(top, from_=10, to=1440, increment=10, width=6, textvariable=self.every).grid(row=12, column=2, sticky="w")
        ttk.Label(top, text="دقيقة").grid(row=12, column=3, sticky="w")
        self.next_lbl = ttk.Label(top, text="")
        self.next_lbl.grid(row=12, column=4, columnspan=3, sticky="w")
        self.times = tk.StringVar(value="")
        self._last_fixed = ""
        ttk.Label(top, text="🕐 مواعيد محددة (HH:MM مفصولة بفاصلة):").grid(row=13, column=0, columnspan=2, sticky="e")
        ttk.Entry(top, textvariable=self.times, width=34).grid(row=13, column=2, columnspan=4, sticky="w")
        self.out = scrolledtext.ScrolledText(self, font=("Consolas", 10), height=18)
        self.out.pack(fill="both", expand=True, padx=8, pady=8)
        self.load()
        self.load_settings()
        for _v in (self.o_orders, self.o_prod, self.o_fin, self.o_oth, self.o_crg, self.write, self.xl):
            _v.trace_add("write", lambda *a: self.save_settings())
        self.folder.trace_add("write", lambda *a: self.save_settings())
        self.times.trace_add("write", lambda *a: self.save_settings())
        self.after(5000, self.tick)

    def log(self, m):
        try:
            if LOG_HOOK:
                LOG_HOOK(str(m) + "\n")
        except Exception:
            pass
        try:
            self.after(0, lambda: (self.out.insert("end", m + "\n"), self.out.see("end")))
        except Exception:
            pass

    def load(self):
        try:
            data = json.load(open(STORES_FILE, encoding="utf-8"))
        except Exception:
            data = []
        for d in data:
            try:
                self._add(TStore(d), save=False)
            except Exception as e:
                self.log(f"⚠️ متجر ناقص بيانات: {e}")

    def save(self):
        json.dump([s.as_dict() for s in self.stores], open(STORES_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    def _add(self, st, save=True):
        self.stores.append(st)
        self.lb.insert("end", f"{st.name}  —  seller {st.sid}  —  {st.sf}")
        if save:
            self.save()

    def add_store(self):
        w = tk.Toplevel(self)
        w.title("متجر تريندول")
        fields = {}
        for i, (k, t, show) in enumerate([("name", "اسم المتجر (للعرض)", ""), ("seller_id", "Seller ID", ""),
                                          ("api_key", "API Key", ""), ("api_secret", "API Secret", "•"),
                                          ("store_front", "storeFrontCode (SA للسعودية)", "")]):
            ttk.Label(w, text=t).grid(row=i, column=0, padx=6, pady=3, sticky="e")
            e = ttk.Entry(w, width=48, show=show)
            e.grid(row=i, column=1, padx=6, pady=3)
            fields[k] = e
        fields["store_front"].insert(0, "SA")

        def ok():
            d = {k: e.get().strip() for k, e in fields.items()}
            if not (d["seller_id"] and d["api_key"] and d["api_secret"]):
                messagebox.showwarning("", "Seller ID و API Key و API Secret لازم", parent=w)
                return
            self._add(TStore(d))
            w.destroy()
        ttk.Button(w, text="حفظ", command=ok).grid(row=5, column=1, pady=8, sticky="e")

    def rm_store(self):
        for i in reversed(self.lb.curselection()):
            self.lb.delete(i)
            del self.stores[i]
        self.save()

    def test(self):
        if not self.stores:
            messagebox.showwarning("", "ضيف متجر الأول")
            return

        def run():
            for st in self.stores:
                try:
                    now = int(time.time() * 1000)
                    j = get_json(st, "/integration/order/sellers/{sid}/v2/orders",
                                 {"startDate": now - 3 * DAY, "endDate": now, "page": 0, "size": 1, "supplierId": st.sid},
                                 self.log, tries=2)
                    c = (j.get("content") or [{}])[0]
                    ln = ((c.get("lines") or [{}])[0]) if c else {}
                    self.log(f"[{st.name}] ✅ الاتصال تمام — باكدجات في آخر 3 أيام: {j.get('totalElements')}")
                    if c:
                        keep = ("lineUnitPrice", "lineGrossAmount", "lineTotalDiscount", "lineSellerDiscount", "lineTyDiscount",
                                "quantity", "commission", "vatRate", "orderLineItemStatusName")
                        self.log(f"[{st.name}] 🔎 مفاتيح الباكدج: {sorted(k for k in c.keys() if 'customer' not in k.lower() and 'address' not in k.lower() and 'identity' not in k.lower())}")
                        self.log(f"[{st.name}] 🔎 قيم سطر (أرقام بس): { {k: ln.get(k) for k in keep} }")
                except Exception as e:
                    self.log(f"[{st.name}] ❌ {e}")
        threading.Thread(target=run, daemon=True).start()

    # ---------- الإعدادات + الجدولة ----------
    def load_settings(self):
        try:
            d = json.load(open(SETTINGS_FILE, encoding="utf-8"))
            self.days.set(int(d.get("days", 30)))
            self.every.set(int(d.get("every", 60)))
            self.auto.set(bool(d.get("auto", False)))
            for _n in ("o_orders", "o_prod", "o_fin", "o_oth", "o_crg", "write", "xl"):
                if _n in d:
                    getattr(self, _n).set(bool(d[_n]))
            if d.get("folder"):
                self.folder.set(str(d["folder"]))
            self.times.set(str(d.get("times", "")))
            if self.auto.get():
                self.next_run = time.time() + 60
                self.log("⏰ التشغيل التلقائي مفعّل (من آخر مرة) — أول تشغيل بعد دقيقة")
        except Exception:
            pass

    def save_settings(self):
        try:
            d = {"days": int(self.days.get()), "every": int(self.every.get()), "auto": bool(self.auto.get()),
                 "folder": self.folder.get(), "times": self.times.get()}
            for _n in ("o_orders", "o_prod", "o_fin", "o_oth", "o_crg", "write", "xl"):
                d[_n] = bool(getattr(self, _n).get())
            json.dump(d, open(SETTINGS_FILE, "w", encoding="utf-8"))
        except Exception:
            pass

    def toggle_auto(self):
        if self.auto.get():
            self.next_run = time.time() + int(self.every.get()) * 60
            self.log(f"⏰ تشغيل تلقائي كل {int(self.every.get())} دقيقة (الجاي: {time.strftime('%H:%M', time.localtime(self.next_run))})")
        else:
            self.next_run = None
            self.log("⏰ التشغيل التلقائي اتقفل")
        self.save_settings()

    def fixed_times(self):
        out = []
        for t in str(self.times.get()).replace("،", ",").replace(" ", ",").split(","):
            t = t.strip()
            try:
                out.append(dt.datetime.strptime(t, "%H:%M").strftime("%H:%M")) if t else None
            except ValueError:
                pass
        return out

    def tick(self):
        try:
            now = time.strftime("%H:%M")
            key = time.strftime("%Y-%m-%d ") + now
            if now in self.fixed_times() and key != self._last_fixed and not self.running:
                self._last_fixed = key
                self.log(f"🕐 موعد محدد {now} — تشغيل تلقائي...")
                self.go(auto=True)
        except Exception:
            pass
        try:
            if self.auto.get() and self.next_run:
                self.next_lbl.configure(text=f"التشغيل الجاي: {time.strftime('%H:%M', time.localtime(self.next_run))}")
                if self.next_run - time.time() <= 0 and not self.running:
                    self.log("⏰ تشغيل تلقائي...")
                    self.go(auto=True)
            else:
                self.next_lbl.configure(text="")
        except Exception:
            pass
        self.after(15000, self.tick)

    def shutdown(self):
        self.auto.set(False)

    def go(self, auto=False):
        if self.running:
            return
        if not self.stores:
            if not auto:
                messagebox.showwarning("", "ضيف متجر الأول")
            else:
                self.log("⚠️ مفيش متاجر تريندول — التشغيل التلقائي اتخطّى")
                self.next_run = time.time() + int(self.every.get()) * 60
            return
        self.auto_run = auto
        self.save_settings()
        self.running = True
        self.btn.state(["disabled"])
        threading.Thread(target=self.job, daemon=True).start()

    def job(self):
        try:
            t0 = time.time()
            days = int(self.days.get())
            up = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.log(f"📅 آخر {days} يوم — {len(self.stores)} متجر")

            def one(st):
                res = {"ord": [], "prd": [], "fin": [], "oth": [], "crg": []}
                try:
                    pm = {}
                    if self.o_prod.get() or self.o_orders.get():
                        prows, pm = build_product_rows(st, fetch_products(st, self.log), up)
                        res["prd"] = prows
                    if self.o_orders.get():
                        res["ord"] = build_order_rows(st, fetch_orders(st, days, self.log), pm, up)
                    if self.o_fin.get():
                        for t in ("Sale", "Return"):
                            items = fetch_finance(st, days, "settlements", t, self.log)
                            self.log(f"[{st.name}] 💰 {t}: {len(items)} سطر")
                            res["fin"] += build_fin_rows(st, t, items)
                    if self.o_oth.get():
                        for t in OTHER_TYPES:
                            items = fetch_finance(st, days, "other", t, self.log)
                            self.log(f"[{st.name}] 🧾 {t}: {len(items)} سطر")
                            res["oth"] += build_fin_rows(st, t, items)
                            if t == "DeductionInvoices" and self.o_crg.get():
                                kinds = {}
                                for i in items:
                                    k = str(i.get("transactionType", "") or "—")
                                    kinds[k] = kinds.get(k, 0) + 1
                                self.log(f"[{st.name}] 🔎 أنواع الخصومات: " + " | ".join(f"{k}: {n}" for k, n in kinds.items()))
                                cinv = [i for i in items if is_cargo_invoice(i)]
                                self.log(f"[{st.name}] 🚚 فواتير شحن لاقيتها: {len(cinv)}")
                                for inv in cinv:
                                    try:
                                        ci = fetch_cargo_items(st, inv.get("id"), self.log)
                                        res["crg"] += build_cargo_rows(st, inv, ci)
                                        self.log(f"[{st.name}] 🚚 فاتورة {inv.get('id')}: {len(ci)} طرد")
                                    except Exception as e:
                                        self.log(f"[{st.name}] ⚠️ فاتورة شحن {inv.get('id')}: {str(e)[:150]}")
                except Exception as e:
                    self.log(f"[{st.name}] ❌ {e}")
                return res
            with ThreadPoolExecutor(max_workers=len(self.stores)) as ex:
                results = list(ex.map(one, self.stores))
            xl_sheets = []
            for key, tab, hdr, on in (("ord", ORD_TAB, ORD_HEADERS, self.o_orders.get()), ("prd", PRD_TAB, PRD_HEADERS, self.o_prod.get()),
                                      ("fin", FIN_TAB, FIN_HEADERS, self.o_fin.get()), ("oth", OTH_TAB, FIN_HEADERS, self.o_oth.get()),
                                      ("crg", CRG_TAB, CRG_HEADERS, self.o_oth.get() and self.o_crg.get())):
                if not on:
                    continue
                rows = [r for res in results for r in res[key]]
                if key == "ord":
                    rows.sort(key=lambda r: r[2], reverse=True)
                self.log(f"📊 {tab}: {len(rows)} صف")
                xl_sheets.append((tab, hdr, rows))
                if self.write.get() and rows:
                    if key == "ord":
                        write_raw_tab(tab, hdr, rows, self.log, keys=("Row Key",), cmp_cols=[h for h in hdr if h != "Date Uploaded"])
                    else:
                        write_raw_tab(tab, hdr, rows, self.log)
            if self.xl.get() and xl_sheets and not self.auto_run:   # التشغيل التلقائي ما بيعملش إكسيل كل مرة
                if openpyxl is None:
                    self.log("⚠️ الإكسيل محتاج:  py -m pip install openpyxl")
                else:
                    folder = self.folder.get().strip() or os.getcwd()
                    os.makedirs(folder, exist_ok=True)
                    path = os.path.join(folder, f"Trendyol_API_{dt.datetime.now():%Y%m%d_%H%M}.xlsx")
                    save_xlsx(path, xl_sheets)
                    self.log(f"✅ الإكسيل جاهز: {path}")
            self.log(f"⏱️ خلص في {int(time.time() - t0)} ثانية")
        except Exception as e:
            self.log(f"❌ {type(e).__name__}: {e}")
        finally:
            self.running = False
            if self.auto.get():
                self.next_run = time.time() + max(10, int(self.every.get())) * 60
            self.after(0, lambda: self.btn.state(["!disabled"]))


def build_ui(root):
    """بيتنادى من البرنامج الموحد: بيبني الواجهة جوه التاب ويرجّع دالة إيقاف."""
    return App(root).shutdown

'''

_EMBEDDED['NoonAccounts'] = r'''# -*- coding: utf-8 -*-
"""
Noon_Accounts.py — دفتر حسابات نون (بوت منفصل، متعدد المتاجر): بيسحب التقرير المالي (على مستوى الصنف) وأوردرات نون من تاريخ البداية
وبيحدّث بالمفتاح (Upsert): الصف المتغيّر يتحدّث والجديد يتضاف، من غير تكرار. التحديث اليومي بيرجع بفترة أمان (lookback) عشان
الأوردرات اللي حالتها اتغيّرت والكشوف الأسبوعية اللي بتنزل متأخر.

المخرجات (جنب الأداة):
  noon_accounts/<المتجر>/finance.csv   — دفتر المالية (مفتاح: Reference Nr + Order Nr + Item Nr + Transaction Type)
  noon_accounts/<المتجر>/orders.csv    — دفتر الأوردرات (مفتاح: item_nr) + أول ظهور + تاريخ آخر تغيير حالة
  noon_accounts/state.json             — آخر تحديث لكل متجر
  + (اختياري) تابين في الشيت: Noon_Finance_API و Noon_Orders_API (كل المتاجر، عمود Store) — بيتبنوا ويتحدّثوا بالمفتاح (مفيش مسح)

التشغيل:  py Noon_Accounts.py            (واجهة)
          py Noon_Accounts.py --auto     (بدون واجهة: تحديث يومي لكل المتاجر المحفوظة — للجدولة)
المتطلبات: py -m pip install requests pyjwt cryptography gspread oauth2client
"""
import csv, io, json, os, sys, glob, time, uuid, zipfile, threading, datetime as dt
import jwt
import requests

BASE = "https://noon-api-gateway.noon.partners"
UA = {"User-Agent": "NoonApiClient/1.0"}
APP_DIR = os.environ.get("PLATFORMS_BASE_DIR") or os.path.dirname(os.path.abspath(__file__))
EMBEDDED = bool(os.environ.get("PLATFORMS_BASE_DIR"))
LOG_HOOK = None   # بيتحط من البرنامج الموحد → السجل المشترك
EXPORT_LOG = os.path.join(APP_DIR, "noon_accounts_exports_log.json")
_elock = threading.Lock()


def read_key_file(path):
    with open(path, encoding="utf-8-sig") as fh:
        raw = json.load(fh)

    def find(o):
        if isinstance(o, dict):
            if ("key_id" in o or "keyId" in o) and ("private_key" in o or "privateKey" in o):
                return o
            for v in o.values():
                r = find(v)
                if r:
                    return r
        if isinstance(o, list):
            for v in o:
                r = find(v)
                if r:
                    return r
        return None
    k = find(raw)
    if not k:
        raise ValueError("مالقيتش key_id / private_key في الملف")
    k = dict(k)
    for a, b in (("key_id", "keyId"), ("private_key", "privateKey"), ("project_code", "projectCode")):
        if not k.get(a) and k.get(b):
            k[a] = k[b]
    if not k.get("project_code") and k.get("default_project_code"):
        k["project_code"] = k["default_project_code"]
    miss = [x for x in ("key_id", "private_key", "project_code") if not k.get(x)]
    if miss:
        raise ValueError("ناقص: " + ", ".join(miss))
    return k


class Store:
    def __init__(self, path):
        self.path = path
        self.name = os.path.splitext(os.path.basename(path))[0]
        self.key = read_key_file(path)
        self.sess = None
        self.lock = threading.Lock()


def login(st, log):
    s = requests.Session()
    s.headers.update(UA)
    try:   # اتصالات أكتر مع بعض (الافتراضي 10 بس وده كان بيخنق الطلبات المتوازية)
        from requests.adapters import HTTPAdapter
        s.mount('https://', HTTPAdapter(pool_connections=48, pool_maxsize=48))
    except Exception:
        pass
    tok = jwt.encode({"sub": st.key["key_id"], "iat": int(time.time()), "jti": str(uuid.uuid4())},
                     st.key["private_key"], algorithm="RS256")
    r = s.post(f"{BASE}/identity/public/v1/api/login",
               json={"token": tok, "default_project_code": st.key["project_code"]}, timeout=30)
    if r.status_code != 200:
        log(f"[{st.name}] ❌ تسجيل الدخول فشل: {r.status_code}")
        return None
    return s


def api_post(st, path, body, log):
    for attempt in (1, 2):
        with st.lock:
            if st.sess is None:
                st.sess = login(st, log)
            s = st.sess
        if s is None:
            return None
        r = None
        for t in (1, 2, 3):     # نون أحيانًا بتتأخر في الرد — نعيد المحاولة بدل ما نوقّف التشغيلة كلها
            try:
                r = s.post(BASE + path, json=body, headers={"Accept": "application/json"}, timeout=60)
                break
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
                log(f"[{st.name}] ⚠️ نون ماردتش في الوقت ({type(e).__name__}) — محاولة {t}/3")
                if t == 3:
                    log(f"[{st.name}] ❌ {path}: فشل بعد 3 محاولات")
                    return None
                time.sleep(5 * t)
        if r.status_code == 401 and attempt == 1:
            with st.lock:
                st.sess = None
            continue
        return r
    return None


def log_export(st, code, ec):
    with _elock:
        try:
            d = json.load(open(EXPORT_LOG, encoding="utf-8"))
        except Exception:
            d = []
        d.append({"store": st.name, "code": code, "ec": ec, "ts": time.time()})
        d = d[-200:]
        try:
            json.dump(d, open(EXPORT_LOG, "w", encoding="utf-8"), ensure_ascii=False)
        except Exception:
            pass


def recent_exports(st, code, hours=36):
    """أكواد التصديرات اللي الأداة طلبتها قبل كده (الأحدث الأول)، خلال آخر X ساعة."""
    try:
        d = json.load(open(EXPORT_LOG, encoding="utf-8"))
    except Exception:
        return []
    out = [x["ec"] for x in d if x.get("store") == st.name and x.get("code") == code
           and time.time() - x.get("ts", 0) < hours * 3600]
    return list(reversed(out))


def export_state(st, ec, log):
    """(حالة, رابط) لتصدير موجود — من غير ما يطلب جديد."""
    rs = api_post(st, "/impex/v1/export/status", {"export_code": ec}, log)
    if rs is None:
        return "", None
    try:
        sd = rs.json()
    except Exception:
        return "", None
    return str(sd.get("export_status", "")).upper(), sd.get("download_url")


def run_export(st, code, params, log, stop, max_wait=40):
    """يطلب تقرير، يستنى، ينزّله، ويرجع صفوف dict (من الذاكرة — من غير ما يحفظ ملف)."""
    for wait_i in range(max_wait):
        r = api_post(st, "/impex/v1/export/create", {"export_category_code": code, "params": params}, log)
        if r is None:
            return None
        try:
            d = r.json()
        except Exception:
            d = {}
        if r.status_code == 400 and "already running" in (r.text or "").lower() and not stop.is_set():
            if wait_i == max_wait - 1:
                log(f"[{st.name}] ⌛ {code}: لسه في تصدير شغال — هكمل من غيره")
                return None
            if wait_i == 0:
                log(f"[{st.name}] ⏳ {code}: في تصدير تاني من نفس النوع شغال عند نون — هستنى وأعيد المحاولة...")
                for old in recent_exports(st, code, 48):
                    stt, url = export_state(st, old, log)
                    if url or (stt and not any(w in stt for w in ("FAIL", "ERROR", "CANCEL"))):
                        log(f"[{st.name}] ↪️ ده غالباً تصدير قديم من عندنا ({old}) — هاخده بدل ما أطلب جديد")
                        return fetch_export(st, code, old, log, stop)
            time.sleep(20)
            continue
        if r.status_code == 400 and "ratelimit" in (r.text or "").lower() and not stop.is_set() and wait_i < 6:
            log(f"[{st.name}] ⏳ {code}: نون قالت Ratelimit — هستنى دقيقة وأعيد ({wait_i + 1}/6)")
            for _ in range(60):
                if stop.is_set():
                    break
                time.sleep(1)
            continue
        break
    if r.status_code != 200 or not d.get("export_code"):
        log(f"[{st.name}] ❌ طلب {code}: {r.status_code} {(r.text or '')[:300]}")
        return None
    ec = d["export_code"]
    log_export(st, code, ec)
    log(f"[{st.name}] ⏳ {code} اتطلب ({ec})")
    return fetch_export(st, code, ec, log, stop)


def fetch_export(st, code, ec, log, stop):
    """ينتظر export_code موجود ويحمّله (من غير ما يطلب تصدير جديد)."""
    url = None
    for _ in range(180):
        if stop.is_set():
            return None
        time.sleep(5)
        rs = api_post(st, "/impex/v1/export/status", {"export_code": ec}, log)
        if rs is None:
            continue
        try:
            sd = rs.json()
        except Exception:
            continue
        if sd.get("download_url"):
            url = sd["download_url"]
            break
        if any(w in str(sd.get("export_status")).upper() for w in ("FAIL", "ERROR", "CANCEL")):
            log(f"[{st.name}] ❌ {code}: {sd}")
            return None
    if not url:
        log(f"[{st.name}] ⌛ {code} ماجهزش في الوقت المحدد")
        return None
    rr = requests.get(url, timeout=600)
    if rr.status_code in (401, 403) and st.sess is not None:
        rr = st.sess.get(url, timeout=600)
    if rr.status_code != 200:
        log(f"[{st.name}] ❌ فشل التحميل {rr.status_code}")
        return None
    data = rr.content
    if data[:2] == b"PK":
        z = zipfile.ZipFile(io.BytesIO(data))
        names = [n for n in z.namelist() if n.lower().endswith(".csv")]
        if not names:
            log(f"[{st.name}] ⚠️ الملف مضغوط ومفيهوش CSV")
            return None
        data = z.read(names[0])
    rows = list(csv.DictReader(io.StringIO(data.decode("utf-8-sig", errors="replace"))))
    log(f"[{st.name}] ✅ {code}: {len(rows)} صف")
    return rows


# ═══════════════════════ تحليل الأعمدة والمفاتيح ═══════════════════════


# ═══════════════════════ الدفتر (Upsert) ═══════════════════════
FIN_CODE = "noon_financeweb_transactionviewreportonitemlevel"
ORD_CODE = "noon_noonoms_ordersexport"
DEFAULT_START = dt.date(2026, 9, 1)
STEP_DAYS = 15
FIN_LOOKBACK = 21      # يوم — الكشوف الأسبوعية + التعديلات المتأخرة
ORD_LOOKBACK = 30      # يوم — تغيّر حالة الأوردر (شحن → تسليم → مرتجع)
ORD_BACKFILL = 60      # يوم — الأوردرات بتتسحب من (بداية الدفتر − 60 يوم) عشان رسوم الأوردرات القديمة تلاقي منتجاتها
LEDGER_DIR = os.path.join(APP_DIR, "noon_accounts")
STORES_FILE = os.path.join(APP_DIR, "noon_stores.json" if EMBEDDED else "noon_accounts_stores.json")   # في البرنامج الموحد: نفس متاجر Noon API
SETTINGS_FILE = os.path.join(APP_DIR, "noon_accounts_settings.json")
STATE_FILE = os.path.join(LEDGER_DIR, "state.json")
GSHEET_CREDS_DIR = r"\\WIN-181U7G4HFM5\Share For Markting\NOONIMG\Copy"
GSHEET_NAME = "Complaints"
GSHEET_SCOPE = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
TAB_FIN, TAB_ORD = "Noon_Finance_API", "Noon_Orders_API"
PERSONAL = ("name", "phone", "mobile", "email", "address", "customer", "city", "area", "street", "zip", "building", "landmark", "recipient")
ORD_DROP = ("partner_name",)  # أعمدة مش محتاجينها
FIN_KEY = ("Reference Nr", "Order Nr", "Item Nr", "Transaction Type")
FIN_NUM = ("Net Proceeds", "Referral Fee including VAT", "Fullfilment & Logistics Fees including VAT", "Shipping Credits including VAT",
           "Other Order Fees including VAT", "Order Subsidies including VAT", "Non-Order Fees including VAT",
           "Non-Order Subsidies including VAT", "Others including VAT", "Total")
_slock = threading.Lock()


def is_personal(col):
    c = str(col).lower()
    return any(w in c for w in PERSONAL)


def fin_key(r):
    return "|".join(str(r.get(k, "") or "").strip() for k in FIN_KEY)


def ord_key(r):
    return str(r.get("item_nr", "") or "").strip() or ("%s|%s" % (r.get("order_nr", ""), r.get("sku", "")))


def read_ledger(path):
    if not os.path.exists(path):
        return [], []
    with open(path, encoding="utf-8-sig", newline="") as fh:
        rd = csv.DictReader(fh)
        cols = list(rd.fieldnames or [])
        return list(rd), cols


def load_dict(path, keyfn):
    rows, cols = read_ledger(path)
    return {keyfn(r): r for r in rows}, cols


def write_ledger(path, rows, cols):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    os.replace(tmp, path)


def merge_cols(old_cols, rows):
    cols = list(old_cols)
    for r in rows:
        for k in r.keys():
            if k not in cols:
                cols.append(k)
    return cols


def date_chunks(d_from, d_to, step):
    a = d_from
    while a <= d_to:
        b = min(a + dt.timedelta(days=step - 1), d_to)
        yield a, b
        a = b + dt.timedelta(days=1)


def pull_range(st, code, d_from, d_to, step, log, stop):
    """يرجع (صفوف, أجزاء فاشلة). الجزء الفاشل مابيتحسبش — وبيتعاد في التحديث الجاي."""
    allrows, failed = [], []
    for a, b in date_chunks(d_from, d_to, step):
        if stop.is_set():
            failed.append((a, b))
            continue
        rows = run_export(st, code, {"from_date": a.isoformat(), "to_date": b.isoformat()}, log, stop)
        if rows is None:
            failed.append((a, b))
            log(f"[{st.name}] ⚠️ الجزء {a} → {b} فشل — هيتعاد في التحديث الجاي")
            continue
        allrows += rows
    return allrows, failed


def load_state():
    try:
        return json.load(open(STATE_FILE, encoding="utf-8"))
    except Exception:
        return {}


def save_state(s):
    os.makedirs(LEDGER_DIR, exist_ok=True)
    with _slock:
        json.dump(s, open(STATE_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def update_finance(st, d_from, d_to, log, stop, step=STEP_DAYS):
    """يسحب المالية للفترة ويعمل Upsert. يرجع (جديد, اتحدّث, عدد الأجزاء الفاشلة)."""
    path = os.path.join(LEDGER_DIR, st.name, "finance.csv")
    old, cols = load_dict(path, fin_key)
    rows, failed = pull_range(st, FIN_CODE, d_from, d_to, step, log, stop)
    new = upd = 0
    for r in rows:
        r = {k: v for k, v in r.items() if not is_personal(k)}
        k = fin_key(r)
        if k in old:
            if any(str(old[k].get(c, "")) != str(r.get(c, "")) for c in r.keys()):
                upd += 1
            old[k].update(r)
        else:
            r["_first_seen"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
            old[k] = r
            new += 1
    cols = merge_cols(cols, list(old.values()))
    if "_first_seen" not in cols:
        cols.append("_first_seen")
    write_ledger(path, list(old.values()), cols)
    log(f"[{st.name}] 💰 المالية: {new} جديد، {upd} اتحدّث، الإجمالي {len(old)} صف، أجزاء فاشلة {len(failed)}")
    return new, upd, failed


def load_orders_handoff(st, max_age_h=12):
    """أوردرات خطوة «Noon API» (نفس التصدير) لو لسه طازة — بنستخدمها بدل ما نسحبها تاني. بترجع (من, إلى, صفوف) أو None."""
    try:
        d = json.load(open(os.path.join(APP_DIR, "noon_orders_handoff", f"{st.name}.json"), encoding="utf-8"))
        if time.time() - float(d.get("ts", 0)) > max_age_h * 3600:
            return None
        return dt.date.fromisoformat(d["from"]), dt.date.fromisoformat(d["to"]), d["rows"]
    except Exception:
        return None


def update_orders(st, d_from, d_to, log, stop, step=STEP_DAYS):
    """الأوردرات: Upsert بـ item_nr، مع تتبّع أول ظهور وتغيّر الحالة."""
    path = os.path.join(LEDGER_DIR, st.name, "orders.csv")
    old, cols = load_dict(path, ord_key)
    hand = load_orders_handoff(st)
    if hand and hand[0] <= d_to and hand[1] >= d_from:
        h_from, h_to, rows = hand
        failed = []
        log(f"[{st.name}] ♻️ بستخدم أوردرات خطوة Noon API ({len(rows)} صف من {h_from} إلى {h_to}) بدل ما أسحبها تاني")
        if d_from < h_from:                       # الأقدم من اللي سحبته خطوة Noon API
            r2, f2 = pull_range(st, ORD_CODE, d_from, h_from - dt.timedelta(days=1), step, log, stop)
            rows, failed = rows + r2, failed + f2
        if h_to < d_to:                           # آخر يوم أو اتنين
            r3, f3 = pull_range(st, ORD_CODE, h_to + dt.timedelta(days=1), d_to, step, log, stop)
            rows, failed = rows + r3, failed + f3
    else:
        rows, failed = pull_range(st, ORD_CODE, d_from, d_to, step, log, stop)
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    new = chg = 0
    for r in rows:
        r = {k: v for k, v in r.items() if not is_personal(k) and k not in ORD_DROP}
        k = ord_key(r)
        o = old.get(k)
        if o is None:
            r["_first_seen"] = now
            r["_prev_status"] = ""
            r["_status_changed_at"] = now
            old[k] = r
            new += 1
        else:
            if str(o.get("item_status", "")) != str(r.get("item_status", "")):
                o["_prev_status"] = o.get("item_status", "")
                o["_status_changed_at"] = now
                chg += 1
            o.update(r)
    cols = merge_cols(cols, list(old.values()))
    for c in ("_first_seen", "_prev_status", "_status_changed_at"):
        if c not in cols:
            cols.append(c)
    write_ledger(path, list(old.values()), cols)
    log(f"[{st.name}] 🧾 الأوردرات: {new} جديد، {chg} تغيّرت حالتهم، الإجمالي {len(old)} صف، أجزاء فاشلة {len(failed)}")
    return new, chg, failed


def run_store(st, mode, start, log, stop, step=STEP_DAYS):
    """mode: 'full' (من start) أو 'daily' (آخر فترة أمان). التحديث اليومي لو مفيش بيانات قبل كده بيتحوّل لكامل."""
    today = dt.date.today()
    state = load_state()
    s = state.get(st.name, {})
    res = {}
    for kind, code_fn, lb, upd_fn in (("finance", None, FIN_LOOKBACK, update_finance), ("orders", None, ORD_LOOKBACK, update_orders)):
        last_ok = s.get(kind + "_done")
        st0 = start - dt.timedelta(days=ORD_BACKFILL) if kind == "orders" else start   # الأوردرات بنرجع فيها قبل بداية الدفتر: رسوم أوردر آخر الشهر اللي فات بتوصل في الدفتر الشهر ده
        if mode == "full" or not last_ok or s.get(kind + "_from") != start.isoformat():   # تاريخ البداية اتغيّر → سحب كامل من التاريخ الجديد مرة واحدة
            a = st0
        else:
            a = max(st0, dt.date.fromisoformat(last_ok) - dt.timedelta(days=lb))
        if kind == "orders" and not s.get("orders_backfill_done"):   # مرة واحدة: اسحب الأوردرات الأقدم حتى لو التحديث اليومي كان خلّص قبل كده
            a = min(a, st0)
        pend = s.get(kind + "_pending") or []
        for p in pend:  # أجزاء فشلت قبل كده
            try:
                a = min(a, dt.date.fromisoformat(p))
            except Exception:
                pass
        log(f"[{st.name}] ▶️ {kind}: من {a} إلى {today}")
        n, c, failed = upd_fn(st, a, today, log, stop, step)
        s[kind + "_pending"] = [x[0].isoformat() for x in failed]
        if not failed and not stop.is_set():
            s[kind + "_done"] = today.isoformat()
            s[kind + "_from"] = start.isoformat()
            if kind == "orders":
                s["orders_backfill_done"] = True
        s[kind + "_at"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
        res[kind] = (n, c, len(failed))
        state[st.name] = s
        save_state(state)
    return res


# ═══════════════════════ Google Sheet (اختياري) ═══════════════════════
# ═══════════════ PostgreSQL (Neon) — نفس اختيار تاب Noon API (ملف noon_use_pg.txt): sheet / pg / both لكل البوتات ═══════════════
import json
PG_MODE = "sheet"
PG_LABELS = {"sheet": "📄 Google Sheet فقط", "pg": "🐘 PostgreSQL فقط", "both": "📄+🐘 الاتنين مع بعض"}
PG_URL_FILE = os.path.join(APP_DIR, "pg_url.txt")
PG_FLAG_FILE = os.path.join(APP_DIR, "noon_use_pg.txt")
PG_SCHEMA = "sheet_copy"


def pg_flag_load():
    try:
        v = open(PG_FLAG_FILE, encoding="utf-8").read().strip()
        return {"1": "pg"}.get(v, v) if {"1": "pg"}.get(v, v) in PG_LABELS else "sheet"
    except Exception:
        return "sheet"


def pg_url():
    import re as _re
    u = os.environ.get("PG_URL", "").strip()
    if not u:
        try:
            u = open(PG_URL_FILE, encoding="utf-8").read().strip()
        except Exception:
            u = ""
    u = _re.sub(r"[&?]channel_binding=[^&]*", "", u.strip().strip('"'))
    if u and "?" not in u and "&" in u:
        u = u.replace("&", "?", 1)
    return u


def pg_conn():
    import psycopg2, re as _re
    u = pg_url()
    if not u:
        raise RuntimeError("مفيش رابط Neon — اضغط زر 🔑 رابط Neon")
    m = _re.search(r"postgres(?:ql)?://([^:@/\s]+):([^@\s]+)@([^:/?&\s'\"]+)(?::(\d+))?/([A-Za-z0-9_\-]+)", u)
    if not m:
        raise RuntimeError("رابط Neon مش مفهوم — لازم يبدأ بـ postgresql://user:password@host/dbname")
    return psycopg2.connect(user=m.group(1), password=m.group(2), host=m.group(3), port=int(m.group(4) or 5432),
                            dbname=_re.sub(r"(sslmode|channel_binding)$", "", m.group(5)) or "neondb", sslmode="require", connect_timeout=20)


def pg_side(log, fn):
    """بينفّذ كتابة PostgreSQL حسب الوضع المختار في تاب Noon API. بيرجع True لو الشيت يتخطّى (وضع PostgreSQL فقط)."""
    global PG_MODE
    PG_MODE = pg_flag_load()
    if PG_MODE == "sheet":
        return False
    try:
        fn()
    except Exception as e:
        if PG_MODE == "pg":
            raise
        log(f"⚠️ PostgreSQL فشل: {str(e)[:150]} — كمّلت على الشيت")
    return PG_MODE == "pg"


def _pg_snake(name, used):
    import re as _re
    s = _re.sub(r"[^0-9a-zA-Z]+", "_", str(name)).strip("_").lower()
    if not s:
        s = f"c_{len(used) + 1}"
    if s[0].isdigit():
        s = "c_" + s
    s = s[:60]
    base, i = s, 2
    while s in used:
        s, i = f"{base}_{i}", i + 1
    used.add(s)
    return s


def _pg_s(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def pg_write_tab(tab, headers, rows, log, keys=None, orphan_set=None, scope=None, since=None):
    """بيكتب التاب في PostgreSQL (schema sheet_copy — نفس شكل PG_Copy_All عشان التطبيق يقراه).
    keys → بناء وتحديث بالمفتاح (القديم بيفضل). من غيره → استبدال كامل. كله في transaction واحدة."""
    import io as _io, csv as _csv
    headers = [str(h).strip() for h in headers]
    used = set()
    cols = [_pg_snake(h, used) for h in headers]
    table = _pg_snake(tab, set())
    c = pg_conn()
    try:
        cur = c.cursor()
        cur.execute(f"CREATE SCHEMA IF NOT EXISTS {PG_SCHEMA}")
        cur.execute(f"CREATE TABLE IF NOT EXISTS {PG_SCHEMA}._meta (table_name TEXT PRIMARY KEY, tab_title TEXT, headers JSONB, col_names JSONB, copied_at TIMESTAMPTZ DEFAULT now())")
        cur.execute(f"SELECT headers, col_names FROM {PG_SCHEMA}._meta WHERE table_name=%s", (table,))
        m = cur.fetchone()
        old = []                                   # الصفوف القديمة كـ dict بالهيدر
        if m:
            ohdr, ocols = m
            cur.execute("SELECT " + ", ".join(f'"{x}"' for x in ocols) + f' FROM {PG_SCHEMA}."{table}" ORDER BY _row')
            old = [dict(zip(ohdr, ["" if v is None else v for v in r])) for r in cur.fetchall()]
        if not rows and m:
            log(f"⚠️ {tab}: مفيش بيانات — القاعدة ما اتلمستش")
            return
        new = [{h: _pg_s(v) for h, v in zip(headers, r)} for r in rows]
        if keys and old:
            kf = lambda d: "|".join(str(d.get(k, "")).strip() for k in keys)
            idx = {}
            for i, d in enumerate(old):
                k = kf(d)
                if k.strip("|") and k not in idx:
                    idx[k] = i
            final = [dict(d) for d in old]
            adds = ups = same = orph = 0
            seen = set()
            for d in new:
                k = kf(d)
                if k in idx:
                    seen.add(k)
                    cur_d = final[idx[k]]
                    if any(str(cur_d.get(h, "")).strip() != d[h].strip() for h in headers):
                        cur_d.update(d)
                        ups += 1
                    else:
                        same += 1
                else:
                    final.append(d)
                    adds += 1
            if orphan_set and scope:
                sc_col, sc_vals = scope
                for k, i in idx.items():
                    d = final[i]
                    if k in seen or str(d.get(sc_col, "")).strip() not in sc_vals:
                        continue
                    if since and str(d.get(since[0], "")).strip()[:10] < since[1]:
                        continue
                    if any(str(d.get(cn, "")).strip().lower() != str(x).strip().lower() for cn, x in orphan_set.items()):
                        d.update(orphan_set)
                        orph += 1
            body = final
            msg = f"{adds} جديد · {ups} اتحدّث" + (f" · {orph} اتعلّموا" if orph else "") + f" · {same} زي ما هو"
        else:
            body = new
            msg = f"اتبدّل بالكامل"
        cur.execute(f'DROP TABLE IF EXISTS {PG_SCHEMA}."{table}"')
        cur.execute(f'CREATE TABLE {PG_SCHEMA}."{table}" (' + ", ".join(f'"{x}" TEXT' for x in cols) + ", _row BIGSERIAL)")
        cur.execute(f"INSERT INTO {PG_SCHEMA}._meta (table_name, tab_title, headers, col_names, copied_at) VALUES (%s,%s,%s,%s,now()) "
                    f"ON CONFLICT (table_name) DO UPDATE SET tab_title=EXCLUDED.tab_title, headers=EXCLUDED.headers, col_names=EXCLUDED.col_names, copied_at=now()",
                    (table, tab, json.dumps(headers, ensure_ascii=False), json.dumps(cols)))
        if body:
            buf = _io.StringIO()
            w = _csv.writer(buf)
            for d in body:
                w.writerow([d.get(h, "") for h in headers])
            buf.seek(0)
            cur.copy_expert(f'COPY {PG_SCHEMA}."{table}" (' + ", ".join(f'"{x}"' for x in cols) + ") FROM STDIN WITH CSV", buf)
        c.commit()
        log(f"🐘 {tab}: {msg} — الإجمالي في PostgreSQL {len(body)}")
    except Exception:
        c.rollback()
        raise
    finally:
        c.close()


STORAGE_TAB = "Noon_Storage_API"
STORAGE_REPORT = "fbn_finance_monthlystoragechargereport"
STORAGE_HEADERS = ["Store", "soa_reference_nr", "description_en", "id_partner", "service_month", "pbarcode", "sku", "year_month",
                   "warehouse_code", "country_code", "average_stock_on_hand", "average_pending_qty", "average_net_qty", "longest_side",
                   "median_side", "shortest_side", "weight", "volumetric_weight", "item_volume", "estimated_total_item_volume",
                   "cost_per_cubic_feet", "charged_amount"]


def pg_push(tab, cols, values, key_cols, log):
    """بناء وتحديث بالمفتاح جوه PostgreSQL نفسها (من غير ما نحمّل التاب كله في الذاكرة): الصفوف اللي مفتاحها اتبعت بتتبدّل والباقي زي ما هو."""
    import io as _io, csv as _csv
    headers = [str(h).strip() for h in cols]
    table = _pg_snake(tab, set())
    c = pg_conn()
    try:
        cur = c.cursor()
        cur.execute(f"CREATE SCHEMA IF NOT EXISTS {PG_SCHEMA}")
        cur.execute(f"CREATE TABLE IF NOT EXISTS {PG_SCHEMA}._meta (table_name TEXT PRIMARY KEY, tab_title TEXT, headers JSONB, col_names JSONB, copied_at TIMESTAMPTZ DEFAULT now())")
        cur.execute(f"SELECT headers, col_names FROM {PG_SCHEMA}._meta WHERE table_name=%s", (table,))
        m = cur.fetchone()
        cur.execute("SELECT to_regclass(%s)", (f'{PG_SCHEMA}."{table}"',))
        if not m or cur.fetchone()[0] is None or not key_cols or any(k not in headers for k in key_cols):
            c.close()
            c = None
            return pg_write_tab(tab, cols, values, log)
        ohdr, ocols = [str(h) for h in m[0]], list(m[1])
        used = set(ocols)
        for h in headers:
            if h not in ohdr:
                cn = _pg_snake(h, used)
                cur.execute(f'ALTER TABLE {PG_SCHEMA}."{table}" ADD COLUMN "{cn}" TEXT')
                ohdr.append(h)
                ocols.append(cn)
        cmap = dict(zip(ohdr, ocols))
        ncols = [cmap[h] for h in headers]
        cur.execute("CREATE TEMP TABLE _pgn (" + ", ".join(f'"{x}" TEXT' for x in ncols) + ", _n BIGSERIAL) ON COMMIT DROP")
        buf = _io.StringIO()
        w = _csv.writer(buf)
        for r in values:
            w.writerow([_pg_s(v) for v in r])
        buf.seek(0)
        cur.copy_expert("COPY _pgn (" + ", ".join(f'"{x}"' for x in ncols) + ") FROM STDIN WITH CSV", buf)
        kc = [cmap[k] for k in key_cols]
        ke = lambda a: "(" + " || '|' || ".join(f"COALESCE(btrim({a}.\"{k}\"), '')" for k in kc) + ")"
        cur.execute(f'DELETE FROM {PG_SCHEMA}."{table}" m USING _pgn n WHERE {ke("m")} = {ke("n")} AND btrim({ke("n")}, \'|\') <> \'\'')
        old_n = cur.rowcount
        cl = ", ".join(f'"{x}"' for x in ncols)
        cur.execute(f'INSERT INTO {PG_SCHEMA}."{table}" ({cl}) SELECT {cl} FROM (SELECT *, row_number() OVER (PARTITION BY {ke("_pgn")} ORDER BY _n DESC) rn FROM _pgn) t '
                    f"WHERE rn = 1 OR btrim({ke('t')}, '|') = ''")
        new_n = cur.rowcount
        cur.execute(f"UPDATE {PG_SCHEMA}._meta SET headers=%s, col_names=%s, copied_at=now() WHERE table_name=%s",
                    (json.dumps(ohdr, ensure_ascii=False), json.dumps(ocols), table))
        c.commit()
        cur.execute(f'SELECT count(*) FROM {PG_SCHEMA}."{table}"')
        total = cur.fetchone()[0]
        log(f"🐘 {tab}: اتكتب {new_n:,} صف ({old_n:,} منهم كانوا موجودين واتبدّلوا) — الإجمالي في PostgreSQL {total:,}")
    finally:
        if c is not None:
            c.close()


# ═══════════════════════ كتابة الشيت: بناء وتحديث (Upsert) — عمر ما بيمسح الأول ═══════════════════════
def _su_norm(v):
    t = str(v if v is not None else "").strip()
    if t == "":
        return ""
    if t.lower() in ("true", "false"):                  # TRUE (الشيت) = True (بايثون)
        return t.lower()
    try:
        return repr(round(float(t.replace(",", "")), 6))
    except ValueError:
        return t


def _su_retry(fn, log, tag, tries=6):
    for a in range(tries):
        try:
            return fn()
        except Exception as e:
            if a == tries - 1:
                raise
            log(f"⚠️ {tag}: محاولة {a + 1} فشلت: {str(e)[:120]}")
            _m = str(e)
            time.sleep(65 if ("429" in _m or "Quota" in _m or "quota" in _m) else 3 * (a + 1))


def _su_fill(ws, cols, values, log, tag, vio="RAW", chunk=2000):
    """يكتب العنوان ثم الصفوف على دفعات — أي دفعة تفشل بعد 5 محاولات بترمي خطأ (مفيش صمت)."""
    _su_retry(lambda: ws.resize(rows=max(len(values) + 10, 100), cols=max(len(cols), 6)), log, tag)
    _su_retry(lambda: ws.update(range_name="A1", values=[list(cols)], value_input_option=vio), log, tag)
    for i in range(0, len(values), chunk):
        part = values[i:i + chunk]
        _su_retry(lambda p=part: ws.append_rows(p, value_input_option=vio), log, tag)
        log(f"⬆️ {tag}: {min(i + chunk, len(values))}/{len(values)}")


def _su_swap(ss, ws, tab, cols, values, log, vio):
    """يبني التاب كامل في تاب مؤقت، ولما يكتمل بس بيبدّل القديم — القديم عمره ما بيتمسح قبل ما الجديد يتكتب."""
    import gspread
    tmp_title = (tab + "_new")[:100]
    try:
        ss.del_worksheet(ss.worksheet(tmp_title))
    except gspread.exceptions.WorksheetNotFound:
        pass
    tmp = ss.add_worksheet(title=tmp_title, rows="100", cols=str(max(len(cols), 6)))
    try:
        _su_fill(tmp, cols, values, log, tab, vio)
    except Exception as e:
        try:
            ss.del_worksheet(tmp)
        except Exception:
            pass
        if "limit of" in str(e) and "cells" in str(e):
            raise RuntimeError(f"ملف Google Sheet وصل حد الـ 20 مليون خلية — محتاج {len(values) * len(cols):,} خلية لتاب {tab}. امسح تابات/صفوف قديمة وجرّب تاني") from None
        raise
    ss.del_worksheet(ws)
    tmp.update_title(tab)


def sheet_replace_safe(ss, tab, cols, values, log, vio="RAW"):
    """للتابات اللي بتتبني كلها كل مرة (مخزون/كتالوج/كاش): بناء تاب جديد كامل ثم تبديل — مفيش مسح قبل الكتابة."""
    import gspread
    cols = list(cols)
    if not values:
        log(f"⚠️ {tab}: مفيش بيانات — الشيت ما اتلمسش")
        return
    if pg_side(log, lambda: pg_write_tab(tab, cols, values, log)):
        return
    try:
        ws = ss.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab, rows="100", cols=str(max(len(cols), 6)))
        _su_fill(ws, cols, values, log, tab, vio)
        log(f"🆕 {tab}: اتبنى من الصفر — {len(values)} صف")
        return
    _su_swap(ss, ws, tab, cols, values, log, vio)
    log(f"🔁 {tab}: اتبنى كامل واتبدّل — {len(values)} صف")


def sheet_upsert(ss, tab, cols, values, key_cols, log, orphan_set=None, scope=None, since=None, vio="RAW", cmp_cols=None):
    """يضيف الصفوف الجديدة ويحدّث اللي اتغيّر (بالمفتاح key_cols) من غير ما يمسح حاجة.
    cmp_cols: الأعمدة اللي تغيّرها يعتبر الصف اتحدّث (الافتراضي كلها).
    orphan_set/scope/since: صفوف في الشيت ومش في البيانات الجديدة (في متاجر scope=(عمود, قيم) ومن تاريخ since=(عمود, تاريخ))
    بتتعلّم بالقيم دي بدل ما تتمسح."""
    import gspread
    from gspread.utils import rowcol_to_a1
    if not values:
        log(f"⚠️ {tab}: مفيش بيانات — الشيت ما اتلمسش")
        return
    if pg_side(log, lambda: (pg_write_tab(tab, cols, values, log, keys=key_cols, orphan_set=orphan_set, scope=scope, since=since)
                             if orphan_set else pg_push(tab, cols, values, key_cols, log))):
        return
    cols = list(cols)
    try:
        ws = ss.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        ws = None
    ex = _su_retry(lambda: ws.get_all_values(), log, tab) if ws is not None else []
    hdr = [str(h).strip() for h in ex[0]] if ex else []
    has_data = len(ex) > 1
    if ws is None:
        ws = ss.add_worksheet(title=tab, rows="100", cols=str(max(len(cols), 6)))
        _su_fill(ws, cols, values, log, tab, vio)
        log(f"🆕 {tab}: اتبنى من الصفر — {len(values)} صف")
        return
    if not hdr or not has_data:
        _su_fill(ws, cols, values, log, tab, vio)
        log(f"🆕 {tab}: كان فاضي — اتكتب {len(values)} صف")
        return
    if not all(k in cols and k in hdr for k in key_cols):     # أول مرة بعد إضافة عمود المفتاح: بناء كامل + تبديل
        _su_swap(ss, ws, tab, cols, values, log, vio)
        log(f"🔁 {tab}: اتبنى من جديد بعمود المفتاح — {len(values)} صف")
        return
    new_cols = [c for c in cols if c not in hdr]
    fh = hdr + new_cols
    if new_cols:
        if ws.col_count < len(fh):
            _su_retry(lambda: ws.resize(cols=len(fh)), log, tab)
        _su_retry(lambda: ws.update(range_name="A1", values=[fh], value_input_option="RAW"), log, tab)
    W = len(fh)
    pos = {c: i for i, c in enumerate(fh)}
    kix = [pos[k] for k in key_cols]
    idx = {}
    for i, row in enumerate(ex[1:], start=2):
        row = list(row) + [""] * (W - len(row))
        k = "|".join(str(row[j]).strip() for j in kix)
        if k.strip("|") and k not in idx:
            idx[k] = (i, row)
    ci = {c: j for j, c in enumerate(cols)}
    newmap = {}
    for v in values:
        newmap["|".join(str(v[ci[k]]).strip() for k in key_cols)] = v
    cc = [c for c in (cmp_cols or cols) if c in ci]
    ups, adds, same, orph, seen = [], [], 0, 0, set()
    _dc, _dx = {}, {}
    for k, v in newmap.items():
        if k in idx:
            i, row = idx[k]
            seen.add(k)
            _dif = [c for c in cc if _su_norm(row[pos[c]]) != _su_norm(v[ci[c]])]
            if _dif:
                for c in _dif:                                   # تشخيص: أنهي أعمدة بتفرّق (عشان نعرف ليه صفوف كتير بتتحدّث)
                    _dc[c] = _dc.get(c, 0) + 1
                    if c not in _dx:
                        _dx[c] = (str(row[pos[c]])[:40], str(v[ci[c]])[:40])
                tgt = list(row)
                for c in cols:
                    tgt[pos[c]] = v[ci[c]]
                ups.append((i, tgt))
            else:
                same += 1
        else:
            adds.append([v[ci[c]] if c in ci else "" for c in fh])
    if orphan_set and scope:
        sc_col, sc_vals = scope
        for k, (i, row) in idx.items():
            if k in seen or str(row[pos[sc_col]]).strip() not in sc_vals:
                continue
            if since and str(row[pos[since[0]]]).strip()[:10] < since[1]:
                continue
            if any(_su_norm(row[pos[c]]) != _su_norm(x) for c, x in orphan_set.items() if c in pos):
                tgt = list(row)
                for c, x in orphan_set.items():
                    if c in pos:
                        tgt[pos[c]] = x
                ups.append((i, tgt))
                orph += 1
    last = "".join(ch for ch in rowcol_to_a1(1, W) if not ch.isdigit())
    for s in range(0, len(ups), 500):
        data = [{"range": f"A{i}:{last}{i}", "values": [r]} for i, r in ups[s:s + 500]]
        _su_retry(lambda d=data: ws.batch_update([dict(x) for x in d], value_input_option=vio), log, tab)
    for s in range(0, len(adds), 2000):
        part = adds[s:s + 2000]
        _su_retry(lambda p=part: ws.append_rows(p, value_input_option=vio), log, tab)
    if _dc and len(ups) - orph > 200:                            # كتير اتحدّث: اطبع أكتر الأعمدة اختلافاً مع مثال (الشيت ← الجديد)
        for c, n in sorted(_dc.items(), key=lambda x: -x[1])[:5]:
            log(f"🔎 {tab}: عمود «{c}» اختلف في {n} صف — مثال: «{_dx[c][0]}» ← «{_dx[c][1]}»")
    log(f"⬆️ {tab}: {len(adds)} جديد · {len(ups) - orph} اتحدّث" + (f" · {orph} اتعلّموا" if orph else "") + f" · {same} زي ما هو — الإجمالي في الشيت {len(idx) + len(adds)}")


def push_sheet(store_names, log):
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    ms = glob.glob(os.path.join(GSHEET_CREDS_DIR, "*.json"))
    if not ms:
        raise RuntimeError(f"مفيش ملف اعتماد Google في {GSHEET_CREDS_DIR}")
    ss = gspread.authorize(ServiceAccountCredentials.from_json_keyfile_name(ms[0], GSHEET_SCOPE)).open(GSHEET_NAME)
    for tab, fname in ((TAB_FIN, "finance.csv"), (TAB_ORD, "orders.csv")):
        allrows, cols = [], ["Store"]
        for sn in store_names:
            rd, c = read_ledger(os.path.join(LEDGER_DIR, sn, fname))
            for c1 in c:
                if c1 not in cols:
                    cols.append(c1)
            for r in rd:
                r["Store"] = sn
                allrows.append(r)
        if not allrows:
            continue
        values = []
        for r in allrows:
            row = []
            for c in cols:
                v = r.get(c, "")
                if c in FIN_NUM or c == "Net Proceeds":
                    try:
                        v = float(v)
                    except Exception:
                        pass
                row.append(v)
            values.append(row)
        sheet_upsert(ss, tab, cols, values, (("Store",) + FIN_KEY) if tab == TAB_FIN else ("Store", "item_nr", "order_nr", "sku"), log)


# ═══════════════════════ الواجهة / التشغيل التلقائي ═══════════════════════
def load_store_paths():
    out = []
    for f in (STORES_FILE, os.path.join(APP_DIR, "noon_stores.json")):
        try:
            for p in json.load(open(f, encoding="utf-8")):
                if os.path.exists(p) and p not in out:
                    out.append(p)
        except Exception:
            pass
    return out


def auto_main():
    stop = threading.Event()
    log = lambda m: print(time.strftime("%H:%M:%S"), m, flush=True)
    names = []
    for p in load_store_paths():
        try:
            st = Store(p)
            names.append(st.name)
            run_store(st, "daily", DEFAULT_START, log, stop)
        except Exception as e:
            log(f"❌ {os.path.basename(p)}: {type(e).__name__}: {e}")
    if "--sheet" in sys.argv and names:
        push_sheet(names, log)


import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox


class App(tk.Frame):
    def __init__(self, root):
        super().__init__(root, bg="#d9d9d9")
        self.pack(fill="both", expand=True, padx=8, pady=8)
        self.stop = threading.Event()
        self.stores = []
        self.running = False
        self.next_run = None
        self._last_fixed = ""
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)
        ttk.Label(top, text="المتاجر (ملفات المفاتيح):").grid(row=0, column=0, sticky="nw")
        self.lb = tk.Listbox(top, height=5, width=80)
        self.lb.grid(row=0, column=1, padx=6)
        bb = ttk.Frame(top)
        bb.grid(row=0, column=2, sticky="n")
        ttk.Button(bb, text="➕ إضافة", command=self.add_key).pack(fill="x")
        ttk.Button(bb, text="➖ حذف", command=self.rm_key).pack(fill="x", pady=2)
        ttk.Label(top, text="تاريخ البداية:").grid(row=1, column=0, sticky="w", pady=6)
        self.dfrom = tk.StringVar(value=DEFAULT_START.isoformat())
        ttk.Entry(top, textvariable=self.dfrom, width=14).grid(row=1, column=1, sticky="w", padx=6)
        self.sheet = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="اكتب كمان في Google Sheet (Noon_Finance_API / Noon_Orders_API)", variable=self.sheet).grid(row=2, column=1, sticky="w")
        self.auto = tk.BooleanVar(value=False)
        self.times = tk.StringVar(value="")
        ttk.Checkbutton(top, text="⏰ تحديث يومي تلقائي في المواعيد دي (HH:MM مفصولة بفاصلة):", variable=self.auto, command=self.save_settings).grid(row=3, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Entry(top, textvariable=self.times, width=24).grid(row=3, column=2, sticky="w")
        self.next_lbl = ttk.Label(top, text="")
        self.next_lbl.grid(row=4, column=1, sticky="w")
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=10)
        self.btns = []
        for txt, mode in (("🔄 تحديث يومي (آخر فترة)", "daily"), ("📥 سحب كامل من تاريخ البداية", "full")):
            b = ttk.Button(bar, text=txt, command=lambda m=mode: self.go(m))
            b.pack(side="left", padx=4, pady=6)
            self.btns.append(b)
        ttk.Button(bar, text="⏹ إيقاف", command=self.stop.set).pack(side="left", padx=4)
        self.box = scrolledtext.ScrolledText(self, font=("Consolas", 10), height=16)
        self.box.pack(fill="both", expand=True, padx=10, pady=8)
        for p in load_store_paths():
            self.add_path(p, save=False)
        self.load_settings()
        for _v in (self.sheet, self.times, self.dfrom):
            _v.trace_add("write", lambda *a: self.save_settings())
        self.log("جاهز. جرّب متجر واحد الأول: ضيف مفتاحه واضغط «سحب كامل». بعدها اليومي بيحدّث بس.")
        self.after(5000, self.tick)

    def log(self, m):
        try:
            if LOG_HOOK:
                LOG_HOOK(str(m) + "\n")
        except Exception:
            pass
        try:
            self.after(0, lambda: (self.box.insert("end", str(m) + "\n"), self.box.see("end")))
        except Exception:
            pass

    def load_settings(self):
        try:
            s = json.load(open(SETTINGS_FILE, encoding="utf-8"))
            self.sheet.set(bool(s.get("sheet", False)))
            self.times.set(str(s.get("times", "")))
            self.auto.set(bool(s.get("auto", False)))
            if s.get("dfrom"):
                self.dfrom.set(str(s["dfrom"]))
        except Exception:
            pass

    def save_settings(self):
        try:
            json.dump({"sheet": bool(self.sheet.get()), "times": self.times.get(), "auto": bool(self.auto.get()), "dfrom": self.dfrom.get()},
                      open(SETTINGS_FILE, "w", encoding="utf-8"), ensure_ascii=False)
        except Exception:
            pass

    def fixed_times(self):
        out = []
        for t in str(self.times.get()).replace("،", ",").replace(" ", ",").split(","):
            t = t.strip()
            try:
                if t:
                    out.append(dt.datetime.strptime(t, "%H:%M").strftime("%H:%M"))
            except ValueError:
                pass
        return out

    def tick(self):
        try:
            ft = self.fixed_times()
            self.next_lbl.configure(text=("المواعيد: " + " ، ".join(ft)) if (self.auto.get() and ft) else "")
            now = time.strftime("%H:%M")
            key = time.strftime("%Y-%m-%d ") + now
            if self.auto.get() and now in ft and key != self._last_fixed and not self.running:
                self._last_fixed = key
                self.log(f"🕐 موعد محدد {now} — تحديث يومي تلقائي...")
                self.go("daily", auto=True)
        except Exception:
            pass
        self.after(20000, self.tick)

    def save(self):
        try:
            json.dump([s.path for s in self.stores], open(STORES_FILE, "w", encoding="utf-8"), ensure_ascii=False)
        except Exception:
            pass

    def add_path(self, p, save=True):
        try:
            st = Store(p)
        except Exception as e:
            self.log(f"⚠️ {os.path.basename(p)}: {e}")
            return
        if any(s.path == p for s in self.stores):
            return
        self.stores.append(st)
        self.lb.insert("end", f"{st.name}  —  {p}")
        if save:
            self.save()

    def add_key(self):
        for p in filedialog.askopenfilenames(title="ملفات مفاتيح نون", filetypes=[("JSON", "*.json"), ("All", "*.*")]):
            self.add_path(p)

    def rm_key(self):
        for i in reversed(self.lb.curselection()):
            self.lb.delete(i)
            del self.stores[i]
        self.save()

    def go(self, mode, auto=False, sheet=None):
        if self.running:
            return
        if not self.stores:
            if not auto:
                messagebox.showwarning("", "ضيف ملف مفتاح واحد على الأقل")
            return
        try:
            start = dt.date.fromisoformat(self.dfrom.get().strip())
        except Exception as e:
            if not auto:
                messagebox.showerror("خطأ", str(e))
            return
        self.stop.clear()
        self.running = True
        for b in self.btns:
            b.state(["disabled"])
        want_sheet = self.sheet.get() if sheet is None else bool(sheet)
        stores = list(self.stores)

        def work():
            try:
                for st in stores:
                    if self.stop.is_set():
                        break
                    try:
                        run_store(st, mode, start, self.log, self.stop)
                    except Exception as e:
                        self.log(f"❌ [{st.name}] {type(e).__name__}: {e}")
                if want_sheet and not self.stop.is_set():
                    self.log("⬆️ كتابة الشيت...")
                    push_sheet([s.name for s in stores], self.log)
                self.log("✅ خلص.")
            except Exception as e:
                self.log(f"❌ {type(e).__name__}: {e}")
            finally:
                self.running = False
                try:
                    self.after(0, lambda: [b.state(["!disabled"]) for b in self.btns])
                except Exception:
                    pass
        threading.Thread(target=work, daemon=True).start()

    def shutdown(self):
        self.stop.set()


def build_ui(root):
    """بيتنادى من البرنامج الموحد: بيبني الواجهة جوه التاب ويرجّع دالة إيقاف."""
    return App(root).shutdown


def main():
    root = tk.Tk()
    root.title("Noon Accounts — دفتر حسابات نون")
    root.geometry("1000x720")
    a = App(root)
    root.protocol("WM_DELETE_WINDOW", lambda: (a.shutdown(), root.destroy()))
    root.mainloop()


if __name__ == "__main__":
    if "--auto" in sys.argv:
        auto_main()
    else:
        main()
'''

_EMBEDDED['TrendyolAccounts'] = r'''# -*- coding: utf-8 -*-
"""
Trendyol_Accounts.py — دفتر حسابات تريندول (متعدد المتاجر): بيسحب كل المعاملات المالية (مبيعات/مرتجعات/خصومات/أوامر الدفع/فواتير الخصم والعمولة)
وبيحدّثها بالمفتاح id (Upsert): الجديد بيتضاف والمتغيّر بيتحدّث من غير تكرار. التحديث اليومي بيرجع بفترة أمان (35 يوم) لأن
paymentOrderId بيتحط على الصفوف القديمة بعد ما الدفعة بتتجهز.

المخرجات (جنب الأداة):
  trendyol_accounts/<المتجر>/finance.csv   — الدفتر (مفتاح: id)
  trendyol_accounts/state.json             — آخر تحديث لكل متجر
  + (اختياري) تاب في الشيت: Trendyol_Finance_API (كل المتاجر، عمود Store) — بيتبني ويتحدّث بالمفتاح (id) من الدفتر — مفيش مسح

المتاجر: نفس ملف trendyol_stores.json بتاع Trendyol API (أو ضيفهم بالزرار).
التشغيل:  py Trendyol_Accounts.py          (واجهة)
          py Trendyol_Accounts.py --auto [--sheet]   (بدون واجهة — للجدولة)
المتطلبات: py -m pip install requests gspread oauth2client
"""
import csv, json, os, sys, glob, time, threading, datetime as dt
import requests

BASE = "https://apigw.trendyol.com"
APP_DIR = os.environ.get("PLATFORMS_BASE_DIR") or os.path.dirname(os.path.abspath(__file__))
EMBEDDED = bool(os.environ.get("PLATFORMS_BASE_DIR"))
LOG_HOOK = None   # بيتحط من البرنامج الموحد → السجل المشترك
STORES_FILE = os.path.join(APP_DIR, "trendyol_stores.json")
SETTINGS_FILE = os.path.join(APP_DIR, "trendyol_accounts_settings.json")
LEDGER_DIR = os.path.join(APP_DIR, "trendyol_accounts")
STATE_FILE = os.path.join(LEDGER_DIR, "state.json")
GSHEET_CREDS_DIR = r"\\WIN-181U7G4HFM5\Share For Markting\NOONIMG\Copy"
GSHEET_NAME = "Complaints"
GSHEET_SCOPE = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
TAB_FIN = "Trendyol_Finance_API"
DEFAULT_START = dt.date(2026, 9, 1)
FIN_LOOKBACK = 35        # يوم — paymentOrderId بيتحط على صفوف قديمة بعد تجهيز الدفعة
WIN_DAYS = 14            # تريندول بتسمح بـ 15 يوم كحد أقصى للطلب
TS_OFFSET_H = 3          # توقيت السعودية
# ترتيب مهم: CommissionInvoice قبل DeductionInvoices (نفس الفاتورة بتظهر في الاتنين، بنحتفظ بنوعها الأصلي)
FIN_TYPES = [("settlements", t) for t in ("Sale", "Return", "Discount", "DiscountCancel", "Coupon", "CouponCancel",
                                           "ProvisionPositive", "ProvisionNegative", "ManualRefund", "ManualRefundCancel",
                                           "TyDiscount", "TyDiscountCancel", "TyCoupon", "TyCouponCancel")] + \
            [("otherfinancials", t) for t in ("PaymentOrder", "CommissionInvoice", "DeductionInvoices", "CreditNote")]
PERSONAL = ("address", "phone", "mobile", "email", "customer", "recipient", "firstname", "lastname", "fullname", "identity")
DROP_COLS = ("receiptId", "storeName", "shipmentPackageId", "orderLineItemId")
NUM_COLS = ("debt", "credit", "commissionAmount", "sellerRevenue", "commissionRate", "paymentPeriod")
_slock = threading.Lock()


class TStore:
    def __init__(self, d):
        self.name = (d.get("name") or d["seller_id"]).strip()
        self.sid = str(d["seller_id"]).strip()
        self.key = d["api_key"].strip()
        self.secret = d["api_secret"].strip()
        self.sf = (d.get("store_front") or "SA").strip().upper()
        self.sess = requests.Session()
        self.sess.auth = (self.key, self.secret)
        self.sess.headers.update({"User-Agent": f"{self.sid} - SelfIntegration", "storeFrontCode": self.sf, "Accept": "application/json"})

    def as_dict(self):
        return {"name": self.name, "seller_id": self.sid, "api_key": self.key, "api_secret": self.secret, "store_front": self.sf}


class ApiError(Exception):
    def __init__(self, code, msg):
        super().__init__(msg)
        self.code = code


def get_json(st, path, params, tries=5):
    url = BASE + path.format(sid=st.sid)
    last = ""
    for a in range(tries):
        try:
            r = st.sess.get(url, params=params, timeout=60)
        except Exception as e:
            last = f"{type(e).__name__}: {str(e)[:120]}"
            time.sleep(2 * (a + 1))
            continue
        if r.status_code == 200:
            try:
                return r.json()
            except Exception:
                raise ApiError(200, f"رد مش JSON: {(r.text or '')[:150]}")
        last = f"HTTP {r.status_code}: {(r.text or '')[:200].replace(chr(10), ' ')}"
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(3 * (a + 1))
            continue
        raise ApiError(r.status_code, last)
    raise ApiError(0, f"فشل بعد {tries} محاولات — {last}")


def ms_to_local(ms):
    try:
        return (dt.datetime.utcfromtimestamp(int(ms) / 1000) + dt.timedelta(hours=TS_OFFSET_H)).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return ""


def day_ms(d, end=False):
    t = dt.datetime.combine(d, dt.time.max if end else dt.time.min) - dt.timedelta(hours=TS_OFFSET_H)
    return int(t.replace(tzinfo=dt.timezone.utc).timestamp() * 1000)


def date_chunks(d_from, d_to, step=WIN_DAYS):
    a = d_from
    while a <= d_to:
        b = min(a + dt.timedelta(days=step - 1), d_to)
        yield a, b
        a = b + dt.timedelta(days=1)


def is_personal(k):
    k = str(k).lower()
    return any(w in k for w in PERSONAL)


def flat(item):
    row = {}
    for k, v in item.items():
        if is_personal(k) or k in DROP_COLS:
            continue
        if isinstance(v, (dict, list)):
            row[k] = json.dumps(v, ensure_ascii=False)[:300]
        elif ("date" in k.lower()) and isinstance(v, (int, float)) and v > 10 ** 11:
            row[k] = ms_to_local(v)
        elif isinstance(v, str):
            row[k] = v.replace("\n", " ").strip()
        else:
            row[k] = "" if v is None else v
    return row


# ═══════════════════════ الدفتر ═══════════════════════
def read_ledger(path):
    if not os.path.exists(path):
        return [], []
    with open(path, encoding="utf-8-sig", newline="") as fh:
        rd = csv.DictReader(fh)
        cols = list(rd.fieldnames or [])
        return list(rd), cols


def write_ledger(path, rows, cols):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    os.replace(tmp, path)


def load_state():
    try:
        return json.load(open(STATE_FILE, encoding="utf-8"))
    except Exception:
        return {}


def save_state(s):
    os.makedirs(LEDGER_DIR, exist_ok=True)
    with _slock:
        json.dump(s, open(STATE_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def pull_type(st, kind, ttype, d_from, d_to, log, stop):
    """يرجع (صفوف, تواريخ بداية الأجزاء الفاشلة). نوع بيرفضه تريندول (400) بيتتخطى من غير ما يعتبر فشل."""
    path = "/integration/finance/che/sellers/{sid}/" + ("settlements" if kind == "settlements" else "otherfinancials")
    out, failed = [], []
    for a, b in date_chunks(d_from, d_to):
        if stop.is_set():
            failed.append(a)
            continue
        page = 0
        try:
            while True:
                j = get_json(st, path, {"startDate": day_ms(a), "endDate": day_ms(b, True), "transactionType": ttype,
                                        "page": page, "size": 1000, "supplierId": st.sid})
                out.extend(j.get("content") or [])
                page += 1
                if page >= int(j.get("totalPages") or 0):
                    break
        except ApiError as e:
            if e.code == 400:
                log(f"[{st.name}] ⚠️ {ttype}: تريندول رفض النوع ده ({str(e)[:90]}) — اتخطّيته")
                return [], []
            log(f"[{st.name}] ⚠️ {ttype} {a}→{b}: {str(e)[:120]}")
            failed.append(a)
    return out, failed


def update_finance(st, d_from, d_to, log, stop):
    """Upsert بالـ id. يرجع (جديد, اتحدّث, تواريخ فاشلة)."""
    path = os.path.join(LEDGER_DIR, st.name, "finance.csv")
    rows_old, cols = read_ledger(path)
    old = {str(r.get("id", "")): r for r in rows_old}
    seen_run, new, upd, failed_all = set(), 0, 0, []
    for kind, tt in FIN_TYPES:
        if stop.is_set():
            failed_all.append(d_from)
            break
        items, failed = pull_type(st, kind, tt, d_from, d_to, log, stop)
        failed_all += failed
        if items:
            log(f"[{st.name}] ✅ {tt}: {len(items)} صف")
        for it in items:
            r = flat(it)
            k = str(r.get("id", "")).strip()
            if not k or k in seen_run:
                continue
            seen_run.add(k)
            r["_ttype"] = tt
            if k in old:
                o = old[k]
                if o.get("_ttype") == "CommissionInvoice":
                    r["_ttype"] = "CommissionInvoice"
                if any(str(o.get(c, "")) != str(r.get(c, "")) for c in r.keys()):
                    upd += 1
                o.update(r)
            else:
                r["_first_seen"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
                old[k] = r
                new += 1
    for r in old.values():
        for c in r.keys():
            if c not in cols:
                cols.append(c)
    for c in ("_ttype", "_first_seen"):
        if c not in cols:
            cols.append(c)
    write_ledger(path, list(old.values()), cols)
    log(f"[{st.name}] 💰 المالية: {new} جديد، {upd} اتحدّث، الإجمالي {len(old)} صف، أجزاء فاشلة {len(failed_all)}")
    return new, upd, failed_all


def run_store(st, mode, start, log, stop):
    """mode: 'full' (من start) أو 'daily' (آخر فترة أمان). لو مفيش تحديث ناجح قبل كده بيتحوّل لكامل."""
    today = dt.date.today()
    state = load_state()
    s = state.get(st.name, {})
    last_ok = s.get("finance_done")
    a = start if (mode == "full" or not last_ok or s.get("finance_from") != start.isoformat()) else max(start, dt.date.fromisoformat(last_ok) - dt.timedelta(days=FIN_LOOKBACK))
    for p in (s.get("finance_pending") or []):
        try:
            a = min(a, dt.date.fromisoformat(p))
        except Exception:
            pass
    log(f"[{st.name}] ▶️ المالية: من {a} إلى {today}")
    n, c, failed = update_finance(st, a, today, log, stop)
    s["finance_pending"] = sorted({x.isoformat() for x in failed})
    if not failed and not stop.is_set():
        s["finance_done"] = today.isoformat()
        s["finance_from"] = start.isoformat()
    s["finance_at"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    state[st.name] = s
    save_state(state)
    return n, c, len(failed)


# ═══════════════════════ Google Sheet (اختياري) ═══════════════════════
# ═══════════════ PostgreSQL (Neon) — نفس اختيار تاب Noon API (ملف noon_use_pg.txt): sheet / pg / both لكل البوتات ═══════════════
import json
PG_MODE = "sheet"
PG_LABELS = {"sheet": "📄 Google Sheet فقط", "pg": "🐘 PostgreSQL فقط", "both": "📄+🐘 الاتنين مع بعض"}
PG_URL_FILE = os.path.join(APP_DIR, "pg_url.txt")
PG_FLAG_FILE = os.path.join(APP_DIR, "noon_use_pg.txt")
PG_SCHEMA = "sheet_copy"


def pg_flag_load():
    try:
        v = open(PG_FLAG_FILE, encoding="utf-8").read().strip()
        return {"1": "pg"}.get(v, v) if {"1": "pg"}.get(v, v) in PG_LABELS else "sheet"
    except Exception:
        return "sheet"


def pg_url():
    import re as _re
    u = os.environ.get("PG_URL", "").strip()
    if not u:
        try:
            u = open(PG_URL_FILE, encoding="utf-8").read().strip()
        except Exception:
            u = ""
    u = _re.sub(r"[&?]channel_binding=[^&]*", "", u.strip().strip('"'))
    if u and "?" not in u and "&" in u:
        u = u.replace("&", "?", 1)
    return u


def pg_conn():
    import psycopg2, re as _re
    u = pg_url()
    if not u:
        raise RuntimeError("مفيش رابط Neon — اضغط زر 🔑 رابط Neon")
    m = _re.search(r"postgres(?:ql)?://([^:@/\s]+):([^@\s]+)@([^:/?&\s'\"]+)(?::(\d+))?/([A-Za-z0-9_\-]+)", u)
    if not m:
        raise RuntimeError("رابط Neon مش مفهوم — لازم يبدأ بـ postgresql://user:password@host/dbname")
    return psycopg2.connect(user=m.group(1), password=m.group(2), host=m.group(3), port=int(m.group(4) or 5432),
                            dbname=_re.sub(r"(sslmode|channel_binding)$", "", m.group(5)) or "neondb", sslmode="require", connect_timeout=20)


def pg_side(log, fn):
    """بينفّذ كتابة PostgreSQL حسب الوضع المختار في تاب Noon API. بيرجع True لو الشيت يتخطّى (وضع PostgreSQL فقط)."""
    global PG_MODE
    PG_MODE = pg_flag_load()
    if PG_MODE == "sheet":
        return False
    try:
        fn()
    except Exception as e:
        if PG_MODE == "pg":
            raise
        log(f"⚠️ PostgreSQL فشل: {str(e)[:150]} — كمّلت على الشيت")
    return PG_MODE == "pg"


def _pg_snake(name, used):
    import re as _re
    s = _re.sub(r"[^0-9a-zA-Z]+", "_", str(name)).strip("_").lower()
    if not s:
        s = f"c_{len(used) + 1}"
    if s[0].isdigit():
        s = "c_" + s
    s = s[:60]
    base, i = s, 2
    while s in used:
        s, i = f"{base}_{i}", i + 1
    used.add(s)
    return s


def _pg_s(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def pg_write_tab(tab, headers, rows, log, keys=None, orphan_set=None, scope=None, since=None):
    """بيكتب التاب في PostgreSQL (schema sheet_copy — نفس شكل PG_Copy_All عشان التطبيق يقراه).
    keys → بناء وتحديث بالمفتاح (القديم بيفضل). من غيره → استبدال كامل. كله في transaction واحدة."""
    import io as _io, csv as _csv
    headers = [str(h).strip() for h in headers]
    used = set()
    cols = [_pg_snake(h, used) for h in headers]
    table = _pg_snake(tab, set())
    c = pg_conn()
    try:
        cur = c.cursor()
        cur.execute(f"CREATE SCHEMA IF NOT EXISTS {PG_SCHEMA}")
        cur.execute(f"CREATE TABLE IF NOT EXISTS {PG_SCHEMA}._meta (table_name TEXT PRIMARY KEY, tab_title TEXT, headers JSONB, col_names JSONB, copied_at TIMESTAMPTZ DEFAULT now())")
        cur.execute(f"SELECT headers, col_names FROM {PG_SCHEMA}._meta WHERE table_name=%s", (table,))
        m = cur.fetchone()
        old = []                                   # الصفوف القديمة كـ dict بالهيدر
        if m:
            ohdr, ocols = m
            cur.execute("SELECT " + ", ".join(f'"{x}"' for x in ocols) + f' FROM {PG_SCHEMA}."{table}" ORDER BY _row')
            old = [dict(zip(ohdr, ["" if v is None else v for v in r])) for r in cur.fetchall()]
        if not rows and m:
            log(f"⚠️ {tab}: مفيش بيانات — القاعدة ما اتلمستش")
            return
        new = [{h: _pg_s(v) for h, v in zip(headers, r)} for r in rows]
        if keys and old:
            kf = lambda d: "|".join(str(d.get(k, "")).strip() for k in keys)
            idx = {}
            for i, d in enumerate(old):
                k = kf(d)
                if k.strip("|") and k not in idx:
                    idx[k] = i
            final = [dict(d) for d in old]
            adds = ups = same = orph = 0
            seen = set()
            for d in new:
                k = kf(d)
                if k in idx:
                    seen.add(k)
                    cur_d = final[idx[k]]
                    if any(str(cur_d.get(h, "")).strip() != d[h].strip() for h in headers):
                        cur_d.update(d)
                        ups += 1
                    else:
                        same += 1
                else:
                    final.append(d)
                    adds += 1
            if orphan_set and scope:
                sc_col, sc_vals = scope
                for k, i in idx.items():
                    d = final[i]
                    if k in seen or str(d.get(sc_col, "")).strip() not in sc_vals:
                        continue
                    if since and str(d.get(since[0], "")).strip()[:10] < since[1]:
                        continue
                    if any(str(d.get(cn, "")).strip().lower() != str(x).strip().lower() for cn, x in orphan_set.items()):
                        d.update(orphan_set)
                        orph += 1
            body = final
            msg = f"{adds} جديد · {ups} اتحدّث" + (f" · {orph} اتعلّموا" if orph else "") + f" · {same} زي ما هو"
        else:
            body = new
            msg = f"اتبدّل بالكامل"
        cur.execute(f'DROP TABLE IF EXISTS {PG_SCHEMA}."{table}"')
        cur.execute(f'CREATE TABLE {PG_SCHEMA}."{table}" (' + ", ".join(f'"{x}" TEXT' for x in cols) + ", _row BIGSERIAL)")
        cur.execute(f"INSERT INTO {PG_SCHEMA}._meta (table_name, tab_title, headers, col_names, copied_at) VALUES (%s,%s,%s,%s,now()) "
                    f"ON CONFLICT (table_name) DO UPDATE SET tab_title=EXCLUDED.tab_title, headers=EXCLUDED.headers, col_names=EXCLUDED.col_names, copied_at=now()",
                    (table, tab, json.dumps(headers, ensure_ascii=False), json.dumps(cols)))
        if body:
            buf = _io.StringIO()
            w = _csv.writer(buf)
            for d in body:
                w.writerow([d.get(h, "") for h in headers])
            buf.seek(0)
            cur.copy_expert(f'COPY {PG_SCHEMA}."{table}" (' + ", ".join(f'"{x}"' for x in cols) + ") FROM STDIN WITH CSV", buf)
        c.commit()
        log(f"🐘 {tab}: {msg} — الإجمالي في PostgreSQL {len(body)}")
    except Exception:
        c.rollback()
        raise
    finally:
        c.close()


STORAGE_TAB = "Noon_Storage_API"
STORAGE_REPORT = "fbn_finance_monthlystoragechargereport"
STORAGE_HEADERS = ["Store", "soa_reference_nr", "description_en", "id_partner", "service_month", "pbarcode", "sku", "year_month",
                   "warehouse_code", "country_code", "average_stock_on_hand", "average_pending_qty", "average_net_qty", "longest_side",
                   "median_side", "shortest_side", "weight", "volumetric_weight", "item_volume", "estimated_total_item_volume",
                   "cost_per_cubic_feet", "charged_amount"]


def pg_push(tab, cols, values, key_cols, log):
    """بناء وتحديث بالمفتاح جوه PostgreSQL نفسها (من غير ما نحمّل التاب كله في الذاكرة): الصفوف اللي مفتاحها اتبعت بتتبدّل والباقي زي ما هو."""
    import io as _io, csv as _csv
    headers = [str(h).strip() for h in cols]
    table = _pg_snake(tab, set())
    c = pg_conn()
    try:
        cur = c.cursor()
        cur.execute(f"CREATE SCHEMA IF NOT EXISTS {PG_SCHEMA}")
        cur.execute(f"CREATE TABLE IF NOT EXISTS {PG_SCHEMA}._meta (table_name TEXT PRIMARY KEY, tab_title TEXT, headers JSONB, col_names JSONB, copied_at TIMESTAMPTZ DEFAULT now())")
        cur.execute(f"SELECT headers, col_names FROM {PG_SCHEMA}._meta WHERE table_name=%s", (table,))
        m = cur.fetchone()
        cur.execute("SELECT to_regclass(%s)", (f'{PG_SCHEMA}."{table}"',))
        if not m or cur.fetchone()[0] is None or not key_cols or any(k not in headers for k in key_cols):
            c.close()
            c = None
            return pg_write_tab(tab, cols, values, log)
        ohdr, ocols = [str(h) for h in m[0]], list(m[1])
        used = set(ocols)
        for h in headers:
            if h not in ohdr:
                cn = _pg_snake(h, used)
                cur.execute(f'ALTER TABLE {PG_SCHEMA}."{table}" ADD COLUMN "{cn}" TEXT')
                ohdr.append(h)
                ocols.append(cn)
        cmap = dict(zip(ohdr, ocols))
        ncols = [cmap[h] for h in headers]
        cur.execute("CREATE TEMP TABLE _pgn (" + ", ".join(f'"{x}" TEXT' for x in ncols) + ", _n BIGSERIAL) ON COMMIT DROP")
        buf = _io.StringIO()
        w = _csv.writer(buf)
        for r in values:
            w.writerow([_pg_s(v) for v in r])
        buf.seek(0)
        cur.copy_expert("COPY _pgn (" + ", ".join(f'"{x}"' for x in ncols) + ") FROM STDIN WITH CSV", buf)
        kc = [cmap[k] for k in key_cols]
        ke = lambda a: "(" + " || '|' || ".join(f"COALESCE(btrim({a}.\"{k}\"), '')" for k in kc) + ")"
        cur.execute(f'DELETE FROM {PG_SCHEMA}."{table}" m USING _pgn n WHERE {ke("m")} = {ke("n")} AND btrim({ke("n")}, \'|\') <> \'\'')
        old_n = cur.rowcount
        cl = ", ".join(f'"{x}"' for x in ncols)
        cur.execute(f'INSERT INTO {PG_SCHEMA}."{table}" ({cl}) SELECT {cl} FROM (SELECT *, row_number() OVER (PARTITION BY {ke("_pgn")} ORDER BY _n DESC) rn FROM _pgn) t '
                    f"WHERE rn = 1 OR btrim({ke('t')}, '|') = ''")
        new_n = cur.rowcount
        cur.execute(f"UPDATE {PG_SCHEMA}._meta SET headers=%s, col_names=%s, copied_at=now() WHERE table_name=%s",
                    (json.dumps(ohdr, ensure_ascii=False), json.dumps(ocols), table))
        c.commit()
        cur.execute(f'SELECT count(*) FROM {PG_SCHEMA}."{table}"')
        total = cur.fetchone()[0]
        log(f"🐘 {tab}: اتكتب {new_n:,} صف ({old_n:,} منهم كانوا موجودين واتبدّلوا) — الإجمالي في PostgreSQL {total:,}")
    finally:
        if c is not None:
            c.close()


# ═══════════════════════ كتابة الشيت: بناء وتحديث (Upsert) — عمر ما بيمسح الأول ═══════════════════════
def _su_norm(v):
    t = str(v if v is not None else "").strip()
    if t == "":
        return ""
    if t.lower() in ("true", "false"):                  # TRUE (الشيت) = True (بايثون)
        return t.lower()
    try:
        return repr(round(float(t.replace(",", "")), 6))
    except ValueError:
        return t


def _su_retry(fn, log, tag, tries=6):
    for a in range(tries):
        try:
            return fn()
        except Exception as e:
            if a == tries - 1:
                raise
            log(f"⚠️ {tag}: محاولة {a + 1} فشلت: {str(e)[:120]}")
            _m = str(e)
            time.sleep(65 if ("429" in _m or "Quota" in _m or "quota" in _m) else 3 * (a + 1))


def _su_fill(ws, cols, values, log, tag, vio="RAW", chunk=2000):
    """يكتب العنوان ثم الصفوف على دفعات — أي دفعة تفشل بعد 5 محاولات بترمي خطأ (مفيش صمت)."""
    _su_retry(lambda: ws.resize(rows=max(len(values) + 10, 100), cols=max(len(cols), 6)), log, tag)
    _su_retry(lambda: ws.update(range_name="A1", values=[list(cols)], value_input_option=vio), log, tag)
    for i in range(0, len(values), chunk):
        part = values[i:i + chunk]
        _su_retry(lambda p=part: ws.append_rows(p, value_input_option=vio), log, tag)
        log(f"⬆️ {tag}: {min(i + chunk, len(values))}/{len(values)}")


def _su_swap(ss, ws, tab, cols, values, log, vio):
    """يبني التاب كامل في تاب مؤقت، ولما يكتمل بس بيبدّل القديم — القديم عمره ما بيتمسح قبل ما الجديد يتكتب."""
    import gspread
    tmp_title = (tab + "_new")[:100]
    try:
        ss.del_worksheet(ss.worksheet(tmp_title))
    except gspread.exceptions.WorksheetNotFound:
        pass
    tmp = ss.add_worksheet(title=tmp_title, rows="100", cols=str(max(len(cols), 6)))
    try:
        _su_fill(tmp, cols, values, log, tab, vio)
    except Exception as e:
        try:
            ss.del_worksheet(tmp)
        except Exception:
            pass
        if "limit of" in str(e) and "cells" in str(e):
            raise RuntimeError(f"ملف Google Sheet وصل حد الـ 20 مليون خلية — محتاج {len(values) * len(cols):,} خلية لتاب {tab}. امسح تابات/صفوف قديمة وجرّب تاني") from None
        raise
    ss.del_worksheet(ws)
    tmp.update_title(tab)


def sheet_replace_safe(ss, tab, cols, values, log, vio="RAW"):
    """للتابات اللي بتتبني كلها كل مرة (مخزون/كتالوج/كاش): بناء تاب جديد كامل ثم تبديل — مفيش مسح قبل الكتابة."""
    import gspread
    cols = list(cols)
    if not values:
        log(f"⚠️ {tab}: مفيش بيانات — الشيت ما اتلمسش")
        return
    if pg_side(log, lambda: pg_write_tab(tab, cols, values, log)):
        return
    try:
        ws = ss.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab, rows="100", cols=str(max(len(cols), 6)))
        _su_fill(ws, cols, values, log, tab, vio)
        log(f"🆕 {tab}: اتبنى من الصفر — {len(values)} صف")
        return
    _su_swap(ss, ws, tab, cols, values, log, vio)
    log(f"🔁 {tab}: اتبنى كامل واتبدّل — {len(values)} صف")


def sheet_upsert(ss, tab, cols, values, key_cols, log, orphan_set=None, scope=None, since=None, vio="RAW", cmp_cols=None):
    """يضيف الصفوف الجديدة ويحدّث اللي اتغيّر (بالمفتاح key_cols) من غير ما يمسح حاجة.
    cmp_cols: الأعمدة اللي تغيّرها يعتبر الصف اتحدّث (الافتراضي كلها).
    orphan_set/scope/since: صفوف في الشيت ومش في البيانات الجديدة (في متاجر scope=(عمود, قيم) ومن تاريخ since=(عمود, تاريخ))
    بتتعلّم بالقيم دي بدل ما تتمسح."""
    import gspread
    from gspread.utils import rowcol_to_a1
    if not values:
        log(f"⚠️ {tab}: مفيش بيانات — الشيت ما اتلمسش")
        return
    if pg_side(log, lambda: (pg_write_tab(tab, cols, values, log, keys=key_cols, orphan_set=orphan_set, scope=scope, since=since)
                             if orphan_set else pg_push(tab, cols, values, key_cols, log))):
        return
    cols = list(cols)
    try:
        ws = ss.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        ws = None
    ex = _su_retry(lambda: ws.get_all_values(), log, tab) if ws is not None else []
    hdr = [str(h).strip() for h in ex[0]] if ex else []
    has_data = len(ex) > 1
    if ws is None:
        ws = ss.add_worksheet(title=tab, rows="100", cols=str(max(len(cols), 6)))
        _su_fill(ws, cols, values, log, tab, vio)
        log(f"🆕 {tab}: اتبنى من الصفر — {len(values)} صف")
        return
    if not hdr or not has_data:
        _su_fill(ws, cols, values, log, tab, vio)
        log(f"🆕 {tab}: كان فاضي — اتكتب {len(values)} صف")
        return
    if not all(k in cols and k in hdr for k in key_cols):     # أول مرة بعد إضافة عمود المفتاح: بناء كامل + تبديل
        _su_swap(ss, ws, tab, cols, values, log, vio)
        log(f"🔁 {tab}: اتبنى من جديد بعمود المفتاح — {len(values)} صف")
        return
    new_cols = [c for c in cols if c not in hdr]
    fh = hdr + new_cols
    if new_cols:
        if ws.col_count < len(fh):
            _su_retry(lambda: ws.resize(cols=len(fh)), log, tab)
        _su_retry(lambda: ws.update(range_name="A1", values=[fh], value_input_option="RAW"), log, tab)
    W = len(fh)
    pos = {c: i for i, c in enumerate(fh)}
    kix = [pos[k] for k in key_cols]
    idx = {}
    for i, row in enumerate(ex[1:], start=2):
        row = list(row) + [""] * (W - len(row))
        k = "|".join(str(row[j]).strip() for j in kix)
        if k.strip("|") and k not in idx:
            idx[k] = (i, row)
    ci = {c: j for j, c in enumerate(cols)}
    newmap = {}
    for v in values:
        newmap["|".join(str(v[ci[k]]).strip() for k in key_cols)] = v
    cc = [c for c in (cmp_cols or cols) if c in ci]
    ups, adds, same, orph, seen = [], [], 0, 0, set()
    _dc, _dx = {}, {}
    for k, v in newmap.items():
        if k in idx:
            i, row = idx[k]
            seen.add(k)
            _dif = [c for c in cc if _su_norm(row[pos[c]]) != _su_norm(v[ci[c]])]
            if _dif:
                for c in _dif:                                   # تشخيص: أنهي أعمدة بتفرّق (عشان نعرف ليه صفوف كتير بتتحدّث)
                    _dc[c] = _dc.get(c, 0) + 1
                    if c not in _dx:
                        _dx[c] = (str(row[pos[c]])[:40], str(v[ci[c]])[:40])
                tgt = list(row)
                for c in cols:
                    tgt[pos[c]] = v[ci[c]]
                ups.append((i, tgt))
            else:
                same += 1
        else:
            adds.append([v[ci[c]] if c in ci else "" for c in fh])
    if orphan_set and scope:
        sc_col, sc_vals = scope
        for k, (i, row) in idx.items():
            if k in seen or str(row[pos[sc_col]]).strip() not in sc_vals:
                continue
            if since and str(row[pos[since[0]]]).strip()[:10] < since[1]:
                continue
            if any(_su_norm(row[pos[c]]) != _su_norm(x) for c, x in orphan_set.items() if c in pos):
                tgt = list(row)
                for c, x in orphan_set.items():
                    if c in pos:
                        tgt[pos[c]] = x
                ups.append((i, tgt))
                orph += 1
    last = "".join(ch for ch in rowcol_to_a1(1, W) if not ch.isdigit())
    for s in range(0, len(ups), 500):
        data = [{"range": f"A{i}:{last}{i}", "values": [r]} for i, r in ups[s:s + 500]]
        _su_retry(lambda d=data: ws.batch_update([dict(x) for x in d], value_input_option=vio), log, tab)
    for s in range(0, len(adds), 2000):
        part = adds[s:s + 2000]
        _su_retry(lambda p=part: ws.append_rows(p, value_input_option=vio), log, tab)
    if _dc and len(ups) - orph > 200:                            # كتير اتحدّث: اطبع أكتر الأعمدة اختلافاً مع مثال (الشيت ← الجديد)
        for c, n in sorted(_dc.items(), key=lambda x: -x[1])[:5]:
            log(f"🔎 {tab}: عمود «{c}» اختلف في {n} صف — مثال: «{_dx[c][0]}» ← «{_dx[c][1]}»")
    log(f"⬆️ {tab}: {len(adds)} جديد · {len(ups) - orph} اتحدّث" + (f" · {orph} اتعلّموا" if orph else "") + f" · {same} زي ما هو — الإجمالي في الشيت {len(idx) + len(adds)}")


def push_sheet(store_names, log):
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    ms = glob.glob(os.path.join(GSHEET_CREDS_DIR, "*.json"))
    if not ms:
        raise RuntimeError(f"مفيش ملف اعتماد Google في {GSHEET_CREDS_DIR}")
    ss = gspread.authorize(ServiceAccountCredentials.from_json_keyfile_name(ms[0], GSHEET_SCOPE)).open(GSHEET_NAME)
    allrows, cols = [], ["Store"]
    for sn in store_names:
        rd, c = read_ledger(os.path.join(LEDGER_DIR, sn, "finance.csv"))
        for c1 in c:
            if c1 not in cols:
                cols.append(c1)
        for r in rd:
            r["Store"] = sn
            allrows.append(r)
    if not allrows:
        log("⚠️ مفيش بيانات تتكتب في الشيت")
        return
    values = []
    for r in allrows:
        row = []
        for c in cols:
            v = r.get(c, "")
            if c in NUM_COLS:
                try:
                    v = float(v)
                except Exception:
                    pass
            row.append(v)
        values.append(row)
    sheet_upsert(ss, TAB_FIN, cols, values, ("Store", "id"), log)


# ═══════════════════════ الواجهة / التشغيل التلقائي ═══════════════════════
def load_stores():
    out = []
    try:
        for d in json.load(open(STORES_FILE, encoding="utf-8")):
            try:
                out.append(TStore(d))
            except Exception:
                pass
    except Exception:
        pass
    return out


def auto_main():
    stop = threading.Event()
    log = lambda m: print(time.strftime("%H:%M:%S"), m, flush=True)
    names = []
    for st in load_stores():
        try:
            names.append(st.name)
            run_store(st, "daily", DEFAULT_START, log, stop)
        except Exception as e:
            log(f"❌ {st.name}: {type(e).__name__}: {e}")
    if "--sheet" in sys.argv and names:
        push_sheet(names, log)


import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox


class App(tk.Frame):
    def __init__(self, root):
        super().__init__(root, bg="#d9d9d9")
        self.pack(fill="both", expand=True, padx=8, pady=8)
        self.stop = threading.Event()
        self.stores = []
        self.running = False
        self._last_fixed = ""
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)
        ttk.Label(top, text="المتاجر:").grid(row=0, column=0, sticky="nw")
        self.lb = tk.Listbox(top, height=5, width=80)
        self.lb.grid(row=0, column=1, padx=6)
        bb = ttk.Frame(top)
        bb.grid(row=0, column=2, sticky="n")
        ttk.Button(bb, text="➕ إضافة", command=self.add_store).pack(fill="x")
        ttk.Button(bb, text="➖ حذف", command=self.rm_store).pack(fill="x", pady=2)
        ttk.Label(top, text="تاريخ البداية:").grid(row=1, column=0, sticky="w", pady=6)
        self.dfrom = tk.StringVar(value=DEFAULT_START.isoformat())
        ttk.Entry(top, textvariable=self.dfrom, width=14).grid(row=1, column=1, sticky="w", padx=6)
        self.sheet = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text=f"اكتب كمان في Google Sheet ({TAB_FIN})", variable=self.sheet).grid(row=2, column=1, sticky="w")
        self.auto = tk.BooleanVar(value=False)
        self.times = tk.StringVar(value="")
        ttk.Checkbutton(top, text="⏰ تحديث يومي تلقائي في المواعيد دي (HH:MM مفصولة بفاصلة):", variable=self.auto, command=self.save_settings).grid(row=3, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Entry(top, textvariable=self.times, width=24).grid(row=3, column=2, sticky="w")
        self.next_lbl = ttk.Label(top, text="")
        self.next_lbl.grid(row=4, column=1, sticky="w")
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=10)
        self.btns = []
        for txt, mode in (("🔄 تحديث يومي (آخر فترة)", "daily"), ("📥 سحب كامل من تاريخ البداية", "full")):
            b = ttk.Button(bar, text=txt, command=lambda m=mode: self.go(m))
            b.pack(side="left", padx=4, pady=6)
            self.btns.append(b)
        ttk.Button(bar, text="⏹ إيقاف", command=self.stop.set).pack(side="left", padx=4)
        self.box = scrolledtext.ScrolledText(self, font=("Consolas", 10), height=16)
        self.box.pack(fill="both", expand=True, padx=10, pady=8)
        for st in load_stores():
            self.stores.append(st)
            self.lb.insert("end", f"{st.name} — seller {st.sid} — {st.sf}")
        self.load_settings()
        for _v in (self.sheet, self.times, self.dfrom):
            _v.trace_add("write", lambda *a: self.save_settings())
        self.log("جاهز. جرّب «سحب كامل» أول مرة، بعدها اليومي بيحدّث بس.")
        self.after(5000, self.tick)

    def log(self, m):
        try:
            if LOG_HOOK:
                LOG_HOOK(str(m) + "\n")
        except Exception:
            pass
        try:
            self.after(0, lambda: (self.box.insert("end", str(m) + "\n"), self.box.see("end")))
        except Exception:
            pass

    def load_settings(self):
        try:
            s = json.load(open(SETTINGS_FILE, encoding="utf-8"))
            self.sheet.set(bool(s.get("sheet", False)))
            self.times.set(str(s.get("times", "")))
            self.auto.set(bool(s.get("auto", False)))
            if s.get("dfrom"):
                self.dfrom.set(str(s["dfrom"]))
        except Exception:
            pass

    def save_settings(self):
        try:
            json.dump({"sheet": bool(self.sheet.get()), "times": self.times.get(), "auto": bool(self.auto.get()), "dfrom": self.dfrom.get()},
                      open(SETTINGS_FILE, "w", encoding="utf-8"), ensure_ascii=False)
        except Exception:
            pass

    def fixed_times(self):
        out = []
        for t in str(self.times.get()).replace("،", ",").replace(" ", ",").split(","):
            t = t.strip()
            try:
                if t:
                    out.append(dt.datetime.strptime(t, "%H:%M").strftime("%H:%M"))
            except ValueError:
                pass
        return out

    def tick(self):
        try:
            ft = self.fixed_times()
            self.next_lbl.configure(text=("المواعيد: " + " ، ".join(ft)) if (self.auto.get() and ft) else "")
            now = time.strftime("%H:%M")
            key = time.strftime("%Y-%m-%d ") + now
            if self.auto.get() and now in ft and key != self._last_fixed and not self.running:
                self._last_fixed = key
                self.log(f"🕐 موعد محدد {now} — تحديث يومي تلقائي...")
                self.go("daily", auto=True)
        except Exception:
            pass
        self.after(20000, self.tick)

    def save(self):
        try:
            json.dump([s.as_dict() for s in self.stores], open(STORES_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        except Exception:
            pass

    def add_store(self):
        w = tk.Toplevel(self)
        w.title("متجر تريندول")
        f = {}
        for i, (k, t, show) in enumerate([("name", "اسم المتجر", ""), ("seller_id", "Seller ID", ""), ("api_key", "API Key", ""),
                                          ("api_secret", "API Secret", "•"), ("store_front", "storeFrontCode", "")]):
            ttk.Label(w, text=t).grid(row=i, column=0, padx=6, pady=3, sticky="e")
            e = ttk.Entry(w, width=48, show=show)
            e.grid(row=i, column=1, padx=6, pady=3)
            f[k] = e
        f["store_front"].insert(0, "SA")

        def ok():
            d = {k: e.get().strip() for k, e in f.items()}
            if not (d["seller_id"] and d["api_key"] and d["api_secret"]):
                return
            s = TStore(d)
            self.stores.append(s)
            self.lb.insert("end", f"{s.name} — seller {s.sid} — {s.sf}")
            self.save()
            w.destroy()
        ttk.Button(w, text="حفظ", command=ok).grid(row=5, column=1, pady=8, sticky="e")

    def rm_store(self):
        for i in reversed(self.lb.curselection()):
            self.lb.delete(i)
            del self.stores[i]
        self.save()

    def go(self, mode, auto=False, sheet=None):
        if self.running:
            return
        if not self.stores:
            if not auto:
                messagebox.showwarning("", "ضيف متجر واحد على الأقل")
            return
        try:
            start = dt.date.fromisoformat(self.dfrom.get().strip())
        except Exception as e:
            if not auto:
                messagebox.showerror("خطأ", str(e))
            return
        self.stop.clear()
        self.running = True
        for b in self.btns:
            b.state(["disabled"])
        want_sheet = self.sheet.get() if sheet is None else bool(sheet)
        stores = list(self.stores)

        def work():
            try:
                for st in stores:
                    if self.stop.is_set():
                        break
                    try:
                        run_store(st, mode, start, self.log, self.stop)
                    except Exception as e:
                        self.log(f"❌ [{st.name}] {type(e).__name__}: {e}")
                if want_sheet and not self.stop.is_set():
                    self.log("⬆️ كتابة الشيت...")
                    push_sheet([s.name for s in stores], self.log)
                self.log("✅ خلص.")
            except Exception as e:
                self.log(f"❌ {type(e).__name__}: {e}")
            finally:
                self.running = False
                try:
                    self.after(0, lambda: [b.state(["!disabled"]) for b in self.btns])
                except Exception:
                    pass
        threading.Thread(target=work, daemon=True).start()

    def shutdown(self):
        self.stop.set()


def build_ui(root):
    """بيتنادى من البرنامج الموحد: بيبني الواجهة جوه التاب ويرجّع دالة إيقاف."""
    return App(root).shutdown


def main():
    root = tk.Tk()
    root.title("Trendyol Accounts — دفتر حسابات تريندول")
    root.geometry("1000x720")
    a = App(root)
    root.protocol("WM_DELETE_WINDOW", lambda: (a.shutdown(), root.destroy()))
    root.mainloop()


if __name__ == "__main__":
    if "--auto" in sys.argv:
        auto_main()
    else:
        main()
'''

_EMBEDDED['AmazonAccounts'] = r'''# -*- coding: utf-8 -*-
"""
Amazon_Accounts.py — بوت أمازون الرسمي (SP-API) بديل Amazon_Server.py، متعدد المتاجر. بيعمل 3 حاجات:
  1) المالية  → دفتر حسابات (Upsert) + مجموعات التسوية (الدفعات) → تابين في الشيت: Amazon_Finance_API و Amazon_Payouts_API
  2) الشحنات → تقرير الشحنات المسلّمة (GET_AMAZON_FULFILLED_SHIPMENTS_DATA_GENERAL) → إضافة تحت آخر صف في DailyOrdersAmazon (من غير تكرار)
  3) المخزون → تقرير مخزون FBA (GET_AFN_INVENTORY_DATA) → StockAmazon (مسح وكتابة، SELLABLE وكمية > 0 بس)
  4) الأوردرات بحالتها (جديد/اتشحن/ملغي...) → دفتر Upsert + تاب Amazon_Orders_API (زي Noon_Orders_API)

ملاحظات مهمة عن مالية أمازون السعودية (من التجربة الحقيقية):
  • الأوردر بيظهر مرتين: مرة وقت الشحن بحالة DEFERRED/DEFERRED_RELEASED، ومرة تانية وقت الإفراج (بعد ~8 أيام) بحالة RELEASED.
    لو اتحسبوا الاتنين المبيعات بتتضاعف، فالدفتر بيربطهم ويحط counted=1 على واحد بس (الأصلي بتاريخ الشحن).
  • مجموع كل صفوف RELEASED في مجموعة تسوية = قيمة الدفعة بالظبط (التحويل البنكي) — ده أساس صفحة التسوية.
  • المؤجّل (DEFERRED) = فلوس لسه محجوزة عند أمازون لحد الإفراج.

المخرجات (جنب الأداة):
  amazon_accounts/<المتجر>/finance.csv  — الدفتر (مفتاح: transactionId#رقم البند)
  amazon_accounts/<المتجر>/groups.csv   — مجموعات التسوية
  amazon_accounts/state.json            — آخر تحديث لكل متجر
المتاجر والمفاتيح: amazon_stores.json (نفس ملف أداة التجربة) — لا يترفع على GitHub ولا يتشارك مع الموظفين.
التشغيل:  py Amazon_Accounts.py                (واجهة)
          py Amazon_Accounts.py --auto [--sheet]   (بدون واجهة — للجدولة)
المتطلبات: py -m pip install requests gspread oauth2client
"""
import csv, io, gzip, json, os, re, sys, glob, time, threading, collections, datetime as dt
import requests

APP_DIR = os.environ.get("PLATFORMS_BASE_DIR") or os.path.dirname(os.path.abspath(__file__))
EMBEDDED = bool(os.environ.get("PLATFORMS_BASE_DIR"))
LOG_HOOK = None   # بيتحط من البرنامج الموحد → السجل المشترك
STORES_FILE = os.path.join(APP_DIR, "amazon_stores.json")
SETTINGS_FILE = os.path.join(APP_DIR, "amazon_accounts_settings.json")
LEDGER_DIR = os.path.join(APP_DIR, "amazon_accounts")
STATE_FILE = os.path.join(LEDGER_DIR, "state.json")
GSHEET_CREDS_DIR = r"\\WIN-181U7G4HFM5\Share For Markting\NOONIMG\Copy"
GSHEET_NAME = "Complaints"
GSHEET_SCOPE = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
LWA_URL = "https://api.amazon.com/auth/o2/token"
DEFAULT_ENDPOINT = "https://sellingpartnerapi-eu.amazon.com"
DEFAULT_MARKETPLACE = "A17E79C6D8DWNP"          # amazon.sa
TAB_FIN = "Amazon_Finance_API"
TAB_GROUPS = "Amazon_Payouts_API"
TAB_ORD = "Amazon_Orders_API"
ORD_REPORT = "GET_FLAT_FILE_ALL_ORDERS_DATA_BY_LAST_UPDATE_GENERAL"   # كل الأوردرات بحالتها (آخر تحديث)، بحد أقصى 30 يوم للطلب
ORD_LOOKBACK = 21        # يوم — حالة الأوردر بتتغيّر بعد إنشائه (اتشحن/اتلغى)
ORD_PERSONAL = ("buyer", "phone", "email", "address", "city", "state", "postal", "recipient", "street", "zip", "bill", "customer", "-name")
ORD_SHEET_COLS = ["Store", "amazon-order-id", "purchase-date", "last-updated-date", "order-status", "item-status", "fulfillment-channel",
                  "sales-channel", "sku", "asin", "product-name", "quantity", "currency", "item-price", "item-tax", "shipping-price",
                  "item-promotion-discount", "ship-promotion-discount", "_first_seen", "_status_changed_at", "_key"]
ORD_NUM = ("quantity", "item-price", "item-tax", "shipping-price", "item-promotion-discount", "ship-promotion-discount")
TAB_IMG = "Amazon_Images_API"
IMG_COLS = ["ASIN", "Image", "Title", "Store"]
IMG_CACHE = os.path.join(LEDGER_DIR, "images.json")
IMG_REFRESH_DAYS = 14
ORDERS_SHEET = "DailyOrdersAmazon"
STOCK_SHEET = "StockAmazon"
# بداية ١٥ أغسطس عشان أول مجموعة تسوية (١٩/٨ → ٢/٩) تبقى كاملة، وعشان أوردرات اتشحنت قبل ١ سبتمبر واتفرج عنها بعده تلاقي أصلها
DEFAULT_START = dt.date(2026, 9, 1)
FIN_LOOKBACK = 45        # يوم — حالة المؤجّل بتتغيّر على الصف القديم (تاريخ الشحن) لما يتفرج عنه، والإفراج بياخد لحد ~18 يوم
WIN_DAYS = 30
TS_OFFSET_H = 3          # توقيت السعودية
ORDERS_HEADERS = ["تاريخ شحنة العميل", "رقم تخزين سلعة التاجر MSKU", "رقم تخزين السلعة لشبكة نظام الشحن (FNSKU)", "ASIN",
                  "حاوية كاملة الحمولة", "الكمية", "رقم الطلب من أمازون", "العملة", "مبلغ المنتج"]
ORDERS_KEY_COLS = (6, 1, 3, 5, 8)   # رقم الطلب + MSKU + ASIN + الكمية + المبلغ (منع التكرار)
STOCK_HEADERS = ["seller-sku", "fulfillment-channel-sku", "asin", "condition-type", "Warehouse-Condition-code", "Quantity Available"]
LEDGER_COLS = ["key", "tx_id", "type", "status", "posted", "order_id", "shipment_id", "group_id", "deferred_id", "sku", "asin", "title",
               "qty", "currency", "product_charges", "shipping", "promo", "commission", "fba_fee", "other_fees", "total",
               "counted", "pay_group", "pay_date", "_first_seen"]
SHEET_COLS = ["Store", "type", "status", "posted", "order_id", "group_id", "sku", "asin", "title", "qty", "product_charges",
              "shipping", "promo", "commission", "fba_fee", "other_fees", "total", "counted", "pay_group", "key"]
NUM_COLS = ("qty", "product_charges", "shipping", "promo", "commission", "fba_fee", "other_fees", "total", "counted")
GROUP_COLS = ["group_id", "start", "end", "processing", "transfer_status", "total", "currency", "transfer_date", "begin_balance"]
GROUP_SHEET_COLS = ["Store"] + GROUP_COLS + ["released_sum", "released_n"]
_slock = threading.Lock()


# ═══════════════════════ الاتصال بـ SP-API ═══════════════════════
class ApiError(Exception):
    def __init__(self, code, msg):
        super().__init__(msg)
        self.code = code


class AStore:
    def __init__(self, d):
        self.name = re.sub(r'[\\/:*?"<>|]+', "_", (d.get("name") or "amazon").strip())
        self.client_id = d["client_id"].strip()
        self.client_secret = d["client_secret"].strip()
        self.refresh = d["refresh_token"].strip()
        self.marketplace = (d.get("marketplace") or DEFAULT_MARKETPLACE).strip()
        self.endpoint = (d.get("endpoint") or DEFAULT_ENDPOINT).strip().rstrip("/")
        self.to_orders = bool(d.get("to_orders", True))
        self.to_stock = bool(d.get("to_stock", True))
        self._tok, self._exp = None, 0
        self.sess = requests.Session()

    def as_dict(self):
        return {"name": self.name, "client_id": self.client_id, "client_secret": self.client_secret, "refresh_token": self.refresh,
                "marketplace": self.marketplace, "endpoint": self.endpoint, "to_orders": self.to_orders, "to_stock": self.to_stock}

    def token(self):
        if self._tok and time.time() < self._exp - 60:
            return self._tok
        r = self.sess.post(LWA_URL, data={"grant_type": "refresh_token", "refresh_token": self.refresh,
                                          "client_id": self.client_id, "client_secret": self.client_secret}, timeout=30)
        if r.status_code != 200:
            raise ApiError(r.status_code, f"LWA فشل: {(r.text or '')[:250]}")
        j = r.json()
        self._tok, self._exp = j["access_token"], time.time() + int(j.get("expires_in", 3600))
        return self._tok


def api_get(st, path, params, tries=6):
    last = ""
    for a in range(tries):
        try:
            r = st.sess.get(st.endpoint + path, params=params, headers={"x-amz-access-token": st.token(), "Accept": "application/json"}, timeout=90)
        except ApiError:
            raise
        except Exception as e:
            last = f"{type(e).__name__}: {str(e)[:120]}"
            time.sleep(3 * (a + 1))
            continue
        if r.status_code == 200:
            return r.json()
        last = f"HTTP {r.status_code}: {(r.text or '')[:300].replace(chr(10), ' ')}"
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(4 * (a + 1))
            continue
        raise ApiError(r.status_code, last)
    raise ApiError(0, f"فشل بعد {tries} محاولات — {last}")


def api_post(st, path, body, tries=6):
    last = ""
    for a in range(tries):
        try:
            r = st.sess.post(st.endpoint + path, json=body, headers={"x-amz-access-token": st.token(), "Accept": "application/json"}, timeout=90)
        except ApiError:
            raise
        except Exception as e:
            last = f"{type(e).__name__}: {str(e)[:120]}"
            time.sleep(3 * (a + 1))
            continue
        if r.status_code in (200, 201, 202):
            return r.json()
        last = f"HTTP {r.status_code}: {(r.text or '')[:300].replace(chr(10), ' ')}"
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(20 * (a + 1))
            continue
        raise ApiError(r.status_code, last)
    raise ApiError(0, f"فشل بعد {tries} محاولات — {last}")


# ═══════════════════════ أدوات ═══════════════════════
def local_ts(s):
    """وقت ISO (UTC) → توقيت السعودية 'YYYY-MM-DD HH:MM:SS'."""
    try:
        d = dt.datetime.fromisoformat(str(s).replace("Z", "+00:00"))
        if d.tzinfo is not None:
            d = d.astimezone(dt.timezone.utc).replace(tzinfo=None)
        return (d + dt.timedelta(hours=TS_OFFSET_H)).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return ""


def iso(d, end=False):
    t = dt.datetime.combine(d, dt.time.max if end else dt.time.min)
    return t.strftime("%Y-%m-%dT%H:%M:%SZ")


def windows(d_from, d_to):
    now = dt.datetime.utcnow() - dt.timedelta(minutes=5)
    a = d_from
    while a <= d_to:
        b = min(a + dt.timedelta(days=WIN_DAYS), d_to + dt.timedelta(days=1))
        pa = dt.datetime.combine(a, dt.time.min)
        pb = min(dt.datetime.combine(b, dt.time.min), now)
        if pb > pa:
            yield pa.strftime("%Y-%m-%dT%H:%M:%SZ"), pb.strftime("%Y-%m-%dT%H:%M:%SZ")
        a = b


def amt(x):
    if isinstance(x, dict):
        v = x.get("currencyAmount", x.get("CurrencyAmount"))
        try:
            return float(v)
        except Exception:
            return None
    try:
        return float(x)
    except Exception:
        return None


def f2(x):
    try:
        return round(float(x), 2)
    except Exception:
        return 0.0


def rel(ids):
    d = {}
    for i in ids or []:
        d[str(i.get("relatedIdentifierName", ""))] = str(i.get("relatedIdentifierValue", ""))
    return d


def leaves(bks, path, acc):
    for b in bks or []:
        p = path + [str(b.get("breakdownType", ""))]
        kids = b.get("breakdowns") or []
        if kids:
            leaves(kids, p, acc)
        else:
            acc.append((p, amt(b.get("breakdownAmount")) or 0.0))


def tx_rows(t):
    """معاملة واحدة → صفوف (صف لكل بند). الأرقام: سعر/شحن/خصم/عمولة/رسوم FBA/باقي الرسوم (= الإجمالي − اللي فوق)."""
    rid = rel(t.get("relatedIdentifiers"))
    tid = str(t.get("transactionId", ""))
    cur = (t.get("totalAmount") or {}).get("currencyCode", "") if isinstance(t.get("totalAmount"), dict) else ""
    base = {"tx_id": tid, "type": str(t.get("transactionType", "")), "status": str(t.get("transactionStatus", "")),
            "posted": local_ts(t.get("postedDate", "")), "order_id": rid.get("ORDER_ID", ""), "shipment_id": rid.get("SHIPMENT_ID", ""),
            "group_id": rid.get("FINANCIAL_EVENT_GROUP_ID", ""), "deferred_id": rid.get("DEFERRED_TRANSACTION_ID", ""), "currency": cur}
    out = []
    items = t.get("items") or []
    if not items:
        tot = amt(t.get("totalAmount")) or 0.0
        r = dict(base)
        r.update({"key": tid + "#0", "sku": "", "asin": "", "title": str(t.get("description", ""))[:100], "qty": "",
                  "product_charges": 0.0, "shipping": 0.0, "promo": 0.0, "commission": 0.0, "fba_fee": 0.0, "other_fees": f2(tot), "total": f2(tot)})
        return [r]
    for i, it in enumerate(items):
        sku = asin = qty = ""
        for c in it.get("contexts") or []:
            sku = sku or str(c.get("sku", "") or "")
            asin = asin or str(c.get("asin", "") or "")
            qty = qty or str(c.get("quantityShipped", "") or "")
        acc = []
        leaves(it.get("breakdowns"), [], acc)
        pc = sh = pr = cm = fb = 0.0
        for p, v in acc:
            if p[0] == "ProductCharges":
                pc += v
            elif p[0] == "Shipping":
                sh += v
            elif p[0] == "PromoRebates":
                pr += v
            elif p[:2] == ["AmazonFees", "Commission"]:
                cm += v
            elif p[:2] == ["AmazonFees", "FBAPerUnitFulfillmentFee"]:
                fb += v
        tot = amt(it.get("totalAmount"))
        if tot is None:
            tot = sum(v for _, v in acc)
        r = dict(base)
        r.update({"key": f"{tid}#{i}", "sku": sku, "asin": asin, "title": str(it.get("description", "") or t.get("description", ""))[:70],
                  "qty": qty, "product_charges": f2(pc), "shipping": f2(sh), "promo": f2(pr), "commission": f2(cm), "fba_fee": f2(fb),
                  "other_fees": f2(tot - pc - sh - pr - cm - fb), "total": f2(tot)})
        out.append(r)
    return out


def fetch_transactions(st, d_from, d_to, log, stop):
    """يرجع (معاملات, تواريخ بداية الفترات الفاشلة)."""
    out, failed = [], []
    for pa, pb in windows(d_from, d_to):
        if stop.is_set():
            failed.append(dt.date.fromisoformat(pa[:10]))
            continue
        token, n, part = None, 0, []
        try:
            while True:
                params = {"nextToken": token} if token else {"postedAfter": pa, "postedBefore": pb, "marketplaceId": st.marketplace}
                j = api_get(st, "/finances/2024-06-19/transactions", params)
                p = j.get("payload", j)
                items = p.get("transactions") or []
                part.extend(items)
                n += len(items)
                token = p.get("nextToken")
                if not token:
                    break
                time.sleep(2)        # الحد الأقصى 0.5 طلب/ثانية
            log(f"[{st.name}]    {pa[:10]} → {pb[:10]}: {n} معاملة")
        except ApiError as e:
            log(f"[{st.name}] ⚠️ المالية {pa[:10]} → {pb[:10]}: {str(e)[:150]}")
            failed.append(dt.date.fromisoformat(pa[:10]))
        out.extend(part)
        time.sleep(2)
    return out, failed


def fetch_groups(st, d_from, log):
    out, token = [], None
    try:
        while True:
            params = {"nextToken": token} if token else {"FinancialEventGroupStartedAfter": iso(d_from), "MaxResultsPerPage": 100}
            j = api_get(st, "/finances/v0/financialEventGroups", params)
            p = j.get("payload", j)
            out.extend(p.get("FinancialEventGroupList") or [])
            token = p.get("NextToken")
            if not token:
                break
            time.sleep(2)
    except ApiError as e:
        log(f"[{st.name}] ⚠️ مجموعات التسوية: {str(e)[:150]}")
        return None
    return out


def group_row(g):
    ot = g.get("OriginalTotal") or {}
    bb = g.get("BeginningBalance") or {}
    return {"group_id": str(g.get("FinancialEventGroupId", "")), "start": local_ts(g.get("FinancialEventGroupStart", ""))[:10],
            "end": local_ts(g.get("FinancialEventGroupEnd", ""))[:10], "processing": str(g.get("ProcessingStatus", "")),
            "transfer_status": str(g.get("FundTransferStatus", "")), "total": ot.get("CurrencyAmount", ""),
            "currency": ot.get("CurrencyCode", ""), "transfer_date": local_ts(g.get("FundTransferDate", ""))[:10],
            "begin_balance": bb.get("CurrencyAmount", "")}


# ═══════════════════════ الدفتر ═══════════════════════
def read_ledger(path):
    if not os.path.exists(path):
        return [], []
    with open(path, encoding="utf-8-sig", newline="") as fh:
        rd = csv.DictReader(fh)
        cols = list(rd.fieldnames or [])
        return list(rd), cols


def write_ledger(path, rows, cols):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    os.replace(tmp, path)


def load_state():
    try:
        return json.load(open(STATE_FILE, encoding="utf-8"))
    except Exception:
        return {}


def save_state(s):
    os.makedirs(LEDGER_DIR, exist_ok=True)
    with _slock:
        json.dump(s, open(STATE_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def num(v):
    try:
        return float(v)
    except Exception:
        return 0.0


def derive(rows):
    """يربط معاملة الإفراج (RELEASED + DEFERRED_TRANSACTION_ID) بأصلها المؤجّل:
       الأصل counted=1 (تاريخ الشحن) + pay_group/pay_date من الإفراج، ومعاملة الإفراج counted=0 (لأنها تكرار).
       ربط أول بالـ id، ولو ماتلاقاش بتجرّب (نوع + أوردر + SKUs + الإجمالي). اللي ملهوش أصل بيتحسب لوحده.
       يرجع إحصائيات للسجل."""
    by_tx = collections.OrderedDict()
    for r in rows:
        by_tx.setdefault(r["tx_id"], []).append(r)

    def txsig(rs):
        return (rs[0]["type"], rs[0]["order_id"], tuple(sorted(x["sku"] for x in rs)), round(sum(num(x["total"]) for x in rs), 2))

    orig = {tid: rs for tid, rs in by_tx.items() if rs[0]["status"] in ("DEFERRED", "DEFERRED_RELEASED")}
    pool = collections.defaultdict(list)      # توقيع → أصول مؤجّلة اتفرج عنها ولسه ماترّبطتش
    for tid, rs in orig.items():
        if rs[0]["status"] == "DEFERRED_RELEASED":
            pool[txsig(rs)].append(tid)
    for k in pool:
        pool[k].sort(key=lambda t: by_tx[t][0]["posted"])
    for r in rows:
        r["counted"] = 0 if r["type"] == "Transfer" else 1
        r["pay_group"], r["pay_date"] = "", ""
    linked, by_id, by_sig, orphan = set(), 0, 0, 0
    releases = [(rs[0]["posted"], tid, rs) for tid, rs in by_tx.items() if rs[0]["status"] == "RELEASED" and rs[0]["deferred_id"]]
    releases.sort(key=lambda x: x[0])
    for posted, tid, rs in releases:
        d = rs[0]["deferred_id"]
        o = None
        if d in orig and d not in linked:
            o, by_id = d, by_id + 1
        else:
            cand = pool.get(txsig(rs)) or []
            for c in cand:
                if c not in linked:
                    o, by_sig = c, by_sig + 1
                    break
        if o is None:
            orphan += 1
            for x in rs:
                x["pay_group"], x["pay_date"] = x["group_id"], x["posted"]
            continue
        linked.add(o)
        for x in rs:
            x["counted"] = 0
        for x in orig[o]:
            x["pay_group"], x["pay_date"] = rs[0]["group_id"], posted
    for tid, rs in by_tx.items():
        s = rs[0]["status"]
        if s == "RELEASED" and not rs[0]["deferred_id"]:
            for x in rs:
                x["pay_group"], x["pay_date"] = x["group_id"], x["posted"]
    unlinked = sum(1 for tid, rs in orig.items() if rs[0]["status"] == "DEFERRED_RELEASED" and tid not in linked)
    held = round(sum(num(x["total"]) for tid, rs in orig.items() if rs[0]["status"] == "DEFERRED" for x in rs), 2)
    return {"by_id": by_id, "by_sig": by_sig, "orphan": orphan, "unlinked": unlinked, "held": held}


def update_finance(st, d_from, d_to, log, stop):
    """Upsert بالمفتاح (transactionId#بند). يرجع (جديد, اتحدّث, تواريخ فاشلة)."""
    path = os.path.join(LEDGER_DIR, st.name, "finance.csv")
    rows_old, cols = read_ledger(path)
    old = collections.OrderedDict((r["key"], r) for r in rows_old if r.get("key"))
    txs, failed = fetch_transactions(st, d_from, d_to, log, stop)
    new = upd = 0
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    for t in txs:
        for r in tx_rows(t):
            k = r["key"]
            if k in old:
                o = old[k]
                if any(str(o.get(c, "")) != str(r.get(c, "")) for c in ("status", "group_id", "total", "posted", "deferred_id")):
                    upd += 1
                o.update(r)
            else:
                r["_first_seen"] = now
                old[k] = r
                new += 1
    allrows = list(old.values())
    ds = derive(allrows)
    cols = list(LEDGER_COLS)
    write_ledger(path, allrows, cols)
    log(f"[{st.name}] 💰 المالية: {new} جديد، {upd} اتحدّث، الإجمالي {len(allrows)} صف، أجزاء فاشلة {len(failed)}")
    log(f"[{st.name}]    ربط الإفراج بالأصل: {ds['by_id']} بالـ id، {ds['by_sig']} بالتوقيع، {ds['orphan']} من غير أصل · "
        f"مؤجّل اتفرج عنه ومالوش إفراج: {ds['unlinked']} · محجوز لسه: {ds['held']:,.2f}")
    return new, upd, failed


def update_groups(st, d_from, log):
    gs = fetch_groups(st, d_from, log)
    if gs is None:
        return
    path = os.path.join(LEDGER_DIR, st.name, "groups.csv")
    old, _ = read_ledger(path)
    d = {r["group_id"]: r for r in old}
    for g in gs:
        r = group_row(g)
        if r["group_id"]:
            d[r["group_id"]] = r
    write_ledger(path, sorted(d.values(), key=lambda r: r["start"]), GROUP_COLS)
    log(f"[{st.name}] 🏦 مجموعات التسوية: {len(d)}")


# ═══════════════════════ دفتر الأوردرات (بحالتها) ═══════════════════════
def _ord_personal(c):
    c = str(c).lower()
    return c != "product-name" and any(w in c for w in ORD_PERSONAL)


def update_orders(st, d_from, d_to, log, stop):
    """أوردرات أمازون بحالتها (تقرير آخر تحديث، 30 يوم للطلب). Upsert بـ (أوردر|SKU|ASIN) مع أول ظهور وتغيّر الحالة. يرجع (جديد, اتغيّرت, أجزاء فاشلة)."""
    path = os.path.join(LEDGER_DIR, st.name, "orders.csv")
    rows_old, cols = read_ledger(path)
    old = collections.OrderedDict((r["_key"], r) for r in rows_old if r.get("_key"))
    new = chg = 0
    failed = []
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    a = d_from
    while a <= d_to:
        if stop.is_set():
            failed.append(a)
            break
        b = min(a + dt.timedelta(days=29), d_to)
        try:
            hdr, rows = download_report(st, ORD_REPORT, a, b, log, stop)
        except ApiError as e:
            log(f"[{st.name}] ⚠️ الأوردرات {a} → {b}: {str(e)[:200]}")
            failed.append(a)
            hdr, rows = [], []
        if hdr:
            names = [h.strip().lower() for h in hdr]
            seen = collections.Counter()
            for r in rows:
                d = {names[i]: (r[i].strip() if i < len(r) else "") for i in range(len(names)) if not _ord_personal(names[i])}
                if not d.get("amazon-order-id"):
                    continue
                for c in ("purchase-date", "last-updated-date"):
                    if d.get(c):
                        d[c] = local_ts(d[c])
                base = "%s|%s|%s" % (d["amazon-order-id"], d.get("sku", ""), d.get("asin", ""))
                seen[base] += 1
                k = base if seen[base] == 1 else f"{base}#{seen[base]}"
                d["_key"] = k
                o = old.get(k)
                if o is None:
                    d["_first_seen"], d["_prev_status"], d["_status_changed_at"] = now, "", now
                    old[k] = d
                    new += 1
                else:
                    if str(o.get("last-updated-date", "")) > str(d.get("last-updated-date", "")):
                        continue
                    if str(o.get("item-status", "")) != str(d.get("item-status", "")):
                        o["_prev_status"] = o.get("item-status", "")
                        o["_status_changed_at"] = now
                        chg += 1
                    o.update(d)
            log(f"[{st.name}] 🧾 أوردرات {a} → {b}: {len(rows)} صف")
        a = b + dt.timedelta(days=1)
        if a <= d_to:
            time.sleep(65)        # createReport: طلب كل دقيقة تقريباً
    allrows = list(old.values())
    cols = list(cols)
    for r in allrows:
        for c in r.keys():
            if c not in cols:
                cols.append(c)
    for c in ("_key", "_first_seen", "_prev_status", "_status_changed_at"):
        if c not in cols:
            cols.append(c)
    write_ledger(path, allrows, cols)
    log(f"[{st.name}] 🧾 الأوردرات: {new} جديد، {chg} تغيّرت حالتهم، الإجمالي {len(allrows)} صف، أجزاء فاشلة {len(failed)}")
    return new, chg, failed


def run_orders_ledger(st, mode, start, log, stop):
    today = dt.date.today()
    state = load_state()
    s = state.get(st.name, {})
    last_ok = s.get("orders_done")
    a = start if (mode == "full" or not last_ok or s.get("orders_from") != start.isoformat()) else max(start, dt.date.fromisoformat(last_ok) - dt.timedelta(days=ORD_LOOKBACK))
    for p in (s.get("orders_pending") or []):
        try:
            a = min(a, dt.date.fromisoformat(p))
        except Exception:
            pass
    log(f"[{st.name}] ▶️ الأوردرات (بحالتها): من {a} إلى {today}")
    n, c, failed = update_orders(st, a, today, log, stop)
    state = load_state()
    s = state.get(st.name, {})
    s["orders_pending"] = sorted({x.isoformat() for x in failed})
    if not failed and not stop.is_set():
        s["orders_done"] = today.isoformat()
        s["orders_from"] = start.isoformat()
    s["orders_at"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    state[st.name] = s
    save_state(state)


# ═══════════════════════ التقارير (شحنات + مخزون) ═══════════════════════
def download_report(st, rtype, d_from, d_to, log, stop, with_dates=True, wait_min=25):
    """يطلب تقرير ويستناه وينزّله. يرجع (أعمدة, صفوف). لو مفيش بيانات يرجع ([], [])."""
    body = {"reportType": rtype, "marketplaceIds": [st.marketplace]}
    if with_dates:
        body["dataStartTime"] = iso(d_from)
        body["dataEndTime"] = iso(d_to, True)
    j = api_post(st, "/reports/2021-06-30/reports", body)
    rid = j.get("reportId")
    log(f"[{st.name}] ⏳ تقرير {rtype} اتطلب ({rid})")
    doc = None
    for _ in range(wait_min * 6):
        if stop.is_set():
            raise ApiError(0, "اتوقّف")
        time.sleep(10)
        j = api_get(st, f"/reports/2021-06-30/reports/{rid}", {}, tries=4)
        s = j.get("processingStatus")
        if s == "DONE":
            doc = j.get("reportDocumentId")
            break
        if s == "CANCELLED":
            log(f"[{st.name}] ℹ️ تقرير {rtype}: اتلغى (غالباً مفيش بيانات في الفترة)")
            return [], []
        if s == "FATAL":
            raise ApiError(0, f"تقرير {rtype}: الحالة FATAL")
    if not doc:
        raise ApiError(0, f"تقرير {rtype}: ماجهزش في {wait_min} دقيقة")
    dj = api_get(st, f"/reports/2021-06-30/documents/{doc}", {}, tries=4)
    raw = requests.get(dj["url"], timeout=300).content
    if str(dj.get("compressionAlgorithm", "")).upper() == "GZIP" or raw[:2] == b"\x1f\x8b":
        raw = gzip.decompress(raw)
    try:
        txt = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        txt = raw.decode("cp1252", errors="replace")
    rows = list(csv.reader(io.StringIO(txt), delimiter="\t", quoting=csv.QUOTE_NONE))
    if not rows:
        return [], []
    return rows[0], rows[1:]


def _ix(hdr, name):
    low = [h.strip().lower() for h in hdr]
    return low.index(name) if name in low else None


def sku_maps(st):
    """SKU → (ASIN, FNSKU) من الدفتر (كل اللي اتباع) — بيتكمّل بتقرير المخزون."""
    m = {}
    rows, _ = read_ledger(os.path.join(LEDGER_DIR, st.name, "finance.csv"))
    for r in rows:
        if r.get("sku") and r.get("asin"):
            m.setdefault(r["sku"], [r["asin"], ""])
    return m


def run_orders(st, days, log, stop):
    """شحنات آخر `days` يوم → صفوف بترتيب ORDERS_HEADERS."""
    today = dt.date.today()
    start = today - dt.timedelta(days=max(1, int(days)))
    maps = sku_maps(st)
    out, a = [], start
    while a <= today:
        b = min(a + dt.timedelta(days=29), today)
        hdr, rows = download_report(st, "GET_AMAZON_FULFILLED_SHIPMENTS_DATA_GENERAL", a, b, log, stop)
        if hdr:
            ix = {k: _ix(hdr, k) for k in ("shipment-date", "purchase-date", "sku", "quantity-shipped", "amazon-order-id", "currency",
                                           "item-price", "fulfillment-center-id", "asin", "fnsku")}
            g = lambda r, k: (r[ix[k]].strip() if ix[k] is not None and ix[k] < len(r) else "")
            for r in rows:
                sku = g(r, "sku")
                if not sku:
                    continue
                asin = g(r, "asin") or (maps.get(sku) or ["", ""])[0]
                fn = g(r, "fnsku") or (maps.get(sku) or ["", ""])[1]
                out.append([local_ts(g(r, "shipment-date") or g(r, "purchase-date")), sku, fn, asin, g(r, "fulfillment-center-id"),
                            g(r, "quantity-shipped"), g(r, "amazon-order-id"), g(r, "currency"), g(r, "item-price")])
            log(f"[{st.name}] 📦 شحنات {a} → {b}: {len(rows)} صف")
        a = b + dt.timedelta(days=1)
        if a <= today:
            time.sleep(65)         # createReport: طلب كل دقيقة تقريباً
    return out


def run_stock(st, log, stop):
    """مخزون FBA: SELLABLE وكمية > 0 بس."""
    hdr, rows = download_report(st, "GET_AFN_INVENTORY_DATA", None, None, log, stop, with_dates=False)
    if not hdr:
        return []
    ix = [_ix(hdr, h.lower()) for h in STOCK_HEADERS]
    if any(i is None for i in ix):
        raise ApiError(0, "أعمدة تقرير المخزون مختلفة: " + " | ".join(hdr))
    out, skipped = [], 0
    for r in rows:
        r = [(r[i].strip() if i < len(r) else "") for i in ix]
        try:
            q = float(r[5].replace(",", "") or 0)
        except ValueError:
            q = 0
        if r[4].upper() != "SELLABLE" or q <= 0:
            skipped += 1
            continue
        out.append(r)
    log(f"[{st.name}] 🏬 مخزون: {len(out)} صف (SELLABLE وكمية > 0) — اتجاهل {skipped}")
    return out


# ═══════════════════════ تشغيل متجر ═══════════════════════
def run_store(st, mode, start, log, stop, do_fin=True):
    today = dt.date.today()
    state = load_state()
    s = state.get(st.name, {})
    if do_fin:
        last_ok = s.get("finance_done")
        a = start if (mode == "full" or not last_ok or s.get("finance_from") != start.isoformat()) else max(start, dt.date.fromisoformat(last_ok) - dt.timedelta(days=FIN_LOOKBACK))
        for p in (s.get("finance_pending") or []):
            try:
                a = min(a, dt.date.fromisoformat(p))
            except Exception:
                pass
        log(f"[{st.name}] ▶️ المالية: من {a} إلى {today}")
        n, c, failed = update_finance(st, a, today, log, stop)
        update_groups(st, start, log)
        s["finance_pending"] = sorted({x.isoformat() for x in failed})
        if not failed and not stop.is_set():
            s["finance_done"] = today.isoformat()
            s["finance_from"] = start.isoformat()
        s["finance_at"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
        state[st.name] = s
        save_state(state)
        return n, c, len(failed)
    return 0, 0, 0


# ═══════════════════════ Google Sheet ═══════════════════════
def open_sheet():
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    ms = glob.glob(os.path.join(GSHEET_CREDS_DIR, "*.json"))
    if not ms:
        raise RuntimeError(f"مفيش ملف اعتماد Google في {GSHEET_CREDS_DIR}")
    return gspread.authorize(ServiceAccountCredentials.from_json_keyfile_name(ms[0], GSHEET_SCOPE)).open(GSHEET_NAME)


def _ws(ss, title, ncols):
    import gspread
    try:
        return ss.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        return ss.add_worksheet(title=title, rows="2000", cols=str(max(ncols, 1)))


def _append(ws, values, log, tag, chunk=2000):
    for i in range(0, len(values), chunk):
        for attempt in range(5):
            try:
                ws.append_rows(values[i:i + chunk], value_input_option="RAW")
                break
            except Exception as e:
                log(f"⚠️ {tag}: دفعة {i // chunk + 1} محاولة {attempt + 1}: {str(e)[:150]}")
                time.sleep(2 * (attempt + 1))
        log(f"⬆️ {tag}: {min(i + chunk, len(values))}/{len(values)}")


# ═══════════════ PostgreSQL (Neon) — نفس اختيار تاب Noon API (ملف noon_use_pg.txt): sheet / pg / both لكل البوتات ═══════════════
import json
PG_MODE = "sheet"
PG_LABELS = {"sheet": "📄 Google Sheet فقط", "pg": "🐘 PostgreSQL فقط", "both": "📄+🐘 الاتنين مع بعض"}
PG_URL_FILE = os.path.join(APP_DIR, "pg_url.txt")
PG_FLAG_FILE = os.path.join(APP_DIR, "noon_use_pg.txt")
PG_SCHEMA = "sheet_copy"


def pg_flag_load():
    try:
        v = open(PG_FLAG_FILE, encoding="utf-8").read().strip()
        return {"1": "pg"}.get(v, v) if {"1": "pg"}.get(v, v) in PG_LABELS else "sheet"
    except Exception:
        return "sheet"


def pg_url():
    import re as _re
    u = os.environ.get("PG_URL", "").strip()
    if not u:
        try:
            u = open(PG_URL_FILE, encoding="utf-8").read().strip()
        except Exception:
            u = ""
    u = _re.sub(r"[&?]channel_binding=[^&]*", "", u.strip().strip('"'))
    if u and "?" not in u and "&" in u:
        u = u.replace("&", "?", 1)
    return u


def pg_conn():
    import psycopg2, re as _re
    u = pg_url()
    if not u:
        raise RuntimeError("مفيش رابط Neon — اضغط زر 🔑 رابط Neon")
    m = _re.search(r"postgres(?:ql)?://([^:@/\s]+):([^@\s]+)@([^:/?&\s'\"]+)(?::(\d+))?/([A-Za-z0-9_\-]+)", u)
    if not m:
        raise RuntimeError("رابط Neon مش مفهوم — لازم يبدأ بـ postgresql://user:password@host/dbname")
    return psycopg2.connect(user=m.group(1), password=m.group(2), host=m.group(3), port=int(m.group(4) or 5432),
                            dbname=_re.sub(r"(sslmode|channel_binding)$", "", m.group(5)) or "neondb", sslmode="require", connect_timeout=20)


def pg_side(log, fn):
    """بينفّذ كتابة PostgreSQL حسب الوضع المختار في تاب Noon API. بيرجع True لو الشيت يتخطّى (وضع PostgreSQL فقط)."""
    global PG_MODE
    PG_MODE = pg_flag_load()
    if PG_MODE == "sheet":
        return False
    try:
        fn()
    except Exception as e:
        if PG_MODE == "pg":
            raise
        log(f"⚠️ PostgreSQL فشل: {str(e)[:150]} — كمّلت على الشيت")
    return PG_MODE == "pg"


def _pg_snake(name, used):
    import re as _re
    s = _re.sub(r"[^0-9a-zA-Z]+", "_", str(name)).strip("_").lower()
    if not s:
        s = f"c_{len(used) + 1}"
    if s[0].isdigit():
        s = "c_" + s
    s = s[:60]
    base, i = s, 2
    while s in used:
        s, i = f"{base}_{i}", i + 1
    used.add(s)
    return s


def _pg_s(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def pg_write_tab(tab, headers, rows, log, keys=None, orphan_set=None, scope=None, since=None):
    """بيكتب التاب في PostgreSQL (schema sheet_copy — نفس شكل PG_Copy_All عشان التطبيق يقراه).
    keys → بناء وتحديث بالمفتاح (القديم بيفضل). من غيره → استبدال كامل. كله في transaction واحدة."""
    import io as _io, csv as _csv
    headers = [str(h).strip() for h in headers]
    used = set()
    cols = [_pg_snake(h, used) for h in headers]
    table = _pg_snake(tab, set())
    c = pg_conn()
    try:
        cur = c.cursor()
        cur.execute(f"CREATE SCHEMA IF NOT EXISTS {PG_SCHEMA}")
        cur.execute(f"CREATE TABLE IF NOT EXISTS {PG_SCHEMA}._meta (table_name TEXT PRIMARY KEY, tab_title TEXT, headers JSONB, col_names JSONB, copied_at TIMESTAMPTZ DEFAULT now())")
        cur.execute(f"SELECT headers, col_names FROM {PG_SCHEMA}._meta WHERE table_name=%s", (table,))
        m = cur.fetchone()
        old = []                                   # الصفوف القديمة كـ dict بالهيدر
        if m:
            ohdr, ocols = m
            cur.execute("SELECT " + ", ".join(f'"{x}"' for x in ocols) + f' FROM {PG_SCHEMA}."{table}" ORDER BY _row')
            old = [dict(zip(ohdr, ["" if v is None else v for v in r])) for r in cur.fetchall()]
        if not rows and m:
            log(f"⚠️ {tab}: مفيش بيانات — القاعدة ما اتلمستش")
            return
        new = [{h: _pg_s(v) for h, v in zip(headers, r)} for r in rows]
        if keys and old:
            kf = lambda d: "|".join(str(d.get(k, "")).strip() for k in keys)
            idx = {}
            for i, d in enumerate(old):
                k = kf(d)
                if k.strip("|") and k not in idx:
                    idx[k] = i
            final = [dict(d) for d in old]
            adds = ups = same = orph = 0
            seen = set()
            for d in new:
                k = kf(d)
                if k in idx:
                    seen.add(k)
                    cur_d = final[idx[k]]
                    if any(str(cur_d.get(h, "")).strip() != d[h].strip() for h in headers):
                        cur_d.update(d)
                        ups += 1
                    else:
                        same += 1
                else:
                    final.append(d)
                    adds += 1
            if orphan_set and scope:
                sc_col, sc_vals = scope
                for k, i in idx.items():
                    d = final[i]
                    if k in seen or str(d.get(sc_col, "")).strip() not in sc_vals:
                        continue
                    if since and str(d.get(since[0], "")).strip()[:10] < since[1]:
                        continue
                    if any(str(d.get(cn, "")).strip().lower() != str(x).strip().lower() for cn, x in orphan_set.items()):
                        d.update(orphan_set)
                        orph += 1
            body = final
            msg = f"{adds} جديد · {ups} اتحدّث" + (f" · {orph} اتعلّموا" if orph else "") + f" · {same} زي ما هو"
        else:
            body = new
            msg = f"اتبدّل بالكامل"
        cur.execute(f'DROP TABLE IF EXISTS {PG_SCHEMA}."{table}"')
        cur.execute(f'CREATE TABLE {PG_SCHEMA}."{table}" (' + ", ".join(f'"{x}" TEXT' for x in cols) + ", _row BIGSERIAL)")
        cur.execute(f"INSERT INTO {PG_SCHEMA}._meta (table_name, tab_title, headers, col_names, copied_at) VALUES (%s,%s,%s,%s,now()) "
                    f"ON CONFLICT (table_name) DO UPDATE SET tab_title=EXCLUDED.tab_title, headers=EXCLUDED.headers, col_names=EXCLUDED.col_names, copied_at=now()",
                    (table, tab, json.dumps(headers, ensure_ascii=False), json.dumps(cols)))
        if body:
            buf = _io.StringIO()
            w = _csv.writer(buf)
            for d in body:
                w.writerow([d.get(h, "") for h in headers])
            buf.seek(0)
            cur.copy_expert(f'COPY {PG_SCHEMA}."{table}" (' + ", ".join(f'"{x}"' for x in cols) + ") FROM STDIN WITH CSV", buf)
        c.commit()
        log(f"🐘 {tab}: {msg} — الإجمالي في PostgreSQL {len(body)}")
    except Exception:
        c.rollback()
        raise
    finally:
        c.close()


STORAGE_TAB = "Noon_Storage_API"
STORAGE_REPORT = "fbn_finance_monthlystoragechargereport"
STORAGE_HEADERS = ["Store", "soa_reference_nr", "description_en", "id_partner", "service_month", "pbarcode", "sku", "year_month",
                   "warehouse_code", "country_code", "average_stock_on_hand", "average_pending_qty", "average_net_qty", "longest_side",
                   "median_side", "shortest_side", "weight", "volumetric_weight", "item_volume", "estimated_total_item_volume",
                   "cost_per_cubic_feet", "charged_amount"]


def pg_push(tab, cols, values, key_cols, log):
    """بناء وتحديث بالمفتاح جوه PostgreSQL نفسها (من غير ما نحمّل التاب كله في الذاكرة): الصفوف اللي مفتاحها اتبعت بتتبدّل والباقي زي ما هو."""
    import io as _io, csv as _csv
    headers = [str(h).strip() for h in cols]
    table = _pg_snake(tab, set())
    c = pg_conn()
    try:
        cur = c.cursor()
        cur.execute(f"CREATE SCHEMA IF NOT EXISTS {PG_SCHEMA}")
        cur.execute(f"CREATE TABLE IF NOT EXISTS {PG_SCHEMA}._meta (table_name TEXT PRIMARY KEY, tab_title TEXT, headers JSONB, col_names JSONB, copied_at TIMESTAMPTZ DEFAULT now())")
        cur.execute(f"SELECT headers, col_names FROM {PG_SCHEMA}._meta WHERE table_name=%s", (table,))
        m = cur.fetchone()
        cur.execute("SELECT to_regclass(%s)", (f'{PG_SCHEMA}."{table}"',))
        if not m or cur.fetchone()[0] is None or not key_cols or any(k not in headers for k in key_cols):
            c.close()
            c = None
            return pg_write_tab(tab, cols, values, log)
        ohdr, ocols = [str(h) for h in m[0]], list(m[1])
        used = set(ocols)
        for h in headers:
            if h not in ohdr:
                cn = _pg_snake(h, used)
                cur.execute(f'ALTER TABLE {PG_SCHEMA}."{table}" ADD COLUMN "{cn}" TEXT')
                ohdr.append(h)
                ocols.append(cn)
        cmap = dict(zip(ohdr, ocols))
        ncols = [cmap[h] for h in headers]
        cur.execute("CREATE TEMP TABLE _pgn (" + ", ".join(f'"{x}" TEXT' for x in ncols) + ", _n BIGSERIAL) ON COMMIT DROP")
        buf = _io.StringIO()
        w = _csv.writer(buf)
        for r in values:
            w.writerow([_pg_s(v) for v in r])
        buf.seek(0)
        cur.copy_expert("COPY _pgn (" + ", ".join(f'"{x}"' for x in ncols) + ") FROM STDIN WITH CSV", buf)
        kc = [cmap[k] for k in key_cols]
        ke = lambda a: "(" + " || '|' || ".join(f"COALESCE(btrim({a}.\"{k}\"), '')" for k in kc) + ")"
        cur.execute(f'DELETE FROM {PG_SCHEMA}."{table}" m USING _pgn n WHERE {ke("m")} = {ke("n")} AND btrim({ke("n")}, \'|\') <> \'\'')
        old_n = cur.rowcount
        cl = ", ".join(f'"{x}"' for x in ncols)
        cur.execute(f'INSERT INTO {PG_SCHEMA}."{table}" ({cl}) SELECT {cl} FROM (SELECT *, row_number() OVER (PARTITION BY {ke("_pgn")} ORDER BY _n DESC) rn FROM _pgn) t '
                    f"WHERE rn = 1 OR btrim({ke('t')}, '|') = ''")
        new_n = cur.rowcount
        cur.execute(f"UPDATE {PG_SCHEMA}._meta SET headers=%s, col_names=%s, copied_at=now() WHERE table_name=%s",
                    (json.dumps(ohdr, ensure_ascii=False), json.dumps(ocols), table))
        c.commit()
        cur.execute(f'SELECT count(*) FROM {PG_SCHEMA}."{table}"')
        total = cur.fetchone()[0]
        log(f"🐘 {tab}: اتكتب {new_n:,} صف ({old_n:,} منهم كانوا موجودين واتبدّلوا) — الإجمالي في PostgreSQL {total:,}")
    finally:
        if c is not None:
            c.close()


# ═══════════════════════ كتابة الشيت: بناء وتحديث (Upsert) — عمر ما بيمسح الأول ═══════════════════════
def _su_norm(v):
    t = str(v if v is not None else "").strip()
    if t == "":
        return ""
    if t.lower() in ("true", "false"):                  # TRUE (الشيت) = True (بايثون)
        return t.lower()
    try:
        return repr(round(float(t.replace(",", "")), 6))
    except ValueError:
        return t


def _su_retry(fn, log, tag, tries=6):
    for a in range(tries):
        try:
            return fn()
        except Exception as e:
            if a == tries - 1:
                raise
            log(f"⚠️ {tag}: محاولة {a + 1} فشلت: {str(e)[:120]}")
            _m = str(e)
            time.sleep(65 if ("429" in _m or "Quota" in _m or "quota" in _m) else 3 * (a + 1))


def _su_fill(ws, cols, values, log, tag, vio="RAW", chunk=2000):
    """يكتب العنوان ثم الصفوف على دفعات — أي دفعة تفشل بعد 5 محاولات بترمي خطأ (مفيش صمت)."""
    _su_retry(lambda: ws.resize(rows=max(len(values) + 10, 100), cols=max(len(cols), 6)), log, tag)
    _su_retry(lambda: ws.update(range_name="A1", values=[list(cols)], value_input_option=vio), log, tag)
    for i in range(0, len(values), chunk):
        part = values[i:i + chunk]
        _su_retry(lambda p=part: ws.append_rows(p, value_input_option=vio), log, tag)
        log(f"⬆️ {tag}: {min(i + chunk, len(values))}/{len(values)}")


def _su_swap(ss, ws, tab, cols, values, log, vio):
    """يبني التاب كامل في تاب مؤقت، ولما يكتمل بس بيبدّل القديم — القديم عمره ما بيتمسح قبل ما الجديد يتكتب."""
    import gspread
    tmp_title = (tab + "_new")[:100]
    try:
        ss.del_worksheet(ss.worksheet(tmp_title))
    except gspread.exceptions.WorksheetNotFound:
        pass
    tmp = ss.add_worksheet(title=tmp_title, rows="100", cols=str(max(len(cols), 6)))
    try:
        _su_fill(tmp, cols, values, log, tab, vio)
    except Exception as e:
        try:
            ss.del_worksheet(tmp)
        except Exception:
            pass
        if "limit of" in str(e) and "cells" in str(e):
            raise RuntimeError(f"ملف Google Sheet وصل حد الـ 20 مليون خلية — محتاج {len(values) * len(cols):,} خلية لتاب {tab}. امسح تابات/صفوف قديمة وجرّب تاني") from None
        raise
    ss.del_worksheet(ws)
    tmp.update_title(tab)


def sheet_replace_safe(ss, tab, cols, values, log, vio="RAW"):
    """للتابات اللي بتتبني كلها كل مرة (مخزون/كتالوج/كاش): بناء تاب جديد كامل ثم تبديل — مفيش مسح قبل الكتابة."""
    import gspread
    cols = list(cols)
    if not values:
        log(f"⚠️ {tab}: مفيش بيانات — الشيت ما اتلمسش")
        return
    if pg_side(log, lambda: pg_write_tab(tab, cols, values, log)):
        return
    try:
        ws = ss.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab, rows="100", cols=str(max(len(cols), 6)))
        _su_fill(ws, cols, values, log, tab, vio)
        log(f"🆕 {tab}: اتبنى من الصفر — {len(values)} صف")
        return
    _su_swap(ss, ws, tab, cols, values, log, vio)
    log(f"🔁 {tab}: اتبنى كامل واتبدّل — {len(values)} صف")


def sheet_upsert(ss, tab, cols, values, key_cols, log, orphan_set=None, scope=None, since=None, vio="RAW", cmp_cols=None):
    """يضيف الصفوف الجديدة ويحدّث اللي اتغيّر (بالمفتاح key_cols) من غير ما يمسح حاجة.
    cmp_cols: الأعمدة اللي تغيّرها يعتبر الصف اتحدّث (الافتراضي كلها).
    orphan_set/scope/since: صفوف في الشيت ومش في البيانات الجديدة (في متاجر scope=(عمود, قيم) ومن تاريخ since=(عمود, تاريخ))
    بتتعلّم بالقيم دي بدل ما تتمسح."""
    import gspread
    from gspread.utils import rowcol_to_a1
    if not values:
        log(f"⚠️ {tab}: مفيش بيانات — الشيت ما اتلمسش")
        return
    if pg_side(log, lambda: (pg_write_tab(tab, cols, values, log, keys=key_cols, orphan_set=orphan_set, scope=scope, since=since)
                             if orphan_set else pg_push(tab, cols, values, key_cols, log))):
        return
    cols = list(cols)
    try:
        ws = ss.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        ws = None
    ex = _su_retry(lambda: ws.get_all_values(), log, tab) if ws is not None else []
    hdr = [str(h).strip() for h in ex[0]] if ex else []
    has_data = len(ex) > 1
    if ws is None:
        ws = ss.add_worksheet(title=tab, rows="100", cols=str(max(len(cols), 6)))
        _su_fill(ws, cols, values, log, tab, vio)
        log(f"🆕 {tab}: اتبنى من الصفر — {len(values)} صف")
        return
    if not hdr or not has_data:
        _su_fill(ws, cols, values, log, tab, vio)
        log(f"🆕 {tab}: كان فاضي — اتكتب {len(values)} صف")
        return
    if not all(k in cols and k in hdr for k in key_cols):     # أول مرة بعد إضافة عمود المفتاح: بناء كامل + تبديل
        _su_swap(ss, ws, tab, cols, values, log, vio)
        log(f"🔁 {tab}: اتبنى من جديد بعمود المفتاح — {len(values)} صف")
        return
    new_cols = [c for c in cols if c not in hdr]
    fh = hdr + new_cols
    if new_cols:
        if ws.col_count < len(fh):
            _su_retry(lambda: ws.resize(cols=len(fh)), log, tab)
        _su_retry(lambda: ws.update(range_name="A1", values=[fh], value_input_option="RAW"), log, tab)
    W = len(fh)
    pos = {c: i for i, c in enumerate(fh)}
    kix = [pos[k] for k in key_cols]
    idx = {}
    for i, row in enumerate(ex[1:], start=2):
        row = list(row) + [""] * (W - len(row))
        k = "|".join(str(row[j]).strip() for j in kix)
        if k.strip("|") and k not in idx:
            idx[k] = (i, row)
    ci = {c: j for j, c in enumerate(cols)}
    newmap = {}
    for v in values:
        newmap["|".join(str(v[ci[k]]).strip() for k in key_cols)] = v
    cc = [c for c in (cmp_cols or cols) if c in ci]
    ups, adds, same, orph, seen = [], [], 0, 0, set()
    _dc, _dx = {}, {}
    for k, v in newmap.items():
        if k in idx:
            i, row = idx[k]
            seen.add(k)
            _dif = [c for c in cc if _su_norm(row[pos[c]]) != _su_norm(v[ci[c]])]
            if _dif:
                for c in _dif:                                   # تشخيص: أنهي أعمدة بتفرّق (عشان نعرف ليه صفوف كتير بتتحدّث)
                    _dc[c] = _dc.get(c, 0) + 1
                    if c not in _dx:
                        _dx[c] = (str(row[pos[c]])[:40], str(v[ci[c]])[:40])
                tgt = list(row)
                for c in cols:
                    tgt[pos[c]] = v[ci[c]]
                ups.append((i, tgt))
            else:
                same += 1
        else:
            adds.append([v[ci[c]] if c in ci else "" for c in fh])
    if orphan_set and scope:
        sc_col, sc_vals = scope
        for k, (i, row) in idx.items():
            if k in seen or str(row[pos[sc_col]]).strip() not in sc_vals:
                continue
            if since and str(row[pos[since[0]]]).strip()[:10] < since[1]:
                continue
            if any(_su_norm(row[pos[c]]) != _su_norm(x) for c, x in orphan_set.items() if c in pos):
                tgt = list(row)
                for c, x in orphan_set.items():
                    if c in pos:
                        tgt[pos[c]] = x
                ups.append((i, tgt))
                orph += 1
    last = "".join(ch for ch in rowcol_to_a1(1, W) if not ch.isdigit())
    for s in range(0, len(ups), 500):
        data = [{"range": f"A{i}:{last}{i}", "values": [r]} for i, r in ups[s:s + 500]]
        _su_retry(lambda d=data: ws.batch_update([dict(x) for x in d], value_input_option=vio), log, tab)
    for s in range(0, len(adds), 2000):
        part = adds[s:s + 2000]
        _su_retry(lambda p=part: ws.append_rows(p, value_input_option=vio), log, tab)
    if _dc and len(ups) - orph > 200:                            # كتير اتحدّث: اطبع أكتر الأعمدة اختلافاً مع مثال (الشيت ← الجديد)
        for c, n in sorted(_dc.items(), key=lambda x: -x[1])[:5]:
            log(f"🔎 {tab}: عمود «{c}» اختلف في {n} صف — مثال: «{_dx[c][0]}» ← «{_dx[c][1]}»")
    log(f"⬆️ {tab}: {len(adds)} جديد · {len(ups) - orph} اتحدّث" + (f" · {orph} اتعلّموا" if orph else "") + f" · {same} زي ما هو — الإجمالي في الشيت {len(idx) + len(adds)}")


def _rewrite(ss, title, cols, values, log):
    ws = _ws(ss, title, len(cols))
    ws.clear()
    try:
        ws.resize(rows=2, cols=len(cols))      # يفضّي المساحة القديمة الأول
        ws.resize(rows=max(len(values) + 10, 100), cols=len(cols))
    except Exception as e:
        if "limit of" in str(e) and "cells" in str(e):
            raise RuntimeError(f"ملف Google Sheet وصل حد الـ 20 مليون خلية — محتاج {len(values) * len(cols):,} خلية لتاب {title}. "
                               f"امسح تابات/صفوف فاضية أو قديمة من الملف وجرّب تاني") from None
        raise
    ws.update(range_name="A1", values=[cols], value_input_option="RAW")
    _append(ws, values, log, title)


def push_finance(ss, store_names, log):
    allrows, grows = [], []
    for sn in store_names:
        rd, _ = read_ledger(os.path.join(LEDGER_DIR, sn, "finance.csv"))
        gs, gn = collections.defaultdict(float), collections.Counter()
        for r in rd:
            if r.get("status") == "RELEASED" and r.get("type") != "Transfer":     # مجموع الإفراج لكل دفعة (للتسوية) — بما فيه المكرر
                gs[r.get("group_id", "")] += num(r.get("total"))
                gn[r.get("group_id", "")] += 1
            if r.get("counted") == "1" and r.get("type") != "Transfer":           # الشيت فيه الصفوف المحسوبة بس (توفير في عدد الخلايا)
                r["Store"] = sn
                allrows.append(r)
        gd, _ = read_ledger(os.path.join(LEDGER_DIR, sn, "groups.csv"))
        for r in gd:
            r["Store"] = sn
            r["released_sum"] = round(gs.get(r["group_id"], 0.0), 2)
            r["released_n"] = gn.get(r["group_id"], 0)
            grows.append(r)
    if not allrows:
        log("⚠️ مفيش بيانات مالية تتكتب في الشيت")
        return
    vals = []
    for r in allrows:
        row = []
        for c in SHEET_COLS:
            v = r.get(c, "")
            if c in NUM_COLS and v != "":
                try:
                    v = float(v)
                except Exception:
                    pass
            row.append(v)
        vals.append(row)
    sheet_upsert(ss, TAB_FIN, SHEET_COLS, vals, ("Store", "key"), log, orphan_set={"counted": 0}, scope=("Store", set(store_names)))
    if grows:
        gv = []
        for r in grows:
            row = []
            for c in GROUP_SHEET_COLS:
                v = r.get(c, "")
                if c in ("total", "begin_balance", "released_sum", "released_n") and v != "":
                    try:
                        v = float(v)
                    except Exception:
                        pass
                row.append(v)
            gv.append(row)
        sheet_upsert(ss, TAB_GROUPS, GROUP_SHEET_COLS, gv, ("Store", "group_id"), log)


def push_orders_api(ss, store_names, log):
    cols = list(ORD_SHEET_COLS)
    vals = []
    for sn in store_names:
        rd, _ = read_ledger(os.path.join(LEDGER_DIR, sn, "orders.csv"))
        for r in rd:
            r["Store"] = sn
            row = []
            for c in cols:
                v = r.get(c, "")
                if c == "product-name":
                    v = str(v)[:60]
                if c in ORD_NUM and v != "":
                    try:
                        v = float(v)
                    except Exception:
                        pass
                row.append(v)
            vals.append(row)
    if not vals:
        log("⚠️ مفيش أوردرات تتكتب في الشيت")
        return
    sheet_upsert(ss, TAB_ORD, cols, vals, ("Store", "_key"), log, cmp_cols=[c for c in cols if c != "_first_seen"])


def _norm_cell(v):
    t = str(v).strip()
    try:
        return f"{float(t.replace(',', '')):.2f}"
    except ValueError:
        return t.lower()


def _pg_push_orders(rows, log):
    """نفس منطق push_orders (إضافة من غير تكرار بنفس المفتاح) بس على جدول PostgreSQL."""
    data = None
    c = pg_conn()
    try:
        cur = c.cursor()
        cur.execute("SELECT to_regclass(%s)", (f"{PG_SCHEMA}._meta",))
        if cur.fetchone()[0] is not None:
            cur.execute(f"SELECT table_name, headers, col_names FROM {PG_SCHEMA}._meta WHERE tab_title=%s", (ORDERS_SHEET,))
            m = cur.fetchone()
            if m:
                table, hdr, cols = m
                cur.execute("SELECT " + ", ".join(f'"{x}"' for x in cols) + f' FROM {PG_SCHEMA}."{table}" ORDER BY _row')
                data = [list(hdr)] + [["" if v is None else v for v in r] for r in cur.fetchall()]
    finally:
        c.close()
    if not data or len(data) < 2:
        pg_write_tab(ORDERS_SHEET, ORDERS_HEADERS, rows, log)
        return
    header = [h.strip() for h in data[0]]
    for h in ORDERS_HEADERS:
        if h not in header:
            header.append(h)
    pos = {h: header.index(h) for h in ORDERS_HEADERS}
    kc = [pos[ORDERS_HEADERS[k]] for k in ORDERS_KEY_COLS]
    have = collections.Counter(tuple(_norm_cell(r[k] if k < len(r) else "") for k in kc) for r in data[1:])
    fresh = []
    for r in rows:
        k = tuple(_norm_cell(r[i]) for i in ORDERS_KEY_COLS)
        if have[k] > 0:
            have[k] -= 1
            continue
        line = [""] * len(header)
        for i, h in enumerate(ORDERS_HEADERS):
            line[pos[h]] = r[i]
        fresh.append(line)
    log(f"➕ 🐘 {ORDERS_SHEET}: {len(fresh)} صف جديد من {len(rows)} (الباقي موجود قبل كده)")
    if fresh:
        pg_write_tab(ORDERS_SHEET, header, [list(r) + [""] * (len(header) - len(r)) for r in data[1:]] + fresh, log)


def push_orders(ss, rows, log):
    """إضافة تحت آخر صف في DailyOrdersAmazon من غير تكرار (بنفس مفتاح البوت القديم). الأعمدة بتتحط بالاسم مش بالمكان."""
    if not rows:
        log("⚠️ مفيش شحنات للرفع — DailyOrdersAmazon ما اتلمسش")
        return
    if pg_side(log, lambda: _pg_push_orders(rows, log)):
        return
    ws = _ws(ss, ORDERS_SHEET, len(ORDERS_HEADERS))
    data = ws.get_all_values()
    header = [h.strip() for h in data[0]] if data and any(data[0]) else []
    if not header:
        header = list(ORDERS_HEADERS)
        ws.update(values=[header], range_name="A1")
    else:
        miss = [h for h in ORDERS_HEADERS if h not in header]
        if miss:
            header = header + miss
            ws.update(values=[header], range_name="A1")
    pos = {h: header.index(h) for h in ORDERS_HEADERS}
    kc = [pos[ORDERS_HEADERS[c]] for c in ORDERS_KEY_COLS]
    have = collections.Counter(tuple(_norm_cell(r[c] if c < len(r) else "") for c in kc) for r in data[1:])
    fresh = []
    for r in rows:
        k = tuple(_norm_cell(r[c]) for c in ORDERS_KEY_COLS)
        if have[k] > 0:
            have[k] -= 1
            continue
        line = [""] * len(header)
        for i, h in enumerate(ORDERS_HEADERS):
            line[pos[h]] = r[i]
        fresh.append(line)
    log(f"➕ {ORDERS_SHEET}: {len(fresh)} صف جديد من {len(rows)} (الباقي موجود قبل كده)")
    if fresh:
        _append(ws, fresh, log, ORDERS_SHEET, chunk=1000)


def push_stock(ss, rows, log):
    if not rows:
        log("⚠️ مفيش صفوف مخزون — StockAmazon ما اتلمسش")
        return
    if pg_side(log, lambda: pg_write_tab(STOCK_SHEET, STOCK_HEADERS, rows, log)):
        return
    ws = _ws(ss, STOCK_SHEET, len(STOCK_HEADERS))
    head = [h.strip() for h in ws.row_values(1)]
    if head != STOCK_HEADERS:
        ws.clear()
        ws.update(values=[STOCK_HEADERS], range_name="A1")
    else:
        ws.batch_clear([f"A2:F{max(ws.row_count, 2)}"])
    if ws.row_count < len(rows) + 5:
        ws.resize(rows=len(rows) + 50, cols=len(STOCK_HEADERS))
    _append(ws, rows, log, STOCK_SHEET, chunk=2000)



# ═══════════════════════ صور المنتجات (Catalog Items API) ═══════════════════════
def _pick_main(images):
    """أحسن صورة MAIN: أصغر واحدة عرضها ≥ 500 (خفيفة وواضحة) وإلا أكبر متاحة."""
    for grp in images or []:
        lst = [i for i in (grp.get("images") or []) if i.get("link")]
        mains = [i for i in lst if str(i.get("variant", "")).upper() == "MAIN"] or lst
        if not mains:
            continue
        big = sorted([i for i in mains if int(i.get("width") or 0) >= 500], key=lambda i: int(i.get("width") or 0))
        return (big[0] if big else sorted(mains, key=lambda i: -int(i.get("width") or 0))[0])["link"]
    return None


def fetch_images(st, asins, log, stop):
    """{ASIN: (image, title)} — 20 ASIN في الطلب."""
    out = {}
    asins = list(asins)
    for i in range(0, len(asins), 20):
        if stop.is_set():
            break
        chunk = asins[i:i + 20]
        try:
            j = api_get(st, "/catalog/2022-04-01/items", {"marketplaceIds": st.marketplace, "identifiers": ",".join(chunk),
                                                         "identifiersType": "ASIN", "includedData": "images,summaries", "pageSize": 20})
        except ApiError as e:
            log(f"[{st.name}] ⚠️ صور (دفعة {i // 20 + 1}): {str(e)[:160]}")
            if getattr(e, "code", 0) in (401, 403):
                raise
            continue
        for it in j.get("items", []):
            a = str(it.get("asin", "")).upper()
            summ = (it.get("summaries") or [{}])[0]
            out[a] = (_pick_main(it.get("images")) or "", str(summ.get("itemName", ""))[:80])
        time.sleep(0.6)
    return out


def run_images(stores, log, stop, force=False):
    """يجمّع كل ASIN من الدفتر (مالية + أوردرات) ويجيب صورته وعنوانه. المحفوظ أقل من 14 يوم ما بيتجابش تاني. يرجع صفوف التاب."""
    try:
        cache = json.load(open(IMG_CACHE, encoding="utf-8"))
    except Exception:
        cache = {}
    now = time.time()
    for st in stores:
        if stop.is_set():
            break
        asins = set()
        for fn in ("finance.csv", "orders.csv"):
            rd, _ = read_ledger(os.path.join(LEDGER_DIR, st.name, fn))
            asins |= {str(r.get("asin", "")).strip().upper() for r in rd if r.get("asin")}
        todo = sorted(a for a in asins if a and (force or a not in cache or not cache[a].get("img")
                                                  or now - cache[a].get("ts", 0) > IMG_REFRESH_DAYS * 86400))
        log(f"[{st.name}] 🖼️ صور: {len(asins)} ASIN — هجيب {len(todo)}")
        if todo:
            try:
                got = fetch_images(st, todo, log, stop)
            except ApiError as e:
                log(f"❌ [{st.name}] الصور: {str(e)[:200]} — غالباً الدور «Product Listing» ناقص في تطبيق SP-API")
                continue
            for a, (img, title) in got.items():
                cache[a] = {"img": img, "title": title, "store": st.name, "ts": now}
            log(f"[{st.name}] 🖼️ اتجابت {sum(1 for v in got.values() if v[0])} صورة من {len(todo)}")
    os.makedirs(LEDGER_DIR, exist_ok=True)
    json.dump(cache, open(IMG_CACHE, "w", encoding="utf-8"), ensure_ascii=False)
    return [[a, v.get("img", ""), v.get("title", ""), v.get("store", "")] for a, v in sorted(cache.items()) if v.get("img")]


def push_images(ss, rows, log):
    if not rows:
        log("⚠️ مفيش صور تتكتب — تاب الصور ما اتلمسش")
        return
    sheet_replace_safe(ss, TAB_IMG, IMG_COLS, rows, log)


# ═══════════════════════ تشغيل كامل ═══════════════════════
def run_all(stores, mode, start, steps, want_sheet, log, stop, orders_days=14):
    """steps: dict(fin, orders, stock). المالية بتتكتب في الشيت لو want_sheet؛ الشحنات والمخزون ليهم هدف واحد (الشيت) فبيتخطّوا من غيره."""
    done = []
    for st in stores:
        if stop.is_set():
            break
        try:
            if steps.get("fin", True):
                run_store(st, mode, start, log, stop)
            done.append(st.name)
        except Exception as e:
            log(f"❌ [{st.name}] المالية: {type(e).__name__}: {str(e)[:200]}")
    if steps.get("ordapi", True):
        for st in stores:
            if stop.is_set():
                break
            try:
                run_orders_ledger(st, mode, start, log, stop)
            except Exception as e:
                log(f"❌ [{st.name}] الأوردرات: {type(e).__name__}: {str(e)[:200]}")
    ss = None
    if want_sheet and not stop.is_set():
        try:
            ss = open_sheet()
        except Exception as e:
            log(f"❌ الاتصال بـ Google Sheets: {type(e).__name__}: {str(e)[:200]}")
    if steps.get("fin", True) and ss and done:
        try:
            push_finance(ss, done, log)
        except Exception as e:
            log(f"❌ كتابة المالية في الشيت: {type(e).__name__}: {str(e)[:200]}")
    if steps.get("ordapi", True) and ss and not stop.is_set():
        try:
            push_orders_api(ss, [s.name for s in stores], log)
        except Exception as e:
            log(f"❌ كتابة الأوردرات في الشيت: {type(e).__name__}: {str(e)[:200]}")
    if steps.get("orders") and not want_sheet:
        log("ℹ️ الشحنات هدفها الشيت بس — فعّل خانة Google Sheet")
    if steps.get("orders") and ss and not stop.is_set():
        rows, ok = [], True
        for st in [s for s in stores if s.to_orders]:
            try:
                rows += run_orders(st, orders_days, log, stop)
            except Exception as e:
                ok = False
                log(f"❌ [{st.name}] الشحنات: {type(e).__name__}: {str(e)[:200]}")
        if rows:
            try:
                push_orders(ss, rows, log)
            except Exception as e:
                log(f"❌ كتابة DailyOrdersAmazon: {type(e).__name__}: {str(e)[:200]}")
        elif ok:
            log("⚠️ مفيش شحنات في الفترة دي")
    if steps.get("stock") and not want_sheet:
        log("ℹ️ المخزون هدفه الشيت بس — فعّل خانة Google Sheet")
    if steps.get("stock") and ss and not stop.is_set():
        rows, ok = [], True
        picked = [s for s in stores if s.to_stock]
        for st in picked:
            try:
                rows += run_stock(st, log, stop)
            except Exception as e:
                ok = False
                log(f"❌ [{st.name}] المخزون: {type(e).__name__}: {str(e)[:200]}")
        if ok and picked:
            try:
                push_stock(ss, rows, log)
            except Exception as e:
                log(f"❌ كتابة StockAmazon: {type(e).__name__}: {str(e)[:200]}")
        elif not ok:
            log("⚠️ متجر فشل — StockAmazon ما اتمسحش (عشان ماتضيعش بيانات)")
    if steps.get("img", True) and ss and not stop.is_set():
        try:
            push_images(ss, run_images(stores, log, stop), log)
        except Exception as e:
            log(f"❌ الصور: {type(e).__name__}: {str(e)[:200]}")


def load_stores():
    out = []
    try:
        for d in json.load(open(STORES_FILE, encoding="utf-8")):
            try:
                out.append(AStore(d))
            except Exception:
                pass
    except Exception:
        pass
    return out


def load_settings():
    try:
        return json.load(open(SETTINGS_FILE, encoding="utf-8"))
    except Exception:
        return {}


def auto_main():
    stop = threading.Event()
    log = lambda m: print(time.strftime("%H:%M:%S"), m, flush=True)
    s = load_settings()
    steps = {"fin": s.get("fin", True), "ordapi": s.get("ordapi", True), "orders": s.get("orders", True), "stock": s.get("stock", True), "img": s.get("img", True)}
    run_all(load_stores(), "daily", DEFAULT_START, steps, "--sheet" in sys.argv, log, stop, int(s.get("orders_days", 14) or 14))


import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox


class App(tk.Frame):
    def __init__(self, root):
        super().__init__(root, bg="#d9d9d9")
        self.pack(fill="both", expand=True, padx=8, pady=8)
        self.stop = threading.Event()
        self.stores = []
        self.running = False
        self._last_fixed = ""
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)
        ttk.Label(top, text="المتاجر:").grid(row=0, column=0, sticky="nw")
        self.lb = tk.Listbox(top, height=4, width=80)
        self.lb.grid(row=0, column=1, padx=6)
        bb = ttk.Frame(top)
        bb.grid(row=0, column=2, sticky="n")
        ttk.Button(bb, text="➕ إضافة", command=lambda: self.store_dialog(None)).pack(fill="x")
        ttk.Button(bb, text="✏️ تعديل", command=self.edit_store).pack(fill="x", pady=2)
        ttk.Button(bb, text="➖ حذف", command=self.rm_store).pack(fill="x")
        ttk.Label(top, text="تاريخ بداية المالية:").grid(row=1, column=0, sticky="w", pady=4)
        self.dfrom = tk.StringVar(value=DEFAULT_START.isoformat())
        ttk.Entry(top, textvariable=self.dfrom, width=14).grid(row=1, column=1, sticky="w", padx=6)
        ttk.Label(top, text="شحنات آخر (يوم):").grid(row=2, column=0, sticky="w")
        self.odays = tk.StringVar(value="14")
        ttk.Entry(top, textvariable=self.odays, width=6).grid(row=2, column=1, sticky="w", padx=6)
        sf = ttk.Frame(top)
        sf.grid(row=3, column=1, sticky="w", pady=2)
        self.s_fin = tk.BooleanVar(value=True)
        self.s_ord = tk.BooleanVar(value=True)
        self.s_stk = tk.BooleanVar(value=True)
        self.s_oapi = tk.BooleanVar(value=True)
        self.s_img = tk.BooleanVar(value=True)
        ttk.Checkbutton(sf, text=f"💰 المالية ({TAB_FIN})", variable=self.s_fin).pack(side="left", padx=4)
        ttk.Checkbutton(sf, text=f"🧾 الأوردرات بحالتها ({TAB_ORD})", variable=self.s_oapi).pack(side="left", padx=4)
        ttk.Checkbutton(sf, text=f"📦 الشحنات → {ORDERS_SHEET}", variable=self.s_ord).pack(side="left", padx=4)
        ttk.Checkbutton(sf, text=f"🏬 المخزون → {STOCK_SHEET}", variable=self.s_stk).pack(side="left", padx=4)
        ttk.Checkbutton(sf, text=f"🖼️ الصور → {TAB_IMG}", variable=self.s_img).pack(side="left", padx=4)
        self.sheet = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="اكتب في Google Sheet (لازم للشحنات والمخزون)", variable=self.sheet).grid(row=4, column=1, sticky="w")
        self.auto = tk.BooleanVar(value=False)
        self.times = tk.StringVar(value="")
        ttk.Checkbutton(top, text="⏰ تحديث يومي تلقائي في المواعيد دي (HH:MM مفصولة بفاصلة):", variable=self.auto).grid(row=5, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Entry(top, textvariable=self.times, width=24).grid(row=5, column=2, sticky="w")
        self.next_lbl = ttk.Label(top, text="")
        self.next_lbl.grid(row=6, column=1, sticky="w")
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=10)
        self.btns = []
        for txt, mode in (("🔄 تحديث يومي", "daily"), ("📥 سحب كامل من تاريخ البداية", "full")):
            b = ttk.Button(bar, text=txt, command=lambda m=mode: self.go(m))
            b.pack(side="left", padx=4, pady=6)
            self.btns.append(b)
        ttk.Button(bar, text="⏹ إيقاف", command=self.stop.set).pack(side="left", padx=4)
        self.box = scrolledtext.ScrolledText(self, font=("Consolas", 10), height=16)
        self.box.pack(fill="both", expand=True, padx=10, pady=8)
        self.reload_stores()
        self.load_settings()
        for _v in (self.sheet, self.times, self.dfrom, self.odays, self.auto, self.s_fin, self.s_ord, self.s_stk, self.s_oapi, self.s_img):
            _v.trace_add("write", lambda *a: self.save_settings())
        self.log("جاهز. أول مرة: «سحب كامل» (المالية + الشحنات + المخزون). بعدها «تحديث يومي».")
        self.after(5000, self.tick)

    def reload_stores(self):
        self.stores = load_stores()
        self.lb.delete(0, "end")
        for st in self.stores:
            self.lb.insert("end", f"{st.name} — {st.marketplace} — شحنات:{'✔' if st.to_orders else '✘'} مخزون:{'✔' if st.to_stock else '✘'}")

    def log(self, m):
        try:
            if LOG_HOOK:
                LOG_HOOK(str(m) + "\n")
        except Exception:
            pass
        try:
            self.after(0, lambda: (self.box.insert("end", str(m) + "\n"), self.box.see("end")))
        except Exception:
            pass

    def load_settings(self):
        s = load_settings()
        try:
            self.sheet.set(bool(s.get("sheet", False)))
            self.times.set(str(s.get("times", "")))
            self.auto.set(bool(s.get("auto", False)))
            self.s_fin.set(bool(s.get("fin", True)))
            self.s_ord.set(bool(s.get("orders", True)))
            self.s_stk.set(bool(s.get("stock", True)))
            self.s_oapi.set(bool(s.get("ordapi", True)))
            self.s_img.set(bool(s.get("img", True)))
            self.odays.set(str(s.get("orders_days", 14)))
            if s.get("dfrom"):
                self.dfrom.set(str(s["dfrom"]))
        except Exception:
            pass

    def save_settings(self):
        try:
            json.dump({"sheet": bool(self.sheet.get()), "times": self.times.get(), "auto": bool(self.auto.get()), "dfrom": self.dfrom.get(),
                       "fin": bool(self.s_fin.get()), "orders": bool(self.s_ord.get()), "stock": bool(self.s_stk.get()), "ordapi": bool(self.s_oapi.get()), "img": bool(self.s_img.get()),
                       "orders_days": self.odays.get()}, open(SETTINGS_FILE, "w", encoding="utf-8"), ensure_ascii=False)
        except Exception:
            pass

    def fixed_times(self):
        out = []
        for t in str(self.times.get()).replace("،", ",").replace(" ", ",").split(","):
            t = t.strip()
            try:
                if t:
                    out.append(dt.datetime.strptime(t, "%H:%M").strftime("%H:%M"))
            except ValueError:
                pass
        return out

    def tick(self):
        try:
            ft = self.fixed_times()
            self.next_lbl.configure(text=("المواعيد: " + " ، ".join(ft)) if (self.auto.get() and ft) else "")
            now = time.strftime("%H:%M")
            key = time.strftime("%Y-%m-%d ") + now
            if self.auto.get() and now in ft and key != self._last_fixed and not self.running:
                self._last_fixed = key
                self.log(f"🕐 موعد محدد {now} — تحديث يومي تلقائي...")
                self.go("daily", auto=True)
        except Exception:
            pass
        self.after(20000, self.tick)

    def save_stores(self):
        try:
            json.dump([s.as_dict() for s in self.stores], open(STORES_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        except Exception:
            pass

    def store_dialog(self, cur):
        w = tk.Toplevel(self)
        w.title("متجر أمازون")
        f = {}
        for i, (k, t, show) in enumerate([("name", "اسم المتجر", ""), ("client_id", "LWA Client ID", ""), ("client_secret", "Client Secret", "•"),
                                          ("refresh_token", "Refresh Token", "•"), ("marketplace", "Marketplace ID", ""), ("endpoint", "Endpoint", "")]):
            ttk.Label(w, text=t).grid(row=i, column=0, padx=6, pady=3, sticky="e")
            e = ttk.Entry(w, width=56, show=show)
            e.grid(row=i, column=1, padx=6, pady=3)
            e.insert(0, (cur.as_dict()[k] if cur else {"marketplace": DEFAULT_MARKETPLACE, "endpoint": DEFAULT_ENDPOINT}.get(k, "")))
            f[k] = e
        vo = tk.BooleanVar(value=cur.to_orders if cur else True)
        vs = tk.BooleanVar(value=cur.to_stock if cur else True)
        ttk.Checkbutton(w, text=f"يغذّي {ORDERS_SHEET} (الشحنات)", variable=vo).grid(row=6, column=1, sticky="w", padx=6)
        ttk.Checkbutton(w, text=f"يغذّي {STOCK_SHEET} (المخزون)", variable=vs).grid(row=7, column=1, sticky="w", padx=6)

        def ok():
            d = {k: e.get().strip() for k, e in f.items()}
            if not (d["client_id"] and d["client_secret"] and d["refresh_token"]):
                return
            d["to_orders"], d["to_stock"] = bool(vo.get()), bool(vs.get())
            s = AStore(d)
            if cur is not None:
                self.stores[self.stores.index(cur)] = s
            else:
                self.stores.append(s)
            self.save_stores()
            self.reload_stores()
            w.destroy()
        ttk.Button(w, text="حفظ", command=ok).grid(row=8, column=1, pady=8, sticky="e")

    def edit_store(self):
        sel = self.lb.curselection()
        if sel:
            self.store_dialog(self.stores[sel[0]])

    def rm_store(self):
        for i in reversed(self.lb.curselection()):
            del self.stores[i]
        self.save_stores()
        self.reload_stores()

    def go(self, mode, auto=False, sheet=None):
        if self.running:
            return
        if not self.stores:
            if not auto:
                messagebox.showwarning("", "ضيف متجر واحد على الأقل")
            return
        try:
            start = dt.date.fromisoformat(self.dfrom.get().strip())
            odays = int(str(self.odays.get()).strip() or 14)
        except Exception as e:
            if not auto:
                messagebox.showerror("خطأ", str(e))
            return
        self.stop.clear()
        self.running = True
        for b in self.btns:
            b.state(["disabled"])
        want_sheet = self.sheet.get() if sheet is None else bool(sheet)
        steps = {"fin": self.s_fin.get(), "ordapi": self.s_oapi.get(), "orders": self.s_ord.get(), "stock": self.s_stk.get(), "img": self.s_img.get()}
        stores = list(self.stores)

        def work():
            try:
                run_all(stores, mode, start, steps, want_sheet, self.log, self.stop, odays)
                self.log("✅ خلص.")
            except Exception as e:
                self.log(f"❌ {type(e).__name__}: {e}")
            finally:
                self.running = False
                try:
                    self.after(0, lambda: [b.state(["!disabled"]) for b in self.btns])
                except Exception:
                    pass
        threading.Thread(target=work, daemon=True).start()

    def shutdown(self):
        self.stop.set()


def build_ui(root):
    """بيتنادى من البرنامج الموحد: بيبني الواجهة جوه التاب ويرجّع دالة إيقاف."""
    return App(root).shutdown


def main():
    root = tk.Tk()
    root.title("Amazon Accounts — حسابات وشحنات ومخزون أمازون")
    root.geometry("1000x760")
    a = App(root)
    root.protocol("WM_DELETE_WINDOW", lambda: (a.shutdown(), root.destroy()))
    root.mainloop()


if __name__ == "__main__":
    if "--auto" in sys.argv:
        auto_main()
    else:
        main()
'''

_EMBEDDED['NoonTracker'] = r'''# -*- coding: utf-8 -*-
"""
Noon_Tracker.py — متابعة منافسين نون (صفحات noon.com العامة، مفيش API) كتاب في البلاتفورم الموحد
- بيسحب صفحات منتجاتنا + المنافسين من جوه كروم (نفس الكوكيز) بدفعات متوازية صغيرة، وبيهدّي نفسه لو نون رفض.
- بيتخطّى الاسكيوهات اللي اتحدّثت حديثًا (إلا اللي ليها منافسين).
- النتائج: جوجل شيت (NoonTrack_Products / Offers / Comps) و/أو إكسيل.
- ملف المنافسين noon_competitors.xlsx هو نفسه بتاع أداة noon_tracker.py (PyQt) — تقدر تفتحها من الزرار.
المتطلبات:  py -m pip install selenium openpyxl gspread oauth2client
"""
import os, re, sys, json, time, glob, random, shutil, tempfile, subprocess, threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

MARKETS = {
    "السعودية": {"path": "saudi-en", "locale": "en-sa"},
    "الإمارات": {"path": "uae-en", "locale": "en-ae"},
    "مصر": {"path": "egypt-en", "locale": "en-eg"},
}
PRODUCT_COLS = [
    "SKU", "اسم المنتج", "البراند", "القسم", "التقييم", "عدد التقييمات",
    "ترتيب/شارات", "صاحب البوكس", "كود صاحب البوكس", "سعر البوكس",
    "السعر قبل الخصم", "نسبة الخصم %", "البوكس FBN", "مخزون البوكس", "التوصيل",
    "عدد البائعين", "متجرنا موجود", "متاجرنا على المنتج", "سعرنا", "ترتيبنا",
    "إحنا البوكس", "أقل سعر", "أقل بائع", "كود أقل بائع", "الفرق عن أقل سعر",
    "الرابط", "الصورة", "الحالة",
]
COMP_COLS = [
    "SKU تبعنا", "اسم منتجنا", "سعرنا", "SKU المنافس", "اسم منتج المنافس",
    "البراند", "بائع المنافس", "كود بائع المنافس", "سعر المنافس",
    "الفرق (المنافس - إحنا)", "أرخص من مين", "أقل بائع عند المنافس",
    "كود أقل بائع عند المنافس", "أقل سعر عند المنافس", "متاجرنا على منتج المنافس",
    "التقييم", "عدد التقييمات", "عدد البائعين", "FBN", "المخزون", "ترتيب/شارات",
    "الرابط", "الحالة",
]
OFFER_COLS = [
    "SKU", "اسم المنتج", "ترتيب نون", "البائع", "كود البائع", "السعر النهائي",
    "السعر قبل الخصم", "FBN", "المخزون", "متاح للشراء", "التوصيل",
    "تقييم البائع", "عدد تقييمات البائع", "صاحب البوكس", "متجرنا",
    "الفرق عن سعرنا",
]
CATALOG_HEADERS = {"x-mp": "noon", "x-platform": "web", "x-content": "desktop"}
GSHEET_CREDS_DIR = r"\\WIN-181U7G4HFM5\Share For Markting\NOONIMG\Copy"
GSHEET_NAME = "Complaints"
GSHEET_SCOPE = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
TAB_PRODUCTS = "NoonTrack_Products"
TAB_OFFERS = "NoonTrack_Offers"
TAB_COMPS = "NoonTrack_Comps"
TAB_SKUS = "NoonTrack_SKUs"
STAMP_COL = "آخر تحديث"
FAILED_STATUSES = ("فشل", "مرفوض من نون")
SKU_IN_URL = re.compile(r"/([A-Z0-9]{8,}(?:-\d+)?)/p", re.I)

def strip_variant(sku):
    """Z123...Z-1 -> Z123...Z (نفس صيغة صفحة المنتج العامة على noon.com)"""
    return re.sub(r"-\d+$", "", str(sku or "").strip().upper())

def gs_open():
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    files = glob.glob(os.path.join(GSHEET_CREDS_DIR, "*.json"))
    if not files:
        raise RuntimeError(f"مفيش ملف اعتماد Google في: {GSHEET_CREDS_DIR}")
    creds = ServiceAccountCredentials.from_json_keyfile_name(files[0], GSHEET_SCOPE)
    return gspread, gspread.authorize(creds).open(GSHEET_NAME)

def gs_target_skus(ss):
    """اسكيوهات نون الظاهرة في تاب المبيعات (التطبيق بيكتبها في تاب NoonTrack_SKUs) بدون -1 وبدون تكرار."""
    try:
        vals = ss.worksheet(TAB_SKUS).get_all_values()
    except Exception:
        raise RuntimeError(f"تاب {TAB_SKUS} مش موجود - افتح تاب المبيعات في التطبيق مرة الأول (بيعمله)")
    out, seen = [], set()
    for row in vals[1:]:
        s = strip_variant(row[0]) if row else ""
        if s and s not in seen:
            seen.add(s)
            out.append(s)
    return out

def gs_merge_write(gspread, ss, tab, headers, new_rows, key_cols, replace_keys, stamp):
    """
    يكتب new_rows في التاب (بيتعمل أول مرة). الصفوف القديمة اللي مفتاحها في replace_keys بتتشال
    وبتتحط الجديدة، وباقي الصفوف القديمة بتفضل زي ما هي (عشان اسكيو فشل سحبه ما يتمسحش).
    """
    full = list(headers) + [STAMP_COL]
    try:
        ws = ss.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title=tab, rows="2000", cols=str(len(full)))
    old = ws.get_all_values()
    keep = []
    if old and [h.strip() for h in old[0]] == full:
        ki = [full.index(c) for c in key_cols]
        for r in old[1:]:
            r = list(r) + [""] * (len(full) - len(r))
            if tuple(r[i] for i in ki) not in replace_keys:
                keep.append(r)

    def cell(v):
        if v is None:
            return ""
        if isinstance(v, str):
            return v[:4000]
        return v
    fresh = [[cell(r.get(c)) for c in headers] + [stamp] for r in new_rows]
    values = [full] + keep + fresh
    ws.clear()
    ws.resize(rows=max(len(values) + 20, 100), cols=len(full))
    for i in range(0, len(values), 2000):
        ws.update(range_name=f"A{i + 1}", values=values[i:i + 2000], value_input_option="RAW")
    return len(keep), len(fresh)

def gs_push_results(gspread, ss, products, offers, comps, log):
    """يرفع نتيجة التشغيل لتابات NoonTrack_*. اسكيو فشل سحبه بيفضل بياناته القديمة."""
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    ok_products = [r for r in products if r.get("الحالة") not in FAILED_STATUSES]
    ok_skus = {r["SKU"] for r in ok_products}
    ok_comps = [r for r in comps if r.get("الحالة") not in FAILED_STATUSES]
    jobs = [
        (TAB_PRODUCTS, PRODUCT_COLS, ok_products, ("SKU",), {(r["SKU"],) for r in ok_products}),
        (TAB_OFFERS, OFFER_COLS, [o for o in offers if o.get("SKU") in ok_skus], ("SKU",),
         {(k,) for k in ok_skus}),
        (TAB_COMPS, COMP_COLS, ok_comps, ("SKU تبعنا", "SKU المنافس"),
         {(r["SKU تبعنا"], r["SKU المنافس"]) for r in ok_comps}),
    ]
    for tab, headers, rows, keys, rkeys in jobs:
        kept, fresh = gs_merge_write(gspread, ss, tab, headers, rows, keys, rkeys, stamp)
        log(f"جوجل شيت {tab}: {fresh} صف جديد + {kept} قديم محفوظ")

def clean_sku(value):
    s = str(value or "").strip()
    if not s or s.lower() == "nan":
        return ""
    m = SKU_IN_URL.search(s)
    if m:
        return m.group(1).upper()
    return s.split()[0].upper()

def detect_columns(headers):
    """يرجع (عمود الاسكيو, عمود المنافس أو None) - 0-based"""
    names = [str(h or "").strip().lower() for h in headers]
    comp = None
    for i, n in enumerate(names):
        if "منافس" in n or "comp" in n:
            comp = i
            break
    sku = None
    for i, n in enumerate(names):
        if i == comp:
            continue
        if "sku" in n or "اسكيو" in n or "رابط" in n or "link" in n:
            sku = i
            break
    if sku is None:
        sku = 1 if comp == 0 else 0
    return sku, comp

def ours_from_pairs(pairs):
    out = []
    for o, _ in pairs:
        if o not in out:
            out.append(o)
    return out

def strip_html(text):
    return re.sub(r"<[^>]+>", "", str(text or "")).strip()

def to_num(v):
    try:
        if v is None or v == "":
            return None
        return float(v)
    except Exception:
        return None

def final_price(offer):
    sale = to_num(offer.get("sale_price"))
    return sale if sale else to_num(offer.get("price"))

def is_ours(offer, our_set):
    if not our_set:
        return False
    name = str(offer.get("store_name") or "").strip().lower()
    code = str(offer.get("partner_code") or "").strip().lower()
    return name in our_set or code in our_set

def pick_variant(product, sku_in):
    variants = product.get("variants") or []
    for v in variants:
        if str(v.get("sku") or "").upper() == sku_in:
            return v
    for v in variants:
        if v.get("offers"):
            return v
    return variants[0] if variants else {}

def breadcrumbs_text(product):
    parts = []
    for b in product.get("breadcrumbs") or []:
        if isinstance(b, dict):
            n = b.get("name") or b.get("title") or ""
            if n:
                parts.append(str(n))
        elif b:
            parts.append(str(b))
    return " > ".join(parts)

def empty_product_row(sku, status):
    row = {c: "" for c in PRODUCT_COLS}
    row["SKU"] = sku
    row["الحالة"] = status
    return row

def parse_product(sku_in, data, our_set, market_path):
    product = data.get("product") or {}
    if not product:
        return empty_product_row(sku_in, "المنتج غير موجود"), []

    variant = pick_variant(product, sku_in)
    offers = variant.get("offers") or []
    title = product.get("product_title") or ""
    sku_cfg = product.get("sku") or sku_in.split("-")[0]

    bb_code = product.get("offer_code")
    buybox = None
    for o in offers:
        if bb_code and o.get("offer_code") == bb_code:
            buybox = o
            break
    if buybox is None and offers:
        buybox = offers[0]

    rating = product.get("product_rating") or {}
    row = empty_product_row(sku_in, "تم")
    row.update({
        "اسم المنتج": title,
        "البراند": product.get("brand") or "",
        "القسم": breadcrumbs_text(product),
        "التقييم": rating.get("value", ""),
        "عدد التقييمات": rating.get("count", ""),
        "عدد البائعين": len(offers),
    })

    slug = (buybox or {}).get("url") or "x"
    row["الرابط"] = f"https://www.noon.com/{market_path}/{slug}/{sku_cfg}/p/"

    imgs = product.get("image_urls") or []
    if imgs:
        row["الصورة"] = imgs[0]
    elif buybox:
        row["الصورة"] = (buybox.get("wishlist_data") or {}).get("img_url", "")

    if buybox:
        fp = final_price(buybox)
        before = to_num(buybox.get("price"))
        row.update({
            "صاحب البوكس": buybox.get("store_name") or "",
            "كود صاحب البوكس": buybox.get("partner_code") or "",
            "سعر البوكس": fp,
            "السعر قبل الخصم": before,
            "نسبة الخصم %": round((before - fp) / before * 100, 1) if before and fp and before > fp else 0,
            "البوكس FBN": "نعم" if "fbn" in (buybox.get("flags") or []) else "لا",
            "مخزون البوكس": buybox.get("stock", ""),
            "التوصيل": strip_html(buybox.get("estimated_delivery")),
            "ترتيب/شارات": " | ".join(
                str(n.get("text")) for n in (buybox.get("nudges") or []) if n.get("text")
            ),
        })

    if not offers:
        row["الحالة"] = "لا يوجد بائعين"
        return row, []

    priced = [o for o in offers if final_price(o)]
    lowest = min(priced, key=final_price) if priced else None
    if lowest:
        row["أقل سعر"] = final_price(lowest)
        row["أقل بائع"] = lowest.get("store_name") or ""
        row["كود أقل بائع"] = lowest.get("partner_code") or ""

    # متاجرنا (ممكن أكتر من متجر من متاجرنا على نفس المنتج) - سعرنا = أقل سعر فيهم
    ours_on = [(i, o) for i, o in enumerate(offers, start=1) if is_ours(o, our_set)]
    our_offer, our_rank = None, None
    for i, o in ours_on:
        if final_price(o) and (our_offer is None or final_price(o) < final_price(our_offer)):
            our_offer, our_rank = o, i
    our_price = final_price(our_offer) if our_offer else None
    row["متجرنا موجود"] = "نعم" if ours_on else "لا"
    row["متاجرنا على المنتج"] = " | ".join(
        f"{o.get('store_name') or ''} ({o.get('partner_code') or ''}) {final_price(o) or ''}" for _, o in ours_on
    )
    row["سعرنا"] = our_price if our_price else ""
    row["ترتيبنا"] = our_rank or ""
    row["إحنا البوكس"] = "نعم" if (buybox and is_ours(buybox, our_set)) else "لا"
    if our_price and lowest:
        row["الفرق عن أقل سعر"] = round(our_price - final_price(lowest), 2)

    offer_rows = []
    for i, o in enumerate(offers, start=1):
        fp = final_price(o)
        sl = o.get("partner_ratings_sellerlab") or {}
        offer_rows.append({
            "SKU": sku_in,
            "اسم المنتج": title,
            "ترتيب نون": i,
            "البائع": o.get("store_name") or "",
            "كود البائع": o.get("partner_code") or "",
            "السعر النهائي": fp,
            "السعر قبل الخصم": to_num(o.get("price")),
            "FBN": "نعم" if "fbn" in (o.get("flags") or []) else "لا",
            "المخزون": o.get("stock", ""),
            "متاح للشراء": "نعم" if o.get("is_buyable") else "لا",
            "التوصيل": strip_html(o.get("estimated_delivery")),
            "تقييم البائع": sl.get("partner_rating", ""),
            "عدد تقييمات البائع": sl.get("num_of_rating", ""),
            "صاحب البوكس": "نعم" if o is buybox else "لا",
            "متجرنا": "نعم" if is_ours(o, our_set) else "لا",
            "الفرق عن سعرنا": round(fp - our_price, 2) if (fp and our_price) else "",
        })
    return row, offer_rows

def build_comp_row(our_sku, our_row, comp_sku, comp_row):
    our_price = to_num((our_row or {}).get("سعرنا")) or to_num((our_row or {}).get("سعر البوكس"))
    comp_price = to_num(comp_row.get("سعر البوكس"))
    diff, who = "", ""
    if our_price and comp_price:
        diff = round(comp_price - our_price, 2)
        who = "المنافس أرخص" if diff < 0 else ("إحنا أرخص" if diff > 0 else "نفس السعر")
    return {
        "SKU تبعنا": our_sku,
        "اسم منتجنا": (our_row or {}).get("اسم المنتج", ""),
        "سعرنا": our_price or "",
        "SKU المنافس": comp_sku,
        "اسم منتج المنافس": comp_row.get("اسم المنتج", ""),
        "البراند": comp_row.get("البراند", ""),
        "بائع المنافس": comp_row.get("صاحب البوكس", ""),
        "كود بائع المنافس": comp_row.get("كود صاحب البوكس", ""),
        "سعر المنافس": comp_price or "",
        "الفرق (المنافس - إحنا)": diff,
        "أرخص من مين": who,
        "أقل بائع عند المنافس": comp_row.get("أقل بائع", ""),
        "كود أقل بائع عند المنافس": comp_row.get("كود أقل بائع", ""),
        "أقل سعر عند المنافس": comp_row.get("أقل سعر", ""),
        "متاجرنا على منتج المنافس": comp_row.get("متاجرنا على المنتج", ""),
        "التقييم": comp_row.get("التقييم", ""),
        "عدد التقييمات": comp_row.get("عدد التقييمات", ""),
        "عدد البائعين": comp_row.get("عدد البائعين", ""),
        "FBN": comp_row.get("البوكس FBN", ""),
        "المخزون": comp_row.get("مخزون البوكس", ""),
        "ترتيب/شارات": comp_row.get("ترتيب/شارات", ""),
        "الرابط": comp_row.get("الرابط", ""),
        "الحالة": comp_row.get("الحالة", ""),
    }

def format_sheet(ws, df, painter=None):
    from openpyxl.styles import Font, PatternFill, Alignment
    header_fill = PatternFill("solid", fgColor="FEEE00")
    ws.sheet_view.rightToLeft = True
    ws.freeze_panes = "B2"
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, col in enumerate(df.columns, start=1):
        values = [str(col)] + [str(v) for v in df[col].tolist()[:200]]
        width = min(max(len(v) for v in values) + 2, 50)
        ws.column_dimensions[ws.cell(1, i).column_letter].width = width
    if painter:
        heads = [c.value for c in ws[1]]
        for r in range(2, ws.max_row + 1):
            vals = {h: ws.cell(r, i + 1).value for i, h in enumerate(heads)}
            fill = painter(vals)
            if fill:
                for c in range(1, ws.max_column + 1):
                    ws.cell(r, c).fill = fill

def our_price_of(row):
    row = row or {}
    return to_num(row.get("سعرنا")) or to_num(row.get("سعر البوكس"))

def write_wide_comparison(book, pairs, our_rows, comp_rows):
    """
    شيت المقارنة: صف لكل اسكيو تبعنا
    [منتجنا] [الملخص] [منافس 1] [منافس 2] ...   (المنافسين من الأرخص للأغلى)
    """
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    green = PatternFill("solid", fgColor="C6EFCE")
    red = PatternFill("solid", fgColor="FFC7CE")
    ours_fill = PatternFill("solid", fgColor="FEEE00")
    summary_fill = PatternFill("solid", fgColor="F8CBAD")
    block_fills = [PatternFill("solid", fgColor="DDEBF7"), PatternFill("solid", fgColor="E2EFDA")]
    thin = Side(style="thin", color="BFBFBF")
    thick = Side(style="medium", color="595959")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bold = Font(bold=True)
    big = Font(bold=True, size=12)

    # (اسم العمود, العرض, دالة القيمة)
    our_cols = [
        ("SKU", 24, lambda o: o.get("SKU")),
        ("اسم المنتج", 34, lambda o: o.get("اسم المنتج")),
        ("سعرنا", 10, lambda o: our_price_of(o)),
        ("صاحب البوكس", 16, lambda o: o.get("صاحب البوكس")),
        ("إحنا البوكس", 9, lambda o: o.get("إحنا البوكس")),
        ("التقييم", 8, lambda o: to_num(o.get("التقييم"))),
        ("عدد التقييمات", 10, lambda o: to_num(o.get("عدد التقييمات"))),
        ("ترتيب/شارات", 22, lambda o: o.get("ترتيب/شارات")),
        ("نسبة الخصم %", 9, lambda o: to_num(o.get("نسبة الخصم %"))),
        ("FBN", 6, lambda o: o.get("البوكس FBN")),
        ("المخزون", 8, lambda o: to_num(o.get("مخزون البوكس"))),
        ("عدد البائعين", 8, lambda o: to_num(o.get("عدد البائعين"))),
    ]
    sum_cols = [("عدد المنافسين", 9), ("أقل سعر منافس", 10), ("أرخص منافس", 16),
                ("الفرق عن الأرخص", 10), ("منافسين أرخص مننا", 10)]
    comp_cols = [
        ("SKU", 24, lambda c: c.get("SKU")),
        ("اسم المنتج", 30, lambda c: c.get("اسم المنتج")),
        ("البائع", 16, lambda c: c.get("صاحب البوكس")),
        ("كود البائع", 10, lambda c: c.get("كود صاحب البوكس")),
        ("السعر", 10, lambda c: to_num(c.get("سعر البوكس")) or c.get("الحالة")),
        ("الفرق عن سعرنا", 10, None),
        ("التقييم", 8, lambda c: to_num(c.get("التقييم"))),
        ("عدد التقييمات", 10, lambda c: to_num(c.get("عدد التقييمات"))),
        ("ترتيب/شارات", 22, lambda c: c.get("ترتيب/شارات")),
        ("نسبة الخصم %", 9, lambda c: to_num(c.get("نسبة الخصم %"))),
        ("FBN", 6, lambda c: c.get("البوكس FBN")),
        ("المخزون", 8, lambda c: to_num(c.get("مخزون البوكس"))),
        ("عدد البائعين", 8, lambda c: to_num(c.get("عدد البائعين"))),
    ]
    C_PRICE, C_DIFF, C_RATE, C_COUNT = 4, 5, 6, 7  # أماكن الأعمدة جوه بلوك المنافس (0-based)

    groups = {}
    for o, c in pairs:
        groups.setdefault(o, []).append(c)
    max_n = max((len(v) for v in groups.values()), default=0)

    ws = book.create_sheet("مقارنة المنافسين", 0)
    book.active = 0
    ws.sheet_view.rightToLeft = True

    # ---- الهيدر (سطرين: اسم الجزء + أسماء الأعمدة) ----
    sections = [("منتجنا", ours_fill, [(n, w) for n, w, _ in our_cols]),
                ("الملخص", summary_fill, sum_cols)]
    for n in range(max_n):
        sections.append((f"منافس {n + 1}", block_fills[n % 2], [(c, w) for c, w, _ in comp_cols]))

    col = 1
    section_starts = []
    for title, fill, cols in sections:
        section_starts.append(col)
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + len(cols) - 1)
        top = ws.cell(1, col, title)
        top.font, top.fill, top.alignment = big, fill, center
        for j, (name, width) in enumerate(cols):
            cell = ws.cell(2, col + j, name)
            cell.font, cell.fill, cell.alignment = bold, fill, center
            ws.column_dimensions[cell.column_letter].width = width
        col += len(cols)
    last_col = col - 1
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 32

    # ---- البيانات ----
    s_start = len(our_cols) + 1
    b_start = s_start + len(sum_cols)
    r = 3
    for our_sku, comp_list in groups.items():
        o = our_rows.get(our_sku) or empty_product_row(our_sku, "لم يتم السحب")
        price = our_price_of(o)
        o_rate, o_count = to_num(o.get("التقييم")), to_num(o.get("عدد التقييمات"))

        comps = [comp_rows.get(c) or empty_product_row(c, "لم يتم السحب") for c in comp_list]
        comps.sort(key=lambda c: (to_num(c.get("سعر البوكس")) is None, to_num(c.get("سعر البوكس")) or 0))

        # منتجنا
        for j, (_, _, fn) in enumerate(our_cols):
            cell = ws.cell(r, 1 + j, fn(o))
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if o.get("الرابط"):
            ws.cell(r, 1).hyperlink = o["الرابط"]
            ws.cell(r, 1).font = Font(color="0563C1", underline="single")
        ws.cell(r, 5).fill = green if o.get("إحنا البوكس") == "نعم" else red

        # الملخص
        prices = [to_num(c.get("سعر البوكس")) for c in comps if to_num(c.get("سعر البوكس"))]
        low = min(prices) if prices else None
        cheapest = next((c for c in comps if to_num(c.get("سعر البوكس")) == low), {}) if low else {}
        cheaper = sum(1 for p in prices if price and p < price)
        vals = [len(comps), low or "", cheapest.get("صاحب البوكس", ""),
                round(price - low, 2) if (price and low) else "", cheaper if price else ""]
        for j, v in enumerate(vals):
            cell = ws.cell(r, s_start + j, v)
            cell.alignment = center
        if price and low:
            ws.cell(r, s_start + 3).fill = red if price > low else green
            ws.cell(r, s_start + 4).fill = red if cheaper else green

        # المنافسين
        for n, c in enumerate(comps):
            c0 = b_start + n * len(comp_cols)
            c_price = to_num(c.get("سعر البوكس"))
            diff = round(c_price - price, 2) if (c_price and price) else None
            for j, (_, _, fn) in enumerate(comp_cols):
                v = diff if j == C_DIFF else fn(c)
                cell = ws.cell(r, c0 + j, "" if v is None else v)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if c.get("الرابط"):
                ws.cell(r, c0).hyperlink = c["الرابط"]
                ws.cell(r, c0).font = Font(color="0563C1", underline="single")
            # السعر: أحمر لو المنافس أرخص
            if diff is not None and diff != 0:
                f = red if diff < 0 else green
                ws.cell(r, c0 + C_PRICE).fill = f
                ws.cell(r, c0 + C_DIFF).fill = f
            # التقييم وعدد التقييمات: أحمر لو المنافس أعلى مننا
            c_rate, c_count = to_num(c.get("التقييم")), to_num(c.get("عدد التقييمات"))
            if c_rate is not None and o_rate is not None and c_rate != o_rate:
                ws.cell(r, c0 + C_RATE).fill = red if c_rate > o_rate else green
            if c_count is not None and o_count is not None and c_count != o_count:
                ws.cell(r, c0 + C_COUNT).fill = red if c_count > o_count else green
        ws.row_dimensions[r].height = 45
        r += 1

    # ---- حدود: خط تقيل في أول كل جزء ----
    starts = set(section_starts)
    for row in ws.iter_rows(min_row=1, max_row=max(r - 1, 2), max_col=last_col):
        for cell in row:
            cell.border = Border(left=thick if cell.column in starts else thin,
                                 right=thick if cell.column == last_col else thin,
                                 top=thin, bottom=thin)
            if isinstance(cell.value, float) and cell.row > 2:
                cell.number_format = "0.00" if cell.value != int(cell.value) else "0"

    ws.freeze_panes = ws.cell(3, 2)
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


# ============================================================================
# إعدادات + ملفات (جنب البرنامج)
# ============================================================================
APP_DIR = os.environ.get("PLATFORMS_BASE_DIR") or os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(APP_DIR, "noon_tracker_tk_settings.json")
QT_SETTINGS_FILE = os.path.join(APP_DIR, "noon_tracker_settings.json")   # إعدادات أداة الـ PyQt القديمة (نقرا منها بس)
COMP_FILE = os.path.join(APP_DIR, "noon_competitors.xlsx")
RESULTS_DIR = os.path.join(APP_DIR, "نتائج")
LOG_HOOK = None

FETCH_BATCH_JS = """
const urls = arguments[0], headers = arguments[1];
const done = arguments[arguments.length - 1];
Promise.all(urls.map(u =>
  fetch(u, {method: 'GET', credentials: 'include', headers: headers})
    .then(r => r.text().then(t => ({status: r.status, body: t})))
    .catch(e => ({status: 0, body: String(e)}))
)).then(done);
"""


def log_write(log, text):
    """log = ويدجت Text أو دالة. آمن من أي Thread."""
    try:
        if callable(log):
            log(text)
        elif log is not None:
            log.after(0, lambda: (log.insert("end", text), log.see("end")))
    except Exception:
        pass


# ============================================================================
# قراءة ملف المنافسين (openpyxl بدل pandas)
# ============================================================================
def _sheet_rows(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = [["" if v is None else str(v) for v in r] for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def read_pairs(path):
    """بالعرض: SKU | منافس 1 | منافس 2 ...   أو بالطول: SKU | SKU المنافس. الخانة تقبل اسكيو أو رابط."""
    rows = _sheet_rows(path)
    if len(rows) < 2:
        return []
    names = [str(h or "").strip().lower() for h in rows[0]]
    comp_idx = [i for i, n in enumerate(names) if "منافس" in n or "comp" in n]
    if not comp_idx:
        return []
    sku_idx = next((i for i, n in enumerate(names) if i not in comp_idx and ("sku" in n or "اسكيو" in n)), None)
    if sku_idx is None:
        sku_idx = next(i for i in range(len(names)) if i not in comp_idx)
    pairs = []
    for r in rows[1:]:
        r = list(r) + [""] * (len(names) - len(r))
        o = clean_sku(r[sku_idx])
        if not o:
            continue
        for ci in comp_idx:
            c = clean_sku(r[ci])
            if c and c != o and (o, c) not in pairs:
                pairs.append((o, c))
    return pairs


def load_mapping():
    if not os.path.exists(COMP_FILE):
        return []
    try:
        return read_pairs(COMP_FILE)
    except Exception:
        return []


# ============================================================================
# حفظ الإكسيل
# ============================================================================
class _Col:
    def __init__(self, vals):
        self.vals = vals

    def tolist(self):
        return self.vals


class _Frame:
    def __init__(self, cols, rows):
        self.columns = list(cols)
        self.rows = rows

    def __getitem__(self, c):
        return _Col([r.get(c, "") for r in self.rows])


def _clean(v):
    if v is None:
        return ""
    if isinstance(v, (int, float, str)):
        return v
    return str(v)


def save_excel(path, products=None, offers=None, comps=None, comp_products=None, wide=None):
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill
    green = PatternFill("solid", fgColor="C6EFCE")
    red = PatternFill("solid", fgColor="FFC7CE")
    sheets = []
    if products is not None:
        sheets.append(("بيانات منتجاتنا", PRODUCT_COLS, products,
                       lambda v: green if v.get("إحنا البوكس") == "نعم"
                       else (red if v.get("متجرنا موجود") == "نعم" else None)))
    if comp_products is not None:
        sheets.append(("بيانات المنافسين", ["منافس لـ"] + PRODUCT_COLS, comp_products, None))
    if offers is not None:
        sheets.append(("البائعين على منتجنا", OFFER_COLS, offers,
                       lambda v: green if v.get("متجرنا") == "نعم" else None))
    if comps is not None:
        sheets.append(("تفاصيل المنافسين", COMP_COLS, comps,
                       lambda v: red if v.get("أرخص من مين") == "المنافس أرخص"
                       else (green if v.get("أرخص من مين") == "إحنا أرخص" else None)))
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)
    for name, cols, rows, painter in sheets:
        ws = wb.create_sheet(name)
        ws.append(list(cols))
        for r in rows:
            ws.append([_clean(r.get(c, "")) for c in cols])
        format_sheet(ws, _Frame(cols, rows), painter)
    if wide is not None:
        write_wide_comparison(wb, *wide)
    wb.save(path)


# ============================================================================
# المتصفح + السحب المتوازي
# ============================================================================
def start_driver(headless=True, log=None):
    """كروم عادي ببروفايل مؤقت (صفحات نون العامة مش محتاجة تسجيل دخول)."""
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    profile_dir = tempfile.mkdtemp(prefix="NoonTrackChrome_")
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1366,900")
    opts.add_argument("--lang=en-US")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--disable-background-timer-throttling")
    opts.add_argument("--disable-backgrounding-occluded-windows")
    opts.add_argument("--disable-renderer-backgrounding")
    opts.add_argument(f"--user-data-dir={profile_dir}")
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    opts.set_capability("unhandledPromptBehavior", "accept")
    opts.page_load_strategy = "eager"
    try:
        driver = webdriver.Chrome(options=opts)
    except Exception:
        shutil.rmtree(profile_dir, ignore_errors=True)
        raise
    driver.set_script_timeout(60)
    driver.set_page_load_timeout(60)
    driver._tmp_profile = profile_dir
    return driver


def close_driver(driver):
    if not driver:
        return
    prof = getattr(driver, "_tmp_profile", None)
    try:
        driver.quit()
    except Exception:
        pass
    if prof:
        shutil.rmtree(prof, ignore_errors=True)


class Fetcher:
    """بيسحب صفحات المنتجات من جوه صفحة نون نفسها (نفس الكوكيز)، دفعات متوازية صغيرة + تهدئة تلقائية."""

    def __init__(self, driver, market, headless, log, stop):
        self.driver, self.market, self.headless, self.L, self.stop = driver, market, headless, log, stop

    def open_home(self):
        self.driver.get(f"https://www.noon.com/{self.market['path']}/")
        time.sleep(4)

    def _headers(self):
        h = dict(CATALOG_HEADERS)
        h["x-locale"] = self.market["locale"]
        return h

    def batch(self, skus):
        urls = [f"/_svc/catalog/api/v3/u/x/{s.split('-')[0]}/p/" for s in skus]
        try:
            res = self.driver.execute_async_script(FETCH_BATCH_JS, urls, self._headers())
        except Exception as e:
            self.L(f"خطأ اتصال: {str(e)[:120]}")
            try:                                   # لو الشباك مات نفتح واحد جديد
                self.driver.title
            except Exception:
                self.L("المتصفح وقع — بفتح واحد جديد")
                close_driver(self.driver)
                self.driver = start_driver(self.headless)
                self.open_home()
            return [(0, None)] * len(skus)
        out = []
        for r in res or []:
            st = (r or {}).get("status", 0)
            body = None
            if st == 200:
                try:
                    body = json.loads(r.get("body") or "{}")
                except Exception:
                    st = 0
            out.append((st, body))
        while len(out) < len(skus):
            out.append((0, None))
        return out


def fetch_all(fetcher, skus, market, our_set, conc, delay, log, stop, progress):
    """يرجع {sku: (status, row, offers)}. دفعات متوازية؛ لو نون بدأ يرفض بنقلّل السرعة ونستنى، وبعد 3 دفعات مرفوضة نقف."""
    L = lambda m: log_write(log, m + "\n")
    result, queue, retry = {}, list(skus), []
    cur = max(1, int(conc))
    clean_batches, blocked_streak, done = 0, 0, 0
    total = len(skus)

    def settle(sku, st, data):
        if st == 200 and data:
            row, offs = parse_product(sku, data, our_set, market["path"])
            return sku, (st, row, offs)
        if st == 404:
            return sku, (st, empty_product_row(sku, "المنتج غير موجود"), [])
        if st == 403:
            return sku, (st, empty_product_row(sku, "مرفوض من نون"), [])
        return None

    while queue and not stop():
        chunk, queue = queue[:cur], queue[cur:]
        res = fetcher.batch(chunk)
        n403 = n_ok = n_bad = 0
        for sku, (st, data) in zip(chunk, res):
            s = settle(sku, st, data)
            if s is None:
                retry.append(sku)
                n_bad += 1
                continue
            result[sku] = s[1]
            done += 1
            if st == 403:
                n403 += 1
            else:
                n_ok += 1
        progress(done, total)
        L(f"[{done}/{total}] دفعة {len(chunk)} (تمام {n_ok} • مرفوض {n403} • إعادة {n_bad}) • توازي {cur}")
        if n403 and not n_ok:
            blocked_streak += 1
            cur = 1
            clean_batches = 0
            if blocked_streak >= 3:
                L("⛔ نون رافض الطلبات — وقفت. استنى شوية وجرّب تاني")
                for sku in chunk:                      # المرفوضين نسيبهم يتحاولوا بعدين مش 'مرفوض'
                    result.pop(sku, None)
                break
            wait = 20 * blocked_streak
            L(f"⏸ رفض من نون — هستنى {wait} ثانية وأعيد فتح الصفحة")
            for sku in chunk:
                result.pop(sku, None)
                done -= 1
            queue = chunk + queue
            end = time.time() + wait
            while time.time() < end and not stop():
                time.sleep(1)
            try:
                fetcher.open_home()
            except Exception:
                pass
            continue
        blocked_streak = 0
        if n403 or n_bad >= max(2, len(chunk) // 2):
            cur = max(1, cur // 2)
            clean_batches = 0
        else:
            clean_batches += 1
            if clean_batches >= 5 and cur < int(conc):
                cur += 1
                clean_batches = 0
        time.sleep(delay + random.uniform(0, delay * 0.5))

    # إعادة المحاولة للّي فشل (واحد واحد وبهدوء)
    for sku in retry:
        if stop():
            break
        time.sleep(2)
        st, data = fetcher.batch([sku])[0]
        s = settle(sku, st, data)
        if s is None:
            result[sku] = (st, empty_product_row(sku, "فشل"), [])
        else:
            result[sku] = s[1]
        done += 1
        progress(done, total)
    return result


def sheet_fresh_skus(ss, hours):
    """اسكيوهات اتحدّثت في NoonTrack_Products من أقل من hours ساعة (ونجحت)."""
    try:
        vals = ss.worksheet(TAB_PRODUCTS).get_all_values()
    except Exception:
        return set()
    if not vals:
        return set()
    head = [h.strip() for h in vals[0]]
    if "SKU" not in head or STAMP_COL not in head:
        return set()
    si, ti = head.index("SKU"), head.index(STAMP_COL)
    ci = head.index("الحالة") if "الحالة" in head else None
    lim = datetime.now().timestamp() - hours * 3600
    out = set()
    for r in vals[1:]:
        r = list(r) + [""] * (len(head) - len(r))
        if ci is not None and r[ci] in FAILED_STATUSES:
            continue
        try:
            t = datetime.strptime(r[ti].strip(), "%Y-%m-%d %H:%M").timestamp()
        except Exception:
            continue
        if t >= lim and r[si]:
            out.add(strip_variant(r[si]))
    return out


def do_run(settings, log, stop=None, progress=None):
    L = lambda m: log_write(log, m + "\n")
    stop = stop or (lambda: False)
    progress = progress or (lambda d, t: None)
    market_name = settings.get("market") or "السعودية"
    market = MARKETS[market_name]
    our_set = {x.strip().lower() for x in re.split(r"[,،;\n]", str(settings.get("our_stores", ""))) if x.strip()}
    conc = int(settings.get("conc", 4))
    delay = float(settings.get("delay", 1.0))
    headless = bool(settings.get("headless", True))
    pairs = [(strip_variant(o), c) for o, c in load_mapping()]
    ours = ours_from_pairs(pairs)
    comp_skus = []
    for _, c in pairs:
        if c not in comp_skus:
            comp_skus.append(c)
    L(f"🏁 {market_name} — منافسين: {len(pairs)} ربط ({len(ours)} اسكيو تبعنا + {len(comp_skus)} منافس)")
    if not our_set:
        L("⚠️ خانة 'متاجرنا' فاضية — مش هعرف أميّز بوكسنا من المنافسين")

    gs = None
    want_gs = settings.get("sheet_skus") or settings.get("push") or settings.get("skip_fresh")
    if want_gs:
        try:
            gs = gs_open()
        except Exception as e:
            L(f"⚠️ تعذر الاتصال بجوجل شيت: {e}")
    if settings.get("sheet_skus") and gs:
        try:
            extra = [s for s in gs_target_skus(gs[1]) if s not in ours]
            ours = ours + extra
            L(f"📋 اسكيوهات تاب المبيعات: {len(extra)} إضافي (إجمالي اسكيوهاتنا {len(ours)})")
        except Exception as e:
            L(f"⚠️ تعذر قراءة اسكيوهات تاب المبيعات: {e}")

    skipped = 0
    if settings.get("skip_fresh") and gs:
        hours = float(settings.get("fresh_hours", 6))
        fresh = sheet_fresh_skus(gs[1], hours)
        protected = set(o for o, _ in pairs) | set(comp_skus)     # اللي ليهم منافسين لازم يتسحبوا عشان المقارنة
        keep = [s for s in ours if not (s in fresh and s not in protected)]
        skipped = len(ours) - len(keep)
        ours = keep
        L(f"⏭ اتخطّيت {skipped} اسكيو اتحدّثوا من أقل من {hours:g} ساعة")

    jobs = [("our", s) for s in ours] + [("comp", s) for s in comp_skus]
    if not jobs:
        L("مفيش اسكيوهات تتسحب (ضيف منافسين أو فعّل اسكيوهات تاب المبيعات)")
        return None
    skus = [s for _, s in jobs]
    driver = None
    fetched = {}
    try:
        L("🌐 فتح المتصفح...")
        driver = start_driver(headless)
        fetcher = Fetcher(driver, market, headless, log, stop)
        fetcher.open_home()
        L(f"🚀 بدء السحب: {len(skus)} صفحة")
        fetched = fetch_all(fetcher, skus, market, our_set, conc, delay, log, stop, progress)
        driver = fetcher.driver
    except Exception as e:
        L(f"❌ خطأ عام: {e}")
    finally:
        close_driver(driver)
    if stop():
        L("⏹ اتوقف")

    products, offers, comps, comp_products = [], [], [], []
    our_cache, comp_cache = {}, {}
    for kind, sku in jobs:
        if sku not in fetched:
            continue
        _, row, rows = fetched[sku]
        if kind == "our":
            our_cache[sku] = row
            products.append(row)
            offers.extend(rows)
    for kind, sku in jobs:
        if kind != "comp" or sku not in fetched:
            continue
        _, row, _ = fetched[sku]
        comp_cache[sku] = row
        owners = [o for o, c in pairs if c == sku]
        comp_products.append({"منافس لـ": " | ".join(owners), **row})
        comps.extend(build_comp_row(o, our_cache.get(o), sku, row) for o in owners)

    ok = sum(1 for r in products + comp_products if r.get("الحالة") not in FAILED_STATUSES)
    bad = len(products) + len(comp_products) - ok
    L(f"📊 اتسحب {ok} بنجاح • {bad} فشل/مرفوض • {len(comps)} صف مقارنة")
    if not (products or comps):
        L("مفيش بيانات اتسحبت عشان تتحفظ")
        return None

    if settings.get("push") and gs:
        try:
            gs_push_results(gs[0], gs[1], products, offers, comps, L)
            L("☁️ اترفعت النتائج لجوجل شيت")
        except Exception as e:
            L(f"❌ فشل رفع النتائج لجوجل شيت: {e}")
    path = ""
    if settings.get("xl") and not (settings.get("auto") and settings.get("push")):
        folder = (settings.get("folder") or "").strip() or RESULTS_DIR
        name = f"Noon_Tracker_{datetime.now():%Y%m%d_%H%M}.xlsx"
        fallback = os.path.join(os.path.expanduser("~"), "Documents", "noon_results", name)
        for target in (os.path.join(folder, name), fallback):
            try:
                save_excel(target, products=products, offers=offers, comps=comps,
                           comp_products=comp_products, wide=(pairs, our_cache, comp_cache))
                path = target
                L(f"✅ الإكسيل جاهز: {path}")
                break
            except Exception as e:
                L(f"⚠️ فشل الحفظ في {target}: {e}")
    return path or True


# ============================================================================
# الواجهة (Tk)
# ============================================================================
class App(tk.Frame):
    def __init__(self, root):
        super().__init__(root, bg="#d9d9d9")
        self.pack(fill="both", expand=True, padx=8, pady=8)
        self.running = False
        self._stop = False
        self.next_run = None
        self.auto_run = False
        self._last_fixed = ""
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")

        self.market = tk.StringVar(value="السعودية")
        self.our = tk.StringVar(value="")
        self.headless = tk.BooleanVar(value=True)
        self.conc = tk.IntVar(value=4)
        self.delay = tk.DoubleVar(value=1.0)
        self.sheet_skus = tk.BooleanVar(value=True)
        self.push = tk.BooleanVar(value=True)
        self.skip_fresh = tk.BooleanVar(value=True)
        self.fresh_hours = tk.IntVar(value=6)
        self.xl = tk.BooleanVar(value=False)
        self.folder = tk.StringVar(value=RESULTS_DIR)
        self.auto = tk.BooleanVar(value=False)
        self.every = tk.IntVar(value=120)
        self.times = tk.StringVar(value="")

        r = 0
        ttk.Label(top, text="السوق:").grid(row=r, column=0, sticky="e")
        ttk.Combobox(top, textvariable=self.market, values=list(MARKETS), state="readonly", width=12).grid(row=r, column=1, sticky="w")
        ttk.Label(top, text="متاجرنا (أكواد/أسماء بفاصلة):").grid(row=r, column=2, sticky="e", padx=(12, 0))
        ttk.Entry(top, textvariable=self.our, width=46).grid(row=r, column=3, columnspan=4, sticky="w")
        r += 1
        ttk.Label(top, text="توازي (صفحات مع بعض):").grid(row=r, column=0, sticky="e")
        ttk.Spinbox(top, from_=1, to=8, width=5, textvariable=self.conc).grid(row=r, column=1, sticky="w")
        ttk.Label(top, text="تهدئة بين الدفعات (ث):").grid(row=r, column=2, sticky="e", padx=(12, 0))
        ttk.Spinbox(top, from_=0.2, to=10, increment=0.2, width=5, textvariable=self.delay).grid(row=r, column=3, sticky="w")
        r += 1
        for v, t in [(self.headless, "تشغيل المتصفح مخفي (headless)"),
                     (self.sheet_skus, "ضيف اسكيوهات تاب المبيعات (NoonTrack_SKUs)"),
                     (self.push, "ارفع النتائج لجوجل شيت (NoonTrack_Products / Offers / Comps)")]:
            ttk.Checkbutton(top, text=t, variable=v).grid(row=r, column=0, columnspan=7, sticky="w")
            r += 1
        ttk.Checkbutton(top, text="⏭ تخطّى الاسكيوهات اللي اتحدّثت من أقل من", variable=self.skip_fresh).grid(row=r, column=0, columnspan=3, sticky="w")
        ttk.Spinbox(top, from_=1, to=168, width=5, textvariable=self.fresh_hours).grid(row=r, column=3, sticky="w")
        ttk.Label(top, text="ساعة (اللي ليهم منافسين بيتسحبوا دايمًا)").grid(row=r, column=4, columnspan=3, sticky="w")
        r += 1
        ttk.Checkbutton(top, text="احفظ ملف إكسيل (التشغيل التلقائي مع الرفع للشيت مابيعملوش)", variable=self.xl).grid(row=r, column=0, columnspan=4, sticky="w")
        r += 1
        ttk.Label(top, text="مكان الإكسيل:").grid(row=r, column=0, sticky="e")
        ttk.Entry(top, textvariable=self.folder, width=55).grid(row=r, column=1, columnspan=5, sticky="w")
        r += 1
        self.comp_lbl = ttk.Label(top, text="")
        self.comp_lbl.grid(row=r, column=0, columnspan=4, sticky="w", pady=(6, 0))
        ttk.Button(top, text="🔄 اعد القراءة", command=self.refresh_comp).grid(row=r, column=4, padx=3, pady=(6, 0))
        ttk.Button(top, text="📂 افتح ملف المنافسين", command=self.open_comp).grid(row=r, column=5, padx=3, pady=(6, 0))
        ttk.Button(top, text="🧩 أداة إدارة المنافسين", command=self.open_qt).grid(row=r, column=6, padx=3, pady=(6, 0))
        r += 1
        self.btn = ttk.Button(top, text="▶ شغّل مقارنة المنافسين", command=self.go)
        self.btn.grid(row=r, column=0, columnspan=2, pady=6, sticky="w")
        self.stop_btn = ttk.Button(top, text="⏹ إيقاف", command=self.stop)
        self.stop_btn.grid(row=r, column=2, pady=6, sticky="w")
        self.prog = ttk.Label(top, text="")
        self.prog.grid(row=r, column=3, columnspan=4, sticky="w")
        r += 1
        ttk.Checkbutton(top, text="⏰ تشغيل تلقائي كل", variable=self.auto, command=self.toggle_auto).grid(row=r, column=0, columnspan=2, sticky="w")
        ttk.Spinbox(top, from_=10, to=1440, increment=10, width=6, textvariable=self.every).grid(row=r, column=2, sticky="w")
        ttk.Label(top, text="دقيقة").grid(row=r, column=3, sticky="w")
        self.next_lbl = ttk.Label(top, text="")
        self.next_lbl.grid(row=r, column=4, columnspan=3, sticky="w")
        r += 1
        ttk.Label(top, text="🕐 مواعيد محددة (HH:MM مفصولة بفاصلة):").grid(row=r, column=0, columnspan=2, sticky="e")
        ttk.Entry(top, textvariable=self.times, width=34).grid(row=r, column=2, columnspan=4, sticky="w")

        self.out = scrolledtext.ScrolledText(self, font=("Consolas", 10), height=16)
        self.out.pack(fill="both", expand=True, padx=8, pady=8)

        self.load_settings()
        for v in (self.market, self.our, self.headless, self.conc, self.delay, self.sheet_skus, self.push,
                  self.skip_fresh, self.fresh_hours, self.xl, self.folder, self.every, self.times):
            v.trace_add("write", lambda *a: self.save_settings())
        self.refresh_comp()
        self.after(5000, self.tick)

    # ---------- مساعدات ----------
    def note(self, m):
        log_write(self.log, str(m) + "\n")

    def log(self, m):
        try:
            if LOG_HOOK:
                LOG_HOOK(str(m) if str(m).endswith("\n") else str(m) + "\n")
        except Exception:
            pass
        try:
            self.after(0, lambda: (self.out.insert("end", m if str(m).endswith("\n") else m + "\n"), self.out.see("end")))
        except Exception:
            pass

    def refresh_comp(self):
        n = len(load_mapping())
        if os.path.exists(COMP_FILE):
            self.comp_lbl.configure(text=f"📒 ملف المنافسين: {n} ربط ({os.path.basename(COMP_FILE)})")
        else:
            self.comp_lbl.configure(text="📒 مفيش ملف منافسين — هتتسحب اسكيوهات تاب المبيعات بس (أو استخدم أداة الإدارة)")

    def open_comp(self):
        try:
            if os.path.exists(COMP_FILE):
                os.startfile(COMP_FILE)
            else:
                messagebox.showinfo("", f"الملف مش موجود:\n{COMP_FILE}\nاعمله من أداة إدارة المنافسين.")
        except Exception as e:
            messagebox.showwarning("", str(e))

    def open_qt(self):
        script = os.path.join(APP_DIR, "noon_tracker.py")
        if not os.path.exists(script):
            messagebox.showinfo("", f"حط noon_tracker.py جنب البرنامج:\n{APP_DIR}")
            return
        try:
            exe = "py" if os.name == "nt" else "python3"
            subprocess.Popen([exe, script], cwd=APP_DIR)
        except Exception as e:
            messagebox.showwarning("", f"تعذّر التشغيل: {e}")

    # ---------- الإعدادات ----------
    _NAMES = ("market", "our", "headless", "conc", "delay", "sheet_skus", "push", "skip_fresh",
              "fresh_hours", "xl", "folder", "every", "times", "auto")

    def load_settings(self):
        d = {}
        try:
            d = json.load(open(SETTINGS_FILE, encoding="utf-8"))
        except Exception:
            try:                                  # أول مرة: ناخد متاجرنا والسوق من أداة الـ Qt
                q = json.load(open(QT_SETTINGS_FILE, encoding="utf-8"))
                d = {"our": q.get("our_stores", ""), "market": q.get("market", "السعودية")}
            except Exception:
                d = {}
        for n in self._NAMES:
            if n in d:
                try:
                    getattr(self, n).set(d[n])
                except Exception:
                    pass
        if self.market.get() not in MARKETS:
            self.market.set("السعودية")
        if self.auto.get():
            self.next_run = time.time() + 60
            self.note("⏰ التشغيل التلقائي مفعّل (من آخر مرة) — أول تشغيل بعد دقيقة")

    def settings_dict(self):
        d = {}
        for n in self._NAMES:
            try:
                d[n] = getattr(self, n).get()
            except Exception:
                pass
        return d

    def save_settings(self):
        try:
            json.dump(self.settings_dict(), open(SETTINGS_FILE, "w", encoding="utf-8"), ensure_ascii=False)
        except Exception:
            pass

    # ---------- الجدولة ----------
    def toggle_auto(self):
        if self.auto.get():
            self.next_run = time.time() + int(self.every.get()) * 60
            self.note(f"⏰ تشغيل تلقائي كل {int(self.every.get())} دقيقة (الجاي: {time.strftime('%H:%M', time.localtime(self.next_run))})")
        else:
            self.next_run = None
            self.note("⏰ التشغيل التلقائي اتقفل")
        self.save_settings()

    def fixed_times(self):
        out = []
        for t in str(self.times.get()).replace("،", ",").replace(" ", ",").split(","):
            t = t.strip()
            try:
                if t:
                    out.append(datetime.strptime(t, "%H:%M").strftime("%H:%M"))
            except ValueError:
                pass
        return out

    def tick(self):
        try:
            now = time.strftime("%H:%M")
            key = time.strftime("%Y-%m-%d ") + now
            if now in self.fixed_times() and key != self._last_fixed and not self.running:
                self._last_fixed = key
                self.note(f"🕐 موعد محدد {now} — تشغيل تلقائي...")
                self.go(auto=True)
        except Exception:
            pass
        try:
            if self.auto.get() and self.next_run:
                self.next_lbl.configure(text=f"التشغيل الجاي: {time.strftime('%H:%M', time.localtime(self.next_run))}")
                if self.next_run - time.time() <= 0 and not self.running:
                    self.note("⏰ تشغيل تلقائي...")
                    self.go(auto=True)
            else:
                self.next_lbl.configure(text="")
        except Exception:
            pass
        self.after(15000, self.tick)

    def shutdown(self):
        self.auto.set(False)
        self._stop = True

    def stop(self):
        if self.running:
            self._stop = True
            self.note("⏹ طلب إيقاف — هيقف بعد الدفعة الحالية")

    def _progress(self, d, t):
        self.after(0, lambda: self.prog.configure(text=f"{d}/{t}"))

    def go(self, auto=False):
        if self.running:
            return
        self.refresh_comp()
        if not auto and not self.sheet_skus.get() and not load_mapping():
            messagebox.showwarning("", "مفيش منافسين ولا اسكيوهات تاب المبيعات — فعّل واحد منهم.")
            return
        self.auto_run = auto
        self._stop = False
        self.save_settings()
        self.running = True
        self.btn.state(["disabled"])
        threading.Thread(target=self.job, daemon=True).start()

    def job(self):
        try:
            t0 = time.time()
            s = self.settings_dict()
            s["auto"] = self.auto_run
            do_run(s, self.log, stop=lambda: self._stop, progress=self._progress)
            self.note(f"⏱️ خلص في {int(time.time() - t0)} ثانية")
        except Exception as e:
            self.note(f"❌ {type(e).__name__}: {e}")
        finally:
            self.running = False
            if self.auto.get():
                self.next_run = time.time() + max(10, int(self.every.get())) * 60
            self.after(0, lambda: self.btn.state(["!disabled"]))


def build_ui(root):
    """بيتنادى من البرنامج الموحد: بيبني الواجهة جوه التاب ويرجّع دالة إيقاف."""
    return App(root).shutdown

'''


if __name__ == "__main__":
    main()
