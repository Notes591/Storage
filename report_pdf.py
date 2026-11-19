# -*- coding: utf-8 -*-
"""تقارير PDF زي G Rep / M Rep عند المحاسب (عربي من اليمين لليسار).
محتاج: pip install reportlab arabic-reshaper python-bidi   (على ويندوز: py -m pip install reportlab arabic-reshaper python-bidi)
الخط: بيدوّر على Tahoma / Arial من ويندوز (أو DejaVu/Noto على لينكس)."""
import io
import os

_FONT = None


def _ar(t):
    import arabic_reshaper
    from bidi.algorithm import get_display
    t = "" if t is None else str(t)
    return get_display(arabic_reshaper.reshape(t)) if t else t


def _font():
    global _FONT
    if _FONT:
        return _FONT
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    cands = [(r"C:\Windows\Fonts\tahoma.ttf", r"C:\Windows\Fonts\tahomabd.ttf"),
             (r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\arialbd.ttf"),
             ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
             ("/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf", "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Bold.ttf")]
    for reg, bold in cands:
        if os.path.exists(reg):
            pdfmetrics.registerFont(TTFont("RptF", reg))
            pdfmetrics.registerFont(TTFont("RptFB", bold if os.path.exists(bold) else reg))
            _FONT = ("RptF", "RptFB")
            return _FONT
    raise RuntimeError("مفيش خط عربي على الجهاز (Tahoma/Arial). نزّل خط وحطه في C:\\Windows\\Fonts")


def _n(v, d=0):
    try:
        t = f"{float(v):,.{d}f}"
        return "0" if t.lstrip("-") .replace(",", "").strip("0.") == "" else t
    except Exception:
        return ""


def build_report_pdf(title, subtitle, sections, grand, first_col="المتجر"):
    """sections: list من {"name","qty","cost","net","profit","rows":[(label,qty,cost,net,profit)]} · grand: {"qty","cost","net","profit"}
    بيرجع bytes بتاعت الـ PDF."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    reg, bold = _font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12 * mm, rightMargin=12 * mm, topMargin=12 * mm, bottomMargin=12 * mm,
                            title=title)
    ts = ParagraphStyle("t", fontName=bold, fontSize=14, alignment=1, leading=20)
    ss = ParagraphStyle("s", fontName=reg, fontSize=10, alignment=1, leading=14, textColor=colors.HexColor("#555555"))
    # الأعمدة من اليمين لليسار: الاسم | الكمية | تكلفة المبيعات | صافي المبيعات | الربح  → بنعكس الترتيب عشان يتعرض RTL
    hdr = [first_col, "الكمية", "تكلفة المبيعات", "صافي المبيعات", "الربح"]
    data = [[_ar(h) for h in reversed(hdr)]]
    style = [("FONT", (0, 0), (-1, -1), reg, 9), ("FONT", (0, 0), (-1, 0), bold, 9),
             ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
             ("ALIGN", (0, 0), (-1, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
             ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#d0d7de")),
             ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]

    def line(label, q, c, n, p, kind):
        r = len(data)
        data.append([_n(p), _n(n), _n(c), _n(q), _ar(label)])
        if kind == "sec":
            style.extend([("FONT", (0, r), (-1, r), bold, 9), ("BACKGROUND", (0, r), (-1, r), colors.HexColor("#ddebf7"))])
        elif kind == "grand":
            style.extend([("FONT", (0, r), (-1, r), bold, 10), ("BACKGROUND", (0, r), (-1, r), colors.HexColor("#1f4e79")),
                          ("TEXTCOLOR", (0, r), (-1, r), colors.white)])
        else:
            style.append(("LEFTPADDING", (4, r), (4, r), 0))
            style.append(("RIGHTPADDING", (4, r), (4, r), 14))
        if kind != "grand":
            style.append(("TEXTCOLOR", (0, r), (0, r), colors.HexColor("#c0392b") if p < 0 else colors.HexColor("#1e7b34")))

    for sc in sections:
        line(sc["name"], sc["qty"], sc["cost"], sc["net"], sc["profit"], "sec")
        for lb, q, c, n, p in sc["rows"]:
            line(lb, q, c, n, p, "row")
    if grand:
        line("Grand Total", grand["qty"], grand["cost"], grand["net"], grand["profit"], "grand")
    t = Table(data, colWidths=[30 * mm, 32 * mm, 30 * mm, 24 * mm, 66 * mm], repeatRows=1)
    t.setStyle(TableStyle(style))
    doc.build([Paragraph(_ar(title), ts), Paragraph(_ar(subtitle), ss), Spacer(1, 6 * mm), t])
    return buf.getvalue()
