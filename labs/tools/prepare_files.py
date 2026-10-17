# /// script
# requires-python = ">=3.10"
# dependencies = ["python-docx>=1.1", "xlsxwriter>=3.2"]
# ///
"""Generate the Al-Waha Shared Services Center file pack used by every lab (labs/files/).

    uv run labs/tools/prepare_files.py            # everything
    uv run labs/tools/prepare_files.py --no-pdf   # skip the PDFs (no Chrome needed)

Deterministic: the same seed gives the same files, so every participant and every answer key match. PDFs are printed
by headless Chrome (Arabic shaping and right-to-left layout come for free); set CHROME=/path/to/chrome if it is not
found. All names, numbers and organisations are fictional; national IDs and mobiles only look real.
"""
from __future__ import annotations

import argparse
import html
import os
import random
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path

import xlsxwriter
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

sys.path.insert(0, str(Path(__file__).parent))
import texts as T  # noqa: E402

LABS = Path(__file__).resolve().parents[1]
OUT = LABS / "files"
KEYS = LABS / "instructor" / "answer_keys.md"
SEED = 1448
TRAINING_NOTE = "نسخة تدريبية افتراضية · التصنيف المفترض للأصل: مقيّد"

# ================================================================ PDF (HTML printed by Chrome)
PDF_CSS = """
@page { size: A4; margin: 0; }
* { box-sizing: border-box; }
body { margin: 0; font-family: "IBM Plex Sans Arabic", "Segoe UI", Tahoma, sans-serif; color: #1e2a36; font-size: 12pt; line-height: 1.7; }
.page { width: 210mm; height: 297mm; padding: 30mm 20mm 22mm; position: relative; overflow: hidden; page-break-after: always; }
.page:last-child { page-break-after: auto; }
.head { position: absolute; top: 10mm; left: 20mm; right: 20mm; display: flex; justify-content: space-between; align-items: center;
        border-bottom: 2px solid #0F766E; padding-bottom: 3mm; font-size: 9.5pt; color: #4b5563; }
.head b { color: #0F766E; font-size: 11pt; }
.foot { position: absolute; bottom: 9mm; left: 20mm; right: 20mm; display: flex; justify-content: space-between; font-size: 9pt; color: #6b7280;
        border-top: 1px solid #d1d5db; padding-top: 2mm; }
.content { height: 100%; overflow: hidden; }
h1 { font-size: 19pt; color: #0F766E; margin: 0 0 4mm; line-height: 1.35; }
h2 { font-size: 14pt; color: #1E3A5F; margin: 6mm 0 2mm; }
h3 { font-size: 12.5pt; color: #1E3A5F; margin: 4mm 0 1mm; }
p { margin: 0 0 3mm; text-align: justify; }
ul, ol { margin: 0 0 3mm; padding-right: 7mm; padding-left: 0; }
li { margin-bottom: 1mm; }
table { width: 100%; border-collapse: collapse; margin: 2mm 0 4mm; font-size: 11pt; }
th { background: #0F766E; color: #fff; text-align: right; padding: 2mm 3mm; font-weight: 600; }
td { border-bottom: 1px solid #e5e7eb; padding: 1.8mm 3mm; vertical-align: top; }
tr:nth-child(even) td { background: #f8f6f1; }
.kv { display: grid; grid-template-columns: 42mm 1fr; border: 1px solid #e5e7eb; border-radius: 3mm; overflow: hidden; margin-bottom: 5mm; }
.kv div { padding: 1.8mm 3mm; border-bottom: 1px solid #eee; }
.kv div:nth-child(odd) { background: #F1E7D3; font-weight: 600; }
.note { background: #fff7e6; border-right: 4px solid #C8963E; padding: 3mm 4mm; margin: 3mm 0; border-radius: 2mm; }
.small { font-size: 9.5pt; color: #6b7280; }
.cover { display: flex; flex-direction: column; justify-content: center; height: 100%; }
.cover .t { font-size: 30pt; color: #0F766E; font-weight: 700; line-height: 1.3; }
.cover .s { font-size: 15pt; color: #1E3A5F; margin-top: 6mm; }
.chart { margin: 3mm 0 5mm; }
.ltr { direction: ltr; unicode-bidi: isolate; }
"""


def esc(s) -> str:
    return html.escape(str(s))


def chart_svg(spec: dict) -> str:
    """Horizontal bar chart: spec = {title, rows: [(label, value)], unit, max}."""
    rows, w, bar_h, gap = spec["rows"], 600, 22, 10
    top = 30
    mx = spec.get("max") or max(v for _, v in rows) * 1.1
    h = top + len(rows) * (bar_h + gap) + 10
    out = [f'<svg viewBox="0 0 {w} {h}" width="100%" xmlns="http://www.w3.org/2000/svg" direction="ltr" style="direction:ltr" font-family="IBM Plex Sans Arabic, Tahoma">',
           f'<text x="{w}" y="18" text-anchor="end" font-size="15" font-weight="700" fill="#1E3A5F">{esc(spec["title"])}</text>']
    for i, (label, v) in enumerate(rows):
        y = top + i * (bar_h + gap)
        bw = (v / mx) * 360
        out.append(f'<text x="{w}" y="{y + 16}" text-anchor="end" font-size="13" fill="#374151">{esc(label)}</text>')
        out.append(f'<rect x="{w - 210 - bw}" y="{y}" width="{bw:.1f}" height="{bar_h}" rx="4" fill="#0F766E"/>')
        val = f'{v:,}' if isinstance(v, int) else f'{v}'
        out.append(f'<text x="{w - 216 - bw}" y="{y + 16}" text-anchor="end" font-size="13" fill="#1E3A5F">{esc(val)}{esc(spec.get("unit", ""))}</text>')
    out.append("</svg>")
    return "".join(out)


def block_html(b) -> str:
    kind, val = b[0], b[1] if len(b) > 1 else None
    if kind in ("h1", "h2", "h3", "p"):
        return f"<{kind}>{esc(val)}</{kind}>"
    if kind == "small":
        return f'<p class="small">{esc(val)}</p>'
    if kind == "note":
        return f'<div class="note">{esc(val)}</div>'
    if kind in ("ul", "ol"):
        return f"<{kind}>" + "".join(f"<li>{esc(x)}</li>" for x in val) + f"</{kind}>"
    if kind == "table":
        head, *rows = val
        return ("<table><tr>" + "".join(f"<th>{esc(c)}</th>" for c in head) + "</tr>"
                + "".join("<tr>" + "".join(f"<td>{esc(c)}</td>" for c in r) + "</tr>" for r in rows) + "</table>")
    if kind == "kv":
        return '<div class="kv">' + "".join(f"<div>{esc(k)}</div><div>{esc(v)}</div>" for k, v in val) + "</div>"
    if kind == "chart":
        return f'<div class="chart">{chart_svg(val)}</div>'
    if kind == "cover":
        return f'<div class="cover"><div class="t">{esc(val["title"])}</div><div class="s">{esc(val["sub"])}</div><p class="small" style="margin-top:12mm">{esc(val["meta"])}</p></div>'
    if kind == "raw":
        return val
    raise ValueError(kind)


def doc_html(title: str, pages: list, header: str = T.CENTER, first_h1: bool = True) -> str:
    n = len(pages)
    body = []
    for i, page in enumerate(pages, 1):
        inner = "".join(block_html(b) for b in page)
        if i == 1 and first_h1 and not any(b[0] == "cover" for b in page):
            inner = f"<h1>{esc(title)}</h1>" + inner
        body.append(f'<section class="page"><div class="head"><b>{esc(header)}</b><span>{esc(TRAINING_NOTE)}</span></div>'
                    f'<div class="content">{inner}</div>'
                    f'<div class="foot"><span>{esc(title)}</span><span>صفحة {i} من {n}</span></div></section>')
    check = ("<script>window.addEventListener('load',()=>{const bad=[...document.querySelectorAll('.content')]"
             ".map((c,i)=>c.scrollHeight>c.clientHeight+2?(i+1)+'(+'+(c.scrollHeight-c.clientHeight)+'px)':0).filter(Boolean);document.title='OVERFLOW:'+bad.join(',');});</script>")
    return (f'<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>{esc(title)}</title>'
            '<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Arabic:wght@400;600;700&display=swap" rel="stylesheet">'
            f"<style>{PDF_CSS}</style>{check}</head><body>{''.join(body)}</body></html>")


def find_chrome() -> str | None:
    cands = [os.environ.get("CHROME"), shutil.which("google-chrome"), shutil.which("chromium"), shutil.which("chrome"),
             r"C:\Program Files\Google\Chrome\Application\chrome.exe", r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
             "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"]
    return next((c for c in cands if c and Path(c).exists()), None)


def write_pdf(path: Path, html_text: str, chrome: str) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "doc.html"
        src.write_text(html_text, encoding="utf-8", newline="\n")
        url = src.as_uri()
        dom = subprocess.run([chrome, "--headless=new", "--disable-gpu", "--virtual-time-budget=8000", "--dump-dom", url],
                             capture_output=True, text=True, encoding="utf-8").stdout
        m = re.search(r"<title>OVERFLOW:([^<]*)</title>", dom)
        if m and m.group(1):
            print(f"  ! {path.name}: content overflows on page(s) {m.group(1)}")
        subprocess.run([chrome, "--headless=new", "--disable-gpu", "--no-pdf-header-footer", "--virtual-time-budget=8000",
                        f"--print-to-pdf={path}", url], capture_output=True, check=True)
    print(f"  {path.relative_to(LABS)}")


FIXED_TIME = datetime(2026, 10, 1, 8, 0, 0)  # file metadata, so regenerating gives identical XLSX/DOCX


def new_wb(path: Path):
    wb = xlsxwriter.Workbook(path)
    wb.set_properties({"author": T.CENTER, "created": FIXED_TIME})
    return wb


def fix_zip_times(path: Path) -> None:
    """python-docx stamps zip entries with the current time; rewrite them with FIXED_TIME."""
    with zipfile.ZipFile(path) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in items:
            info.date_time = FIXED_TIME.timetuple()[:6]
            z.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)


# ================================================================ DOCX (right-to-left)
def _rtl_paragraph(p):
    ppr = p._p.get_or_add_pPr()
    ppr.append(OxmlElement("w:bidi"))
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for r in p.runs:
        _rtl_run(r)


def _rtl_run(r, size=12, bold=False, color=None):
    rpr = r._r.get_or_add_rPr()
    rpr.append(OxmlElement("w:rtl"))
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.append(fonts)
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        fonts.set(qn(a), "Arial")
    r.font.size = Pt(size)
    r.font.bold = bold
    szcs = OxmlElement("w:szCs")
    szcs.set(qn("w:val"), str(size * 2))
    rpr.append(szcs)
    if bold:
        rpr.append(OxmlElement("w:bCs"))
    if color:
        r.font.color.rgb = RGBColor.from_string(color)


def docx_par(d, text, size=12, bold=False, color=None, style=None):
    p = d.add_paragraph(style=style)
    r = p.add_run(text)
    _rtl_run(r, size, bold, color)
    ppr = p._p.get_or_add_pPr()
    ppr.append(OxmlElement("w:bidi"))
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    return p


def docx_table(d, rows):
    t = d.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    tblpr = t._tbl.tblPr
    tblpr.append(OxmlElement("w:bidiVisual"))
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = t.cell(i, j)
            cell.text = ""
            p = cell.paragraphs[0]
            r = p.add_run(str(val))
            _rtl_run(r, 11, bold=(i == 0))
            p._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    d.add_paragraph()


def write_docx(path: Path, title: str, pages: list, header=T.CENTER) -> None:
    d = Document()
    cp = d.core_properties
    cp.author, cp.last_modified_by, cp.created, cp.modified, cp.revision = T.CENTER, T.CENTER, FIXED_TIME, FIXED_TIME, 1
    sec = d.sections[0]
    sec.header.paragraphs[0].text = ""
    hp = sec.header.paragraphs[0]
    _rtl_run(hp.add_run(f"{header} · {TRAINING_NOTE}"), 9, color="6B7280")
    hp._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    docx_par(d, title, 18, True, "0F766E")
    for pi, page in enumerate(pages):
        if pi:
            d.add_page_break()
        for b in page:
            kind, val = b[0], b[1]
            if kind in ("h1", "h2", "h3"):
                docx_par(d, val, {"h1": 16, "h2": 14, "h3": 12}[kind], True, "1E3A5F")
            elif kind in ("p", "note"):
                docx_par(d, val)
            elif kind == "small":
                docx_par(d, val, 9, color="6B7280")
            elif kind in ("ul", "ol"):
                for k, x in enumerate(val, 1):
                    docx_par(d, (f"{k}. " if kind == "ol" else "• ") + x)
            elif kind == "table":
                docx_table(d, val)
            elif kind == "kv":
                docx_table(d, [[k, v] for k, v in val])
    d.save(path)
    fix_zip_times(path)
    print(f"  {path.relative_to(LABS)}")


# ================================================================ data
def gen_inquiries(rng: random.Random):
    rows, topics = [], []
    start = date(2026, 7, 1)
    weights = [0.22, 0.14, 0.20, 0.17, 0.15, 0.12]  # entity mix
    for topic, (count, templates) in T.INQUIRY_TOPICS.items():
        for _ in range(count):
            text = rng.choice(templates).format(
                e=rng.choice(T.ENTITIES), s=rng.choice(T.SERVICES), n=rng.randint(70000, 79999), d=rng.randint(5, 21))
            topics.append(topic)
            rows.append(text)
    order = list(range(len(rows)))
    rng.shuffle(order)
    records = []
    for k, i in enumerate(order):
        female = rng.random() < 0.48
        first = rng.choice(T.FIRST_F if female else T.FIRST_M)
        last = rng.choice(T.LAST)
        d = start + timedelta(days=int(k * 91 / len(order)) + rng.randint(0, 1))
        d = min(d, date(2026, 9, 30))
        status = rng.choice(T.STATUSES)
        records.append({
            "id": f"INQ-{26000 + k + 1}", "date": d, "entity": rng.choices(T.ENTITIES, weights)[0],
            "channel": rng.choice(T.CHANNELS), "name": f"{first} {last}",
            "mobile": "05" + "".join(str(rng.randint(0, 9)) for _ in range(8)),
            "email": f"user{rng.randint(1000, 9999)}@mail.example",
            "nid": "1" + "".join(str(rng.randint(0, 9)) for _ in range(9)),
            "text": rows[i], "topic": topics[i], "status": status,
            "days": "" if status == "قيد المعالجة" else rng.randint(1, 14),
        })
    dups = rng.sample(records, 6)  # deliberate duplicates, re-sent the next day
    for r in dups:
        records.append(dict(r))
    records.sort(key=lambda r: (r["date"], r["id"]))
    return records, dups


def write_inquiries(records) -> None:
    path = OUT / "beneficiary_inquiries_q3.xlsx"
    wb = new_wb(path)
    ws = wb.add_worksheet("الاستفسارات")
    ws.right_to_left()
    head = wb.add_format({"bold": True, "bg_color": "#0F766E", "font_color": "#FFFFFF", "border": 1, "font_name": "Arial"})
    cell = wb.add_format({"font_name": "Arial", "text_wrap": True, "valign": "top"})
    dfmt = wb.add_format({"font_name": "Arial", "num_format": "yyyy-mm-dd", "valign": "top"})
    cols = [("رقم الطلب", 12), ("تاريخ الاستلام", 13), ("الجهة", 28), ("القناة", 15), ("اسم المستفيد", 20), ("الجوال", 13),
            ("البريد الإلكتروني", 24), ("رقم الهوية", 13), ("نص الاستفسار", 60), ("الحالة", 13), ("مدة المعالجة (يوم)", 10)]
    for j, (c, w) in enumerate(cols):
        ws.write(0, j, c, head)
        ws.set_column(j, j, w)
    for i, r in enumerate(records, 1):
        vals = [r["id"], r["date"], r["entity"], r["channel"], r["name"], r["mobile"], r["email"], r["nid"], r["text"], r["status"], r["days"]]
        for j, v in enumerate(vals):
            if isinstance(v, date):
                ws.write_datetime(i, j, datetime(v.year, v.month, v.day), dfmt)
            else:
                ws.write_string(i, j, str(v), cell) if j != 10 or v == "" else ws.write_number(i, j, v, cell)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(records), len(cols) - 1)
    info = wb.add_worksheet("عن الملف")
    info.right_to_left()
    info.set_column(0, 0, 100)
    for k, line in enumerate([
        "استفسارات مكتوبة وردت إلى مركز الواحة للخدمات المشتركة عبر البوابة والبريد ونموذج الموقع، من 1 يوليو إلى 30 سبتمبر 2026م.",
        "الملف تدريبي افتراضي: الأسماء والجوالات وأرقام الهوية وهمية، لكنها تُعامل كبيانات شخصية لأغراض التمرين.",
        "قبل رفع الملف إلى أي تطبيق: احذف الأعمدة الشخصية أو رمّزها، واحتفظ بجدول الربط على جهازك.",
        "لا يوجد عمود للتصنيف: تصنيف الاستفسارات جزء من التمرين.",
    ]):
        info.write(k, 0, line, cell)
    wb.close()
    print(f"  {path.relative_to(LABS)}")


def write_budget() -> list[str]:
    """Training budget with deliberate errors; returns the list of errors for the answer key."""
    path = OUT / "training_budget_2026.xlsx"
    wb = new_wb(path)
    ws = wb.add_worksheet("خطة التدريب 2026")
    ws.right_to_left()
    head = wb.add_format({"bold": True, "bg_color": "#0F766E", "font_color": "#FFFFFF", "border": 1, "text_wrap": True, "font_name": "Arial"})
    txt = wb.add_format({"font_name": "Arial", "border": 1})
    num = wb.add_format({"font_name": "Arial", "border": 1, "num_format": "#,##0"})
    pct = wb.add_format({"font_name": "Arial", "border": 1, "num_format": "0%"})
    tot = wb.add_format({"font_name": "Arial", "border": 1, "bold": True, "num_format": "#,##0", "bg_color": "#F1E7D3"})
    heads = ["البرنامج", "الإدارة", "عدد المتدربين", "تكلفة المتدرب (ريال)", "التكلفة المخططة", "المصروف حتى سبتمبر", "المتبقي", "نسبة الصرف"]
    widths = [38, 22, 12, 16, 16, 18, 14, 12]
    for j, (h, w) in enumerate(zip(heads, widths)):
        ws.write(0, j, h, head)
        ws.set_column(j, j, w)
    progs = [
        ("مهارات خدمة المستفيدين المتقدمة", "خدمة المستفيدين", 40, 1500, 52000),
        ("التدريب على بوابة الخدمات الذاتية", "خدمة المستفيدين", 40, 900, 0),
        ("الطلاقة في الذكاء الاصطناعي", "جميع الإدارات", 60, 2200, 61000),
        ("إدارة المشاريع الاحترافية", "تقنية المعلومات", 9, 5000, 45000),  # E hardcoded 48,000 below
        ("التحليل المالي للقطاع العام", "المالية", 12, 4200, 30400),
        ("المشتريات الحكومية", "المشتريات والعقود", 10, 3800, 38000),
        ("الكتابة الرسمية والمراسلات", "جميع الإدارات", 45, 800, 21600),
        ("الأمن السيبراني للموظفين", "جميع الإدارات", 120, 350, 42000),  # G formula points at wrong row
        ("قيادة الفرق", "الإدارة العليا", 14, 6500, 52000),
        ("تحليل البيانات بالجداول", "التخطيط والأداء", 18, 2400, 19200),
        ("إمكانية الوصول الرقمي", "تقنية المعلومات", 6, "3,200 ريال", 9600),  # text number
        ("التفاوض مع الموردين", "المشتريات والعقود", 8, 14500, 58000),  # above the 12,000 policy cap
        ("الإسعافات الأولية", "الخدمات العامة", 30, 450, 10800),  # % inverted below
        ("العرض والإلقاء", "الاتصال المؤسسي", 10, 1900, 15200),
    ]
    first, last = 2, 1 + len(progs)  # Excel rows (1-based)
    for i, (name, dept, n, cost, spent) in enumerate(progs):
        r = i + 1  # 0-based row index
        R = r + 1  # Excel row number
        ws.write(r, 0, name, txt)
        ws.write(r, 1, dept, txt)
        ws.write_number(r, 2, n, num)
        if isinstance(cost, str):
            ws.write_string(r, 3, cost, txt)
        else:
            ws.write_number(r, 3, cost, num)
        if name == "إدارة المشاريع الاحترافية":
            ws.write_number(r, 4, 48000, num)  # hardcoded, should be C*D = 45,000
            planned = 48000
        elif isinstance(cost, str):
            ws.write_formula(r, 4, f"=C{R}*D{R}", num, "#VALUE!")
            planned = None
        else:
            planned = n * cost
            ws.write_formula(r, 4, f"=C{R}*D{R}", num, planned)
        ws.write_number(r, 5, spent, num)
        if name == "الأمن السيبراني للموظفين":
            prev_spent = progs[i - 1][4]
            ws.write_formula(r, 6, f"=E{R}-F{R - 1}", num, planned - prev_spent)
        elif planned is None:
            ws.write_formula(r, 6, f"=E{R}-F{R}", num, "#VALUE!")
        else:
            ws.write_formula(r, 6, f"=E{R}-F{R}", num, planned - spent)
        if name == "الإسعافات الأولية":
            ws.write_formula(r, 7, f"=E{R}/F{R}", pct, planned / spent)
        elif planned is None:
            ws.write_formula(r, 7, f"=F{R}/E{R}", pct, "#VALUE!")
        else:
            ws.write_formula(r, 7, f"=F{R}/E{R}", pct, (spent / planned) if planned else 0)
    tr = last  # 0-based index of the total row
    ws.write(tr, 0, "الإجمالي", tot)
    ws.write(tr, 1, "", tot)
    ws.write_formula(tr, 2, f"=SUM(C{first}:C{last})", tot, sum(p[2] for p in progs))
    ws.write(tr, 3, "", tot)
    ws.write_formula(tr, 4, f"=SUM(E{first}:E{last - 3})", tot, "#VALUE!")  # misses the last 3 rows, and hits #VALUE!
    ws.write_formula(tr, 5, f"=SUM(F{first}:F{last})", tot, sum(p[4] for p in progs))
    ws.write(tr, 6, "", tot)
    ws.write(tr, 7, "", tot)
    s = wb.add_worksheet("ملخص للإدارة")
    s.right_to_left()
    s.set_column(0, 1, 34)
    s.write(0, 0, "البند", head)
    s.write(0, 1, "القيمة", head)
    s.write(1, 0, "الميزانية المعتمدة للتدريب 2026", txt)
    s.write_number(1, 1, 520000, num)
    s.write(2, 0, "إجمالي التكلفة المخططة", txt)
    s.write_formula(2, 1, "='خطة التدريب 2026'!E16", num, "#VALUE!")
    s.write(3, 0, "إجمالي المصروف حتى سبتمبر", txt)
    s.write_formula(3, 1, "=SUM(#REF!)", num, "#REF!")
    s.write(4, 0, "المتبقي من الميزانية المعتمدة", txt)
    s.write_formula(4, 1, "=B2-B4", num, "#REF!")
    wb.close()
    print(f"  {path.relative_to(LABS)}")
    return [
        "«إدارة المشاريع الاحترافية»: التكلفة المخططة مكتوبة يدويًا 48,000 بدل الصيغة =C*D (الصحيح 45,000).",
        "«الأمن السيبراني للموظفين»: المتبقي =E-F للصف السابق (F من الصف الأعلى) بدل الصف نفسه.",
        "«إمكانية الوصول الرقمي»: تكلفة المتدرب نص «3,200 ريال» لا رقم، فتظهر #VALUE! في التكلفة والمتبقي والنسبة.",
        "«الإسعافات الأولية»: نسبة الصرف =E/F مقلوبة (يجب =F/E).",
        "صف الإجمالي: =SUM(E2:E12) لا يشمل آخر ثلاثة برامج، ويتأثر بخطأ #VALUE!.",
        "ورقة «ملخص للإدارة»: المصروف =SUM(#REF!) مرجع محذوف، فيتعطل حساب المتبقي.",
        "«التفاوض مع الموردين»: تكلفة المتدرب 14,500 تتجاوز الحد الأعلى 12,000 في أسئلة الموارد البشرية الشائعة (يحتاج اعتمادًا).",
        "«التدريب على بوابة الخدمات الذاتية»: لا مصروف (0) لأن التدريب في نوفمبر؛ ليس خطأ.",
    ]


def write_kpis(rng: random.Random) -> None:
    """Monthly service KPIs per entity, January–September 2026; Q3 totals match the quarterly report."""
    path = OUT / "service_kpis_2026.xlsx"
    wb = new_wb(path)
    ws = wb.add_worksheet("المؤشرات الشهرية")
    ws.right_to_left()
    head = wb.add_format({"bold": True, "bg_color": "#0F766E", "font_color": "#FFFFFF", "border": 1, "font_name": "Arial", "text_wrap": True})
    cell = wb.add_format({"font_name": "Arial", "border": 1})
    pct = wb.add_format({"font_name": "Arial", "border": 1, "num_format": "0%"})
    one = wb.add_format({"font_name": "Arial", "border": 1, "num_format": "0.0"})
    cols = ["الشهر", "الجهة", "الطلبات", "متوسط مدة المعالجة (يوم)", "رضا المستفيدين", "الحل من أول تواصل", "نسبة القنوات الرقمية"]
    for j, c in enumerate(cols):
        ws.write(0, j, c, head)
    ws.set_column(0, 0, 10)
    ws.set_column(1, 1, 30)
    ws.set_column(2, 6, 14)
    months = ["2026-01", "2026-02", "2026-03", "2026-04", "2026-05", "2026-06", "2026-07", "2026-08", "2026-09"]
    r = 1
    for ent, q3, days, sat in T.REPORT_ENTITIES:
        sat_v = int(sat.rstrip("%")) / 100
        split = [0.327, 0.335]
        q3m = [round(q3 * split[0]), round(q3 * split[1])]
        q3m.append(q3 - sum(q3m))
        for mi, m in enumerate(months):
            if mi >= 6:
                req = q3m[mi - 6]
            else:
                req = round(q3 / 3 * rng.uniform(0.86, 0.98))
            trend = (8 - mi) * 0.06
            ws.write(r, 0, m, cell)
            ws.write(r, 1, ent, cell)
            ws.write_number(r, 2, req, cell)
            ws.write_number(r, 3, round(days + trend + rng.uniform(-0.2, 0.2), 1), one)
            ws.write_number(r, 4, round(min(0.95, sat_v - trend / 6 + rng.uniform(-0.015, 0.015)), 2), pct)
            ws.write_number(r, 5, round(0.55 + mi * 0.014 + rng.uniform(-0.02, 0.02), 2), pct)
            ws.write_number(r, 6, round(0.28 + mi * 0.016 + rng.uniform(-0.02, 0.02), 2), pct)
            r += 1
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, r - 1, len(cols) - 1)
    wb.close()
    print(f"  {path.relative_to(LABS)}")


def write_glossary() -> None:
    path = OUT / "glossary_en_ar.xlsx"
    wb = new_wb(path)
    ws = wb.add_worksheet("المسرد")
    ws.right_to_left()
    head = wb.add_format({"bold": True, "bg_color": "#0F766E", "font_color": "#FFFFFF", "font_name": "Arial"})
    cell = wb.add_format({"font_name": "Arial"})
    ws.write(0, 0, "المصطلح الإنجليزي", head)
    ws.write(0, 1, "المقابل العربي المعتمد في المركز", head)
    ws.set_column(0, 1, 40)
    for i, (en, ar) in enumerate(T.GLOSSARY, 1):
        ws.write(i, 0, en, cell)
        ws.write(i, 1, ar, cell)
    wb.close()
    print(f"  {path.relative_to(LABS)}")


def write_kit() -> None:
    """The participant's personal productivity kit: task map, productivity log, prompt library (Labs 1.1, 1.4 and on)."""
    path = OUT / "my_productivity_kit.xlsx"
    wb = new_wb(path)
    head = wb.add_format({"bold": True, "bg_color": "#1C355E", "font_color": "#FFFFFF", "border": 1, "text_wrap": True, "valign": "top", "font_name": "Arial"})
    cell = wb.add_format({"font_name": "Arial", "border": 1, "text_wrap": True, "valign": "top"})
    ex = wb.add_format({"font_name": "Arial", "border": 1, "text_wrap": True, "valign": "top", "font_color": "#5B6472", "italic": True})
    num = wb.add_format({"font_name": "Arial", "border": 1, "valign": "top"})
    sheets = [
        ("خريطة المهام", [("المهمة", 34), ("كم مرة شهريًا", 12), ("الوقت المعتاد (دقيقة)", 14), ("مستوى التحقق المطلوب", 16),
                          ("النمط", 14), ("أين قد يخطئ التطبيق؟", 40)],
         ["تحديث المدير بقرارات لجنة", 4, 35, "مرتفع", "أتمتة", "ينسب موعدًا أو مسؤولًا غير وارد في المحضر"],
         {3: ["منخفض", "متوسط", "مرتفع"], 4: ["أتمتة", "تعزيز", "توكيل", "لي وحدي"]}),
        ("سجل الإنتاجية", [("التاريخ", 12), ("المهمة", 30), ("التطبيق", 14), ("قبل (دقيقة)", 11), ("بعد (دقيقة)", 11), ("الوفر (دقيقة)", 11),
                           ("معيار القبول", 36), ("تحقّق؟", 12), ("ما صحّحته", 32)],
         ["2026-10-18", "تحديث المدير بقرارات اللجنة", "Claude", 35, 12, None, "180 كلمة؛ لكل بند مسؤول وموعد من المحضر", "بعد تصحيح", "موعد منسوب لغير مصدره"],
         {7: ["نعم", "بعد تصحيح", "لا"]}),
        ("مكتبة الأوامر", [("الاسم (فعل + مخرج)", 24), ("متى أستخدمه وعلى أي مادة", 28), ("الأمر بعناصر نائبة [ ]", 60), ("معيار القبول", 32),
                           ("أخطاء شائعة وكيف أتحقق منها", 34)],
         ["تحديث تنفيذي من محضر", "بعد كل اجتماع لجنة لا يحضره المدير", "تصرّف كمساعد تنفيذي… القارئ: [المنصب]… بريد لا يتجاوز [عدد] كلمة…",
          "لكل بند مسؤول وموعد من المحضر؛ لا معلومة غير واردة", "يختلق موعدًا: أطابق كل موعد مع المحضر"], {}),
    ]
    for name, cols, example, lists in sheets:
        ws = wb.add_worksheet(name)
        ws.right_to_left()
        for j, (c, w) in enumerate(cols):
            ws.write(0, j, c, head)
            ws.set_column(j, j, w)
        for j, v in enumerate(example):
            if name == "سجل الإنتاجية" and j == 5:
                ws.write_formula(1, j, "=IF(AND(D2<>\"\",E2<>\"\"),D2-E2,\"\")", num, 23)
            elif isinstance(v, int):
                ws.write_number(1, j, v, ex)
            else:
                ws.write(1, j, v, ex)
        for r in range(2, 40):
            for j in range(len(cols)):
                if name == "سجل الإنتاجية" and j == 5:
                    ws.write_formula(r, j, f'=IF(AND(D{r + 1}<>"",E{r + 1}<>""),D{r + 1}-E{r + 1},"")', num, "")
                else:
                    ws.write_blank(r, j, None, cell)
        for j, opts in lists.items():
            ws.data_validation(1, j, 39, j, {"validate": "list", "source": opts})
        ws.freeze_panes(1, 0)
    t = wb.add_worksheet("الإجمالي")
    t.right_to_left()
    t.set_column(0, 1, 30)
    for r, (k, f) in enumerate([("مهام مسجّلة", "=COUNTA('سجل الإنتاجية'!B3:B40)"), ("الوقت الموفّر (دقيقة)", "=SUM('سجل الإنتاجية'!F3:F40)"),
                                ("مهام تحقّقت من أول مرة", "=COUNTIF('سجل الإنتاجية'!H3:H40,\"نعم\")"),
                                ("مهام لم تنجح", "=COUNTIF('سجل الإنتاجية'!H3:H40,\"لا\")")]):
        t.write(r, 0, k, head)
        t.write_formula(r, 1, f, num, 0)
    t.write(5, 0, "الصف الثاني في كل ورقة مثال: احذفه أو اكتب فوقه. الإجمالي يبدأ من الصف الثالث.", cell)
    wb.close()
    print(f"  {path.relative_to(LABS)}")


# ================================================================ quarterly report pages
def report_pages() -> list:
    kpi_rows = [["المؤشر", "الربع الثاني", "الربع الثالث", "المستهدف"]] + [list(k) for k in T.REPORT_KPIS]
    ent_rows = [["الجهة", "الطلبات", "متوسط المدة (يوم)", "الرضا"]] + [[e, f"{n:,}", d, s] for e, n, d, s in T.REPORT_ENTITIES]
    topic_counts = [(t, c) for t, (c, _) in T.INQUIRY_TOPICS.items()]
    return [
        [("cover", {"title": "تقرير الأداء الربعي", "sub": f"{T.CENTER} · {T.QUARTER}",
                    "meta": "إعداد: إدارة التخطيط والأداء · تاريخ الإصدار: 12 أكتوبر 2026م · الإصدار 1.0"})],
        [("h1", "الملخص التنفيذي"),
         ("p", "حقق المركز في الربع الثالث تحسنًا في أغلب مؤشراته الرئيسية؛ إذ عالج 48,600 طلب بزيادة 7.5% عن الربع الثاني، "
               "وانخفض متوسط مدة المعالجة من 4.2 إلى 3.8 يوم عمل، وارتفع رضا المستفيدين إلى 84%."),
         ("p", "وما زالت ثلاثة مؤشرات دون المستهدف: مدة المعالجة (المستهدف 3 أيام)، ونسبة القنوات الرقمية (41% مقابل 50%)، "
               "والالتزام باتفاقيات مستوى الخدمة (91% مقابل 95%). ويتركز الأداء الأضعف في صندوق السدرة للإقراض."),
         ("h2", "أبرز ما يحتاج قرارًا"),
         ("ol", ["اعتماد أمر التغيير الخاص بمتطلبات الوصول في بوابة الخدمات الذاتية (186,000 ريال) قبل 22 أكتوبر.",
                 "تعزيز مكتب خدمة المستفيدين مؤقتًا في أسبوع إطلاق البوابة لمواجهة زيادة متوقعة بنحو 30% في الاستفسارات.",
                 "إعادة جدولة مشروع نظام إدارة الوثائق بعد تعثره (نسبة الإنجاز 34% مقابل 60% مخطط)."]),
         ("h2", "أرقام الربع في سطر"),
         ("kv", [("الطلبات المعالجة", "48,600"), ("متوسط مدة المعالجة", "3.8 يوم عمل"), ("القنوات الرقمية", "41%"), ("تنفيذ الميزانية حتى سبتمبر", "71%")])],
        [("h1", "1. المؤشرات الرئيسية"), ("table", kpi_rows),
         ("chart", {"title": "الطلبات المعالجة شهريًا في الربع الثالث", "rows": [("يوليو", 15900), ("أغسطس", 16300), ("سبتمبر", 16400)], "max": 18000}),
         ("p", "ارتفعت الطلبات في سبتمبر بسبب بدء العام الدراسي وطلبات تحديث البيانات، مع ثبات عدد الموظفين.")],
        [("h1", "2. الأداء حسب الجهة المستفيدة"), ("table", ent_rows),
         ("chart", {"title": "متوسط مدة المعالجة (يوم عمل)", "rows": [(e, d) for e, _, d, _ in T.REPORT_ENTITIES], "max": 5.5}),
         ("p", "سجل صندوق السدرة للإقراض أطول مدة معالجة وأدنى رضا، وترتبط معظم طلباته بالمستحقات المالية وشهادات الراتب التي تعتمد على بيانات من أنظمة متعددة.")],
        [("h1", "3. قنوات الخدمة"),
         ("chart", {"title": "توزيع الطلبات حسب القناة", "rows": [("البوابة الإلكترونية", 41), ("الهاتف", 34), ("البريد الإلكتروني", 15), ("الحضور الشخصي", 10)], "unit": "%", "max": 50}),
         ("p", "ارتفعت حصة البوابة من 36% إلى 41%، وما زال الهاتف القناة الثانية. ومتوسط الانتظار الهاتفي 2.6 دقيقة في الربع، وبلغ 4.5 دقائق في الأسبوع الأخير من سبتمبر بسبب نقص الموظفين."),
         ("p", "ويُتوقع أن ترتفع حصة القنوات الرقمية بعد إطلاق بوابة الخدمات الذاتية في 15 نوفمبر 2026م.")],
        [("h1", "4. الاستفسارات المكتوبة"),
         ("p", "وصل إلى المركز 452 استفسارًا مكتوبًا عبر البوابة والبريد ونموذج الموقع، صُنّفت يدويًا كما يلي:"),
         ("table", [["الموضوع", "العدد"]] + [[t, str(c)] for t, c in topic_counts] + [["الإجمالي", "452"]]),
         ("chart", {"title": "الاستفسارات المكتوبة حسب الموضوع", "rows": topic_counts, "max": 130})],
        [("h1", "5. الموارد البشرية"),
         ("table", [["المؤشر", "القيمة"], ["عدد الموظفين", "312"], ["الوظائف الشاغرة", "14"], ["معدل الدوران الوظيفي (سنوي)", "4.1%"],
                    ["ساعات التدريب المنفذة", "6,240 ساعة"], ["متوسط ساعات التدريب للموظف", "20 ساعة"], ["نسبة السعوديين", "96%"]]),
         ("p", "تتركز الشواغر في مركز الاتصال (8 وظائف) وتقنية المعلومات (4 وظائف). ويُنفذ التدريب على البوابة لأربعين موظفًا في نوفمبر.")],
        [("h1", "6. الأداء المالي"),
         ("table", [["البند", "الميزانية السنوية", "المصروف حتى سبتمبر", "نسبة التنفيذ"], ["الرواتب والبدلات", "26,400,000", "19,650,000", "74%"],
                    ["التشغيل والصيانة", "5,800,000", "3,830,000", "66%"], ["المشاريع الرأسمالية", "5,780,000", "3,450,000", "60%"],
                    ["التدريب", "520,000", "370,000", "71%"], ["الإجمالي", "38,500,000", "27,300,000", "71%"]]),
         ("p", "بلغ تنفيذ الميزانية 71% مقابل مستهدف 75% بنهاية الربع الثالث، ويعود الانخفاض أساسًا إلى تأخر مشروع نظام إدارة الوثائق.")],
        [("h1", "7. المشاريع"),
         ("table", [["المشروع", "الإنجاز الفعلي", "المخطط", "الحالة"], ["بوابة الخدمات الذاتية", "78%", "85%", "متأخر قليلًا"],
                    ["أتمتة المطابقة المالية", "92%", "90%", "في المسار"], ["نظام إدارة الوثائق", "34%", "60%", "متعثر"],
                    ["تحديث نظام الاتصال", "100%", "100%", "مكتمل"]]),
         ("p", "تعثر نظام إدارة الوثائق بسبب انسحاب المورد من مرحلة الترحيل، وتجري إدارة المشتريات تقييم بدائل.")],
        [("h1", "8. المخاطر"),
         ("table", [["الخطر", "الاحتمال", "الأثر", "الإجراء"], ["تأخر التكامل مع نظام الموارد البشرية", "مرتفع", "مرتفع", "تأجيل شهادة الراتب للمرحلة الثانية"],
                    ["ضغط الاستفسارات في أسبوع الإطلاق", "مرتفع", "متوسط", "تعزيز مؤقت للفريق"], ["استمرار تعثر نظام الوثائق", "متوسط", "مرتفع", "تقييم موردين بديلين"],
                    ["انخفاض نسبة الاستجابة للاستبيان", "مرتفع", "متوسط", "إعادة الاستبيان إلى نهاية المكالمة"]])],
        [("h1", "9. التوصيات"),
         ("ol", ["اعتماد أمر التغيير الخاص بمتطلبات الوصول قبل 22 أكتوبر.", "تعزيز مكتب خدمة المستفيدين بستة موظفين مؤقتًا في الأسبوع الأول من الإطلاق.",
                 "إعداد دليل إجابات موحد لأكثر عشرين سؤالًا تكرارًا ونشره في البوابة.", "إعادة الاستبيان إلى نهاية المكالمة لرفع نسبة الاستجابة.",
                 "رفع خطة بديلة لنظام إدارة الوثائق في الربع الرابع.", "معالجة أسباب تأخر طلبات صندوق السدرة مع إدارتي المالية والموارد البشرية."])],
        [("h1", "ملحق: تعريف المؤشرات والمنهجية"),
         ("table", [["المؤشر", "التعريف"], ["متوسط مدة المعالجة", "متوسط أيام العمل من استلام الطلب إلى إغلاقه، للطلبات المغلقة في الربع"],
                    ["رضا المستفيدين", "نسبة من قيّموا الخدمة بـ4 أو 5 من 5 في الاستبيان القصير"], ["الحل من أول تواصل", "نسبة الطلبات التي أُغلقت دون تواصل ثانٍ من المستفيد خلال 7 أيام"],
                    ["القنوات الرقمية", "البوابة والبريد ونموذج الموقع"]]),
         ("p", "تغيرت طريقة جمع الاستبيان في سبتمبر من سؤال في نهاية المكالمة إلى رسالة نصية بعد إغلاق الطلب، فانخفضت نسبة الاستجابة من 31% إلى 12%. "
               "ولذلك تُقرأ نتيجة رضا سبتمبر بحذر."),
         ("small", "مصدر البيانات: نظام إدارة الطلبات ونظام الاتصال ونظام الموارد البشرية والنظام المالي، حتى 30 سبتمبر 2026م.")],
    ]


def hr_faq_pages() -> list:
    rows = [("h2" if False else "h3", q) for q, _ in T.HR_FAQ]
    pages, cur = [], []
    for i, (q, a) in enumerate(T.HR_FAQ, 1):
        cur += [("h3", f"{i}. {q}"), ("p", a)]
        if len(cur) >= 14:
            pages.append(cur)
            cur = []
    if cur:
        pages.append(cur)
    pages[0].insert(0, ("kv", [("رقم الوثيقة", "HR-FAQ-03"), ("آخر تحديث", "1 أكتوبر 2026م"), ("الجهة المالكة", "إدارة الموارد البشرية")]))
    return pages


def identity_pages() -> list:
    swatches = "".join(f'<div style="display:inline-block;margin-left:4mm;text-align:center;font-size:9pt"><div style="width:22mm;height:12mm;border-radius:2mm;background:{c}"></div>{c}</div>'
                       for c in ("#0F766E", "#F1E7D3", "#C8963E", "#1E3A5F"))
    page = list(T.IDENTITY["pages"][0])
    page.insert(2, ("raw", f'<div style="margin:2mm 0 4mm">{swatches}</div>'))
    return [page] + T.IDENTITY["pages"][1:]


# ================================================================ main
def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-pdf", action="store_true")
    args = ap.parse_args()
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "circulars").mkdir(exist_ok=True)
    KEYS.parent.mkdir(parents=True, exist_ok=True)
    rng = random.Random(SEED)

    print("XLSX")
    records, dups = gen_inquiries(rng)
    write_inquiries(records)
    budget_errors = write_budget()
    write_kpis(random.Random(SEED + 1))
    write_glossary()
    write_kit()

    print("DOCX")
    write_docx(OUT / "accommodation_request.docx", T.ACCOMMODATION["title"], T.ACCOMMODATION["pages"])
    write_docx(OUT / "report_excerpt_en.docx", T.EN_EXCERPT_TITLE, [[("p", p) for p in T.EN_EXCERPT]], header="International Digital Services Review 2026")
    for num, day, issuer, cls, subject, body in T.CIRCULARS:
        name = f"{num.replace('ت-', 'C-')}_{day.split()[0]}oct.docx"
        write_docx(OUT / "circulars" / name, f"تعميم رقم {num}: {subject}",
                   [[("kv", [("رقم التعميم", num), ("التاريخ", day + "م"), ("الجهة المصدرة", issuer), ("التصنيف", cls)]), ("p", body),
                     ("small", "يُعمل بهذا التعميم من تاريخه.")]])
    write_docx(OUT / "circulars" / "bulletin_template.docx", "النشرة الأسبوعية للتعاميم · قالب",
               [[("p", "الأسبوع: [من – إلى]"), ("h2", "مطلوب منك هذا الأسبوع"),
                 ("table", [["التعميم", "المطلوب", "الفئة المعنية", "الموعد النهائي"], ["[رقم وعنوان]", "[سطر واحد]", "[من]", "[تاريخ]"]]),
                 ("h2", "للعلم"), ("ul", ["[رقم التعميم: ملخص من سطرين]"]),
                 ("h2", "يحتاج تأكيدًا"), ("ul", ["[ما كان غامضًا في التعاميم]"]),
                 ("small", "أعدت هذه النشرة بمساعدة الذكاء الاصطناعي وراجعها [الاسم والصفة] قبل الإرسال.")]])

    print("TXT")
    lines = [T.TRANSCRIPT_TITLE, "", "المتحدثون:"] + [f"- {k}: {v}" for k, v in T.TRANSCRIPT_SPEAKERS.items()] + [""]
    for ts, who, text in T.TRANSCRIPT:
        if ts == "38:50":
            lines += [T.TRANSCRIPT_GAP, ""]
        lines.append(f"[{ts}] {who}: {text}")
    (OUT / "meeting_transcript_sept_review.txt").write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")
    inbox = []
    for i, (frm, org, when, subj, body) in enumerate(T.INBOX, 1):
        inbox += [f"===== الرسالة {i} =====", f"من: {frm}", f"الجهة: {org}", f"التاريخ: {when}", f"الموضوع: {subj}", "", body, ""]
    (OUT / "inbox_week43.txt").write_text("\n".join(inbox), encoding="utf-8", newline="\n")
    print("  files/meeting_transcript_sept_review.txt\n  files/inbox_week43.txt")

    if not args.no_pdf:
        chrome = find_chrome()
        if not chrome:
            sys.exit("Chrome not found: set CHROME=/path/to/chrome or run with --no-pdf")
        print("PDF")
        write_pdf(OUT / "minutes_portal_committee.pdf", doc_html(T.MINUTES_PORTAL["title"], T.MINUTES_PORTAL["pages"]), chrome)
        write_pdf(OUT / "quarterly_report_q3_2026.pdf", doc_html("تقرير الأداء الربعي · الربع الثالث 2026", report_pages()), chrome)
        write_pdf(OUT / "policy_flexible_work.pdf", doc_html(T.POLICY_FLEX["title"], T.POLICY_FLEX["pages"]), chrome)
        write_pdf(OUT / "procurement_guide.pdf", doc_html(T.POLICY_PROCUREMENT["title"], T.POLICY_PROCUREMENT["pages"]), chrome)
        write_pdf(OUT / "hr_faq.pdf", doc_html("الأسئلة الشائعة للموظفين · الموارد البشرية", hr_faq_pages()), chrome)
        write_pdf(OUT / "visual_identity_guide.pdf", doc_html(T.IDENTITY["title"], identity_pages()), chrome)
        write_pdf(OUT / "lecture_05_service_quality.pdf", doc_html(T.LECTURE["title"], T.LECTURE["pages"], header="جامعة افتراضية · كلية الإدارة"), chrome)
        for key, name in (("a", "quote_a_alofuq.pdf"), ("b", "quote_b_naseej.pdf"), ("c", "quote_c_rawasi.pdf")):
            q = T.QUOTES[key]
            write_pdf(OUT / name, doc_html(f"عرض {q['company']}", q["pages"], header=q["company"]), chrome)

    # ------------------------------------------------------------ answer keys
    counts = Counter(r["topic"] for r in records[: len(records)])
    unique = len({r["id"] for r in records})
    c3 = [520000, round(520000 * 1.07), round(520000 * 1.07 ** 2)]
    k = ["# Answer keys (generated by labs/tools/prepare_files.py)", "",
         "Instructor material. Regenerating the files regenerates this page; do not edit by hand.", "",
         "## beneficiary_inquiries_q3.xlsx", "",
         f"- Rows: {len(records)} ({unique} unique request IDs + {len(dups)} duplicated rows: "
         + ", ".join(sorted(r['id'] for r in dups)) + ").",
         "- Channel names are inconsistent on purpose: «البوابة» / «البوابه», «البريد الإلكتروني» / «بريد إلكتروني».",
         "- Topic counts used to generate the texts (an AI classification will be close, not identical; duplicates counted once):", ""]
    k += [f"  - {t}: {c}" for t, (c, _) in T.INQUIRY_TOPICS.items()]
    k += ["", "## quarterly_report_q3_2026.pdf", "",
          "- **Planted inconsistency:** the executive summary (page 2) says satisfaction rose to **84%**; the KPI table (page 3) says **81%**. "
          "81% is the measured figure. The September survey switched to SMS and response fell to 12% (appendix, page 12).",
          "- Pages: 12. Q3 requests 48,600 = 15,900 + 16,300 + 16,400.", "",
          "## Supplier quotations (Lab 2.2)", "",
          "| | Al-Ofuq (A) | Naseej (B) | Rawasi (C) |", "|---|---|---|---|",
          "| Headline (p. 1) | 1,450,000 incl. VAT | 1,180,000 | «from 520,000 / year» |",
          "| Real 3-year cost | 1,330,000 + VAT = 1,529,500 incl. 15% VAT, but p. 1 says 1,450,000: **mismatch to query** (p. 4 vs p. 1) | 1,180,000 + Arabic chatbot 260,000 + integration (unpriced) + VAT (p. 3, p. 5) | "
          + f"{c3[0]:,} + {c3[1]:,} + {c3[2]:,} = {sum(c3):,} incl. VAT (7% escalation, p. 3) |",
          "| Hosting | Riyadh (p. 2) | **Frankfurt** (p. 5): fails the in-Kingdom hosting rule in the procurement guide | Dammam (p. 2) |",
          "| Arabic chatbot | included (p. 2) | **optional +260,000** (p. 5) | included (p. 2) |",
          "| Availability SLA | 99.5% (p. 3) | **99%** (p. 2): below the 99.5% minimum | 99.9% (p. 2) |",
          "| Integration with portal + HR | included (p. 2) | **not included** (p. 5) | included (p. 2) |",
          "| Payment | 40/30/30 (p. 4) | **50% advance** (p. 4): above the 10% advance cap | **100% yearly in advance** (p. 3): above the cap |",
          "| Warranty | 12 months (p. 3) | 6 months (p. 4) | subscription (p. 3) |", "",
          "Procurement route: every option exceeds 1,000,000 riyals over three years, so **public competition** applies (procurement guide, section 2).", "",
          "## training_budget_2026.xlsx (Lab 2.3)", ""]
    k += [f"{i}. {e}" for i, e in enumerate(budget_errors, 1)]
    k += ["", "## minutes_portal_committee.pdf (Lab 1.3)", "",
          "- Decisions (3): postpone the salary-certificate service to phase 2; extend user testing to 25 Oct and redesign the attachment step; "
          "escalate the 186,000 change order to the Director General (above the committee's 150,000 limit).",
          "- Tasks: IT, integration test by 28 Oct; Beneficiary Services, 25 FAQ answers (date: «before launch», not a date); Finance, impact memo "
          "within three working days; Communications, campaign plan (**no date**); HR, train 40 staff, week of 2 Nov.",
          "- Traps: the soft-launch idea is **not** a decision; the chatbot is **deferred**, not decided; the meeting-room AV issue is noise.",
          "- Needs the Director General: the change order, before 22 Oct (after that it cannot ship before the 15 Nov launch).",
          "- Accessibility test **not scheduled**: a high risk with no owner.", "",
          "## meeting_transcript_sept_review.txt (Lab 2.2)", "",
          "- Actions: unified answer guide (Majed + Lama, 31 Oct); link it to the portal FAQ (Bandar, with Majed, no date); two training "
          "batches, schedule to supervisors on Sunday (Sara); survey back at end of call (Bandar, before end of month); AI guidelines (Lama, 15 Nov).",
          "- Open: temporary extra staff for launch week (Noura to come back next week). Next meeting Thursday 22 Oct.",
          "- September satisfaction: 80% approved (82% was before removing duplicate responses).",
          "- 28 minutes were not transcribed (off-topic); say so in the minutes.", "",
          "## circulars/ (Lab 4.1)", "",
          "- Six files: C-2026-118 appears twice (original + reminder: deduplicate); C-2026-121 is **restricted** and must not go to the AI "
          "application; C-2026-117 cancels an earlier circular.",
          "- Expected bulletin: action items 118 (update data by 31 Oct, all staff) and 120 (mandatory training 1–4 Nov, Beneficiary Services); "
          "for information 117 and 119; nothing from 121.", ""]
    KEYS.write_text("\n".join(k) + "\n", encoding="utf-8", newline="\n")
    print(f"KEYS\n  {KEYS.relative_to(LABS)}")


if __name__ == "__main__":
    main()
