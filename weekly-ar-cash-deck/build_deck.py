#!/usr/bin/env python3
"""
build_deck.py — builds the weekly "AR, Cash & Payables Position" PowerPoint from inputs/weekly_inputs.xlsx.

Usage:
  python3 build_deck.py --inputs inputs/weekly_inputs.xlsx --out output/JESA_AR_Cash_Position_2026-10-21.pptx

Everything on the slides comes from the input workbook (see extract_inputs.py / README.md).
"""
import argparse, datetime as dt, copy
import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_LABEL_POSITION
from pptx.oxml.ns import qn
from lxml import etree

# ----------------------------------------------------------------------------- design tokens
NAVY = RGBColor(0x1B, 0x2F, 0x4B); NAVY2 = RGBColor(0x0F, 0x3B, 0x63); BLUE = RGBColor(0x2E, 0x6D, 0xA4)
LBLUE = RGBColor(0xC9, 0xDD, 0xEE); GOLD = RGBColor(0xC8, 0x97, 0x2B); TEAL = RGBColor(0x3A, 0x9C, 0x8C)
RED = RGBColor(0xB5, 0x47, 0x3A); GREEN = RGBColor(0x2E, 0x7D, 0x4F); GREY = RGBColor(0x7A, 0x87, 0x94)
DARK = RGBColor(0x1F, 0x2A, 0x37); BG = RGBColor(0xF4, 0xF8, 0xFB); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xE9, 0xF0, 0xF6); LIGHT2 = RGBColor(0xF7, 0xFA, 0xFC); SLATE = RGBColor(0x6B, 0x7B, 0x8C)
ORANGE = RGBColor(0xE0, 0x8A, 0x2E); PURPLE = RGBColor(0x5B, 0x4B, 0x8A)
FONT = "Arial"
SW, SH = 13.333, 7.5

def rgb_hex(c): return "%02X%02X%02X" % (c[0], c[1], c[2])

# ----------------------------------------------------------------------------- inputs
def read_inputs(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    data = {}
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        # skip note row if present (first row with a single text cell and no header keywords)
        h = 0
        if rows and rows[0] and isinstance(rows[0][0], str) and all(v is None for v in rows[0][1:]) and rows[0][0] not in ("key",):
            h = 1
        header = [str(c) if c is not None else "" for c in rows[h]]
        recs = []
        for r in rows[h + 1:]:
            if all(v is None for v in r): continue
            recs.append({header[i]: r[i] for i in range(len(header))})
        data[ws.title] = recs
    meta = {r["key"]: r["value"] for r in data["META"]}
    return data, meta

def f0(v):
    if v is None: return "–"
    v = float(v)
    if abs(v) < 0.5: return "0"
    return f"{v:,.0f}".replace("-", "−")

def f1(v):
    if v is None: return "–"
    return f"{float(v):,.1f}".replace("-", "−")

def fd(v, decimals=0):
    """signed delta."""
    if v is None: return "–"
    v = float(v)
    if abs(v) < (0.5 if decimals == 0 else 0.05): return "0"
    s = f"{abs(v):,.{decimals}f}"
    return ("+" if v > 0 else "−") + s

def fpct(v):
    if v is None: return "–"
    return f"{v*100:+.1f}%".replace("-", "−")

def delta_color(v, good_when_negative=True, neutral_below=0.5):
    if v is None or abs(float(v)) < neutral_below: return GREY
    neg = float(v) < 0
    return GREEN if (neg == good_when_negative) else RED

def datef(s, fmt="%d %b %Y"):
    if s is None: return ""
    if isinstance(s, (dt.date, dt.datetime)): return s.strftime(fmt)
    try: return dt.date.fromisoformat(str(s)[:10]).strftime(fmt)
    except Exception: return str(s)

# ----------------------------------------------------------------------------- primitives
def rect(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None: s.line.fill.background()
    else: s.line.color.rgb = line; s.line.width = Pt(0.75)
    s.shadow.inherit = False
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    return s

def tb(slide, x, y, w, h, text, size=12, bold=False, color=DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
       italic=False, font=FONT, margin=0.04, line_spacing=None):
    """text: str, or list of paragraphs; each paragraph str or list of runs (text, {size,bold,color,italic})."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(margin); tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = anchor
    paras = text if isinstance(text, list) else [text]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        if line_spacing: para.line_spacing = line_spacing
        runs = p if isinstance(p, list) else [(p, {})]
        for rt, st in runs:
            r = para.add_run(); r.text = rt
            r.font.name = st.get("font", font); r.font.size = Pt(st.get("size", size)); r.font.bold = st.get("bold", bold)
            r.font.italic = st.get("italic", italic); r.font.color.rgb = st.get("color", color)
    return box

def card(slide, x, y, w, h, label, value, sub=None, fill=NAVY, fg=WHITE, value_size=26, label_size=9.5, sub_color=None, sub_size=8.5):
    rect(slide, x, y, w, h, fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
    tb(slide, x + 0.1, y + 0.07, w - 0.2, 0.28, label.upper(), size=label_size, bold=True, color=fg)
    vs = min(value_size, int((h - 0.62) * 72 / 1.25)) if sub else value_size
    tb(slide, x + 0.1, y + 0.3, w - 0.2, vs * 1.3 / 72 + 0.05, value, size=vs, bold=True, color=fg)
    if sub:
        tb(slide, x + 0.1, y + h - 0.3, w - 0.2, 0.28, sub, size=sub_size, color=sub_color or fg)

def kpi_card(slide, x, y, w, h, label, value, delta_wow=None, delta_ms=None, fill=WHITE, border=LBLUE, good_when_negative=True, value_size=22, ms_label="vs baseline"):
    rect(slide, x, y, w, h, fill, line=border, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.1)
    rect(slide, x, y + 0.12, 0.07, h - 0.24, GOLD)
    tb(slide, x + 0.15, y + 0.07, w - 0.2, 0.28, label.upper(), size=9, bold=True, color=SLATE)
    if value is not None:
        tb(slide, x + 0.15, y + 0.3, w - 0.2, 0.5, f0(value), size=value_size, bold=True, color=NAVY)
    if delta_wow is not None:
        yw = (y + h - 0.56) if (delta_ms is not None or h >= 1.2) else (y + h - 0.32)
        tb(slide, x + 0.15, yw, w - 0.2, 0.26, [[("WoW  ", {"color": SLATE, "size": 8.5}), (fd(delta_wow), {"color": delta_color(delta_wow, good_when_negative), "bold": True, "size": 10})]], size=9)
    if delta_ms is not None:
        tb(slide, x + 0.15, y + h - 0.32, w - 0.2, 0.26, [[(ms_label + "  ", {"color": SLATE, "size": 8.5}), (fd(delta_ms), {"color": delta_color(delta_ms, good_when_negative), "bold": True, "size": 10})]], size=9)

def circle_num(slide, x, y, d, n, fill=NAVY2):
    s = rect(slide, x, y, d, d, fill, shape=MSO_SHAPE.OVAL)
    tf = s.text_frame; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER; r = p.add_run(); r.text = str(n)
    r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

# ----------------------------------------------------------------------------- table
def table(slide, x, y, w, header, rows, col_w=None, font_size=9, header_size=9, row_h=0.26, header_h=0.34,
          align=None, bold_rows=(), fills=None, colors=None, zebra=True, header_fill=NAVY, total_last=False, first_col_bold=False):
    """rows: list of lists of str. colors: dict (r,c)->RGB. fills: dict r->RGB."""
    nr, nc = len(rows) + 1, len(header)
    h = header_h + row_h * len(rows)
    gs = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(h))
    t = gs.table
    tblPr = gs._element.graphic.graphicData.tbl.tblPr
    tblPr.set("firstRow", "0"); tblPr.set("bandRow", "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is not None: sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # No style, no grid
    if col_w:
        tot = sum(col_w)
        for j, cw in enumerate(col_w): t.columns[j].width = Inches(w * cw / tot)
    t.rows[0].height = Inches(header_h)
    for i in range(1, nr): t.rows[i].height = Inches(row_h)
    def setcell(cell, text, size, bold, color, fill, al):
        cell.fill.solid(); cell.fill.fore_color.rgb = fill
        cell.margin_left = cell.margin_right = Inches(0.05); cell.margin_top = cell.margin_bottom = Inches(0.015)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = al; p.font.size = Pt(size)
        r = p.add_run(); r.text = "" if text is None else str(text)
        r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; r.font.name = FONT
    for j, hname in enumerate(header):
        setcell(t.cell(0, j), hname, header_size, True, WHITE, header_fill, PP_ALIGN.CENTER if j else PP_ALIGN.LEFT)
    for i, row in enumerate(rows, 1):
        is_total = total_last and i == nr - 1
        base_fill = (LIGHT if (zebra and i % 2 == 0) else WHITE)
        if fills and (i - 1) in fills: base_fill = fills[i - 1]
        if is_total: base_fill = LBLUE
        for j, v in enumerate(row):
            al = (align[j] if align else (PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT))
            col = DARK
            if colors and (i - 1, j) in colors: col = colors[(i - 1, j)]
            bold = is_total or ((i - 1) in bold_rows) or (first_col_bold and j == 0)
            setcell(t.cell(i, j), v, font_size, bold, col, base_fill, al)
    return gs

# ----------------------------------------------------------------------------- charts
def _style_axes(ch, cat_size=9, val_visible=False, val_size=8, gridlines=False):
    ca = ch.category_axis; ca.tick_labels.font.size = Pt(cat_size); ca.tick_labels.font.name = FONT; ca.tick_labels.font.color.rgb = DARK
    ca.format.line.color.rgb = RGBColor(0xBF, 0xC8, 0xD2); ca.has_major_gridlines = False
    va = ch.value_axis; va.visible = val_visible; va.has_major_gridlines = gridlines
    if gridlines: va.major_gridlines.format.line.color.rgb = RGBColor(0xE1, 0xE7, 0xED)
    va.tick_labels.font.size = Pt(val_size); va.tick_labels.font.name = FONT
    va.format.line.fill.background()

def bar_chart(slide, x, y, w, h, cats, series, colors, stacked=False, horizontal=False, labels=True, legend=True,
              number_format='#,##0', label_size=8, cat_size=9, gap=60, legend_pos=XL_LEGEND_POSITION.TOP, title=None, val_axis=False, label_pos=None, overlap=None, label_colors=None):
    cd = CategoryChartData(); cd.categories = cats
    nd = 1 if "0.0" in number_format else 0
    series = [(name, [round(float(v) if v is not None else 0.0, nd) for v in vals]) for name, vals in series]
    if ";" not in number_format: number_format = number_format + ";−" + number_format + ';""'
    elif number_format.endswith(";0"): number_format = number_format[:-2] + ';""'
    for name, vals in series: cd.add_series(name, vals)
    if horizontal: ctype = XL_CHART_TYPE.BAR_STACKED if stacked else XL_CHART_TYPE.BAR_CLUSTERED
    else: ctype = XL_CHART_TYPE.COLUMN_STACKED if stacked else XL_CHART_TYPE.COLUMN_CLUSTERED
    gf = slide.shapes.add_chart(ctype, Inches(x), Inches(y), Inches(w), Inches(h), cd); ch = gf.chart
    ch.font.name = FONT; ch.font.size = Pt(9)
    ch.has_title = bool(title)
    if title:
        ch.chart_title.text_frame.text = title
        p = ch.chart_title.text_frame.paragraphs[0]; p.runs[0].font.size = Pt(11); p.runs[0].font.bold = True; p.runs[0].font.color.rgb = NAVY
    ch.has_legend = legend and len(series) > 1
    if ch.has_legend:
        ch.legend.position = legend_pos; ch.legend.include_in_layout = False; ch.legend.font.size = Pt(8.5); ch.legend.font.name = FONT
    plot = ch.plots[0]; plot.gap_width = gap
    if stacked: plot.overlap = 100
    elif overlap is not None: plot.overlap = overlap
    for i, s in enumerate(plot.series):
        s.format.fill.solid(); s.format.fill.fore_color.rgb = colors[i % len(colors)]
        s.format.line.fill.background()
    if labels:
        plot.has_data_labels = True; dl = plot.data_labels
        dl.number_format = number_format; dl.number_format_is_linked = False
        dl.font.size = Pt(label_size); dl.font.name = FONT; dl.font.color.rgb = DARK
        if label_pos is not None: dl.position = label_pos
        if label_colors:
            for i, ser in enumerate(plot.series):
                sdl = ser.data_labels; sdl.number_format = number_format; sdl.number_format_is_linked = False
                sdl.font.size = Pt(label_size); sdl.font.name = FONT; sdl.font.color.rgb = label_colors[i % len(label_colors)]; sdl.show_value = True
                if label_pos is not None: sdl.position = label_pos
    _style_axes(ch, cat_size=cat_size, val_visible=val_axis)
    if any((float(v) if v is not None else 0.0) < 0 for _, vals in series for v in vals):
        ch.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    if horizontal: ch.category_axis.reverse_order = True
    return ch

def hide_small_labels(ch, series_values, threshold):
    """delete the data label of points below threshold (keeps the series-level labels for the others)."""
    for si, vals in enumerate(series_values):
        ser = ch.plots[0].series[si]
        for i, v in enumerate(vals):
            if abs(float(v or 0)) < threshold:
                dLbl = ser.points[i].data_label._get_or_add_dLbl()
                for child in list(dLbl):
                    if child.tag != qn("c:idx"): dLbl.remove(child)
                etree.SubElement(dLbl, qn("c:delete")).set("val", "1")

def color_points(ch, series_idx, colors):
    s = ch.plots[0].series[series_idx]
    for i, c in enumerate(colors):
        pt = s.points[i]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb = c; pt.format.line.fill.background()

def doughnut(slide, x, y, w, h, cats, vals, colors, labels=False, legend=False, hole=55, label_size=9, number_format='#,##0'):
    cd = CategoryChartData(); cd.categories = cats; cd.add_series("s", [float(v) for v in vals])
    gf = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(x), Inches(y), Inches(w), Inches(h), cd); ch = gf.chart
    ch.font.name = FONT; ch.has_title = False
    ch.has_legend = legend
    if legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM; ch.legend.include_in_layout = False; ch.legend.font.size = Pt(8.5)
    plot = ch.plots[0]
    color_points(ch, 0, colors)
    if labels:
        tot = sum(float(v) for v in vals) or 1.0
        plot.has_data_labels = True
        for i, (c, v) in enumerate(zip(cats, vals)):
            pt = plot.series[0].points[i]; pt.data_label.has_text_frame = True
            pt.data_label.text_frame.text = f"{c}, {float(v)/tot*100:.0f}%\n{f0(v)}"
            for para in pt.data_label.text_frame.paragraphs:
                for r in para.runs: r.font.size = Pt(label_size); r.font.bold = True; r.font.color.rgb = DARK; r.font.name = FONT
            pt.data_label.position = XL_LABEL_POSITION.OUTSIDE_END
    # hole size
    hs = plot._element.find(qn("c:holeSize"))
    if hs is None:
        hs = etree.SubElement(plot._element, qn("c:holeSize"))
    hs.set("val", str(hole))
    return ch

def waterfall(slide, x, y, w, h, cats, values, kinds, colors=None, label_size=8.5, cat_size=8.5, number_format='#,##0'):
    """kinds: 'total' or 'delta'. Builds a stacked column with invisible base."""
    base, vis, cols = [], [], []
    run = 0.0
    for v, k in zip(values, kinds):
        if k == "total":
            base.append(0.0); vis.append(abs(v)); run = v; cols.append(NAVY)
        else:
            new = run + v
            base.append(min(run, new)); vis.append(abs(v)); run = new
            cols.append(GREEN if v < 0 else RED)
    cd = CategoryChartData(); cd.categories = cats
    cd.add_series("base", base); cd.add_series("value", vis)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(x), Inches(y), Inches(w), Inches(h), cd); ch = gf.chart
    ch.font.name = FONT; ch.has_legend = False
    plot = ch.plots[0]; plot.gap_width = 40; plot.overlap = 100
    b = plot.series[0]; b.format.fill.background(); b.format.line.fill.background()
    s = plot.series[1]; s.format.line.fill.background()
    for i, c in enumerate(cols):
        pt = s.points[i]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb = c
        pt.data_label.has_text_frame = True
        txt = (f0(values[i]) if kinds[i] == "total" else fd(values[i]))
        pt.data_label.text_frame.text = txt
        r = pt.data_label.text_frame.paragraphs[0].runs[0]; r.font.size = Pt(label_size); r.font.bold = True; r.font.name = FONT
        r.font.color.rgb = WHITE
        pt.data_label.position = XL_LABEL_POSITION.INSIDE_END if kinds[i] == "total" else XL_LABEL_POSITION.CENTER
    _style_axes(ch, cat_size=cat_size)
    return ch

# ----------------------------------------------------------------------------- slide chrome
class Deck:
    def __init__(self, meta):
        self.prs = Presentation(); self.prs.slide_width = Inches(SW); self.prs.slide_height = Inches(SH)
        self.blank = self.prs.slide_layouts[6]; self.n = 0; self.meta = meta

    def slide(self, title=None, subtitle=None, dark=False):
        s = self.prs.slides.add_slide(self.blank); self.n += 1
        bg = s.background; bg.fill.solid(); bg.fill.fore_color.rgb = (NAVY if dark else BG)
        rect(s, 0, 0, 1.7, 0.1, GOLD); rect(s, 1.7, 0, SW - 1.7, 0.1, NAVY2)
        # wordmark (replace with logo image in the master if available)
        tb(s, SW - 2.3, 0.18, 1.9, 0.5, "JESA", size=26, bold=True, color=(WHITE if dark else RGBColor(0x1E, 0x3F, 0x8F)), align=PP_ALIGN.RIGHT)
        if title:
            tb(s, 0.5, 0.22, 9.8, 0.55, title, size=26, bold=True, color=(WHITE if dark else NAVY))
        if subtitle:
            tb(s, 0.52, 0.78, 10.5, 0.35, subtitle, size=12.5, bold=True, color=TEAL)
        tb(s, 4.5, SH - 0.3, 4.3, 0.22, "JESA - INTERNAL USE ONLY", size=7.5, color=GREY, align=PP_ALIGN.CENTER)
        tb(s, SW - 1.3, SH - 0.3, 0.9, 0.22, f"{self.n}", size=8, color=GREY, align=PP_ALIGN.RIGHT)
        return s

    def save(self, path): self.prs.save(path)

# ----------------------------------------------------------------------------- data access helpers
def by_metric(data, sheet, key="metric"):
    return {r[key]: r for r in data[sheet]}

BU_SHORT = {"Water & Environment": "Water & Env.", "Fertilizers & Chemicals": "Fert. & Chem.", "SP2M Ad Hoc Program": "SP2M", "AM": "AM",
            "Energy Operations": "Energy Ops", "Global Mining": "Global Mining", "Building & Infrastructure": "B&I", "JESA Technologies": "JESA Tech",
            "Advisory": "Advisory", "Building & Capabilities": "B&C", "Power to X": "Power to X", "Other": "Other"}
BU_ORDER = ["Water & Environment", "Fertilizers & Chemicals", "SP2M Ad Hoc Program", "AM", "Energy Operations", "Global Mining",
            "Building & Infrastructure", "JESA Technologies", "Advisory", "Building & Capabilities", "Power to X"]
def short(bu): return BU_SHORT.get(bu, bu)

def short_action(a):
    """map the free-text 'Action' column of the LC situation to a short label."""
    l = (str(a) if a else "").strip().lower()
    if not l or l == "na": return "No action recorded"
    rules = [("24-month", "24-month rescheduling to finalise"), ("split amount", "Amount to be split by delivery"), ("pushed for 2027", "Remaining payment pushed to 2027"),
             ("finance to check", "Finance to check with bank"), ("lc to close", "LC to close"), ("lc closed", "LC closed"), ("to pay by 2026", "To pay by 2026"),
             ("lc copy", "Awaiting LC copy from buyer"), ("draft lc", "Awaiting draft LC (procurement)"), ("official response", "Awaiting JESA response (suspension)"),
             ("client agreement", "Client agreement on two-lot schedule"), ("opening committee", "LC opening committee / client approval"), ("new schedule", "Approval of new schedule"),
             ("signed wo", "Signed WO from client needed"), ("not open", "LC not opened")]
    for k, v in rules:
        if k in l: return v
    return (str(a)[:30] + "…") if len(str(a)) > 31 else str(a)

# ============================================================================= SLIDES
def build(data, meta, out):
    D = Deck(meta)
    kpi = by_metric(data, "AR_KPI")
    def K(name):
        for k, v in kpi.items():
            if k.lower().startswith(name.lower()): return v
        raise KeyError(name)
    adv, unb, unp, tot = K("Advance"), K("Unbilled"), K("Unpaid"), K("Total AR")
    unrec, newar = K("Unrecognized"), K("Total NEW AR")
    ap_dp, ap_pk, ap_44, ap_lc, ap_tot = K("AP Downpayment"), K("AP Parked"), K("AP Invoices 4411"), K("LC"), K("Total Payables")
    cur_d, prev_d, ms_d = datef(meta["ar_current_date"]), datef(meta["ar_prev_date"]), datef(meta["ar_month_start_date"])
    cur_s, prev_s, ms_s = datef(meta["ar_current_date"], "%d-%b"), datef(meta["ar_prev_date"], "%d-%b"), datef(meta["ar_month_start_date"], "%d-%b")
    ms_label = "vs month start"
    asof = datef(meta["as_of_date"])
    splits = {r["metric"]: r["value"] for r in data["AR_SPLITS"]}
    dso = {r["bu"]: r for r in data.get("DSO", [])}
    def dso_txt(b):
        r = dso.get(b)
        if not r or r.get("dso_current") in (None, ""): return None
        cur_ = int(float(r["dso_current"])); pr = r.get("dso_prior_month")
        return f"{cur_} days" + (f" (prior {int(float(pr))})" if pr not in (None, "") else "")
    arbu = {}
    for r in data["AR_BU"]: arbu.setdefault(r["bu"], {})[r["component"]] = r
    def dwow(r): return round(r["current"] or 0) - round(r["prev_week"] or 0)
    def dms(r): return round(r["current"] or 0) - round(r["month_start"] or 0)

    # ------------------------------------------------------------------ 1. Title
    s = D.slide(dark=False)
    tb(s, 0.9, 1.5, 9, 0.4, "JESA  ·  ACCOUNTS RECEIVABLE, CASH & PAYABLES", size=14, bold=True, color=GOLD)
    tb(s, 0.9, 1.95, 11.5, 1.1, "AR, Cash & Payables Position", size=44, bold=True, color=NAVY)
    rect(s, 0.9, 3.25, 1.9, 0.04, GOLD)
    tb(s, 0.9, 3.4, 9, 0.5, "Weekly Executive Review", size=24, color=NAVY)
    tb(s, 0.9, 3.95, 11.5, 0.4, f"Deck as of {asof}   ·   AR snapshot {cur_d} (vs {prev_d} and {ms_d} baseline)   ·   Cash to {meta['collections_to']}   ·   Figures in MMAD", size=12.5, color=SLATE)
    cards = [("Gross exposure", newar, RED), ("Unrecognized revenue", unrec, BLUE), ("Total AR", tot, SLATE), ("Unbilled", unb, TEAL), ("Unpaid", unp, GOLD), ("Advance", adv, NAVY2)]
    cw = (SW - 0.6 - 0.12 * 5) / 6
    for i, (lab, r, col) in enumerate(cards):
        card(s, 0.3 + i * (cw + 0.12), 5.55, cw, 1.35, lab, f0(r["current"]), f"vs {prev_s}:  {fd(dwow(r))}", fill=col, value_size=24, sub_size=10)

    # ------------------------------------------------------------------ 1b. Agenda
    s = D.slide("Agenda", f"Weekly review – {asof}")
    items = [("1", "Key messages", "Headline position: AR, unrecognized revenue, collections, cash, payables and letters of credit"),
             ("2", "Accounts receivable", "Executive summary, movement vs last week and vs month start, position by business unit, unbilled and unpaid status, non-OCP clients, unbilled under JESA responsibility"),
             ("3", "Unrecognized revenue & suppliers", "Executed scope not yet booked as AR, by business unit and status; supplier PO book"),
             ("4", "Payments received", "Invoices paid in the month and year to date, by business unit"),
             ("5", "Cash position – procurement", "Cash in versus cash out by month and by business unit; paid, approved and pending payments"),
             ("6", "Payables & letters of credit", "Supplier payables by business unit and due date; LC commitments and maturity schedule"),
             ("7", "Zoom by business unit", "Position, movement, projects and unrecognized revenue for each business unit"),
             ("8", "Backups & sources", "LC and payables detail; sources and definitions")]
    y0 = 1.3; step = (SH - 0.6 - y0) / len(items)
    for i, (n, t, d_) in enumerate(items):
        y = y0 + i * step
        circle_num(s, 0.6, y + 0.05, 0.42, n)
        tb(s, 1.3, y - 0.02, 4.0, 0.4, t, size=14, bold=True, color=NAVY)
        tb(s, 5.2, y + 0.02, 7.7, step - 0.1, d_, size=10.5, color=DARK)
        if i < len(items) - 1: rect(s, 1.3, y + step - 0.1, 11.6, 0.012, RGBColor(0xD5, 0xDE, 0xE6))

    # ------------------------------------------------------------------ 2. Key messages
    s = D.slide("Key Messages", "The AR, cash and payables position at a glance")
    msgs = data["KEY_MESSAGES"][:6]
    y0 = 1.35; step = (SH - 0.6 - y0) / max(len(msgs), 1)
    for i, m in enumerate(msgs):
        y = y0 + i * step
        circle_num(s, 0.55, y + 0.08, 0.48, m["n"])
        tb(s, 1.3, y, 11.5, 0.4, str(m["title"]), size=13 if len(msgs) > 5 else 13.5, bold=True, color=NAVY)
        tb(s, 1.3, y + 0.36, 11.5, step - 0.55, str(m["body"]), size=10.5 if len(msgs) > 5 else 11.5, color=DARK, line_spacing=1.05)
        if i < len(msgs) - 1: rect(s, 1.3, y + step - 0.08, 11.5, 0.012, RGBColor(0xD5, 0xDE, 0xE6))

    # ------------------------------------------------------------------ 3. AR executive summary
    pct_unb = unb["current"] / tot["current"] * 100
    s = D.slide("AR Executive Summary", f"Unbilled is {pct_unb:.0f}% of total AR; AR {fd(dwow(tot))} MMAD week on week and {fd(dms(tot))} vs the month-start baseline ({ms_s})")
    cw = (SW - 0.6 - 0.12 * 5) / 6
    for i, (lab, r, gwn) in enumerate([("Gross exposure", newar, True), ("Unrecognized revenue", unrec, True), ("Total AR", tot, True),
                                        ("Unbilled", unb, True), ("Unpaid", unp, True), ("Advance", adv, True)]):
        kpi_card(s, 0.3 + i * (cw + 0.12), 1.25, cw, 1.25, lab, r["current"], dwow(r), (dms(r) if r["month_start"] else None), good_when_negative=gwn, ms_label=ms_label)
    # donut
    tb(s, 0.4, 2.7, 5.5, 0.35, "AR by component", size=14, bold=True, color=NAVY)
    tb(s, 0.4, 3.02, 5.5, 0.3, "Unbilled dominates the receivable book (MMAD)", size=10, color=SLATE)
    doughnut(s, 0.2, 3.3, 3.0, 3.7, ["Unbilled", "Unpaid", "Advance"], [unb["current"], unp["current"], adv["current"]], [NAVY2, GREY, LBLUE])
    for i, (lab, v, c) in enumerate([("Unbilled", unb["current"], NAVY2), ("Unpaid", unp["current"], GREY), ("Advance", adv["current"], LBLUE)]):
        yy = 3.9 + i * 0.85
        rect(s, 3.35, yy + 0.08, 0.22, 0.22, c)
        tb(s, 3.65, yy, 1.6, 0.3, lab, size=11, bold=True, color=NAVY)
        tb(s, 3.65, yy + 0.3, 1.6, 0.3, f"{f0(v)}  ·  {v/tot['current']*100:.0f}% of AR", size=9.5, color=SLATE)
    # OCP callout
    ocp = splits.get("AR OCP", 0); nocp = splits.get("AR NON OCP", 0)
    rect(s, 5.3, 2.75, 7.7, 1.1, WHITE, line=LBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08); rect(s, 5.3, 2.75, 0.08, 1.1, NAVY2)
    tb(s, 5.55, 2.8, 2.2, 0.9, f"{ocp/(ocp+nocp)*100:.0f}%", size=36, bold=True, color=BLUE, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, 7.7, 2.85, 5.2, 0.95, [[("of AR is OCP-related", {"size": 15, "bold": True, "color": NAVY})], [(f"OCP {f0(ocp)}  ·  Non-OCP {f0(nocp)} ({nocp/(ocp+nocp)*100:.1f}%)", {"size": 11, "color": SLATE})],
                                  [((("DSO excl. advances, JESA Group: " + dso_txt("JESA Group") + "  ·  ") if dso_txt("JESA Group") else "") + "Receivables remain almost entirely concentrated with OCP", {"size": 10, "color": SLATE})]], anchor=MSO_ANCHOR.MIDDLE)
    # labor vs non-labor
    tb(s, 5.4, 4.0, 7, 0.35, "Labor vs non-labor split", size=14, bold=True, color=NAVY)
    nl_pct = splits.get("Total AR non-labor", 0) / tot["current"] * 100
    tb(s, 5.4, 4.32, 7, 0.3, f"{nl_pct:.0f}% of the book is non-labor (procurement scope)", size=10, color=SLATE)
    bar_chart(s, 5.3, 4.55, 7.7, 2.55, ["Total AR", "Unbilled", "Unpaid"],
              [("Labor", [splits.get("Total AR labor"), splits.get("Unbilled labor"), splits.get("Unpaid labor")]),
               ("Non-labor", [splits.get("Total AR non-labor"), splits.get("Unbilled non-labor"), splits.get("Unpaid non-labor")])],
              [NAVY2, LBLUE], stacked=True, gap=80, label_size=9, label_colors=[WHITE, DARK])

    # ------------------------------------------------------------------ 4. AR movement (comparatives)
    s = D.slide("AR Movement – week on week and vs baseline",
                f"This week {cur_s} vs last week {prev_s} and vs the month-start baseline ({ms_s} AR position)")
    hdr = ["Component", f"Month start\n{ms_s}", f"Last week\n{prev_s}", f"This week\n{cur_s}", "Δ WoW", "Δ% WoW", "Δ vs\nmonth start", "Δ% vs\nmonth start"]
    rows, colors = [], {}
    comp_rows = [("Advance", adv), ("Unbilled", unb), ("Unpaid", unp), ("Total AR", tot), ("Unrecognized revenue", unrec), ("Gross exposure", newar)]
    for i, (lab, r) in enumerate(comp_rows):
        ms = r["month_start"] or None; pw = r["prev_week"]; cu = r["current"]
        d1 = cu - pw; d2 = (cu - ms) if ms else None
        rows.append([lab, f0(ms) if ms else "n/a", f0(pw), f0(cu), fd(d1), fpct(d1 / pw) if pw else "–", fd(d2) if d2 is not None else "n/a", fpct(d2 / ms) if ms else "n/a"])
        colors[(i, 4)] = delta_color(d1); colors[(i, 5)] = delta_color(d1); colors[(i, 6)] = delta_color(d2); colors[(i, 7)] = delta_color(d2)
    table(s, 0.4, 1.3, 7.6, hdr, rows, col_w=[2.6, 1.1, 1.1, 1.1, 0.9, 0.9, 1.0, 0.95], font_size=9.5, bold_rows=(3,), row_h=0.3)
    # narrative: biggest BU drivers
    def drivers(comp, key):
        items = sorted(((bu, key(arbu[bu][comp])) for bu in arbu if comp in arbu[bu]), key=lambda t: t[1])
        neg = [f"{short(b)} {fd(v)}" for b, v in items if v <= -1][:4]; pos = [f"{short(b)} {fd(v)}" for b, v in reversed(items) if v >= 1][:3]
        return neg, pos
    lines = []
    for comp, r in (("Unbilled", unb), ("Unpaid", unp), ("Total AR", tot)):
        n, p = drivers(comp, dwow)
        lines.append([(f"{comp} {fd(dwow(r))} WoW: ", {"bold": True, "color": NAVY}), ("decreases " + ", ".join(n) if n else "no decrease", {}), ("; increases " + ", ".join(p) if p else "", {})])
    for comp, r in (("Unbilled", unb), ("Unpaid", unp), ("Total AR", tot)):
        n, p = drivers(comp, dms)
        lines.append([(f"{comp} {fd(dms(r))} vs month start: ", {"bold": True, "color": NAVY}), ("decreases " + ", ".join(n) if n else "no decrease", {}), ("; increases " + ", ".join(p) if p else "", {})])
    tb(s, 0.4, 3.75, 7.6, 0.3, "What moved", size=13, bold=True, color=NAVY)
    tb(s, 0.4, 4.05, 7.6, 2.0, lines, size=9.5, color=DARK, line_spacing=1.15)
    # aging table
    ag = data["AR_AGING"]
    ag_rows, ag_col = [], {}
    keep = [r for r in ag if str(r["year"]) in ("2023", "2024", "2025", "2026")]
    older = [r for r in ag if str(r["year"]) not in ("2023", "2024", "2025", "2026")]
    agg = [("2022 & before", sum(r["month_start"] or 0 for r in older), sum(r["prev_week"] or 0 for r in older), sum(r["current"] or 0 for r in older))] + [(str(r["year"]), r["month_start"], r["prev_week"], r["current"]) for r in keep]
    for i, (lab, m, p, c) in enumerate(agg):
        ag_rows.append([lab, f0(m), f0(p), f0(c), fd(c - p), fd(c - m)]); ag_col[(i, 4)] = delta_color(c - p); ag_col[(i, 5)] = delta_color(c - m)
    ag_rows.append(["Total AR", f0(tot["month_start"]), f0(tot["prev_week"]), f0(tot["current"]), fd(dwow(tot)), fd(dms(tot))])
    tb(s, 8.3, 1.25, 4.7, 0.3, "AR aging by year of origin", size=13, bold=True, color=NAVY)
    table(s, 8.3, 1.6, 4.7, ["Year", f"M. start\n{ms_s}", prev_s, cur_s, "Δ WoW", "Δ vs\nm. start"], ag_rows, col_w=[1.4, 0.9, 0.9, 0.9, 0.8, 0.9], font_size=9, colors=ag_col, total_last=True, row_h=0.28)
    # chart deltas by component
    bar_chart(s, 8.2, 3.75, 4.9, 3.3, ["Advance", "Unbilled", "Unpaid", "Total AR"],
              [("Δ WoW", [dwow(adv), dwow(unb), dwow(unp), dwow(tot)]), ("Δ vs month start", [dms(adv), dms(unb), dms(unp), dms(tot)])],
              [BLUE, GOLD], gap=80, number_format='+#,##0;−#,##0;0', title="Movement by component (MMAD)", label_size=8)

    # ------------------------------------------------------------------ 5. AR by BU
    s = D.slide("AR Position by Business Unit", f"This week vs last week and vs the month-start baseline ({ms_s}), MMAD; DSO excl. advances in days (BU meeting report)")
    hdr = ["Business Unit", f"Month start\n{ms_s}", f"Last week\n{prev_s}", f"This week\n{cur_s}", "Δ WoW", "Δ vs\nmonth start", "Unbilled", "Δ WoW", "Unpaid", "Δ WoW", "Advance", "Unrecog.\nrevenue", "Gross\nexposure", "DSO\n(days)"]
    rows, colors = [], {}
    bus = [b for b in BU_ORDER if b in arbu]
    sums = {k: 0.0 for k in ("ms", "pw", "cu", "ub", "ubp", "up", "upp", "ad", "un")}
    for i, b in enumerate(bus):
        t_, u_, p_, a_ = arbu[b]["Total AR"], arbu[b]["Unbilled"], arbu[b]["Unpaid"], arbu[b]["Advance"]
        un = (arbu[b].get("Unrecognized") or {}).get("current") or 0
        dr = dso.get(b, {}); dsv = (str(int(float(dr["dso_current"]))) if dr.get("dso_current") not in (None, "") else "–")
        rows.append([short(b), f0(t_["month_start"]), f0(t_["prev_week"]), f0(t_["current"]), fd(dwow(t_)), fd(dms(t_)), f0(u_["current"]), fd(dwow(u_)), f0(p_["current"]), fd(dwow(p_)), f0(a_["current"]), f0(un), f0(t_["current"] + un), dsv])
        colors[(i, 4)] = delta_color(dwow(t_)); colors[(i, 5)] = delta_color(dms(t_)); colors[(i, 7)] = delta_color(dwow(u_)); colors[(i, 9)] = delta_color(dwow(p_))
        sums["ms"] += t_["month_start"] or 0; sums["pw"] += t_["prev_week"] or 0; sums["cu"] += t_["current"] or 0; sums["ub"] += u_["current"] or 0; sums["ubp"] += u_["prev_week"] or 0
        sums["up"] += p_["current"] or 0; sums["upp"] += p_["prev_week"] or 0; sums["ad"] += a_["current"] or 0; sums["un"] += un
    gd = dso.get("JESA Group", {}); gdv = (str(int(float(gd["dso_current"]))) if gd.get("dso_current") not in (None, "") else "–")
    rows.append(["Total JESA", f0(sums["ms"]), f0(sums["pw"]), f0(sums["cu"]), fd(sums["cu"] - sums["pw"]), fd(sums["cu"] - sums["ms"]), f0(sums["ub"]), fd(sums["ub"] - sums["ubp"]), f0(sums["up"]), fd(sums["up"] - sums["upp"]), f0(sums["ad"]), f0(sums["un"]), f0(sums["cu"] + sums["un"]), gdv])
    n = len(rows) - 1
    colors[(n, 4)] = delta_color(sums["cu"] - sums["pw"]); colors[(n, 5)] = delta_color(sums["cu"] - sums["ms"]); colors[(n, 7)] = delta_color(sums["ub"] - sums["ubp"]); colors[(n, 9)] = delta_color(sums["up"] - sums["upp"])
    table(s, 0.35, 1.25, 12.65, hdr, rows, col_w=[1.6, 0.95, 0.95, 0.95, 0.8, 0.9, 0.9, 0.8, 0.85, 0.8, 0.8, 0.9, 0.95, 0.7], font_size=9, header_size=8.5, row_h=0.265, header_h=0.5, total_last=True, first_col_bold=True)
    yb = 1.25 + 0.5 + 0.265 * len(rows) + 0.2
    hb = SH - 0.45 - yb
    bar_chart(s, 0.3, yb, 6.3, hb, [short(b) for b in bus[:9]], [("Δ WoW", [dwow(arbu[b]["Total AR"]) for b in bus[:9]])], [BLUE], legend=False, number_format='+#,##0;−#,##0;0', title="Total AR – Δ week on week by BU (MMAD)", cat_size=8, gap=50)
    bar_chart(s, 6.75, yb, 6.3, hb, [short(b) for b in bus[:9]], [(f"Δ vs {ms_s}", [dms(arbu[b]["Total AR"]) for b in bus[:9]])], [GOLD], legend=False, number_format='+#,##0;−#,##0;0', title=f"Total AR – Δ vs month start ({ms_s}) by BU (MMAD)", cat_size=8, gap=50)

    # ------------------------------------------------------------------ 6. Unbilled & unpaid status
    s = D.slide("Unbilled & Unpaid – status view", "Where the receivable book sits in the process, with week-on-week movement")
    own = {}
    for r in data["AR_UNBILLED_OWNER"]:
        o = own.setdefault(r["category"], {"ms": 0, "pw": 0, "cu": 0}); o["ms"] += r["month_start"] or 0; o["pw"] += r["prev_week"] or 0; o["cu"] += r["current"] or 0
    tb(s, 0.4, 1.25, 6, 0.3, "Unbilled by ownership (MMAD)", size=13, bold=True, color=NAVY)
    rows, colors = [], {}
    for i, k in enumerate(["Under OCP", "SS postponed", "Under JESA", "At risk"]):
        o = own.get(k, {"ms": 0, "pw": 0, "cu": 0})
        rows.append([k, f0(o["ms"]), f0(o["pw"]), f0(o["cu"]), f"{o['cu']/unb['current']*100:.0f}%", fd(o["cu"] - o["pw"]), fd(o["cu"] - o["ms"])]); colors[(i, 5)] = delta_color(o["cu"] - o["pw"]); colors[(i, 6)] = delta_color(o["cu"] - o["ms"])
    rows.append(["Unbilled total", f0(unb["month_start"]), f0(unb["prev_week"]), f0(unb["current"]), "100%", fd(dwow(unb)), fd(dms(unb))])
    table(s, 0.4, 1.6, 6.2, ["Category", f"M. start\n{ms_s}", prev_s, cur_s, "% AR", "Δ WoW", "Δ vs\nm. start"], rows, col_w=[1.7, 0.85, 0.85, 0.85, 0.7, 0.8, 0.9], font_size=9.5, colors=colors, total_last=True, row_h=0.3)
    # top statuses
    st = sorted([r for r in data["AR_UNBILLED_STATUS"] if r["group"] in ("OCP", "JESA", "At Risk")], key=lambda r: -(r["current"] or 0))[:9]
    tb(s, 0.4, 3.75, 6.2, 0.3, "Largest unbilled statuses (MMAD)", size=13, bold=True, color=NAVY)
    rows, colors = [], {}
    for i, r in enumerate(st):
        lab = str(r["status"]); lab = (lab[:43] + "…") if len(lab) > 44 else lab
        d1 = (r["current"] or 0) - (r["prev_week"] or 0)
        rows.append([f"{r['group']}", lab, f0(r["prev_week"]), f0(r["current"]), fd(d1)]); colors[(i, 4)] = delta_color(d1)
    table(s, 0.4, 4.1, 6.2, ["Owner", "Status", prev_s, cur_s, "Δ WoW"], rows, col_w=[0.7, 3.2, 0.8, 0.8, 0.7], font_size=8.5, colors=colors, row_h=0.27, header_h=0.3, align=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])
    # unpaid
    tb(s, 6.95, 1.25, 6, 0.3, "Unpaid by OCP processing status (MMAD)", size=13, bold=True, color=NAVY)
    ups = data["AR_UNPAID_STATUS"]; rows, colors = [], {}
    for i, r in enumerate(ups):
        d1 = (r["current"] or 0) - (r["prev_week"] or 0); d2 = (r["current"] or 0) - (r["month_start"] or 0)
        rows.append([r["status"], f0(r["month_start"]), f0(r["prev_week"]), f0(r["current"]), f"{(r['current'] or 0)/unp['current']*100:.0f}%", fd(d1), fd(d2)]); colors[(i, 5)] = delta_color(d1); colors[(i, 6)] = delta_color(d2)
    rows.append(["Unpaid total", f0(unp["month_start"]), f0(unp["prev_week"]), f0(unp["current"]), "100%", fd(dwow(unp)), fd(dms(unp))])
    table(s, 6.95, 1.6, 6.05, ["Status", f"M. start\n{ms_s}", prev_s, cur_s, "% unpaid", "Δ WoW", "Δ vs\nm. start"], rows, col_w=[1.7, 0.8, 0.8, 0.8, 0.75, 0.75, 0.85], font_size=9.5, colors=colors, total_last=True, row_h=0.3)
    bar_chart(s, 6.9, 4.05, 6.15, 3.05, [r["status"] for r in ups], [(prev_s, [r["prev_week"] for r in ups]), (cur_s, [r["current"] for r in ups])], [LBLUE, NAVY2], gap=60, title="Unpaid by status – last week vs this week", cat_size=8, label_size=8)


    # ------------------------------------------------------------------ 6b. Non-OCP receivables
    no = {}
    for r in data.get("NONOCP_BU", []): no.setdefault(r["bu"], {})[r["component"]] = r
    def nv(b, c, k): return ((no.get(b, {}).get(c) or {}).get(k)) or 0
    n_unb = sum(nv(b, "Unbilled", "current") for b in no); n_unp = sum(nv(b, "Unpaid", "current") for b in no)
    n_unb_ms = sum(nv(b, "Unbilled", "month_start") for b in no); n_unp_ms = sum(nv(b, "Unpaid", "month_start") for b in no)
    s = D.slide("Non-OCP Clients – unbilled & unpaid", f"Non-OCP receivables of {f0(n_unb + n_unp)} MMAD ({(n_unb+n_unp)/tot['current']*100:.1f}% of AR): unbilled {f0(n_unb)}, unpaid {f0(n_unp)}; {fd(round(n_unb+n_unp)-round(n_unb_ms+n_unp_ms))} vs month start")
    cw = (SW - 0.6 - 0.12 * 3) / 4
    kpi_card(s, 0.3, 1.2, cw, 1.1, "Non-OCP AR (unbilled + unpaid)", n_unb + n_unp, delta_ms=round(n_unb + n_unp) - round(n_unb_ms + n_unp_ms), value_size=20, ms_label="vs month start")
    kpi_card(s, 0.3 + (cw + 0.12), 1.2, cw, 1.1, "Non-OCP unbilled", n_unb, delta_ms=round(n_unb) - round(n_unb_ms), value_size=20, ms_label="vs month start")
    kpi_card(s, 0.3 + 2 * (cw + 0.12), 1.2, cw, 1.1, "Non-OCP unpaid", n_unp, delta_ms=round(n_unp) - round(n_unp_ms), value_size=20, ms_label="vs month start")
    kpi_card(s, 0.3 + 3 * (cw + 0.12), 1.2, cw, 1.1, "Share of total AR", None, value_size=20)
    tb(s, 0.3 + 3 * (cw + 0.12) + 0.15, 1.5, cw - 0.2, 0.5, f"{(n_unb+n_unp)/tot['current']*100:.1f}%", size=20, bold=True, color=NAVY)
    tb(s, 0.4, 2.45, 5.6, 0.3, "By business unit (MMAD)", size=12, bold=True, color=NAVY)
    rows, colors = [], {}
    nb = sorted([b for b in no if nv(b, "Unbilled", "current") + nv(b, "Unpaid", "current") > 0.5], key=lambda b: -(nv(b, "Unbilled", "current") + nv(b, "Unpaid", "current")))
    for i, b in enumerate(nb):
        ub_, up_ = nv(b, "Unbilled", "current"), nv(b, "Unpaid", "current"); tot_ = ub_ + up_
        d_ = round(tot_) - round(nv(b, "Unbilled", "month_start") + nv(b, "Unpaid", "month_start"))
        rows.append([short(b), f0(ub_), f0(up_), f0(tot_), f"{tot_/(n_unb+n_unp)*100:.0f}%", fd(d_)]); colors[(i, 5)] = delta_color(d_)
    rows.append(["Total non-OCP", f0(n_unb), f0(n_unp), f0(n_unb + n_unp), "100%", fd(round(n_unb+n_unp)-round(n_unb_ms+n_unp_ms))])
    table(s, 0.4, 2.75, 5.6, ["BU", "Unbilled", "Unpaid", "Total", "%", "Δ vs MS"], rows, col_w=[1.7, 0.8, 0.8, 0.8, 0.6, 0.8], font_size=9, header_size=8.5, row_h=0.3, header_h=0.32, colors=colors, total_last=True)
    tb(s, 0.4, 2.75 + 0.32 + 0.3 * len(rows) + 0.15, 5.6, 1.0, [[(f"Non-OCP clients carry {f0(n_unb + n_unp)} MMAD, of which Global Mining (Ma'aden) accounts for {nv('Global Mining', 'Unbilled', 'current') + nv('Global Mining', 'Unpaid', 'current'):,.0f} MMAD. Statuses: C-* pending with the client, J-* pending with JESA.", {})]], size=9, color=DARK, line_spacing=1.1)
    npr = data.get("NONOCP_PROJECTS", [])[:14]
    tb(s, 6.3, 2.45, 6.8, 0.3, "Projects driving non-OCP receivables (MMAD)", size=12, bold=True, color=NAVY)
    def clean_txt(v, n):
        v = str(v or "").replace("\xa0", " ").strip(); return (v[:n - 1] + "…") if len(v) > n else v
    rows = [[short(r["bu"]), clean_txt(r["client"], 24), clean_txt(r["project"], 30), f1(r["unbilled"]), f1(r["unpaid"]), clean_txt(r["status"], 27)] for r in npr]
    rows.append([f"Top {len(npr)} total", "", "", f1(sum((r["unbilled"] or 0) for r in npr)), f1(sum((r["unpaid"] or 0) for r in npr)), f"{sum((r['unbilled'] or 0)+(r['unpaid'] or 0) for r in npr)/(n_unb+n_unp)*100:.0f}% of non-OCP"])
    table(s, 6.3, 2.75, 6.8, ["BU", "Client", "Project", "Unbilled", "Unpaid", "Status"], rows, col_w=[0.95, 1.5, 1.85, 0.6, 0.6, 1.55], font_size=7.5, header_size=7.5, row_h=0.255, header_h=0.28, total_last=True,
          align=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.LEFT])

    # ------------------------------------------------------------------ 6c. Unbilled under JESA responsibility
    ju = {}
    for r in data.get("JESA_UNBILLED_BU", []): ju.setdefault(r["bu"], {})[r["split"]] = r
    def jv(b, sp, k): return ((ju.get(b, {}).get(sp) or {}).get(k)) or 0
    j_lab = sum(jv(b, "labor", "current") for b in ju); j_nl = sum(jv(b, "non_labor", "current") for b in ju); j_tot = j_lab + j_nl
    j_lab_ms = sum(jv(b, "labor", "month_start") for b in ju); j_nl_ms = sum(jv(b, "non_labor", "month_start") for b in ju)
    owner_rows = [r for r in data["AR_UNBILLED_OWNER"] if r["category"] == "Under JESA"]
    j_prev = sum((r["prev_week"] or 0) for r in owner_rows) if owner_rows else None
    s = D.slide("Unbilled under JESA responsibility", f"{f0(j_tot)} MMAD of unbilled sits with JESA ({j_tot/unb['current']*100:.0f}% of unbilled): PSR {f0(j_lab)} / Procurement {f0(j_nl)}; {fd(round(j_tot)-round(j_lab_ms+j_nl_ms))} vs month start" + (f", {fd(round(j_tot)-round(j_prev))} WoW" if j_prev is not None else ""))
    cw = (SW - 0.6 - 0.12 * 3) / 4
    kpi_card(s, 0.3, 1.2, cw, 1.2, "Unbilled under JESA", j_tot, delta_wow=(round(j_tot) - round(j_prev)) if j_prev is not None else None, delta_ms=round(j_tot) - round(j_lab_ms + j_nl_ms), value_size=20, ms_label="vs month start")
    kpi_card(s, 0.3 + (cw + 0.12), 1.2, cw, 1.2, "PSR (labor)", j_lab, delta_ms=round(j_lab) - round(j_lab_ms), value_size=20, ms_label="vs month start")
    kpi_card(s, 0.3 + 2 * (cw + 0.12), 1.2, cw, 1.2, "Procurement (non-labor)", j_nl, delta_ms=round(j_nl) - round(j_nl_ms), value_size=20, ms_label="vs month start")
    kpi_card(s, 0.3 + 3 * (cw + 0.12), 1.2, cw, 1.2, "Share of total unbilled", None, value_size=20)
    tb(s, 0.3 + 3 * (cw + 0.12) + 0.15, 1.5, cw - 0.2, 0.5, f"{j_tot/unb['current']*100:.0f}%", size=20, bold=True, color=NAVY)
    tb(s, 0.4, 2.55, 6.2, 0.3, "By business unit – PSR vs Procurement (MMAD)", size=12, bold=True, color=NAVY)
    rows, colors = [], {}
    jb = sorted([b for b in ju if abs(jv(b, "labor", "current")) + abs(jv(b, "non_labor", "current")) > 0.05], key=lambda b: -(jv(b, "labor", "current") + jv(b, "non_labor", "current")))
    for i, b in enumerate(jb):
        l_, n_ = jv(b, "labor", "current"), jv(b, "non_labor", "current"); t_ = l_ + n_
        d_ = round(t_) - round(jv(b, "labor", "month_start") + jv(b, "non_labor", "month_start"))
        ub_b = (arbu[b]["Unbilled"]["current"] or 0) if b in arbu else 0
        rows.append([short(b), f0(l_), f0(n_), f0(t_), f"{t_/ub_b*100:.0f}%" if ub_b else "–", fd(d_)]); colors[(i, 5)] = delta_color(d_)
    rows.append(["Total JESA", f0(j_lab), f0(j_nl), f0(j_tot), f"{j_tot/unb['current']*100:.0f}%", fd(round(j_tot)-round(j_lab_ms+j_nl_ms))])
    table(s, 0.4, 2.85, 6.2, ["BU", "PSR", "Procurement", "Total", "% of BU unbilled", "Δ vs MS"], rows, col_w=[1.6, 0.8, 1.0, 0.8, 1.1, 0.8], font_size=9, header_size=8.5, row_h=0.3, header_h=0.34, colors=colors, total_last=True)
    tb(s, 0.4, 2.85 + 0.34 + 0.3 * len(rows) + 0.1, 6.2, 0.5, "PSR = labor scope; Procurement = non-labor scope. Negative procurement amounts are credit lines pending JESA PM approval (AM, Global Mining).", size=8, italic=True, color=GREY)
    js = data.get("JESA_UNBILLED_STATUS", [])
    tb(s, 6.9, 2.55, 6.2, 0.3, "By status – what JESA still has to do (MMAD)", size=12, bold=True, color=NAVY)
    rows = [[(str(r["status"])[:44] + "…") if len(str(r["status"])) > 45 else str(r["status"]), f0(r["labor"]), f0(r["non_labor"]), f0((r["labor"] or 0) + (r["non_labor"] or 0)), f"{((r['labor'] or 0)+(r['non_labor'] or 0))/j_tot*100:.0f}%" if j_tot else "–"] for r in js]
    rows.append(["Total under JESA", f0(j_lab), f0(j_nl), f0(j_tot), "100%"])
    table(s, 6.9, 2.85, 6.2, ["Status", "PSR", "Procurement", "Total", "%"], rows, col_w=[3.1, 0.75, 0.95, 0.75, 0.65], font_size=8, header_size=8, row_h=0.25, header_h=0.28, total_last=True, align=[PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])
    tb(s, 6.9, 2.85 + 0.28 + 0.25 * len(rows) + 0.1, 6.2, 0.7, [[("J-* statuses are the steps under JESA control (attachment creation, internal reviews, PM approval, service-sheet submission, invoicing); service sheets postponed and at-risk scope are excluded. Negative values are credit lines pending JESA PM approval.", {})]], size=8.5, color=DARK, line_spacing=1.1)

    # ------------------------------------------------------------------ 7. Unrecognized revenue & suppliers
    ut = data["UNRECOGNIZED"]
    MK = 1.06  # JESA cost + fee markup: unrecognized revenue = supplier cost x 1.06 (UR consolidation columns already include it)
    ur_rev = unrec["current"] or 0; ur_cost = ur_rev / MK
    ur_prev, ur_ms = unrec.get("prev_week"), unrec.get("month_start")
    sub_txt = f"Unrecognized revenue of {f0(ur_rev)} MMAD = supplier balance ({f0(ur_cost)}) plus a 6% markup; {fd(dwow(unrec))} vs {prev_s}"
    if ur_ms is not None: sub_txt += f", {fd(dms(unrec))} vs month start"
    s = D.slide("Unrecognized Revenue & Suppliers", sub_txt)
    card(s, 0.4, 1.3, 6.1, 1.15, "Unrecognized revenue", f"{f0(ur_rev)} MMAD", f"= unrecognized supplier cost × {MK} cost-and-fee markup  ·  {prev_s}: {f0(ur_prev)}" + (f"  ·  month start: {f0(ur_ms)}" if ur_ms is not None else ""), fill=BLUE, value_size=24, sub_size=8.5)
    card(s, 6.85, 1.3, 6.15, 1.15, "Unrecognized suppliers — cost", f"{f0(ur_cost)} MMAD", "Executed but unbooked supplier & equipment scope (HT, excl. DP)", fill=SLATE, value_size=24, sub_size=8.5)
    # by business unit (revenue incl. markup), one colour per BU
    BU_COL = {"SP2M Ad Hoc Program": GOLD, "Water & Environment": NAVY2, "Fertilizers & Chemicals": TEAL, "Energy Operations": GREEN, "JESA Technologies": BLUE,
              "Global Mining": GREY, "Building & Infrastructure": RED, "AM": NAVY, "Advisory": PURPLE, "Building & Capabilities": ORANGE, "Power to X": LBLUE}
    utb = sorted([r for r in ut if (r["grand_total"] or 0) > 0.05], key=lambda r: -(r["grand_total"] or 0))
    tb(s, 0.4, 2.6, 6, 0.3, "Unrecognized Revenue by Business Unit", size=13, bold=True, color=NAVY)
    tb(s, 0.4, 2.88, 6, 0.25, "6% JESA cost + fee markup included (MMAD)" + (f"  ·  Δ vs {prev_s} in brackets" if ur_prev is not None else ""), size=9, color=SLATE)
    cats = [short(r["bu"]) for r in utb]
    if ur_prev is not None:
        cats = [f"{short(r['bu'])}  ({fd(round(r['grand_total'] or 0) - round(r['prev_week_total'] or 0))})" if r.get("prev_week_total") is not None else short(r["bu"]) for r in utb]
    ch = bar_chart(s, 0.3, 3.15, 6.3, 3.6, cats, [("Unrecognized revenue", [r["grand_total"] for r in utb])], [BLUE], horizontal=True, gap=35, cat_size=8.5, label_size=8.5, legend=False)
    color_points(ch, 0, [BU_COL.get(r["bu"], BLUE) for r in utb])
    tb(s, 0.4, 6.78, 6.2, 0.26, "No revenue recognition has been booked on these balances to date.", size=8.5, bold=True, color=RED)
    # Table 7 - unrecognized suppliers by status, at cost
    tb(s, 6.85, 2.6, 6.2, 0.3, "Table 7 — Unrecognized Suppliers by Status", size=13, bold=True, color=NAVY)
    tb(s, 6.85, 2.88, 6.2, 0.25, "Executed but unbooked scope, at cost (MMAD)", size=9, color=SLATE)
    stat = [("Pend. Submission", "pending_submission"), ("Pend. Approval", "pending_approval"), ("Under CN", "under_cn"), ("At-Risk", "at_risk")]
    def st_sum(pfx, key):
        v = sum((r[f"{pfx}_{key}"] or 0) for r in ut)
        if key == "pending_approval": v += sum((r[f"{pfx}_signed_by_client"] or 0) for r in ut)
        return v / MK
    eq = [st_sum("eq", k) for _, k in stat]; sb = [st_sum("sub", k) for _, k in stat]
    ch = bar_chart(s, 6.75, 3.1, 6.35, 2.05, [n for n, _ in stat], [("Equipment", eq), ("Subcontractors", sb)], [NAVY2, GOLD], stacked=True, horizontal=True, gap=45, cat_size=8.5, label_size=8, label_colors=[WHITE, DARK], legend_pos=XL_LEGEND_POSITION.TOP)
    hide_small_labels(ch, [eq, sb], 60)
    rows = [["Pending Approval", "Attachments submitted (PROCORE etc.), under the approval workflow; includes attachments already signed by the client"],
            ["Pending Submission", "Executed scope not yet submitted for approval; may include potential / impending scope"],
            ["Under CN", "Executed scope via a Change Notice (or a potential CN)"],
            ["At-Risk", "Executed scope without any formal contract"]]
    table(s, 6.85, 5.2, 6.2, ["Status (Table 7)", "Definition"], rows, col_w=[1.5, 4.7], font_size=8.5, header_size=9, row_h=0.33, header_h=0.3, align=[PP_ALIGN.LEFT, PP_ALIGN.LEFT], first_col_bold=True, colors={(i, 0): BLUE for i in range(len(rows))})
    tb(s, 6.85, 6.78, 6.2, 0.26, "*** Amounts in HT — executed progress, excluding DP amortization. Status split at cost (= revenue ÷ 1.06).", size=8, italic=True, color=SLATE)

    # ------------------------------------------------------------------ 8. Collections (payments received)
    cb = data["COLLECTIONS_BU"]; cm = data["COLLECTIONS_MONTHLY"]
    oct_cur = sum(r["month_current"] or 0 for r in cb); oct_pw = sum(r["month_prev_week"] or 0 for r in cb); prior = sum(r["prior_month"] or 0 for r in cb)
    ytd = sum(r["ytd_current"] or 0 for r in cb); ytd_pw = sum(r["ytd_prev_week"] or 0 for r in cb); ninv = sum(int(r["invoices_ytd"] or 0) for r in cb)
    month_name = datef(meta["as_of_date"], "%B"); prior_name = (dt.date.fromisoformat(str(meta["as_of_date"])[:10]).replace(day=1) - dt.timedelta(days=1)).strftime("%B")
    dep = sum((r.get("deposited_month") or 0) for r in cb)
    s = D.slide("Payments Received (Collections)", f"{f0(oct_cur)} MMAD of invoices paid in {month_name} to {cur_s} ({fd(oct_cur-oct_pw)} vs {prev_s}); {ytd/1000:,.2f} Bn MAD year to date")
    cw = (SW - 0.6 - 0.12 * 4) / 5
    kpi_card(s, 0.3, 1.25, cw, 1.15, f"{month_name} to date", oct_cur, delta_wow=oct_cur - oct_pw, good_when_negative=False)
    kpi_card(s, 0.3 + (cw + 0.12), 1.25, cw, 1.15, f"{prior_name} (full month)", prior)
    kpi_card(s, 0.3 + 2 * (cw + 0.12), 1.25, cw, 1.15, "YTD received", ytd, delta_wow=ytd - ytd_pw, good_when_negative=False)
    lab_ytd = sum(r["labor"] or 0 for r in cm); nl_ytd = sum(r["non_labor"] or 0 for r in cm)
    kpi_card(s, 0.3 + 3 * (cw + 0.12), 1.25, cw, 1.15, "Labor / non-labor YTD", None, value_size=18)
    tb(s, 0.3 + 3 * (cw + 0.12) + 0.15, 1.55, cw - 0.2, 0.8, [[(f"{f0(lab_ytd)} ", {"bold": True, "color": BLUE, "size": 16}), (f"· {lab_ytd/(lab_ytd+nl_ytd)*100:.0f}% labor", {"size": 9, "color": SLATE})], [(f"{f0(nl_ytd)} ", {"bold": True, "color": ORANGE, "size": 16}), (f"· {nl_ytd/(lab_ytd+nl_ytd)*100:.0f}% non-labor", {"size": 9, "color": SLATE})]], size=10)
    kpi_card(s, 0.3 + 4 * (cw + 0.12), 1.25, cw, 1.15, f"Invoices deposited in {month_name[:3]}", dep)
    tb(s, 0.3 + 4 * (cw + 0.12) + 0.15, 2.08, cw - 0.2, 0.26, f"{ninv:,} invoices paid YTD", size=8.5, color=SLATE)
    months = [r for r in cm if (r["labor"] or 0) + (r["non_labor"] or 0) > 0]
    tb(s, 0.4, 2.6, 7, 0.3, f"Monthly payments received 2026 (MMAD, amount invoiced by payment date) – {month_name} in progress", size=12.5, bold=True, color=NAVY)
    ch = bar_chart(s, 0.3, 2.9, 7.3, 4.2, [r["month"] for r in months], [("Labor (PSR)", [r["labor"] for r in months]), ("Non-labor (procurement)", [r["non_labor"] for r in months])], [BLUE, ORANGE], stacked=True, gap=45, label_size=7.5, label_colors=[WHITE, DARK])
    hide_small_labels(ch, [[r["labor"] for r in months], [r["non_labor"] for r in months]], 45)
    tb(s, 7.85, 2.6, 5.2, 0.3, "By business unit (MMAD)", size=13, bold=True, color=NAVY)
    rows, colors = [], {}
    cbs = sorted(cb, key=lambda r: -(r["ytd_current"] or 0))
    for i, r in enumerate(cbs):
        d1 = (r["month_current"] or 0) - (r["month_prev_week"] or 0)
        rows.append([short(r["bu"]), f0(r["month_prev_week"]), f0(r["month_current"]), fd(d1), f0(r["prior_month"]), f0(r["ytd_current"]), f"{(r['ytd_current'] or 0)/ytd*100:.0f}%"]); colors[(i, 3)] = delta_color(d1, False)
    rows.append(["Total", f0(oct_pw), f0(oct_cur), fd(oct_cur - oct_pw), f0(prior), f0(ytd), "100%"])
    table(s, 7.85, 2.9, 5.2, ["BU", f"{month_name[:3]} to {prev_s}", f"{month_name[:3]} to {cur_s}", "Δ WoW", prior_name[:3], "YTD", "% YTD"], rows, col_w=[1.4, 0.8, 0.8, 0.7, 0.7, 0.8, 0.6], font_size=8.5, header_size=8, colors=colors, total_last=True, row_h=0.26, header_h=0.4)

    # ------------------------------------------------------------------ 9. Cash position – procurement
    cmo = {r["line"]: r for r in data["CASH_MONTHLY"]}
    cin = cmo["Total cash IN (procurement)"]; cout = cmo["Total cash OUT (procurement)"]; net = cmo["NET position (cash IN + cash OUT)"]
    sup = cmo["Cash OUT suppliers (by tier)"]; lcp = cmo["Cash OUT LC payments"]; cus = cmo["Cash OUT customs"]; div = cmo["Dividends / PROC compensable"]; proc = cmo["Procurement collections (OCP payment report)"]
    psr = cmo["PSR collections (labor)"]; prov = cmo["PSR provisions (AM ring-fence + W&E reserve)"]; npsr = cmo["NET PSR after provisions"]
    ong = {r["bucket"]: r for r in data["ONGOING"]}
    appr_lc = ong["Approved – October LC maturities"]["amount"]; appr_sup = ong["Approved – supplier invoices"]["amount"]; pend = ong["Pending payment approval"]["amount"]; paid_oct = ong["Paid in October to date (suppliers + LC + customs)"]["amount"]
    s = D.slide("Cash Position – Procurement (JESA Group incl. AM)", f"Jul–{month_name[:3]} 2026: cash IN {f0(cin['cumulative'])} vs cash OUT {f0(-cout['cumulative'])} MMAD incl. {month_name} commitments → net {fd(net['cumulative'])} MMAD")
    cw = (SW - 0.6 - 0.12 * 4) / 5
    card(s, 0.3, 1.25, cw, 1.15, "Cash IN Jul–Oct", f0(cin["cumulative"]), f"OCP payments {f0(proc['cumulative'])} + dividends {f0(div['cumulative'])}", fill=TEAL, value_size=22)
    card(s, 0.3 + (cw + 0.12), 1.25, cw, 1.15, "Cash OUT actuals", f0(-(cout["jul"] + cout["aug"] + cout["sep"] + cout["oct_actual"])), f"Jul–Sep {f0(-(cout['jul']+cout['aug']+cout['sep']))} + Oct to {meta['payments_to']} {f0(-cout['oct_actual'])}", fill=NAVY2, value_size=22)
    card(s, 0.3 + 2 * (cw + 0.12), 1.25, cw, 1.15, f"{month_name} commitments", f0(-cout["oct_ongoing"]), f"Approved {f0(appr_lc + appr_sup)} (LC {f0(appr_lc)}) + pending {f0(pend)}", fill=GOLD, value_size=22)
    card(s, 0.3 + 3 * (cw + 0.12), 1.25, cw, 1.15, "Net position", fd(net["cumulative"]), "Cash IN + cash OUT incl. commitments", fill=(RED if net["cumulative"] < 0 else GREEN), value_size=22)
    card(s, 0.3 + 4 * (cw + 0.12), 1.25, cw, 1.15, "PSR (labor) net", fd(npsr["cumulative"]), f"Collections {f0(psr['cumulative'])} less provisions {f0(-prov['cumulative'])}", fill=SLATE, value_size=22)
    mcats = ["Jul", "Aug", "Sep", f"Oct to date", "Oct commitments"]
    keys = ["jul", "aug", "sep", "oct_actual", "oct_ongoing"]
    tb(s, 0.4, 2.6, 7, 0.3, "Monthly cash IN vs cash OUT – procurement (MMAD)", size=13, bold=True, color=NAVY)
    ch = bar_chart(s, 0.3, 2.9, 6.9, 4.0, mcats, [("Cash IN", [cin[k] for k in keys]), ("Cash OUT", [-cout[k] for k in keys]), ("Net", [net[k] for k in keys])], [TEAL, NAVY2, GOLD], gap=70, label_size=8, number_format='#,##0;−#,##0;0', overlap=-10)
    tb(s, 7.45, 2.6, 5.6, 0.3, "Cash position by month (MMAD)", size=13, bold=True, color=NAVY)
    rows, colors = [], {}
    def rowf(lab, r, sign=1, bold=False):
        return [lab] + [f0(sign * (r[k] or 0)) for k in keys] + [f0(sign * (r["cumulative"] or 0))]
    rows.append(rowf("OCP procurement collections", proc)); rows.append(rowf("Dividends / PROC compensable", div)); rows.append(rowf("Total cash IN", cin))
    rows.append(rowf("Suppliers (by tier)", sup)); rows.append(rowf("LC payments", lcp)); rows.append(rowf("Customs", cus)); rows.append(rowf("Total cash OUT", cout))
    rows.append([ "NET position"] + [fd(net[k]) for k in keys] + [fd(net["cumulative"])])
    for j in range(1, 7): colors[(7, j)] = delta_color(net[keys[j - 1]] if j < 6 else net["cumulative"], good_when_negative=False)
    rows.append(rowf("PSR collections (labor)", psr)); rows.append(rowf("PSR provisions", prov)); rows.append(["NET PSR after provisions"] + [fd(npsr[k]) for k in keys] + [fd(npsr["cumulative"])])
    table(s, 7.45, 2.9, 5.6, ["MMAD", "Jul", "Aug", "Sep", f"Oct to {meta['payments_to']}", "Oct ongoing", "Cumul."], rows, col_w=[2.1, 0.65, 0.65, 0.65, 0.8, 0.8, 0.75], font_size=8.5, header_size=8, bold_rows=(2, 6, 7, 10), colors=colors, row_h=0.27, header_h=0.36)
    tb(s, 7.45, 6.55, 5.6, 0.55, f"Basis: {meta['cash_basis']}", size=7, italic=True, color=GREY)

    # ------------------------------------------------------------------ 10. Cash by BU
    cbu = sorted(data["CASH_BU"], key=lambda r: -(r["cash_in"] or 0))
    s = D.slide("Cash Position by Business Unit – Procurement", f"Cumulative Jul–{month_name[:3]} 2026 cash IN, cash OUT and net balance by BU, before and after {month_name} commitments (MMAD)")
    rows, colors = [], {}
    tot_r = {k: 0.0 for k in ("cash_in", "cash_out_suppliers", "cash_out_lc", "cash_out_customs", "net_actuals", "oct_approved", "net_after_approved", "oct_pending", "net_all_ongoing")}
    for i, r in enumerate(cbu):
        rows.append([short(r["bu"]), f0(r["cash_in"]), f0(-(r["cash_out_suppliers"] or 0)), f0(-(r["cash_out_lc"] or 0)), f0(-(r["cash_out_customs"] or 0)), fd(r["net_actuals"]), f0(-(r["oct_approved"] or 0)), fd(r["net_after_approved"]), f0(-(r["oct_pending"] or 0)), fd(r["net_all_ongoing"])])
        for j, k in ((5, "net_actuals"), (7, "net_after_approved"), (9, "net_all_ongoing")): colors[(i, j)] = delta_color(r[k], False)
        for k in tot_r: tot_r[k] += r[k] or 0
    rows.append(["Total BUs", f0(tot_r["cash_in"]), f0(-tot_r["cash_out_suppliers"]), f0(-tot_r["cash_out_lc"]), f0(-tot_r["cash_out_customs"]), fd(tot_r["net_actuals"]), f0(-tot_r["oct_approved"]), fd(tot_r["net_after_approved"]), f0(-tot_r["oct_pending"]), fd(tot_r["net_all_ongoing"])])
    n = len(rows) - 1
    for j, k in ((5, "net_actuals"), (7, "net_after_approved"), (9, "net_all_ongoing")): colors[(n, j)] = delta_color(tot_r[k], False)
    table(s, 0.35, 1.25, 12.65, ["Business Unit", "Cash IN", "Paid suppliers", "Paid LC", "Paid customs", "Net actuals", f"{month_name[:3]} approved", "Net after approved", f"{month_name[:3]} pending", "Net all ongoing"], rows,
          col_w=[1.7, 1, 1.1, 0.9, 1, 1.05, 1.05, 1.2, 1.0, 1.15], font_size=9, header_size=8.5, row_h=0.27, header_h=0.42, total_last=True, first_col_bold=True)
    yb = 1.25 + 0.42 + 0.27 * len(rows) + 0.2
    bar_chart(s, 0.3, yb, 6.3, SH - 0.45 - yb, [short(r["bu"]) for r in cbu], [("Cash IN", [r["cash_in"] for r in cbu]), ("Cash OUT (incl. commitments)", [-((r["cash_out_suppliers"] or 0) + (r["cash_out_lc"] or 0) + (r["cash_out_customs"] or 0) + (r["oct_approved"] or 0) + (r["oct_pending"] or 0)) for r in cbu])], [TEAL, NAVY2], gap=50, cat_size=7.5, label_size=7, title="Cash IN vs cash OUT by BU (MMAD)")
    bar_chart(s, 6.75, yb, 6.3, SH - 0.45 - yb, [short(r["bu"]) for r in cbu], [("Net after all ongoing", [r["net_all_ongoing"] for r in cbu])], [GOLD], legend=False, gap=50, cat_size=7.5, label_size=7.5, title="Net balance after all ongoing commitments (MMAD)", number_format='+#,##0;−#,##0;0')
    ch = s.shapes[-1].chart; color_points(ch, 0, [GREEN if (r["net_all_ongoing"] or 0) >= 0 else RED for r in cbu])

    # ------------------------------------------------------------------ 11. Ongoing payments – October
    s = D.slide(f"{month_name} Payments – paid, approved and pending", f"{f0(paid_oct)} MMAD paid to {meta['payments_to']}; {f0(appr_lc+appr_sup)} MMAD approved (incl. {f0(appr_lc)} LC maturities) and {f0(pend)} MMAD pending approval")
    tb(s, 0.4, 1.25, 6.2, 0.3, f"{month_name} cash OUT: actual and committed (MMAD)", size=13, bold=True, color=NAVY)
    total_oct = round(paid_oct) + round(appr_lc) + round(appr_sup) + round(pend)
    waterfall(s, 0.3, 1.55, 6.3, 2.45, [f"Paid to {meta['payments_to']}", "Approved – LC maturities", "Approved – supplier invoices", "Pending approval", f"Total {month_name} outflow"],
              [paid_oct, appr_lc, appr_sup, pend, total_oct], ["total", "delta", "delta", "delta", "total"])
    # override colors: all additions in gold shades
    ch = s.shapes[-1].chart; color_points(ch, 1, [NAVY2, GOLD, GOLD, ORANGE, NAVY2])
    obo = sorted(data["CASH_BU_OCT"], key=lambda r: ((r["oct_paid_suppliers"] or 0) + (r["oct_paid_lc"] or 0) + (r["oct_paid_customs"] or 0) + (r["oct_ongoing_suppliers"] or 0) + (r["oct_ongoing_lc"] or 0)))
    tb(s, 0.4, 4.05, 6.2, 0.3, f"{month_name} by business unit (MMAD)", size=13, bold=True, color=NAVY)
    rows, colors = [], {}
    t = [0.0] * 5
    for i, r in enumerate(obo):
        paid = -((r["oct_paid_suppliers"] or 0) + (r["oct_paid_lc"] or 0) + (r["oct_paid_customs"] or 0)); os_ = -(r["oct_ongoing_suppliers"] or 0); ol = -(r["oct_ongoing_lc"] or 0); ci = r["oct_cash_in"] or 0
        netv = round(ci) - round(paid) - round(os_) - round(ol)
        rows.append([short(r["bu"]), f0(ci), f0(paid), f0(os_), f0(ol), fd(netv)]); colors[(i, 5)] = delta_color(netv, False)
        for j, v in enumerate((ci, paid, os_, ol, netv)): t[j] += v
    rows.append(["Total", f0(t[0]), f0(t[1]), f0(t[2]), f0(t[3]), fd(t[4])]); colors[(len(rows) - 1, 5)] = delta_color(t[4], False)
    table(s, 0.4, 4.35, 6.2, ["BU", "Cash IN to date", f"Paid to {meta['payments_to']}", "Ongoing suppliers", "Ongoing LC", "Net"], rows, col_w=[1.5, 1, 1, 1, 0.9, 0.8], font_size=8.5, header_size=8, row_h=0.215, header_h=0.34, colors=colors, total_last=True)
    # right: tracking log status
    tl = [r for r in data["ONGOING"] if str(r["bucket"]).startswith("Tracking log –") and not str(r["bucket"]).endswith("Total approved supplier payments") and (r["amount"] or 0) > 0.05]
    tl_tot = [r for r in data["ONGOING"] if str(r["bucket"]).endswith("Total approved supplier payments")]
    tl_tot = tl_tot[0] if tl_tot else {"amount": sum(r["amount"] for r in tl), "lines": sum(r["lines"] or 0 for r in tl)}
    tb(s, 6.95, 1.25, 6.1, 0.3, "Approved supplier invoices – status vs payment tracking log", size=13, bold=True, color=NAVY)
    tb(s, 6.95, 1.55, 6.1, 0.3, f"{f0(tl_tot['amount'])} MMAD / {int(tl_tot['lines'] or 0)} approved supplier lines by payment status (approved position excl. LC: {f0(appr_sup)} MMAD)", size=9.5, color=SLATE)
    def clean(lbl):
        l = str(lbl).replace("Tracking log – ", "").lower()
        if "pending ap" in l: return "Pending AP (unpaid)"
        if "pending treso" in l: return "Pending Treasury (unpaid)"
        if "paid after" in l: return "Paid after cut-off, awaiting SAP"
        if "not in sap" in l: return "Paid per log, not yet in SAP"
        if "reversed" in l: return "Paid, SAP clearing reversed"
        if "no date" in l: return "Paid per log, no date"
        if "cancelled" in l: return "Cancelled – to review"
        if "not in tracking" in l: return "Not in tracking log"
        return str(lbl).replace("Tracking log – ", "")[:40]
    tl_s = sorted(tl, key=lambda r: -(r["amount"] or 0))
    bar_chart(s, 6.85, 1.85, 6.2, 2.5, [clean(r["bucket"])[:32] for r in tl_s], [("Approved", [r["amount"] for r in tl_s])], [BLUE], horizontal=True, legend=False, gap=35, cat_size=8, label_size=8.5, number_format='#,##0.0')
    rows = [[clean(r["bucket"])[:60], str(int(r["lines"] or 0)), f1(r["amount"]), f"{(r['amount'] or 0)/(tl_tot['amount'] or 1)*100:.0f}%"] for r in tl_s]
    rows.append(["Total approved supplier payments", str(int(tl_tot["lines"] or 0)), f1(tl_tot["amount"]), "100%"])
    table(s, 6.95, 4.45, 6.1, ["Tracking-log status", "Lines", "MMAD", "%"], rows, col_w=[3.6, 0.7, 0.9, 0.7], font_size=8.5, header_size=8, row_h=0.22, header_h=0.3, total_last=True)
    pri = [r for r in data["ONGOING"] if str(r["bucket"]).startswith("Priority list")]
    pri_tot = next((r["amount"] for r in pri if "total approved" in str(r["bucket"])), None)
    pri_cls = [r for r in pri if str(r["bucket"]).startswith("Priority list – P")]
    if pri_tot:
        tb(s, 6.95, 4.45 + 0.3 + 0.22 * len(rows) + 0.1, 6.1, 0.7, [[("Priority list (02.10): ", {"bold": True, "color": NAVY}), (f"{f0(pri_tot)} MMAD approved for payment — " + " · ".join(f"{str(r['bucket']).replace('Priority list – ','')} {f0(r['amount'])}" for r in sorted(pri_cls, key=lambda r: str(r['bucket']))), {})]], size=9, line_spacing=1.1)

    # ------------------------------------------------------------------ 12. Accounts payable
    apb = data["AP_BU"]; bk = data["AP_BUCKETS"]
    s = D.slide("Accounts Payable – supplier payables & LC", f"Payables incl. LC of {f0(ap_tot['current'])} MMAD ({fd(ap_tot['current']-ap_tot['prev_week'])} WoW): parked {f0(ap_pk['current'])}, 4411 {f0(ap_44['current'])}, downpayments {f0(ap_dp['current'])}, LC {f0(ap_lc['current'])}")
    cw = (SW - 0.6 - 0.12 * 4) / 5
    for i, (lab, r) in enumerate([("Total payables incl. LC", ap_tot), ("Parked invoices", ap_pk), ("Invoices 4411 (open items)", ap_44), ("Downpayments", ap_dp), ("LC opened (bank situation)", ap_lc)]):
        kpi_card(s, 0.3 + i * (cw + 0.12), 1.25, cw, 1.05, lab, r["current"], delta_wow=r["current"] - r["prev_week"], good_when_negative=True)
    tb(s, 0.4, 2.45, 6.5, 0.3, "Payables by business unit – this week and Δ WoW (MMAD)", size=13, bold=True, color=NAVY)
    rows, colors = [], {}
    apbs = sorted([r for r in apb if sum((r[k] or 0) for k in ("downpayment_cur", "parked_cur", "inv4411_cur", "lc_cur")) > 0.05], key=lambda r: -sum((r[k] or 0) for k in ("downpayment_cur", "parked_cur", "inv4411_cur", "lc_cur")))
    for i, r in enumerate(apbs):
        cur = sum((r[k] or 0) for k in ("downpayment_cur", "parked_cur", "inv4411_cur", "lc_cur")); prev = sum((r[k] or 0) for k in ("downpayment_prev", "parked_prev", "inv4411_prev", "lc_prev"))
        rows.append([short(r["bu"]), f0(r["downpayment_cur"]), f0(r["parked_cur"]), f0(r["inv4411_cur"]), f0(r["lc_cur"]), f0(cur), fd(cur - prev)]); colors[(i, 6)] = delta_color(cur - prev)
    rows.append(["Total", f0(ap_dp["current"]), f0(ap_pk["current"]), f0(ap_44["current"]), f0(ap_lc["current"]), f0(ap_tot["current"]), fd(ap_tot["current"] - ap_tot["prev_week"])])
    table(s, 0.4, 2.75, 6.5, ["BU", "Down-payment", "Parked", "4411", "LC", "Total", "Δ WoW"], rows, col_w=[1.5, 0.9, 0.8, 0.8, 0.8, 0.85, 0.85], font_size=9.5, header_size=8.5, row_h=0.33, header_h=0.36, colors=colors, total_last=True)
    tb(s, 0.4, 2.75 + 0.36 + 0.33 * len(rows) + 0.1, 6.5, 0.5, "LC = outstanding on opened LCs; parked = invoices parked in SAP pending validation; 4411 = posted supplier open items; downpayments = advance invoices due to suppliers.", size=8, italic=True, color=GREY)
    # right: due date buckets from the invoice file
    origins = ["Invoices - 4411", "Parked Invoice", "Downpayment"]; buckets = ["Overdue", "Due ≤ 30d", "Due 31–60d", "Due > 60d", "No due date"]
    bm = {(r["origin"], r["bucket"]): r["amount"] or 0 for r in bk}
    buckets = [b for b in buckets if any(bm.get((o, b), 0) > 0.05 for o in origins)]
    ap_file_tot = sum(bm.values()); overdue = sum(bm.get((o, "Overdue"), 0) for o in origins)
    tb(s, 7.2, 2.45, 5.9, 0.3, f"Open supplier invoices by due date – AP file {meta['ap_label']} (MMAD)", size=13, bold=True, color=NAVY)
    tb(s, 7.2, 2.75, 5.9, 0.3, f"{f0(ap_file_tot)} MMAD of open lines, of which {f0(overdue)} MMAD ({overdue/ap_file_tot*100:.0f}%) past due at {asof}", size=9.5, color=SLATE)
    ch = bar_chart(s, 7.1, 3.0, 6.0, 2.45, buckets, [(o.replace("Invoices - ", ""), [bm.get((o, b), 0) for b in buckets]) for o in origins], [NAVY2, BLUE, GOLD], stacked=True, gap=50, cat_size=8.5, label_size=8, legend_pos=XL_LEGEND_POSITION.BOTTOM, label_colors=[WHITE, WHITE, DARK])
    hide_small_labels(ch, [[bm.get((o, b), 0) for b in buckets] for o in origins], 30)
    tiers = sorted(data["AP_TIERS"], key=lambda r: -(r["amount"] or 0))
    def tier_group(t):
        t = str(t).lower()
        if "international" in t: return "International"
        if "national" in t: return "National (tiers 1–3)"
        return "Other / not tiered"
    tg = {}
    for r in tiers: tg[tier_group(r["tier"])] = tg.get(tier_group(r["tier"]), 0) + (r["amount"] or 0)
    tb(s, 7.2, 5.45, 5.9, 0.3, "By supplier tier (MMAD)", size=11, bold=True, color=NAVY)
    rows = [[k, f0(v), f"{v/ap_file_tot*100:.0f}%"] for k, v in sorted(tg.items(), key=lambda x: -x[1])]
    table(s, 7.2, 5.72, 5.9, ["Tier group", "MMAD", "%"], rows, col_w=[3, 1.4, 1.4], font_size=8.5, header_size=8, row_h=0.22, header_h=0.26)
    tb(s, 7.2, 5.72 + 0.26 + 0.22 * len(rows) + 0.04, 5.9, 0.4, "Due-date buckets follow the invoice due date; lines without a due date are not bucketed.", size=7.5, italic=True, color=GREY)

    # ------------------------------------------------------------------ 13. Letters of credit
    lcl = data["LC_LINES"]; months = [k for k in lcl[0].keys() if str(k).startswith("202")]
    lc_all = sum(r["encours_mad"] or 0 for r in lcl); lc_open = sum((r["encours_mad"] or 0) for r in lcl if r["opened"] == "Yes"); lc_noo = lc_all - lc_open
    mt = {m: sum((r[m] or 0) for r in lcl) for m in months}
    p26 = sum(v for m, v in mt.items() if str(m).startswith("2026")); p27 = sum(v for m, v in mt.items() if str(m).startswith("2027"))
    s = D.slide("Letters of Credit – outstanding & maturities", f"{f0(lc_all)} MMAD of LC commitments at {meta['lc_date']}: {f0(lc_open)} opened, {f0(lc_noo)} still to be opened; {f0(appr_lc)} MMAD maturing in {month_name}")
    cw = (SW - 0.6 - 0.12 * 4) / 5
    card(s, 0.3, 1.25, cw, 1.05, "Total LC commitments", f0(lc_all), "Encours in MMAD, bank LC situation", fill=NAVY, value_size=22)
    card(s, 0.3 + (cw + 0.12), 1.25, cw, 1.05, "Opened LCs", f0(lc_open), "LC reference and opening date exist", fill=NAVY2, value_size=22)
    card(s, 0.3 + 2 * (cw + 0.12), 1.25, cw, 1.05, "To be opened", f0(lc_noo), "Committed, LC not yet issued", fill=SLATE, value_size=22)
    card(s, 0.3 + 3 * (cw + 0.12), 1.25, cw, 1.05, f"{month_name} maturities", f0(appr_lc), "Counted as approved payments", fill=GOLD, value_size=22)
    card(s, 0.3 + 4 * (cw + 0.12), 1.25, cw, 1.05, "2026 / 2027 maturities", f"{f0(p26)} / {f0(p27)}", f"Per monthly schedule; unscheduled {f0(lc_all - p26 - p27)}", fill=TEAL, value_size=20)
    tb(s, 0.4, 2.45, 7.5, 0.3, "LC maturity schedule by month (MMAD)", size=13, bold=True, color=NAVY)
    last_nz = max(i for i, m in enumerate(months) if mt[m] > 0.05); mkeep = months[:last_nz + 1]
    bar_chart(s, 0.3, 2.75, 7.6, 2.5, [dt.date.fromisoformat(m + "-01").strftime("%b-%y") for m in mkeep], [("Maturities", [mt[m] for m in mkeep])], [NAVY2], legend=False, gap=40, cat_size=8, label_size=8)
    # by BU / project
    byp = {}
    for r in lcl:
        k = r["bu"]; d = byp.setdefault(k, {"open": 0, "noo": 0, "oct": 0})
        d["open" if r["opened"] == "Yes" else "noo"] += r["encours_mad"] or 0; d["oct"] += r[months[0]] or 0
    rows = [[short(k), f0(v["open"]), f0(v["noo"]), f0(v["open"] + v["noo"]), f0(v["oct"])] for k, v in sorted(byp.items(), key=lambda x: -(x[1]["open"] + x[1]["noo"]))]
    rows.append(["Total", f0(lc_open), f0(lc_noo), f0(lc_all), f0(mt[months[0]])])
    tb(s, 0.4, 5.3, 7.5, 0.3, "By business unit (MMAD)", size=11, bold=True, color=NAVY)
    table(s, 0.4, 5.58, 7.5, ["BU", "Opened", "To be opened", "Total", f"{dt.date.fromisoformat(months[0]+'-01').strftime('%b-%y')} maturity"], rows, col_w=[2, 1.2, 1.3, 1.2, 1.6], font_size=8.5, header_size=8, row_h=0.2, header_h=0.26, total_last=True)
    # by action
    act = {}
    for r in lcl:
        a = short_action(r["action"]); act[a] = act.get(a, 0) + (r["encours_mad"] or 0)
    acts = sorted([(k, v) for k, v in act.items() if v > 0.05], key=lambda x: -x[1])
    if len(acts) > 8: acts = acts[:8] + [("Other actions", sum(v for _, v in acts[8:]))]
    tb(s, 8.2, 2.45, 4.9, 0.3, "By action status (MMAD)", size=13, bold=True, color=NAVY)
    bar_chart(s, 8.1, 2.75, 5.0, 4.2, [k for k, _ in acts], [("Encours", [v for _, v in acts])], [TEAL], horizontal=True, legend=False, gap=30, cat_size=7.5, label_size=8)

    # ------------------------------------------------------------------ 7b. Zoom per business unit
    det = {}
    for r in data.get("AR_BU_DETAIL", []): det.setdefault(r["bu"], {}).setdefault(r["section"], []).append(r)
    projs = {}
    for r in data.get("AR_BU_PROJECTS", []): projs.setdefault((r["bu"], r["type"]), []).append(r)
    urt = {r["bu"]: r for r in data["UNRECOGNIZED"]}
    att = {r["bu"]: r for r in data.get("UR_ATTACHMENT_STATUS", [])}
    has_prev_detail = any(r["prev_week"] is not None for r in data.get("AR_BU_DETAIL", []) if r["section"] != "owner")
    zoom_bus = [b for b in BU_ORDER if b in arbu and (arbu[b]["Total AR"]["current"] or 0) >= 20]
    for b in zoom_bus:
        t_, u_, p_, a_ = arbu[b]["Total AR"], arbu[b]["Unbilled"], arbu[b]["Unpaid"], arbu[b]["Advance"]
        unr = arbu[b].get("Unrecognized") or {"current": 0, "prev_week": None, "month_start": None}
        unr_c = unr["current"] or 0; unr_p = unr["prev_week"]; unr_m = unr.get("month_start")
        tc = t_["current"] or 1
        s = D.slide(f"Zoom – {b}", f"Total AR {f0(t_['current'])} MMAD ({fd(dwow(t_))} WoW, {fd(dms(t_))} vs month start); unbilled {u_['current']/tc*100:.0f}% / unpaid {p_['current']/tc*100:.0f}% / advance {a_['current']/tc*100:.0f}%" + (f" · DSO {dso_txt(b)}" if dso_txt(b) else ""))
        cw = (SW - 0.6 - 0.12 * 5) / 6
        kpi_card(s, 0.3, 1.2, cw, 1.15, "Gross exposure", t_["current"] + unr_c, delta_wow=(dwow(t_) + (unr_c - unr_p)) if unr_p is not None else None, value_size=20)
        kpi_card(s, 0.3 + (cw + 0.12), 1.2, cw, 1.15, "Unrecognized revenue", unr_c, delta_wow=(round(unr_c) - round(unr_p)) if unr_p is not None else None, delta_ms=(round(unr_c) - round(unr_m)) if unr_m is not None else None, value_size=20, ms_label="vs month start")
        for i, (lab, r) in enumerate([("Total AR", t_), ("Unbilled", u_), ("Unpaid", p_), ("Advance", a_)], start=2):
            kpi_card(s, 0.3 + i * (cw + 0.12), 1.2, cw, 1.15, lab, r["current"], delta_wow=dwow(r), delta_ms=dms(r), value_size=20, ms_label="vs month start")
        # unbilled by ownership (doughnut + legend)
        own_rows = det.get(b, {}).get("owner", [])
        own = {r["item"]: (r["current"] or 0) for r in own_rows}
        tb(s, 0.4, 2.5, 4.4, 0.3, "Unbilled split by ownership", size=12, bold=True, color=NAVY)
        cats = [("Under OCP", NAVY2), ("SS postponed", TEAL), ("Under JESA", LBLUE), ("At risk", RED)]
        vals = [max(own.get(c, 0), 0) for c, _ in cats]
        if sum(vals) > 0: doughnut(s, 0.2, 2.75, 2.3, 2.35, [c for c, _ in cats], vals, [col for _, col in cats], hole=50)
        ub = u_["current"] or 1
        shown = [(c, col) for c, col in cats if own.get(c, 0) > 0.5 or own.get(c, 0) < -0.5]
        for i, (c, col) in enumerate(shown):
            yy = 2.85 + i * 0.55
            rect(s, 2.6, yy + 0.06, 0.16, 0.16, col, line=SLATE)
            r0 = next((r for r in own_rows if r["item"] == c), None)
            dw = ((r0["current"] or 0) - (r0["prev_week"] or 0)) if (r0 and r0["prev_week"] is not None) else None
            tb(s, 2.82, yy - 0.02, 2.0, 0.26, c, size=9.5, bold=True, color=NAVY)
            tb(s, 2.82, yy + 0.2, 2.1, 0.26, [[(f"{f0(own.get(c, 0))} · {own.get(c, 0)/ub*100:.0f}%", {"size": 8.5, "color": SLATE})] + ([("   " + fd(dw), {"size": 8.5, "bold": True, "color": delta_color(dw)})] if dw is not None else [])], size=8.5)
        # unbilled by status table
        st_rows = sorted([r for r in det.get(b, {}).get("unbilled_status", []) if abs(r["current"] or 0) >= 0.5 or abs(r["month_start"] or 0) >= 0.5], key=lambda r: -abs(r["current"] or 0))[:8]
        tb(s, 5.0, 2.5, 4.2, 0.3, "Largest unbilled statuses (MMAD)", size=12, bold=True, color=NAVY)
        rows, colors = [], {}
        for i, r in enumerate(st_rows):
            lab = str(r["item"]); lab = (lab[:46] + "…") if len(lab) > 47 else lab
            dw = (round(r["current"] or 0) - round(r["prev_week"] or 0)) if r["prev_week"] is not None else None
            dm = (round(r["current"] or 0) - round(r["month_start"] or 0)) if r["month_start"] is not None else None
            row = [lab, f0(r["current"]), f"{(r['current'] or 0)/ub*100:.0f}%"] + ([fd(dw)] if has_prev_detail else []) + [fd(dm) if dm is not None else "n/a"]
            rows.append(row)
            if has_prev_detail: colors[(i, 3)] = delta_color(dw); colors[(i, 4)] = delta_color(dm)
            else: colors[(i, 3)] = delta_color(dm)
        hdrs = ["Status", "MMAD", "%"] + (["Δ WoW"] if has_prev_detail else []) + ["Δ vs MS"]
        cw_ = [2.55, 0.55, 0.45] + ([0.55] if has_prev_detail else []) + [0.6]
        table(s, 5.0, 2.8, 4.2, hdrs, rows, col_w=cw_, font_size=7.5, header_size=7.5, row_h=0.235, header_h=0.26, colors=colors, align=[PP_ALIGN.LEFT] + [PP_ALIGN.RIGHT] * (len(hdrs) - 1))
        # unpaid split + phase
        ug = {r["item"]: (r["current"] or 0) for r in det.get(b, {}).get("unpaid_group", [])}
        tb(s, 9.45, 2.5, 3.7, 0.3, "Unpaid split", size=12, bold=True, color=NAVY)
        ucats = [("Past due", NAVY2), ("Rejected", LBLUE), ("Current", GREY)]
        uvals = [max(ug.get(c, 0), 0) for c, _ in ucats]; upb = p_["current"] or 1
        if sum(uvals) > 0: doughnut(s, 9.3, 2.75, 1.9, 1.9, [c for c, _ in ucats], uvals, [col for _, col in ucats], hole=50)
        for i, (c, col) in enumerate([(c, col) for c, col in ucats if ug.get(c, 0) > 0.5]):
            yy = 2.9 + i * 0.5
            rect(s, 11.3, yy + 0.05, 0.16, 0.16, col, line=SLATE)
            tb(s, 11.5, yy - 0.03, 1.7, 0.26, c, size=9, bold=True, color=NAVY)
            tb(s, 11.5, yy + 0.19, 1.7, 0.24, f"{f0(ug.get(c, 0))} · {ug.get(c, 0)/upb*100:.0f}%", size=8, color=SLATE)
        ph = sorted([r for r in det.get(b, {}).get("unpaid_phase", []) if (r["current"] or 0) > 0.05], key=lambda r: -(r["current"] or 0))[:6]
        if ph:
            bar_chart(s, 9.3, 4.65, 3.85, min(2.45, 0.6 + 0.38 * len(ph)), [((str(r["item"])[:19] + "…") if len(str(r["item"])) > 20 else str(r["item"])) for r in ph], [("Unpaid", [r["current"] for r in ph])], [NAVY2], horizontal=True, legend=False, gap=30, cat_size=7.5, label_size=7.5, title="Unpaid by phase (MMAD)")
        # variance table (bottom left)
        tb(s, 0.4, 5.15, 4.5, 0.3, "Movement (MMAD)", size=12, bold=True, color=NAVY)
        rows, colors = [], {}
        comps = [("Advance", a_), ("Unbilled", u_), ("Unpaid", p_), ("Total AR", t_)]
        for i, (lab, r) in enumerate(comps):
            rows.append([lab, f0(r["month_start"]), f0(r["prev_week"]), f0(r["current"]), fd(dwow(r)), fd(dms(r))]); colors[(i, 4)] = delta_color(dwow(r)); colors[(i, 5)] = delta_color(dms(r))
        rows.append(["Unrecognized", f0(unr_m) if unr_m is not None else "n/a", f0(unr_p) if unr_p is not None else "n/a", f0(unr_c), fd(round(unr_c) - round(unr_p)) if unr_p is not None else "n/a", fd(round(unr_c) - round(unr_m)) if unr_m is not None else "n/a"])
        colors[(4, 4)] = delta_color((unr_c - unr_p) if unr_p is not None else None); colors[(4, 5)] = delta_color((unr_c - unr_m) if unr_m is not None else None)
        table(s, 0.4, 5.45, 4.5, ["", f"M. start\n{ms_s}", prev_s, cur_s, "Δ WoW", "Δ vs MS"], rows, col_w=[1.1, 0.75, 0.65, 0.65, 0.65, 0.7], font_size=8, header_size=7.5, row_h=0.235, header_h=0.3, colors=colors, bold_rows=(3,))
        # aging buckets (bottom middle)
        agu = {r["item"]: (r["current"] or 0) for r in det.get(b, {}).get("aging_unbilled", [])}; agp = {r["item"]: (r["current"] or 0) for r in det.get(b, {}).get("aging_unpaid", [])}
        bks = ["< 30 days", "31–60 days", "61–90 days", "91–120 days", "> 120 days"]
        bar_chart(s, 5.0, 5.05, 4.2, 2.05, bks, [("Unbilled", [agu.get(k, 0) for k in bks]), ("Unpaid", [agp.get(k, 0) for k in bks])], [NAVY2, GOLD], gap=50, cat_size=7, label_size=7, title="Aging – unbilled vs unpaid (MMAD)", legend_pos=XL_LEGEND_POSITION.BOTTOM)

        # ---- page B: projects & unrecognized
        s = D.slide(f"Zoom – {short(b)} · projects & unrecognized", f"{b}: largest unbilled and unpaid projects at {cur_s}; unrecognized supplier scope by status (incl. markup)")
        for j, (typ, x0) in enumerate((("Unbilled", 0.4), ("Unpaid", 6.75))):
            lst = sorted(projs.get((b, typ), []), key=lambda r: -(r["amount"] or 0))[:8]
            tot_t = (u_ if typ == "Unbilled" else p_)["current"] or 1
            tb(s, x0, 1.2, 6.0, 0.3, f"Top {typ.lower()} projects (MMAD)", size=12, bold=True, color=NAVY)
            def short_status(st):
                st = str(st or ""); return (st[:34] + "…") if len(st) > 35 else st
            rows = [[(str(r["wbs"]) or "")[:13], ((str(r["project"]) or "")[:33] + "…") if len(str(r["project"] or "")) > 34 else str(r["project"] or ""), short_status(r["status"]), f1(r["amount"]), f"{(r['amount'] or 0)/tot_t*100:.0f}%"] for r in lst]
            if rows: table(s, x0, 1.5, 6.15, ["WBS", "Project", "Status", "MMAD", ("% unb." if typ == "Unbilled" else "% unp.")], rows, col_w=[0.95, 2.25, 2.05, 0.5, 0.45], font_size=7.5, header_size=7.5, row_h=0.27, header_h=0.32, align=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])
        u = urt.get(b)
        if u and (u["grand_total"] or 0) > 0.05:
            tb(s, 0.4, 4.2, 6.0, 0.3, f"Unrecognized suppliers – {f0(u['grand_total'])} MMAD" + (f" ({fd(round(u['grand_total'] or 0) - round(u['prev_week_total'] or 0))} WoW)" if u.get("prev_week_total") is not None else ""), size=12, bold=True, color=NAVY)
            cats = ["Pending approval", "Signed by client", "Pending submission", "Under CN", "At risk"]
            eq = [u[k] or 0 for k in ("eq_pending_approval", "eq_signed_by_client", "eq_pending_submission", "eq_under_cn", "eq_at_risk")]
            sb = [u[k] or 0 for k in ("sub_pending_approval", "sub_signed_by_client", "sub_pending_submission", "sub_under_cn", "sub_at_risk")]
            ch = bar_chart(s, 0.3, 4.5, 6.2, 2.6, cats, [("Equipment", eq), ("Subcontracts", sb)], [NAVY2, BLUE], stacked=True, gap=50, cat_size=8, label_size=8, label_colors=[WHITE, WHITE], legend_pos=XL_LEGEND_POSITION.BOTTOM)
            hide_small_labels(ch, [eq, sb], max(10, (u["grand_total"] or 0) * 0.05))
            at = att.get(b)
            if at and (at["total"] or 0) > 0.05:
                tb(s, 6.75, 4.2, 6.2, 0.3, f"JESA attachments to client – status ({f0(at['total'])} MMAD)", size=12, bold=True, color=NAVY)
                acats = ["Under JESA approval", "Under JESA preparation", "Under client approval", "Signed by client", "Accrual posted"]
                aeq = [at[k] or 0 for k in ("eq_under_jesa_approval", "eq_under_jesa_preparation", "eq_under_client_approval", "eq_signed_by_client", "eq_accrual_posted")]
                asb = [at[k] or 0 for k in ("sub_under_jesa_approval", "sub_under_jesa_preparation", "sub_under_client_approval", "sub_signed_by_client", "sub_accrual_posted")]
                ch = bar_chart(s, 6.65, 4.5, 6.4, 2.6, acats, [("Equipment", aeq), ("Subcontracts", asb)], [NAVY2, BLUE], stacked=True, gap=50, cat_size=8, label_size=8, label_colors=[WHITE, WHITE], legend_pos=XL_LEGEND_POSITION.BOTTOM)
                hide_small_labels(ch, [aeq, asb], max(10, (at["total"] or 0) * 0.05))
        else:
            tb(s, 0.4, 4.3, 12, 0.4, "No unrecognized supplier scope reported for this business unit.", size=10, italic=True, color=SLATE)

    # ------------------------------------------------------------------ 14. Backup: LC detail
    s = D.slide("Backup – Letters of Credit detail", f"Largest LC lines by outstanding amount (bank LC situation {meta['lc_date']}, MMAD)")
    top = sorted(lcl, key=lambda r: -(r["encours_mad"] or 0))[:18]
    rows = [[r["project"], (r["supplier"] or "")[:34], r["lc_ref"] or "not opened", datef(r["expiry"], "%d-%b-%y") if r["expiry"] and str(r["expiry"])[:2] == "20" else (r["expiry"] or "n/a"), r["ccy"], f1(r["encours_mad"]), r["opened"], f0(r["part_2026"]), f0(r["part_2027"]), short_action(r["action"])] for r in top]
    table(s, 0.35, 1.25, 12.65, ["Project", "Supplier", "LC reference", "Expiry", "Ccy", "Outstanding", "Opened", "2026", "2027", "Action"], rows, col_w=[1.1, 2.7, 1.5, 0.95, 0.5, 0.85, 0.7, 0.7, 0.7, 2.85], font_size=8, header_size=8, row_h=0.27, header_h=0.3,
          align=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.CENTER, PP_ALIGN.RIGHT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.LEFT])

    # ------------------------------------------------------------------ 15. Backup: AP detail
    s = D.slide("Backup – Accounts Payable detail", f"Open supplier invoices by BU and origin, and top suppliers (AP file {meta['ap_label']}, MMAD)")
    abo = data["AP_BU_ORIGIN"]; m = {}
    for r in abo: m.setdefault(r["bu"], {})[r["origin"]] = (r["amount"] or 0, int(r["lines"] or 0))
    rows = []
    for b, d in sorted(m.items(), key=lambda x: -sum(v[0] for v in x[1].values())):
        rows.append([short(b)] + [f0(d.get(o, (0, 0))[0]) for o in origins] + [f0(sum(v[0] for v in d.values())), f"{sum(v[1] for v in d.values()):,}"])
    rows.append(["Total"] + [f0(sum(d.get(o, (0, 0))[0] for d in m.values())) for o in origins] + [f0(ap_file_tot), f"{sum(v[1] for d in m.values() for v in d.values()):,}"])
    tb(s, 0.4, 1.25, 6, 0.3, "By business unit and origin", size=12, bold=True, color=NAVY)
    table(s, 0.4, 1.55, 6.0, ["BU", "4411", "Parked", "Downpayment", "Total", "Lines"], rows, col_w=[1.7, 0.9, 0.9, 1.0, 0.9, 0.7], font_size=8.5, header_size=8, row_h=0.26, header_h=0.3, total_last=True)
    tops = data["AP_TOP_SUPPLIERS"]
    tb(s, 6.8, 1.25, 6, 0.3, "Top 15 suppliers by open payables", size=12, bold=True, color=NAVY)
    rows = [[str(i + 1), (r["supplier"] or "")[:34], short(r["bu"]), f1(r["amount"]), f"{(r['amount'] or 0)/ap_file_tot*100:.1f}%"] for i, r in enumerate(tops)]
    table(s, 6.8, 1.55, 6.2, ["#", "Supplier", "BU", "MMAD", "% of AP"], rows, col_w=[0.4, 3.0, 1.4, 0.8, 0.8], font_size=8.5, header_size=8, row_h=0.26, header_h=0.3, align=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])
    bar_chart(s, 0.3, 1.55 + 0.3 + 0.26 * len(m) + 0.5, 6.2, SH - 0.5 - (1.55 + 0.3 + 0.26 * len(m) + 0.5), [b for b in buckets], [(o.replace("Invoices - ", ""), [bm.get((o, b), 0) for b in buckets]) for o in origins], [NAVY2, BLUE, GOLD], stacked=False, gap=60, cat_size=8, label_size=7.5, title="Due-date buckets by origin (MMAD)", legend_pos=XL_LEGEND_POSITION.BOTTOM)

    # ------------------------------------------------------------------ 16. Sources & definitions
    s = D.slide("Sources & definitions", "Where each figure comes from and how it is defined")
    rows = [
        ["Accounts receivable (advance, unbilled, unpaid), statuses, aging, non-OCP, JESA responsibility, collections", "AR extract of the week", f"{cur_d}; previous week {prev_d}; month start {ms_d}"],
        ["Unrecognized revenue by business unit and status, supplier PO book", "UR tracking consolidation", f"{asof}; previous week and month start from the earlier reviews"],
        ["DSO excluding advances, group and by business unit", "BU meeting report", "October 2026 (prior month: September)"],
        ["Cash position, October paid / approved / pending payments, priority list", "Cash flow analysis by business unit", f"{datef(meta['cash_data_as_of'])}; collections to {meta['collections_to']}, payments to {meta['payments_to']}"],
        ["Supplier payables by business unit, due date, tier and supplier", "Supplier invoice extract (4411, parked, downpayments)", f"AP extract {meta['ap_label']}"],
        ["Letters of credit: outstanding, maturity schedule, action status", "Bank LC situation", f"LC situation {meta['lc_date']}"],
    ]
    table(s, 0.4, 1.25, 12.5, ["Figures", "Source", "As of"], rows, col_w=[5.6, 3.6, 3.3], font_size=9, header_size=9, row_h=0.42, header_h=0.32, align=[PP_ALIGN.LEFT] * 3)
    defs = [("Total AR", "advance payments + unbilled + unpaid, amounts HT in MMAD."),
            ("Unbilled", "delivered scope not yet invoiced, excluding unearned amounts on lump-sum contracts; split by ownership of the next action (client / OCP, JESA, service sheets postponed, at risk)."),
            ("Unpaid", "invoiced, not yet collected; split current / past due / rejected and by OCP processing phase."),
            ("Unrecognized revenue", "supplier and equipment scope executed but not yet booked as AR, including the 6% JESA cost + fee."),
            ("Gross exposure", "total AR + unrecognized revenue."),
            ("Week on week / vs month start", f"position at {cur_s} compared with {prev_s} and with the first AR position of the month ({ms_s})."),
            ("Payments received", "invoices paid by payment date, on the amount-invoiced basis used in the BU meeting report."),
            ("Cash position", "procurement (non-labor) cash in and out for the JESA Group including AM; PSR (labor) collections shown net of the AM ring-fence and the Water & Environment reserve."),
            ("DSO", "days sales outstanding excluding advances, as reported in the BU meeting report.")]
    tb(s, 0.4, 4.15, 12.5, 2.9, [[(t + ": ", {"bold": True, "color": NAVY}), (d_, {})] for t, d_ in defs], size=9, color=DARK, line_spacing=1.1)

    D.save(out)
    print("saved", out, "slides:", D.n)

if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--inputs", default="inputs/weekly_inputs.xlsx"); ap.add_argument("--out", required=True)
    a = ap.parse_args()
    data, meta = read_inputs(a.inputs)
    build(data, meta, a.out)
