"""
AI-powered Executive Report generator for myCompetency Scorecard.
Dashboard-style PPTX with native charts, matching the Streamlit scorecard layout.

Credentials loaded from .env:
    LLM_API_KEY / AZURE_OPENAI_API_KEY
    AZURE_OPENAI_ENDPOINT
    AZURE_OPENAI_DEPLOYMENT
"""

import io
import json
import os
from datetime import datetime

from dotenv import load_dotenv
from lxml import etree
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt, Emu

load_dotenv()

# ---------------------------------------------------------------------------
# Brand palette — matching the Streamlit dashboard
# ---------------------------------------------------------------------------
NAVY   = RGBColor(0x1E, 0x27, 0x61)
PURPLE = RGBColor(0x4B, 0x2D, 0x7F)
ICE    = RGBColor(0xCA, 0xDC, 0xFC)
TEAL   = RGBColor(0x02, 0x80, 0x90)
ORANGE = RGBColor(0xE5, 0x5C, 0x2A)
GREEN  = RGBColor(0x22, 0xC5, 0x5E)
RED    = RGBColor(0xEF, 0x44, 0x44)
AMBER  = RGBColor(0xFA, 0xCC, 0x15)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
OFFWHT = RGBColor(0xF5, 0xF7, 0xFA)
GREY   = RGBColor(0x6B, 0x7B, 0x8D)
LTGREY = RGBColor(0xD9, 0xDE, 0xE6)

PROF_HEX = {
    "P0": "EF4444",
    "P1": "F97316",
    "P2": "FACC15",
    "P3": "22C55E",
    "Expert eligible": "16A34A",
    "Unspecified": "9CA3AF",
    "No Assessment": "E55C2A",
}

# Slide canvas — 16:9 widescreen
W = Inches(13.33)
H = Inches(7.5)


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------

def _rect(slide, x, y, w, h, fill: RGBColor):
    shape = slide.shapes.add_shape(1, x, y, w, h)
    shape.line.fill.background()
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    return shape


def _txt(slide, text, x, y, w, h, size, bold=False, color=WHITE,
         align=PP_ALIGN.LEFT, wrap=True, italic=False):
    txb = slide.shapes.add_textbox(x, y, w, h)
    txb.text_frame.word_wrap = wrap
    txb.text_frame.margin_left   = Emu(0)
    txb.text_frame.margin_right  = Emu(0)
    txb.text_frame.margin_top    = Emu(0)
    txb.text_frame.margin_bottom = Emu(0)
    p = txb.text_frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text)
    run.font.size   = Pt(size)
    run.font.bold   = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txb


def _add_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _pct_color(pct):
    if pct >= 80:
        return TEAL
    if pct >= 60:
        return ORANGE
    return RED


def _hex_to_rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _set_donut_point_colors(chart, color_list: list[str]):
    """Set individual slice colors on the first series of a donut/pie chart."""
    ns_c = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    series_elem = chart.series[0]._element
    for i, hex_color in enumerate(color_list):
        dp = etree.fromstring(
            f'<c:dPt xmlns:c="{ns_c}" xmlns:a="{ns_a}">'
            f'<c:idx val="{i}"/><c:bubble3D val="0"/>'
            f'<c:spPr><a:solidFill><a:srgbClr val="{hex_color}"/></a:solidFill>'
            f'<a:ln><a:noFill/></a:ln></c:spPr>'
            f'</c:dPt>'
        )
        series_elem.append(dp)


def _set_series_color(series, hex_color: str):
    """Set the fill color of a chart series."""
    ns_c = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    elem = series._element
    sp_pr = etree.fromstring(
        f'<c:spPr xmlns:c="{ns_c}" xmlns:a="{ns_a}">'
        f'<a:solidFill><a:srgbClr val="{hex_color}"/></a:solidFill>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</c:spPr>'
    )
    elem.append(sp_pr)


def _hide_chart_border(chart):
    """Remove the border box around the chart plot area."""
    try:
        from pptx.oxml.ns import qn
        plot_area = chart._element.find(
            ".//{http://schemas.openxmlformats.org/drawingml/2006/chart}plotArea"
        )
        if plot_area is not None:
            sp_pr = etree.SubElement(
                plot_area,
                "{http://schemas.openxmlformats.org/drawingml/2006/chart}spPr"
            )
            ln = etree.SubElement(
                sp_pr,
                "{http://schemas.openxmlformats.org/drawingml/2006/main}ln"
            )
            etree.SubElement(
                ln,
                "{http://schemas.openxmlformats.org/drawingml/2006/main}noFill"
            )
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Slide 1 — Title
# ---------------------------------------------------------------------------

def _slide_title(prs, business_group: str, report_date: str):
    slide = _add_slide(prs)
    _rect(slide, 0, 0, W, H, NAVY)
    _rect(slide, 0, H - Inches(0.5), W, Inches(0.5), TEAL)
    _rect(slide, Inches(0.6), Inches(0.55), Inches(0.16), Inches(0.42), TEAL)

    _txt(slide, "myCompetency",
         Inches(0.85), Inches(0.45), Inches(8), Inches(0.55),
         size=13, color=ICE)
    _txt(slide, "Scorecard\nExecutive Report",
         Inches(0.6), Inches(1.4), Inches(11), Inches(2.2),
         size=52, bold=True, color=WHITE, wrap=True)
    _txt(slide, business_group,
         Inches(0.6), Inches(3.7), Inches(10), Inches(0.65),
         size=22, color=ICE)
    _txt(slide, f"Report generated: {report_date}",
         Inches(0.6), Inches(4.45), Inches(9), Inches(0.45),
         size=14, color=GREY)
    _txt(slide, "CONFIDENTIAL  ·  Internal Use Only",
         Inches(0.6), H - Inches(0.48), W - Inches(1.2), Inches(0.42),
         size=10, color=WHITE, align=PP_ALIGN.CENTER)


# ---------------------------------------------------------------------------
# Slide 2 — Executive Dashboard (mirrors the Streamlit layout)
# ---------------------------------------------------------------------------

def _kpi_card_6(slide, label, value, sub, x, y, w, h, accent: RGBColor,
                delta: str = "", delta_positive: bool = True):
    _rect(slide, x, y, w, h, OFFWHT)
    _rect(slide, x, y, w, Inches(0.06), accent)
    _txt(slide, str(value),
         x + Inches(0.1), y + Inches(0.1), w - Inches(0.2), Inches(0.62),
         size=32, bold=True, color=accent, align=PP_ALIGN.CENTER)
    _txt(slide, label,
         x + Inches(0.08), y + Inches(0.72), w - Inches(0.16), Inches(0.26),
         size=9, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    if sub:
        _txt(slide, sub,
             x + Inches(0.08), y + Inches(0.98), w - Inches(0.16), Inches(0.20),
             size=8, color=GREY, align=PP_ALIGN.CENTER)
    if delta:
        delta_color = GREEN if delta_positive else RED
        _txt(slide, delta,
             x + Inches(0.08), y + Inches(1.18), w - Inches(0.16), Inches(0.20),
             size=7.5, bold=True, color=delta_color, align=PP_ALIGN.CENTER)


def _slide_dashboard(prs, metrics: dict, proficiency_rows: list, career_rows: list,
                     career_kpis: dict, insights: dict):
    slide = _add_slide(prs)

    # ── Header ────────────────────────────────────────────────────────────────
    _rect(slide, 0, 0, W, Inches(1.05), NAVY)
    _txt(slide, metrics.get("business_group", ""),
         Inches(0.5), Inches(0.08), Inches(12), Inches(0.4),
         size=11, color=ICE)
    _txt(slide, "Primary Skill Scorecard",
         Inches(0.5), Inches(0.42), Inches(12), Inches(0.52),
         size=26, bold=True, color=WHITE)

    # ── 6 KPI cards ───────────────────────────────────────────────────────────
    total     = metrics.get("total_resources", 0)
    assessed  = metrics.get("assessed_resources", 0)
    comp_pct  = metrics.get("completion_pct", 0)
    compl_pct = metrics.get("compliance_pct", 0)
    no_asmt   = metrics.get("no_assessment", 0)
    below     = metrics.get("below_target", 0)

    comp_delta   = metrics.get("completion_delta", "")
    comp_dpos    = metrics.get("completion_delta_pos", True)
    compl_delta  = metrics.get("compliance_delta", "")
    compl_dpos   = metrics.get("compliance_delta_pos", True)
    noasmt_delta = metrics.get("no_assessment_delta", "")
    noasmt_dpos  = metrics.get("no_assessment_delta_pos", True)
    below_delta  = metrics.get("below_target_delta", "")
    below_dpos   = metrics.get("below_target_delta_pos", True)

    has_delta = any([comp_delta, compl_delta, noasmt_delta, below_delta])
    card_w = Inches(2.02)
    card_h = Inches(1.45) if has_delta else Inches(1.25)
    gap    = Inches(0.115)
    kpi_y  = Inches(1.12)
    x0     = Inches(0.32)

    cards = [
        ("Total Resources",   f"{total:,}",        "in scope",                   NAVY,    "",           True),
        ("Assessed Resources",f"{assessed:,}",      f"{comp_pct:.0f}% completion",_pct_color(comp_pct),"",True),
        ("Completion",        f"{comp_pct:.1f}%",   f"of {total:,} resources",    _pct_color(comp_pct), comp_delta,  comp_dpos),
        ("Target Compliance", f"{compl_pct:.0f}%",  "meeting target",             _pct_color(compl_pct),compl_delta, compl_dpos),
        ("No Assessment",     f"{no_asmt:,}",       "remaining to complete",      ORANGE,  noasmt_delta,noasmt_dpos),
        ("Below Target",      f"{below:,}",         "upskilling pool",            RED,     below_delta, below_dpos),
    ]
    for i, (lbl, val, sub, col, dlt, dpos) in enumerate(cards):
        _kpi_card_6(slide, lbl, val, sub,
                    x0 + i * (card_w + gap), kpi_y,
                    card_w, card_h, col,
                    delta=dlt, delta_positive=dpos)

    # ── Divider ────────────────────────────────────────────────────────────────
    _rect(slide, Inches(6.635), Inches(2.52), Inches(0.03), Inches(4.7), LTGREY)

    # ── LEFT PANEL: Skill Proficiency Mix ──────────────────────────────────────
    lx = Inches(0.32)
    ly = Inches(2.52)
    lw = Inches(6.2)

    _txt(slide, "Skill Proficiency Mix",
         lx, ly, lw, Inches(0.32),
         size=11, bold=True, color=NAVY)
    _rect(slide, lx, ly + Inches(0.33), Inches(0.55), Inches(0.03), NAVY)

    # Table: Proficiency | Count | Share%
    rows = proficiency_rows or []
    if rows:
        th = Inches(0.26)
        ty0 = ly + Inches(0.43)
        col_x = [lx, lx + Inches(1.55), lx + Inches(2.35)]
        col_w = [Inches(1.5), Inches(0.75), Inches(0.85)]
        hdrs  = ["Proficiency", "Count", "Share %"]

        for ci, (hdr, cx, cw) in enumerate(zip(hdrs, col_x, col_w)):
            _rect(slide, cx, ty0, cw - Inches(0.02), th, NAVY)
            _txt(slide, hdr, cx + Inches(0.05), ty0 + Inches(0.04),
                 cw - Inches(0.08), th - Inches(0.06),
                 size=8, bold=True, color=WHITE)

        for ri, row in enumerate(rows):
            bg  = OFFWHT if ri % 2 == 0 else WHITE
            ry  = ty0 + th + ri * th
            for ci, (cx, cw) in enumerate(zip(col_x, col_w)):
                _rect(slide, cx, ry, cw - Inches(0.02), th, bg)
            prof    = str(row.get("Proficiency", ""))
            cnt     = str(row.get("Count", ""))
            pct_val = row.get("Percent", 0)
            pct_str = f"{pct_val:.1f}%"

            prof_hex = PROF_HEX.get(prof, "6B7B8D")
            prof_rgb = _hex_to_rgb(prof_hex)

            _txt(slide, prof, col_x[0] + Inches(0.05),
                 ry + Inches(0.04), col_w[0] - Inches(0.08), th - Inches(0.06),
                 size=9, bold=True, color=prof_rgb)
            _txt(slide, cnt, col_x[1] + Inches(0.05),
                 ry + Inches(0.04), col_w[1] - Inches(0.08), th - Inches(0.06),
                 size=9, color=NAVY)
            _txt(slide, pct_str, col_x[2] + Inches(0.05),
                 ry + Inches(0.04), col_w[2] - Inches(0.08), th - Inches(0.06),
                 size=9, color=NAVY)

        table_bottom = ty0 + th + len(rows) * th

        # Donut chart to the right of the table
        if len(rows) > 0:
            profs  = [r.get("Proficiency", "") for r in rows]
            counts = [max(1, int(r.get("Count", 0) or 0)) for r in rows]
            colors_list = [PROF_HEX.get(p, "9CA3AF") for p in profs]

            cd = ChartData()
            cd.categories = profs
            cd.add_series("", counts)

            chart_x = lx + Inches(3.25)
            chart_y = ly + Inches(0.38)
            chart_w = Inches(2.65)
            chart_h = Inches(2.65)

            chart_frame = slide.shapes.add_chart(
                XL_CHART_TYPE.DOUGHNUT,
                chart_x, chart_y, chart_w, chart_h, cd
            )
            ch = chart_frame.chart
            ch.has_legend = False
            ch.has_title  = False
            _set_donut_point_colors(ch, colors_list)

        # Legend below chart
        legend_y = ly + Inches(3.2)
        for i, row in enumerate(rows[:6]):
            prof     = str(row.get("Proficiency", ""))
            hex_col  = PROF_HEX.get(prof, "9CA3AF")
            rgb_col  = _hex_to_rgb(hex_col)
            ix       = lx + (Inches(3.1) if i >= 3 else Inches(0.0))
            iy_off   = (i % 3) * Inches(0.3)
            _rect(slide, ix, legend_y + iy_off, Inches(0.18), Inches(0.18), rgb_col)
            _txt(slide, prof,
                 ix + Inches(0.22), legend_y + iy_off - Inches(0.02),
                 Inches(0.9), Inches(0.24), size=8, color=NAVY)

    # AI insight for proficiency mix
    summary = insights.get("executive_summary", "")
    if summary:
        nar_y = ly + Inches(4.1)
        _txt(slide, summary,
             lx, nar_y, lw, H - nar_y - Inches(0.15),
             size=8.5, color=GREY, italic=True, wrap=True)

    # ── RIGHT PANEL: Career Level Competency Health ────────────────────────────
    rx  = Inches(6.85)
    ry0 = Inches(2.52)
    rw  = Inches(6.15)

    _txt(slide, "Career Level Competency Health",
         rx, ry0, rw, Inches(0.32),
         size=11, bold=True, color=NAVY)
    _rect(slide, rx, ry0 + Inches(0.33), Inches(0.55), Inches(0.03), TEAL)

    # 4 mini-KPIs
    kpis = career_kpis or {}
    mini_cards = [
        ("Avg Completion",   f"{kpis.get('avg_completion', 0):.1f}%",  TEAL),
        ("Avg Compliance",   f"{kpis.get('avg_compliance', 0):.1f}%",  _pct_color(kpis.get("avg_compliance", 0))),
        ("Widest Gap",       f"{kpis.get('widest_gap', 0)}%",          ORANGE),
        ("Priority Segment", str(kpis.get("priority_segment", "—")),   RED),
    ]
    mcw = Inches(1.42)
    mch = Inches(0.75)
    mgap = Inches(0.085)
    my = ry0 + Inches(0.45)
    for i, (mlbl, mval, mcol) in enumerate(mini_cards):
        mx = rx + i * (mcw + mgap)
        _rect(slide, mx, my, mcw, mch, OFFWHT)
        _rect(slide, mx, my, mcw, Inches(0.05), mcol)
        _txt(slide, mval, mx + Inches(0.06), my + Inches(0.06),
             mcw - Inches(0.12), Inches(0.38),
             size=22, bold=True, color=mcol, align=PP_ALIGN.CENTER)
        _txt(slide, mlbl, mx + Inches(0.04), my + Inches(0.44),
             mcw - Inches(0.08), Inches(0.25),
             size=8, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

    # Career Level grouped bar chart: Completion% vs Target Compliance%
    if career_rows:
        cl_rows    = sorted(career_rows, key=lambda r: int(r.get("Career Level", 0)))
        cl_labels  = [f"CL{r['Career Level']}" for r in cl_rows]
        comp_vals  = [float(r.get("Completion %", 0)) for r in cl_rows]
        compl_vals = [float(r.get("Target Compliance %", 0)) for r in cl_rows]

        cd2 = ChartData()
        cd2.categories = cl_labels
        cd2.add_series("Completion %", comp_vals)
        cd2.add_series("Target Compliance %", compl_vals)

        chart_frame2 = slide.shapes.add_chart(
            XL_CHART_TYPE.BAR_CLUSTERED,
            rx, my + mch + Inches(0.2),
            rw, H - (my + mch + Inches(0.3)),
            cd2
        )
        ch2 = chart_frame2.chart
        ch2.has_legend = True
        ch2.has_title  = False
        _set_series_color(ch2.series[0], "028090")   # Teal for Completion
        _set_series_color(ch2.series[1], "E55C2A")   # Orange for Compliance

        try:
            ch2.value_axis.has_major_gridlines = False
            ch2.category_axis.tick_labels.font.size = Pt(8)
            ch2.value_axis.tick_labels.font.size = Pt(8)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Slide 3 — Career Level Detail
# ---------------------------------------------------------------------------

def _slide_career_detail(prs, career_rows: list):
    slide = _add_slide(prs)
    _rect(slide, 0, 0, W, Inches(1.05), NAVY)
    _txt(slide, "Career Level Competency Health",
         Inches(0.5), Inches(0.12), Inches(12), Inches(0.4),
         size=26, bold=True, color=WHITE)
    _txt(slide, "Completion and compliance breakdown by career level",
         Inches(0.5), Inches(0.65), Inches(12), Inches(0.35),
         size=11, color=ICE)

    if not career_rows:
        _txt(slide, "No career level data available.",
             Inches(1), Inches(2), Inches(10), Inches(1),
             size=16, color=NAVY)
        return

    # ── Table (left side) ─────────────────────────────────────────────────────
    cols_def = [
        ("Career Level",   Inches(1.4)),
        ("Total",          Inches(0.8)),
        ("No Assessment",  Inches(1.3)),
        ("Below Target",   Inches(1.2)),
        ("Completion %",   Inches(1.2)),
        ("Compliance %",   Inches(1.2)),
    ]
    col_labels = [c[0] for c in cols_def]
    col_widths = [c[1] for c in cols_def]
    col_x      = [Inches(0.4)]
    for cw in col_widths[:-1]:
        col_x.append(col_x[-1] + cw)

    row_h = Inches(0.4)
    hdr_y = Inches(1.2)
    data_y = hdr_y + row_h

    for lbl, cx, cw in zip(col_labels, col_x, col_widths):
        _rect(slide, cx, hdr_y, cw - Inches(0.03), row_h, NAVY)
        _txt(slide, lbl, cx + Inches(0.05), hdr_y + Inches(0.08),
             cw - Inches(0.1), row_h - Inches(0.1),
             size=9, bold=True, color=WHITE)

    sorted_rows = sorted(career_rows, key=lambda r: int(r.get("Career Level", 0)), reverse=True)
    for ri, row in enumerate(sorted_rows[:12]):
        bg = OFFWHT if ri % 2 == 0 else WHITE
        ry = data_y + ri * row_h
        _rect(slide, col_x[0], ry, sum(col_widths) - Inches(0.03), row_h, bg)

        vals = [
            f"CL{int(row.get('Career Level', ''))}",
            str(row.get("Total Resources", "")),
            str(row.get("No Assessment", "")),
            str(row.get("Below Target", "")),
            f"{float(row.get('Completion %', 0)):.0f}%",
            f"{float(row.get('Target Compliance %', 0)):.0f}%",
        ]
        for ci, (val, cx, cw) in enumerate(zip(vals, col_x, col_widths)):
            txt_color = NAVY
            if ci == 4:
                txt_color = _pct_color(float(row.get("Completion %", 0)))
            elif ci == 5:
                txt_color = _pct_color(float(row.get("Target Compliance %", 0)))
            _txt(slide, val, cx + Inches(0.06), ry + Inches(0.1),
                 cw - Inches(0.1), row_h - Inches(0.1),
                 size=10, color=txt_color, bold=(ci in [4, 5]))

    # ── Grouped bar chart (right side) ────────────────────────────────────────
    sorted_asc = sorted(career_rows, key=lambda r: int(r.get("Career Level", 0)))
    cl_labels  = [f"CL{r['Career Level']}" for r in sorted_asc]
    comp_vals  = [float(r.get("Completion %", 0)) for r in sorted_asc]
    compl_vals = [float(r.get("Target Compliance %", 0)) for r in sorted_asc]

    cd = ChartData()
    cd.categories = cl_labels
    cd.add_series("Completion %", comp_vals)
    cd.add_series("Target Compliance %", compl_vals)

    chart_x = col_x[0] + sum(col_widths) + Inches(0.3)
    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        chart_x, Inches(1.15),
        W - chart_x - Inches(0.3), H - Inches(1.4),
        cd
    )
    ch = chart_frame.chart
    ch.has_legend = True
    ch.has_title  = False
    _set_series_color(ch.series[0], "028090")
    _set_series_color(ch.series[1], "E55C2A")

    try:
        ch.value_axis.has_major_gridlines = False
        ch.category_axis.tick_labels.font.size = Pt(9)
        ch.value_axis.tick_labels.font.size = Pt(9)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Slide 4 — Skill Gap Analysis
# ---------------------------------------------------------------------------

def _slide_skill_gaps(prs, skill_rows: list, insights: dict):
    slide = _add_slide(prs)
    _rect(slide, 0, 0, W, Inches(1.05), NAVY)
    _txt(slide, "Primary Skill Gap Analysis",
         Inches(0.5), Inches(0.12), Inches(12), Inches(0.4),
         size=26, bold=True, color=WHITE)
    _txt(slide, "Top skills ranked by resources needing action (Below Target + No Assessment)",
         Inches(0.5), Inches(0.65), Inches(12), Inches(0.35),
         size=11, color=ICE)

    # Sort by Resources To Chase descending (skill_export may come in alphabetical order)
    top = sorted(
        skill_rows or [],
        key=lambda r: int(r.get("Resources To Chase", 0) or 0),
        reverse=True,
    )[:10]
    if not top:
        _txt(slide, "No skill gap data available.",
             Inches(1), Inches(2), Inches(10), Inches(1),
             size=16, color=NAVY)
        return

    # ── Table (left half) ──────────────────────────────────────────────────────
    tx   = Inches(0.35)
    ty0  = Inches(1.2)
    tcols = [
        ("Primary Skill",  Inches(2.4)),
        ("Total",          Inches(0.65)),
        ("Below Tgt",      Inches(0.8)),
        ("No Asmt",        Inches(0.75)),
        ("To Chase",       Inches(0.75)),
    ]
    t_labels = [c[0] for c in tcols]
    t_widths = [c[1] for c in tcols]
    t_x      = [tx]
    for cw in t_widths[:-1]:
        t_x.append(t_x[-1] + cw)

    th = Inches(0.36)
    for lbl, cx, cw in zip(t_labels, t_x, t_widths):
        _rect(slide, cx, ty0, cw - Inches(0.02), th, NAVY)
        _txt(slide, lbl, cx + Inches(0.04), ty0 + Inches(0.07),
             cw - Inches(0.08), th - Inches(0.08),
             size=8, bold=True, color=WHITE)

    for ri, row in enumerate(top):
        bg = OFFWHT if ri % 2 == 0 else WHITE
        ry = ty0 + th + ri * th
        _rect(slide, t_x[0], ry, sum(t_widths) - Inches(0.02), th, bg)

        skill   = str(row.get("Primary Skill", row.get("SkillName", "")))[:28]
        total   = int(row.get("TotalResources", row.get("Resources To Chase", 0)) or 0)
        below   = int(row.get("Below Target Only", row.get("BelowTargetOnly", 0)) or 0)
        no_asmt = int(row.get("No Assessment", row.get("NoAssessment", 0)) or 0)
        chase   = int(row.get("Resources To Chase", 0) or 0)

        vals = [skill, str(total), str(below), str(no_asmt), str(chase)]
        for ci, (val, cx, cw) in enumerate(zip(vals, t_x, t_widths)):
            tc = ORANGE if ci == 4 else NAVY
            _txt(slide, val, cx + Inches(0.04), ry + Inches(0.07),
                 cw - Inches(0.08), th - Inches(0.08),
                 size=9, color=tc, bold=(ci == 4))

    # ── Stacked bar chart (right half) ────────────────────────────────────────
    skill_names = [
        str(r.get("Primary Skill", r.get("SkillName", "")))[:22]
        for r in top
    ]
    below_vals   = [int(r.get("Below Target Only", r.get("BelowTargetOnly", 0)) or 0) for r in top]
    no_asmt_vals = [int(r.get("No Assessment", r.get("NoAssessment", 0)) or 0) for r in top]

    cd = ChartData()
    cd.categories = skill_names
    cd.add_series("Below Target", below_vals)
    cd.add_series("No Assessment", no_asmt_vals)

    chart_x = t_x[-1] + t_widths[-1] + Inches(0.25)
    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_STACKED,
        chart_x, Inches(1.15),
        W - chart_x - Inches(0.35), H - Inches(1.4),
        cd
    )
    ch = chart_frame.chart
    ch.has_legend = True
    ch.has_title  = False
    _set_series_color(ch.series[0], "028090")   # Teal for Below Target
    _set_series_color(ch.series[1], "E55C2A")   # Orange for No Assessment

    try:
        ch.plots[0].gap_width = 80
        ch.value_axis.has_major_gridlines = False
        ch.category_axis.tick_labels.font.size = Pt(8)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Slide 5 — AI Insights + Recommendations
# ---------------------------------------------------------------------------

def _slide_insights_recs(prs, insights: dict):
    slide = _add_slide(prs)
    _rect(slide, 0, 0, W, Inches(1.05), TEAL)
    _txt(slide, "AI Executive Insights & Recommendations",
         Inches(0.5), Inches(0.12), Inches(12), Inches(0.4),
         size=26, bold=True, color=WHITE)
    _txt(slide, "Generated by Azure OpenAI based on current scorecard data",
         Inches(0.5), Inches(0.65), Inches(12), Inches(0.35),
         size=11, color=WHITE)

    summary  = insights.get("executive_summary", "")
    findings = insights.get("key_findings", [])
    recs     = insights.get("recommendations", [])

    # Left column: Summary + Key Findings
    _rect(slide, Inches(0.35), Inches(1.12), Inches(0.04), Inches(5.9), TEAL)
    _txt(slide, "Executive Summary",
         Inches(0.5), Inches(1.15), Inches(5.8), Inches(0.3),
         size=11, bold=True, color=NAVY)
    if summary:
        _txt(slide, summary,
             Inches(0.5), Inches(1.5), Inches(5.8), Inches(1.3),
             size=10, color=NAVY, italic=True, wrap=True)

    _txt(slide, "Key Findings",
         Inches(0.5), Inches(2.9), Inches(5.8), Inches(0.3),
         size=11, bold=True, color=NAVY)

    fy = Inches(3.25)
    for bullet in findings[:5]:
        _rect(slide, Inches(0.5), fy + Inches(0.07),
              Inches(0.1), Inches(0.1), TEAL)
        _txt(slide, bullet,
             Inches(0.7), fy, Inches(5.5), Inches(0.6),
             size=10, color=NAVY, wrap=True)
        fy += Inches(0.7)

    # Right column: Recommendations
    _rect(slide, Inches(6.75), Inches(1.12), Inches(0.04), Inches(5.9), ORANGE)
    _txt(slide, "Recommended Next Steps",
         Inches(6.9), Inches(1.15), Inches(6.0), Inches(0.3),
         size=11, bold=True, color=NAVY)

    ry_start = Inches(1.5)
    for j, rec in enumerate(recs[:6]):
        badge_col = TEAL if j % 2 == 0 else ORANGE
        _rect(slide, Inches(6.9), ry_start + Inches(0.05),
              Inches(0.35), Inches(0.35), badge_col)
        _txt(slide, str(j + 1),
             Inches(6.9), ry_start + Inches(0.04),
             Inches(0.35), Inches(0.35),
             size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _txt(slide, rec,
             Inches(7.35), ry_start, Inches(5.6), Inches(0.6),
             size=10, color=NAVY, wrap=True)
        ry_start += Inches(0.75)

    # Footer
    _rect(slide, 0, H - Inches(0.4), W, Inches(0.4), NAVY)
    _txt(slide, "myCompetency Scorecard  ·  Confidential  ·  ✦ Generated by Azure OpenAI",
         Inches(0.4), H - Inches(0.38), W - Inches(0.8), Inches(0.36),
         size=9, color=WHITE, align=PP_ALIGN.CENTER)


# ---------------------------------------------------------------------------
# Slide 6 — Executive One-Pager (summary of everything)
# ---------------------------------------------------------------------------

def _slide_one_pager(prs, metrics: dict, proficiency_rows: list, skill_rows: list,
                     career_kpis: dict, insights: dict):
    """Single-slide summary: KPIs + Skill Gap chart + AI insights."""
    slide = _add_slide(prs)

    # Full background — light off-white
    _rect(slide, 0, 0, W, H, OFFWHT)
    # Top header stripe
    _rect(slide, 0, 0, W, Inches(0.72), NAVY)
    _txt(slide, "Executive One-Pager",
         Inches(0.4), Inches(0.06), Inches(7), Inches(0.36),
         size=18, bold=True, color=WHITE)
    _txt(slide, metrics.get("business_group", ""),
         Inches(7.5), Inches(0.06), Inches(5.4), Inches(0.36),
         size=12, color=ICE, align=PP_ALIGN.RIGHT)
    _txt(slide, "Primary Skill Scorecard  ·  " + datetime.now().strftime("%B %Y"),
         Inches(0.4), Inches(0.42), Inches(12), Inches(0.26),
         size=9, color=ICE)

    # ── LEFT COLUMN (KPIs + proficiency table) x: 0.3" – 4.35" ──────────────
    lx  = Inches(0.28)
    lw  = Inches(4.05)
    ly0 = Inches(0.85)

    # Section label
    _txt(slide, "KEY METRICS",
         lx, ly0, lw, Inches(0.22),
         size=8, bold=True, color=NAVY)
    _rect(slide, lx, ly0 + Inches(0.22), lw, Inches(0.025), NAVY)

    # 6 compact KPI cards — 2 rows × 3
    total     = metrics.get("total_resources", 0)
    assessed  = metrics.get("assessed_resources", 0)
    comp_pct  = metrics.get("completion_pct", 0)
    compl_pct = metrics.get("compliance_pct", 0)
    no_asmt   = metrics.get("no_assessment", 0)
    below     = metrics.get("below_target", 0)

    _comp_delta   = metrics.get("completion_delta", "")
    _comp_dpos    = metrics.get("completion_delta_pos", True)
    _compl_delta  = metrics.get("compliance_delta", "")
    _compl_dpos   = metrics.get("compliance_delta_pos", True)
    _noasmt_delta = metrics.get("no_assessment_delta", "")
    _noasmt_dpos  = metrics.get("no_assessment_delta_pos", True)
    _below_delta  = metrics.get("below_target_delta", "")
    _below_dpos   = metrics.get("below_target_delta_pos", True)

    _has_delta = any([_comp_delta, _compl_delta, _noasmt_delta, _below_delta])

    kpi_items = [
        ("Total",         f"{total:,}",         NAVY,                  "",            True),
        ("Assessed",      f"{assessed:,}",       TEAL,                  "",            True),
        ("Completion",    f"{comp_pct:.0f}%",    _pct_color(comp_pct),  _comp_delta,  _comp_dpos),
        ("Compliance",    f"{compl_pct:.0f}%",   _pct_color(compl_pct), _compl_delta, _compl_dpos),
        ("No Assessment", f"{no_asmt:,}",        ORANGE,                _noasmt_delta,_noasmt_dpos),
        ("Below Target",  f"{below:,}",          RED,                   _below_delta, _below_dpos),
    ]
    mcw  = Inches(1.26)
    mch  = Inches(1.0) if _has_delta else Inches(0.82)
    mgap = Inches(0.065)
    krow = ly0 + Inches(0.28)

    for i, (lbl, val, col, dlt, dpos) in enumerate(kpi_items):
        col_i = i % 3
        row_i = i // 3
        mx = lx + col_i * (mcw + mgap)
        my = krow + row_i * (mch + mgap)
        _rect(slide, mx, my, mcw, mch, WHITE)
        _rect(slide, mx, my, mcw, Inches(0.055), col)
        _txt(slide, val,
             mx + Inches(0.07), my + Inches(0.07),
             mcw - Inches(0.14), Inches(0.44),
             size=24, bold=True, color=col, align=PP_ALIGN.CENTER)
        _txt(slide, lbl,
             mx + Inches(0.05), my + Inches(0.52),
             mcw - Inches(0.1), Inches(0.20),
             size=8, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        if dlt:
            dlt_color = GREEN if dpos else RED
            _txt(slide, dlt,
                 mx + Inches(0.05), my + Inches(0.73),
                 mcw - Inches(0.1), Inches(0.20),
                 size=6.5, bold=True, color=dlt_color, align=PP_ALIGN.CENTER)

    # Proficiency mix table (compact)
    prof_y = krow + 2 * (mch + mgap) + Inches(0.22)
    _txt(slide, "PROFICIENCY MIX",
         lx, prof_y, lw, Inches(0.2),
         size=8, bold=True, color=NAVY)
    _rect(slide, lx, prof_y + Inches(0.2), lw, Inches(0.025), TEAL)

    ph = Inches(0.29)
    hdr_y2 = prof_y + Inches(0.26)
    pcol_x  = [lx,              lx + Inches(2.15), lx + Inches(3.0)]
    pcol_w  = [Inches(2.1),     Inches(0.8),       Inches(1.0)]
    for ci, (hdr, cx, cw) in enumerate(zip(["Proficiency","Count","Share %"], pcol_x, pcol_w)):
        _rect(slide, cx, hdr_y2, cw - Inches(0.02), ph, NAVY)
        _txt(slide, hdr, cx + Inches(0.04), hdr_y2 + Inches(0.06),
             cw - Inches(0.06), ph - Inches(0.06),
             size=7, bold=True, color=WHITE)

    for ri, row in enumerate((proficiency_rows or [])[:6]):
        bg  = OFFWHT if ri % 2 == 0 else WHITE
        ry  = hdr_y2 + ph + ri * ph
        for cx, cw in zip(pcol_x, pcol_w):
            _rect(slide, cx, ry, cw - Inches(0.02), ph, bg)
        prof    = str(row.get("Proficiency", ""))
        rgb_col = _hex_to_rgb(PROF_HEX.get(prof, "6B7B8D"))
        _txt(slide, prof,
             pcol_x[0] + Inches(0.04), ry + Inches(0.06),
             pcol_w[0] - Inches(0.06), ph - Inches(0.06),
             size=8, bold=True, color=rgb_col)
        _txt(slide, str(row.get("Count", "")),
             pcol_x[1] + Inches(0.04), ry + Inches(0.06),
             pcol_w[1] - Inches(0.06), ph - Inches(0.06),
             size=8, color=NAVY)
        _txt(slide, f"{row.get('Percent', 0):.1f}%",
             pcol_x[2] + Inches(0.04), ry + Inches(0.06),
             pcol_w[2] - Inches(0.06), ph - Inches(0.06),
             size=8, color=NAVY)

    # Career Level KPIs below proficiency table
    kpis = career_kpis or {}
    cl_y = hdr_y2 + ph * (len(proficiency_rows or []) + 1) + Inches(0.2)
    _txt(slide, "CAREER LEVEL",
         lx, cl_y, lw, Inches(0.2),
         size=8, bold=True, color=NAVY)
    _rect(slide, lx, cl_y + Inches(0.2), lw, Inches(0.025), ORANGE)

    cl_items = [
        ("Avg Completion",   f"{kpis.get('avg_completion', 0):.0f}%",    TEAL),
        ("Avg Compliance",   f"{kpis.get('avg_compliance', 0):.0f}%",    _pct_color(kpis.get("avg_compliance", 0))),
        ("Widest Gap",       f"{kpis.get('widest_gap', 0)}%",            ORANGE),
        ("Priority Segment", str(kpis.get("priority_segment", "—")),     RED),
    ]
    clw = Inches(0.94)
    cl_card_y = cl_y + Inches(0.28)
    for ci, (clbl, cval, ccol) in enumerate(cl_items):
        cx2 = lx + ci * (clw + Inches(0.055))
        _rect(slide, cx2, cl_card_y, clw, Inches(0.72), WHITE)
        _rect(slide, cx2, cl_card_y, clw, Inches(0.045), ccol)
        _txt(slide, cval,
             cx2 + Inches(0.04), cl_card_y + Inches(0.05),
             clw - Inches(0.08), Inches(0.4),
             size=18, bold=True, color=ccol, align=PP_ALIGN.CENTER)
        _txt(slide, clbl,
             cx2 + Inches(0.03), cl_card_y + Inches(0.45),
             clw - Inches(0.06), Inches(0.22),
             size=6.5, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

    # ── MIDDLE COLUMN (skill gap chart) x: 4.5" – 8.8" ──────────────────────
    mx0  = Inches(4.5)
    mw   = Inches(4.3)
    my0  = Inches(0.85)

    _txt(slide, "TOP SKILL GAPS",
         mx0, my0, mw, Inches(0.22),
         size=8, bold=True, color=NAVY)
    _rect(slide, mx0, my0 + Inches(0.22), mw, Inches(0.025), TEAL)

    top5 = sorted(
        skill_rows or [],
        key=lambda r: int(r.get("Resources To Chase", 0) or 0),
        reverse=True,
    )[:8]

    if top5:
        s_names   = [str(r.get("Primary Skill", r.get("SkillName", "")))[:30] for r in top5]
        below_v   = [int(r.get("Below Target Only", r.get("BelowTargetOnly", 0)) or 0) for r in top5]
        no_asmt_v = [int(r.get("No Assessment", r.get("NoAssessment", 0)) or 0) for r in top5]

        cd = ChartData()
        cd.categories = s_names
        cd.add_series("No Assessment", no_asmt_v)
        cd.add_series("Below Target",  below_v)

        cf = slide.shapes.add_chart(
            XL_CHART_TYPE.BAR_STACKED,
            mx0, my0 + Inches(0.3),
            mw, H - my0 - Inches(0.4),
            cd
        )
        ch = cf.chart
        ch.has_legend = True
        ch.has_title  = False
        _set_series_color(ch.series[0], "028090")   # Teal = No Assessment
        _set_series_color(ch.series[1], "4B2D7F")   # Purple = Below Target

        try:
            ch.plots[0].gap_width = 70
            ch.value_axis.has_major_gridlines = False
            ch.category_axis.tick_labels.font.size = Pt(7)
            ch.value_axis.tick_labels.font.size = Pt(7)
            ch.legend.position = 4   # bottom
            ch.legend.include_in_layout = False
        except Exception:
            pass

    # ── RIGHT COLUMN (AI insights) x: 8.95" – 12.95" ────────────────────────
    rx0  = Inches(8.95)
    rw   = Inches(4.0)
    ry0  = Inches(0.85)

    _txt(slide, "AI EXECUTIVE INSIGHTS",
         rx0, ry0, rw, Inches(0.22),
         size=8, bold=True, color=NAVY)
    _rect(slide, rx0, ry0 + Inches(0.22), rw, Inches(0.025), ORANGE)

    summary = insights.get("executive_summary", "")
    _rect(slide, rx0, ry0 + Inches(0.3), rw, Inches(1.5), WHITE)
    if summary:
        _txt(slide, summary,
             rx0 + Inches(0.1), ry0 + Inches(0.35),
             rw - Inches(0.2), Inches(1.4),
             size=8, color=NAVY, italic=True, wrap=True)

    findings = insights.get("key_findings", [])
    fy2 = ry0 + Inches(1.95)
    _txt(slide, "Key Findings",
         rx0, fy2, rw, Inches(0.2),
         size=8, bold=True, color=NAVY)
    fy2 += Inches(0.22)
    for bullet in findings[:4]:
        _rect(slide, rx0, fy2 + Inches(0.05), Inches(0.08), Inches(0.08), TEAL)
        _txt(slide, bullet,
             rx0 + Inches(0.13), fy2, rw - Inches(0.14), Inches(0.45),
             size=7.5, color=NAVY, wrap=True)
        fy2 += Inches(0.48)

    recs = insights.get("recommendations", [])
    ry2 = fy2 + Inches(0.1)
    _txt(slide, "Recommended Actions",
         rx0, ry2, rw, Inches(0.2),
         size=8, bold=True, color=NAVY)
    ry2 += Inches(0.22)
    for j, rec in enumerate(recs[:4]):
        badge_col = TEAL if j % 2 == 0 else ORANGE
        _rect(slide, rx0, ry2 + Inches(0.03), Inches(0.24), Inches(0.24), badge_col)
        _txt(slide, str(j + 1),
             rx0, ry2 + Inches(0.02), Inches(0.24), Inches(0.26),
             size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _txt(slide, rec,
             rx0 + Inches(0.28), ry2, rw - Inches(0.3), Inches(0.45),
             size=7.5, color=NAVY, wrap=True)
        ry2 += Inches(0.5)

    # Footer
    _rect(slide, 0, H - Inches(0.3), W, Inches(0.3), NAVY)
    _txt(slide, "myCompetency Scorecard  ·  Confidential  ·  ✦ Generated by Azure OpenAI",
         Inches(0.4), H - Inches(0.28), W - Inches(0.8), Inches(0.25),
         size=8, color=WHITE, align=PP_ALIGN.CENTER)


# ---------------------------------------------------------------------------
# Azure OpenAI call
# ---------------------------------------------------------------------------

def generate_insights(metrics: dict, skill_rows: list, project_rows: list) -> dict:
    api_key    = os.environ.get("LLM_API_KEY") or os.environ.get("AZURE_OPENAI_API_KEY", "")
    base_url   = os.environ.get("AZURE_OPENAI_ENDPOINT", "").rstrip("/")
    deployment = os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-4o")

    from openai import OpenAI
    client = OpenAI(api_key=api_key, base_url=base_url)

    top_skills = [
        f"  - {r.get('Primary Skill', r.get('SkillName', ''))}: "
        f"{r.get('Resources To Chase', 0)} needing action "
        f"({r.get('No Assessment', r.get('NoAssessment', 0))} no asmt, "
        f"{r.get('Below Target Only', r.get('BelowTargetOnly', 0))} below target)"
        for r in skill_rows[:6]
    ]
    top_projects = [
        f"  - {r.get('Project', '')}: {r.get('Resources To Chase', 0)} to chase "
        f"({r.get('Chase %', 0):.0f}%)"
        for r in (project_rows or [])[:5]
        if r.get("Resources To Chase", 0) > 0
    ]

    prompt = f"""You are a Talent & Capability analytics assistant preparing an executive report.

Current myCompetency Scorecard for {metrics.get('business_group', 'the team')}:

METRICS
- Total: {metrics.get('total_resources', 0)}
- Assessed: {metrics.get('assessed_resources', 0)} ({metrics.get('completion_pct', 0):.1f}%)
- Target Compliance: {metrics.get('compliance_pct', 0):.1f}%
- No Assessment: {metrics.get('no_assessment', 0)}
- Below Target: {metrics.get('below_target', 0)}

TOP SKILL GAPS:
{chr(10).join(top_skills) or "  - None identified"}

TOP PROJECTS REQUIRING FOLLOW-UP:
{chr(10).join(top_projects) or "  - All on track"}

Respond with ONLY this JSON:
{{
  "executive_summary": "<2-3 sentence summary>",
  "key_findings": ["<finding 1>", "<finding 2>", "<finding 3>", "<finding 4>", "<finding 5>"],
  "recommendations": ["<action 1>", "<action 2>", "<action 3>", "<action 4>", "<action 5>", "<action 6>"]
}}

Be factual, data-driven, and specific. No fluff."""

    response = client.chat.completions.create(
        model=deployment,
        messages=[{"role": "user", "content": prompt}],
        max_tokens=1024,
    )

    raw = response.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def build_executive_pptx(
    business_group: str,
    metrics: dict,
    career_rows: list,
    skill_rows: list,
    project_rows: list,
    proficiency_rows: list = None,
    career_kpis: dict = None,
) -> io.BytesIO:
    """
    Generate AI Executive Report PPTX. Returns BytesIO for st.download_button.

    Parameters
    ----------
    business_group   : Display name shown in header/title
    metrics          : total_resources, assessed_resources, completion_pct,
                       compliance_pct, no_assessment, below_target
    career_rows      : [{Career Level, Total Resources, No Assessment,
                        Completion %, Target Compliance %, Below Target}]
    skill_rows       : [{Primary Skill/SkillName, No Assessment/NoAssessment,
                        Below Target Only/BelowTargetOnly, Resources To Chase,
                        TotalResources}]
    project_rows     : [{Project, Resources To Chase, Chase %}]
    proficiency_rows : [{Proficiency, Count, Percent}]  (optional)
    career_kpis      : {avg_completion, avg_compliance, widest_gap,
                        priority_segment, priority_sub}  (optional)
    """
    metrics["business_group"] = business_group
    report_date = datetime.now().strftime("%B %d, %Y")

    insights = generate_insights(metrics, skill_rows or [], project_rows or [])

    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    _slide_title(prs, business_group, report_date)
    _slide_dashboard(prs, metrics, proficiency_rows or [], career_rows or [],
                     career_kpis or {}, insights)
    _slide_career_detail(prs, career_rows or [])
    _slide_skill_gaps(prs, skill_rows or [], insights)
    _slide_insights_recs(prs, insights)
    _slide_one_pager(prs, metrics, proficiency_rows or [], skill_rows or [],
                     career_kpis or {}, insights)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
