# -*- coding: utf-8 -*-
"""
Isocal processing module — web-callable version.
All CLI interaction removed; use run_isocal() as the single entry point.
"""

import os
import math
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import ScatterChart, Reference, Series
from lxml import etree


# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────

AUTO_CUTOFF_MINUTES = 120.0

YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")
ORANGE_FILL = PatternFill("solid", fgColor="FFC000")
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
BOLD        = Font(bold=True)
CENTER      = Alignment(horizontal="center", vertical="center")
THIN        = Side(style="thin")
BOX         = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ─────────────────────────────────────────────
# RAW DATA READING
# ─────────────────────────────────────────────

def read_raw_data(raw_file, original_filename=None):
    """
    Read raw instrument xlsx.
    Sheet 'Raw data': Col A=time(s), Col D=Heat Flow(W), Col E=Heat(J).
    Rows 1-2 are headers/units. Data starts row 3.
    Valid data: t >= 60s.

    original_filename: the user's original upload filename, used for sample_name
                       derivation so that temp file paths don't pollute the name.

    Returns: sample_name (str), data (list of dicts)
    """
    wb = openpyxl.load_workbook(raw_file, data_only=True)

    # Derive sample name from original filename (not temp path)
    name_source = original_filename if original_filename else raw_file
    sample_name = os.path.splitext(os.path.basename(name_source))[0]

    if "Experiment info" in wb.sheetnames:
        ws_info = wb["Experiment info"]
        for row in ws_info.iter_rows(values_only=True):
            for i, val in enumerate(row):
                if isinstance(val, str) and "name" in val.lower():
                    try:
                        nv = row[i + 1]
                        if nv:
                            sample_name = str(nv).replace(".rslt", "").strip()
                    except IndexError:
                        pass

    ws_raw = wb["Raw data"]
    data = []
    for row in ws_raw.iter_rows(min_row=3, values_only=True):
        t = row[0]
        if t is None or not isinstance(t, (int, float)):
            continue
        t = float(t)
        if t < 60:
            continue
        hf = float(row[3]) if row[3] is not None else 0.0
        h  = float(row[4]) if row[4] is not None else 0.0
        data.append({"time_s": t, "heat_flow_W": hf, "heat_J": h})

    # Trim trailing flatlined rows
    while data and abs(data[-1]["heat_flow_W"]) < 1e-5 and abs(data[-1]["heat_J"]) < 1e-3:
        data.pop()

    return sample_name, data


# ─────────────────────────────────────────────
# MIX DESIGN
# ─────────────────────────────────────────────

def compute_mix_design(paste, clinker, gypsum, ns, noh, limestone, cc, water):
    solids           = clinker + gypsum + ns + noh + limestone + cc
    total            = solids + water
    solids_in_paste  = solids  * paste / total
    clinker_in_paste = clinker * paste / total
    return dict(
        clinker=clinker, gypsum=gypsum, ns=ns, noh=noh,
        limestone=limestone, cc=cc, solids=solids,
        water=water, total=total, paste=paste,
        solids_in_paste=solids_in_paste,
        clinker_in_paste=clinker_in_paste,
    )


# ─────────────────────────────────────────────
# CUTOFF
# ─────────────────────────────────────────────

def resolve_cutoff(cutoff_arg, data):
    """Returns (cutoff_idx, cutoff_min, auto_flag)."""
    cutoff_min = cutoff_arg if cutoff_arg is not None else AUTO_CUTOFF_MINUTES
    auto       = (cutoff_arg is None)
    cutoff_idx = 0
    for i, row in enumerate(data):
        if row["time_s"] / 60.0 <= cutoff_min:
            cutoff_idx = i
        else:
            break
    return cutoff_idx, cutoff_min, auto


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def cl(n):
    return get_column_letter(n)

def header_cell(ws, row, col, val):
    c = ws.cell(row=row, column=col, value=val)
    c.font = BOLD
    c.alignment = CENTER
    c.fill = HEADER_FILL
    c.border = BOX
    return c

def data_cell(ws, row, col, val):
    c = ws.cell(row=row, column=col, value=val)
    return c


# ─────────────────────────────────────────────
# CHART BUILDER
# ─────────────────────────────────────────────

def _make_rich_text(text, sz_hundredths):
    from openpyxl.chart.text import RichText, Text
    from openpyxl.chart.title import Title
    from openpyxl.drawing.text import (RichTextProperties, ListStyle,
                                        Paragraph, ParagraphProperties,
                                        RegularTextRun, CharacterProperties)
    rpr  = CharacterProperties(sz=sz_hundredths)
    run  = RegularTextRun(t=text, rPr=rpr)
    para = Paragraph(r=[run], pPr=ParagraphProperties())
    rt   = RichText(bodyPr=RichTextProperties(), lstStyle=ListStyle(), p=[para])
    return Text(rich=rt)

def nearest_half_above(value):
    return math.ceil(value * 2) / 2


def add_chart(ws, title, x_col, y_col, data_start, last_row,
              x_max, y_min, y_max, color, x_title, y_title, anchor):
    from openpyxl.drawing.colors import ColorChoice, SchemeColor
    from openpyxl.chart.title import Title

    chart = ScatterChart()
    chart.style = None
    chart.title = Title(tx=_make_rich_text(title, 1440))

    x_ref  = Reference(ws, min_col=x_col, min_row=data_start, max_row=last_row)
    y_ref  = Reference(ws, min_col=y_col, min_row=data_start, max_row=last_row)
    series = Series(y_ref, x_ref, title=title)
    series.marker.symbol = "none"
    series.graphicalProperties.line.width = 28575
    series.graphicalProperties.line.solidFill = ColorChoice(
        schemeClr=SchemeColor(val=color)
    )
    chart.series.append(series)

    chart.x_axis.title = Title(tx=_make_rich_text(x_title, 1200))
    chart.y_axis.title = Title(tx=_make_rich_text(y_title, 1200))

    if x_max is not None:
        chart.x_axis.scaling.max = x_max
    if y_min is not None:
        chart.y_axis.scaling.min = y_min
    if y_max is not None:
        chart.y_axis.scaling.max = nearest_half_above(y_max)
        chart.y_axis.majorUnit   = 0.5

    chart.x_axis.numFmt = "0.0"
    chart.y_axis.numFmt = "0.0"

    chart.x_axis.majorTickMark = "out"
    chart.y_axis.majorTickMark = "out"

    ns_uri = "http://schemas.openxmlformats.org/drawingml/2006/main"
    def _black_line_sppr():
        spPr_el = etree.Element(f"{{{ns_uri}}}spPr")
        ln_el   = etree.SubElement(spPr_el, f"{{{ns_uri}}}ln", w="12700")
        sf_el   = etree.SubElement(ln_el,   f"{{{ns_uri}}}solidFill")
        etree.SubElement(sf_el, f"{{{ns_uri}}}srgbClr", val="000000")
        from openpyxl.chart.shapes import GraphicalProperties
        return GraphicalProperties.from_tree(spPr_el)
    chart.x_axis.spPr = _black_line_sppr()
    chart.y_axis.spPr = _black_line_sppr()

    def _white_border_sppr():
        spPr_el = etree.Element(f"{{{ns_uri}}}spPr")
        ln_el   = etree.SubElement(spPr_el, f"{{{ns_uri}}}ln", w="12700")
        sf_el   = etree.SubElement(ln_el,   f"{{{ns_uri}}}solidFill")
        etree.SubElement(sf_el, f"{{{ns_uri}}}srgbClr", val="FFFFFF")
        from openpyxl.chart.shapes import GraphicalProperties
        return GraphicalProperties.from_tree(spPr_el)
    chart.plot_area.spPr = _white_border_sppr()

    from openpyxl.chart.legend import Legend
    leg = Legend()
    leg.position = "b"
    chart.legend = leg

    chart.x_axis.majorGridlines = None
    chart.y_axis.majorGridlines = None

    chart.width  = 15
    chart.height = 7.5

    ws.add_chart(chart, anchor)


# ─────────────────────────────────────────────
# MAIN OUTPUT BUILDER
# ─────────────────────────────────────────────

DATA_START = 4

def build_output(sample_name, data, mix, cutoff_idx, cutoff_min, auto_cutoff, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = sample_name

    n          = len(data)
    last       = DATA_START + n - 1
    cutoff_row = DATA_START + cutoff_idx

    # Row 1: sample name
    ws["A1"] = sample_name
    ws["A1"].font = Font(bold=True, size=12)

    # Row 2: group headers
    for col, val in [(5,"Paste"), (8,"Solids"), (11,"Clinker")]:
        header_cell(ws, 2, col, val)

    # Row 3: column headers
    col_headers = [
        "Time (min)", "Time (Days)", "Heat Flow (W)", "Heat (J)",
        "Heat Flow (mW/g)", "Heat (J/g)", "Adj Heat (J/g)",
        "Heat Flow (mW/g)", "Heat (J/g)", "Adj Heat (J/g)",
        "Heat Flow (mW/g)", "Heat (J/g)", "Adj Heat (J/g)",
    ]
    for j, h in enumerate(col_headers, 1):
        header_cell(ws, 3, j, h)

    # Mix design table — row 2 headers
    for j, h in enumerate(["Clinker","Gypsum","NS","NOH","Limestone","CC","Solids","Water","Total"], 15):
        header_cell(ws, 2, j, h)

    # Row 5: component masses
    for j, v in enumerate([mix["clinker"], mix["gypsum"], mix["ns"], mix["noh"],
                            mix["limestone"], mix["cc"], None, mix["water"], None], 15):
        c = ws.cell(row=5, column=j)
        if v is not None:
            c.value = v
        c.border = BOX
        c.number_format = "0.000"

    ws["U5"] = "=SUM(O5:T5)"
    ws["W5"] = "=U5+V5"
    ws["U5"].border = BOX
    ws["W5"].border = BOX

    # Row 7: derived headers
    for j, h in enumerate(["Paste", "Solids (Paste)", "Clinker (Paste)"], 15):
        header_cell(ws, 7, j, h)

    # Row 8: derived values
    ws["O8"] = mix["paste"]
    ws["P8"] = "=U5*O8/W5"
    ws["Q8"] = "=O5*O8/W5"
    for j in range(15, 18):
        ws.cell(row=8, column=j).border = BOX
        ws.cell(row=8, column=j).number_format = "0.000"

    # Cutoff cell P10
    ws["O10"] = "Cutoff Time (min):"
    ws["O10"].font = BOLD
    ws["O10"].alignment = Alignment(horizontal="right")

    ws["P10"].value         = cutoff_min
    ws["P10"].fill          = ORANGE_FILL if auto_cutoff else YELLOW_FILL
    ws["P10"].font          = Font(bold=True)
    ws["P10"].border        = BOX
    ws["P10"].number_format = "0.0"

    ws["Q10"] = "← AUTO (120 min) — edit to override" if auto_cutoff else "← User-specified"
    ws["Q10"].font = Font(italic=True, color="888888")

    # Data rows
    for i, rd in enumerate(data):
        r    = DATA_START + i
        t_s  = rd["time_s"]
        hfW  = rd["heat_flow_W"]
        hJ   = rd["heat_J"]
        tmin = t_s / 60.0
        tday = tmin / (24 * 60)

        ws.cell(r, 1, tmin)
        ws.cell(r, 2, tday)
        ws.cell(r, 3, hfW)
        ws.cell(r, 4, hJ)

        ws.cell(r, 5,  f"=C{r}*1000/$O$8")
        ws.cell(r, 6,  f"=D{r}/$O$8")
        ws.cell(r, 7,  f"=IF($A{r}<=$P$10,0,F{r}-F{cutoff_row})")

        ws.cell(r, 8,  f"=C{r}*1000/$P$8")
        ws.cell(r, 9,  f"=D{r}/$P$8")
        ws.cell(r, 10, f"=IF($A{r}<=$P$10,0,I{r}-I{cutoff_row})")

        ws.cell(r, 11, f"=C{r}*1000/$Q$8")
        ws.cell(r, 12, f"=D{r}/$Q$8")
        ws.cell(r, 13, f"=IF($A{r}<=$P$10,0,L{r}-L{cutoff_row})")

    # Number formats
    for r in range(DATA_START, DATA_START + n):
        ws.cell(r, 1).number_format = "0.00"
        ws.cell(r, 2).number_format = "0.00000"
        for c in range(3, 14):
            ws.cell(r, c).number_format = "0.000000"

    # Column widths
    widths = {1:12, 2:12, 3:14, 4:12,
              5:18, 6:13, 7:18,
              8:18, 9:13, 10:18,
              11:18, 12:13, 13:18,
              15:11, 16:14, 17:16, 18:8,
              19:12, 20:8, 21:10, 22:10, 23:10}
    for col, w in widths.items():
        ws.column_dimensions[cl(col)].width = w

    ws.freeze_panes = "A4"

    # Charts
    paste_mass   = mix["paste"]
    solids_mass  = mix["solids_in_paste"]
    clinker_mass = mix["clinker_in_paste"]

    post_cutoff_data = [rd for rd in data if rd["time_s"] / 60.0 > cutoff_min]
    if post_cutoff_data:
        max_hf_W = max(rd["heat_flow_W"] for rd in post_cutoff_data)
        hf_ymax_paste   = round(max_hf_W * 1000 / paste_mass   * 1.2, 2) if paste_mass   > 0 else None
        hf_ymax_solids  = round(max_hf_W * 1000 / solids_mass  * 1.2, 2) if solids_mass  > 0 else None
        hf_ymax_clinker = round(max_hf_W * 1000 / clinker_mass * 1.2, 2) if clinker_mass > 0 else None
    else:
        hf_ymax_paste = hf_ymax_solids = hf_ymax_clinker = None

    charts = [
        (f"{sample_name} (Paste)",   2,  5, hf_ymax_paste,   "accent2", "Time (Days)", "Heat Flow (mW/g)", f"{cl(23)}8"),
        (f"{sample_name} (Solids)",  2, 10, None,            "accent1", "Time (Days)", "Heat (J/g)",       f"{cl(31)}26"),
        (f"{sample_name} (Solids)",  2,  8, hf_ymax_solids,  "accent2", "Time (Days)", "Heat Flow (mW/g)", f"{cl(30)}8"),
        (f"{sample_name} (Paste)",   2,  7, None,            "accent1", "Time (Days)", "Heat (J/g)",       f"{cl(23)}26"),
        (f"{sample_name} (Clinker)", 2, 11, hf_ymax_clinker, "accent2", "Time (Days)", "Heat Flow (mW/g)", f"{cl(39)}8"),
        (f"{sample_name} (Clinker)", 2, 13, None,            "accent1", "Time (Days)", "Heat (J/g)",       f"{cl(39)}26"),
    ]

    for title, x_col, y_col, y_max, color, x_title, y_title, anchor in charts:
        add_chart(ws, title, x_col, y_col, DATA_START, last,
                  None, None, y_max, color, x_title, y_title, anchor)

    wb.save(output_path)
    return {
        "sample_name": sample_name,
        "data_rows": n,
        "cutoff_min": cutoff_min,
        "auto_cutoff": auto_cutoff,
    }


# ─────────────────────────────────────────────
# WEB ENTRY POINT
# ─────────────────────────────────────────────

def run_isocal(input_path, output_path, original_filename=None,
               paste=0.0, clinker=0.0, gypsum=0.0, ns=0.0,
               noh=0.0, limestone=0.0, cc=0.0, water=0.0,
               cutoff=None):
    """
    Process a raw isocalorimeter xlsx file and write the analysis to output_path.
    Raises ValueError for invalid/empty data.
    Returns a summary dict.
    """
    sample_name, data = read_raw_data(input_path, original_filename)
    if not data:
        raise ValueError("No valid data found in the uploaded file. "
                         "Check that the file contains a 'Raw data' sheet with data at t >= 60s.")

    mix = compute_mix_design(paste, clinker, gypsum, ns, noh, limestone, cc, water)
    cutoff_idx, cutoff_min, auto_cutoff = resolve_cutoff(cutoff, data)
    return build_output(sample_name, data, mix, cutoff_idx, cutoff_min, auto_cutoff, output_path)
