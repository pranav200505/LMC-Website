# -*- coding: utf-8 -*-
"""
Created on Fri Mar 13 16:32:12 2026

@author: Pranav
"""

"""
Isocal Automation Script
========================
Processes raw Isocalorimeter .xlsx output and generates a formatted analysis file.

Run with:
    python isocal_automation.py

The script will prompt you for all inputs interactively.

Cutoff behaviour:
    - Press Enter to use AUTO mode (120 min), highlighted ORANGE in output
    - Enter a value to specify manually, highlighted YELLOW in output
    - The cutoff cell (P10) is always editable in Excel after generation.
"""

import sys
import os
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import ScatterChart, Reference, Series


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
# INTERACTIVE INPUT
# ─────────────────────────────────────────────

def prompt_float(label, default=None):
    """Prompt for a float. If default given, pressing Enter uses it."""
    if default is not None:
        suffix = f" [{default}]: "
    else:
        suffix = ": "
    while True:
        raw = input(f"  {label}{suffix}").strip()
        if raw == "" and default is not None:
            return default
        try:
            return float(raw)
        except ValueError:
            print(f"    ✗ Please enter a number.")

def prompt_path(label):
    """Prompt for a file path, re-prompting until valid."""
    while True:
        raw = input(f"  {label}: ").strip()
        if os.path.exists(raw):
            return raw
        print(f"    ✗ File not found: {raw}")

def gather_inputs():
    print("=" * 55)
    print("=" * 55)

    print("\n[1/3] Input file")
    raw_file = prompt_path("Path to raw instrument .xlsx file")

    print("\n[2/3] Mix design (press Enter to use 0 for optional fields)")
    paste     = prompt_float("Paste mass (g)")
    clinker   = prompt_float("Clinker mass (g)")
    gypsum    = prompt_float("Gypsum mass (g)",    default=0.0)
    ns        = prompt_float("NS mass (g)",         default=0.0)
    noh       = prompt_float("NOH mass (g)",        default=0.0)
    limestone = prompt_float("Limestone mass (g)",  default=0.0)
    cc        = prompt_float("CC mass (g)",         default=0.0)
    water     = prompt_float("Water mass (g)")

    print("\n[3/3] Cutoff time")
    print("  Press Enter for AUTO (120 min), or enter a value in minutes:")
    raw = input("  Cutoff time (min) [auto=120]: ").strip()
    cutoff = float(raw) if raw != "" else None

    output_default = os.path.splitext(raw_file)[0] + "_output.xlsx"
    print(f"\n  Output file [{output_default}]:")
    raw = input("  Output path (press Enter for default): ").strip()
    output_path = raw if raw != "" else output_default

    return raw_file, paste, clinker, gypsum, ns, noh, limestone, cc, water, cutoff, output_path


# ─────────────────────────────────────────────
# RAW DATA READING
# ─────────────────────────────────────────────

def read_raw_data(raw_file):
    """
    Read raw instrument xlsx.
    Sheet 'Raw data': Col A=time(s), Col D=Heat Flow(W), Col E=Heat(J).
    Rows 1-2 are headers/units. Data starts row 3.
    Valid data: t >= 60s.
    Returns: sample_name (str), data (list of dicts)
    """
    wb = openpyxl.load_workbook(raw_file, data_only=True)

    # Extract sample name from Experiment info sheet
    sample_name = os.path.splitext(os.path.basename(raw_file))[0]
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

    # Trim trailing rows where instrument has flatlined:
    # remove from the end any rows where both HF and H are effectively zero
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

import math
from lxml import etree

def _make_rich_text(text, sz_hundredths):
    """Build a chart Text object with specified font size (in hundredths of a pt)."""
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

def _make_axis_txpr(sz_hundredths):
    """Build a RichText object for axis tick label font size."""
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import (RichTextProperties, ListStyle,
                                        Paragraph, ParagraphProperties,
                                        RegularTextRun, CharacterProperties)
    rpr  = CharacterProperties(sz=sz_hundredths)
    run  = RegularTextRun(t="", rPr=rpr)
    para = Paragraph(r=[run], pPr=ParagraphProperties())
    return RichText(bodyPr=RichTextProperties(), lstStyle=ListStyle(), p=[para])

def nearest_half_above(value):
    """Round value up to nearest 0.5."""
    return math.ceil(value * 2) / 2


def add_chart(ws, title, x_col, y_col, data_start, last_row,
              x_max, y_min, y_max, color, x_title, y_title, anchor):
    from openpyxl.drawing.colors import ColorChoice, SchemeColor
    from openpyxl.chart.title import Title

    chart = ScatterChart()
    chart.style = None

    # Chart title: 14.4pt → 1440 hundredths
    chart.title = Title(tx=_make_rich_text(title, 1440))

    # Series
    x_ref  = Reference(ws, min_col=x_col, min_row=data_start, max_row=last_row)
    y_ref  = Reference(ws, min_col=y_col, min_row=data_start, max_row=last_row)
    series = Series(y_ref, x_ref, title=title)
    series.marker.symbol = "none"
    series.graphicalProperties.line.width = 28575
    series.graphicalProperties.line.solidFill = ColorChoice(
        schemeClr=SchemeColor(val=color)
    )
    chart.series.append(series)

    # Axis titles: 12pt → 1200 hundredths
    chart.x_axis.title = Title(tx=_make_rich_text(x_title, 1200))
    chart.y_axis.title = Title(tx=_make_rich_text(y_title, 1200))

    # Note: do NOT set txPr on axes — openpyxl writes a malformed empty run
    # that causes Excel to hide tick labels in embedded view. Leave at default.

    # Scaling
    if x_max is not None:
        chart.x_axis.scaling.max = x_max
    if y_min is not None:
        chart.y_axis.scaling.min = y_min
    if y_max is not None:
        chart.y_axis.scaling.max = nearest_half_above(y_max)
        chart.y_axis.majorUnit   = 0.5

    # Axis number format: 1 decimal place
    chart.x_axis.numFmt = "0.0"
    chart.y_axis.numFmt = "0.0"

    # Tick marks: outside, black, 1pt (12700 EMU)
    chart.x_axis.majorTickMark = "out"
    chart.y_axis.majorTickMark = "out"
    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    def _black_line_sppr():
        spPr_el = etree.Element(f"{{{ns}}}spPr")
        ln_el   = etree.SubElement(spPr_el, f"{{{ns}}}ln", w="12700")
        sf_el   = etree.SubElement(ln_el,   f"{{{ns}}}solidFill")
        etree.SubElement(sf_el, f"{{{ns}}}srgbClr", val="000000")
        from openpyxl.chart.shapes import GraphicalProperties
        return GraphicalProperties.from_tree(spPr_el)
    chart.x_axis.spPr = _black_line_sppr()
    chart.y_axis.spPr = _black_line_sppr()

    # Chart border: white (FFFFFF)
    def _white_border_sppr():
        spPr_el = etree.Element(f"{{{ns}}}spPr")
        ln_el   = etree.SubElement(spPr_el, f"{{{ns}}}ln", w="12700")
        sf_el   = etree.SubElement(ln_el,   f"{{{ns}}}solidFill")
        etree.SubElement(sf_el, f"{{{ns}}}srgbClr", val="FFFFFF")
        from openpyxl.chart.shapes import GraphicalProperties
        return GraphicalProperties.from_tree(spPr_el)
    chart.plot_area.spPr = _white_border_sppr()

    # Legend at bottom
    from openpyxl.chart.legend import Legend
    leg = Legend()
    leg.position = "b"
    chart.legend = leg

    # Remove all gridlines
    chart.x_axis.majorGridlines = None
    chart.y_axis.majorGridlines = None

    chart.width  = 15
    chart.height = 7.5

    ws.add_chart(chart, anchor)


# ─────────────────────────────────────────────
# MAIN OUTPUT BUILDER
# ─────────────────────────────────────────────

DATA_START = 4   # first data row (1-indexed)

def build_output(sample_name, data, mix, cutoff_idx, cutoff_min, auto_cutoff, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = sample_name

    n     = len(data)
    last  = DATA_START + n - 1
    cutoff_row = DATA_START + cutoff_idx  # Excel row of cutoff

    # ── Row 1: sample name ──────────────────────────────────────────
    ws["A1"] = sample_name
    ws["A1"].font = Font(bold=True, size=12)

    # ── Row 2: group headers ────────────────────────────────────────
    for col, val in [(5,"Paste"), (8,"Solids"), (11,"Clinker")]:
        header_cell(ws, 2, col, val)

    # ── Row 3: column headers ───────────────────────────────────────
    col_headers = [
        "Time (min)", "Time (Days)", "Heat Flow (W)", "Heat (J)",
        "Heat Flow (mW/g)", "Heat (J/g)", "Adj Heat (J/g)",
        "Heat Flow (mW/g)", "Heat (J/g)", "Adj Heat (J/g)",
        "Heat Flow (mW/g)", "Heat (J/g)", "Adj Heat (J/g)",
    ]
    for j, h in enumerate(col_headers, 1):
        header_cell(ws, 3, j, h)

    # ── Mix design table (cols O=15 onward) ─────────────────────────
    # Row 2: component headers
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

    ws["U5"] = "=SUM(O5:T5)"   # Solids
    ws["W5"] = "=U5+V5"         # Total
    ws["U5"].border = BOX
    ws["W5"].border = BOX

    # Row 7: derived headers
    for j, h in enumerate(["Paste", "Solids (Paste)", "Clinker (Paste)"], 15):
        header_cell(ws, 7, j, h)

    # Row 8: derived values
    ws["O8"] = mix["paste"]       # paste (user input)
    ws["P8"] = "=U5*O8/W5"        # solids in paste
    ws["Q8"] = "=O5*O8/W5"        # clinker in paste
    for j in range(15, 18):
        ws.cell(row=8, column=j).border = BOX
        ws.cell(row=8, column=j).number_format = "0.000"

    # ── Cutoff cell P10 ─────────────────────────────────────────────
    ws["O10"] = "Cutoff Time (min):"
    ws["O10"].font = BOLD
    ws["O10"].alignment = Alignment(horizontal="right")

    ws["P10"].value        = cutoff_min
    ws["P10"].fill         = ORANGE_FILL if auto_cutoff else YELLOW_FILL
    ws["P10"].font         = Font(bold=True)
    ws["P10"].border       = BOX
    ws["P10"].number_format = "0.0"

    ws["Q10"] = "← AUTO (120 min) — edit to override" if auto_cutoff else "← User-specified"
    ws["Q10"].font = Font(italic=True, color="888888")

    # ── Data rows ────────────────────────────────────────────────────
    for i, rd in enumerate(data):
        r    = DATA_START + i
        t_s  = rd["time_s"]
        hfW  = rd["heat_flow_W"]
        hJ   = rd["heat_J"]
        tmin = t_s / 60.0
        tday = tmin / (24 * 60)

        ws.cell(r, 1, tmin)    # A: Time (min)
        ws.cell(r, 2, tday)    # B: Time (Days)
        ws.cell(r, 3, hfW)     # C: Heat Flow (W)
        ws.cell(r, 4, hJ)      # D: Heat (J)

        # Paste
        ws.cell(r, 5, f"=C{r}*1000/$O$8")       # E: HF mW/g paste
        ws.cell(r, 6, f"=D{r}/$O$8")             # F: Heat J/g paste
        ws.cell(r, 7, f"=IF($A{r}<=$P$10,0,F{r}-F{cutoff_row})")  # G: Adj paste

        # Solids
        ws.cell(r, 8,  f"=C{r}*1000/$P$8")       # H: HF mW/g solids
        ws.cell(r, 9,  f"=D{r}/$P$8")             # I: Heat J/g solids
        ws.cell(r, 10, f"=IF($A{r}<=$P$10,0,I{r}-I{cutoff_row})") # J: Adj solids

        # Clinker
        ws.cell(r, 11, f"=C{r}*1000/$Q$8")       # K: HF mW/g clinker
        ws.cell(r, 12, f"=D{r}/$Q$8")             # L: Heat J/g clinker
        ws.cell(r, 13, f"=IF($A{r}<=$P$10,0,L{r}-L{cutoff_row})") # M: Adj clinker

    # ── Number formats ──────────────────────────────────────────────
    for r in range(DATA_START, DATA_START + n):
        ws.cell(r, 1).number_format = "0.00"
        ws.cell(r, 2).number_format = "0.00000"
        for c in range(3, 14):
            ws.cell(r, c).number_format = "0.000000"

    # ── Column widths ───────────────────────────────────────────────
    widths = {1:12, 2:12, 3:14, 4:12,
              5:18, 6:13, 7:18,
              8:18, 9:13, 10:18,
              11:18, 12:13, 13:18,
              15:11, 16:14, 17:16, 18:8,
              19:12, 20:8, 21:10, 22:10, 23:10}
    for col, w in widths.items():
        ws.column_dimensions[cl(col)].width = w

    ws.freeze_panes = "A4"

    # ── Charts ─────────────────────────────────────────────────────
    # ── Compute y_max for Heat Flow charts (suppress initial spike) ──
    # Use max post-cutoff Heat Flow value × 1.2 as y_max for all HF charts.
    # This ignores the large dissolution spike at t=0 so the hump is visible.
    paste_mass  = mix["paste"]
    solids_mass = mix["solids_in_paste"]
    clinker_mass = mix["clinker_in_paste"]

    post_cutoff_data = [rd for rd in data if rd["time_s"] / 60.0 > cutoff_min]
    if post_cutoff_data:
        max_hf_W = max(rd["heat_flow_W"] for rd in post_cutoff_data)
        hf_ymax_paste   = round(max_hf_W * 1000 / paste_mass   * 1.2, 2) if paste_mass   > 0 else None
        hf_ymax_solids  = round(max_hf_W * 1000 / solids_mass  * 1.2, 2) if solids_mass  > 0 else None
        hf_ymax_clinker = round(max_hf_W * 1000 / clinker_mass * 1.2, 2) if clinker_mass > 0 else None
    else:
        hf_ymax_paste = hf_ymax_solids = hf_ymax_clinker = None

    # All axes fully auto-scaled to the data
    charts = [
        # (title,                      x_col, y_col, y_max,             color,     x_title,        y_title,            anchor)
        (f"{sample_name} (Paste)",    2,  5, hf_ymax_paste,   "accent2", "Time (Days)", "Heat Flow (mW/g)", f"{cl(23)}8"),
        (f"{sample_name} (Solids)",   2, 10, None,            "accent1", "Time (Days)", "Heat (J/g)",       f"{cl(31)}26"),
        (f"{sample_name} (Solids)",   2,  8, hf_ymax_solids,  "accent2", "Time (Days)", "Heat Flow (mW/g)", f"{cl(30)}8"),
        (f"{sample_name} (Paste)",    2,  7, None,            "accent1", "Time (Days)", "Heat (J/g)",       f"{cl(23)}26"),
        (f"{sample_name} (Clinker)",  2, 11, hf_ymax_clinker, "accent2", "Time (Days)", "Heat Flow (mW/g)", f"{cl(39)}8"),
        (f"{sample_name} (Clinker)",  2, 13, None,            "accent1", "Time (Days)", "Heat (J/g)",       f"{cl(39)}26"),
    ]

    for title, x_col, y_col, y_max, color, x_title, y_title, anchor in charts:
        add_chart(ws, title, x_col, y_col, DATA_START, last,
                  None, None, y_max, color, x_title, y_title, anchor)

    # ── Save ────────────────────────────────────────────────────────
    wb.save(output_path)

    print(f"\n✓ Done: {output_path}")
    print(f"  Sample name  : {sample_name}")
    print(f"  Data rows    : {n}")
    print(f"  Cutoff       : {cutoff_min} min {'[AUTO]' if auto_cutoff else '[user]'} → Excel row {cutoff_row}")
    print(f"  Paste mass   : {mix['paste']} g")
    print(f"  Solids/paste : {mix['solids_in_paste']:.4f} g  (from formula =U5*O8/W5)")
    print(f"  Clinker/paste: {mix['clinker_in_paste']:.4f} g  (from formula =O5*O8/W5)")


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────

def main():
    raw_file, paste, clinker, gypsum, ns, noh, limestone, cc, water, cutoff, output_path = gather_inputs()

    print(f"\nReading: {raw_file}")
    sample_name, data = read_raw_data(raw_file)
    print(f"  Sample: '{sample_name}'  |  {len(data)} valid data points")

    mix = compute_mix_design(
        paste=paste, clinker=clinker, gypsum=gypsum,
        ns=ns, noh=noh, limestone=limestone,
        cc=cc, water=water,
    )

    cutoff_idx, cutoff_min, auto_cutoff = resolve_cutoff(cutoff, data)

    if auto_cutoff:
        print(f"  Cutoff: AUTO ({AUTO_CUTOFF_MINUTES} min)")
    else:
        print(f"  Cutoff: {cutoff_min} min (user-specified)")

    build_output(sample_name, data, mix, cutoff_idx, cutoff_min, auto_cutoff, output_path)


if __name__ == "__main__":
    main()