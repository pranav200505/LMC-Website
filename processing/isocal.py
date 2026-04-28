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

def _default_sample_name(original_filename, raw_file):
    name_source = original_filename if original_filename else raw_file
    return os.path.splitext(os.path.basename(name_source))[0]


def _finalize_data(data):
    """Trim trailing flatlined rows (heat flow and heat both ~0)."""
    while data and abs(data[-1]["heat_flow_W"]) < 1e-5 and abs(data[-1]["heat_J"]) < 1e-3:
        data.pop()
    return data


def _read_raw_xlsx(raw_file, sample_name):
    """Read .xlsx / .xlsm via openpyxl. Returns (sample_name, data)."""
    wb = openpyxl.load_workbook(raw_file, data_only=True)

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

    if "Raw data" not in wb.sheetnames:
        raise ValueError("Expected a sheet named 'Raw data' in the uploaded file.")

    ws_raw = wb["Raw data"]
    data = []
    for row in ws_raw.iter_rows(min_row=3, values_only=True):
        if not row:
            continue
        t = row[0]
        if t is None or not isinstance(t, (int, float)):
            continue
        t = float(t)
        if t < 60:
            continue
        hf = float(row[3]) if len(row) > 3 and row[3] is not None else 0.0
        h  = float(row[4]) if len(row) > 4 and row[4] is not None else 0.0
        data.append({"time_s": t, "heat_flow_W": hf, "heat_J": h})

    return sample_name, _finalize_data(data)


def _read_raw_xls(raw_file, sample_name):
    """Read legacy .xls via xlrd. Returns (sample_name, data)."""
    try:
        import xlrd
    except ImportError as e:
        raise ValueError(
            "Legacy .xls files require the 'xlrd' package. "
            "Please install it (xlrd>=1.2.0,<2.0.0)."
        ) from e

    book = xlrd.open_workbook(raw_file)
    sheet_names = book.sheet_names()

    if "Experiment info" in sheet_names:
        ws_info = book.sheet_by_name("Experiment info")
        for row_idx in range(ws_info.nrows):
            row = ws_info.row_values(row_idx)
            for i, val in enumerate(row):
                if isinstance(val, str) and "name" in val.lower():
                    if i + 1 < len(row):
                        nv = row[i + 1]
                        if nv:
                            sample_name = str(nv).replace(".rslt", "").strip()

    if "Raw data" not in sheet_names:
        raise ValueError("Expected a sheet named 'Raw data' in the uploaded file.")

    ws_raw = book.sheet_by_name("Raw data")
    data = []
    # Data starts at 1-indexed row 3 -> 0-indexed row 2
    for row_idx in range(2, ws_raw.nrows):
        row = ws_raw.row_values(row_idx)
        if not row:
            continue
        t = row[0]
        if t == "" or t is None or not isinstance(t, (int, float)):
            continue
        t = float(t)
        if t < 60:
            continue
        hf_raw = row[3] if len(row) > 3 else ""
        h_raw  = row[4] if len(row) > 4 else ""
        hf = float(hf_raw) if isinstance(hf_raw, (int, float)) else 0.0
        h  = float(h_raw)  if isinstance(h_raw,  (int, float)) else 0.0
        data.append({"time_s": t, "heat_flow_W": hf, "heat_J": h})

    return sample_name, _finalize_data(data)


def read_raw_data(raw_file, original_filename=None):
    """
    Read raw instrument file (.xlsx, .xlsm, or .xls).

    Expected layout (same for all formats):
      - Sheet 'Raw data': Col A=time(s), Col D=Heat Flow(W), Col E=Heat(J).
      - Rows 1-2 are headers/units; data starts row 3.
      - Valid data: t >= 60s.
      - Optional 'Experiment info' sheet used to pull the sample name.

    original_filename: the user's original upload filename, used for extension
                       detection and sample_name derivation.

    Returns: (sample_name, data) where data is a list of dicts.
    """
    sample_name = _default_sample_name(original_filename, raw_file)

    ext_source = original_filename if original_filename else raw_file
    ext = os.path.splitext(ext_source)[1].lower()

    if ext in (".xlsx", ".xlsm"):
        return _read_raw_xlsx(raw_file, sample_name)
    if ext == ".xls":
        return _read_raw_xls(raw_file, sample_name)

    # Unknown extension (form validation should prevent this); try xlsx then xls
    try:
        return _read_raw_xlsx(raw_file, sample_name)
    except Exception:
        try:
            return _read_raw_xls(raw_file, sample_name)
        except Exception as e:
            raise ValueError(
                f"Unsupported file type '{ext}'. Please upload a .xlsx, .xlsm, or .xls file."
            ) from e


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

def nearest_half_above(value):
    return math.ceil(value * 2) / 2


def _pt_to_emu(pt):
    return int(round(pt * 12700))


def add_chart(ws, title, x_col, y_col, data_start, last_row,
              x_max, y_min, y_max, color, x_title, y_title, anchor,
              chart_style=None):
    from openpyxl.drawing.colors import ColorChoice, SchemeColor

    cs = chart_style or {}

    chart = ScatterChart()
    chart.title = cs.get("chart_title") or title
    chart.x_axis.title = cs.get("x_title") or x_title
    chart.y_axis.title = cs.get("y_title") or y_title

    x_ref  = Reference(ws, min_col=x_col, min_row=data_start, max_row=last_row)
    y_ref  = Reference(ws, min_col=y_col, min_row=data_start, max_row=last_row)
    series = Series(y_ref, x_ref, title=title)
    series.marker.symbol = "none"
    series.graphicalProperties.line.width = _pt_to_emu(cs.get("line_thickness_pt", 2.25))

    custom_line_color = cs.get("line_color")
    if custom_line_color:
        series.graphicalProperties.line.solidFill = custom_line_color.lstrip("#")
    else:
        series.graphicalProperties.line.solidFill = ColorChoice(
            schemeClr=SchemeColor(val=color)
        )
    chart.series.append(series)

    if x_max is not None:
        chart.x_axis.scaling.max = x_max
    if y_min is not None:
        chart.y_axis.scaling.min = y_min
    if y_max is not None:
        chart.y_axis.scaling.max = nearest_half_above(y_max)
        chart.y_axis.majorUnit   = 0.5

    chart.x_axis.tickLblPos = "low"
    chart.y_axis.tickLblPos = "low"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.majorTickMark = "out"
    chart.y_axis.majorTickMark = "out"
    chart.x_axis.minorTickMark = "none"
    chart.y_axis.minorTickMark = "none"
    chart.x_axis.majorGridlines = None
    chart.y_axis.majorGridlines = None

    chart.legend = None
    chart.width  = 15
    chart.height = 7.5

    ws.add_chart(chart, anchor)


# ─────────────────────────────────────────────
# MAIN OUTPUT BUILDER
# ─────────────────────────────────────────────

DATA_START = 4

def build_output(sample_name, data, mix, cutoff_idx, cutoff_min, auto_cutoff, output_path,
                 hf_style=None, h_style=None):
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

    # Resolve chart titles and axis labels from custom styles or defaults
    _hf = hf_style or {}
    _h  = h_style or {}

    def _hf_title(base):
        custom = _hf.get("chart_title", "").strip()
        return custom if custom else f"{sample_name} ({base})"

    def _h_title(base):
        custom = _h.get("chart_title", "").strip()
        return custom if custom else f"{sample_name} ({base})"

    hf_x = _hf.get("x_title", "Time (Days)") or "Time (Days)"
    hf_y = _hf.get("y_title", "Heat Flow (mW/g)") or "Heat Flow (mW/g)"
    h_x  = _h.get("x_title", "Time (Days)") or "Time (Days)"
    h_y  = _h.get("y_title", "Heat (J/g)") or "Heat (J/g)"

    charts = [
        (_hf_title("Paste"),   2,  5, hf_ymax_paste,   "accent2", hf_x, hf_y, f"{cl(23)}8",  _hf),
        (_h_title("Solids"),   2, 10, None,             "accent1", h_x,  h_y,  f"{cl(31)}26", _h),
        (_hf_title("Solids"),  2,  8, hf_ymax_solids,   "accent2", hf_x, hf_y, f"{cl(30)}8",  _hf),
        (_h_title("Paste"),    2,  7, None,             "accent1", h_x,  h_y,  f"{cl(23)}26", _h),
        (_hf_title("Clinker"), 2, 11, hf_ymax_clinker,  "accent2", hf_x, hf_y, f"{cl(39)}8",  _hf),
        (_h_title("Clinker"),  2, 13, None,             "accent1", h_x,  h_y,  f"{cl(39)}26", _h),
    ]

    for title, x_col, y_col, y_max, color, x_title, y_title, anchor, style in charts:
        add_chart(ws, title, x_col, y_col, DATA_START, last,
                  None, None, y_max, color, x_title, y_title, anchor,
                  chart_style=style)

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
               cutoff=None, hf_style=None, h_style=None):
    """
    Process a raw isocalorimeter xlsx file and write the analysis to output_path.
    Raises ValueError for invalid/empty data.

    hf_style / h_style: optional dicts with chart appearance overrides for
    Heat Flow and Heat charts respectively.  Keys:
        chart_title, x_title, y_title, line_thickness_pt, line_color,
        x_font_size_pt, x_font_color, y_font_size_pt, y_font_color,
        border_color
    Returns a summary dict.
    """
    sample_name, data = read_raw_data(input_path, original_filename)
    if not data:
        raise ValueError("No valid data found in the uploaded file. "
                         "Check that the file contains a 'Raw data' sheet with data at t >= 60s.")

    mix = compute_mix_design(paste, clinker, gypsum, ns, noh, limestone, cc, water)
    cutoff_idx, cutoff_min, auto_cutoff = resolve_cutoff(cutoff, data)
    return build_output(sample_name, data, mix, cutoff_idx, cutoff_min, auto_cutoff, output_path,
                        hf_style=hf_style, h_style=h_style)
