#!/usr/bin/env python3
"""
Project Breakdowns — build_summary.py
Generates a presale pricing summary Excel workbook from structured JSON data.

Usage:
    python build_summary.py --data project_data.json --output "Harth_Pricing_Summary.xlsx"
"""

import argparse
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side
)
from openpyxl.utils import get_column_letter


# ── Styles ──────────────────────────────────────────────────────────────────

YELLOW     = PatternFill("solid", fgColor="FFFF00")
LIGHT_GREY = PatternFill("solid", fgColor="D9D9D9")
HEADER_BG  = PatternFill("solid", fgColor="404040")

HEADER_FONT  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
BOLD_FONT    = Font(name="Calibri", bold=True, size=10)
NORMAL_FONT  = Font(name="Calibri", size=10)
TITLE_FONT   = Font(name="Calibri", bold=True, size=13)

THIN_SIDE  = Side(style="thin")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

CENTER = Alignment(horizontal="center", vertical="center")
LEFT   = Alignment(horizontal="left",   vertical="center")
RIGHT  = Alignment(horizontal="right",  vertical="center")


def money(v):
    return f"${v:,.0f}" if v else ""


def apply_header(cell, text):
    cell.value = text
    cell.font = HEADER_FONT
    cell.fill = HEADER_BG
    cell.alignment = CENTER
    cell.border = THIN_BORDER


def apply_data(cell, value, align=LEFT, number_format=None, fill=None, bold=False):
    cell.value = value
    cell.font = BOLD_FONT if bold else NORMAL_FONT
    cell.alignment = align
    cell.border = THIN_BORDER
    if number_format:
        cell.number_format = number_format
    if fill:
        cell.fill = fill


# ── Summary Sheet ────────────────────────────────────────────────────────────

SUMMARY_HEADERS = [
    "Unit Type",
    "Plan Name(s)",
    "Count",
    "Min Interior SF",
    "Max Interior SF",
    "Starting Price",
    "Est. Top-End Price",
    "Starting PPSF",
]

COL_WIDTHS_SUMMARY = [18, 28, 9, 16, 16, 18, 20, 14]


def build_summary_sheet(ws, project_name: str, rows: list[dict]):
    ws.title = "Summary"

    # Title
    ws.merge_cells("A1:H1")
    title_cell = ws["A1"]
    title_cell.value = project_name
    title_cell.font = TITLE_FONT
    title_cell.alignment = CENTER
    title_cell.fill = LIGHT_GREY

    # Subtitle
    ws.merge_cells("A2:H2")
    ws["A2"].value = "Pricing Summary — Unit Type Breakdown"
    ws["A2"].font = BOLD_FONT
    ws["A2"].alignment = CENTER

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 16

    # Blank row
    ws.row_dimensions[3].height = 6

    # Column widths
    for i, w in enumerate(COL_WIDTHS_SUMMARY, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Header row (row 4)
    for col_idx, header in enumerate(SUMMARY_HEADERS, start=1):
        apply_header(ws.cell(row=4, column=col_idx), header)
    ws.row_dimensions[4].height = 20

    # Data rows (starting row 5)
    data_start = 5
    for row_num, r in enumerate(rows, start=data_start):
        count = r.get("count") or 0
        min_sf = r.get("min_sf")
        max_sf = r.get("max_sf")
        start_price = r.get("starting_price")
        top_price = r.get("top_end_price")

        apply_data(ws.cell(row=row_num, column=1), r.get("unit_type", ""))
        apply_data(ws.cell(row=row_num, column=2), r.get("plan_names", ""))
        # Count cell — yellow (editable input)
        c_count = ws.cell(row=row_num, column=3)
        c_count.value = count
        c_count.font = NORMAL_FONT
        c_count.fill = YELLOW
        c_count.alignment = CENTER
        c_count.border = THIN_BORDER

        apply_data(ws.cell(row=row_num, column=4), min_sf, align=CENTER)
        apply_data(ws.cell(row=row_num, column=5), max_sf, align=CENTER)

        # Starting price
        c_sp = ws.cell(row=row_num, column=6)
        c_sp.value = start_price
        c_sp.number_format = '"$"#,##0'
        c_sp.font = NORMAL_FONT
        c_sp.alignment = RIGHT
        c_sp.border = THIN_BORDER

        # Top-end price
        c_tp = ws.cell(row=row_num, column=7)
        c_tp.value = top_price
        c_tp.number_format = '"$"#,##0'
        c_tp.font = NORMAL_FONT
        c_tp.alignment = RIGHT
        c_tp.border = THIN_BORDER

        # Starting PPSF = Starting Price / Min SF (formula)
        c_psf = ws.cell(row=row_num, column=8)
        if min_sf and start_price:
            c_psf.value = f"=IF(D{row_num}>0,F{row_num}/D{row_num},\"-\")"
        else:
            c_psf.value = "-"
        c_psf.number_format = '"$"#,##0'
        c_psf.font = NORMAL_FONT
        c_psf.alignment = RIGHT
        c_psf.border = THIN_BORDER

        ws.row_dimensions[row_num].height = 18

    # Weighted average PPSF footer
    footer_row = data_start + len(rows)
    ws.row_dimensions[footer_row].height = 20

    apply_data(ws.cell(row=footer_row, column=1), "WEIGHTED AVERAGE PPSF", bold=True, fill=LIGHT_GREY)
    ws.cell(row=footer_row, column=1).alignment = RIGHT
    for col in range(2, 8):
        ws.cell(row=footer_row, column=col).fill = LIGHT_GREY
        ws.cell(row=footer_row, column=col).border = THIN_BORDER

    # SUMPRODUCT formula: counts × starting PPSF / total count
    c_range = f"C{data_start}:C{footer_row - 1}"
    h_range = f"H{data_start}:H{footer_row - 1}"
    avg_cell = ws.cell(row=footer_row, column=8)
    avg_cell.value = f"=IF(SUMPRODUCT({c_range},{h_range})>0,SUMPRODUCT({c_range},{h_range})/SUM({c_range}),\"-\")"
    avg_cell.number_format = '"$"#,##0'
    avg_cell.font = BOLD_FONT
    avg_cell.alignment = RIGHT
    avg_cell.border = THIN_BORDER
    avg_cell.fill = LIGHT_GREY

    # Yellow note below
    note_row = footer_row + 1
    ws.merge_cells(f"A{note_row}:H{note_row}")
    ws[f"A{note_row}"].value = "* Yellow cells (Count) are editable — update to recalculate the weighted average PPSF automatically."
    ws[f"A{note_row}"].font = Font(name="Calibri", size=9, italic=True, color="666666")
    ws[f"A{note_row}"].fill = PatternFill("solid", fgColor="FFFACD")


# ── Plan Detail Sheet ────────────────────────────────────────────────────────

DETAIL_HEADERS = [
    "Unit Type",
    "Plan Name",
    "Count",
    "Interior SF",
    "Starting Price",
    "Starting PPSF",
]

COL_WIDTHS_DETAIL = [18, 16, 9, 14, 18, 14]


def build_detail_sheet(ws, project_name: str, rows: list[dict]):
    ws.title = "Plan Detail"

    ws.merge_cells("A1:F1")
    title_cell = ws["A1"]
    title_cell.value = f"{project_name} — Plan Detail"
    title_cell.font = TITLE_FONT
    title_cell.alignment = CENTER
    title_cell.fill = LIGHT_GREY
    ws.row_dimensions[1].height = 24

    for i, w in enumerate(COL_WIDTHS_DETAIL, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for col_idx, header in enumerate(DETAIL_HEADERS, start=1):
        apply_header(ws.cell(row=2, column=col_idx), header)
    ws.row_dimensions[2].height = 20

    detail_row = 3
    for r in rows:
        # Expand plan_names into individual rows if comma-separated
        plan_names_raw = r.get("plan_names", "")
        plans = [p.strip() for p in plan_names_raw.split(",")] if plan_names_raw else [""]

        min_sf = r.get("min_sf")
        max_sf = r.get("max_sf")
        start_price = r.get("starting_price")
        count = r.get("count", 0)
        unit_type = r.get("unit_type", "")

        # For multi-plan types, distribute count evenly (rough split)
        per_plan_count = round(count / len(plans)) if plans else count

        for plan in plans:
            apply_data(ws.cell(row=detail_row, column=1), unit_type)
            apply_data(ws.cell(row=detail_row, column=2), plan)
            c_count = ws.cell(row=detail_row, column=3)
            c_count.value = per_plan_count
            c_count.font = NORMAL_FONT
            c_count.fill = YELLOW
            c_count.alignment = CENTER
            c_count.border = THIN_BORDER

            # SF — show range if min != max, else single value
            if min_sf and max_sf and min_sf != max_sf:
                sf_display = f"{min_sf}–{max_sf}"
            else:
                sf_display = min_sf or ""
            apply_data(ws.cell(row=detail_row, column=4), sf_display, align=CENTER)

            c_sp = ws.cell(row=detail_row, column=5)
            c_sp.value = start_price
            c_sp.number_format = '"$"#,##0'
            c_sp.font = NORMAL_FONT
            c_sp.alignment = RIGHT
            c_sp.border = THIN_BORDER

            c_psf = ws.cell(row=detail_row, column=6)
            if min_sf and start_price:
                psf = round(start_price / min_sf)
                c_psf.value = psf
            else:
                c_psf.value = ""
            c_psf.number_format = '"$"#,##0'
            c_psf.font = NORMAL_FONT
            c_psf.alignment = RIGHT
            c_psf.border = THIN_BORDER

            ws.row_dimensions[detail_row].height = 18
            detail_row += 1


# ── Main ─────────────────────────────────────────────────────────────────────

def build(data_path: str, output_path: str):
    with open(data_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    project_name = data.get("project_name", "Project")
    rows = data.get("rows", [])

    wb = Workbook()
    ws_summary = wb.active
    build_summary_sheet(ws_summary, project_name, rows)

    ws_detail = wb.create_sheet()
    build_detail_sheet(ws_detail, project_name, rows)

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Build a presale pricing summary Excel workbook.")
    parser.add_argument("--data",   required=True, help="Path to JSON data file")
    parser.add_argument("--output", required=True, help="Output .xlsx file path")
    args = parser.parse_args()
    build(args.data, args.output)
