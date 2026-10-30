#!/usr/bin/env python3
"""Build dist/renovation_budget_sample.xlsx from sample_data/sample.json."""

from __future__ import annotations

import json
import re
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent
DATA_PATH = ROOT / "sample_data" / "sample.json"
OUT_DIR = ROOT / "dist"
OUT_PATH = OUT_DIR / "renovation_budget_sample.xlsx"

SHEET_DASH = "Executive Dashboard"
SHEET_BUDGET = "Line-Item Budget & Actuals"
SHEET_BIDS = "Contractor Bid Leveling"
SHEET_CO = "Contingency & Change Order Log"
SHEET_SCHEDULE = "Draws & Milestone Payments"
SHEET_LISTS = "Lists"

SHEETS = [
    SHEET_DASH,
    SHEET_BUDGET,
    SHEET_BIDS,
    SHEET_CO,
    SHEET_SCHEDULE,
    SHEET_LISTS,
]

HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
INPUT_FILL = PatternFill("solid", fgColor="DDEBF7")
TITLE_FONT = Font(bold=True, size=14)
HEADER_FONT = Font(bold=True)
UNLOCKED = Protection(locked=False)
LOCKED = Protection(locked=True)


def load_data() -> dict:
    with DATA_PATH.open(encoding="utf-8") as f:
        return json.load(f)


def parse_iso_date(value: str) -> date:
    return date.fromisoformat(value)


def sheet_q(name: str) -> str:
    escaped = name.replace("'", "''")
    return f"'{escaped}'"


def set_input(cell) -> None:
    cell.protection = UNLOCKED
    cell.fill = INPUT_FILL


def set_formula(cell, formula: str) -> None:
    cell.value = formula
    cell.protection = LOCKED


def col_widths(ws, widths: dict[str, float]) -> None:
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def budget_last_row(n_items: int) -> int:
    return 1 + n_items


def payment_section_rows(n_draws: int, n_payments: int) -> dict[str, int]:
    draw_last = 1 + n_draws
    label_row = draw_last + 2
    header_row = draw_last + 3
    data_start = draw_last + 4
    data_last = data_start + n_payments - 1
    total_row = data_last + 1
    return {
        "draw_last": draw_last,
        "label_row": label_row,
        "header_row": header_row,
        "data_start": data_start,
        "data_last": data_last,
        "total_row": total_row,
    }


def write_lists(wb: Workbook, data: dict) -> None:
    ws = wb[SHEET_LISTS]
    lists = data["lists"]
    headers = ["Rooms", "Trades", "Statuses", "Vendors"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.protection = LOCKED
    keys = ["rooms", "trades", "statuses", "vendors"]
    max_len = max(len(lists[k]) for k in keys)
    for r in range(max_len):
        row = r + 2
        for ci, key in enumerate(keys, start=1):
            vals = lists[key]
            val = vals[r] if r < len(vals) else None
            cell = ws.cell(row=row, column=ci, value=val)
            cell.protection = LOCKED
    col_widths(ws, {"A": 16, "B": 14, "C": 12, "D": 22})
    ws.sheet_state = "hidden"


def write_budget(wb: Workbook, data: dict) -> tuple[int, int]:
    ws = wb[SHEET_BUDGET]
    bq = sheet_q(SHEET_BUDGET)
    headers = ["Room", "Trade", "Line item", "Estimate", "Actual", "Variance", "Notes"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.protection = LOCKED
    items = data["budget"]
    for idx, item in enumerate(items):
        row = idx + 2
        set_input(ws.cell(row=row, column=1, value=item["room"]))
        set_input(ws.cell(row=row, column=2, value=item["trade"]))
        set_input(ws.cell(row=row, column=3, value=item["lineItem"]))
        set_input(ws.cell(row=row, column=4, value=item["estimate"]))
        set_input(ws.cell(row=row, column=5, value=item["actual"]))
        set_formula(ws.cell(row=row, column=6), f"=E{row}-D{row}")
        set_input(ws.cell(row=row, column=7, value=""))
        for col in (4, 5, 6):
            ws.cell(row=row, column=col).number_format = "#,##0"
    last = budget_last_row(len(items))
    total_row = last + 1
    ws.cell(row=total_row, column=3, value="Total").font = HEADER_FONT
    set_formula(ws.cell(row=total_row, column=4), f"=SUM(D2:D{last})")
    set_formula(ws.cell(row=total_row, column=5), f"=SUM(E2:E{last})")
    set_formula(ws.cell(row=total_row, column=6), f"=E{total_row}-D{total_row}")
    for col in (4, 5, 6):
        ws.cell(row=total_row, column=col).number_format = "#,##0"
        ws.cell(row=total_row, column=col).font = HEADER_FONT

    room_dv = DataValidation(
        type="list",
        formula1=f"={sheet_q(SHEET_LISTS)}!$A$2:$A$6",
        allow_blank=False,
    )
    trade_dv = DataValidation(
        type="list",
        formula1=f"={sheet_q(SHEET_LISTS)}!$B$2:$B$7",
        allow_blank=False,
    )
    ws.add_data_validation(room_dv)
    ws.add_data_validation(trade_dv)
    room_dv.add(f"A2:A{last}")
    trade_dv.add(f"B2:B{last}")

    red_fill = PatternFill("solid", fgColor="FFC7CE")
    ws.conditional_formatting.add(
        f"E2:E{last}",
        FormulaRule(formula=[f"E2>D2"], fill=red_fill),
    )
    col_widths(ws, {"A": 14, "B": 12, "C": 36, "D": 12, "E": 12, "F": 12, "G": 18})
    return last, total_row


def write_bids(wb: Workbook, data: dict) -> None:
    ws = wb[SHEET_BIDS]
    headers = [
        "Trade",
        "Bidder A",
        "Amount A",
        "Bidder B",
        "Amount B",
        "Bidder C",
        "Amount C",
        "Low bid",
        "Low bidder",
    ]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.protection = LOCKED
    items = data["bids"]
    last = 1 + len(items)
    for idx, item in enumerate(items):
        row = idx + 2
        set_input(ws.cell(row=row, column=1, value=item["trade"]))
        set_input(ws.cell(row=row, column=2, value=item["bidderA"]))
        set_input(ws.cell(row=row, column=3, value=item["amountA"]))
        set_input(ws.cell(row=row, column=4, value=item["bidderB"]))
        set_input(ws.cell(row=row, column=5, value=item["amountB"]))
        set_input(ws.cell(row=row, column=6, value=item["bidderC"]))
        set_input(ws.cell(row=row, column=7, value=item["amountC"]))
        set_formula(ws.cell(row=row, column=8), f"=MIN(C{row},E{row},G{row})")
        set_formula(
            ws.cell(row=row, column=9),
            f"=IF(H{row}=C{row},B{row},IF(H{row}=E{row},D{row},F{row}))",
        )
        for col in (3, 5, 7, 8):
            ws.cell(row=row, column=col).number_format = "#,##0"

    green = PatternFill("solid", fgColor="C6EFCE")
    for r in range(2, last + 1):
        ws.conditional_formatting.add(
            f"C{r}",
            CellIsRule(operator="equal", formula=[f"$H{r}"], fill=green),
        )
        ws.conditional_formatting.add(
            f"E{r}",
            CellIsRule(operator="equal", formula=[f"$H{r}"], fill=green),
        )
        ws.conditional_formatting.add(
            f"G{r}",
            CellIsRule(operator="equal", formula=[f"$H{r}"], fill=green),
        )
    col_widths(
        ws,
        {"A": 12, "B": 22, "C": 12, "D": 22, "E": 12, "F": 22, "G": 12, "H": 12, "I": 22},
    )


def write_change_orders(wb: Workbook, data: dict) -> None:
    ws = wb[SHEET_CO]
    headers = ["CO ID", "Date", "Room", "Trade", "Description", "Amount", "Status"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.protection = LOCKED
    items = data["changeOrders"]
    last = 1 + len(items)
    for idx, item in enumerate(items):
        row = idx + 2
        set_input(ws.cell(row=row, column=1, value=item["id"]))
        set_input(ws.cell(row=row, column=2, value=item["date"]))
        set_input(ws.cell(row=row, column=3, value=item["room"]))
        set_input(ws.cell(row=row, column=4, value=item["trade"]))
        set_input(ws.cell(row=row, column=5, value=item["description"]))
        set_input(ws.cell(row=row, column=6, value=item["amount"]))
        set_input(ws.cell(row=row, column=7, value=item["status"]))
        ws.cell(row=row, column=6).number_format = "#,##0"
    status_dv = DataValidation(type="list", formula1=f"={sheet_q(SHEET_LISTS)}!$C$2:$C$5")
    room_dv = DataValidation(type="list", formula1=f"={sheet_q(SHEET_LISTS)}!$A$2:$A$6")
    trade_dv = DataValidation(type="list", formula1=f"={sheet_q(SHEET_LISTS)}!$B$2:$B$7")
    for dv in (status_dv, room_dv, trade_dv):
        ws.add_data_validation(dv)
    status_dv.add(f"G2:G{last}")
    room_dv.add(f"C2:C{last}")
    trade_dv.add(f"D2:D{last}")
    col_widths(ws, {"A": 10, "B": 12, "C": 14, "D": 12, "E": 34, "F": 10, "G": 12})


def write_schedule(wb: Workbook, data: dict, budget_total_row: int) -> dict[str, int]:
    ws = wb[SHEET_SCHEDULE]
    sq = sheet_q(SHEET_SCHEDULE)
    bq = sheet_q(SHEET_BUDGET)
    draws = data["draws"]
    payments = data["payments"]
    layout = payment_section_rows(len(draws), len(payments))

    draw_headers = ["Draw", "Milestone", "Pct of budget", "Scheduled amount", "Status"]
    for i, h in enumerate(draw_headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.protection = LOCKED
    for idx, item in enumerate(draws):
        row = idx + 2
        set_input(ws.cell(row=row, column=1, value=item["draw"]))
        set_input(ws.cell(row=row, column=2, value=item["milestone"]))
        c_pct = ws.cell(row=row, column=3, value=item["pctOfBudget"])
        set_input(c_pct)
        c_pct.number_format = "0%"
        set_formula(
            ws.cell(row=row, column=4),
            f"=ROUND(C{row}*{bq}!D${budget_total_row},0)",
        )
        ws.cell(row=row, column=4).number_format = "#,##0"
        set_input(ws.cell(row=row, column=5, value=item["status"]))

    draw_last = layout["draw_last"]
    status_dv = DataValidation(type="list", formula1=f"={sheet_q(SHEET_LISTS)}!$C$2:$C$5")
    ws.add_data_validation(status_dv)
    status_dv.add(f"E2:E{draw_last}")

    ws.cell(row=layout["label_row"], column=1, value="Payment log").font = HEADER_FONT
    ws.cell(row=layout["label_row"], column=1).protection = LOCKED

    pay_headers = ["Date", "Vendor", "Draw", "Room", "Trade", "Amount", "Memo"]
    hr = layout["header_row"]
    for i, h in enumerate(pay_headers, start=1):
        c = ws.cell(row=hr, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.protection = LOCKED

    for idx, item in enumerate(payments):
        row = layout["data_start"] + idx
        date_cell = ws.cell(row=row, column=1, value=parse_iso_date(item["date"]))
        set_input(date_cell)
        date_cell.number_format = "yyyy-mm-dd"
        set_input(ws.cell(row=row, column=2, value=item["vendor"]))
        set_input(ws.cell(row=row, column=3, value=item["draw"]))
        set_input(ws.cell(row=row, column=4, value=item["room"]))
        set_input(ws.cell(row=row, column=5, value=item["trade"]))
        set_input(ws.cell(row=row, column=6, value=item["amount"]))
        set_input(ws.cell(row=row, column=7, value=item["memo"]))
        ws.cell(row=row, column=6).number_format = "#,##0"

    vendor_dv = DataValidation(type="list", formula1=f"={sheet_q(SHEET_LISTS)}!$D$2:$D$6")
    room_dv = DataValidation(type="list", formula1=f"={sheet_q(SHEET_LISTS)}!$A$2:$A$6")
    trade_dv = DataValidation(type="list", formula1=f"={sheet_q(SHEET_LISTS)}!$B$2:$B$7")
    for dv in (vendor_dv, room_dv, trade_dv):
        ws.add_data_validation(dv)
    d0 = layout["data_start"]
    d1 = layout["data_last"]
    vendor_dv.add(f"B{d0}:B{d1}")
    room_dv.add(f"D{d0}:D{d1}")
    trade_dv.add(f"E{d0}:E{d1}")

    tr = layout["total_row"]
    ws.cell(row=tr, column=5, value="Total paid").font = HEADER_FONT
    ws.cell(row=tr, column=5).protection = LOCKED
    set_formula(ws.cell(row=tr, column=6), f"=SUM(F{d0}:F{d1})")
    ws.cell(row=tr, column=6).number_format = "#,##0"
    ws.cell(row=tr, column=6).font = HEADER_FONT

    col_widths(ws, {"A": 12, "B": 22, "C": 8, "D": 14, "E": 12, "F": 12, "G": 28})
    return layout


def month_start_expr(month_cell: str) -> str:
    """First day of YYYY-MM in month_cell; works in Excel and LibreOffice."""
    return (
        f"DATE(VALUE(LEFT({month_cell},4)),VALUE(MID({month_cell},6,2)),1)"
    )


def monthly_spend_formula(
    sched: str,
    pay_start: int,
    pay_last: int,
    month_cell: str,
) -> str:
    start = month_start_expr(month_cell)
    return (
        f"=SUMIFS({sched}!$F${pay_start}:$F${pay_last},"
        f"{sched}!$A${pay_start}:$A${pay_last},\">=\"&{start},"
        f"{sched}!$A${pay_start}:$A${pay_last},\"<\"&EDATE({start},1))"
    )


def write_dashboard(
    wb: Workbook,
    data: dict,
    budget_last: int,
    budget_total_row: int,
    schedule_layout: dict[str, int],
) -> None:
    ws = wb[SHEET_DASH]
    bq = sheet_q(SHEET_BUDGET)
    co = sheet_q(SHEET_CO)
    sq = sheet_q(SHEET_SCHEDULE)
    pay_start = schedule_layout["data_start"]
    pay_last = schedule_layout["data_last"]

    project = data["project"]["name"]
    ws["A1"] = "Renovation budget dashboard"
    ws["A1"].font = TITLE_FONT
    ws["A1"].protection = LOCKED
    ws["A2"] = "Project"
    set_input(ws["B2"])
    ws["B2"] = project
    ws["A3"] = "Author"
    ws["B3"] = "reypage-tech"
    ws["B3"].protection = LOCKED

    labels = [
        ("A5", "Total estimate"),
        ("A6", "Total actual"),
        ("A7", "Variance (actual - estimate)"),
        ("A8", "Contingency rate (editable)"),
        ("A9", "Contingency amount"),
        ("A10", "Budget plus contingency"),
        ("A11", "Approved change orders"),
        ("A12", "Total paid to date"),
        ("A13", "Remaining (budget+contingency - paid)"),
    ]
    for addr, text in labels:
        ws[addr] = text
        ws[addr].font = Font(bold=True)
        ws[addr].protection = LOCKED

    set_formula(ws["B5"], f"={bq}!D{budget_total_row}")
    set_formula(ws["B6"], f"={bq}!E{budget_total_row}")
    set_formula(ws["B7"], "=B6-B5")
    ws["B8"] = data["project"]["contingencyRate"]
    ws["B8"].number_format = "0%"
    set_input(ws["B8"])
    set_formula(ws["B9"], "=ROUND(B5*B8,0)")
    set_formula(ws["B10"], "=B5+B9")
    set_formula(
        ws["B11"],
        f'=SUMIFS({co}!$F$2:$F$20,{co}!$G$2:$G$20,"Approved")',
    )
    set_formula(ws["B12"], f"=SUM({sq}!F{pay_start}:F{pay_last})")
    set_formula(ws["B13"], "=B10-B12")
    for addr in ("B5", "B6", "B7", "B9", "B10", "B11", "B12", "B13"):
        ws[addr].number_format = "#,##0"

    ws["D5"] = "Estimate by room (summary)"
    ws["D5"].font = HEADER_FONT
    room_headers = ["Room", "Estimate", "Actual", "Variance"]
    for i, h in enumerate(room_headers, start=4):
        c = ws.cell(row=6, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.protection = LOCKED
    rooms = data["lists"]["rooms"]
    summary_start = 7
    for idx, room in enumerate(rooms):
        row = summary_start + idx
        ws.cell(row=row, column=4, value=room).protection = LOCKED
        set_formula(
            ws.cell(row=row, column=5),
            f"=SUMIFS({bq}!$D$2:$D${budget_last},{bq}!$A$2:$A${budget_last},D{row})",
        )
        set_formula(
            ws.cell(row=row, column=6),
            f"=SUMIFS({bq}!$E$2:$E${budget_last},{bq}!$A$2:$A${budget_last},D{row})",
        )
        set_formula(ws.cell(row=row, column=7), f"=F{row}-E{row}")
        for col in (5, 6, 7):
            ws.cell(row=row, column=col).number_format = "#,##0"

    summary_end = summary_start + len(rooms) - 1

    bar = BarChart()
    bar.type = "col"
    bar.title = "Estimate vs actual by room"
    bar.y_axis.title = "USD"
    data_ref = Reference(ws, min_col=5, min_row=6, max_row=summary_end, max_col=6)
    cats = Reference(ws, min_col=4, min_row=summary_start, max_row=summary_end)
    bar.add_data(data_ref, titles_from_data=True)
    bar.set_categories(cats)
    bar.height = 9
    bar.width = 11
    ws.add_chart(bar, "D15")

    ws["J5"] = "Monthly spend"
    ws["J5"].font = HEADER_FONT
    ws["J5"].protection = LOCKED
    for i, h in enumerate(["Month", "Spend"], start=10):
        c = ws.cell(row=6, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.protection = LOCKED
    months = ["2026-02", "2026-03", "2026-04", "2026-05", "2026-06"]
    for idx, month in enumerate(months):
        row = 7 + idx
        ws.cell(row=row, column=10, value=month).protection = LOCKED
        set_formula(
            ws.cell(row=row, column=11),
            monthly_spend_formula(sq, pay_start, pay_last, f"J{row}"),
        )
        ws.cell(row=row, column=11).number_format = "#,##0"

    line = BarChart()
    line.type = "col"
    line.title = "Monthly spend"
    line.y_axis.title = "USD"
    spend_ref = Reference(ws, min_col=11, min_row=6, max_row=11, max_col=11)
    month_ref = Reference(ws, min_col=10, min_row=7, max_row=11)
    line.add_data(spend_ref, titles_from_data=True)
    line.set_categories(month_ref)
    line.height = 9
    line.width = 11
    ws.add_chart(line, "M15")

    col_widths(
        ws,
        {"A": 42, "B": 16, "D": 22, "E": 12, "F": 12, "G": 12, "J": 14, "K": 12},
    )


def protect_sheet(ws, password: str | None = None) -> None:
    ws.protection.sheet = True
    ws.protection.enableObjects = False
    ws.protection.enableScenarios = False
    if password:
        ws.protection.set_password(password)


def apply_protection(wb: Workbook) -> None:
    for name in SHEETS:
        protect_sheet(wb[name])


def build_workbook(data: dict) -> Workbook:
    wb = Workbook()
    default = wb.active
    wb.remove(default)
    for name in SHEETS:
        wb.create_sheet(name)
    budget_last, budget_total_row = write_budget(wb, data)
    write_lists(wb, data)
    write_bids(wb, data)
    write_change_orders(wb, data)
    schedule_layout = write_schedule(wb, data, budget_total_row)
    write_dashboard(wb, data, budget_last, budget_total_row, schedule_layout)
    apply_protection(wb)
    return wb


def sync_code_gs(data: dict) -> None:
    """Embed sample_data JSON into src/Code.gs for Google Sheets setup."""
    code_path = ROOT / "src" / "Code.gs"
    text = code_path.read_text(encoding="utf-8")
    marker_start = "// SAMPLE_DATA_JSON_START"
    marker_end = "// SAMPLE_DATA_JSON_END"
    payload = json.dumps(data, ensure_ascii=True)
    injected = (
        f"{marker_start}\n"
        f"var SAMPLE_DATA_JSON = {json.dumps(payload)};\n"
        f"{marker_end}"
    )
    pattern = re.compile(
        re.escape(marker_start) + r".*?" + re.escape(marker_end),
        re.DOTALL,
    )
    if not pattern.search(text):
        raise SystemExit("Code.gs missing SAMPLE_DATA_JSON markers")
    text = pattern.sub(injected, text)
    code_path.write_text(text, encoding="utf-8")


def main() -> None:
    data = load_data()
    sync_code_gs(data)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb = build_workbook(data)
    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH}")


if __name__ == "__main__":
    main()
