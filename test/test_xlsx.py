#!/usr/bin/env python3
"""Validate dist/renovation_budget_sample.xlsx structure and sample totals."""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

ROOT = Path(__file__).resolve().parent.parent
BUILD = ROOT / "build_xlsx.py"
XLSX = ROOT / "dist" / "renovation_budget_sample.xlsx"
DATA = ROOT / "sample_data" / "sample.json"

SHEET_DASH = "Executive Dashboard"
SHEET_BUDGET = "Line-Item Budget & Actuals"
SHEET_BIDS = "Contractor Bid Leveling"
SHEET_CO = "Contingency & Change Order Log"
SHEET_SCHEDULE = "Draws & Milestone Payments"
SHEET_LISTS = "Lists"

EXPECTED_SHEETS = [
    SHEET_DASH,
    SHEET_BUDGET,
    SHEET_BIDS,
    SHEET_CO,
    SHEET_SCHEDULE,
    SHEET_LISTS,
]

INPUT_FILL = "DDEBF7"


def load_sample() -> dict:
    with DATA.open(encoding="utf-8") as f:
        return json.load(f)


def cell_value(ws, coord):
    cell = ws[coord]
    if isinstance(cell, MergedCell):
        return None
    return cell.value


def iter_cells(ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            yield cell


def budget_totals(data: dict) -> tuple[int, int]:
    est = sum(row["estimate"] for row in data["budget"])
    act = sum(row["actual"] for row in data["budget"])
    return est, act


def approved_change_total(data: dict) -> int:
    return sum(
        co["amount"]
        for co in data["changeOrders"]
        if co["status"] == "Approved"
    )


def payment_total(data: dict) -> int:
    return sum(p["amount"] for p in data["payments"])


def month_key(date_str: str) -> str:
    return date_str[:7]


def monthly_totals(data: dict) -> dict[str, int]:
    totals: dict[str, int] = {}
    for p in data["payments"]:
        key = month_key(p["date"])
        totals[key] = totals.get(key, 0) + p["amount"]
    return totals


def kitchen_first_estimate(data: dict) -> int:
    for row in data["budget"]:
        if row["room"] == "Kitchen":
            return row["estimate"]
    raise AssertionError("no kitchen row")


def payment_section_rows(n_draws: int, n_payments: int) -> dict[str, int]:
    draw_last = 1 + n_draws
    data_start = draw_last + 4
    data_last = data_start + n_payments - 1
    total_row = data_last + 1
    return {"data_start": data_start, "data_last": data_last, "total_row": total_row}


def assert_no_bad_formula_refs(wb) -> None:
    risky = ("#REF", "#NAME", "XLOOKUP", "FILTER(")
    for ws in wb.worksheets:
        for cell in iter_cells(ws):
            val = cell.value
            if cell.data_type == "f" and val:
                upper = str(val).upper()
                for token in risky:
                    if token.upper() in upper or token in str(val):
                        raise AssertionError(
                            f"disallowed formula token {token!r} in {ws.title}!{cell.coordinate}: {val}"
                        )


def assert_no_excel_errors(wb) -> None:
    for ws in wb.worksheets:
        for cell in iter_cells(ws):
            val = cell.value
            if val is not None and str(val).strip().startswith("#"):
                raise AssertionError(
                    f"error value in {ws.title}!{cell.coordinate}: {val!r}"
                )


def recalc_with_libreoffice(xlsx_path: Path) -> Path | None:
    soffice = shutil.which("soffice")
    if not soffice:
        return None
    with tempfile.TemporaryDirectory(prefix="lo_recalc_") as tmp:
        subprocess.run(
            [
                soffice,
                "--headless",
                "--convert-to",
                "xlsx",
                "--outdir",
                tmp,
                str(xlsx_path),
            ],
            check=True,
            capture_output=True,
            text=True,
        )
        out = Path(tmp) / xlsx_path.name
        if not out.is_file():
            raise AssertionError(f"LibreOffice did not produce {out.name}")
        dest = xlsx_path.parent / f".lo_recalc_{xlsx_path.name}"
        dest.write_bytes(out.read_bytes())
        return dest


def assert_libreoffice_recalc(xlsx_path: Path) -> None:
    recalc_path = recalc_with_libreoffice(xlsx_path)
    if recalc_path is None:
        print("test_xlsx.py: soffice not found; skipping LibreOffice recalc test")
        return
    try:
        wb = load_workbook(recalc_path, data_only=True)
        assert_no_excel_errors(wb)
        dash = wb[SHEET_DASH]
        monthly_sum = 0
        for row in range(7, 12):
            val = cell_value(dash, f"K{row}")
            if isinstance(val, (int, float)):
                monthly_sum += val
        assert monthly_sum > 0, f"monthly spend total was {monthly_sum!r}"
    finally:
        recalc_path.unlink(missing_ok=True)


def assert_input_fill(cell) -> None:
    assert cell.protection.locked is False, f"{cell.coordinate} should be unlocked"
    fill = cell.fill
    if fill and fill.fgColor and fill.fgColor.rgb:
        rgb = fill.fgColor.rgb
        if isinstance(rgb, str) and len(rgb) >= 6:
            assert rgb[-6:].upper() == INPUT_FILL, f"{cell.coordinate} fill {rgb}"


def main() -> int:
    subprocess.run([sys.executable, str(BUILD)], check=True, cwd=ROOT)
    assert XLSX.is_file(), "xlsx output missing"

    data = load_sample()
    est_total, act_total = budget_totals(data)
    contingency = round(est_total * data["project"]["contingencyRate"])
    budget_plus = est_total + contingency
    layout = payment_section_rows(len(data["draws"]), len(data["payments"]))

    wb = load_workbook(XLSX, data_only=False)
    assert wb.sheetnames == EXPECTED_SHEETS, wb.sheetnames
    assert wb[SHEET_LISTS].sheet_state == "hidden"

    assert_no_bad_formula_refs(wb)

    budget = wb[SHEET_BUDGET]
    n = len(data["budget"])
    total_row = n + 2
    assert cell_value(budget, f"D{total_row}") == f"=SUM(D2:D{n + 1})"

    for row in range(2, n + 2):
        for col in ("A", "B", "C", "D", "E", "G"):
            assert_input_fill(budget[f"{col}{row}"])
        assert budget[f"F{row}"].protection.locked is True, f"F{row} variance formula"
        assert str(budget[f"F{row}"].value).startswith("=")

    for col in ("D", "E", "F"):
        assert budget[f"{col}{total_row}"].protection.locked is True

    dash = wb[SHEET_DASH]
    assert_input_fill(dash["B2"])
    assert_input_fill(dash["B8"])
    assert dash["B5"].protection.locked is True
    assert "SUMIFS" in str(dash["B11"].value)

    assert payment_total(data) == 52300
    assert est_total == 80600
    assert act_total == 80630
    assert approved_change_total(data) == 1270
    assert contingency == 8060
    assert budget_plus == 88660

    sched = wb[SHEET_SCHEDULE]
    d0 = layout["data_start"]
    d1 = layout["data_last"]
    tr = layout["total_row"]
    assert str(sched[f"F{tr}"].value) == f"=SUM(F{d0}:F{d1})"
    assert_input_fill(sched[f"A{d0}"])

    months = monthly_totals(data)
    assert months["2026-02"] == 15200
    assert months["2026-06"] == 8000
    assert "SUMIFS" in str(dash["K7"].value)
    assert "Draws & Milestone Payments" in str(dash["K7"].value)
    assert "DATE(VALUE(LEFT(J7" in str(dash["K7"].value)
    assert "EDATE" in str(dash["K7"].value)

    assert_libreoffice_recalc(XLSX)

    bids = wb[SHEET_BIDS]
    assert "MIN" in str(bids["H2"].value)
    low_formula = str(bids["I2"].value)
    assert "IF(" in low_formula
    assert "XLOOKUP" not in low_formula.upper()

    assert hasattr(wb[SHEET_DASH], "_charts") and len(wb[SHEET_DASH]._charts) >= 2  # noqa: SLF001

    code_gs = (ROOT / "src" / "Code.gs").read_text(encoding="utf-8")
    marker = "var SAMPLE_DATA_JSON = "
    start = code_gs.index(marker) + len(marker)
    end = code_gs.index(";\n", start)
    embedded = json.loads(json.loads(code_gs[start:end]))
    assert embedded["budget"][0]["estimate"] == data["budget"][0]["estimate"]
    assert "Executive Dashboard" in code_gs
    assert "XLOOKUP" not in code_gs.upper()

    assert kitchen_first_estimate(data) == 14200

    print("test_xlsx.py: all checks passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
