"""Build starter-pack/claims-tracker-90-day.xlsx from the tested rules."""

from __future__ import annotations

from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from rules import DEFAULT_BUFFER_DAYS, OLDER_SUPPORT_CAP, RULE_START

OUT = Path(__file__).resolve().parent / "claims-tracker-90-day.xlsx"

HEADER_ROW = 4
FIRST_DATA = 5
LAST_DATA = 104

HEADERS = [
    "Participant code",
    "Plan management",
    "Who usually claims (auto)",
    "Support date",
    "Claim-by date (guide only)",
    "Your target date (auto)",
    "Days left (auto)",
    "Status",
    "Alert (auto)",
    "Submission date",
    "Reference / confirmation",
    "Follow-up needed",
    "Where the record is kept",
]

INPUT_COLS = {1, 2, 4, 8, 10, 11, 12, 13}
DATE_COLS = {4, 5, 6, 10}
FORMULA_COLS = {3, 5, 6, 7, 9}

DISCLAIMER = (
    "Independent decision-support resource. Not affiliated with, endorsed by or "
    "produced by the NDIA, the NDIS Quality and Safeguards Commission or the "
    "Australian Government. General information only, not legal, tax, financial "
    "or compliance advice. This is a planning aid. It does not calculate a "
    "personalised legal deadline."
)

EXAMPLES = [
    ("EXAMPLE-1", "NDIA-managed", date(2026, 9, 2)),
    ("EXAMPLE-2", "Plan-managed", date(2026, 9, 3)),
    ("EXAMPLE-3", "Self-managed", date(2026, 12, 2)),
]


def _fills():
    return {
        "navy": PatternFill("solid", fgColor="1F4E79"),
        "input": PatternFill("solid", fgColor="FFF8E7"),
        "auto": PatternFill("solid", fgColor="E7EEF5"),
        "banner": PatternFill("solid", fgColor="F4F1EA"),
        "example": PatternFill("solid", fgColor="FFF4D6"),
        "red": PatternFill("solid", fgColor="F4C7C3"),
        "amber": PatternFill("solid", fgColor="FCE4B3"),
        "green": PatternFill("solid", fgColor="D9EAD3"),
        "grey": PatternFill("solid", fgColor="E6E6E6"),
    }


def build(path: Path = OUT) -> Path:
    fills = _fills()
    thin = Border(
        left=Side(style="thin", color="D0D7DE"),
        right=Side(style="thin", color="D0D7DE"),
        top=Side(style="thin", color="D0D7DE"),
        bottom=Side(style="thin", color="D0D7DE"),
    )
    wb = Workbook()
    ws = wb.active
    ws.title = "Tracker"
    _settings(wb, fills)
    _tracker(ws, fills, thin)
    _how_to(wb, fills)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def _settings(wb: Workbook, fills: dict) -> None:
    ws = wb.create_sheet("Settings", 0)
    ws["A1"] = "Safety buffer (days)"
    ws["B1"] = DEFAULT_BUFFER_DAYS
    ws["A2"] = "90-day rule starts"
    ws["B2"] = RULE_START
    ws["A3"] = "Guide-date cap for supports before the rule starts"
    ws["B3"] = OLDER_SUPPORT_CAP
    ws["A5"] = (
        "Your target date is the guide date minus the safety buffer. "
        "Raise the buffer to be reminded earlier. The guide date itself does not move "
        "when it falls on a weekend or public holiday."
    )
    ws["A5"].alignment = Alignment(wrap_text=True)
    for row in (2, 3):
        ws[f"B{row}"].number_format = "DD MMM YYYY"
    ws["B1"].fill = fills["input"]
    ws["A1"].font = Font(bold=True)
    ws.column_dimensions["A"].width = 62
    ws.column_dimensions["B"].width = 18
    ws.row_dimensions[5].height = 48
    ws.sheet_properties.tabColor = "1F4E79"


def _formula(row: int) -> dict[int, str]:
    return {
        3: (
            f'IF(B{row}="","",'
            f'IF(B{row}="NDIA-managed","You",'
            f'IF(B{row}="Plan-managed","Plan manager",'
            f'IF(B{row}="Self-managed","Participant",""))))'
        ),
        5: (
            f'IF(D{row}="","",'
            f'IF(D{row}<Settings!$B$2,MIN(D{row}+89,Settings!$B$3),D{row}+89))'
        ),
        6: f'IF(E{row}="","",E{row}-Settings!$B$1)',
        7: f'IF(OR(F{row}="",J{row}<>""),"",F{row}-TODAY())',
        9: (
            f'IF(D{row}="","",'
            f'IF(J{row}<>"","Lodged",'
            f'IF(G{row}="","",'
            f'IF(G{row}<0,"Past guide date",'
            f'IF(G{row}<=Settings!$B$1,"Due soon","On track")))))'
        ),
    }


def _tracker(ws, fills: dict, thin: Border) -> None:
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.left.text = "90-day claims tracker · guide only"
    ws.oddFooter.left.text = "Not a personalised legal deadline. Facts checked 2 November 2026."
    ws.oddFooter.right.text = "Page &P of &N"
    ws.print_title_rows = "1:4"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.tabColor = "C4A35A"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    ws["A1"] = "90-day claims tracker (guide only)"
    ws["A1"].font = Font(bold=True, size=16, color="1F4E79")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(HEADERS))
    ws["A2"] = DISCLAIMER
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws["A2"].fill = fills["banner"]
    ws.row_dimensions[2].height = 48

    header_font = Font(bold=True, color="FFFFFF", size=10)
    for col, name in enumerate(HEADERS, start=1):
        cell = ws.cell(HEADER_ROW, col, name)
        cell.font = header_font
        cell.fill = fills["navy"]
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[get_column_letter(col)].width = 16 if col != 13 else 28
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["K"].width = 24
    ws.row_dimensions[HEADER_ROW].height = 32
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A{HEADER_ROW}:M{LAST_DATA}"
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.sheet_view.zoomScale = 110

    for row in range(FIRST_DATA, LAST_DATA + 1):
        formulas = _formula(row)
        for col in range(1, len(HEADERS) + 1):
            cell = ws.cell(row, col, f"={formulas[col]}" if col in formulas else None)
            cell.border = thin
            cell.alignment = Alignment(vertical="center", wrap_text=(col == 13))
            cell.fill = fills["auto"] if col in FORMULA_COLS else fills["input"]
            if col in DATE_COLS:
                cell.number_format = "DD MMM YYYY"
            if col == 7:
                cell.number_format = "0"

    for offset, (code, plan, support) in enumerate(EXAMPLES):
        row = FIRST_DATA + offset
        ws.cell(row, 1, code)
        ws.cell(row, 2, plan)
        ws.cell(row, 4, support)
        for col in (1, 2, 4):
            ws.cell(row, col).fill = fills["example"]

    plans = DataValidation(
        type="list",
        formula1='"NDIA-managed,Plan-managed,Self-managed"',
        allow_blank=True,
    )
    plans.error = "Choose NDIA-managed, Plan-managed or Self-managed."
    plans.errorTitle = "Plan management"
    plans.add(f"B{FIRST_DATA}:B{LAST_DATA}")
    ws.add_data_validation(plans)

    statuses = DataValidation(
        type="list",
        formula1='"Not lodged,Invoiced,Lodged,Paid,Needs follow-up"',
        allow_blank=True,
    )
    statuses.add(f"H{FIRST_DATA}:H{LAST_DATA}")
    ws.add_data_validation(statuses)

    red = fills["red"]
    amber = fills["amber"]
    green = fills["green"]
    grey = fills["grey"]
    days = f"G{FIRST_DATA}:G{LAST_DATA}"
    alerts = f"I{FIRST_DATA}:I{LAST_DATA}"
    ws.conditional_formatting.add(
        days,
        CellIsRule(operator="lessThan", formula=["0"], fill=red, stopIfTrue=True),
    )
    ws.conditional_formatting.add(
        days,
        FormulaRule(formula=[f'AND(G{FIRST_DATA}<>"",G{FIRST_DATA}<=Settings!$B$1)'], fill=amber),
    )
    ws.conditional_formatting.add(
        days,
        FormulaRule(formula=[f'AND(G{FIRST_DATA}<>"",G{FIRST_DATA}>Settings!$B$1)'], fill=green),
    )
    for label, fill in (
        ("Past guide date", red),
        ("Due soon", amber),
        ("On track", green),
        ("Lodged", grey),
    ):
        ws.conditional_formatting.add(
            alerts,
            CellIsRule(operator="equal", formula=[f'"{label}"'], fill=fill),
        )


def _how_to(wb: Workbook, fills: dict) -> None:
    ws = wb.create_sheet("How to use", 1)
    ws.sheet_properties.tabColor = "5B7C99"
    lines = [
        "How to use this tracker",
        DISCLAIMER,
        "Facts checked 2 November 2026. Rules can change. Check the official websites before you rely on a date.",
        "One row per support. Fill the cream cells. Leave the blue cells alone; the sheet fills them in.",
        "Participant code is your own reference. Do not put a participant's name in the sheet if you can avoid it.",
        "Plan management: NDIA-managed, Plan-managed or Self-managed. Who usually claims is you, the plan manager, or the participant, in that order.",
        "Claim-by date (guide only) is the support date plus 89 days, so the day of the support counts as day 1. A shift on 2 December 2026 has a guide date of 1 March 2027.",
        "For supports delivered before 1 December 2026, the guide date stops at 30 November 2026. As at 2 November 2026, no transitional rules had been made for this change. To be safe, make sure claims for older supports are lodged by whoever claims before 1 December 2026. Not legal advice.",
        "If the guide date falls on a weekend or public holiday, the legal deadline may move to the next business day. This sheet does not move it. Don't rely on that movement.",
        "On the Settings tab, set a safety buffer so Your target date comes up before the guide date. Days left and Alert use the target date, and they are blank once you enter a Submission date.",
        "Alert is a reminder, not a legal status. Past guide date means the target date is already behind you.",
        "The three EXAMPLE rows are the edge cases checked on 3 November 2026: 2 September 2026, 3 September 2026 and 2 December 2026. Delete them before you add real supports.",
        "Colour is only on Days left and Alert.",
    ]
    for i, line in enumerate(lines, start=1):
        ws.cell(i, 1, line)
        ws.cell(i, 1).alignment = Alignment(wrap_text=True, vertical="center")
        if i == 1:
            ws.cell(i, 1).font = Font(bold=True, size=16, color="1F4E79")
        elif i == 2:
            ws.cell(i, 1).fill = fills["banner"]
    ws.column_dimensions["A"].width = 120
    ws.row_dimensions[2].height = 48
    for row in range(3, len(lines) + 1):
        ws.row_dimensions[row].height = 32


if __name__ == "__main__":
    print(build())
