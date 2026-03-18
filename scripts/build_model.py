"""
build_model.py
==============
Main script for building the institutional stock screener Excel model.

Usage:
    python scripts/build_model.py

Output:
    output/stock_screener.xlsx

After running:
    python scripts/recalc.py output/stock_screener.xlsx
"""

import json
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side,
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

# ── Color palette (institutional style) ──────────────────────────────────────
NAVY = "1F3864"
WHITE = "FFFFFF"
LIGHT_GRAY = "F2F2F2"
MID_GRAY = "D9D9D9"
DARK_GRAY = "595959"
BLUE_INPUT = "0000FF"
BLACK_FORMULA = "000000"
GREEN_LINK = "008000"
GREEN_BG = "E2EFDA"
YELLOW_BG = "FFEB9C"
ORANGE_BG = "FFCC99"
RED_BG = "FFC7CE"
GREEN_TEXT = "375623"
YELLOW_TEXT = "9C6500"
RED_TEXT = "9C0006"
ACCENT_BLUE = "2E75B6"

# ── Sector colors (pastel, for ticker column background) ─────────────────────
SECTOR_COLORS = {
    "Healthcare": "BDD7EE",
    "Defense / Aerospace": "E2EFDA",
    "Consumer Staples": "FFF2CC",
    "Industrials": "FCE4D6",
    "Utilities": "EDEDED",
    "International": "E2D9F3",
}

SECTOR_HEADER_COLORS = {
    "Healthcare": "2E75B6",
    "Defense / Aerospace": "548235",
    "Consumer Staples": "BF8F00",
    "Industrials": "C55A11",
    "Utilities": "595959",
    "International": "7030A0",
}

SECTOR_ORDER = [
    "Healthcare", "Defense / Aerospace", "Consumer Staples",
    "Industrials", "Utilities", "International",
]


def load_research_data():
    """Load the research data compiled by Claude Code from web research."""
    data_path = "output/research_data.json"
    if not os.path.exists(data_path):
        raise FileNotFoundError(
            f"Research data not found at {data_path}. "
            "Claude Code must complete web research and save to output/research_data.json first."
        )
    with open(data_path, "r") as f:
        return json.load(f)


def make_header_style(bold=True, size=11, color=WHITE, bg=NAVY, wrap=False, halign="center"):
    font = Font(name="Arial", bold=bold, size=size, color=color)
    fill = PatternFill("solid", fgColor=bg)
    align = Alignment(horizontal=halign, vertical="center", wrap_text=wrap)
    return font, fill, align


def make_border(style="thin"):
    s = Side(style=style, color=MID_GRAY)
    return Border(left=s, right=s, top=s, bottom=s)


def apply_header_row(ws, row, headers, col_start=1, bg=NAVY, text_color=WHITE, height=30):
    font, fill, align = make_header_style(bg=bg, color=text_color)
    border = make_border()
    ws.row_dimensions[row].height = height
    for i, h in enumerate(headers):
        cell = ws.cell(row=row, column=col_start + i, value=h)
        cell.font = font
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border


def style_data_cell(cell, value, is_formula=False, is_input=False, is_link=False,
                    num_format=None, row_shade=False, bold=False):
    cell.value = value
    if is_input:
        color = BLUE_INPUT
    elif is_link:
        color = GREEN_LINK
    else:
        color = BLACK_FORMULA
    cell.font = Font(name="Arial", size=10, color=color, bold=bold)
    if row_shade:
        cell.fill = PatternFill("solid", fgColor=LIGHT_GRAY)
    cell.border = make_border()
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
    if num_format:
        cell.number_format = num_format


def _apply_sector_color_to_ticker(cell, sector):
    """Apply pastel sector color as background fill to a ticker cell."""
    color = SECTOR_COLORS.get(sector, "FFFFFF")
    cell.fill = PatternFill("solid", fgColor=color)


def _apply_disqualified_style(ws, row, num_cols):
    """Apply red background + strikethrough text for DISQUALIFIED rows."""
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = Font(name="Arial", size=10, color=RED_TEXT, bold=False, strike=True)
        cell.fill = PatternFill("solid", fgColor=RED_BG)


def _write_stock_ranking_rows(ws, stocks_list, start_row, num_cols=17):
    """Write a ranked list of stocks into rows. Returns next available row."""
    for i, s in enumerate(stocks_list):
        row = start_row + i
        shade = (i % 2 == 1)
        score = s.get("composite_score", 0)
        is_disqualified = s.get("disqualifier_reason") is not None

        if score >= 7.5:
            rec_fill = PatternFill("solid", fgColor=GREEN_BG)
            rec_font = Font(name="Arial", size=10, color=GREEN_TEXT, bold=True)
        elif score >= 5.0:
            rec_fill = PatternFill("solid", fgColor=YELLOW_BG)
            rec_font = Font(name="Arial", size=10, color=YELLOW_TEXT, bold=True)
        else:
            rec_fill = PatternFill("solid", fgColor=RED_BG)
            rec_font = Font(name="Arial", size=10, color=RED_TEXT, bold=True)

        criminal_flag = s.get("criminal_flag", False)
        criminal_text = "Y" if criminal_flag else "N"

        row_data = [
            (i + 1,                                          "#,##0",       False),
            (s.get("ticker", ""),                             "@",           True),
            (s.get("company", ""),                            "@",           False),
            (s.get("sector", ""),                             "@",           False),
            (s.get("current_price", ""),                      "$#,##0.00",   True),
            (s.get("pct_off_high", ""),                       "0.0%",        True),
            (score,                                           "0.0",         False),
            (s.get("recommendation", ""),                     "@",           False),
            (s.get("target_price", ""),                       "$#,##0.00",   True),
            (s.get("implied_upside", ""),                     "0.0%",        False),
            (s.get("expected_value_pct", ""),                 "0.0%",        False),
            (s.get("key_risk", ""),                           "@",           False),
            (s.get("key_catalyst", ""),                       "@",           False),
            (s.get("analyst_consensus", ""),                  "@",           False),
            (s.get("catalyst_name", "") + " — " + s.get("catalyst_date", ""), "@", False),
            (s.get("selloff_type", ""),                       "@",           False),
            (criminal_text,                                   "@",           False),
        ]

        for col_idx, (val, fmt, is_inp) in enumerate(row_data, 1):
            cell = ws.cell(row=row, column=col_idx, value=val)
            if col_idx == 8:  # Recommendation column
                cell.font = rec_font
                cell.fill = rec_fill
                cell.border = make_border()
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = "@"
            else:
                style_data_cell(cell, val, is_input=is_inp, num_format=fmt, row_shade=shade)

        # Apply sector color to ticker cell (col 2)
        _apply_sector_color_to_ticker(ws.cell(row=row, column=2), s.get("sector", ""))

        # Apply DISQUALIFIED styling if applicable
        if is_disqualified:
            _apply_disqualified_style(ws, row, num_cols)

        ws.row_dimensions[row].height = 20

    return start_row + len(stocks_list)


def build_summary_dashboard(wb, data):
    """Sheet 1: Summary Dashboard — Master ranking + per-sector top 5 + screened out."""
    ws = wb.create_sheet("SUMMARY DASHBOARD")
    ws.sheet_view.showGridLines = False

    num_cols = 17
    last_col = get_column_letter(num_cols)

    # Title block
    ws.merge_cells(f"A1:{last_col}1")
    title_cell = ws["A1"]
    title_cell.value = "EQUITY SCREENER — BEATEN-DOWN VALUE WITH UPSIDE"
    title_cell.font = Font(name="Arial", bold=True, size=16, color=WHITE)
    title_cell.fill = PatternFill("solid", fgColor=NAVY)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 45

    ws.merge_cells(f"A2:{last_col}2")
    sub_cell = ws["A2"]
    sectors_present = sorted(set(s.get("sector", "") for s in data.get("stocks", [])))
    sub_cell.value = (
        f"Strategy: Contrarian Value | Macro: Bearish Credit & US Economy | "
        f"Sectors: {' · '.join(sectors_present)} | "
        f"Top 5 per Sector | As of: {datetime.today().strftime('%B %d, %Y')}"
    )
    sub_cell.font = Font(name="Arial", size=10, italic=True, color=DARK_GRAY)
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    ws.merge_cells(f"A3:{last_col}3")

    # ── MASTER CROSS-SECTOR RANKING ──────────────────────────────────────────
    master_title_row = 4
    ws.merge_cells(f"A{master_title_row}:{last_col}{master_title_row}")
    mt = ws.cell(row=master_title_row, column=1,
                 value="MASTER RANKING — ALL SECTORS (sorted by Composite Score)")
    mt.font = Font(name="Arial", bold=True, size=12, color=WHITE)
    mt.fill = PatternFill("solid", fgColor=NAVY)
    mt.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[master_title_row].height = 28

    headers = [
        "Rank", "Ticker", "Company", "Sector", "Price", "% Off High",
        "Composite\nScore", "Recommendation",
        "Target\nPrice", "Upside\n(%)", "Expected\nValue (%)",
        "Key Risk", "Key Catalyst", "Analyst\nConsensus",
        "Catalyst Name\n& Date", "Selloff\nType", "Criminal/\nReg Flag"
    ]
    header_row = master_title_row + 1
    apply_header_row(ws, header_row, headers, height=40)

    stocks = data.get("stocks", [])
    sorted_stocks = sorted(stocks, key=lambda x: x.get("composite_score", 0), reverse=True)

    data_start = header_row + 1
    next_row = _write_stock_ranking_rows(ws, sorted_stocks, data_start, num_cols)

    # Portfolio construction note
    note_row = next_row + 1
    ws.merge_cells(f"A{note_row}:{last_col}{note_row}")
    note = ws.cell(row=note_row, column=1)
    note.value = (
        "PORTFOLIO CONSTRUCTION GUIDANCE  |  Strong Buy (>=8.0): 3-5% position  |  "
        "Buy (6.5-7.9): 2-3% position  |  Speculative (5.0-6.4): 1% or watchlist only  |  "
        "Max single position: 5%  |  Total strategy exposure: 15-25% of portfolio  |  "
        "All positions: define stop-loss at 15-20% from entry before initiating"
    )
    note.font = Font(name="Arial", size=9, italic=True, color=DARK_GRAY)
    note.fill = PatternFill("solid", fgColor="EBF3FB")
    note.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    note.border = make_border()
    ws.row_dimensions[note_row].height = 30

    # ── PER-SECTOR TOP 5 TABLES ──────────────────────────────────────────────
    current_row = note_row + 3
    for sector in SECTOR_ORDER:
        sector_stocks = [s for s in stocks if s.get("sector", "") == sector]
        sector_sorted = sorted(sector_stocks, key=lambda x: x.get("composite_score", 0), reverse=True)[:5]
        if not sector_sorted:
            continue

        sector_color = SECTOR_HEADER_COLORS.get(sector, ACCENT_BLUE)

        # Sector header
        ws.merge_cells(f"A{current_row}:{last_col}{current_row}")
        sec_cell = ws.cell(row=current_row, column=1,
                           value=f"TOP 5 — {sector.upper()}")
        sec_cell.font = Font(name="Arial", bold=True, size=12, color=WHITE)
        sec_cell.fill = PatternFill("solid", fgColor=sector_color)
        sec_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[current_row].height = 28
        current_row += 1

        apply_header_row(ws, current_row, headers, height=35, bg=sector_color)
        current_row += 1

        current_row = _write_stock_ranking_rows(ws, sector_sorted, current_row, num_cols)
        current_row += 2  # gap between sectors

    # ── SCREENED-OUT TABLE ────────────────────────────────────────────────────
    screened_out = data.get("screened_out", [])
    if screened_out:
        so_row = current_row + 1
        so_cols = 7
        so_last = get_column_letter(so_cols)
        ws.merge_cells(f"A{so_row}:{so_last}{so_row}")
        header = ws.cell(row=so_row, column=1, value="SCREENED OUT — DID NOT PASS HARD FILTERS")
        header.font = Font(name="Arial", bold=True, size=11, color=WHITE)
        header.fill = PatternFill("solid", fgColor=DARK_GRAY)
        header.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[so_row].height = 22

        so_headers = [
            "Ticker", "Company", "Sector", "Failure Type",
            "Reason", "Revisit Trigger", "Disqualifier"
        ]
        apply_header_row(ws, so_row + 1, so_headers, bg=MID_GRAY, text_color="000000")

        for j, so in enumerate(screened_out):
            r = so_row + 2 + j
            vals = [
                so.get("ticker", ""), so.get("company", ""),
                so.get("sector", ""), so.get("filter_failed", ""),
                so.get("reason", ""), so.get("revisit_if", ""),
                so.get("disqualifier", "") or ""
            ]
            for k, val in enumerate(vals, 1):
                cell = ws.cell(row=r, column=k, value=val)
                style_data_cell(cell, val, row_shade=(j % 2 == 1))

            # Apply sector color to ticker
            _apply_sector_color_to_ticker(ws.cell(row=r, column=1), so.get("sector", ""))

            # Strikethrough for disqualified entries
            if so.get("disqualifier"):
                for k in range(1, so_cols + 1):
                    cell = ws.cell(row=r, column=k)
                    cell.font = Font(name="Arial", size=10, color=RED_TEXT, strike=True)
                    cell.fill = PatternFill("solid", fgColor=RED_BG)

    # Column widths
    col_widths = [6, 8, 28, 22, 10, 10, 10, 18, 10, 8, 10, 35, 35, 14, 30, 22, 10]
    for ci, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = f"A{data_start}"
    return ws


# ── 7-Criteria scoring system ────────────────────────────────────────────────
CRITERIA = [
    ("Valuation", 0.25),
    ("Balance Sheet", 0.20),
    ("Catalyst Specificity", 0.20),
    ("Selloff Quality", 0.15),
    ("Competitive Moat Integrity", 0.10),
    ("Macro Alignment", 0.10),
    ("Analyst Conviction", 0.05),
]

CRITERIA_KEYS = [
    "valuation", "balance_sheet", "catalyst_specificity",
    "selloff_quality", "competitive_moat_integrity",
    "macro_alignment", "analyst_conviction",
]


def build_stock_scores(wb, data):
    """Sheet 2: Full 7-criteria scoring matrix — grouped by sector."""
    ws = wb.create_sheet("STOCK SCORES")
    ws.sheet_view.showGridLines = False

    num_score_cols = 2 + len(CRITERIA) + 2  # Ticker, Sector, 7 criteria, Composite, Rec
    last_col = get_column_letter(num_score_cols)

    ws.merge_cells(f"A1:{last_col}1")
    title = ws["A1"]
    title.value = "COMPOSITE SCORING MATRIX — 7 CRITERIA (v2 Post-Mortem Updated)"
    title.font = Font(name="Arial", bold=True, size=14, color=WHITE)
    title.fill = PatternFill("solid", fgColor=NAVY)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    weight_headers = (
        ["Ticker", "Sector"]
        + [f"{c[0]}\n(w={c[1]:.0%})" for c in CRITERIA]
        + ["Composite\nScore", "Rec."]
    )
    apply_header_row(ws, 2, weight_headers, height=40)

    # Note row
    ws.merge_cells(f"A3:{last_col}3")
    note = ws["A3"]
    note.value = (
        "Scores: 0-10 per criterion  |  Blue = Hardcoded inputs  |  "
        "Black = Formulas  |  Weighted composite = sum of (score x weight)  |  "
        "DISQUALIFIED rows have red strikethrough"
    )
    note.font = Font(name="Arial", size=9, italic=True, color=DARK_GRAY)
    note.fill = PatternFill("solid", fgColor="EBF3FB")
    note.alignment = Alignment(horizontal="left", vertical="center")

    stocks = data.get("stocks", [])
    current_row = 4
    for sector in SECTOR_ORDER:
        sector_stocks = [s for s in stocks if s.get("sector", "") == sector]
        sector_sorted = sorted(sector_stocks, key=lambda x: x.get("composite_score", 0), reverse=True)
        if not sector_sorted:
            continue

        # Sector divider row
        sector_color = SECTOR_HEADER_COLORS.get(sector, ACCENT_BLUE)
        ws.merge_cells(f"A{current_row}:{last_col}{current_row}")
        sec_cell = ws.cell(row=current_row, column=1, value=sector.upper())
        sec_cell.font = Font(name="Arial", bold=True, size=11, color=WHITE)
        sec_cell.fill = PatternFill("solid", fgColor=sector_color)
        sec_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[current_row].height = 24
        current_row += 1

        for i, s in enumerate(sector_sorted):
            row = current_row
            shade = (i % 2 == 1)
            scores = s.get("scores", {})
            is_disqualified = s.get("disqualifier_reason") is not None

            # Ticker + Sector
            ticker_cell = ws.cell(row=row, column=1, value=s.get("ticker", ""))
            ticker_cell.font = Font(name="Arial", bold=True, size=11)
            ticker_cell.border = make_border()
            _apply_sector_color_to_ticker(ticker_cell, s.get("sector", ""))

            sector_cell = ws.cell(row=row, column=2, value=s.get("sector", ""))
            sector_cell.font = Font(name="Arial", size=9, color=DARK_GRAY)
            sector_cell.border = make_border()

            # 7 criteria scores (columns C through I)
            for j, key in enumerate(CRITERIA_KEYS):
                score_val = scores.get(key, {}).get("score", "")
                cell = ws.cell(row=row, column=3 + j, value=score_val)
                style_data_cell(cell, score_val, is_input=True, num_format="0.0", row_shade=shade)

            # Composite formula (columns C-I are 7 scores, J is composite)
            score_cols = [get_column_letter(3 + j) for j in range(len(CRITERIA))]
            weights = [c[1] for c in CRITERIA]
            formula_parts = [f"{col}{row}*{w}" for col, w in zip(score_cols, weights)]
            composite_col = 3 + len(CRITERIA)  # Column J (10)
            composite_cell = ws.cell(row=row, column=composite_col)
            composite_cell.value = "=" + "+".join(formula_parts)
            composite_cell.font = Font(name="Arial", size=11, bold=True, color=BLACK_FORMULA)
            composite_cell.number_format = "0.0"
            composite_cell.alignment = Alignment(horizontal="center", vertical="center")
            composite_cell.border = make_border()

            # Recommendation formula
            rec_col = composite_col + 1  # Column K (11)
            comp_ref = f"{get_column_letter(composite_col)}{row}"
            rec_cell = ws.cell(row=row, column=rec_col)
            rec_cell.value = (
                f'=IF({comp_ref}>=8,"STRONG BUY",IF({comp_ref}>=6.5,"BUY",'
                f'IF({comp_ref}>=5,"SPECULATIVE BUY","PASS")))'
            )
            rec_cell.font = Font(name="Arial", size=10, color=BLACK_FORMULA, bold=True)
            rec_cell.alignment = Alignment(horizontal="center", vertical="center")
            rec_cell.border = make_border()

            # DISQUALIFIED styling
            if is_disqualified:
                _apply_disqualified_style(ws, row, num_score_cols)

            current_row += 1

            # Justification row
            ws.cell(row=current_row, column=1, value="Rationale:").font = Font(
                name="Arial", italic=True, size=9, color=DARK_GRAY)
            justifications = scores.get("justifications", {})
            for j, key in enumerate(CRITERIA_KEYS):
                just_text = justifications.get(key, "")
                cell = ws.cell(row=current_row, column=3 + j, value=just_text)
                cell.font = Font(name="Arial", italic=True, size=8, color=DARK_GRAY)
                cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            ws.row_dimensions[current_row].height = 50
            current_row += 1

        current_row += 1  # gap between sectors

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 20
    for j in range(len(CRITERIA)):
        ws.column_dimensions[get_column_letter(3 + j)].width = 16
    ws.column_dimensions[get_column_letter(composite_col)].width = 12
    ws.column_dimensions[get_column_letter(rec_col)].width = 16

    ws.freeze_panes = "A4"
    return ws


def build_valuation_comps(wb, data):
    """Sheet 3: Valuation Comps"""
    ws = wb.create_sheet("VALUATION COMPS")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:Q1")
    t = ws["A1"]
    t.value = "VALUATION ANALYSIS & COMPARABLE COMPANIES"
    t.font = Font(name="Arial", bold=True, size=14, color=WHITE)
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    headers = [
        "Ticker", "Sector", "Price",
        "Fwd P/E", "Sector Avg P/E", "5-Yr Hist P/E", "P/E vs. Hist (%)",
        "EV/EBITDA", "Peer Avg EV/EBITDA", "EV/EBITDA vs. Peers (%)",
        "P/FCF", "EV/Sales",
        "Bear Target", "Base Target", "Bull Target",
        "Analyst Avg Target", "Implied Upside"
    ]
    apply_header_row(ws, 2, headers, height=40)

    stocks = data.get("stocks", [])
    for i, s in enumerate(stocks):
        row = 3 + i
        shade = (i % 2 == 1)
        v = s.get("valuation", {})

        row_vals = [
            (s.get("ticker", ""),                  "@",           True),
            (s.get("sector", ""),                   "@",           False),
            (s.get("current_price", ""),            "$#,##0.00",   True),
            (v.get("fwd_pe", ""),                   "0.0x",        True),
            (v.get("sector_avg_pe", ""),            "0.0x",        True),
            (v.get("hist_5yr_pe", ""),              "0.0x",        True),
            (f"=(D{row}-F{row})/F{row}",           "0.0%",        False),
            (v.get("ev_ebitda", ""),                "0.0x",        True),
            (v.get("peer_avg_ev_ebitda", ""),       "0.0x",        True),
            (f"=(H{row}-I{row})/I{row}",           "0.0%",        False),
            (v.get("p_fcf", ""),                    "0.0x",        True),
            (v.get("ev_sales", ""),                 "0.0x",        True),
            (v.get("bear_target", ""),              "$#,##0.00",   True),
            (v.get("base_target", ""),              "$#,##0.00",   True),
            (v.get("bull_target", ""),              "$#,##0.00",   True),
            (v.get("analyst_avg_target", ""),       "$#,##0.00",   True),
            (f"=(P{row}-C{row})/C{row}",           "0.0%",        False),
        ]

        for col_idx, (val, fmt, is_inp) in enumerate(row_vals, 1):
            cell = ws.cell(row=row, column=col_idx, value=val)
            style_data_cell(cell, val, is_input=is_inp, num_format=fmt, row_shade=shade)

        # Sector color on ticker
        _apply_sector_color_to_ticker(ws.cell(row=row, column=1), s.get("sector", ""))
        ws.row_dimensions[row].height = 18

    for ci, w in enumerate([10, 20, 10, 8, 12, 12, 12, 10, 14, 14, 8, 8, 10, 10, 10, 14, 12], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "A3"
    return ws


def build_fundamental_data(wb, data):
    """Sheet 4: Fundamental Data"""
    ws = wb.create_sheet("FUNDAMENTAL DATA")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:T1")
    t = ws["A1"]
    t.value = "FUNDAMENTAL DATA — INCOME STATEMENT, BALANCE SHEET & CASH FLOW SUMMARY"
    t.font = Font(name="Arial", bold=True, size=14, color=WHITE)
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    headers = [
        "Ticker", "Company", "Sector", "Mkt Cap ($B)", "EV ($B)",
        "Rev LTM ($B)", "Rev NTM ($B)", "Rev Growth (%)",
        "EBITDA Mgn (%)", "EPS LTM", "EPS NTM", "EPS 3Y CAGR (%)",
        "Net Cash/(Debt) ($B)", "ND/EBITDA", "Interest Cov.", "Credit Rating",
        "FCF LTM ($B)", "FCF Yield (%)", "Capex/Rev (%)",
        "Div Yield (%)"
    ]
    apply_header_row(ws, 2, headers, height=40)

    stocks = data.get("stocks", [])
    for i, s in enumerate(stocks):
        row = 3 + i
        shade = (i % 2 == 1)
        f = s.get("fundamentals", {})
        bs = f.get("balance_sheet", {})
        cf = f.get("cash_flow", {})

        row_vals = [
            (s.get("ticker", ""),                  "@",                              True),
            (s.get("company", ""),                  "@",                              False),
            (s.get("sector", ""),                   "@",                              False),
            (f.get("market_cap_b", ""),             "$#,##0.0",                       True),
            (f.get("ev_b", ""),                     "$#,##0.0",                       True),
            (f.get("rev_ltm_b", ""),                "$#,##0.0",                       True),
            (f.get("rev_ntm_b", ""),                "$#,##0.0",                       True),
            (f"=(G{row}-F{row})/F{row}",            "0.0%",                           False),
            (f.get("ebitda_margin", ""),             "0.0%",                           True),
            (f.get("eps_ltm", ""),                   "$#,##0.00",                      True),
            (f.get("eps_ntm", ""),                   "$#,##0.00",                      True),
            (f.get("eps_3yr_cagr", ""),              "0.0%",                           True),
            (bs.get("net_cash_debt_b", ""),          "$#,##0.0;($#,##0.0);-",         True),
            (bs.get("nd_ebitda", ""),                "0.0x",                           True),
            (bs.get("interest_coverage", ""),         "0.0x",                           True),
            (bs.get("credit_rating", ""),             "@",                              True),
            (cf.get("fcf_ltm_b", ""),                "$#,##0.0",                       True),
            (f"=Q{row}/D{row}",                      "0.0%",                           False),
            (cf.get("capex_rev_pct", ""),             "0.0%",                           True),
            (f.get("div_yield", ""),                  "0.0%",                           True),
        ]

        for col_idx, (val, fmt, is_inp) in enumerate(row_vals, 1):
            cell = ws.cell(row=row, column=col_idx, value=val)
            style_data_cell(cell, val, is_input=is_inp, num_format=fmt, row_shade=shade)

        _apply_sector_color_to_ticker(ws.cell(row=row, column=1), s.get("sector", ""))
        ws.row_dimensions[row].height = 18

    for ci, w in enumerate([10, 28, 20, 10, 8, 10, 10, 10, 10, 8, 8, 10, 14, 10, 10, 10, 10, 10, 10, 10], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "A3"
    return ws


def build_catalyst_tracker(wb, data):
    """Sheet 5: Catalyst Tracker"""
    ws = wb.create_sheet("CATALYST TRACKER")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:H1")
    t = ws["A1"]
    t.value = "CATALYST TRACKER — UPCOMING EVENTS & PRICE IMPACT ESTIMATES"
    t.font = Font(name="Arial", bold=True, size=14, color=WHITE)
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    headers = [
        "Ticker", "Catalyst Description", "Timeline", "Probability",
        "Impact if Hit (%)", "Impact if Miss (%)", "Expected Impact (%)", "Monitor Signal"
    ]
    apply_header_row(ws, 2, headers, height=30)

    row = 3
    stocks = data.get("stocks", [])
    for s in stocks:
        catalysts = s.get("catalysts", [])
        ticker = s.get("ticker", "")

        for j, cat in enumerate(catalysts):
            shade = (row % 2 == 0)
            prob = cat.get("probability_decimal", 0)
            hit = cat.get("impact_if_hit_pct", 0)
            miss = cat.get("impact_if_miss_pct", 0)

            expected_val = f"=D{row}*E{row}+(1-D{row})*F{row}" if prob else ""

            row_vals = [
                (ticker if j == 0 else "",             "@",    True),
                (cat.get("description", ""),            "@",    False),
                (cat.get("timeline", ""),               "@",    True),
                (prob,                                  "0%",   True),
                (hit / 100 if isinstance(hit, (int, float)) else hit, "0.0%", True),
                (miss / 100 if isinstance(miss, (int, float)) else miss, "0.0%", True),
                (expected_val,                          "0.0%", False),
                (cat.get("monitor_signal", ""),         "@",    False),
            ]

            for col_idx, (val, fmt, is_inp) in enumerate(row_vals, 1):
                cell = ws.cell(row=row, column=col_idx, value=val)
                style_data_cell(cell, val, is_input=is_inp, num_format=fmt, row_shade=shade)
            ws.row_dimensions[row].height = 20
            row += 1

        if catalysts:
            row += 1

    for ci, w in enumerate([10, 45, 14, 11, 12, 12, 14, 35], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "A3"
    return ws


def build_risk_matrix(wb, data):
    """Sheet 6: Risk Matrix"""
    ws = wb.create_sheet("RISK MATRIX")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:N1")
    t = ws["A1"]
    t.value = "RISK MATRIX — BULL / BASE / BEAR SCENARIOS & EXPECTED VALUE"
    t.font = Font(name="Arial", bold=True, size=14, color=WHITE)
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    headers = [
        "Ticker",
        "Bull Case Desc.", "Bull Target", "Bull Return (%)", "P(Bull)",
        "Base Case Desc.", "Base Target", "Base Return (%)", "P(Base)",
        "Bear Case Desc.", "Bear Target", "Bear Return (%)", "P(Bear)",
        "Expected\nValue (%)",
    ]
    apply_header_row(ws, 2, headers, height=40)

    stocks = data.get("stocks", [])
    for i, s in enumerate(stocks):
        row = 3 + i
        shade = (i % 2 == 1)
        sc = s.get("scenarios", {})
        bull = sc.get("bull", {})
        base = sc.get("base", {})
        bear = sc.get("bear", {})

        p_bull = bull.get("probability", 0)
        p_base = base.get("probability", 0)
        bull_ret_val = bull.get("return_pct", "")
        base_ret_val = base.get("return_pct", "")
        bear_ret_val = bear.get("return_pct", "")

        if all(isinstance(x, (int, float)) for x in [p_bull, p_base]):
            p_bear_val = round(1 - p_bull - p_base, 2)
        else:
            p_bear_val = ""

        row_vals = [
            (s.get("ticker", ""),                "@",           True),
            (bull.get("description", ""),         "@",           True),
            (bull.get("price_target", ""),        "$#,##0.00",   True),
            (bull_ret_val,                        "0.0%",        True),
            (p_bull,                              "0%",          True),
            (base.get("description", ""),         "@",           True),
            (base.get("price_target", ""),        "$#,##0.00",   True),
            (base_ret_val,                        "0.0%",        True),
            (p_base,                              "0%",          True),
            (bear.get("description", ""),         "@",           True),
            (bear.get("price_target", ""),        "$#,##0.00",   True),
            (bear_ret_val,                        "0.0%",        True),
            (p_bear_val,                          "0%",          True),
        ]

        for col_idx, (val, fmt, is_inp) in enumerate(row_vals, 1):
            cell = ws.cell(row=row, column=col_idx, value=val)
            style_data_cell(cell, val, is_input=is_inp, num_format=fmt, row_shade=shade)

        # EV formula
        ev_formula = f"=E{row}*D{row}+I{row}*H{row}+M{row}*L{row}"
        ev_cell = ws.cell(row=row, column=14, value=ev_formula)
        style_data_cell(ev_cell, ev_formula, num_format="0.0%", row_shade=shade)

        _apply_sector_color_to_ticker(ws.cell(row=row, column=1), s.get("sector", ""))
        ws.row_dimensions[row].height = 22

    for ci, w in enumerate([10, 30, 10, 10, 8, 30, 10, 10, 8, 30, 10, 10, 8, 10], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "A3"
    return ws


def build_nvo_deep_dive(wb, data):
    """Sheet 7: NVO Deep Dive"""
    ws = wb.create_sheet("NVO DEEP DIVE")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:J1")
    t = ws["A1"]
    t.value = "NVO (NOVO NORDISK) — INSTITUTIONAL DEEP DIVE: IS THE BAD NEWS PRICED IN?"
    t.font = Font(name="Arial", bold=True, size=14, color=WHITE)
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 45

    nvo = next((s for s in data.get("stocks", []) if s.get("ticker") == "NVO"), {})
    nvo_deep = nvo.get("deep_dive", {})

    # Section: Selloff Timeline
    current_row = 3
    ws.merge_cells(f"A{current_row}:J{current_row}")
    section = ws.cell(row=current_row, column=1, value="SELLOFF TIMELINE — KEY EVENTS")
    section.font = Font(name="Arial", bold=True, size=12, color=WHITE)
    section.fill = PatternFill("solid", fgColor=ACCENT_BLUE)
    section.alignment = Alignment(vertical="center")
    ws.row_dimensions[current_row].height = 24
    current_row += 1

    timeline_headers = ["Date", "Event", "Stock Price", "% Change", "Interpretation"]
    apply_header_row(ws, current_row, timeline_headers, col_start=1, bg=MID_GRAY, text_color="000000")
    current_row += 1

    timeline = nvo_deep.get("selloff_timeline", [])
    for i, event in enumerate(timeline):
        for col_idx, (key, fmt) in enumerate([
            ("date", "@"), ("event", "@"), ("price", "$#,##0.00"), ("pct_change", "0.0%"), ("interpretation", "@")
        ], 1):
            cell = ws.cell(row=current_row, column=col_idx, value=event.get(key, ""))
            style_data_cell(cell, event.get(key, ""), is_input=True, num_format=fmt, row_shade=(i % 2 == 1))
        ws.row_dimensions[current_row].height = 20
        current_row += 1

    current_row += 1

    # Section: Pipeline Status
    ws.merge_cells(f"A{current_row}:J{current_row}")
    section2 = ws.cell(row=current_row, column=1, value="PIPELINE STATUS TABLE")
    section2.font = Font(name="Arial", bold=True, size=12, color=WHITE)
    section2.fill = PatternFill("solid", fgColor=ACCENT_BLUE)
    section2.alignment = Alignment(vertical="center")
    ws.row_dimensions[current_row].height = 24
    current_row += 1

    pipe_headers = ["Asset", "Indication", "Phase", "Status", "Timeline",
                    "Market Size ($B)", "Probability", "Bear Rev ($M)", "Base Rev ($M)", "Bull Rev ($M)"]
    apply_header_row(ws, current_row, pipe_headers, col_start=1, bg=MID_GRAY, text_color="000000")
    current_row += 1

    pipeline = nvo_deep.get("pipeline", [])
    for i, asset in enumerate(pipeline):
        for col_idx, (key, fmt) in enumerate([
            ("asset", "@"), ("indication", "@"), ("phase", "@"), ("status", "@"),
            ("timeline", "@"), ("market_size_b", "$#,##0.0"), ("probability", "0%"),
            ("bear_rev_m", "$#,##0"), ("base_rev_m", "$#,##0"), ("bull_rev_m", "$#,##0")
        ], 1):
            cell = ws.cell(row=current_row, column=col_idx, value=asset.get(key, ""))
            style_data_cell(cell, asset.get(key, ""), is_input=True, num_format=fmt, row_shade=(i % 2 == 1))
        ws.row_dimensions[current_row].height = 20
        current_row += 1

    current_row += 1

    # Section: IRA Pricing Impact
    ws.merge_cells(f"A{current_row}:J{current_row}")
    section3 = ws.cell(row=current_row, column=1, value="IRA PRICING IMPACT MODEL (2027+)")
    section3.font = Font(name="Arial", bold=True, size=12, color=WHITE)
    section3.fill = PatternFill("solid", fgColor=ACCENT_BLUE)
    section3.alignment = Alignment(vertical="center")
    ws.row_dimensions[current_row].height = 24
    current_row += 1

    ira = nvo_deep.get("ira_impact", {})
    ira_data = [
        ("Current Ozempic/Wegovy Net Price (US avg)", ira.get("current_net_price", ""), "$#,##0"),
        ("IRA Max Fair Price — Ozempic", ira.get("ira_ozempic_price", ""), "$#,##0"),
        ("IRA Max Fair Price — Wegovy", ira.get("ira_wegovy_price", ""), "$#,##0"),
        ("Estimated Revenue at Risk (2027)", ira.get("revenue_at_risk_b", ""), "$#,##0.0"),
        ("% of 2025 Revenue at Risk", ira.get("pct_revenue_at_risk", ""), "0.0%"),
        ("EPS Impact (Bear Case)", ira.get("eps_impact_bear", ""), "$#,##0.00"),
        ("EPS Impact (Base Case)", ira.get("eps_impact_base", ""), "$#,##0.00"),
    ]
    for key, val, fmt in ira_data:
        cell_a = ws.cell(row=current_row, column=1, value=key)
        cell_a.font = Font(name="Arial", size=10, bold=False, color=DARK_GRAY)
        cell_a.alignment = Alignment(horizontal="left", vertical="center")
        cell_b = ws.cell(row=current_row, column=2, value=val)
        style_data_cell(cell_b, val, is_input=True, num_format=fmt)
        current_row += 1

    current_row += 1

    # Section: Competitive Positioning
    ws.merge_cells(f"A{current_row}:J{current_row}")
    section_comp = ws.cell(row=current_row, column=1, value="COMPETITIVE POSITIONING — NVO vs. LLY")
    section_comp.font = Font(name="Arial", bold=True, size=12, color=WHITE)
    section_comp.fill = PatternFill("solid", fgColor=ACCENT_BLUE)
    section_comp.alignment = Alignment(vertical="center")
    ws.row_dimensions[current_row].height = 24
    current_row += 1

    comp = nvo_deep.get("competitive_positioning", {})
    comp_data = [
        ("NVO GLP-1 Market Share", comp.get("nvo_glp1_market_share_pct", ""), "0.0%"),
        ("LLY GLP-1 Market Share", comp.get("lly_glp1_market_share_pct", ""), "0.0%"),
        ("Market Share Trend", comp.get("trend", ""), "@"),
        ("Notes", comp.get("notes", ""), "@"),
    ]
    for key, val, fmt in comp_data:
        cell_a = ws.cell(row=current_row, column=1, value=key)
        cell_a.font = Font(name="Arial", size=10, bold=False, color=DARK_GRAY)
        cell_a.alignment = Alignment(horizontal="left", vertical="center")
        cell_b = ws.cell(row=current_row, column=2, value=val)
        style_data_cell(cell_b, val, is_input=True, num_format=fmt)
        current_row += 1

    current_row += 1

    # Section: M&A Capacity
    ws.merge_cells(f"A{current_row}:J{current_row}")
    section_ma = ws.cell(row=current_row, column=1, value="M&A CAPACITY ANALYSIS")
    section_ma.font = Font(name="Arial", bold=True, size=12, color=WHITE)
    section_ma.fill = PatternFill("solid", fgColor=ACCENT_BLUE)
    section_ma.alignment = Alignment(vertical="center")
    ws.row_dimensions[current_row].height = 24
    current_row += 1

    ma = nvo_deep.get("ma_capacity", {})
    ma_data = [
        ("Available Cash ($B)", ma.get("available_cash_b", ""), "$#,##0.0"),
        ("Leverage Headroom ($B)", ma.get("leverage_headroom_b", ""), "$#,##0.0"),
        ("Total Firepower ($B)", ma.get("total_firepower_b", ""), "$#,##0.0"),
        ("Likely Targets", ma.get("likely_targets", ""), "@"),
    ]
    for key, val, fmt in ma_data:
        cell_a = ws.cell(row=current_row, column=1, value=key)
        cell_a.font = Font(name="Arial", size=10, bold=False, color=DARK_GRAY)
        cell_a.alignment = Alignment(horizontal="left", vertical="center")
        cell_b = ws.cell(row=current_row, column=2, value=val)
        style_data_cell(cell_b, val, is_input=True, num_format=fmt)
        current_row += 1

    current_row += 1

    # Section: Verdict
    ws.merge_cells(f"A{current_row}:J{current_row}")
    section4 = ws.cell(row=current_row, column=1, value="VERDICT: IS THE BAD NEWS PRICED IN?")
    section4.font = Font(name="Arial", bold=True, size=12, color=WHITE)
    section4.fill = PatternFill("solid", fgColor=ACCENT_BLUE)
    section4.alignment = Alignment(vertical="center")
    ws.row_dimensions[current_row].height = 24
    current_row += 1

    verdict = nvo_deep.get("verdict", {})
    verdict_cell_label = ws.cell(row=current_row, column=1, value="Conclusion:")
    verdict_cell_label.font = Font(name="Arial", bold=True, size=11)
    verdict_cell = ws.cell(row=current_row, column=2, value=verdict.get("conclusion", ""))
    verdict_cell.font = Font(name="Arial", bold=True, size=13,
                             color=GREEN_TEXT if verdict.get("bullish") else RED_TEXT)
    current_row += 1

    ws.merge_cells(f"A{current_row}:J{current_row + 4}")
    rationale_cell = ws.cell(row=current_row, column=1, value=verdict.get("rationale", ""))
    rationale_cell.font = Font(name="Arial", size=10)
    rationale_cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
    ws.row_dimensions[current_row].height = 90

    for ci, w in enumerate([22, 22, 12, 18, 12, 12, 10, 12, 12, 12], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "A3"
    return ws


def build_assumptions(wb, data):
    """Sheet 8: Assumptions"""
    ws = wb.create_sheet("ASSUMPTIONS")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:D1")
    t = ws["A1"]
    t.value = "MODEL ASSUMPTIONS & DATA SOURCES"
    t.font = Font(name="Arial", bold=True, size=14, color=WHITE)
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    assumptions = data.get("assumptions", {})

    current_row = 3
    apply_header_row(ws, current_row, ["MACRO ASSUMPTIONS", "Value", "Notes", "Source"], bg=ACCENT_BLUE)
    current_row += 1

    macro_items = assumptions.get("macro", [])
    for i, item in enumerate(macro_items):
        shade = (i % 2 == 1)
        for col_idx, (key, fmt) in enumerate([("label", "@"), ("value", "@"), ("notes", "@"), ("source", "@")], 1):
            cell = ws.cell(row=current_row, column=col_idx, value=item.get(key, ""))
            style_data_cell(cell, item.get(key, ""), is_input=(col_idx == 2), row_shade=shade)
        current_row += 1

    current_row += 1

    apply_header_row(ws, current_row, ["VALUATION ASSUMPTIONS", "Value", "Notes", "Source"], bg=ACCENT_BLUE)
    current_row += 1

    val_items = assumptions.get("valuation", [])
    for i, item in enumerate(val_items):
        shade = (i % 2 == 1)
        for col_idx, (key, fmt) in enumerate([("label", "@"), ("value", "@"), ("notes", "@"), ("source", "@")], 1):
            cell = ws.cell(row=current_row, column=col_idx, value=item.get(key, ""))
            style_data_cell(cell, item.get(key, ""), is_input=(col_idx == 2), row_shade=shade)
        current_row += 1

    current_row += 1

    apply_header_row(ws, current_row, ["PEER GROUP DEFINITIONS", "Peers Included", "Rationale", "Exclusions"], bg=ACCENT_BLUE)
    current_row += 1

    peer_groups = assumptions.get("peer_groups", [])
    for i, pg in enumerate(peer_groups):
        shade = (i % 2 == 1)
        for col_idx, key in enumerate(["ticker", "peers", "rationale", "exclusions"], 1):
            cell = ws.cell(row=current_row, column=col_idx, value=pg.get(key, ""))
            style_data_cell(cell, pg.get(key, ""), row_shade=shade)
        current_row += 1

    current_row += 1

    apply_header_row(ws, current_row, ["DATA SOURCES", "Description", "As-of Date", "Notes"], bg=ACCENT_BLUE)
    current_row += 1

    data_sources = assumptions.get("data_sources", [])
    if isinstance(data_sources, list):
        for i, src in enumerate(data_sources):
            if isinstance(src, str):
                cell = ws.cell(row=current_row, column=1, value=src)
                style_data_cell(cell, src)
            elif isinstance(src, dict):
                for col_idx, key in enumerate(["name", "description", "as_of_date", "notes"], 1):
                    cell = ws.cell(row=current_row, column=col_idx, value=src.get(key, ""))
                    style_data_cell(cell, src.get(key, ""), row_shade=(i % 2 == 1))
            current_row += 1

    for ci, w in enumerate([35, 30, 40, 30], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "A3"
    return ws


def build_screener_logic(wb, data):
    """Sheet 9: Screener Logic — Documents the updated screening methodology."""
    ws = wb.create_sheet("SCREENER LOGIC")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:D1")
    t = ws["A1"]
    t.value = "SCREENER LOGIC — METHODOLOGY & DECISION FRAMEWORK"
    t.font = Font(name="Arial", bold=True, size=14, color=WHITE)
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    current_row = 3

    # Helper to write a section header
    def section_header(row, title, bg=ACCENT_BLUE):
        ws.merge_cells(f"A{row}:D{row}")
        cell = ws.cell(row=row, column=1, value=title)
        cell.font = Font(name="Arial", bold=True, size=12, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 26
        return row + 1

    def data_row(row, col1, col2="", col3="", col4="", bold_first=False, shade=False):
        vals = [col1, col2, col3, col4]
        for ci, val in enumerate(vals, 1):
            cell = ws.cell(row=row, column=ci, value=val)
            cell.font = Font(name="Arial", size=10,
                             bold=(ci == 1 and bold_first),
                             color=DARK_GRAY if ci > 1 else BLACK_FORMULA)
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            cell.border = make_border()
            if shade:
                cell.fill = PatternFill("solid", fgColor=LIGHT_GRAY)
        ws.row_dimensions[row].height = 22
        return row + 1

    # ── Section 1: 7 Scoring Criteria ────────────────────────────────────────
    current_row = section_header(current_row, "7-CRITERIA SCORING SYSTEM (v2 — Post-Mortem Updated)")
    apply_header_row(ws, current_row, ["Criterion", "Weight", "Key Question", "Threshold Notes"], bg=MID_GRAY, text_color="000000")
    current_row += 1

    criteria_data = [
        ("Valuation", "25%", "How cheap is the stock on absolute and relative basis?",
         "9-10: Multi-year trough, 40%+ implied upside; 5-6: Modestly below fair value"),
        ("Balance Sheet", "20%", "Can it survive and fund itself through the downturn?",
         "9-10: Net cash, IG credit, strong FCF; Hard filter: ND/EBITDA < 3.5x, int cov > 3x"),
        ("Catalyst Specificity", "20%", "Named, dated, binary catalysts?",
         "9-10: 2+ named catalysts with dates within 12mo, binary outcome; 3-4: Vague/long-dated"),
        ("Selloff Quality", "15%", "Single-event overreaction or chronic decay?",
         "9-10: Classic single-event overreaction, core franchise intact; 1-2: Chronic pattern"),
        ("Competitive Moat Integrity", "10%", "Is the competitive position intact?",
         "9-10: Monopoly/near-monopoly intact; Penalty: -1.5pts for structural share loss"),
        ("Macro Alignment", "10%", "Is macro a tailwind or headwind?",
         "9-10: Strong tailwind; Cap at 3.0 if bipartisan political targeting/hostile regulation"),
        ("Analyst Conviction", "5%", "What does smart money think?",
         "9-10: 75%+ buys, recent upgrades; Must have more Buys than Holds+Sells combined"),
    ]
    for i, (crit, wt, question, notes) in enumerate(criteria_data):
        current_row = data_row(current_row, crit, wt, question, notes, bold_first=True, shade=(i % 2 == 1))

    current_row += 1

    # ── Section 2: Score Interpretation ──────────────────────────────────────
    current_row = section_header(current_row, "SCORE INTERPRETATION")
    apply_header_row(ws, current_row, ["Score Range", "Label", "Action", "Position Size"], bg=MID_GRAY, text_color="000000")
    current_row += 1

    interp_data = [
        ("8.0 - 10.0", "STRONG BUY", "High conviction entry", "3-5% position"),
        ("6.5 - 7.9", "BUY", "Solid entry", "2-3% position"),
        ("5.0 - 6.4", "SPECULATIVE BUY", "Small position or watchlist", "1% or watchlist"),
        ("3.0 - 4.9", "PASS", "Do not initiate", "N/A"),
        ("0 - 2.9", "AVOID", "Does not fit strategy", "N/A"),
        ("DISQUALIFIED", "DISQUALIFIED", "Hard disqualifier triggered — do not invest", "N/A"),
    ]
    for i, (rng, label, action, size) in enumerate(interp_data):
        current_row = data_row(current_row, rng, label, action, size, shade=(i % 2 == 1))

    current_row += 1

    # ── Section 3: Hard Filters ──────────────────────────────────────────────
    current_row = section_header(current_row, "HARD FILTERS (All Must Pass)")
    apply_header_row(ws, current_row, ["Filter", "Threshold", "Rationale", "Override Condition"], bg=MID_GRAY, text_color="000000")
    current_row += 1

    filters_data = [
        ("ATH Selloff", "Must be >30% below all-time high",
         "Target beaten-down stocks with asymmetric upside", "None — absolute requirement"),
        ("Recovery Filter", "(price - 52w_low) / (52w_high - 52w_low) < 50%",
         "Catch early-stage bottoming, not stocks that have already recovered", "None"),
        ("Balance Sheet", "ND/EBITDA < 3.5x AND interest coverage > 3x",
         "Must survive downturn without diluting equity holders", "None"),
        ("Criminal DOJ Investigation", "Active DOJ/FBI criminal investigation → DISQUALIFIED",
         "Unbounded downside risk from fines, departures, reputational damage",
         "Investigation formally closed with no charges AND stock not recovered"),
        ("Chronic Underperformance", "3+ years underperforming peers, no confirmed inflection → DISQUALIFIED",
         "'This time is different' is almost never true without concrete evidence",
         "Confirmed inflection: mgmt action taken + quantitative improvement + external validation"),
    ]
    for i, (filt, threshold, rationale, override) in enumerate(filters_data):
        current_row = data_row(current_row, filt, threshold, rationale, override, bold_first=True, shade=(i % 2 == 1))

    current_row += 1

    # ── Section 4: Catalyst Specificity Gate ─────────────────────────────────
    current_row = section_header(current_row, "CATALYST SPECIFICITY GATE")
    apply_header_row(ws, current_row, ["Requirement", "Description", "Pass Example", "Fail Example"], bg=MID_GRAY, text_color="000000")
    current_row += 1

    catalyst_data = [
        ("Named", "Catalyst has a specific name, not 'things will improve'",
         "CHAMPION-AF trial; FARAPULSE PFA approval", "'salesforce transition'; 'management focus'"),
        ("Dated", "Quarter/year known for the event",
         "ACC March 28, 2026; MiniMed spin end-2026", "'over the next 18 months'; 'eventually'"),
        ("Binary/Near-Binary", "Clear yes/no outcome with quantifiable price impact",
         "FDA approval (approved/not); Trial readout (positive/negative)",
         "'gradual improvement'; 'long-term margin expansion'"),
    ]
    for i, (req, desc, good, bad) in enumerate(catalyst_data):
        current_row = data_row(current_row, req, desc, good, bad, bold_first=True, shade=(i % 2 == 1))

    current_row += 1

    # ── Section 5: Positive Patterns (BSX/MDT) ──────────────────────────────
    current_row = section_header(current_row, "POSITIVE PATTERNS (Learned from BSX, MDT Deep Dives)", bg="375623")
    apply_header_row(ws, current_row, ["Pattern", "Description", "Example", "Score Impact"], bg=MID_GRAY, text_color="000000")
    current_row += 1

    positive_data = [
        ("Single-Event Overreaction",
         "Selloff caused by one discrete event; underlying business intact",
         "BSX: EP grew 35% vs ~40% expected — revenue beat, EPS beat, FCF +38%",
         "Selloff Quality: 9-10"),
        ("Named/Dated/Binary Catalysts",
         "Catalysts with specific name, date, and binary outcome",
         "BSX: CHAMPION-AF at ACC March 28, 2026 — TAM 5M→20M patients",
         "Catalyst Specificity: 9-10"),
        ("Analyst Support Despite Selloff",
         "Strong buy consensus maintained or upgraded during selloff",
         "BSX: 35 analysts, 0 sells, mean target +49% above current",
         "Analyst Conviction: 9-10"),
        ("Intact Core Franchise",
         "Core business undamaged — selloff is about fear, not fundamentals",
         "BSX: WATCHMAN 600K+ implants, Farawave ~70% US PFA share",
         "Competitive Moat: 9-10"),
        ("Management Acting (Not Talking)",
         "Concrete actions taken: acquisitions, board changes, strategic reviews",
         "MDT: Elliott added ex-Stryker CFO, MiniMed spin announced with timeline",
         "Catalyst Specificity: +1-2 pts"),
    ]
    for i, (pattern, desc, example, impact) in enumerate(positive_data):
        current_row = data_row(current_row, pattern, desc, example, impact, bold_first=True, shade=(i % 2 == 1))

    current_row += 1

    # ── Section 6: Failure Patterns (ZBH/UNH) ───────────────────────────────
    current_row = section_header(current_row, "FAILURE PATTERNS (Learned from ZBH, UNH Deep Dives)", bg="9C0006")
    apply_header_row(ws, current_row, ["Pattern", "Description", "Example", "Score Impact"], bg=MID_GRAY, text_color="000000")
    current_row += 1

    failure_data = [
        ("Criminal Investigation",
         "Active DOJ/FBI criminal probe → DISQUALIFIED. Unbounded tail risk.",
         "UNH: DOJ criminal + civil MA fraud investigation; mgmt denied then forced to acknowledge",
         "DISQUALIFIED — automatic fail"),
        ("Chronic Underperformance",
         "3+ years losing share to peers without confirmed inflection",
         "ZBH: Underperformed S&P 4 of 5 years; 1-3% organic growth vs industry avg",
         "DISQUALIFIED — automatic fail"),
        ("Structural Competitor Share Loss",
         "Better-resourced competitor structurally taking share in core segment",
         "SYK MAKO taking robotic knee share from ZBH ROSA; LLY gaining GLP-1 share from NVO",
         "Selloff Quality: -2pts; Moat: -1.5pts"),
        ("Macro Directly Hostile",
         "Bipartisan political targeting, regulatory crackdown on core business",
         "UNH: bipartisan healthcare reform pressure; CMS rate squeeze",
         "Macro Alignment: cap at 3.0"),
    ]
    for i, (pattern, desc, example, impact) in enumerate(failure_data):
        current_row = data_row(current_row, pattern, desc, example, impact, bold_first=True, shade=(i % 2 == 1))

    current_row += 1

    # ── Section 7: Sector Balance Rule ───────────────────────────────────────
    current_row = section_header(current_row, "SECTOR BALANCE RULE")
    apply_header_row(ws, current_row, ["Rule", "Description", "Sectors Targeted", "Notes"], bg=MID_GRAY, text_color="000000")
    current_row += 1

    sector_data = [
        ("Minimum Sectors", "Final output must include candidates from >= 3 sectors",
         "Healthcare, Defense/Aerospace, Consumer Staples, Industrials, Utilities, International",
         "Do not force bad picks to meet targets"),
        ("Healthcare Cap", "If >4 healthcare names and <2 non-healthcare, continue screening",
         "N/A", "Healthcare is naturally good fit but need diversification"),
        ("Sector Valuation", "Use sector-specific P/E benchmarks, not S&P average",
         "HC Devices 18-25x; Defense 18-22x; Staples 20-24x; Industrials 18-22x; Utilities 16-20x",
         "Don't penalize utility at 15x or reward defense at 20x"),
    ]
    for i, (rule, desc, sectors, notes) in enumerate(sector_data):
        current_row = data_row(current_row, rule, desc, sectors, notes, bold_first=True, shade=(i % 2 == 1))

    current_row += 1

    # ── Section 8: Screened Out Failure Types ────────────────────────────────
    current_row = section_header(current_row, "SCREENED-OUT FAILURE TYPE CATEGORIES")
    apply_header_row(ws, current_row, ["Category", "Description", "Action", "Revisit Trigger"], bg=MID_GRAY, text_color="000000")
    current_row += 1

    failure_types = [
        ("HARD FILTER: Criminal/DOJ probe",
         "Active criminal investigation by DOJ/FBI",
         "Immediate disqualification", "Investigation closed with no charges"),
        ("HARD FILTER: Chronic underperformer",
         "3+ years underperforming, no confirmed inflection",
         "Immediate disqualification", "2+ quarters of above-industry organic growth"),
        ("HARD FILTER: Structural share loss",
         "Competitor structurally gaining core segment share",
         "Disqualify unless clear fix", "Documented share stabilization or gain"),
        ("HARD FILTER: Macro hostile",
         "Regulatory/political environment actively hostile",
         "Cap macro score at 3.0", "Policy environment reversal"),
        ("SOFT PASS: Valuation insufficient",
         "Stock doesn't meet 30% ATH selloff or recovery filter",
         "Move to screened out", "Further selloff creates entry"),
        ("SOFT PASS: Catalyst too vague",
         "No named/dated/binary catalyst within 12-18 months",
         "Move to watchlist", "Specific catalyst emerges"),
    ]
    for i, (cat, desc, action, trigger) in enumerate(failure_types):
        current_row = data_row(current_row, cat, desc, action, trigger, bold_first=True, shade=(i % 2 == 1))

    # Column widths
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 40

    ws.freeze_panes = "A3"
    return ws


def main():
    print("Loading research data...")
    data = load_research_data()

    print("Building Excel model...")
    wb = Workbook()

    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    print("  -> Sheet 1: Summary Dashboard")
    build_summary_dashboard(wb, data)

    print("  -> Sheet 2: Stock Scores")
    build_stock_scores(wb, data)

    print("  -> Sheet 3: Valuation Comps")
    build_valuation_comps(wb, data)

    print("  -> Sheet 4: Fundamental Data")
    build_fundamental_data(wb, data)

    print("  -> Sheet 5: Catalyst Tracker")
    build_catalyst_tracker(wb, data)

    print("  -> Sheet 6: Risk Matrix")
    build_risk_matrix(wb, data)

    print("  -> Sheet 7: NVO Deep Dive")
    build_nvo_deep_dive(wb, data)

    print("  -> Sheet 8: Assumptions")
    build_assumptions(wb, data)

    print("  -> Sheet 9: Screener Logic")
    build_screener_logic(wb, data)

    os.makedirs("output", exist_ok=True)
    out_path = "output/stock_screener.xlsx"
    wb.save(out_path)
    print(f"\nModel saved to {out_path}")
    print("Now run: python scripts/recalc.py output/stock_screener.xlsx")


if __name__ == "__main__":
    main()
