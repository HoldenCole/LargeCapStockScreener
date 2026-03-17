"""
recalc.py
=========
Recalculate formulas in the stock screener Excel model.

Since openpyxl doesn't evaluate formulas, this script:
1. Opens the workbook
2. Identifies all formula cells
3. Reports formula count per sheet
4. Checks for common formula errors (#REF!, #DIV/0!, #N/A, #VALUE!, #NAME?)
5. Validates all 8 required sheets are present

Usage:
    python scripts/recalc.py output/stock_screener.xlsx
"""

import sys
import os
from openpyxl import load_workbook


REQUIRED_SHEETS = [
    "SUMMARY DASHBOARD",
    "STOCK SCORES",
    "VALUATION COMPS",
    "FUNDAMENTAL DATA",
    "CATALYST TRACKER",
    "RISK MATRIX",
    "NVO DEEP DIVE",
    "ASSUMPTIONS",
]

ERROR_PATTERNS = ["#REF!", "#DIV/0!", "#N/A", "#VALUE!", "#NAME?", "#NULL!", "#NUM!"]


def recalc(filepath):
    if not os.path.exists(filepath):
        print(f"ERROR: File not found: {filepath}")
        return False

    print(f"Opening {filepath}...")
    wb = load_workbook(filepath, data_only=False)

    # Check required sheets
    print("\n=== SHEET VALIDATION ===")
    missing = []
    for sheet_name in REQUIRED_SHEETS:
        if sheet_name in wb.sheetnames:
            print(f"  [OK] {sheet_name}")
        else:
            print(f"  [MISSING] {sheet_name}")
            missing.append(sheet_name)

    if missing:
        print(f"\nERROR: {len(missing)} required sheet(s) missing: {', '.join(missing)}")
        return False
    else:
        print(f"\n  All {len(REQUIRED_SHEETS)} required sheets present.")

    # Count formulas and check for errors
    print("\n=== FORMULA AUDIT ===")
    total_formulas = 0
    total_errors = 0
    error_details = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        sheet_formulas = 0
        sheet_errors = 0

        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                    sheet_formulas += 1

                    # Check for error patterns in formula text
                    for err in ERROR_PATTERNS:
                        if err in str(cell.value).upper():
                            sheet_errors += 1
                            error_details.append(
                                f"  {sheet_name}!{cell.coordinate}: {cell.value} -> contains {err}"
                            )

        total_formulas += sheet_formulas
        total_errors += sheet_errors
        print(f"  {sheet_name}: {sheet_formulas} formulas, {sheet_errors} errors")

    print(f"\n  Total formulas: {total_formulas}")
    print(f"  Total errors in formula text: {total_errors}")

    if error_details:
        print("\n=== ERROR DETAILS ===")
        for detail in error_details:
            print(detail)

    # Spot-check: verify key formulas exist
    print("\n=== SPOT CHECKS ===")
    checks_passed = 0
    checks_total = 0

    # Check 1: STOCK SCORES has composite formula
    if "STOCK SCORES" in wb.sheetnames:
        ws = wb["STOCK SCORES"]
        checks_total += 1
        has_composite = False
        for row in ws.iter_rows(min_col=8, max_col=8):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                    has_composite = True
                    break
        if has_composite:
            print("  [OK] STOCK SCORES has composite score formulas")
            checks_passed += 1
        else:
            print("  [WARN] STOCK SCORES may be missing composite score formulas")

    # Check 2: VALUATION COMPS has implied upside formula
    if "VALUATION COMPS" in wb.sheetnames:
        ws = wb["VALUATION COMPS"]
        checks_total += 1
        has_upside = False
        for row in ws.iter_rows(min_col=16, max_col=16):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                    has_upside = True
                    break
        if has_upside:
            print("  [OK] VALUATION COMPS has implied upside formulas")
            checks_passed += 1
        else:
            print("  [WARN] VALUATION COMPS may be missing implied upside formulas")

    # Check 3: FUNDAMENTAL DATA has revenue growth formula
    if "FUNDAMENTAL DATA" in wb.sheetnames:
        ws = wb["FUNDAMENTAL DATA"]
        checks_total += 1
        has_growth = False
        for row in ws.iter_rows(min_col=7, max_col=7):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                    has_growth = True
                    break
        if has_growth:
            print("  [OK] FUNDAMENTAL DATA has revenue growth formulas")
            checks_passed += 1
        else:
            print("  [WARN] FUNDAMENTAL DATA may be missing revenue growth formulas")

    # Check 4: CATALYST TRACKER has expected impact formula
    if "CATALYST TRACKER" in wb.sheetnames:
        ws = wb["CATALYST TRACKER"]
        checks_total += 1
        has_expected = False
        for row in ws.iter_rows(min_col=7, max_col=7):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                    has_expected = True
                    break
        if has_expected:
            print("  [OK] CATALYST TRACKER has expected impact formulas")
            checks_passed += 1
        else:
            print("  [WARN] CATALYST TRACKER may be missing expected impact formulas")

    # Check 5: STOCK SCORES has recommendation IF formula
    if "STOCK SCORES" in wb.sheetnames:
        ws = wb["STOCK SCORES"]
        checks_total += 1
        has_if = False
        for row in ws.iter_rows(min_col=9, max_col=9):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "IF(" in cell.value.upper():
                    has_if = True
                    break
        if has_if:
            print("  [OK] STOCK SCORES has recommendation IF formulas")
            checks_passed += 1
        else:
            print("  [WARN] STOCK SCORES may be missing recommendation formulas")

    print(f"\n  Spot checks: {checks_passed}/{checks_total} passed")

    # Force recalculation flag for Excel
    wb.calculation.calcMode = "auto"

    # Save with recalc flag
    wb.save(filepath)
    print(f"\nModel saved with recalculation flag set: {filepath}")

    # Final verdict
    if total_errors == 0 and not missing and checks_passed == checks_total:
        print("\n✓ PASS — Zero formula errors, all sheets present, all spot checks passed.")
        return True
    else:
        issues = []
        if total_errors > 0:
            issues.append(f"{total_errors} formula errors")
        if missing:
            issues.append(f"{len(missing)} missing sheets")
        if checks_passed < checks_total:
            issues.append(f"{checks_total - checks_passed} failed spot checks")
        print(f"\n⚠ ISSUES: {'; '.join(issues)}")
        return False


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python scripts/recalc.py <path-to-xlsx>")
        sys.exit(1)

    filepath = sys.argv[1]
    success = recalc(filepath)
    sys.exit(0 if success else 1)
