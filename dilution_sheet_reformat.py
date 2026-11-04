"""
Python port of DilutionSheetReFormat.osts (Excel Office Script).

Takes the raw DilutionSheet export (.xlsx) and reformats it into the
weighing-workflow layout: reorders columns, inserts blank BIN/WELL/
INVENTORY AMOUNT fields, and adds BATCH CHECK / BARCODE CHECK columns
with live formulas + conditional formatting for manual verification
during weighing.

Requires: pandas, openpyxl
"""

from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

DILUTION_VOLUME_COLS = [
    "50uL DMSO for 10.0mM",
    "100uL DMSO for 10.0mM",
    "150uL DMSO for 10.0mM",
    "200uL DMSO for 10.0mM",
    "350uL DMSO for 10.0mM",
    "500uL DMSO for 10.0mM",
]

NAMED_SOURCE_COLS = {
    "ID", "Name", "Barcode", "FW", "mg in vial",
    "mg for 100uL 10.0mM", "uL DMSO for 10.0mM", "Note/External ID",
    "AUTOWEIGH_TYPE", *DILUTION_VOLUME_COLS,
}

GREEN_FILL = {"fgColor": "C6EFCE"}
RED_FILL = {"fgColor": "FFC7CE"}


def _validate_headers(df: pd.DataFrame) -> None:
    missing = [col for col in NAMED_SOURCE_COLS if col not in df.columns]
    if missing:
        raise ValueError(
            "Input file is missing expected column(s): " + ", ".join(sorted(missing))
        )


def reformat_dilution_sheet(input_path: str) -> str:
    """Reformats input_path and writes the result alongside it as
    "<input file name>_FormattedDilutionSheet.xlsx". Returns the output path.
    """
    input_file = Path(input_path)
    output_path = input_file.with_name(f"{input_file.stem}_FormattedDilutionSheet.xlsx")

    df = pd.read_excel(input_path)
    _validate_headers(df)

    # Any columns in the source beyond the named set (e.g. a trailing
    # unnamed/duplicate column some exports include) are dropped rather
    # than carried through.
    ordered_cols = (
        ["Barcode", "Name", "FW"]
        + ["BIN", "WELL", "INVENTORY AMOUNT"]
        + ["mg in vial", "mg for 100uL 10.0mM", "uL DMSO for 10.0mM", "Note/External ID"]
        + DILUTION_VOLUME_COLS
        + ["AUTOWEIGH_TYPE"]
    )

    out = pd.DataFrame(index=df.index)
    for col in ordered_cols:
        out[col] = df[col] if col in df.columns else ""

    out["BATCH CHECK"] = ""
    out["BARCODE CHECK"] = ""

    out.to_excel(output_path, index=False)

    _apply_formatting(output_path, len(out))

    return str(output_path)


def _apply_formatting(path: str, n_rows: int) -> None:
    wb = load_workbook(path)
    ws = wb.active

    headers = {cell.value: cell.column for cell in ws[1]}
    barcode_col = headers["Barcode"]
    name_col = headers["Name"]
    last_dilution_col = headers[DILUTION_VOLUME_COLS[-1]]
    batch_check_col = headers["BATCH CHECK"]
    barcode_check_col = headers["BARCODE CHECK"]

    first_data_row = 2
    last_data_row = first_data_row + n_rows - 1

    # Insert the live-formula columns right after BARCODE CHECK, then
    # right after BATCH CHECK (rightmost first so earlier column indices
    # stay valid).
    ws.insert_cols(barcode_check_col + 1)
    ws.insert_cols(batch_check_col + 1)

    # BATCH CHECK's formula column shifted right by the BARCODE CHECK
    # insert above it, so barcode_check_col grew by 1; recompute.
    barcode_check_col += 1
    batch_check_formula_col = batch_check_col + 1
    barcode_check_formula_col = barcode_check_col + 1

    batch_letter = get_column_letter(name_col)
    barcode_letter = get_column_letter(barcode_col)
    batch_check_letter = get_column_letter(batch_check_col)
    barcode_check_letter = get_column_letter(barcode_check_col)
    batch_formula_letter = get_column_letter(batch_check_formula_col)
    barcode_formula_letter = get_column_letter(barcode_check_formula_col)

    for row in range(first_data_row, last_data_row + 1):
        ws.cell(row=row, column=batch_check_formula_col).value = (
            f'=IF({batch_check_letter}{row}={batch_letter}{row},"GOOD","BAD")'
        )
        ws.cell(row=row, column=barcode_check_formula_col).value = (
            f'=IF({barcode_check_letter}{row}={barcode_letter}{row},"GOOD","BAD")'
        )

    # Borders on Barcode..last dilution-volume column (matches the
    # original script's A:P range).
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(
        min_row=1, max_row=last_data_row,
        min_col=barcode_col, max_col=last_dilution_col,
    ):
        for cell in row:
            cell.border = border

    # Approximate autofit.
    for col_cells in ws.iter_cols(min_row=1, max_row=last_data_row):
        width = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = width + 2

    # Conditional formatting: GOOD -> green, BAD -> red, across the
    # formula-check / manual-entry columns (mirrors original S:U range).
    cf_range = f"{batch_formula_letter}{first_data_row}:{barcode_formula_letter}{last_data_row}"
    anchor = f"{batch_formula_letter}{first_data_row}"
    ws.conditional_formatting.add(
        cf_range,
        FormulaRule(formula=[f'NOT(ISERROR(SEARCH("GOOD",{anchor})))'], fill=_fill(GREEN_FILL)),
    )
    ws.conditional_formatting.add(
        cf_range,
        FormulaRule(formula=[f'NOT(ISERROR(SEARCH("BAD",{anchor})))'], fill=_fill(RED_FILL)),
    )

    wb.save(path)


def _fill(colors):
    from openpyxl.styles import PatternFill
    return PatternFill(start_color=colors["fgColor"], end_color=colors["fgColor"], fill_type="solid")


if __name__ == "__main__":
    import sys
    print(reformat_dilution_sheet(sys.argv[1]))
