"""
Python port of Browser Export Duplicate Reformat.osts (Excel Office Script).

Drops ID / Corporate ID / Initial Amount columns, sorts rows ascending by
Batch Label, and flags (via conditional formatting, no data is removed):
  - duplicate Batch Label values (red)
  - Vol between 0.01 and 60 (red) - low remaining volume
  - "CONSUMED" in Plate/BOX barcode (yellow)

Requires: pandas, openpyxl
"""

from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill, Font
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.utils import get_column_letter

DROPPED_COLS = ["ID", "Corporate ID", "Initial Amount"]

NAMED_SOURCE_COLS = {
    "ID", "Barcode", "Corporate ID", "Batch Label", "Plate/BOX barcode",
    "Site", "Location", "Well", "Container Type", "Comments",
    "Initial Amount", "Current Amount", "Amount Unit", "TARE_WEIGHT_GM",
    "Solvent", "Vol", "Vol Units", "Conc", "Conc Unit", "Parent Barcode",
    "SOURCE", "SOURCE_ID", "FORMULA_WEIGHT", "CREATED_DATE",
    "MODIFIED_DATE", "AUTOWEIGH_TYPE",
}

RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
RED_FONT = Font(color="9C0006")
YELLOW_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
YELLOW_FONT = Font(color="9C5700")


def _validate_headers(df: pd.DataFrame) -> None:
    missing = [col for col in NAMED_SOURCE_COLS if col not in df.columns]
    if missing:
        raise ValueError(
            "Input file is missing expected column(s): " + ", ".join(sorted(missing))
        )


def reformat_sample_search(input_path: str) -> str:
    """Reformats input_path and writes the result alongside it as
    "<input file name>_FormattedSampleSearch.xlsx". Returns the output path.
    """
    input_file = Path(input_path)
    output_path = input_file.with_name(f"{input_file.stem}_FormattedSampleSearch.xlsx")

    df = pd.read_excel(input_path)
    _validate_headers(df)

    out = df.drop(columns=DROPPED_COLS)
    out = out.sort_values("Batch Label", ascending=True, kind="stable").reset_index(drop=True)

    out.to_excel(output_path, index=False)
    _apply_formatting(output_path, len(out))

    return str(output_path)


def _apply_formatting(path, n_rows: int) -> None:
    wb = load_workbook(path)
    ws = wb.active

    headers = {cell.value: cell.column for cell in ws[1]}
    batch_label_col = headers["Batch Label"]
    vol_col = headers["Vol"]
    plate_box_col = headers["Plate/BOX barcode"]
    last_col = ws.max_column

    first_data_row = 2
    last_data_row = first_data_row + n_rows - 1

    # Borders + approximate autofit across the full table.
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=1, max_row=last_data_row, min_col=1, max_col=last_col):
        for cell in row:
            cell.border = border

    for col_cells in ws.iter_cols(min_row=1, max_row=last_data_row):
        width = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = width + 2

    ws.auto_filter.ref = f"A1:{get_column_letter(last_col)}{last_data_row}"

    # Duplicate Batch Label values -> red.
    batch_letter = get_column_letter(batch_label_col)
    batch_range = f"{batch_letter}{first_data_row}:{batch_letter}{last_data_row}"
    ws.conditional_formatting.add(
        batch_range,
        FormulaRule(
            formula=[
                f'COUNTIF(${batch_letter}${first_data_row}:${batch_letter}${last_data_row},'
                f'{batch_letter}{first_data_row})>1'
            ],
            fill=RED_FILL, font=RED_FONT,
        ),
    )

    # Vol between 0.01 and 60 -> red (low remaining volume).
    vol_letter = get_column_letter(vol_col)
    vol_range = f"{vol_letter}{first_data_row}:{vol_letter}{last_data_row}"
    ws.conditional_formatting.add(
        vol_range,
        CellIsRule(operator="between", formula=["0.01", "60"], fill=RED_FILL, font=RED_FONT),
    )

    # "CONSUMED" in Plate/BOX barcode -> yellow.
    plate_box_letter = get_column_letter(plate_box_col)
    plate_box_range = f"{plate_box_letter}{first_data_row}:{plate_box_letter}{last_data_row}"
    ws.conditional_formatting.add(
        plate_box_range,
        FormulaRule(
            formula=[f'NOT(ISERROR(SEARCH("CONSUMED",{plate_box_letter}{first_data_row})))'],
            fill=YELLOW_FILL, font=YELLOW_FONT,
        ),
    )

    wb.save(path)


if __name__ == "__main__":
    import sys
    print(reformat_sample_search(sys.argv[1]))
