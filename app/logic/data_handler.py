import datetime
from copy import copy
from openpyxl.utils import column_index_from_string
from openpyxl.styles import PatternFill
from .sorter import sort_data_rows
from .formula import apply_translated_formulas
from .formatting import apply_font_colors
from .zero_handler import replace_zeros_with_none_in_sheet


def process_data_per_month(sheet_a, sheet_b, month_value, month_abbreviation, header_columns_a, column_mapping):
    """
    Process and transfer monthly data from Sheet A to Sheet B.
    Includes data backup, clearing, copying, sorting, restoring styles,
    applying formulas, font coloring, and replacing zeros.
    """
    print(f"⏳ Processing month {month_abbreviation.upper()}...")

    # --- Locate the start row in Sheet B based on month abbreviation ---
    cut_start_row = None
    for row in range(1, sheet_b.max_row + 1):
        cell_val = sheet_b.cell(row=row, column=2).value
        if cell_val and isinstance(cell_val, str) and cell_val.strip().lower().startswith(month_abbreviation):
            cut_start_row = row + 3  # Data usually starts 3 rows below the header
            break
    if cut_start_row is None:
        print(f"❌ Month {month_abbreviation.upper()} not found in Sheet B.")
        return

    # --- Find the end row of the data block ---
    cut_end_row = cut_start_row
    while cut_end_row <= sheet_b.max_row and sheet_b.cell(row=cut_end_row, column=3).value:
        cut_end_row += 1
    cut_end_row -= 1

    # --- Define key column indexes ---
    n_col = column_index_from_string('N')
    bi_col = column_index_from_string('BJ')
    bl_col = column_index_from_string('BL')
    bm_col = column_index_from_string('BM')
    bs_col = column_index_from_string('BS')
    extra_cols = [bl_col, bm_col, bs_col]

    # --- Backup values, fills, and fonts from Sheet B (N–BI range + extra columns) ---
    cut_data_dict = {}
    for row in sheet_b.iter_rows(min_row=cut_start_row, max_row=cut_end_row,
                                 min_col=n_col, max_col=bs_col):
        row_idx = row[0].row
        vessel_name = sheet_b.cell(row=row_idx, column=5).value
        values_and_styles = {}
        for cell in row:
            col_idx = cell.column
            if n_col <= col_idx <= bi_col:
                values_and_styles[col_idx] = (
                    cell.value,
                    copy(cell.fill),
                    copy(cell.font)   # also keep font
                )
            elif col_idx in extra_cols:
                values_and_styles[col_idx] = (cell.value, None, None)  # value only
        cut_data_dict[vessel_name] = values_and_styles

    # --- Clear old block in Sheet B ---
    # N–BI: clear values + fill
    for row in sheet_b.iter_rows(min_row=cut_start_row, max_row=cut_end_row,
                                 min_col=n_col, max_col=bi_col):
        for cell in row:
            cell.value = None
            cell.fill = PatternFill()
    # BL, BM, BS: clear values only
    for col in extra_cols:
        for row in sheet_b.iter_rows(min_row=cut_start_row, max_row=cut_end_row,
                                     min_col=col, max_col=col):
            for cell in row:
                cell.value = None

    # --- Copy data from Sheet A to Sheet B ---
    current_row_b = cut_end_row + 1
    for row in range(2, sheet_a.max_row + 1):
        if sheet_a.cell(row=row, column=header_columns_a['Month']).value == month_value:
            for col_name, col_letter_b in column_mapping.items():
                col_index_a = header_columns_a.get(col_name)
                col_index_b = column_index_from_string(col_letter_b)
                if col_index_a is not None:
                    value = sheet_a.cell(row=row, column=col_index_a).value
                    if col_name == 'Month':
                        # Format month as date (January, February, etc.)
                        try:
                            date_obj = datetime.datetime(2025, int(value), 1)
                            cell_b = sheet_b.cell(row=current_row_b, column=col_index_b)
                            cell_b.value = date_obj
                            cell_b.number_format = '[$-en-US]mmm;@'
                        except Exception:
                            sheet_b.cell(row=current_row_b, column=col_index_b).value = value
                    else:
                        sheet_b.cell(row=current_row_b, column=col_index_b).value = value
            current_row_b += 1

    # --- Define sorting range ---
    sort_start = cut_start_row
    sort_end = current_row_b - 1

    # --- Extract data for sorting ---
    data_rows = []
    for row in sheet_b.iter_rows(min_row=sort_start, max_row=sort_end, values_only=False):
        data_rows.append([cell.value for cell in row])

    # --- Sort data ---
    data_rows_sorted = sort_data_rows(data_rows)

    # --- Overwrite Sheet B with sorted rows ---
    for i, row_data in enumerate(data_rows_sorted):
        for j, value in enumerate(row_data):
            sheet_b.cell(row=sort_start + i, column=j + 1, value=value)

    # --- Restore backed up values, fills, and fonts (N–BI range only) ---
    sorted_vessel_names = [row[4] for row in data_rows_sorted]
    for i, vessel_name in enumerate(sorted_vessel_names):
        values_and_styles = cut_data_dict.get(vessel_name)
        if values_and_styles:
            for col_idx, (val, fill, font) in values_and_styles.items():
                target_cell = sheet_b.cell(row=cut_start_row + i, column=col_idx)
                target_cell.value = val
                if n_col <= col_idx <= bi_col:  # restore style only for N–BI
                    target_cell.fill = fill
                    target_cell.font = font

    # --- Apply translated formulas to specific columns ---
    apply_translated_formulas(
        sheet_b,
        start_row=sort_start,
        end_row=sort_end,
        included_columns=['B', 'BJ', 'BO', 'ANO'],
        formulas={
            'B': '=IFERROR(1+OFFSET(B{row},-1,0,1,1),1)',
            'BJ': '=IFERROR(SUM(N{row}:BI{row}),"NULL")',
            'BO': '=(SUMIF($N$317:$BI$317,D{row},N{row}:BI{row}))/BJ{row}',
            'ANO': '=BJ{row}/AOA{row}'
        }
    )

    # --- Apply font coloring rules ---
    apply_font_colors(sheet_b, start_row=sort_start, end_row=sort_end)

    # --- Replace zeros with None (to avoid showing 0s in output) ---
    replace_zeros_with_none_in_sheet(sheet_b, start_row=sort_start, end_row=sort_end)

    print(f"✅ Successfully processed month {month_abbreviation.upper()} from Sheet A → Sheet B.")