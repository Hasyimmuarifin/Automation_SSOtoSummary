# month_block_finder.py
from openpyxl import load_workbook

def find_month_block(sheet, month_value):
    """
    Find the start row & end row of a month block based on month number (1-12).
    Example: 1 -> 'jan', 2 -> 'feb', etc.
    Return tuple: (start_row, end_row)
    If montgh block not found, return (None, None)
    
    Parameter:
    - sheet : worksheet object (openpyxl)
    - month_value : int, ex: '1', '2', '3'
    """

    # --- Convert month_value to month prefix ---
    month_map = {
        1: "jan", 2: "feb", 3: "mar", 4: "apr",
        5: "may", 6: "jun", 7: "jul", 8: "aug",
        9: "sep", 10: "oct", 11: "nov", 12: "dec"
    }

    if month_value not in month_map:
        raise ValueError("month_value must be number 1–12")

    month_abbreviation = month_map[month_value]

    cut_start_row = None
    target = month_abbreviation.lower().strip()

    # --- Cari baris start (bulan) pada kolom B ---
    # --- Find start row (month) in column B ---
    for row in range(1, sheet.max_row + 1):
        cell_val = sheet.cell(row=row, column=2).value

        if isinstance(cell_val, str) and cell_val.strip().lower().startswith(target):
            cut_start_row = row + 3      # offset like the Logic in original code
            break

    if cut_start_row is None:
        # nothing month found
        return None, None

    # --- Find end row (month) in column B ---
    cut_end_row = cut_start_row
    while cut_end_row <= sheet.max_row:
        val_c = sheet.cell(row=cut_end_row, column=2).value

        # stop when column B is empty
        if val_c is None or str(val_c).strip() == "":
            break

        cut_end_row += 1

    cut_end_row -= 1  # back to last valid month row

    return cut_start_row, cut_end_row


def find_month_block_from_file(filepath, sheet_name, month_abbreviation):
    """
    Helper if want to directly call from Excel file.
    Return (start_row, end_row)
    """
    wb = load_workbook(filepath, data_only=True)
    sheet = wb[sheet_name]

    return find_month_block(sheet, month_abbreviation)