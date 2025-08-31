from openpyxl.utils import get_column_letter, column_index_from_string

def replace_zeros_with_none_in_sheet(sheet, start_row, end_row, start_col=1, end_col=None):
    """
    Replace all 0 (int, float, or string "0") with None directly in the worksheet,
    except in columns AKC–AKI, AKK–AKQ, and AOG.
    
    Parameters:
        sheet (Worksheet): The target sheet to clean.
        start_row (int): Starting row of the range.
        end_row (int): Ending row of the range.
        start_col (int): Starting column (default = 1).
        end_col (int): Ending column (default = sheet.max_column).
    """
    if end_col is None:
        end_col = sheet.max_column

    # Tentukan index kolom yang harus di-skip
    skip_cols = set()

    # AKC–AKI
    start_skip = column_index_from_string("AKC")
    end_skip   = column_index_from_string("AKI")
    skip_cols.update(range(start_skip, end_skip + 1))

    # AKK–AKQ
    start_skip = column_index_from_string("AKK")
    end_skip   = column_index_from_string("AKQ")
    skip_cols.update(range(start_skip, end_skip + 1))

    # AOG
    skip_cols.add(column_index_from_string("AOG"))

    for row in sheet.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col):
        for cell in row:
            if cell.column in skip_cols:
                continue  # lewati kolom yang tidak boleh diubah
            if cell.value is not None and str(cell.value).strip() == "0":
                cell.value = None