import datetime
import openpyxl

def delete_old_plan_rows(sheet_b, sheet_loading_old, header_columns_a, column_mapping):
    """
    Hapus baris di ITM Summary (sheet_b) yang memiliki Status='Plan'
    dan (Company, Vessel name, End user) cocok dengan data di sheet Loading lama.
    """
    plan_rows_keys = set()

    # ambil keys dari Loading lama
    for row in range(2, sheet_loading_old.max_row + 1):
        comp = sheet_loading_old.cell(row=row, column=header_columns_a['Company']).value
        ves  = sheet_loading_old.cell(row=row, column=header_columns_a['Vessel name']).value
        eus  = sheet_loading_old.cell(row=row, column=header_columns_a['End user']).value
        plan_rows_keys.add((comp, ves, eus))

    company_col = openpyxl.utils.column_index_from_string(column_mapping['Company'])
    vessel_col  = openpyxl.utils.column_index_from_string(column_mapping['Vessel name'])
    enduser_col = openpyxl.utils.column_index_from_string(column_mapping['End user'])
    status_col  = openpyxl.utils.column_index_from_string(column_mapping['Status'])

    # iterasi terbalik agar aman delete row
    for row in range(sheet_b.max_row, 1, -1):
        comp = sheet_b.cell(row=row, column=company_col).value
        ves  = sheet_b.cell(row=row, column=vessel_col).value
        eus  = sheet_b.cell(row=row, column=enduser_col).value
        status = sheet_b.cell(row=row, column=status_col).value
        if status == "Plan" and (comp, ves, eus) in plan_rows_keys:
            sheet_b.delete_rows(row, 1)

def month_to_abbreviation(month_number):
    """
    Convert a numeric month (1–12) to its lowercase 3-letter English abbreviation.
    
    Example:
        1 -> 'jan', 2 -> 'feb', ..., 12 -> 'dec'
    
    Parameters:
        month_number (int): The numeric representation of the month.
    
    Returns:
        str: The lowercase abbreviated month name.
    """
    return datetime.date(2025, month_number, 1).strftime('%b').lower()


def get_header_columns_a(sheet_a, column_mapping):
    """
    Map column headers in Sheet A to their corresponding column indices,
    based on the defined column mapping.

    Parameters:
        sheet_a (Worksheet): The source worksheet to analyze.
        column_mapping (dict): Dictionary of required column names.

    Returns:
        dict: A dictionary mapping column names to their index in Sheet A.
    """
    header_columns = {}
    for col in range(1, sheet_a.max_column + 1):
        val = sheet_a.cell(row=1, column=col).value
        if val in column_mapping:
            header_columns[val] = col
    return header_columns