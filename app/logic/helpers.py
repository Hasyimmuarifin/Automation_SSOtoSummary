import datetime
# import openpyxl
from .renumber_blocks import renumber_month_blocks
from openpyxl.utils import column_index_from_string, get_column_letter
from .auto_separator import get_formula_separator

sep = get_formula_separator()
# 📌 mapping formula kolom → pattern (bisa diperluas sesuai kebutuhan)
formulas={
    # 'B': f"=ROW()-ROW($B${sort_start})+1", # nomor urut otomatis
    'BJ': '=IFERROR(SUM(N{row}:BI{row}),"NULL")',
    'BO': '=(SUMIF($N$317:$BI$317,D{row},N{row}:BI{row}))/BJ{row}',
    'AKK': '=(AOH{row}/BJ{row})*-1',
    'ANO': '=IFERROR(BJ{row}/AOA{row},0)',
    'ANQ': '=J{row}',
    'ANS': '=ANQ{row}+(ANR{row}/24)',
    'ANT': '=K{row}',
    'ANU': '=L{row}',
    'ANX': '=BJ{row}',
    'AOA': '=(ANU{row}-ANT{row})*24',
    'AOB': '=(ANT{row}-ANS{row})*24',
    'AOC': '=(ANU{row}-ANS{row})*24',
    'AOD': f'=IF(BS{{row}}="Stevedore"{sep}10000{sep} IF(H{{row}}="BoCT"{sep}40000{sep} IF(H{{row}}="SMD Anc"{sep}25000{sep} IF(H{{row}}="GPK Port"{sep}10000{sep} IF(H{{row}}="Bunyut"{sep}25000{sep}0)))))',
    'AOE': '=(BJ{row}/AOD{row})*24',
    'AOF': '=(AOC{row}-AOE{row})/24',
    'AOH': '=AOF{row}*AOG{row}',
    'AOI': '=AOF{row}*-1',
    'AOJ': '=AOG{row}/2',
    'AOK': '=AOH{row}/2'
}

def reapply_formulas(sheet, month_blocks, formula_map):
    """
    Terapkan ulang formula berdasarkan formulas.
    start_row default = 2 (anggap baris 1 header).
    """
    col_b_index = column_index_from_string("B")

    for (start_row, end_row) in month_blocks:
        for row in range(start_row, end_row + 1):
            if not sheet.cell(row=row, column=col_b_index).value:
                break
            for col_letter, formula_template in formula_map.items():
                col_index = column_index_from_string(col_letter)
                sheet.cell(row=row, column=col_index).value = formula_template.format(row=row)

def delete_old_plan_rows(sheet_b, sheet_loading_old, header_columns_a, column_mapping):
    """
    Hapus baris di ITM Summary (sheet_b) yang memiliki Status='Plan'
    dan (Company, Vessel name, End user) cocok dengan data di sheet Loading lama.
    Setelah penghapusan, formulas di setiap blok bulan di-reapply ulang.
    """
    plan_rows_keys = set()

    # ambil keys dari Loading lama
    for row in range(2, sheet_loading_old.max_row + 1):
        comp = sheet_loading_old.cell(row=row, column=header_columns_a['Company']).value
        ves  = sheet_loading_old.cell(row=row, column=header_columns_a['Vessel name']).value
        eus  = sheet_loading_old.cell(row=row, column=header_columns_a['End user']).value
        plan_rows_keys.add((comp, ves, eus))

    company_col = column_index_from_string(column_mapping['Company'])
    vessel_col  = column_index_from_string(column_mapping['Vessel name'])
    enduser_col = column_index_from_string(column_mapping['End user'])
    status_col  = column_index_from_string(column_mapping['Status'])

    # iterasi terbalik agar aman delete row --> Hapus baris yang match
    for row in range(sheet_b.max_row, 1, -1):
        comp = sheet_b.cell(row=row, column=company_col).value
        ves  = sheet_b.cell(row=row, column=vessel_col).value
        eus  = sheet_b.cell(row=row, column=enduser_col).value
        status = sheet_b.cell(row=row, column=status_col).value
        if status == "Plan" and (comp, ves, eus) in plan_rows_keys:
            sheet_b.delete_rows(row, 1)

    # --- Renumber ulang blok bulan ---
    month_blocks = renumber_month_blocks(sheet_b)

    # 📌 setelah semua delete → reapply formula hanya dalam blok bulan
    reapply_formulas(sheet_b, month_blocks, formulas)

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