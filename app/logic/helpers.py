import datetime
# import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from .auto_separator import get_formula_separator

sep = get_formula_separator()
# 📌 mapping formula kolom → pattern (bisa diperluas sesuai kebutuhan)
formulas={
    # 'B': f"=ROW()-ROW($B${sort_start})+1", # nomor urut otomatis
    'BJ': '=IFERROR(SUM(N{r}:BI{r}),"NULL")',
    'BO': '=(SUMIF($N$317:$BI$317,D{r},N{r}:BI{r}))/BJ{r}',
    'AKK': '=(AOH{r}/BJ{r})*-1',
    'ANO': '=IFERROR(BJ{r}/AOA{r},0)',
    'ANQ': '=J{r}',
    'ANS': '=ANQ{r}+(ANR{r}/24)',
    'ANT': '=K{r}',
    'ANU': '=L{r}',
    'ANX': '=BJ{r}',
    'AOA': '=(ANU{r}-ANT{r})*24',
    'AOB': '=(ANT{r}-ANS{r})*24',
    'AOC': '=(ANU{r}-ANS{r})*24',
    'AOD': f'=IF(BS{{r}}="Stevedore"{sep}10000{sep} IF(H{{r}}="BoCT"{sep}40000{sep} IF(H{{r}}="SMD Anc"{sep}25000{sep} IF(H{{r}}="GPK Port"{sep}10000{sep} IF(H{{r}}="Bunyut"{sep}25000{sep}0)))))',
    'AOE': '=(BJ{r}/AOD{r})*24',
    'AOF': '=(AOC{r}-AOE{r})/24',
    'AOH': '=AOF{r}*AOG{r}',
    'AOI': '=AOF{r}*-1',
    'AOJ': '=AOG{r}/2',
    'AOK': '=AOH{r}/2'
}

def reapply_formulas(sheet, start_row=2):
    """
    Terapkan ulang formula berdasarkan formulas.
    start_row default = 2 (anggap baris 1 header).
    """
    max_row = sheet.max_row
    for col_letter, pattern in formulas.items():
        for row in range(start_row, max_row + 1):
            formula = pattern.format(r=row)
            sheet[f"{col_letter}{row}"].value = formula

def delete_old_plan_rows(sheet_b, sheet_loading_old, header_columns_a, column_mapping):
    """
    Hapus baris di ITM Summary (sheet_b) yang memiliki Status='Plan'
    dan (Company, Vessel name, End user) cocok dengan data di sheet Loading lama.
    Setelah penghapusan, formula pada kolom target di-reapply ulang.
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

    # iterasi terbalik agar aman delete row
    for row in range(sheet_b.max_row, 1, -1):
        comp = sheet_b.cell(row=row, column=company_col).value
        ves  = sheet_b.cell(row=row, column=vessel_col).value
        eus  = sheet_b.cell(row=row, column=enduser_col).value
        status = sheet_b.cell(row=row, column=status_col).value
        if status == "Plan" and (comp, ves, eus) in plan_rows_keys:
            sheet_b.delete_rows(row, 1)

    # 📌 setelah semua delete → reapply formula
    reapply_formulas(sheet_b, start_row=2)

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