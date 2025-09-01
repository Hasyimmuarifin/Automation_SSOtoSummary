import datetime
# import openpyxl
from .renumber_blocks import renumber_month_blocks
from openpyxl.utils import column_index_from_string, get_column_letter
from .auto_separator import get_formula_separator
from .formula import reapply_formulas
from copy import copy

sep = get_formula_separator()

def normalize_month_block_rows(sheet, month_blocks, reference_col=2, renumber_func=None, formulas=None):
    """
    Normalisasi setiap blok bulan di sheet Excel:
    - Setiap blok bulan minimal 100 baris.
    - Jika kurang, tambahkan baris baru.
    - Copy style dan formula dari baris terakhir yang memiliki value di kolom B.
    - Setelah setiap blok selesai → renumbering blocks agar update.
    - Setelah setiap blok selesai → panggil reapply_formulas untuk blok itu.
    - Print log proses untuk debugging.

    Args:
        sheet: openpyxl worksheet object
        month_blocks: list of tuples (start_row, end_row) hasil renumber_month_blocks
        reference_col: kolom yang dijadikan acuan (default 2 = kolom B)
        renumber_func: fungsi untuk renumbering blok bulan (misalnya renumber_month_blocks)
        formulas: daftar formula yang akan diaplikasikan
    """
    print("🔧 Starting normalize_month_block_rows...")

    i = 0
    while i < len(month_blocks):
        start_row, end_row = month_blocks[i]
        print(f"\n📦 Processing Block #{i+1}: start_row={start_row}, end_row={end_row}")

        # Cari baris terakhir di blok ini yang memiliki value di kolom B
        last_row = end_row
        for row in range(end_row, start_row - 1, -1):
            if sheet.cell(row=row, column=reference_col).value not in (None, ""):
                last_row = row
                break
        print(f"📌 Last row with value in col B: {last_row}")

        current_block_size = end_row - start_row + 1
        if current_block_size >= 100:
            print("✅ Block already has 100 or more rows. Skipping...")
        else:
            rows_to_add = 100 - current_block_size
            print(f"➕ Adding {rows_to_add} row(s) to reach 100 rows.")

            for _ in range(rows_to_add):
                sheet.insert_rows(last_row + 1)
                for col in range(1, sheet.max_column + 1):
                    old_cell = sheet.cell(row=last_row, column=col)
                    new_cell = sheet.cell(row=last_row + 1, column=col)

                    # Copy style
                    if old_cell.has_style:
                        new_cell._style = copy(old_cell._style)
                    
                    # Copy formula jika ada
                    if old_cell.data_type == 'f':
                        new_cell.value = old_cell.value
                    else:
                        # Kosongkan value kecuali kolom B
                        if col == reference_col and old_cell.data_type == 'f':
                            new_cell.value = old_cell.value
                        else:
                            new_cell.value = None
                last_row += 1

        # 🔄 Setelah SETIAP blok selesai → renumber
        if renumber_func is not None:
            print("🔄 Renumbering month blocks after current block...")
            month_blocks = renumber_func(sheet)
            print(f"📌 Updated month_blocks: {month_blocks}")

                
        i += 1  # lanjut ke blok berikutnya dengan list yang sudah update

    print("✅ normalize_month_block_rows complete.")


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
    renumber_month_blocks(sheet_b)
    # month_blocks = renumber_month_blocks(sheet_b)

    # # 📌 setelah semua delete → reapply formula hanya dalam blok bulan
    # reapply_formulas(sheet_b, month_blocks, formulas)

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