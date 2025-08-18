import datetime
# import openpyxl
from .renumber_blocks import renumber_month_blocks
from openpyxl.utils import column_index_from_string, get_column_letter
from .auto_separator import get_formula_separator

# Variabel Declaration :
TOTAL_BOCT_MARKER = "Total Coal Loading of BoCT"
TOTAL_MAHAKAM_MARKER = "Total Coal Loading of Mahakam"
TOTAL_MARKER = "Total Coal Loading of ITM (Coal Demand)"

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
    Terapkan ulang formula pada setiap blok bulan:
      - Isi ulang formula_map untuk setiap baris data (stop kalau kolom B kosong).
      - Khusus baris total (kolom L == TOTAL_BOCT_MARKER) di baris end_row+2:
        set formula SUMIF untuk kolom N..BJ menjumlahkan seluruh baris data (start_row..end_row).
      - Khusus baris total (kolom L == TOTAL_MAHAKAM_MARKER) di baris end_row+3:
        set formula SUMIFS untuk kolom N..BJ menjumlahkan seluruh baris data (start_row..end_row).
      - Khusus baris total (kolom J == TOTAL_MARKER) di baris end_row+4:
        set formula SUM untuk kolom N..BJ menjumlahkan seluruh baris data (start_row..end_row).
    start_row default = 2 (anggap baris 1 header).
    """
    col_b_index = column_index_from_string("B")
    col_j_index = column_index_from_string("J")
    col_l_index = column_index_from_string("L")
    n_col_idx   = column_index_from_string("N")
    bj_col_idx  = column_index_from_string("BJ")

    for (start_row, end_row) in month_blocks:
        # 1) Re-apply formula baris data (hingga B kosong)
        for row in range(start_row, end_row + 1):
            b_val = sheet.cell(row=row, column=col_b_index).value
            if b_val is None or str(b_val).strip() == "":
                # berhenti untuk blok ini ketika kolom B kosong
                break

            # apply semua formula map (placeholder {row})
            for col_letter, template in formula_map.items():
                col_idx = column_index_from_string(col_letter)
                sheet.cell(row=row, column=col_idx).value = template.format(row=row)

        # 1) Tangani baris total (diasumsikan 2 baris di bawah data akhir: end_row + 2)
        total_boct_row = end_row + 2
        j_val = sheet.cell(row=total_boct_row, column=col_l_index).value
        if isinstance(j_val, str) and j_val.strip() == TOTAL_BOCT_MARKER:
            # SUM seluruh blok bulan pada kolom N..BJ (hanya baris data: start_row..end_row
            if end_row >= start_row:
                for col_idx in range(n_col_idx, bj_col_idx + 1):
                    col_letter = get_column_letter(col_idx)
                    # Range Dinamis
                    sum_range = f"{col_letter}{start_row}:{col_letter}{end_row}"
                    crit_range = f"$H${start_row}:$H${end_row}"
                    # Formula dengan not equal "BoCT"
                    formula = f'=SUMIF({crit_range}{sep}"BoCT"{sep}{sum_range})'
                    sheet.cell(row=total_boct_row, column=col_idx).value = formula
                
            else :
                # Tidak ada data di blok → set 0
                for col_idx in range(n_col_idx, bj_col_idx + 1):
                    sheet.cell(row=total_boct_row, column=col_idx).value = 0

        # 2) Tangani baris total (diasumsikan 3 baris di bawah data akhir: end_row + 3)
        total_mahakam_row = end_row + 3
        j_val = sheet.cell(row=total_mahakam_row, column=col_l_index).value
        if isinstance(j_val, str) and j_val.strip() == TOTAL_MAHAKAM_MARKER:
            # SUM seluruh blok bulan pada kolom N..BJ (hanya baris data: start_row..end_row
            if end_row >= start_row:
                for col_idx in range(n_col_idx, bj_col_idx + 1):
                    col_letter = get_column_letter(col_idx)
                    # Range Dinamis
                    sum_range = f"{col_letter}{start_row}:{col_letter}{end_row}"
                    crit_range = f"$H${start_row}:$H${end_row}"
                    # Formula dengan not equal "BoCT"
                    formula = f'=SUMIFS({sum_range}{sep}{crit_range}{sep}"<>"&"BoCT")'
                    sheet.cell(row=total_mahakam_row, column=col_idx).value = formula
                
            else :
                # Tidak ada data di blok → set 0
                for col_idx in range(n_col_idx, bj_col_idx + 1):
                    sheet.cell(row=total_row, column=col_idx).value = 0

        # 3) Tangani baris total (diasumsikan 4 baris di bawah data akhir: end_row + 4)
        total_row = end_row + 4
        j_val = sheet.cell(row=total_row, column=col_j_index).value
        if isinstance(j_val, str) and j_val.strip() == TOTAL_MARKER:
            # SUM seluruh blok bulan pada kolom N..BJ (hanya baris data: start_row..end_row
            if end_row >= start_row:
                for col_idx in range(n_col_idx, bj_col_idx + 1):
                    col_letter = get_column_letter(col_idx)
                    sheet.cell(row=total_row, column=col_idx).value = (
                        f"=SUM({col_letter}{start_row}:{col_letter}{end_row})"
                    )
                
            else :
                # Tidak ada data di blok → set 0
                for col_idx in range(n_col_idx, bj_col_idx + 1):
                    sheet.cell(row=total_row, column=col_idx).value = 0


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