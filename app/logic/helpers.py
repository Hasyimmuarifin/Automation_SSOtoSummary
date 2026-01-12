import datetime
from openpyxl.utils import column_index_from_string, get_column_letter
from .auto_separator import get_formula_separator
from copy import copy

sep = get_formula_separator()

def normalize_month_block_rows(sheet, month_blocks, reference_col=2, renumber_func=None):
    """
    Normalize each monthly block in the Excel sheet:
      - Each monthly block must have 100 rows.
      - If there are fewer rows, add new rows.
      - Copy the style and formula from the last row that has a value in column B.
      - After each month block is complete → renumber the month blocks to update them.
      - Print the process log for debugging.

    Args:
        sheet: openpyxl worksheet object
        reference_col: reference column (default 2 = column B)
        formulas: list of formulas to be applied
    """
    print("   🔧 Starting normalize_month_block_rows...")

    i = 0
    while i < len(month_blocks):
        start_row, end_row = month_blocks[i]
        print(f"\n📦 Processing Block #{i+1}: start_row={start_row}, end_row={end_row}")

        # Find the last row in this block that has a value in column B.
        last_row = end_row
        for row in range(end_row, start_row - 1, -1):
            if sheet.cell(row=row, column=reference_col).value not in (None, ""):
                last_row = row
                break
        print(f"   📌 Last row with value in col B: {last_row}")

        current_block_size = end_row - start_row + 1
        if current_block_size >= 100:
            print("   ✅ Block already has 100 or more rows. Skipping...")
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
                    
                    # Copy formula if already exist
                    if old_cell.data_type == 'f':
                        new_cell.value = old_cell.value
                    else:
                        # Clear the values except for column B.
                        if col == reference_col and old_cell.data_type == 'f':
                            new_cell.value = old_cell.value
                        else:
                            new_cell.value = None
                last_row += 1

        # 🔄 After EVERY block is completed → renumber
        if renumber_func is not None:
            print("   🔄 Renumbering month blocks after current block...")
            month_blocks = renumber_func(sheet)
            print(f"\n📌 Updated month_blocks: {month_blocks}")

                
        i += 1  # proceed to the next block with the updated list

    print("\n   ✅ normalize_month_block_rows complete.")

def insert_boct_formulas(sheet, month_blocks, reference_col=2, loadport_col="H"):
    """
    For each monthly block:
       - Find the last row with content in column B (reference_col).
       - Find boct_start_row (the first row in the block with ‘BoCT’ in column H).
       - Find boct_end_row (the last row in the block with ‘BoCT’ in column H).
       - Add the formula in columns AKC and AKK, 2 rows below last_row.
    """

    loadport_idx = column_index_from_string(loadport_col)
    akc_idx = column_index_from_string("AKC")
    akk_idx = column_index_from_string("AKK")
    ano_idx = column_index_from_string("ANO")
    bj_idx = column_index_from_string("BJ")

    for block_no, (start_row, end_row) in enumerate(month_blocks, start=1):
        print(f"\n   📦 Processing BoCT Block #{block_no}: start={start_row}, end={end_row}")

        # Find last_row in column B
        last_row = end_row
        for row in range(end_row, start_row - 1, -1):
            if sheet.cell(row=row, column=reference_col).value not in (None, ""):
                last_row = row
                break
        print(f"   📌 Last row with value in col B: {last_row}")

        # Find boct_start_row & boct_end_row in column H
        boct_start_row, boct_end_row = None, None
        for row in range(start_row, end_row + 1):
            val = sheet.cell(row=row, column=loadport_idx).value
            if isinstance(val, str) and val.strip().lower() == "boct":
                if boct_start_row is None:
                    boct_start_row = row
                boct_end_row = row  # overwrite overwrite continuously → final result
        print(f"🔎 boct_start_row={boct_start_row}, boct_end_row={boct_end_row}")

        if boct_start_row and boct_end_row:
            target_row = last_row + 2  # dua baris di bawah last_row
            akc_col = get_column_letter(akc_idx)
            akk_col = get_column_letter(akk_idx)
            ano_col = get_column_letter(ano_idx)
            bj_col = get_column_letter(bj_idx)

            akc_formula = f"=SUMPRODUCT(${akc_col}${boct_start_row}:${akc_col}${boct_end_row}{sep}{bj_col}{boct_start_row}:{bj_col}{boct_end_row})/SUM(${bj_col}${boct_start_row}:${bj_col}${boct_end_row})"
            akk_formula = f"=SUMPRODUCT(${akk_col}${boct_start_row}:${akk_col}${boct_end_row}{sep}{bj_col}{boct_start_row}:{bj_col}{boct_end_row})/SUM(${bj_col}${boct_start_row}:${bj_col}${boct_end_row})"
            ano_formula = f"=AVERAGE({ano_col}{boct_start_row}:{ano_col}{boct_end_row})"

            sheet.cell(row=target_row, column=akc_idx).value = akc_formula
            sheet.cell(row=target_row, column=akk_idx).value = akk_formula
            sheet.cell(row=target_row, column=ano_idx).value = ano_formula

            print(f"✅ Inserted AKC formula at {akc_col}{target_row}: {akc_formula}")
            print(f"✅ Inserted AKK formula at {akk_col}{target_row}: {akk_formula}")
            print(f"✅ Inserted ANO formula at {ano_col}{target_row}: {ano_formula}")
        else:
            print("⚠️ Tidak ditemukan baris dengan Load Port = 'BoCT' pada blok ini.")

def insert_mahakam_formulas(sheet, month_blocks, reference_col=2, loadport_col="H", vessel_col="E"):
    """
    Untuk setiap blok bulan:
    - Cari baris terakhir dengan isi di kolom B (reference_col).
    - Cari mahakam_start_row (baris pertama dalam blok dengan 'SMD Anc', 'GPK Port', 'JBG Anc', 'Jorong', 'Bunyut' di kolom H).
    - Cari mahakam_end_row (baris terakhir dalam blok dengan 'SMD Anc', 'GPK Port', 'JBG Anc', 'Jorong', 'Bunyut' di kolom H).
    - subset MV. dan BG. (kolom E, Name of Vessel) dari range mahakam
    - Tambahkan formula:
        • AKC & AKK → last_row + 3
        • MV. (AKC) → last_row + 6
        • BG. (AKC) → last_row + 7
        • GLR TPH BoCT (ANO) → last_row + 2
        • GLR TPH Mahakam (ANO) → last_row + 3
    """

    loadport_idx = column_index_from_string(loadport_col)  # kolom H
    vessel_idx = column_index_from_string(vessel_col)      # kolom E
    akc_idx = column_index_from_string("AKC")
    akk_idx = column_index_from_string("AKK")
    ano_idx = column_index_from_string("ANO")
    bj_idx = column_index_from_string("BJ")

    target_loadports = {"smd anc", "gpk port", "jbg anc", "jorong", "bunyut"}

    for block_no, (start_row, end_row) in enumerate(month_blocks, start=1):
        print(f"\n📦 Processing Mahakam Block #{block_no}: start={start_row}, end={end_row}")

        # Cari last_row di kolom B
        last_row = end_row
        for row in range(end_row, start_row - 1, -1):
            if sheet.cell(row=row, column=reference_col).value not in (None, ""):
                last_row = row
                break
        print(f"📌 Last row with value in col B: {last_row}")

        # Cari mahakam_start_row & mahakam_end_row untuk Mahakam (berdasarkan Load Port)
        mahakam_start_row, mahakam_end_row = None, None
        for row in range(start_row, end_row + 1):
            val = sheet.cell(row=row, column=loadport_idx).value
            if isinstance(val, str) and val.strip().lower() in target_loadports:
                if mahakam_start_row is None:
                    mahakam_start_row = row
                mahakam_end_row = row  # overwrite terus → hasilnya terakhir
        print(f"🔎 mahakam_start_row={mahakam_start_row}, mahakam_end_row={mahakam_end_row}")

        if not (mahakam_start_row and mahakam_end_row):
            print("⚠️ Tidak ditemukan baris Mahakam pada blok ini.")
            continue

        akc_col = get_column_letter(akc_idx)
        akk_col = get_column_letter(akk_idx)
        ano_col = get_column_letter(ano_idx)
        bj_col = get_column_letter(bj_idx)

        # --- Formula Mahakam (semua target loadport)
        if mahakam_start_row and mahakam_end_row:
            target_row = last_row + 3
            akc_formula = f"=SUMPRODUCT(${akc_col}${mahakam_start_row}:${akc_col}${mahakam_end_row}{sep}${bj_col}${mahakam_start_row}:${bj_col}${mahakam_end_row})/SUM(${bj_col}${mahakam_start_row}:${bj_col}${mahakam_end_row})"
            akk_formula = f"=SUMPRODUCT(${akk_col}${mahakam_start_row}:${akk_col}${mahakam_end_row}{sep}${bj_col}${mahakam_start_row}:${bj_col}${mahakam_end_row})/SUM(${bj_col}${mahakam_start_row}:${bj_col}${mahakam_end_row})"
            ano_formula = f"=AVERAGE({ano_col}{mahakam_start_row}:{ano_col}{mahakam_end_row})"

            sheet.cell(row=target_row, column=akc_idx).value = akc_formula
            sheet.cell(row=target_row, column=akk_idx).value = akk_formula
            sheet.cell(row=target_row, column=ano_idx).value = ano_formula

            print(f"✅ Inserted AKC formula at {akc_col}{target_row}: {akc_formula}")
            print(f"✅ Inserted AKK formula at {akk_col}{target_row}: {akk_formula}")
            print(f"✅ Inserted ANO formula at {ano_col}{target_row}: {ano_formula}")

        # --- Formula MV. (subset vessel name contains "MV.")
        mahakam_mv_start, mahakam_mv_end = None, None
        for row in range(mahakam_start_row, mahakam_end_row + 1):
            val = sheet.cell(row=row, column=vessel_idx).value
            if isinstance(val, str) and val.strip().upper().startswith("MV."):
                if mahakam_mv_start is None:
                    mahakam_mv_start = row
                mahakam_mv_end = row
        print(f"🔎 mahakam_mv_start={mahakam_mv_start}, mahakam_mv_end={mahakam_mv_end}")

        if mahakam_mv_start and mahakam_mv_end:
            target_row = last_row + 6
            formula = f"=SUMPRODUCT(${akc_col}${mahakam_mv_start}:${akc_col}${mahakam_mv_end}{sep}${bj_col}${mahakam_mv_start}:${bj_col}${mahakam_mv_end})/SUM(${bj_col}${mahakam_mv_start}:${bj_col}${mahakam_mv_end})"
            sheet.cell(row=target_row, column=akc_idx).value = formula
            print(f"✅ Inserted MV formula at {akc_col}{target_row}: {formula}")

        # --- Cari BG. subset
        mahakam_bg_start, mahakam_bg_end = None, None
        for row in range(mahakam_start_row, mahakam_end_row + 1):
            val = sheet.cell(row=row, column=vessel_idx).value
            if isinstance(val, str) and val.strip().upper().startswith("BG."):
                if mahakam_bg_start is None:
                    mahakam_bg_start = row
                mahakam_bg_end = row
        print(f"🔎 mahakam_bg_start={mahakam_bg_start}, mahakam_bg_end={mahakam_bg_end}")

        if mahakam_bg_start and mahakam_bg_end:
            # Jika ditemukan subset BG (SMD Anc, GPK Port, dll)
            target_row = last_row + 7
            formula = f"=SUMPRODUCT(${akc_col}${mahakam_bg_start}:${akc_col}${mahakam_bg_end}{sep}${bj_col}${mahakam_bg_start}:${bj_col}${mahakam_bg_end})/SUM(${bj_col}${mahakam_bg_start}:${bj_col}${mahakam_bg_end})"
            sheet.cell(row=target_row, column=akc_idx).value = formula
            print(f"✅ Inserted BG formula at {akc_col}{target_row}: {formula}")

        else:
            # Default → isi 0 jika tidak ada subset BG
            target_row = last_row + 7
            sheet.cell(row=target_row, column=akc_idx).value = 0
            print(f"⚠️ Tidak ada subset BG, set {akc_col}{target_row} = 0")
            # # Default fallback pakai range mahakam langsung
            # if mahakam_start_row and mahakam_end_row:
            #     target_row = last_row + 7
            #     formula = f"=SUMPRODUCT(${akc_col}${mahakam_start_row}:${akc_col}${mahakam_end_row}{sep}${bj_col}${mahakam_start_row}:${bj_col}${mahakam_end_row})/SUM(${bj_col}${mahakam_start_row}:${bj_col}${mahakam_end_row})"
            #     sheet.cell(row=target_row, column=akc_idx).value = formula
            #     print(f"⚠️ Tidak ada subset BG, pakai default Mahakam range → {akc_col}{target_row}: {formula}")
            # else:
            #     print("❌ Tidak ada baris Mahakam sama sekali di blok ini.")

# 📌 mapping formula kolom → pattern (bisa diperluas sesuai kebutuhan)
formulas={
    # 'B': f"=ROW()-ROW($B${sort_start})+1",
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
    'AOD': f'=IF(BS{{row}}="Stevedore"{sep}10000{sep}IF(H{{row}}="BoCT"{sep}40000{sep}IF(H{{row}}="SMD Anc"{sep}25000{sep}IF(H{{row}}="GPK Port"{sep}10000{sep}IF(H{{row}}="Bunyut"{sep}25000{sep}IF(H{{row}}="Jorong"{sep}7000{sep}IF(H{{row}}="JBG Anc"{sep}10000{sep}0)))))))',
    'AOE': '=(BJ{row}/AOD{row})*24',
    'AOF': '=(AOC{row}-AOE{row})/24',
    'AOH': '=AOF{row}*AOG{row}',
    'AOI': '=AOF{row}*-1',
    'AOJ': '=AOG{row}/2',
    'AOK': '=AOH{row}/2'
}

def delete_or_clear_plan_rows(sheet_b, column_mapping, target_months: set[int]):
    """
    Hapus baris dengan Status == 'Plan' pada blok bulan yang termasuk target_months.
    Bulan di luar target_months hanya akan di-clear (kolom C..AOT).
    """

    status_col = column_index_from_string(column_mapping['Status'])
    month_header_col = column_index_from_string('B')
    start_clear_col = column_index_from_string("C")
    end_clear_col = column_index_from_string("AOT")

    print(f"📌 Target months for delete: {sorted(target_months)}")

    # Cari header bulan
    header_rows = []
    for r in range(1, sheet_b.max_row + 1):
        cell_val = sheet_b.cell(row=r, column=month_header_col).value
        if isinstance(cell_val, str):
            abbr = cell_val.strip().lower()[:3]
            if abbr in {'jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'}:
                header_rows.append((r, abbr))

    if not header_rows:
        print("⚠️ Tidak ditemukan blok bulan di kolom B.")
        return

    # Proses dari bawah ke atas
    header_rows.sort(key=lambda x: x[0], reverse=True)

    for header_row, abbr in header_rows:
        month_num = datetime.datetime.strptime(abbr.title(), "%b").month
        is_selected_month = month_num in target_months

        print(f"\n📅 Processing month block: {abbr.title()} "
              f"({'TARGET' if is_selected_month else 'other'})")

        data_start = header_row + 2
        data_end = min(sheet_b.max_row, data_start + 100 - 1)

        rows_to_delete = []

        for r in range(data_start, data_end + 1):
            status_val = sheet_b.cell(row=r, column=status_col).value
            if str(status_val).strip().lower() == "plan":
                if is_selected_month:
                    rows_to_delete.append(r)
                    print(f"   🗑️ Delete row {r}")
                else:
                    for c in range(start_clear_col, end_clear_col + 1):
                        sheet_b.cell(row=r, column=c).value = None
                    print(f"   🧹 Clear row {r}")

        for rr in reversed(rows_to_delete):
            sheet_b.delete_rows(rr, 1)

    print("\n✅ delete_or_clear_plan_rows selesai.")

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
    return datetime.date(2026, month_number, 1).strftime('%b').lower()

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