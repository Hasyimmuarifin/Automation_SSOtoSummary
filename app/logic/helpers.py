import datetime
# import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from .auto_separator import get_formula_separator
from copy import copy

sep = get_formula_separator()

def normalize_month_block_rows(sheet, month_blocks, reference_col=2, renumber_func=None):
    """
    Normalisasi setiap blok bulan di sheet Excel:
    - Setiap blok bulan minimal 100 baris.
    - Jika kurang, tambahkan baris baru.
    - Copy style dan formula dari baris terakhir yang memiliki value di kolom B.
    - Setelah setiap blok selesai → renumbering blocks agar update.
    - Print log proses untuk debugging.

    Args:
        sheet: openpyxl worksheet object
        reference_col: kolom yang dijadikan acuan (default 2 = kolom B)
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
    'AOD': f'=IF(BS{{row}}="Stevedore"{sep}10000{sep}IF(H{{row}}="BoCT"{sep}40000{sep}IF(H{{row}}="SMD Anc"{sep}25000{sep}IF(H{{row}}="GPK Port"{sep}10000{sep}IF(H{{row}}="Bunyut"{sep}25000{sep}IF(H{{row}}="Jorong"{sep}7000{sep}IF(H{{row}}="JBG Anc"{sep}10000{sep}0)))))))',
    'AOE': '=(BJ{row}/AOD{row})*24',
    'AOF': '=(AOC{row}-AOE{row})/24',
    'AOH': '=AOF{row}*AOG{row}',
    'AOI': '=AOF{row}*-1',
    'AOJ': '=AOG{row}/2',
    'AOK': '=AOH{row}/2'
}

def delete_or_clear_plan_rows(sheet_b, column_mapping, selected_month):
    """
    Hapus baris dengan Status == 'Plan' hanya pada blok bulan terpilih,
    sedangkan pada bulan lain hanya clear isi value dari kolom C sampai AOT.
    - Mendeteksi header blok bulan pada kolom B (singkatan bulan: Jan, Feb, ...).
    - Setelah menemukan header -> turun 2 baris untuk masuk ke baris pertama blok.
    - Periksa 100 baris (atau sampai akhir sheet), lakukan delete/clear sesuai aturan.
    ➕ Menambahkan log dengan emoji untuk melacak proses.
    """

    # index kolom
    status_col = column_index_from_string(column_mapping['Status'])
    month_header_col = column_index_from_string('B')  # kolom B berisi header bulan
    start_clear_col = column_index_from_string("C")
    end_clear_col = column_index_from_string("AOT")

    # normalisasi selected_month => singkatan (3-letter, e.g. 'Jan')
    if isinstance(selected_month, int):
        try:
            selected_abbrev = datetime.date(2025, selected_month, 1).strftime('%b')
        except Exception:
            selected_abbrev = str(selected_month)[:3]
    else:
        sm = str(selected_month).strip()
        # mencoba parsing full month name ("January") -> "Jan"
        try:
            selected_abbrev = datetime.datetime.strptime(sm, '%B').strftime('%b')
        except Exception:
            # jika sudah singkatan atau lain -> ambil 3 huruf pertama
            selected_abbrev = sm[:3]

    selected_abbrev = selected_abbrev.lower()

    # siapkan daftar header bulan yang ada di sheet (baris dimana kolom B berisi bulan)
    header_rows = []
    for r in range(1, sheet_b.max_row + 1):
        cell_val = sheet_b.cell(row=r, column=month_header_col).value
        if not cell_val:
            continue
        if isinstance(cell_val, str):
            txt = cell_val.strip().lower()
            # cek apakah mulai dengan 3-letter month abbreviation
            if txt[:3] in {'jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'}:
                header_rows.append((r, txt))  # simpan tuple (baris, teks)

    if not header_rows:
        print("⚠️ Tidak ditemukan blok bulan di kolom B.")
        return

    # proses header dari bawah ke atas supaya penghapusan row tidak merusak indeks header yang belum diproses
    header_rows.sort(key=lambda x: x[0], reverse=True)

    for header_row, header_txt in header_rows:
        # mulai data setelah header: turun 2 baris sesuai instruksi
        data_start = header_row + 2
        data_end = min(sheet_b.max_row, data_start + 100 - 1)  # 100 baris ke bawah (inklusive)

        # cek apakah header ini adalah bulan terpilih (bandingkan 3-letter)
        is_selected_month = header_txt.startswith(selected_abbrev[:3].lower())
        print(f"\n📅 Processing month block: {header_txt.title()} "
              f"({'TARGET' if is_selected_month else 'other'})")

        rows_to_delete = []
        # iterasi di block (naik) untuk mengecek status
        for r in range(data_start, data_end + 1):
            status_val = sheet_b.cell(row=r, column=status_col).value
            if status_val is None:
                continue
            if str(status_val).strip().lower() == "plan":
                if is_selected_month:
                    # tandai untuk dihapus (hapusnya nanti dibalik urutan)
                    rows_to_delete.append(r)
                    print(f"   🗑️ Delete row {r} (Status=Plan, {header_txt.title()})")
                else:
                    # Hanya clear nilai dari kolom C..AOT
                    for c in range(start_clear_col, end_clear_col + 1):
                        sheet_b.cell(row=r, column=c).value = None
                    print(f"   🧹 Clear row {r} (Status=Plan, {header_txt.title()})")

        # lakukan penghapusan baris (jika ada), lakukan dari bawah ke atas
        if rows_to_delete:
            for rr in reversed(rows_to_delete):
                sheet_b.delete_rows(rr, 1)

    print("✅ delete_or_clear_plan_rows selesai.")

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