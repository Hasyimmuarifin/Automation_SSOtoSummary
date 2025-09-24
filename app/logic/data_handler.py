import datetime
from copy import copy
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.styles import PatternFill
from .sorter import sort_data_rows
from .formula import apply_translated_formulas
from .formatting import apply_font_colors
from .zero_handler import replace_zeros_with_none_in_sheet
from .auto_separator import get_formula_separator


def process_data_per_month(sheet_a, sheet_b, month_value, month_abbreviation, header_columns_a, column_mapping):
    """
    Process & transfer monthly data (no-duplicate). 
    - If (Company, Vessel name, End user) already exists in the month block of 'ITM Summary',
      the row is refreshed (non-key columns cleared and refilled with latest values).
    - New combos are appended.
    - Styles for N–BI are preserved; for updated rows, only styles are restored (not old values).
    """

    print(f"⏳ Processing month {month_abbreviation.upper()}...")

    # --- Locate the start row (month block) in Sheet B ---
    cut_start_row = None
    print(f"🔎 Looking for month: {month_abbreviation}")
    for row in range(1, sheet_b.max_row + 1):
        cell_val = sheet_b.cell(row=row, column=2).value
        if cell_val:
            print(f"Row {row}, Col B = {cell_val}")  # 👈 cek isi nyata di Excel
        if cell_val and isinstance(cell_val, str) and cell_val.strip().lower().startswith(month_abbreviation.lower()):
            cut_start_row = row + 3  # data starts 3 rows below header
            break
    if cut_start_row is None:
        print(f"❌ Month {month_abbreviation.upper()} not found in Sheet B.")
        return
    else:
        print(f"✅ Found {month_abbreviation.upper()} starting at row {cut_start_row}")

    # --- Find the end row of the month block (stop when Column C empty) ---
    cut_end_row = cut_start_row
    while cut_end_row <= sheet_b.max_row :
        val_b = sheet_b.cell(row=cut_end_row, column=3).value
        # berhenti kalau benar-benar kosong (None atau string kosong)
        if val_b is None or str(val_b).strip() == "":
            break
        cut_end_row += 1
    cut_end_row -= 1

    # --- Key columns / extra columns & indexes ---
    n_col  = column_index_from_string('N')
    bi_col = column_index_from_string('BI')
    bl_col = column_index_from_string('BL')
    bm_col = column_index_from_string('BM')
    bs_col = column_index_from_string('BS')
    akc_col = column_index_from_string('AKC')
    aki_col = column_index_from_string('AKI')
    akk_col = column_index_from_string('AKK')
    akq_col = column_index_from_string('AKQ')
    extra_cols = [bl_col, bm_col, bs_col]

    # --- Backup values, fills, fonts for the whole block (needed for style restore) ---
    cut_data_dict = {}
    for row in sheet_b.iter_rows(min_row=cut_start_row, max_row=cut_end_row,
                                 min_col=n_col, max_col=bs_col):
        row_idx = row[0].row
        vessel_name = sheet_b.cell(row=row_idx, column=5).value  # Column E
        values_and_styles = {}
        for cell in row:
            col_idx = cell.column
            if n_col <= col_idx <= bi_col:
                values_and_styles[col_idx] = (
                    cell.value,
                    copy(cell.fill),
                    copy(cell.font)
                )
            elif col_idx in extra_cols:
                values_and_styles[col_idx] = (cell.value, None, None)  # value only
        cut_data_dict[vessel_name] = values_and_styles

    # --- Backup values AKK–AKQ ---
    akk_data_dict = {}
    for row in sheet_b.iter_rows(min_row=cut_start_row, max_row=cut_end_row,
                                min_col=akk_col, max_col=akq_col):
        row_idx = row[0].row
        vessel_name = sheet_b.cell(row=row_idx, column=5).value  # Column E (unique key)
        values_dict = {}
        for cell in row:
            values_dict[cell.column] = cell.value
        akk_data_dict[vessel_name] = values_dict

    # --- Backup values AKC–AKI ---
    akc_data_dict = {}
    for row in sheet_b.iter_rows(min_row=cut_start_row, max_row=cut_end_row,
                                min_col=akc_col, max_col=aki_col):
        row_idx = row[0].row
        vessel_name = sheet_b.cell(row=row_idx, column=5).value  # Column E (unique key)
        values_dict = {}
        for cell in row:
            values_dict[cell.column] = cell.value
        akc_data_dict[vessel_name] = values_dict

    # --- Clear old block (only N–BI values + BL/BM/BS values), (AKC-AKQ values) ---
    for row in sheet_b.iter_rows(min_row=cut_start_row, max_row=cut_end_row,
                                 min_col=n_col, max_col=bi_col):
        for cell in row:
            cell.value = None
            cell.fill = PatternFill()
    for col in extra_cols:
        for row in sheet_b.iter_rows(min_row=cut_start_row, max_row=cut_end_row,
                                     min_col=col, max_col=col):
            for cell in row:
                cell.value = None

    for row in sheet_b.iter_rows(min_row=cut_start_row, max_row=cut_end_row,
                                min_col=akc_col, max_col=aki_col):
        for cell in row:
            cell.value = None
    for row in sheet_b.iter_rows(min_row=cut_start_row, max_row=cut_end_row,
                                min_col=akk_col, max_col=akq_col):
        for cell in row:
            cell.value = None

    # === PREP: matching helpers ===
    # Key fields MUST match the names in column_mapping exactly
    key_field_names = {'Company', 'Vessel name', 'End user'}
    key_cols_b = {
        'Company':     column_index_from_string(column_mapping['Company']),
        'Vessel name': column_index_from_string(column_mapping['Vessel name']),
        'End user':    column_index_from_string(column_mapping['End user']),
    }

    def get_keys_from_sheet(sheet, row_idx):
        return (
            sheet.cell(row=row_idx, column=key_cols_b['Company']).value,
            sheet.cell(row=row_idx, column=key_cols_b['Vessel name']).value,
            sheet.cell(row=row_idx, column=key_cols_b['End user']).value
        )

    def find_matching_row_in_block(keys_tuple):
        comp, ves, eus = keys_tuple
        for r in range(cut_start_row, cut_end_row + 1):
            if get_keys_from_sheet(sheet_b, r) == keys_tuple:
                return r
        return None

    # Track which vessel names are UPDATED so we can avoid restoring old values
    updated_vessel_names = set()

    # --- Copy or Update from Sheet A to Sheet B ---
    # IMPORTANT: keep this as the ORIGINAL to ensure sort range covers whole block
    current_row_b = cut_end_row + 1

    for row in range(2, sheet_a.max_row + 1):
        if sheet_a.cell(row=row, column=header_columns_a['Month']).value != month_value:
            continue

        keys_tuple = (
            sheet_a.cell(row=row, column=header_columns_a['Company']).value,
            sheet_a.cell(row=row, column=header_columns_a['Vessel name']).value,
            sheet_a.cell(row=row, column=header_columns_a['End user']).value
        )

        match_row = find_matching_row_in_block(keys_tuple)

        # helper to set a value with Month formatting
        def _write_value(dest_row, col_name, col_idx_a, col_idx_b):
            val = sheet_a.cell(row=row, column=col_idx_a).value
            if col_name == 'Month':
                try:
                    date_obj = datetime.datetime(2025, int(val), 1)
                    cell_b = sheet_b.cell(row=dest_row, column=col_idx_b)
                    cell_b.value = date_obj
                    cell_b.number_format = '[$-en-US]mmm;@'
                except Exception:
                    sheet_b.cell(row=dest_row, column=col_idx_b).value = val
            else:
                sheet_b.cell(row=dest_row, column=col_idx_b).value = val
                
        # helper untuk track kolom yang tidak boleh dihapus
        protected_columns = set(key_field_names) | {"Status"}

        if match_row:
            # --- UPDATE: clear non-protected columns, then refill with latest values ---
            for col_name, col_letter_b in column_mapping.items():
                if col_name in protected_columns:
                    continue  # keep keys + Status
                col_b = column_index_from_string(col_letter_b)
                sheet_b.cell(row=match_row, column=col_b).value = None

            for col_name, col_letter_b in column_mapping.items():
                col_a = header_columns_a.get(col_name)
                if col_a is None:
                    continue
                col_b = column_index_from_string(col_letter_b)
                _write_value(match_row, col_name, col_a, col_b)

            # mark this vessel as updated (used in restore step)
            updated_vessel_names.add(keys_tuple[1])  # Vessel name (Column E)
        else:
            # --- APPEND: write new row at current_row_b ---
            for col_name, col_letter_b in column_mapping.items():
                col_a = header_columns_a.get(col_name)
                if col_a is None:
                    continue
                col_b = column_index_from_string(col_letter_b)
                _write_value(current_row_b, col_name, col_a, col_b)
            current_row_b += 1

    # 📝 Isi default untuk kolom BQ (Status) jika kosong (SEBELUM sorting)
    print("📝 Updating BQ (Status) if Empty before sorting...")
    for row in range(cut_start_row, current_row_b):
        col_bq_val = str(sheet_b[f"BQ{row}"].value or "").upper().strip()
        match True:
            case _ if col_bq_val == "":
                sheet_b[f"BQ{row}"].value = "Plan"

    # --- Define sorting range (covers entire original block + any appends) ---
    sort_start = cut_start_row
    sort_end = current_row_b - 1

    # --- Extract & sort ---
    data_rows = []
    for row in sheet_b.iter_rows(min_row=sort_start, max_row=sort_end, values_only=False):
        data_rows.append([cell.value for cell in row])
    data_rows_sorted = sort_data_rows(data_rows)

    # --- Overwrite with sorted rows ---
    for i, row_data in enumerate(data_rows_sorted):
        for j, value in enumerate(row_data):
            sheet_b.cell(row=sort_start + i, column=j + 1, value=value)

    # --- Restore styles/values (N–BI, plus extra cols) carefully ---
    # For UPDATED vessels:
    #   - N–BI: restore fill/font ONLY (keep new values)
    #   - BL/BM/BS: skip restoring values (keep new)
    sorted_vessel_names = [row[4] for row in data_rows_sorted]  # Column E
    for i, vessel_name in enumerate(sorted_vessel_names):
        values_and_styles = cut_data_dict.get(vessel_name)
        if not values_and_styles:
            continue

        for col_idx, (val, fill, font) in values_and_styles.items():
            target_cell = sheet_b.cell(row=sort_start + i, column=col_idx)

            if vessel_name in updated_vessel_names:
                # Updated rows:
                if n_col <= col_idx <= bi_col:
                    # keep NEW value, restore only style
                    if fill is not None:
                        target_cell.fill = fill
                    if font is not None:
                        target_cell.font = font
                elif col_idx in extra_cols:
                    # keep NEW value in BL/BM/BS (do nothing)
                    pass
            else:
                # Unchanged rows: restore previous values + styles
                target_cell.value = val
                if n_col <= col_idx <= bi_col:
                    if fill is not None:
                        target_cell.fill = fill
                    if font is not None:
                        target_cell.font = font

    for i, vessel_name in enumerate(sorted_vessel_names):  # hasil sort (E)
        values_dict = akc_data_dict.get(vessel_name)
        if not values_dict:
            continue

        for col_idx, val in values_dict.items():
            target_cell = sheet_b.cell(row=sort_start + i, column=col_idx)
            target_cell.value = val
    for i, vessel_name in enumerate(sorted_vessel_names):  # hasil sort (E)
        values_dict = akk_data_dict.get(vessel_name)
        if not values_dict:
            continue

        for col_idx, val in values_dict.items():
            target_cell = sheet_b.cell(row=sort_start + i, column=col_idx)
            target_cell.value = val

    # --- Update Kolom AOG berdasarkan prefiks di Kolom E ---
    print("📝 Updating AOG column based on Vessel prefixes...")
    for row in range(sort_start, sort_end + 1):
        col_e_val = str(sheet_b[f"E{row}"].value or "").upper().strip()

        match True:
            case _ if col_e_val.startswith("MV"):
                sheet_b[f"AOG{row}"].value = 18000
            case _ if col_e_val.startswith(("BG", "DUMP")):
                sheet_b[f"AOG{row}"].value = 0
            case _ if col_e_val == "":
                sheet_b[f"AOG{row}"].value = None
            case _:
                # Kalau tidak cocok apapun, kosongkan cell
                sheet_b[f"AOG{row}"].value = None

    # 📝 Isi default untuk kolom BQ dan ANR jika kosong
    print("📝 Updating BQ (Status) and ANR (Time) if Empty...")
    for row in range(sort_start, sort_end + 1):
        col_anr_val = str(sheet_b[f"ANR{row}"].value or "").upper().strip()

        # Kolom ANR (Time) default 12
        match True:
            case _ if col_anr_val == "":
                sheet_b[f"ANR{row}"].value = 12

    sep = get_formula_separator()
    # --- Formulas, font colors, zero cleanup ---
    apply_translated_formulas(
        sheet_b,
        start_row=sort_start,
        end_row=sort_end,
        included_columns=['B', 'BJ', 'BO', 'AKK', 'ANO', 'ANQ', 'ANS', 'ANT', 'ANU', 'ANX', 'AOA', 'AOB', 'AOC', 'AOD', 'AOE', 'AOF', 'AOH', 'AOI', 'AOJ', 'AOK'],
        formulas={
            # 'B': f"=ROW()-ROW($B${sort_start})+1", # nomor urut otomatis
            'BJ': '=IFERROR(SUM(N{row}:BI{row}),"NULL")',
            'BO': '=(SUMIF($N$892:$BI$892,D{row},N{row}:BI{row}))/BJ{row}',
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
    )

    # --- Tambahan untuk membuat formula dinamis FC Quality Master ---
    def build_fc_formula(col_letter, row, sheet_b, sep):
        """
        Build Excel formula untuk kolom tertentu (col_letter) dan baris row,
        berdasarkan month_val (Jan, Feb, dst).
        """
        month_row_map = {
            "Jan": 6, "Feb": 57, "Mar": 108, "Apr": 159,
            "May": 209, "Jun": 261, "Jul": 312, "Aug": 363,
            "Sep": 414, "Oct": 465, "Nov": 516, "Dec": 567,
        }
        # ambil nilai bulan di kolom C untuk baris tersebut
        month_val = sheet_b.cell(row=row, column=3).value
        if month_val:
            try:
                month_val_str = month_val.strftime("%b")  # "Jan", "Feb", ... "Dec"
            except AttributeError:
                month_val_str = str(month_val).strip()[:3]  # fallback kalau bukan datetime
        else:
            month_val_str = "Jul"

        base_row = month_row_map.get(month_val_str, 363)  # default ke Jul bila tidak dikenali

        return (
            f"=IFERROR("
            f"INDEX(OFFSET('FC Quality Master'!$D${base_row}{sep}0{sep}0{sep}49{sep}13)"
            f"{sep}MATCH({col_letter}$893{sep}"
            f"OFFSET('FC Quality Master'!$D${base_row}{sep}0{sep}0{sep}49{sep}1){sep}0)"
            f"{sep}MATCH({col_letter}$892{sep}"
            f"OFFSET('FC Quality Master'!$D${base_row}{sep}0{sep}0{sep}1{sep}13){sep}0))"
            f"{sep}\"NULL\")"
        )

    # --- APPLY FORMULAS KE 432 KOLOM ---
    fc_formula_columns = [
        ("CE", "DZ"),
        ("EB", "FW"),
        ("FY", "HT"),
        ("HV", "JQ"),
        ("JS", "LN"),
        ("LP", "NK"),
        ("NM", "PH"),
        ("PJ", "RE"),
        ("RG", "TB"),
    ]

    for start_col, end_col in fc_formula_columns:
        start_idx = column_index_from_string(start_col)
        end_idx   = column_index_from_string(end_col)

        for r in range(sort_start, sort_end + 1):
            for c in range(start_idx, end_idx + 1):
                col_letter = get_column_letter(c)
                formula = build_fc_formula(col_letter, r, sheet_b, sep)
                sheet_b[f"{col_letter}{r}"].value = formula

    # --- Tambahan: APPLY FORMULAS KE 432 KOLOM (TD–AKA) ---
    custom_formula_columns = [
        ("TD", "UY"),
        ("VA", "WV"),
        ("WX", "YS"),
        ("YU", "AAP"),
        ("AAR", "ACM"),
        ("ACO", "AEJ"),
        ("AEL", "AGG"),
        ("AGI", "AID"),
        ("AIF", "AKA"),
    ]

    # Range kolom untuk bagian kiri (ulang tiap blok)
    left_start = column_index_from_string("N")
    left_end   = column_index_from_string("BI")
    left_range = list(range(left_start, left_end + 1))

    # flatten daftar kolom kanan sesuai fc_formula_columns
    right_columns = []
    for start_col, end_col in fc_formula_columns:
        start_idx = column_index_from_string(start_col)
        end_idx   = column_index_from_string(end_col)
        right_columns.extend(range(start_idx, end_idx + 1))

    right_iter = iter(right_columns)  # supaya bisa maju sesuai definisi blok fc

    for start_col, end_col in custom_formula_columns:
        start_idx = column_index_from_string(start_col)
        end_idx   = column_index_from_string(end_col)

        # ulang lagi untuk left tiap blok
        for offset, c in enumerate(range(start_idx, end_idx + 1)):
            col_letter = get_column_letter(c)
            left_col_letter = get_column_letter(left_range[offset])   # N–BI (ulang)
            right_col_letter = get_column_letter(next(right_iter))    # CE–... (maju sesuai fc_formula_columns)

            for r in range(sort_start, sort_end + 1):
                sheet_b[f"{col_letter}{r}"].value = (
                    f'=IFERROR({left_col_letter}{r}*{right_col_letter}{r}/$BJ{r},"NULL")'
                )

    # --- Tambahan: APPLY FORMULAS untuk KOLOM (BU - CC) ---
    # gunakan separator lokal
    # (pakai variabel `sep` yang sudah kamu set beberapa baris di atas dengan get_formula_separator())
    # mapping bulan ke offset
    month_offset_map = {
        1: 4,     # Jan
        2: 130,   # Feb
        3: 256,   # Mar
        4: 382,   # Apr
        5: 509,   # May
        6: 636,   # Jun
        7: 764,   # Jul
        8: 892,   # Aug
        9: 1020,  # Sep
        10: 1148, # Oct
        11: 1276, # Nov
        12: 1404, # Dec
    }
    offset_row = month_offset_map.get(month_value)
    if offset_row:
        sumif_start = column_index_from_string("BU")
        sumif_end   = column_index_from_string("CC")

        # mulai persis dari sort_start (di kode kamu sort_start = header+3, jadi sudah benar)
        for r in range(sort_start, sort_end + 1):  # mulai dari sort_start langsung
            for c in range(sumif_start, sumif_end + 1):
                col_letter = get_column_letter(c)
                # =IFERROR(SUMIF(OFFSET($TD$off;0;0;1;1428); col$off+1; OFFSET($TDrow;0;0;1;1428)) ; "NULL")
                formula = (
                    f'=IFERROR('
                    f'SUMIF(OFFSET($TD${offset_row}{sep}0{sep}0{sep}1{sep}1428)'
                    f'{sep}{col_letter}${offset_row+1}'
                    f'{sep}OFFSET($TD{r}{sep}0{sep}0{sep}1{sep}1428))'
                    f'{sep}\"NULL\")'
                )
                sheet_b[f"{col_letter}{r}"].value = formula

    # --- Apply font coloring rules ---
    apply_font_colors(sheet_b, start_row=sort_start, end_row=sort_end)

    # --- Replace zeros with None (to avoid showing 0s in output) ---
    replace_zeros_with_none_in_sheet(sheet_b, start_row=sort_start, end_row=sort_end)

    print(f"✅ Successfully processed month {month_abbreviation.upper()} from Sheet A → Sheet B.")