from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string, get_column_letter
from .auto_separator import get_formula_separator

# Marker Baris Total Declaration :
TOTAL_BOCT_MARKER = "Total Coal Loading of BoCT"
TOTAL_MAHAKAM_MARKER = "Total Coal Loading of Mahakam"
TOTAL_MARKER = "Total Coal Loading of ITM (Coal Demand)"

sep = get_formula_separator()

def apply_translated_formulas(sheet, template_row=None, start_row=None, end_row=None, included_columns=None, formulas=None):
    """
    Apply formulas either from a template row or directly from provided formulas.
    
    :param sheet: Worksheet object
    :param template_row: Row number to copy formulas from (if used)
    :param start_row: First row where formulas will be applied
    :param end_row: Last row where formulas will be applied
    :param included_columns: List of column letters where formulas will be applied
    :param formulas: Dict {column_letter: formula_string} to insert formulas directly
                     Example: {'B': '=A{row}+1', 'BJ': '=SUM(C{row}:E{row})'}
    """

    for row in range(start_row, end_row + 1):
        for col in included_columns:
            cell = sheet[f"{col}{row}"]

            if formulas and col in formulas:  
                # Format formula with row number
                formula = formulas[col].replace("{row}", str(row))
                cell.value = formula
            elif template_row:
                # Copy formula from template row if available
                template_cell = sheet[f"{col}{template_row}"]
                if template_cell.value and isinstance(template_cell.value, str) and template_cell.value.startswith("="):
                    cell.value = Translator(template_cell.value, origin=template_cell.coordinate).translate_formula(cell.coordinate)

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
        # 0) Re-apply formula baris data (hingga B kosong)
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