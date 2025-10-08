# main_logic.py
from config.column_mapping import column_mapping
from .helpers import month_to_abbreviation, get_header_columns_a, normalize_month_block_rows, delete_or_clear_plan_rows
from .data_handler import process_data_per_month
from .move_sheet import copy_sheet_full   # ✅ Utility to copy entire sheet
from .renumber_blocks import renumber_month_blocks
from .formula import reapply_formulas
from .auto_separator import get_formula_separator
from .backup_restore_plan import backup_plan_rows, restore_plan_rows
from .fill_empty_with_zero import fill_empty_range_with_zero
import openpyxl

sep = get_formula_separator()
# 📌 mapping formula kolom → pattern (bisa diperluas sesuai kebutuhan)
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

def run_excel_process(input_file: str, output_file: str, selected_month: int) -> str:
    """
    Main function to process the Excel file.

    Processing steps:
    1. Copy the "Loading" sheet from the input file to the output file.
    2. Read the header and column mapping from the "Loading" sheet.
    3. Iterate over each row of data by month, ensuring:
       - The month value is valid (integer).
       - The same month is not processed more than once.
    4. Process the data for each month via `process_data_per_month`.
    5. Save the final result to the Excel file.

    Args:
        input_file (str): Path to the source Excel file (input).
        output_file (str): Path to the destination Excel file (output).

    Returns:
        str: Success message after the process is completed.
    """

    # Step 1: Copy sheet "Loading" baru ke source, namanya "Loading2"
    copy_sheet_full(input_file, output_file, sheet_name="Loading", new_name="Loading2")

    # Step 1.5: Buka workbook output_file untuk isi cell kosong → 0
    wb_temp = openpyxl.load_workbook(input_file)
    for target_sheet in ["Loading", "Loading2"]:
        if target_sheet in wb_temp.sheetnames:
            fill_empty_range_with_zero(wb_temp[target_sheet], check_col="A", start_col="O", end_col="AW", start_row=2)
    wb_temp.save(input_file)
    wb_temp.close()

    # Step 2: Open the workbook for processing
    wb = openpyxl.load_workbook(input_file)
    sheet_b = wb['ITM Summary']     # Destination sheet (processed results)

    # ambil sheet lama (Loading) jika ada
    sheet_loading_old = wb['Loading'] if 'Loading' in wb.sheetnames else None
    sheet_loading_new = wb['Loading2']

    if sheet_loading_old:
        header_columns_old = get_header_columns_a(sheet_loading_old, column_mapping)

        backup_plan_rows(wb, sheet_b)
        delete_or_clear_plan_rows(sheet_b, column_mapping, selected_month)

        # --- Step: Normalisasi blok bulan setelah delete plan rows ---
        month_blocks = renumber_month_blocks(sheet_b)
        normalize_month_block_rows(sheet_b, month_blocks, reference_col=2, renumber_func=renumber_month_blocks)
        # reapply_formulas(sheet_b,month_blocks, formulas)

        # hapus sheet lama
        wb.remove(sheet_loading_old)
    else:
        # Jika tidak ada sheet loading lama, tetap lakukan backup
        backup_plan_rows(wb, sheet_b)
        delete_or_clear_plan_rows(sheet_b, column_mapping, selected_month)

        # --- Step: Normalisasi blok bulan setelah delete plan rows ---
        month_blocks = renumber_month_blocks(sheet_b)
        normalize_month_block_rows(sheet_b, month_blocks, reference_col=2, renumber_func=renumber_month_blocks)

    # rename Loading2 → Loading
    sheet_loading_new.title = "Loading"
    sheet_a = wb['Loading']

    # Get column positions based on defined mapping
    header_columns_a = get_header_columns_a(sheet_a, column_mapping)

    # A set to track already processed months (to avoid duplicates)
    processed_months = set()

    # Step 3: Iterate through each row in the source sheet
    for row in range(2, sheet_a.max_row + 1):  # Start from row 2 (skip header)
        month_value = sheet_a.cell(row=row, column=header_columns_a['Month']).value

        # Validate the month value: must be an integer and not already processed
        if not isinstance(month_value, int) or month_value in processed_months:
            continue

        # Mark this month as processed
        processed_months.add(month_value)

        # Convert numeric month to abbreviation (e.g., 1 -> Jan)
        month_abbreviation = month_to_abbreviation(month_value)

        # Step 4: Process data for this month
        process_data_per_month(
            sheet_a, sheet_b, month_value,
            month_abbreviation, header_columns_a, column_mapping
        )

        renumber_month_blocks(sheet_b)
    restore_plan_rows(wb, sheet_b)

    # Step 5: Save the result back to the input file (final output)
    wb.save(input_file)
    return f"🎉 Processing complete! Data copied and saved to {input_file}"