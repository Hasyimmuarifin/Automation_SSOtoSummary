# main_logic.py
from config.column_mapping import column_mapping
from .auto_separator import get_formula_separator
from .move_sheet import copy_sheet_full
from .fill_empty_with_zero import fill_empty_range_with_zero
from .backup_restore_plan import backup_plan_rows, restore_plan_rows, clear_aon_block
from .backup_restore_quality import backup_quality_rows, restore_quality_rows, clear_plan_fill
from .helpers import month_to_abbreviation, get_header_columns_a, normalize_month_block_rows, delete_or_clear_plan_rows, insert_boct_formulas, insert_mahakam_formulas
from .renumber_blocks import renumber_month_blocks
from .data_handler import process_data_per_month
from .apply_color_font import apply_status_font
from .month_block_finder import find_month_block
from .dem_rate_backup import backup_dem_rate, restore_dem_rate

import openpyxl

sep = get_formula_separator()
# 📌 Column mapping formula → pattern (can be expanded as needed)
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

def run_excel_process(input_file: str, output_file: str, month_start: int, month_end: int | None) -> str:
    """
    Main function to process the Excel file.

    Processing steps:
    1. Copy the "Loading" sheet from the output SSO file to the input file (Summary).
    2. Read the header and column mapping from the "Loading" sheet.
    3. Iterate over each row of data by month, ensuring:
       - The month value is valid (integer).
       - The same month is not processed more than once.
    4. Process the data for each month via `process_data_per_month`.
    5. Save the final result to the Excel file.

    Args:
        input_file (str): Path to the source Excel file (summary).
        output_file (str): Path to the destination Excel file (SSO output).
        mohth_start (int): Starting month for processing.
        month_end (int | None): Ending month for processing (optional).

    Returns:
        str: Success message after the process is completed.
    """

    # Step 1: Copy the new “Loading” sheet to the source, name it “Loading2”
    print(f"🟢 Starting Copy Sheet Loading Process...")
    copy_sheet_full(input_file, output_file, sheet_name="Loading", new_name="Loading2")

    # Step 2: Open the input_file (summary) workbook in the ‘Loading’ sheet to fill in the empty cells in the blending column → 0
    print(f"\n🟡 Open Workbook {input_file} to fill empty cell in sheet 'Loading' with 0")
    wb_temp = openpyxl.load_workbook(input_file)
    for target_sheet in ["Loading", "Loading2"]:
        if target_sheet in wb_temp.sheetnames:
            print(f"🟢 Sheet '{target_sheet}' found. Filling empty cells in column range O:AW start from row 2...")
            fill_empty_range_with_zero(wb_temp[target_sheet], check_col="A", start_col="O", end_col="AW", start_row=2)
        else:
            print(f"⚠️ Sheet '{target_sheet}' not found, skipped.")

    # Decide sheet_a
    sheet_a = wb_temp["Loading"] if "Loading" in wb_temp.sheetnames else wb_temp["Loading2"]

    # Load ITM Summary early
    if "ITM Summary" not in wb_temp.sheetnames:
        raise ValueError("Sheet 'ITM Summary' tidak ditemukan!")
    sheet_b = wb_temp["ITM Summary"]

    # Get header columns
    header_columns_a = get_header_columns_a(sheet_a, column_mapping)

    # Get all unique months from sheet_a
    all_months = set()
    for r in range(2, sheet_a.max_row + 1):
        val = sheet_a.cell(row=r, column=header_columns_a["Month"]).value
        if isinstance(val, int):
            all_months.add(val)

    if not all_months:
        raise ValueError("No valid month values found in the Loading/Loading2 sheet!")

    print(f"📌 Found a unique month in the Loading sheet: {sorted(all_months)}")

    # Backup range for each month that appears
    backup_ranges = []  # to be used later during restore

    print(f"💾 Save changes to {input_file}...")
    wb_temp.save(input_file)
    wb_temp.close()

    # Step 3: Open the workbook for processing
    print(f"🟡 Reopening Workbook {input_file} for processing...")
    wb = openpyxl.load_workbook(input_file)

    print("🟡 Open the 'ITM Summary' sheet as the destination sheet for the process results...")
    sheet_b = wb['ITM Summary']

    # Step 4: Take the old sheet (Loading) if available
    if 'Loading' in wb.sheetnames:
        sheet_loading_old = wb['Loading']
        print("🟢 Old sheet 'Loading' found.")
    else:
        sheet_loading_old = None
        print("🟡 Old sheet 'Loading' not found.")

    sheet_loading_new = wb['Loading2']
    print("🟢 Sheet 'Loading2' found & ready to use.")

    if sheet_loading_old:
        print("🟡 Old sheet 'Loading' found, doing backup plan rows...")
        backup_plan_rows(wb, sheet_b)
        print("\n🟡 Doing backup Quality, SOS Month, Remark of Penalty's Cause, Remark Demurrage's Cause, and Remark column in status 'complete/loading/in progress' rows...")
        backup_quality_rows(wb, sheet_b)
        # ✅ Delete sheet if already exist (reset backup)
        for sheet_name in ("dem_plan", "dem_complete"):
            if sheet_name in wb.sheetnames:
                print(f"\n 🗑️ Deleting old sheet: {sheet_name}")
                del wb[sheet_name]
        print("🟡 Backing up demurrage rates for all monthly blocks...")
        for month_value in sorted(all_months):
            start_row, end_row = find_month_block(sheet_b, month_value)

            if not start_row or not end_row:
                print(f"⚠️ Month Block {month_value} Not found in ITM Summary, skipped.")
                continue

            print(f"\n🔒 Backup Demurrage Rate → Month {month_value}: Row {start_row}–{end_row}")
            backup_dem_rate(sheet_b, "dem_plan", "dem_complete", start_row, end_row)

            # Save range for later restore
            backup_ranges.append((month_value, start_row, end_row))
        print("\n🟡 Doing delete or clean plan rows...")
        delete_or_clear_plan_rows(sheet_b, column_mapping, month_start, month_end)

        # --- Step: Normalization month block to 100 ---
        print("🟡 Normalization of the month block to 100 after deleting plan rows...")
        month_blocks = renumber_month_blocks(sheet_b)
        normalize_month_block_rows(sheet_b, month_blocks, reference_col=2, renumber_func=renumber_month_blocks)

        # Delete sheet old 'Loading'
        print("🗑️ Delete the old ‘Loading’ sheet...")
        wb.remove(sheet_loading_old)
    else:
        # If there is no old sheet 'loading', still perform a backup
        print("🟡 The old ‘Loading’ sheet is missing. Continue to back up plan rows....")
        backup_plan_rows(wb, sheet_b)
        print("🟡 Doing backup Quality, SOS Month, Remark of Penalty's Cause, Remark Demurrage's Cause, and Remark column in status 'complete/loading/in progress' rows...")
        backup_quality_rows(wb, sheet_b)
        # ✅ Delete sheet if already exist (reset backup)
        for sheet_name in ("dem_plan", "dem_complete"):
            if sheet_name in wb.sheetnames:
                print(f"\n 🗑️ Deleting old sheet: {sheet_name}")
                del wb[sheet_name]
        print("🟡 Backing up demurrage rates for all monthly blocks...")
        for month_value in sorted(all_months):
            start_row, end_row = find_month_block(sheet_b, month_value)

            if not start_row or not end_row:
                print(f"⚠️ Month Block {month_value} Not found in ITM Summary, skipped.")
                continue

            print(f"🔒 Backup Demurrage Rate → Month {month_value}: Row {start_row}–{end_row}")
            backup_dem_rate(sheet_b, "dem_plan", "dem_complete", start_row, end_row)

            # Save range for later restore
            backup_ranges.append((month_value, start_row, end_row))
        print("🟡 Doing delete or clean plan rows...")
        delete_or_clear_plan_rows(sheet_b, column_mapping, month_start, month_end)

        # --- Step: Normalization month block to 100 ---
        print("🟡 Normalization of the month block to 100 after deleting plan rows...")
        month_blocks = renumber_month_blocks(sheet_b)
        normalize_month_block_rows(sheet_b, month_blocks, reference_col=2, renumber_func=renumber_month_blocks)

    # rename Loading2 → Loading
    print("✒️ Rename sheet 'Loading2' to 'Loading'...")
    sheet_loading_new.title = "Loading"
    sheet_a = wb['Loading']

    # Get column positions based on defined mapping
    print("🟡 Reading column headers from the ‘Loading’ sheet...")
    header_columns_a = get_header_columns_a(sheet_a, column_mapping)

    # A set to track already processed months (to avoid duplicates)
    processed_months = set()

    # Step 5: Iterate through each row in the source sheet
    print("\n🟡 Start iterate every row in sheet 'Loading'...")
    for row in range(2, sheet_a.max_row + 1):  # Start from row 2 (skip header)
        month_value = sheet_a.cell(row=row, column=header_columns_a['Month']).value

        # Validate the month value: must be an integer and not already processed
        if not isinstance(month_value, int):
            print(f"[WARNING] Row {row}: Nilai bulan tidak valid ({month_value}), dilewati.")
            continue
        if month_value in processed_months:
            print(f"🟡 Row {row}: Bulan {month_value} sudah diproses, dilewati.")
            continue

        # Mark this month as processed
        processed_months.add(month_value)

        # Convert numeric month to abbreviation (e.g., 1 -> Jan)
        month_abbreviation = month_to_abbreviation(month_value)
        print(f"🟡 Row {row}: Memproses data bulan {month_value} ({month_abbreviation})...")

        # Step: Process data for this month
        process_data_per_month(
            sheet_a, sheet_b, month_value,
            month_abbreviation, header_columns_a, column_mapping
        )

        # After the monthly process, perform renumbering to refresh month_block and add formulas.
        print("🟢 Refresh Month_Block after all Process & Normalization...")
        month_blocks = renumber_month_blocks(sheet_b)
        print("✒️ Adding the BoCT formula (AKC & AKK) for each monthly block...")
        insert_boct_formulas(sheet_b, month_blocks, reference_col=2, loadport_col="H")
        print("✒️ Adding the Mahakam formula (AKC & AKK) for each monthly block...")
        insert_mahakam_formulas(sheet_b, month_blocks, reference_col=2, loadport_col="H")
        # renumber_month_blocks(sheet_b)
        print(f"🟢 Data in Month {month_value} ({month_abbreviation}) Finished Processing.")

    print("🟡 Restore plan rows after processing all months...")
    clear_aon_block(sheet_b)
    restore_plan_rows(wb, sheet_b)
    restore_quality_rows(wb, sheet_b)
    print("🟢 Restoring demurrage rates for all monthly blocks...")
    for month_value, start_row, end_row in backup_ranges:
        start_row, end_row = find_month_block(sheet_b, month_value)
        if not start_row or not end_row:
            print(f"⚠️ Month Block {month_value} Not found in ITM Summary during restore, skipped.")
            continue
        print(f"🔄 Restore Dem Rate → Month {month_value}: Row {start_row}–{end_row}")
        restore_dem_rate(sheet_b, "dem_plan", "dem_complete", start_row, end_row)
    clear_plan_fill(wb, sheet_b)
    apply_status_font(sheet_b)

    # 🗑️ Cleanup: delete backup sheet after restore is complete
    print("\n=== 🧹 CLEANUP BACKUP SHEETS ===")
    for sheet_name in ("dem_plan", "dem_complete"):
        if sheet_name in wb.sheetnames:
            print(f"🗑️ Deleting sheet: {sheet_name}")
            del wb[sheet_name]

    print("✅ Restore finished & backup sheet was Deleted\n")

    # Step 5: Save the result back to the input file (final output)
    print(f"💾 Save the final result to a file {input_file}...")
    wb.save(input_file)
    return f" ✅ Automation Complete! Data Copied and Saved to 💾 {input_file}"