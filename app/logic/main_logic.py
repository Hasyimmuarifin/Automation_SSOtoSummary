# main_logic.py
from config.column_mapping import column_mapping
from .helpers import month_to_abbreviation, get_header_columns_a, delete_old_plan_rows
from .data_handler import process_data_per_month
from .move_sheet import copy_sheet_full   # ✅ Utility to copy entire sheet
import openpyxl


def run_excel_process(input_file: str, output_file: str) -> str:
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

    # Step 2: Open the workbook for processing
    wb = openpyxl.load_workbook(input_file)
    sheet_b = wb['ITM Summary']     # Destination sheet (processed results)

    # ambil sheet lama (Loading) jika ada
    sheet_loading_old = wb['Loading'] if 'Loading' in wb.sheetnames else None
    sheet_loading_new = wb['Loading2']

    if sheet_loading_old:
        header_columns_old = get_header_columns_a(sheet_loading_old, column_mapping)
        delete_old_plan_rows(sheet_b, sheet_loading_old, header_columns_old, column_mapping)

        # hapus sheet lama
        wb.remove(sheet_loading_old)

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

    # Step 5: Save the result back to the input file (final output)
    wb.save(input_file)
    return f"🎉 Processing complete! Data copied and saved to {input_file}"