import xlwings as xw

def copy_sheet_full(source_file, target_file, sheet_name="Loading", new_name="Loading2"):
    """
    Copy 1 sheet from the source workbook to the target workbook while maintaining the format, charts, and layout.

    Args:
        source_file (str): source Excel file path (.xlsx, .xlsm) / Summary file
        target_file (str): Excel destination file path (SSO Output)
        sheet_name (str): name of the sheet to be copied (“Loading”)
        new_name (str): name of new loading sheet (“Loading2”)
    """
    # Run Excell App (Invisible in the Background)
    print(f"[START] Copy Sheet Full Process ...")
    print(f"[INFO] Open Excel App (visible=False)...")
    app = xw.App(visible=False)
    
    try:
        print(f"[INFO] Open workbook source: {source_file}")
        wb_source = app.books.open(source_file)

        print(f"[INFO] Open workbook target: {target_file}")
        wb_target = app.books.open(target_file)

        # Searching for the target sheet
        print(f"[INFO] Searching sheet '{sheet_name}' in the workbook target...")
        sheet_target = None
        for sh in wb_target.sheets:
            print(f"   - Found sheet: {sh.name}")
            if sh.name.strip() == sheet_name:
                sheet_target = sh
                break
        if not sheet_target:
            raise ValueError(f"Sheet '{sheet_name}' not found in the {target_file}")
        else:
            print(f"[OK] Sheet '{sheet_name}' Found.")

        # Copy sheet from target to source, give temporary name
        print(f"[INFO] Copy sheet '{sheet_name}' to workbook source...")
        sheet_target.api.Copy(Before=wb_source.sheets[0].api)

        # Ensure that the sheet names are consistent
        print(f"[INFO] Change sheet name result of copy to '{new_name}'...")
        wb_source.sheets[0].name = new_name

        # Save the result to source_file
        print(f"[INFO] Saving Changes to {source_file}...")
        wb_source.save()

        print("[SUCCESS] The copying process is complete..")

    finally:
        print("[INFO] Close workbook and Excel App...")
        wb_source.close()
        wb_target.close()
        app.quit()
        print("[OK] All resources have been closed.")