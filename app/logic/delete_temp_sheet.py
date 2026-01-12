def delete_backup_sheets(wb, sheet_names=("FPG_plan", "FPG_comp", "dem_plan", "dem_complete")):
    """
    Delete temporary backup sheets after restore is complete.
    """
    print("\n === 🧹 CLEANUP BACKUP SHEETS ===")
    for sheet_name in sheet_names:
        if sheet_name in wb.sheetnames:
            print(f"🗑️ Deleting sheet: {sheet_name}")
            del wb[sheet_name]