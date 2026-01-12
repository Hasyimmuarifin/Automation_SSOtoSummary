# backup_restore_quality.py
from openpyxl.styles import PatternFill

def get_fill_color(cell):
    if cell.fill and cell.fill.fgColor.type == "rgb":
        rgb = cell.fill.fgColor.rgb
        if rgb and rgb not in ("00000000", "FFFFFFFF"):  # ignore default black/white
            return rgb
    return None

def backup_complete_quality_rows(wb, sheet_b, backup_sheet_name="backup_complete_quality"):
    """
    Backup row with the status is 'Completed' or 'Loading' or 'In Progress'
       - Keep Value in Column : Month (C), Company (D), Vessel (E), Buyer (F), End User (G)
       - Keep also Value in Column BU–CC with the fill cell color
    """
    # delete the old sheet if there is already exist
    if backup_sheet_name in wb.sheetnames:
        del wb[backup_sheet_name]

    ws_backup = wb.create_sheet(backup_sheet_name)

    # header
    headers = (
        ["Month", "Company", "Vessel", "Buyer", "End User"]
        + [f"{col}_Val" for col in range(73, 82)]  # BU–CC (value)
        + [f"{col}_Fill" for col in range(73, 82)]  # BU–CC (fill)
    )
    ws_backup.append(headers)

    # index column
    COL_BQ = 69       # Status
    COL_C = 3         # Month
    COL_D = 4         # Company
    COL_E = 5         # Vessel
    COL_F = 6         # Buyer
    COL_G = 7         # End User
    COL_BU = 73
    COL_CC = 81

    for row in range(2, sheet_b.max_row + 1):
        status = sheet_b.cell(row=row, column=COL_BQ).value
        if status in ("Completed", "Loading", "In Progress"):
            month   = sheet_b.cell(row=row, column=COL_C).value
            company = sheet_b.cell(row=row, column=COL_D).value
            vessel  = sheet_b.cell(row=row, column=COL_E).value
            buyer  = sheet_b.cell(row=row, column=COL_F).value
            enduser = sheet_b.cell(row=row, column=COL_G).value

            # take the BU–CC value
            values = []
            for col in range(COL_BU, COL_CC + 1):
                cell = sheet_b.cell(row=row, column=col)
                val = cell.value if isinstance(cell.value, (int, float)) else None
                fill = get_fill_color(cell)
                values.append(val)
            for col in range(COL_BU, COL_CC + 1):
                cell = sheet_b.cell(row=row, column=col)
                fill = get_fill_color(cell)
                values.append(fill)

            row_data = [month, company, vessel, buyer, enduser] + values
            ws_backup.append(row_data)

            print(f"   [BACKUP-COMPLETE-QUALITY] Row {row} backed up.")


def restore_complete_quality_rows(wb, sheet_b, backup_sheet_name="backup_complete_quality"):
    """
    Restore data from the backup_complete_quality sheet to the ITM Summary sheet (sheet_b).
    Matching based on 4 columns: Month, Company, Vessel, End User.
    """
    if backup_sheet_name not in wb.sheetnames:
        print("   [RESTORE-QUALITY] There is no backup_complete_quality sheet. Restore canceled.")
        return

    ws_backup = wb[backup_sheet_name]

    # index column
    COL_C = 3         # Month
    COL_D = 4         # Company
    COL_E = 5         # Vessel
    COL_F = 6         # Buyer
    COL_G = 7         # End User
    COL_BU = 73
    COL_CC = 81

    # iterate through all backup rows
    for row in range(2, ws_backup.max_row + 1):
        month_bkp   = ws_backup.cell(row=row, column=1).value
        company_bkp = ws_backup.cell(row=row, column=2).value
        vessel_bkp  = ws_backup.cell(row=row, column=3).value
        buyer_bkp  = ws_backup.cell(row=row, column=4).value
        enduser_bkp = ws_backup.cell(row=row, column=5).value

        values_bkp = [
            ws_backup.cell(row=row, column=col).value
            for col in range(6, 6 + (COL_CC - COL_BU + 1))
        ]
        fills_bkp = [
            ws_backup.cell(row=row, column=col).value
            for col in range(6 + (COL_CC - COL_BU + 1), 6 + 2*(COL_CC - COL_BU + 1))
        ]

        # find matching rows in the ITM Summary sheet
        for r in range(2, sheet_b.max_row + 1):
            if (
                sheet_b.cell(r, COL_C).value == month_bkp and
                sheet_b.cell(r, COL_D).value == company_bkp and
                sheet_b.cell(r, COL_E).value == vessel_bkp and
                sheet_b.cell(r, COL_F).value == buyer_bkp and
                sheet_b.cell(r, COL_G).value == enduser_bkp
            ):
                # restore BU–CC
                for idx, col in enumerate(range(COL_BU, COL_CC + 1)):
                    val = values_bkp[idx]
                    if val is not None:
                        sheet_b.cell(row=r, column=col).value = val
                    if fills_bkp[idx]:
                        sheet_b.cell(row=r, column=col).fill = PatternFill(start_color=fills_bkp[idx], end_color=fills_bkp[idx], fill_type="solid")

                print(f"   [RESTORE-COMPLETE-QUALITY] Row {r} updated with the value + fill.")
                break
    # delete the backup sheet after finishing
    del wb[backup_sheet_name]
    print("   [RESTORE-COMPLETE-QUALITY] Sheet backup_complete_quality successfully deleted after restore.")

def backup_plan_quality_rows(wb, sheet_b, backup_sheet_name="backup_plan_quality"):
    """
    Backup row with the status is 'Plan'
       - Keep Value in Column : Month (C), Company (D), Vessel (E), Buyer (F), End User (G)
       - Keep also Value in Column BU–CC with the fill cell color
    """
    # delete the old sheet if there is already exist
    if backup_sheet_name in wb.sheetnames:
        del wb[backup_sheet_name]

    ws_backup = wb.create_sheet(backup_sheet_name)

    # header
    headers = (
        ["Month", "Company", "Vessel", "Buyer", "End User"]
        + [f"{col}_Val" for col in range(73, 82)]  # BU–CC (value)
        + [f"{col}_Fill" for col in range(73, 82)]  # BU–CC (fill)
    )
    ws_backup.append(headers)

    # index column
    COL_BQ = 69       # Status
    COL_C = 3         # Month
    COL_D = 4         # Company
    COL_E = 5         # Vessel
    COL_F = 6         # Buyer
    COL_G = 7         # End User
    COL_BU = 73
    COL_CC = 81

    for row in range(2, sheet_b.max_row + 1):
        status = sheet_b.cell(row=row, column=COL_BQ).value
        actual_quality = sheet_b.cell(row=row, column=COL_BU).value
        if status == "Plan" and isinstance(actual_quality, str) and actual_quality.startswith("="):
            continue  # skip formula in actual quality

        if status == "Plan":
            month   = sheet_b.cell(row=row, column=COL_C).value
            company = sheet_b.cell(row=row, column=COL_D).value
            vessel  = sheet_b.cell(row=row, column=COL_E).value
            buyer  = sheet_b.cell(row=row, column=COL_F).value
            enduser = sheet_b.cell(row=row, column=COL_G).value

            # take the BU–CC value
            values = []
            for col in range(COL_BU, COL_CC + 1):
                cell = sheet_b.cell(row=row, column=col)
                val = cell.value if isinstance(cell.value, (int, float)) else None
                fill = get_fill_color(cell)
                values.append(val)
            for col in range(COL_BU, COL_CC + 1):
                cell = sheet_b.cell(row=row, column=col)
                fill = get_fill_color(cell)
                values.append(fill)

            row_data = [month, company, vessel, buyer, enduser] + values
            ws_backup.append(row_data)

            print(f"   [BACKUP-PLAN-QUALITY] Row {row} backed up.")


def restore_plan_quality_rows(wb, sheet_b, backup_sheet_name="backup_plan_quality"):
    """
    Restore data from the backup_plan_quality sheet to the ITM Summary sheet (sheet_b).
    Matching based on 4 columns: Month, Company, Vessel, End User.
    """
    if backup_sheet_name not in wb.sheetnames:
        print("   [RESTORE-PLAN-QUALITY] There is no backup_plan_quality sheet. Restore canceled.")
        return

    ws_backup = wb[backup_sheet_name]

    # index column
    COL_C = 3         # Month
    COL_D = 4         # Company
    COL_E = 5         # Vessel
    COL_F = 6         # Buyer
    COL_G = 7         # End User
    COL_BU = 73
    COL_CC = 81

    # iterate through all backup rows
    for row in range(2, ws_backup.max_row + 1):
        month_bkp   = ws_backup.cell(row=row, column=1).value
        company_bkp = ws_backup.cell(row=row, column=2).value
        vessel_bkp  = ws_backup.cell(row=row, column=3).value
        buyer_bkp  = ws_backup.cell(row=row, column=4).value
        enduser_bkp = ws_backup.cell(row=row, column=5).value

        values_bkp = [
            ws_backup.cell(row=row, column=col).value
            for col in range(6, 6 + (COL_CC - COL_BU + 1))
        ]
        fills_bkp = [
            ws_backup.cell(row=row, column=col).value
            for col in range(6 + (COL_CC - COL_BU + 1), 6 + 2*(COL_CC - COL_BU + 1))
        ]

        # find matching rows in the ITM Summary sheet
        for r in range(2, sheet_b.max_row + 1):
            if (
                sheet_b.cell(r, COL_C).value == month_bkp and
                sheet_b.cell(r, COL_D).value == company_bkp and
                sheet_b.cell(r, COL_E).value == vessel_bkp and
                sheet_b.cell(r, COL_F).value == buyer_bkp and
                sheet_b.cell(r, COL_G).value == enduser_bkp
            ):
                # restore BU–CC
                for idx, col in enumerate(range(COL_BU, COL_CC + 1)):
                    val = values_bkp[idx]
                    if val is not None:
                        sheet_b.cell(row=r, column=col).value = val
                    if fills_bkp[idx]:
                        sheet_b.cell(row=r, column=col).fill = PatternFill(start_color=fills_bkp[idx], end_color=fills_bkp[idx], fill_type="solid")

                print(f"   [RESTORE-PLAN-QUALITY] Row {r} updated with the value + fill.")
                break
    # delete the backup sheet after finishing
    del wb[backup_sheet_name]
    print("   [RESTORE-PLAN-QUALITY] Sheet backup_plan_quality successfully deleted after restore.")

def clear_plan_fill(wb, sheet_b):
    """
    Remove the fill color (make it white/default) in the BU–CC column for each row that has Status == ‘Plan’ (column BQ).
    """
    COL_BQ = 69   # Status
    COL_BU = 73
    COL_CC = 81

    count = 0
    for r in range(2, sheet_b.max_row + 1):
        status = str(sheet_b.cell(r, COL_BQ).value or "").strip().lower()
        if status == "plan":
            for col in range(COL_BU, COL_CC + 1):
                sheet_b.cell(row=r, column=col).fill = PatternFill()  # clear fill
            count += 1

    print(f"[CLEAR-FILL] {count}  rows with the status ‘Plan’ are colored white (no fill) in columns BU–CC.")