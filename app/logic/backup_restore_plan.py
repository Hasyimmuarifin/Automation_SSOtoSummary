# backup_restore_plan.py

from openpyxl.utils import get_column_letter

def backup_plan_rows(wb, sheet_b, backup_sheet_name="Backup_Plan", debug=True):
    """
    Back up rows with the status ‘Plan’ to a temporary sheet.
    - If 'Plan' → backup Month, Company, Vessel, End User, with AKC–AKQ column (numeric).
    - debug=True → print additional information for AON/AOP/AOR/AOT verification
    """
    # delete the backup sheet if it already exists
    if backup_sheet_name in wb.sheetnames:
        del wb[backup_sheet_name]

    ws_backup = wb.create_sheet(backup_sheet_name)

    # main header (numeric data)
    headers = (
        ['Month', 'Company', 'Vessel', 'Buyer', 'End User'] + [f"AK{chr(c)}" for c in range(ord('C'), ord('R'))] + ["AON", "AOP", "AOR", "AOT"]
    )
    ws_backup.append(headers)

    # index column
    COL_BQ = 69   # Status
    COL_C = 3     # Month
    COL_D = 4     # Company
    COL_E = 5     # Vessel
    COL_F = 6     # Buyer
    COL_G = 7     # End User
    COL_AKC = 965 
    COL_AKQ = 979
    COL_AON = 1080
    COL_AOP = 1082
    COL_AOR = 1084
    COL_AOT = 1086

    # sanity checks (debug)
    if debug:
        print(f"\n   [DEBUG] Sheet 'ITM Summary' max_row={sheet_b.max_row}, max_column={sheet_b.max_column}")

    # expected calculation
    ak_count = COL_AKQ - COL_AKC + 1
    expected_len = 5 + ak_count + 4
    print(f"   [DEBUG] Total expected header length: {expected_len} column.")

    for row in range(2, sheet_b.max_row + 1):
        status = sheet_b.cell(row=row, column=COL_BQ).value
        if status == "Plan":
            month = sheet_b.cell(row=row, column=COL_C).value
            company = sheet_b.cell(row=row, column=COL_D).value
            vessel = sheet_b.cell(row=row, column=COL_E).value
            buyer = sheet_b.cell(row=row, column=COL_F).value
            enduser = sheet_b.cell(row=row, column=COL_G).value

            if debug:
                print(f"\n                     [DEBUG] ---- Processing row {row} ----")
                print(f"   [DEBUG] Status='{status}' Month={month!r}, Company={company!r}, Vessel={vessel!r}, Buyer={buyer!r}, EndUser={enduser!r}")

            # take the numeric value AKC–AKQ
            values = []
            for col in range(COL_AKC, COL_AKQ + 1):
                cell = sheet_b.cell(row=row, column=col).value
                values.append(cell if isinstance(cell, (int, float)) else None)
            if debug:
                first_col_letter = get_column_letter(COL_AKC)
                last_col_letter = get_column_letter(COL_AKQ)
                print(f"   [DEBUG] AK values ({first_col_letter}{row}..{last_col_letter}{row})  |  count={len(values)}  |  sample: {values[:7]} ... {values[-7:]}")

            # take additional AON, AOP, AOR, AOT columns
            aon_values = []
            for col in (COL_AON, COL_AOP, COL_AOR, COL_AOT):
                val = sheet_b.cell(row=row, column=col).value
                values.append(val)
                aon_values.append((col, get_column_letter(col), val))
            if debug:
                for col_idx, col_letter, val in aon_values:
                    print(f"   [DEBUG] AON...AOT block: Col {col_idx} ({col_letter}{row}) = {val!r}")
                # check if all None
                if all(v is None for (_, _, v) in aon_values):
                    print(f"   [WARN] All Value in AON..AOT Column is 'None' in row {row} (check column mapping / does the cell contain a formula without a value).")

            row_data = [month, company, vessel, buyer, enduser] + values

            # debug length check before appending
            if debug:
                print(f"   [DEBUG] row_data length={len(row_data)} expected={expected_len}. row_data head: {row_data[:8]}")

            ws_backup.append(row_data)

            # Month column format (column A on the backup sheet to "mmm")
            ws_backup.cell(row=ws_backup.max_row, column=1).number_format = "mmm"
            print(f"   [BACKUP] Row {row} → appended to Backup_Plan row {ws_backup.max_row}")

def restore_plan_rows(wb, sheet_b, backup_sheet_name="Backup_Plan"):
    """
    Restore data from the Backup_Plan sheet to the ITM Summary sheet (sheet_b).
    Matching based on 4 columns: Month, Company, Vessel, End User.
    If it match, fill in the AKC–AKQ values again.
    - Restore numeric AKC–AKQ based on 4 key column (month, company, vessel, enduser).
    After the restore is complete, the Backup_Plan sheet is deleted.
    """

    if backup_sheet_name not in wb.sheetnames:
        print("   [RESTORE] There is no Backup_Plan sheet. Restore canceled.")
        return

    ws_backup = wb[backup_sheet_name]

    # index column
    COL_C = 3         # Month
    COL_D = 4         # Company
    COL_E = 5         # Vessel
    COL_F = 6         # Buyer
    COL_G = 7         # End User
    COL_AKC = 965
    COL_AKQ = 979
    COL_AON = 1080
    COL_AOP = 1082
    COL_AOR = 1084
    COL_AOT = 1086

    # iterate through all data in the backup (starting from row 2, because row 1 is the header)
    for row in range(2, ws_backup.max_row + 1):
        month_bkp = ws_backup.cell(row=row, column=1).value
        company_bkp = ws_backup.cell(row=row, column=2).value
        vessel_bkp = ws_backup.cell(row=row, column=3).value
        buyer_bkp = ws_backup.cell(row=row, column=4).value
        enduser_bkp = ws_backup.cell(row=row, column=5).value

        values_bkp = [
            ws_backup.cell(row=row, column=col).value
            for col in range(6, ws_backup.max_column + 1)
        ]

        # find the matching row in the ITM Summary sheet
        for r in range(2, sheet_b.max_row + 1):
            month_val = sheet_b.cell(row=r, column=COL_C).value
            company_val = sheet_b.cell(row=r, column=COL_D).value
            vessel_val = sheet_b.cell(row=r, column=COL_E).value
            buyer_val = sheet_b.cell(row=r, column=COL_F).value
            enduser_val = sheet_b.cell(row=r, column=COL_G).value

            if (
                month_val == month_bkp
                and company_val == company_bkp
                and vessel_val == vessel_bkp
                and buyer_val == buyer_bkp
                and enduser_val == enduser_bkp
            ):
                # match → restore numeric value in AKC–AKQ
                for idx, col in enumerate(range(COL_AKC, COL_AKQ + 1)):
                    val = values_bkp[idx] if idx < len(values_bkp) else None
                    if val is not None:
                        sheet_b.cell(row=r, column=col).value = val

                # additional restore AON–AOT column (4 column after AKQ)
                extra_cols = [COL_AON, COL_AOP, COL_AOR, COL_AOT]
                for i, col in enumerate(extra_cols, start=(COL_AKQ - COL_AKC + 1)):
                    val = values_bkp[i] if i < len(values_bkp) else None
                    if val is not None:
                        sheet_b.cell(row=r, column=col).value = val

                print(f"   [RESTORE] Row {r} updated from backup : (Month={month_bkp}, Company={company_bkp})")
                break

    # delete the backup sheet after finishing
    del wb[backup_sheet_name]
    print("   [RESTORE] The Backup_Plan sheet was successfully deleted after restore..")

def clear_aon_block(sheet_b, debug=True):
    """
    Delete the values in columns AON, AOP, AOR, AOT for rows with Status=‘Plan’ (column BQ).
    Called after the monthly process is complete and before restore_plan_rows().
    """
    COL_BQ = 69    # Status
    COL_AON = 1080
    COL_AOP = 1082
    COL_AOR = 1084
    COL_AOT = 1086

    cleared_rows = 0

    for row in range(2, sheet_b.max_row + 1):
        status = sheet_b.cell(row=row, column=COL_BQ).value
        if status == "Plan":
            for col in (COL_AON, COL_AOP, COL_AOR, COL_AOT):
                if sheet_b.cell(row=row, column=col).value is not None:
                    sheet_b.cell(row=row, column=col).value = None
                    if debug:
                        from openpyxl.utils import get_column_letter
                        print(f"   [CLEAR] Row {row}, Col {get_column_letter(col)} → cleared")
                    cleared_rows += 1

    print(f"   [CLEAR] Total {cleared_rows} rows with Status='Plan' cleared in AON–AOT columns.")