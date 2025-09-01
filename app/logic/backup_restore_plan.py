# backup_restore_plan.py

import openpyxl

def backup_plan_rows(wb, sheet_b, backup_sheet_name="Backup_Plan"):
    """
    Backup baris dengan status 'Plan' (kolom BQ) ke sheet sementara.
    Data yang disimpan: kolom D, E, G, serta AKC–AKQ yang berupa angka (bukan formula).
    """
    # hapus sheet backup jika sudah ada
    if backup_sheet_name in wb.sheetnames:
        del wb[backup_sheet_name]
    ws_backup = wb.create_sheet(backup_sheet_name)

    # header
    headers = ['Company', 'Vessel', 'End User'] + [f"AK{chr(c)}" for c in range(ord('C'), ord('R'))]  # AKC–AKQ
    ws_backup.append(headers)

    # kolom index
    COL_STATUS = 69  # BQ → kolom ke-69
    COL_D = 4
    COL_E = 5
    COL_G = 7
    COL_AKC = 965  # AKC
    COL_AKQ = 979  # AKQ

    for row in range(2, sheet_b.max_row + 1):
        status = sheet_b.cell(row=row, column=COL_STATUS).value
        if status == "Plan":
            company = sheet_b.cell(row=row, column=COL_D).value
            vessel = sheet_b.cell(row=row, column=COL_E).value
            enduser = sheet_b.cell(row=row, column=COL_G).value

            # ambil nilai numeric AKC–AKQ
            values = []
            for col in range(COL_AKC, COL_AKQ + 1):
                cell_val = sheet_b.cell(row=row, column=col).value
                values.append(cell_val if isinstance(cell_val, (int, float)) else None)

            ws_backup.append([company, vessel, enduser] + values)
            print(f"[BACKUP] Row {row} → Company={company}, Vessel={vessel}, End User={enduser}, Values={values}")