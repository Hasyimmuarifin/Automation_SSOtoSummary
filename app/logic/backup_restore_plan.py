# backup_restore_plan.py

import openpyxl

def backup_plan_rows(wb, sheet_b, backup_sheet_name="Backup_Plan"):
    """
    Backup baris dengan status 'Plan' (kolom BQ) ke sheet sementara.
   - Jika 'Plan' → backup Month, Company, Vessel, End User, serta AKC–AKQ (numeric).
    """
    # hapus sheet backup jika sudah ada
    if backup_sheet_name in wb.sheetnames:
        del wb[backup_sheet_name]
    ws_backup = wb.create_sheet(backup_sheet_name)

    # header
    headers = ['Month', 'Company', 'Vessel', 'End User'] + [f"AK{chr(c)}" for c in range(ord('C'), ord('R'))]  # AKC–AKQ
    ws_backup.append(headers)

    # kolom index
    COL_STATUS = 69   # BQ
    COL_C = 3         # Month
    COL_D = 4         # Company
    COL_E = 5         # Vessel
    COL_G = 7         # End User
    COL_AKC = 965     # AKC
    COL_AKQ = 979     # AKQ

    for row in range(2, sheet_b.max_row + 1):
        status = sheet_b.cell(row=row, column=COL_STATUS).value
        if status == "Plan":
            month = sheet_b.cell(row=row, column=COL_C).value
            company = sheet_b.cell(row=row, column=COL_D).value
            vessel = sheet_b.cell(row=row, column=COL_E).value
            enduser = sheet_b.cell(row=row, column=COL_G).value

            # ambil nilai numeric AKC–AKQ
            values = []
            for col in range(COL_AKC, COL_AKQ + 1):
                cell_val = sheet_b.cell(row=row, column=col).value
                values.append(cell_val if isinstance(cell_val, (int, float)) else None)

            row_data = [month, company, vessel, enduser] + values
            ws_backup.append(row_data)

            # format kolom Month (kolom A pada sheet backup)
            ws_backup.cell(row=ws_backup.max_row, column=1).number_format = "mmm"

            print(f"[BACKUP] Row {row} → {row_data}")

def restore_plan_rows(wb, sheet_b, backup_sheet_name="Backup_Plan"):
    """
    Restore data dari sheet Backup_Plan ke sheet ITM Summary (sheet_b).
    Pencocokan berdasarkan 4 kolom: Month, Company, Vessel, End User.
    Jika cocok, isi kembali nilai AKC–AKQ.
    Setelah restore selesai, sheet Backup_Plan dihapus.
    """

    if backup_sheet_name not in wb.sheetnames:
        print("[RESTORE] Tidak ada sheet Backup_Plan. Restore dibatalkan.")
        return

    ws_backup = wb[backup_sheet_name]

    # index kolom pada sheet ITM Summary
    COL_C = 3         # Month
    COL_D = 4         # Company
    COL_E = 5         # Vessel
    COL_G = 7         # End User
    COL_AKC = 965     # AKC
    COL_AKQ = 979     # AKQ

    # iterasi semua data di backup (mulai dari row 2, karena row 1 adalah header)
    for row in range(2, ws_backup.max_row + 1):
        month_bkp = ws_backup.cell(row=row, column=1).value
        company_bkp = ws_backup.cell(row=row, column=2).value
        vessel_bkp = ws_backup.cell(row=row, column=3).value
        enduser_bkp = ws_backup.cell(row=row, column=4).value

        values_bkp = [
            ws_backup.cell(row=row, column=col).value
            for col in range(5, ws_backup.max_column + 1)
        ]

        # cari baris yang cocok di sheet ITM Summary
        for r in range(2, sheet_b.max_row + 1):
            month_val = sheet_b.cell(row=r, column=COL_C).value
            company_val = sheet_b.cell(row=r, column=COL_D).value
            vessel_val = sheet_b.cell(row=r, column=COL_E).value
            enduser_val = sheet_b.cell(row=r, column=COL_G).value

            if (
                month_val == month_bkp
                and company_val == company_bkp
                and vessel_val == vessel_bkp
                and enduser_val == enduser_bkp
            ):
                # cocok → restore nilai numeric AKC–AKQ
                for idx, col in enumerate(range(COL_AKC, COL_AKQ + 1)):
                    val = values_bkp[idx] if idx < len(values_bkp) else None
                    if val is not None:
                        sheet_b.cell(row=r, column=col).value = val

                print(f"[RESTORE] Row {r} diperbarui dari backup (Month={month_bkp}, Company={company_bkp})")
                break  # sudah ketemu baris, tidak perlu cari lagi

    # hapus sheet backup setelah selesai
    del wb[backup_sheet_name]
    print("[RESTORE] Sheet Backup_Plan berhasil dihapus setelah restore.")