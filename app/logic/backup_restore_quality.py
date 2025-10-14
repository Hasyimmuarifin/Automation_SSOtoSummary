# backup_restore_quality.py

def backup_quality_rows(wb, sheet_b, backup_sheet_name="backup_complete_quality"):
    """
    Backup baris dengan status 'Completed' atau 'Loading' atau 'In Progress'
    - Simpan kolom: Month (C), Company (D), Vessel (E), End User (G)
    - Simpan juga nilai BU–CC, AON, AOP, AOR, AOT
    """
    # hapus sheet lama kalau ada
    if backup_sheet_name in wb.sheetnames:
        del wb[backup_sheet_name]
    ws_backup = wb.create_sheet(backup_sheet_name)

    # header
    headers = (
        ["Month", "Company", "Vessel", "End User"]
        + [f"Col_{col}" for col in range(73, 83)]  # BU–CC (col 73–82)
        + ["AON", "AOP", "AOR", "AOT"]
    )
    ws_backup.append(headers)

    # kolom index
    COL_STATUS = 69   # BQ
    COL_C = 3         # Month
    COL_D = 4         # Company
    COL_E = 5         # Vessel
    COL_G = 7         # End User
    COL_BU = 73
    COL_CC = 82
    COL_AON = 1041
    COL_AOP = 1043
    COL_AOR = 1045
    COL_AOT = 1047

    for row in range(2, sheet_b.max_row + 1):
        status = sheet_b.cell(row=row, column=COL_STATUS).value
        if status in ("Completed", "Loading", "In Progress"):
            month   = sheet_b.cell(row=row, column=COL_C).value
            company = sheet_b.cell(row=row, column=COL_D).value
            vessel  = sheet_b.cell(row=row, column=COL_E).value
            enduser = sheet_b.cell(row=row, column=COL_G).value

            # ambil nilai BU–CC
            values = []
            for col in range(COL_BU, COL_CC + 1):
                cell_val = sheet_b.cell(row=row, column=col).value
                values.append(cell_val if isinstance(cell_val, (int, float)) else None)

            # ambil tambahan AON, AOP, AOR, AOT
            for col in (COL_AON, COL_AOP, COL_AOR, COL_AOT):
                cell_val = sheet_b.cell(row=row, column=col).value
                values.append(cell_val if isinstance(cell_val, (int, float)) else None)

            row_data = [month, company, vessel, enduser] + values
            ws_backup.append(row_data)

            print(f"[BACKUP-QUALITY] Row {row} → {row_data}")


def restore_quality_rows(wb, sheet_b, backup_sheet_name="backup_complete_quality"):
    """
    Restore data dari sheet backup_complete_quality ke sheet ITM Summary (sheet_b).
    Matching berdasarkan 4 kolom: Month, Company, Vessel, End User.
    """
    if backup_sheet_name not in wb.sheetnames:
        print("[RESTORE-QUALITY] Tidak ada sheet backup_complete_quality. Restore dibatalkan.")
        return

    ws_backup = wb[backup_sheet_name]

    # index kolom pada sheet ITM Summary
    COL_C = 3         # Month
    COL_D = 4         # Company
    COL_E = 5         # Vessel
    COL_G = 7         # End User
    COL_BU = 73
    COL_CC = 82
    COL_AON = 1041
    COL_AOP = 1043
    COL_AOR = 1045
    COL_AOT = 1047

    # iterasi semua baris backup
    for row in range(2, ws_backup.max_row + 1):
        month_bkp   = ws_backup.cell(row=row, column=1).value
        company_bkp = ws_backup.cell(row=row, column=2).value
        vessel_bkp  = ws_backup.cell(row=row, column=3).value
        enduser_bkp = ws_backup.cell(row=row, column=4).value

        values_bkp = [
            ws_backup.cell(row=row, column=col).value
            for col in range(5, ws_backup.max_column + 1)
        ]

        # cari baris cocok di sheet ITM Summary
        for r in range(2, sheet_b.max_row + 1):
            if (
                sheet_b.cell(r, COL_C).value == month_bkp and
                sheet_b.cell(r, COL_D).value == company_bkp and
                sheet_b.cell(r, COL_E).value == vessel_bkp and
                sheet_b.cell(r, COL_G).value == enduser_bkp
            ):
                # restore BU–CC
                for idx, col in enumerate(range(COL_BU, COL_CC + 1)):
                    val = values_bkp[idx] if idx < len(values_bkp) else None
                    if val is not None:
                        sheet_b.cell(row=r, column=col).value = val

                # restore tambahan AON–AOT (4 kolom setelah BU–CC)
                extra_cols = [COL_AON, COL_AOP, COL_AOR, COL_AOT]
                for i, col in enumerate(extra_cols, start=(COL_CC - COL_BU + 1)):
                    val = values_bkp[i] if i < len(values_bkp) else None
                    if val is not None:
                        sheet_b.cell(row=r, column=col).value = val

                print(f"[RESTORE-QUALITY] Row {r} diperbarui (Month={month_bkp}, Company={company_bkp})")
                break  # sudah ketemu
    # hapus sheet backup setelah selesai
    del wb[backup_sheet_name]
    print("[RESTORE-QUALITY] Sheet backup_complete_quality berhasil dihapus setelah restore.")