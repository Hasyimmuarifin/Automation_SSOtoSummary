# backup_restore_quality.py
from openpyxl.styles import PatternFill

def get_fill_color(cell):
    if cell.fill and cell.fill.fgColor.type == "rgb":
        rgb = cell.fill.fgColor.rgb
        if rgb and rgb not in ("00000000", "FFFFFFFF"):  # abaikan default hitam/putih
            return rgb
    return None

def backup_quality_rows(wb, sheet_b, backup_sheet_name="backup_complete_quality"):
    """
    Backup baris dengan status 'Completed' atau 'Loading' atau 'In Progress'
    - Simpan kolom: Month (C), Company (D), Vessel (E), End User (G)
    - Simpan juga nilai BU–CC, AON, AOP, AOR, AOT beserta warna fill cell
    """
    # hapus sheet lama kalau ada
    if backup_sheet_name in wb.sheetnames:
        del wb[backup_sheet_name]
    ws_backup = wb.create_sheet(backup_sheet_name)

    # header
    headers = (
        ["Month", "Company", "Vessel", "End User"]
        + [f"Col_{col}_Val" for col in range(73, 83)]  # BU–CC (value)
        + [f"Col_{col}_Fill" for col in range(73, 83)]  # BU–CC (fill)
        + ["AON_Val", "AOP_Val", "AOR_Val", "AOT_Val", "AON_Fill", "AOP_Fill", "AOR_Fill", "AOT_Fill"]
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
                cell = sheet_b.cell(row=row, column=col)
                val = cell.value if isinstance(cell.value, (int, float)) else None
                fill = get_fill_color(cell)
                values.append(val)
            for col in range(COL_BU, COL_CC + 1):
                cell = sheet_b.cell(row=row, column=col)
                fill = get_fill_color(cell)
                values.append(fill)

            # ambil tambahan AON, AOP, AOR, AOT
            for col in (COL_AON, COL_AOP, COL_AOR, COL_AOT):
                cell = sheet_b.cell(row=row, column=col)
                values.append(cell.value)

            row_data = [month, company, vessel, enduser] + values
            ws_backup.append(row_data)

            print(f"[BACKUP-QUALITY] Row {row} dibackup.")


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
            for col in range(5, 5 + (COL_CC - COL_BU + 1))
        ]
        fills_bkp = [
            ws_backup.cell(row=row, column=col).value
            for col in range(5 + (COL_CC - COL_BU + 1), 5 + 2*(COL_CC - COL_BU + 1))
        ]

        start_extra = 5 + 2*(COL_CC - COL_BU + 1)
        extra_vals = [
            ws_backup.cell(row=row, column=col).value
            for col in range(start_extra, start_extra + 4)
        ]
        extra_fills = [
            ws_backup.cell(row=row, column=col).value
            for col in range(start_extra + 4, start_extra + 8)
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
                    val = values_bkp[idx]
                    if val is not None:
                        sheet_b.cell(row=r, column=col).value = val
                    if fills_bkp[idx]:
                        sheet_b.cell(row=r, column=col).fill = PatternFill(start_color=fills_bkp[idx], end_color=fills_bkp[idx], fill_type="solid")

                # restore tambahan AON–AOT (4 kolom setelah BU–CC)
                for i, col in enumerate((COL_AON, COL_AOP, COL_AOR, COL_AOT)):
                    val = extra_vals[i]
                    if val is not None:
                        sheet_b.cell(row=r, column=col).value = val
                    if extra_fills[i]:
                        sheet_b.cell(row=r, column=col).fill = PatternFill(start_color=extra_fills[i], end_color=extra_fills[i], fill_type="solid")

                print(f"[RESTORE-QUALITY] Row {r} diperbaruidengan nilai + fill.")
                break  # sudah ketemu
    # hapus sheet backup setelah selesai
    del wb[backup_sheet_name]
    print("[RESTORE-QUALITY] Sheet backup_complete_quality berhasil dihapus setelah restore.")

def clear_plan_fill(wb, sheet_b):
    """
    Menghapus warna fill (membuat putih/default) pada kolom BU–CC
    untuk setiap baris yang memiliki Status == 'Plan' (kolom BQ).
    """
    COL_STATUS = 69   # BQ
    COL_BU = 73
    COL_CC = 82

    count = 0
    for r in range(2, sheet_b.max_row + 1):
        status = str(sheet_b.cell(r, COL_STATUS).value or "").strip().lower()
        if status == "plan":
            for col in range(COL_BU, COL_CC + 1):
                sheet_b.cell(row=r, column=col).fill = PatternFill()  # clear fill
            count += 1

    print(f"[CLEAR-FILL] {count} baris dengan status 'Plan' diwarnai putih (tanpa fill).")