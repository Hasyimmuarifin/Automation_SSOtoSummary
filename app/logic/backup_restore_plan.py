# backup_restore_plan.py

from openpyxl.comments import Comment

def get_match_score(row_vals, backup_vals, target_col, COL_C, COL_D, COL_E, COL_G):
    """
    Hitung skor kecocokan dengan prioritas dinamis berdasarkan kolom asal komentar.
    row_vals    = (month, company, vessel, enduser)
    backup_vals = (month_bkp, company_bkp, vessel_bkp, enduser_bkp)
    target_col  = kolom tempat komentar berasal
    """
    month_val, company_val, vessel_val, enduser_val = row_vals
    month_bkp, company_bkp, vessel_bkp, enduser_bkp = backup_vals

    # default priority
    priority = {
        "enduser": 8,
        "vessel": 4,
        "month": 2,
        "company": 1,
    }

    # kalau komentar berasal dari kolom Vessel → Vessel lebih penting
    if target_col == COL_E:
        priority["vessel"] = 10
        priority["enduser"] = 5

    # kalau komentar berasal dari kolom EndUser → EndUser lebih penting
    elif target_col == COL_G:
        priority["enduser"] = 10
        priority["vessel"] = 5

    # kalau komentar berasal dari kolom Month → Month paling penting
    elif target_col == COL_C:
        priority["month"] = 10

    # kalau komentar berasal dari kolom Company → Company paling penting
    elif target_col == COL_D:
        priority["company"] = 10

    score = 0
    if month_val == month_bkp:
        score += priority["month"]
    if company_val == company_bkp:
        score += priority["company"]
    if vessel_val == vessel_bkp:
        score += priority["vessel"]
    if enduser_val == enduser_bkp:
        score += priority["enduser"]

    return score

def backup_plan_rows(wb, sheet_b, backup_sheet_name="Backup_Plan"):
    """
    Backup baris dengan status 'Plan' atau 'Complete' (kolom BQ) ke sheet sementara.
    - Jika 'Plan' → backup Month, Company, Vessel, End User, serta AKC–AKQ (numeric).
    - Tambahan: Komentar di semua cell kolom C–AOT.
    """
    # hapus sheet backup jika sudah ada
    if backup_sheet_name in wb.sheetnames:
        del wb[backup_sheet_name]
    ws_backup = wb.create_sheet(backup_sheet_name)

    # header utama (data numeric)
    headers = ['Month', 'Company', 'Vessel', 'End User'] + [f"AK{chr(c)}" for c in range(ord('C'), ord('R'))]  # AKC–AKQ
    ws_backup.append(headers)

    # header komentar
    ws_backup.append(["#COMMENT#", "Month", "Company", "Vessel", "EndUser", "Column", "Comment"])

    # kolom index
    COL_STATUS = 69   # BQ
    COL_C = 3         # Month
    COL_D = 4         # Company
    COL_E = 5         # Vessel
    COL_G = 7         # End User
    COL_AKC = 965     # AKC
    COL_AKQ = 979     # AKQ
    COL_MAX = 1086     # AOT

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

            # backup komentar di kolom C–AOT
            for col in range(COL_C, COL_MAX + 1):
                comment_obj = sheet_b.cell(row=row, column=col).comment
                if comment_obj and comment_obj.text:
                    ws_backup.append([
                        "#COMMENT#", month, company, vessel, enduser, col, comment_obj.text
                    ])
                    print(f"[BACKUP] Comment ({row},{col}) → {comment_obj.text}")

def restore_plan_rows(wb, sheet_b, backup_sheet_name="Backup_Plan"):
    """
    Restore data dari sheet Backup_Plan ke sheet ITM Summary (sheet_b).
    Pencocokan berdasarkan 4 kolom: Month, Company, Vessel, End User.
    Jika cocok, isi kembali nilai AKC–AKQ.
    - Restore numeric AKC–AKQ berdasarkan 4 kolom kunci.
    - Restore komentar menggunakan sistem pencocokan berbasis skor.
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

    # pisahkan data utama dan komentar
    comment_rows = []
    # iterasi semua data di backup (mulai dari row 2, karena row 1 adalah header)
    for row in range(2, ws_backup.max_row + 1):
        tag = ws_backup.cell(row=row, column=1).value

        if tag == "#COMMENT#":
            month_bkp = ws_backup.cell(row=row, column=2).value
            company_bkp = ws_backup.cell(row=row, column=3).value
            vessel_bkp = ws_backup.cell(row=row, column=4).value
            enduser_bkp = ws_backup.cell(row=row, column=5).value
            col = ws_backup.cell(row=row, column=6).value
            txt = ws_backup.cell(row=row, column=7).value
            comment_rows.append((month_bkp, company_bkp, vessel_bkp, enduser_bkp, col, txt))
            continue

        if tag is None or tag == "#COMMENTS#":
            continue

        # restore nilai numeric AKC–AKQ
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
            if (
                sheet_b.cell(r, COL_C).value == month_bkp and
                sheet_b.cell(r, COL_D).value == company_bkp and
                sheet_b.cell(r, COL_E).value == vessel_bkp and
                sheet_b.cell(r, COL_G).value == enduser_bkp
            ):
                # cocok → restore nilai numeric AKC–AKQ
                for idx, col in enumerate(range(COL_AKC, COL_AKQ + 1)):
                    val = values_bkp[idx] if idx < len(values_bkp) else None
                    if val is not None:
                        sheet_b.cell(row=r, column=col).value = val

                print(f"[RESTORE] Row {r} diperbarui dari backup (Month={month_bkp}, Company={company_bkp})")
                break  # sudah ketemu baris, tidak perlu cari lagi

    # restore komentar pakai sistem skor (prioritas: End User > Vessel > Month > Company)
    unmatched_comments = []

    for month_bkp, company_bkp, vessel_bkp, enduser_bkp, col, txt in comment_rows:
        best_row = None
        best_score = -1

        print("\n[DEBUG] --- Mencari match untuk komentar ---")
        print(f"Target Backup → Month={month_bkp}, Company={company_bkp}, Vessel={vessel_bkp}, EndUser={enduser_bkp}, Col={col}")

        for r in range(2, sheet_b.max_row + 1):
            month_val   = sheet_b.cell(r, COL_C).value
            company_val = sheet_b.cell(r, COL_D).value
            vessel_val  = sheet_b.cell(r, COL_E).value
            enduser_val = sheet_b.cell(r, COL_G).value

            score = get_match_score(
                (month_val, company_val, vessel_val, enduser_val),
                (month_bkp, company_bkp, vessel_bkp, enduser_bkp),
                col, COL_C, COL_D, COL_E, COL_G
            )

            if score > best_score:
                best_score = score
                best_row = r

        # restore kalau skornya cukup kuat
        if best_score >= 4 and best_row:
            sheet_b.cell(best_row, col).comment = Comment(txt, "BackupRestore")
            print(f"[RESTORE] Comment dikembalikan ke (Row={best_row}, Col={col}) → {txt} | Score={best_score}")
        else:
            unmatched_comments.append({
                "Month": month_bkp,
                "Company": company_bkp,
                "Vessel": vessel_bkp,
                "EndUser": enduser_bkp,
                "Column": col,
                "Text": txt,
                "BestScore": best_score
            })

    # fallback unmatched → buat sheet khusus untuk komentar yang tidak bisa direstore
    if unmatched_comments:
        # filter: buang data kosong / header palsu
        fallback_name = "Unmatched_Comments"
        valid_unmatched = []
        for item in unmatched_comments:
            if (
                not any([item["Month"], item["Company"], item["Vessel"], item["EndUser"], item["Text"]])
                or str(item["Month"]).lower() == "month"
                or str(item["Company"]).lower() == "company"
                or str(item["Vessel"]).lower() == "vessel"
                or str(item["EndUser"]).lower() == "enduser"
                or str(item["Text"]).lower() in ("text", "comment")
            ):
                print("[DEBUG] Skip unmatched_comment karena data kosong/header →", item)
                continue
            valid_unmatched.append(item)

        # hanya buat sheet kalau ada data valid
        if valid_unmatched:
            fallback_name = "Unmatched_Comments"
            if fallback_name in wb.sheetnames:
                ws_fallback = wb[fallback_name]
            else:
                ws_fallback = wb.create_sheet(fallback_name)
                ws_fallback.append(["Month", "Company", "Vessel", "EndUser", "Column", "Text", "BestScore"])

            for item in unmatched_comments:
                ws_fallback.append([
                    item["Month"], item["Company"], item["Vessel"], item["EndUser"],
                    item["Column"], item["Text"], item["BestScore"]
                ])
            print(f"[RESTORE] {len(unmatched_comments)} komentar gagal dipetakan → disimpan di sheet '{fallback_name}'")

    # hapus sheet backup setelah selesai
    del wb[backup_sheet_name]
    print("[RESTORE] Sheet Backup_Plan berhasil dihapus setelah restore.")