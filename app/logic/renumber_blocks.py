def renumber_month_blocks(sheet_b):
    """
    Cari semua blok bulan di Sheet B, lalu perbarui kolom B dengan nomor urut
    Return daftar blok bulan sebagai list of tuples: [(start_row, end_row), ...]
    """

    print("🔢 Renumbering month blocks in Sheet B...")

    month_headers = []  # simpan posisi row header bulan
    month_blocks = []    # hasil akhir: daftar blok bulan
    
    # --- Cari semua header bulan di kolom B ---
    for row in range(1, sheet_b.max_row + 1):
        val = sheet_b.cell(row=row, column=2).value
        if val and isinstance(val, str):
            # cek apakah ini header bulan (contoh: "Jan", "Feb", "September", dll.)
            val_low = val.strip().lower()
            if val_low.startswith(("jan", "feb", "mar", "apr", "may", "jun",
                                   "jul", "aug", "sep", "oct", "nov", "dec")):
                month_headers.append(row)

    # --- Proses setiap blok bulan ---
    for i, hdr_row in enumerate(month_headers):
        cut_start_row = hdr_row + 3   # data mulai 3 baris di bawah header
        if i < len(month_headers) - 1:
            # kalau masih ada bulan berikutnya → batasnya sebelum bulan berikut
            cut_end_row = month_headers[i+1] - 1
        else:
            # kalau bulan terakhir → batasnya sampai data terakhir
            cut_end_row = sheet_b.max_row

        # validasi: pastikan blok tidak kosong
        if cut_start_row > cut_end_row:
            continue

        print(f"   📌 Block {i+1}: Rows {cut_start_row}–{cut_end_row}")

        # --- Cari baris terakhir data valid berdasarkan kolom B ---
        last_data_row = cut_start_row
        while last_data_row <= cut_end_row and sheet_b.cell(row=last_data_row, column=2).value:
            last_data_row += 1
        last_data_row -= 1

        if last_data_row >= cut_start_row:
            # Simpan blok valid
            month_blocks.append((cut_start_row, last_data_row))

            # --- Isi ulang nomor urut di kolom B ---
            for r in range(cut_start_row, last_data_row + 1):
                sheet_b.cell(row=r, column=2).value = f"=ROW()-ROW($B${cut_start_row})+1"

    print("✅ Renumbering done.")
    return month_blocks