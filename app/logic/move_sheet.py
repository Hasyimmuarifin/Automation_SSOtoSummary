import xlwings as xw

def copy_sheet_full(source_file, target_file, sheet_name="Loading", new_name="Loading2"):
    """
    Menyalin 1 sheet dari workbook sumber ke workbook target 
    dengan format, chart, dan layout tetap terjaga.

    Args:
        source_file (str): path file Excel sumber (.xlsx, .xlsm)
        target_file (str): path file Excel tujuan
        sheet_name (str): nama sheet yang akan dicopy ("Loading")
        new_name (str): nama sheet Loading baru ("Loading2")
    """
    # Jalankan Excel (tidak terlihat)
    print(f"[INFO] Membuka Excel App (visible=False)...")
    app = xw.App(visible=False)
    
    try:
        print(f"[INFO] Membuka workbook sumber: {source_file}")
        wb_source = app.books.open(source_file)

        print(f"[INFO] Membuka workbook target: {target_file}")
        wb_target = app.books.open(target_file)

        # Cari sheet dari target
        print(f"[INFO] Mencari sheet '{sheet_name}' di workbook target...")
        sheet_target = None
        for sh in wb_target.sheets:
            print(f"   - Ditemukan sheet: {sh.name}")
            if sh.name.strip() == sheet_name:
                sheet_target = sh
                break
        if not sheet_target:
            raise ValueError(f"Sheet '{sheet_name}' tidak ditemukan di {target_file}")
        else:
            print(f"[OK] Sheet '{sheet_name}' ditemukan.")

        # Copy sheet dari target ke source, beri nama sementara
        print(f"[INFO] Menyalin sheet '{sheet_name}' ke workbook sumber...")
        sheet_target.api.Copy(Before=wb_source.sheets[0].api)

        # Pastikan nama sheet konsisten
        print(f"[INFO] Mengubah nama sheet hasil copy menjadi '{new_name}'...")
        wb_source.sheets[0].name = new_name

        # Save hasil ke source_file
        print(f"[INFO] Menyimpan perubahan ke {source_file}...")
        wb_source.save()

        print("[SUCCESS] Proses penyalinan selesai.")

    finally:
        print("[INFO] Menutup workbook dan Excel App...")
        wb_source.close()
        wb_target.close()
        app.quit()
        print("[OK] Semua resource sudah ditutup.")