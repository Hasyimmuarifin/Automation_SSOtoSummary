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
    app = xw.App(visible=False)
    
    try:
        wb_source = app.books.open(source_file)
        wb_target = app.books.open(target_file)

        # Cari sheet dari target
        sheet_target = None
        for sh in wb_target.sheets:
            if sh.name.strip() == sheet_name:
                sheet_target = sh
                break
        if not sheet_target:
            raise ValueError(f"Sheet '{sheet_name}' tidak ditemukan di {target_file}")

        # # Hapus sheet lama di source kalau ada
        # for sh in wb_source.sheets:
        #     if sh.name.strip() == sheet_name:
        #         sh.delete()
        #         break

        # Copy sheet dari target ke source, beri nama sementara
        sheet_target.api.Copy(Before=wb_source.sheets[0].api)

        # Pastikan nama sheet konsisten
        wb_source.sheets[0].name = new_name

        # Save hasil ke source_file
        wb_source.save()

    finally:
        wb_source.close()
        wb_target.close()
        app.quit()