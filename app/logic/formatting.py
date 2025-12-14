from openpyxl.styles import Font, PatternFill

def apply_font_colors(sheet, start_row, end_row):
    """
    Apply font colors to cells in columns B to L based on the vessel name and status.
    
    - If the status in column BQ is 'Completed', font color is set to black.
    - If the vessel name (column E) starts with 'MV. TBN' or 'BG. TBN', font color is blue.
    - Otherwise, font color is green.

    Parameters:
        sheet (Worksheet): The target worksheet to format.
        start_row (int): The starting row for formatting.
        end_row (int): The ending row for formatting.
    """
    for i in range(start_row, end_row + 1):
        vessel_cell = sheet.cell(row=i, column=5)   # Kolom E: Vessel name
        status_cell = sheet.cell(row=i, column=69)  # Kolom BQ: Status

        status_value = str(status_cell.value).strip() if status_cell.value else ""
        vessel_value = str(vessel_cell.value).strip().upper() if vessel_cell.value else ""

        # Default warna
        font_color = "FF00B050"  # Hijau

        # 🔹 Prioritaskan status selain Plan
        if status_value == "Completed":
            font_color = "FF000000"  # Hitam
        elif status_value == "Loading":
            font_color = "FFFFA500"  # Oranye
        elif status_value == "In Progress":
            font_color = "FF800080"  # Ungu
        elif status_value == "Plan":
            # 🔹 Plan → cek vessel
            if vessel_value.startswith("MV. TBN") or vessel_value.startswith("BG. TBN"):
                font_color = "FF0070C0"  # Biru
            else:
                font_color = "FF00B050"  # Hijau
        else:
            # Kalau status lain yang tidak dikenali → fallback vessel
            if vessel_value.startswith("MV. TBN") or vessel_value.startswith("BG. TBN"):
                font_color = "FF0070C0"  # Biru
            else:
                font_color = "FF00B050"  # Hijau

        # Terapkan ke kolom B–L + BQ
        for j in list(range(2, 13)) + [69]:
            cell = sheet.cell(row=i, column=j)
            if cell.value is not None:
                current_font = cell.font or Font()
                # Create a new Font object preserving the original font attributes
                cell.font = Font(
                    name=current_font.name,
                    size=current_font.size,
                    bold=current_font.bold,
                    italic=current_font.italic,
                    vertAlign=current_font.vertAlign,
                    underline=current_font.underline,
                    strike=current_font.strike,
                    color=font_color
                )

def clear_cell_fill(wb, sheet_b):
    """
    Remove the fill color (make it white/default) in :
    - Column C to G (Month, Company, Vessel, Buyer, End User)
    - Column ANQ
    Only if Status (BQ) is one of:
    complete, loading, in progress, plan
    """
    COL_BQ = 69   # Status
    COL_C = 3
    COL_G = 7
    COL_ANQ = 1057

    count = 0

    for r in range(2, sheet_b.max_row + 1):
        status = str(sheet_b.cell(r, COL_BQ).value or "").strip().lower()

        if status in ("completed", "loading", "in progress", "plan"):
            for col in range(COL_C, COL_G + 1):
                sheet_b.cell(row=r, column=col).fill = PatternFill()  # clear fill
            sheet_b.cell(row=r, column=COL_ANQ).fill = PatternFill()
            count += 1

    print(f"[CLEAR-FILL] {count} rows are colored white (no fill) in C–G and ANQ based on Status (BQ).")