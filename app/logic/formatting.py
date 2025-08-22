from openpyxl.styles import Font

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