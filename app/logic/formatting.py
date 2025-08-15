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
        vessel_cell = sheet.cell(row=i, column=5)   # Column E: Vessel name
        status_cell = sheet.cell(row=i, column=69)  # Column BQ: Status
        
        # Determine the font color based on the cell values
        if str(status_cell.value).strip() == "Completed":
            font_color = "000000"  # Black
        elif str(status_cell.value).strip() == "Loading":
            font_color = "FFA500"  # Orange 
        elif str(status_cell.value).strip() == "In Progress":
            font_color = "800080"  # Purple
        else:
            vessel_value = str(vessel_cell.value).strip().upper()
            if vessel_value.startswith("MV. TBN") or vessel_value.startswith("BG. TBN"):
                font_color = "0070C0"  # Blue
            else:
                font_color = "00B050"  # Green


        # Apply the font color to columns B to L (columns 2 to 12)
        for j in range(2, 13):
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