# dem_rate_backup.py
from openpyxl.styles import Font, Color
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet

# Valid black color
# BLACK = {"000000", "00000000", "FF000000"}

def is_black(color_obj: Color):
    """
    Determining whether a color is black in various Excel formats:
    - RGB / ARGB
    - theme color default ( )
    - indexed black color
    - None value considered to black (Excel default)
    """
    if not color_obj:
        return True  # default Excel font = black

    # --- 1. RGB / ARGB ---
    rgb = color_obj.rgb
    if rgb:
        hex_val = getattr(rgb, "value", rgb)
        if isinstance(hex_val, str) and hex_val.upper() in ("FF000000", "000000", "00000000"):
                return True

    # --- 2. Indexed colors which is considered black ---
    if color_obj.type == "indexed":
        # 8 and 81 are generally black/dark
        if color_obj.indexed in (0, 8, 64, 81):
            return True

    # --- 3. Theme black ---
    if color_obj.type == "theme":
        # theme=1 or theme=0 often turns black
        if color_obj.theme in (0, 1):
            return True

    return False

def get_key(ws, row):
    """Return tuple(Month, Company, Vessel, Buyer, EndUser)"""
    return tuple(str(ws.cell(row=row, column=c).value or "").strip()
                 for c in range(3, 8))

def create_backup_sheet(wb, sheet_name):
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        headers = ["Month", "Company", "Vessel", "Buyer", "End User", "Value", "HexColor"]
        ws.append(headers)
    else:
        ws = wb[sheet_name]

    return ws

def find_row(ws, key):
    """Search for rows in the backup based on the key (columns 1..5)."""
    for r in range(2, ws.max_row + 1):
        row_key = tuple(str(ws.cell(row=r, column=c).value or "").strip()
                        for c in range(1, 6))
        if row_key == key:
            return r
    return None

def backup_dem_rate(ws_summary: Worksheet, plan_sheet_name: str, comp_sheet_name: str, start_row: int, end_row: int):
    """
    Backup DEM Rate with debug print.
    """
    wb = ws_summary.parent

    ws_plan = create_backup_sheet(wb, plan_sheet_name)
    ws_comp = create_backup_sheet(wb, comp_sheet_name)

    if end_row is None:
        end_row = ws_summary.max_row

    try:
        AOG_col = column_index_from_string("AOG")
        BQ_col = column_index_from_string("BQ")
    except ValueError:
        raise ValueError("ERROR: 'AOG' Column or 'BQ' Column are not valid name Column in Excel.")

    print(f"=== 📦 BACKUP DEM RATE ({start_row} → {end_row}) ===")

    for r in range(start_row, end_row + 1):
        aog_cell = ws_summary.cell(row=r, column=AOG_col)
        val = aog_cell.value

        if not isinstance(val, (int, float)):
            continue

        if not is_black(aog_cell.font.color):
            continue

        status = str(ws_summary.cell(row=r, column=BQ_col).value or "").lower()
        key = get_key(ws_summary, r)

        if status == "plan":
            backup_ws = ws_plan
            sheet_label = "PLAN"
        elif status in ("loading", "in progress", "completed"):
            backup_ws = ws_comp
            sheet_label = "COMP"
        else:
            continue

        row_b = find_row(backup_ws, key)
        if not row_b:
            row_b = backup_ws.max_row + 1
            for i, k in enumerate(key, 1):
                backup_ws.cell(row=row_b, column=i).value = k

        def extract_hex(color_obj):
            if not color_obj:
                return "FF000000"

            rgb = getattr(color_obj, "rgb", None)
            if isinstance(rgb, str):
                return rgb  # RGB atau ARGB

            if color_obj.type == "theme":
                return f"FF000000"  # Considered to black for theme 0 or 1
            if color_obj.type == "indexed":
                return f"INDEXED_{color_obj.indexed}"

            return "UNKNOWN"
        
        rgb = extract_hex(aog_cell.font.color)

        backup_ws.cell(row=row_b, column=6).value = val
        backup_ws.cell(row=row_b, column=7).value = rgb

        print(f"[BACKUP] Row {r} → {sheet_label} row {row_b}")
        print(f"         Key     : {key}")
        print(f"         Value   : {val}")
        print(f"         Color   : {rgb}\n")


def restore_dem_rate(ws_summary: Worksheet, plan_sheet_name: str, comp_sheet_name: str, start_row: int, end_row: int):
    """
    Restore DEM Rate with debug print.
    """
    wb = ws_summary.parent

    ws_plan = wb[plan_sheet_name] if plan_sheet_name in wb.sheetnames else None
    ws_comp = wb[comp_sheet_name] if comp_sheet_name in wb.sheetnames else None

    if end_row is None:
        end_row = ws_summary.max_row

    try:
        AOG_col = column_index_from_string("AOG")
    except ValueError:
        raise ValueError("ERROR: 'AOG' Column are not valid name Column in Excel.")

    print(f"\n=== ♻ RESTORE DEM RATE ({start_row} → {end_row}) ===")

    def get_key_from_summary(row):
        return tuple(str(ws_summary.cell(row=row, column=c).value or "").strip() for c in range(3, 8))

    def get_key_from_backup(ws, row):
        return tuple(str(ws.cell(row=row, column=c).value or "").strip() for c in range(1, 6))

    def restore_from_backup(ws_backup, label):
        if not ws_backup:
            print(f"[RESTORE] Sheet {label} not found.")
            return

        print(f"\n--- Restore from sheet {label} ---")

        for r_backup in range(2, ws_backup.max_row + 1):
            backup_key = get_key_from_backup(ws_backup, r_backup)
            backup_value = ws_backup.cell(row=r_backup, column=6).value
            backup_color = ws_backup.cell(row=r_backup, column=7).value

            if backup_value is None:
                continue

            print(f"[CHECK] Backup row {r_backup}")
            print(f"        Key   : {backup_key}")
            print(f"        Value : {backup_value}")
            print(f"        Color : {backup_color}")

            for r_sum in range(start_row, end_row + 1):
                sum_key = get_key_from_summary(r_sum)

                if sum_key != backup_key:
                    continue

                print(f"   → MATCH in summary row {r_sum}, restore data!")

                target_cell = ws_summary.cell(row=r_sum, column=AOG_col)
                target_cell.value = backup_value

                if backup_color:
                    try:
                        target_cell.font = Font(color=backup_color)
                    except Exception:
                        print("     (WARN) Color Format are not Valid, skip font")

                print(f"     ✔ Restored value={backup_value}, color={backup_color}\n")

    restore_from_backup(ws_plan, "PLAN")
    restore_from_backup(ws_comp, "COMP")