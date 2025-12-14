from openpyxl.styles import PatternFill
from openpyxl.styles.colors import Color

"""
Contains 2 functions:
1. backup_fill_by_status()
   - Backup KEY: values in columns C–G (Month, Company, Ship, Buyer, End User)
   - Backup VALUE: hex values of the contents of columns C–G and ANQ
   - Only for rows with a specific Status (BQ)

2. restore_fill_by_status()
   - Restores the contents based on the C–G KEY
"""

COL_BQ = 69        # Status
COL_C = 3          # Month
COL_G = 7          # End User
COL_ANQ = 1057       # Free Pratique Granted

VALID_STATUS = ("completed", "loading", "in progress", "plan")

def _normalize(val):
    return str(val or "").strip().lower()


def _get_fill_hex(cell):
    """Extract ARGB hex from cell fill safely."""
    fill = cell.fill
    if not fill or not fill.fgColor:
        return None

    fg = fill.fgColor
    print(fg.type, fg.rgb)

    # RGB
    if fg.type == "rgb" and isinstance(fg.rgb, str):
        rgb = fg.rgb.upper()
        if rgb not in ("00000000", "FFFFFFFF"):
            return {"type": "rgb", "value": rgb}

    # THEME
    if fg.type == "theme":
        return {
            "type": "theme",
            "value": fg.theme,
            "tint": fg.tint
        }

    # # RGB / ARGB
    # rgb = getattr(fg, "rgb", None)
    # if not isinstance(rgb, str):
    #     return None
    
    # rgb = rgb.upper()

    return None


def _make_fill(hex_color):
    if not hex_color:
        return None

    if hex_color["type"] == "rgb":
        return PatternFill(
            fill_type="solid",
            fgColor=hex_color["value"]
        )

    if hex_color["type"] == "theme":
        return PatternFill(
            fill_type="solid",
            fgColor=Color(
                theme=hex_color["value"],
                tint=hex_color.get("tint")
            )
        )

    return None

def _apply_fill(cell, hex_color):
    fill = _make_fill(hex_color)
    if fill:
        cell.fill = fill

def backup_fill_by_status(sheet_b):
    """
    Backup fill colors based on key (C–G).

    Return structure:
    {
      (month, company, vessel, buyer, end_user): {
          "C": hex,
          "D": hex,
          "E": hex,
          "F": hex,
          "G": hex,
          "ANQ": hex
      }
    }
    """
    backup = {}
    count = 0

    for r in range(2, sheet_b.max_row + 1):
        status = _normalize(sheet_b.cell(r, COL_BQ).value)
        if status not in VALID_STATUS:
            continue

        key = tuple(sheet_b.cell(r, col).value for col in range(COL_C, COL_G + 1))

        backup[key] = {
            "C": _get_fill_hex(sheet_b.cell(r, COL_C)),
            "D": _get_fill_hex(sheet_b.cell(r, COL_C + 1)),
            "E": _get_fill_hex(sheet_b.cell(r, COL_C + 2)),
            "F": _get_fill_hex(sheet_b.cell(r, COL_C + 3)),
            "G": _get_fill_hex(sheet_b.cell(r, COL_G)),
            "ANQ": _get_fill_hex(sheet_b.cell(r, COL_ANQ)),
        }

        count += 1

    print(f"[BACKUP] {count} rows fill backed up")
    return backup

def restore_fill_by_status(sheet_b, backup_data):
    """
    Restore fill colors using key (C–G).
    """
    restored = 0

    for r in range(2, sheet_b.max_row + 1):
        key = tuple(sheet_b.cell(r, col).value for col in range(COL_C, COL_G + 1))

        if key not in backup_data:
            continue

        fills = backup_data[key]

        _apply_fill(sheet_b.cell(r, COL_C), fills.get("C"))
        _apply_fill(sheet_b.cell(r, COL_C + 1), fills.get("D"))
        _apply_fill(sheet_b.cell(r, COL_C + 2), fills.get("E"))
        _apply_fill(sheet_b.cell(r, COL_C + 3), fills.get("F"))
        _apply_fill(sheet_b.cell(r, COL_G), fills.get("G"))
        _apply_fill(sheet_b.cell(r, COL_ANQ), fills.get("ANQ"))

        restored += 1

    print(f"[RESTORE] {restored} rows fill restored")