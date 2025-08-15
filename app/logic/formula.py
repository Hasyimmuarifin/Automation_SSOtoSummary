from openpyxl.formula.translate import Translator

def apply_translated_formulas(sheet, template_row=None, start_row=None, end_row=None, included_columns=None, formulas=None):
    """
    Apply formulas either from a template row or directly from provided formulas.
    
    :param sheet: Worksheet object
    :param template_row: Row number to copy formulas from (if used)
    :param start_row: First row where formulas will be applied
    :param end_row: Last row where formulas will be applied
    :param included_columns: List of column letters where formulas will be applied
    :param formulas: Dict {column_letter: formula_string} to insert formulas directly
                     Example: {'B': '=A{row}+1', 'BJ': '=SUM(C{row}:E{row})'}
    """

    for row in range(start_row, end_row + 1):
        for col in included_columns:
            cell = sheet[f"{col}{row}"]

            if formulas and col in formulas:  
                # Format formula with row number
                formula = formulas[col].replace("{row}", str(row))
                cell.value = formula
            elif template_row:
                # Copy formula from template row if available
                template_cell = sheet[f"{col}{template_row}"]
                if template_cell.value and isinstance(template_cell.value, str) and template_cell.value.startswith("="):
                    cell.value = Translator(template_cell.value, origin=template_cell.coordinate).translate_formula(cell.coordinate)    