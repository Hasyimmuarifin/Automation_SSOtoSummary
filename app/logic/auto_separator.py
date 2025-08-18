import locale

def get_formula_separator():
    try:
        conv = locale.localeconv()
        decimal_sep = conv['decimal_point']
        # Kalau decimal separator koma, berarti Excel pakai ; untuk formula
        return ";" if decimal_sep == "," else ","
    except:
        return ";"  # fallback aman