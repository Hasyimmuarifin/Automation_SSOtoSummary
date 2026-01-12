def handle_cross_year_december(sheet_a, sheet_b, sheet_c, valid_months: set, header_columns_a: dict, column_mapping: dict, month_to_abbreviation, process_data_per_month) :
    """
    Handle special case for cross-year (Dec → Jan).

    Rules:
    - Triggered ONLY when valid_months contains both 12 and 1
    - If December exists in raw data (sheet_a):
        → process normally (update data)
    - If December does NOT exist in raw data:
        → normalization only (no data update)
    """

    # 🔒 Guard clause — only cross-year
    if not (12 in valid_months and 1 in valid_months):
        return False  # not handled

    print("🔁 Cross-year detected (Dec → Jan)")

    # Step 1: Check if December exists in raw data
    has_december_in_raw = False

    for row in range(2, sheet_a.max_row + 1):
        val = sheet_a.cell(row=row, column=header_columns_a['Month']).value
        if val == 12:
            has_december_in_raw = True
            break

    # Step 2: Process December
    month_abbreviation = month_to_abbreviation(12)

    if has_december_in_raw:
        print("🟢 December found in raw data → normal processing")
    else:
        print("🧹 December NOT found in raw data → normalization only")

    # process_data_per_month is SAFE:
    # - will not append data if no raw match
    # - will normalize + apply formula
    process_data_per_month(
        sheet_a,
        sheet_b,
        sheet_c,
        12,
        month_abbreviation,
        header_columns_a,
        column_mapping
    )

    return True  # handled