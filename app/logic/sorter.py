import datetime

# Custom order for loading ports
loading_port_group = {
    "BoCT": 0,
    # "SMD Anc: 1"
    "Muara Berau": 1,
    "Muara Jawa": 1,
    "Bunyut": 2,
    "GPK Port": 3
}

# Prefixes used to prioritize vessel names
custom_order_e = ["MV", "BG", "DUMP"]

def get_order_index_h(val):
    """
    Get the index of the loading port from the custom order loading port group.

    Parameters:
        val : The loading port value.

    Returns:
        int: The index in loading_port_group if found, otherwise a large index to sort it last.
    """
    return loading_port_group.get(val, 999)


def get_order_index_e(val):
    """
    Get the index of a vessel name based on its prefix (e.g., MV, BG, DUMP).

    Parameters:
        val (str): The vessel name.

    Returns:
        int: The index in custom_order_e based on prefix, or a large index if no match.
    """
    if val is None:
        return len(custom_order_e)
    val_str = str(val).upper()
    for i, prefix in enumerate(custom_order_e):
        if val_str.startswith(prefix):
            return i
    return len(custom_order_e)


def get_order_index_k(val):
    """
    Convert a datetime string or object to a datetime object for sorting.

    Parameters:
        val (str or datetime): The ETB value to parse.

    Returns:
        datetime: Parsed datetime object or datetime.max if parsing fails.
    """
    try:
        if isinstance(val, datetime.datetime):
            return val
        return datetime.datetime.strptime(str(val), "%d/%m/%Y %H:%M:%S")
    except:
        return datetime.datetime.max


def sort_data_rows(data_rows):
    """
    Sort the data rows based on:
      1. Loading port (column H, index 7)
      2. Vessel name (column E, index 4)
      3. ETB time (column K, index 10)

    Parameters:
        data_rows (list of list): The rows of data to sort.

    Returns:
        list of list: The sorted data rows.
    """
    return sorted(
        data_rows,
        key=lambda row: (
            get_order_index_h(row[7]),   # Custom order for loading port
            get_order_index_e(row[4]),   # Prefix order for vessel name
            get_order_index_k(row[10])   # Datetime parsing for ETB
        )
    )