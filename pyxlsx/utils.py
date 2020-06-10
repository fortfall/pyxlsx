def trim(value):
    if isinstance(value, str):
        return value.strip()
    return value