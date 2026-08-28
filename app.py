def clean_date_str(val):
    """將 ISO 時間格式清理為 YYYY-MM-DD"""
    if not val:
        return ""
    val_str = str(val).strip()
    if 'T' in val_str:
        return val_str.split('T')[0]
    if len(val_str) >= 10 and val_str[4] == '-' and val_str[7] == '-':
        return val_str[:10]
    return val_str
