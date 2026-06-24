import pandas as pd


def format_number(num):
    if pd.isna(num):
        return ''
    if abs(num) >= 1e6:
        return f"{num/1e6:.0f}M"
    elif abs(num) >= 1e3:
        return f"{num/1e3:.0f}K"
    else:
        return "<1K"


def format_percentage(num):
    if pd.isna(num):
        return ''
    return f"{num:.1f}%"
