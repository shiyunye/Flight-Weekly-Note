import pandas as pd


def format_number(num):
    if pd.isna(num):
        return ''
    if abs(num) >= 1e6:
        return f"{num/1e6:.1f}M"
    elif abs(num) >= 1e3:
        return f"{num/1e3:.0f}K"
    else:
        return "<1K"


def format_percentage(num):
    if pd.isna(num):
        return ''
    return f"{num:.1f}%"


def calculate_metrics(data, end_date, pp_end_date, group_col,
                      value_cy, value_ly, display_name,
                      filters=None, suffix=""):
    """Shared ticket/revenue metric builder used by kpi.py and kpi_rev.py."""
    if filters:
        for col, val in filters.items():
            data = data[data[col] == val]

    grouped_current  = data[data['wk_ending'] == end_date].groupby(group_col)
    grouped_previous = data[data['wk_ending'] == pp_end_date].groupby(group_col)

    df = pd.DataFrame()
    df[display_name]           = grouped_current[value_cy].sum().astype(int)
    df[f'{display_name}_cwly'] = grouped_current[value_ly].sum().astype(int)
    df[f'{display_name}_pw']   = grouped_previous[value_cy].sum().astype(int)
    df[f'{display_name}_pwly'] = grouped_previous[value_ly].sum().astype(int)

    df.loc['Total', display_name:f'{display_name}_pwly'] = [
        df[display_name].sum(), df[f'{display_name}_cwly'].sum(),
        df[f'{display_name}_pw'].sum(), df[f'{display_name}_pwly'].sum(),
    ]

    df['YoY']    = ((df[display_name]           / df[f'{display_name}_cwly'] - 1) * 100).apply(format_percentage)
    df['YoY PW'] = ((df[f'{display_name}_pw']   / df[f'{display_name}_pwly'] - 1) * 100).apply(format_percentage)

    df = df[[display_name, 'YoY', 'YoY PW']]
    df[display_name] = df[display_name].apply(format_number)

    return df.add_suffix(suffix)
