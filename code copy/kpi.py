import pandas as pd

def format_number(num):
    if pd.isna(num):
        return ''
    if abs(num) >= 1e6:
        return f"{num/1e6:.0f}M"
    elif abs(num) >= 1e3:
        return f"{num/1e3:.0f}K"
    elif abs(num) < 1e3:
        return "<1K"
    return f"{num:.0f}"

def format_percentage(num):
    if pd.isna(num):
        return ''
    return f"{num:.1f}%"

def round_to_nearest_10(num):
  return round(num / 10) * 10

def calculate_metrics(data,end_date, pp_end_date, group_col, filters=None, suffix=""):
    if filters:
        for col, val in filters.items():
            data = data[data[col] == val]

    grouped_current = data[data['wk_ending'] == end_date].groupby(group_col)
    grouped_previous = data[data['wk_ending'] == pp_end_date].groupby(group_col)

    df = pd.DataFrame()
    df['Net Tickets'] = grouped_current['net_tkts_cy'].sum().astype(int)
    df['Net Tickets_cwly'] = grouped_current['net_tkts_ly'].sum().astype(int)
    df['Net Tickets_PW'] = grouped_previous['net_tkts_cy'].sum().astype(int)
    df['Net Tickets_PWly'] = grouped_previous['net_tkts_ly'].sum().astype(int)

    # Add totals
    df.loc['Total', 'Net Tickets':'Net Tickets_PWly'] = [
        df['Net Tickets'].sum(), df['Net Tickets_cwly'].sum(),
        df['Net Tickets_PW'].sum(), df['Net Tickets_PWly'].sum()
    ]

    # Calculate YoY and YoY PW
    df['YoY'] = ((df['Net Tickets'] / df['Net Tickets_cwly'] - 1) * 100).round(1).astype(str) + '%'
    df['YoY PW'] = ((df['Net Tickets_PW'] / df['Net Tickets_PWly'] - 1) * 100).round(1).astype(str) + '%'

    # Select and format columns
    df = df[['Net Tickets', 'YoY', 'YoY PW']]
    df['Net Tickets'] = df['Net Tickets'].apply(format_number)

    return df.add_suffix(suffix)

# Example application of the generalized function for each use case

def calculate_business_metrics(df_pricelince,end_date,pp_end_date):
    df_business= pd.DataFrame()
    df_standalone = calculate_metrics(
        df_pricelince,
        end_date, 
        pp_end_date, 
        'company', 
        filters={'offer_type': 'Flights Only'}, 
        suffix='_standalone'
    )
    df_package = calculate_metrics(
        df_pricelince, 
        end_date,
        pp_end_date, 
        group_col='company', 
        filters={'offer_type': 'Packages'},
        suffix='_package'
    )
    df_total = calculate_metrics(
        df_pricelince, 
        end_date,
        pp_end_date, 
        'company', 
        suffix='_total'
    )

    df_business = pd.concat([df_standalone, df_package, df_total], axis=1)
    df_business.index = df_business.index.str.replace('Priceline B2C', 'B2C')
    df_business.index = df_business.index.str.replace('Priceline B2B', 'B2B')
    df_business = df_business.reindex([
        'B2C', 'B2B', 'Total'
    ]).fillna('')
    return df_business


def calculate_carrier_metrics(df_pricelince_b2c_standalone,end_date, pp_end_date):
    df_retail = calculate_metrics(
        df_pricelince_b2c_standalone,end_date, pp_end_date, 'carrier', filters={'offer_method_code': 'Retail (Disclosed)'}, suffix='_Retail'
    )
    df_opaque = calculate_metrics(
        df_pricelince_b2c_standalone,end_date, pp_end_date, 'carrier', filters={'offer_method_code': 'Opaque (Non-disclosed)'}, suffix='_Opaque'
    )
    df_total = calculate_metrics(
        df_pricelince_b2c_standalone,end_date, pp_end_date, 'carrier', suffix='_Total'
    )

    df_carrier = pd.concat([df_retail, df_opaque, df_total], axis=1)
    df_carrier = df_carrier.reindex([
        'American Airlines (AA)', 'Delta Air Lines (DL)', 'United Airlines (UA)','Southwest Airlines (WN)',
        'Spirit Airlines (NK)', 'Frontier Airlines (F9)', 'Alaska Airlines (AS)',
        'JetBlue Airways (B6)','Other', 'Total'
    ]).fillna('')
    return df_carrier

def calculate_channel_metrics(df_pricelince_b2c_standalone,end_date, pp_end_date):
    df_app = calculate_metrics(
        df_pricelince_b2c_standalone,end_date, pp_end_date, 'search_channel_group', filters={'application': 'App'}, suffix='_App'
    )
    df_desk_mweb = calculate_metrics(
        df_pricelince_b2c_standalone,end_date, pp_end_date, 'search_channel_group', filters={'application': 'Desk/MWEB'}, suffix='_Desk/MWEB'
    )
    df_total = calculate_metrics(
        df_pricelince_b2c_standalone,end_date, pp_end_date, 'search_channel_group', suffix='_Total'
    )

    df_channel = pd.concat([df_app, df_desk_mweb, df_total], axis=1)
    # df_channel.index = df_channel.index.str.replace('SEM_Core', 'SEM Core')
    # df_channel.index = df_channel.index.str.replace('SEM_Brand', 'SEM Brand')
    # df_channel.index = df_channel.index.str.replace('SEM_Brand', 'SEM Brand')
    df_channel = df_channel.reindex([
        'Direct', 'Web Marketing', 'Shop PPC', 'Affiliate','Total'
    ]).fillna('')
    return df_channel

def calculate_source_metrics(df_pricelince_b2c_standalone,end_date, pp_end_date):
    df_published = calculate_metrics(
        df_pricelince_b2c_standalone,end_date, pp_end_date, 'gds_booking_category', filters={'fare_type_group': 'Published'}, suffix='_Published'
    )
    df_private = calculate_metrics(
        df_pricelince_b2c_standalone,end_date, pp_end_date, 'gds_booking_category', filters={'fare_type_group': 'Private'}, suffix='_Private'
    )
    df_total = calculate_metrics(
        df_pricelince_b2c_standalone,end_date, pp_end_date, 'gds_booking_category', suffix='_Total'
    )

    df_source = pd.concat([df_published, df_private, df_total], axis=1)
    df_source.index = df_source.index.str.replace('Indirect', 'Indirect Connect')
    df_source = df_source.reindex(['Direct Connect', 'Indirect Connect','Phone Sales','Total']).fillna('')
    return df_source

def calculate_brand_metrics(df_weekly,end_date, pp_end_date):
    df_us_origin = calculate_metrics(
        df_weekly,end_date, pp_end_date, 'company', filters={'orig_country_group': 'US Origin'}, suffix='_US Origin'
    )
    df_intl_origin = calculate_metrics(
        df_weekly,end_date, pp_end_date, 'company', filters={'orig_country_group': 'Intl Origin'}, suffix='_Intl Origin'
    )
    df_total = calculate_metrics(
        df_weekly,end_date, pp_end_date, 'company', suffix='_Total'
    )

    df_brand = pd.concat([df_us_origin, df_intl_origin, df_total], axis=1)
    return df_brand