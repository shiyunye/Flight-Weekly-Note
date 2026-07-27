import pandas as pd
from utils import calculate_metrics as _calculate_metrics


def calculate_metrics(data, end_date, pp_end_date, group_col, filters=None, suffix=""):
    return _calculate_metrics(
        data, end_date, pp_end_date, group_col,
        value_cy='net_contr_fee_cy', value_ly='net_contr_fee_ly',
        display_name='Net Cont w/ Fee',
        filters=filters, suffix=suffix,
    )


def calculate_business_metrics(df_priceline, end_date, pp_end_date):
    df_standalone = calculate_metrics(
        df_priceline, end_date, pp_end_date,
        'company', filters={'offer_type': 'Flights Only'}, suffix='_standalone'
    )
    df_package = calculate_metrics(
        df_priceline, end_date, pp_end_date,
        group_col='company', filters={'offer_type': 'Packages'}, suffix='_package'
    )
    df_total = calculate_metrics(
        df_priceline, end_date, pp_end_date,
        'company', suffix='_total'
    )

    df_business = pd.concat([df_standalone, df_package, df_total], axis=1)
    df_business.index = df_business.index.str.replace('Priceline B2C', 'B2C')
    df_business.index = df_business.index.str.replace('Priceline B2B', 'B2B')
    df_business = df_business.reindex(['B2C', 'B2B', 'Total']).fillna('')
    return df_business


def calculate_carrier_metrics(df_priceline_b2c_standalone, end_date, pp_end_date):
    df_retail = calculate_metrics(
        df_priceline_b2c_standalone, end_date, pp_end_date,
        'carrier', filters={'offer_method_code': 'Retail (Disclosed)'}, suffix='_Retail'
    )
    df_opaque = calculate_metrics(
        df_priceline_b2c_standalone, end_date, pp_end_date,
        'carrier', filters={'offer_method_code': 'Opaque (Non-disclosed)'}, suffix='_Opaque'
    )
    df_total = calculate_metrics(
        df_priceline_b2c_standalone, end_date, pp_end_date,
        'carrier', suffix='_Total'
    )

    df_carrier = pd.concat([df_retail, df_opaque, df_total], axis=1)
    df_carrier = df_carrier.reindex([
        'American Airlines (AA)', 'Delta Air Lines (DL)', 'United Airlines (UA)',
        'Southwest Airlines (WN)', 'Frontier Airlines (F9)',
        'Alaska Airlines (AS)', 'JetBlue Airways (B6)', 'Other', 'Total'
    ]).fillna('')
    return df_carrier


def calculate_channel_metrics(df_priceline_b2c_standalone, end_date, pp_end_date):
    df_app = calculate_metrics(
        df_priceline_b2c_standalone, end_date, pp_end_date,
        'search_channel_group', filters={'application': 'App'}, suffix='_App'
    )
    df_desk_mweb = calculate_metrics(
        df_priceline_b2c_standalone, end_date, pp_end_date,
        'search_channel_group', filters={'application': 'Desk/MWEB'}, suffix='_Desk/MWEB'
    )
    df_total = calculate_metrics(
        df_priceline_b2c_standalone, end_date, pp_end_date,
        'search_channel_group', suffix='_Total'
    )

    df_channel = pd.concat([df_app, df_desk_mweb, df_total], axis=1)
    df_channel = df_channel.reindex([
        'Direct', 'Web Marketing', 'Shop PPC', 'Affiliate', 'Total'
    ]).fillna('')
    return df_channel


def calculate_source_metrics(df_priceline_b2c_standalone, end_date, pp_end_date):
    df_published = calculate_metrics(
        df_priceline_b2c_standalone, end_date, pp_end_date,
        'gds_booking_category', filters={'fare_type_group': 'Published'}, suffix='_Published'
    )
    df_private = calculate_metrics(
        df_priceline_b2c_standalone, end_date, pp_end_date,
        'gds_booking_category', filters={'fare_type_group': 'Private'}, suffix='_Private'
    )
    df_total = calculate_metrics(
        df_priceline_b2c_standalone, end_date, pp_end_date,
        'gds_booking_category', suffix='_Total'
    )

    df_source = pd.concat([df_published, df_private, df_total], axis=1)
    df_source.index = df_source.index.str.replace('Indirect', 'Indirect Connect')
    df_source = df_source.reindex(['Direct Connect', 'Indirect Connect', 'Phone Sales', 'Total']).fillna('')
    return df_source


def calculate_brand_metrics(df_weekly, end_date, pp_end_date):
    df_us_origin = calculate_metrics(
        df_weekly, end_date, pp_end_date,
        'company', filters={'orig_country_group': 'US Origin'}, suffix='_US Origin'
    )
    df_intl_origin = calculate_metrics(
        df_weekly, end_date, pp_end_date,
        'company', filters={'orig_country_group': 'Intl Origin'}, suffix='_Intl Origin'
    )
    df_total = calculate_metrics(
        df_weekly, end_date, pp_end_date,
        'company', suffix='_Total'
    )

    df_brand = pd.concat([df_us_origin, df_intl_origin, df_total], axis=1)
    return df_brand
