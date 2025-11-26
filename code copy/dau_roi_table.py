import pandas as pd

def calculate_dau_conversion(dau_conversion,format_percentage,end_date, pp_end_date, cwly_date, pwly_date):

    #DAU
    dau_conversion['week_ending'] = pd.to_datetime(dau_conversion['week_ending'], errors='coerce')
    df_dau_converison= pd.DataFrame()
    df_dau_converison['DAU'] =  dau_conversion[dau_conversion['week_ending'] == end_date.strftime("%Y-%m-%d")].groupby(['channel'])['engaged_DAU'].sum().astype(int)
    df_dau_converison['DAU_cwly'] = dau_conversion[dau_conversion['week_ending'] == cwly_date.strftime("%Y-%m-%d")].groupby(['channel'])['engaged_DAU'].sum().astype(int)
    df_dau_converison['DAU_pw'] = dau_conversion[dau_conversion['week_ending'] == pp_end_date.strftime("%Y-%m-%d")].groupby(['channel'])['engaged_DAU'].sum().astype(int)
    df_dau_converison['DAU_pwly'] = dau_conversion[dau_conversion['week_ending'] == pwly_date.strftime("%Y-%m-%d")].groupby(['channel'])['engaged_DAU'].sum().astype(int)


    # Add totals
    df_dau_converison.loc['Total', 'DAU':'DAU_pwly'] = [
        df_dau_converison['DAU'].sum(), df_dau_converison['DAU_cwly'].sum(),
        df_dau_converison['DAU_pw'].sum(), df_dau_converison['DAU_pwly'].sum()
    ]


    df_dau_converison['YoY']=(((df_dau_converison['DAU']/df_dau_converison['DAU_cwly'])- 1) * 100).round(1).astype(str) + '%'
    df_dau_converison['YoY_PW']=(((df_dau_converison['DAU_pw']/df_dau_converison['DAU_pwly'])-1)* 100).round(1).astype(str) + '%'


    #Conversion
    df_dau_converison['converted'] =  dau_conversion[dau_conversion['week_ending'] == end_date.strftime("%Y-%m-%d")].groupby(['channel'])['converted'].sum()
    df_dau_converison['converted_cwly'] = dau_conversion[dau_conversion['week_ending'] == cwly_date.strftime("%Y-%m-%d")].groupby(['channel'])['converted'].sum()
    df_dau_converison['converted_pw'] = dau_conversion[dau_conversion['week_ending'] == pp_end_date.strftime("%Y-%m-%d")].groupby(['channel'])['converted'].sum()
    df_dau_converison['converted_pwly'] = dau_conversion[dau_conversion['week_ending'] == pwly_date.strftime("%Y-%m-%d")].groupby(['channel'])['converted'].sum()


     # Add totals
    df_dau_converison.loc['Total', 'converted':'converted_pwly'] = [
        df_dau_converison['converted'].sum(), df_dau_converison['converted_cwly'].sum(),
        df_dau_converison['converted_pw'].sum(), df_dau_converison['converted_pwly'].sum()
    ]


    df_dau_converison['conversion'] =  (df_dau_converison['converted']/df_dau_converison['DAU'])

    # Conversion rates for historical dates
    df_dau_converison['conversion_cwly'] = df_dau_converison['converted_cwly'] / df_dau_converison['DAU_cwly']
    df_dau_converison['conversion_pw'] = df_dau_converison['converted_pw'] / df_dau_converison['DAU_pw']
    df_dau_converison['conversion_pwly'] = df_dau_converison['converted_pwly'] / df_dau_converison['DAU_pwly']

    # YoY and YoY_PW for conversion
    df_dau_converison['conversion_YoY'] = (
        ((df_dau_converison['conversion'] / df_dau_converison['conversion_cwly']) - 1) * 100
    ).round(1).astype(str) + '%'

    df_dau_converison['conversion_YoY_PW'] = (
        ((df_dau_converison['conversion_pw'] / df_dau_converison['conversion_pwly']) - 1) * 100
    ).round(1).astype(str) + '%'
    
    return df_dau_converison


def calculate_roi(df_roi, end_date, pp_end_date, cwly_date, pwly_date):

    df_roi_section= pd.DataFrame()
    df_roi['week_ending'] = pd.to_datetime(df_roi['week_ending'], errors='coerce')

    df_roi_section['contribution'] =  df_roi[df_roi['week_ending'] == end_date.strftime("%Y-%m-%d")].groupby(['channel'])['contribution'].sum().astype(int)
    df_roi_section['contribution_cwly'] = df_roi[df_roi['week_ending'] == cwly_date.strftime("%Y-%m-%d")].groupby(['channel'])['contribution'].sum().astype(int)
    df_roi_section['contribution_pw'] = df_roi[df_roi['week_ending'] == pp_end_date.strftime("%Y-%m-%d")].groupby(['channel'])['contribution'].sum().astype(int)
    df_roi_section['contribution_pwly'] = df_roi[df_roi['week_ending'] == pwly_date.strftime("%Y-%m-%d")].groupby(['channel'])['contribution'].sum().astype(int)

        # Add totals
    df_roi_section.loc['Total', 'contribution':'contribution_pwly'] = [
        df_roi_section['contribution'].sum(), df_roi_section['contribution_cwly'].sum(),
        df_roi_section['contribution_pw'].sum(), df_roi_section['contribution_pwly'].sum()
    ]

    df_roi_section['cost'] =  df_roi[df_roi['week_ending'] == end_date.strftime("%Y-%m-%d")].groupby(['channel'])['cost'].sum()
    df_roi_section['cost_cwly'] = df_roi[df_roi['week_ending'] == cwly_date.strftime("%Y-%m-%d")].groupby(['channel'])['cost'].sum()
    df_roi_section['cost_pw'] = df_roi[df_roi['week_ending'] == pp_end_date.strftime("%Y-%m-%d")].groupby(['channel'])['cost'].sum()
    df_roi_section['cost_pwly'] = df_roi[df_roi['week_ending'] == pwly_date.strftime("%Y-%m-%d")].groupby(['channel'])['cost'].sum()

    df_roi_section.loc['Total', 'cost':'cost_pwly'] = [
        df_roi_section['cost'].sum(), df_roi_section['cost_cwly'].sum(),
        df_roi_section['cost_pw'].sum(), df_roi_section['cost_pwly'].sum()
    ]

    df_roi_section['ROI'] =  df_roi_section['contribution']/df_roi_section['cost']
    df_roi_section['ROI_cwly'] = df_roi_section['contribution_cwly'] / df_roi_section['cost_cwly']
    df_roi_section['ROI_pw'] = df_roi_section['contribution_pw'] / df_roi_section['cost_pw']
    df_roi_section['ROI_pwly'] = df_roi_section['contribution_pwly'] / df_roi_section['cost_pwly']

    # YoY and YoY_PW for conversion
    df_roi_section['ROI_YoY'] = (
        ((df_roi_section['ROI'] / df_roi_section['ROI_cwly']) - 1) * 100
    ).round(1).astype(str) + '%'

    df_roi_section['ROI_YoY_PW'] = (
        ((df_roi_section['ROI_pw'] / df_roi_section['ROI_pwly']) - 1) * 100
    ).round(1).astype(str) + '%'
    
    return df_roi_section


def create_roi_table(df_roi_section,df_dau_converison, end_date, pp_end_date, cwly_date, pwly_date,format_number):
    df_roi_section = df_roi_section.rename(index={
        'CHEAPFLIGHTS': 'Shop PPC Cheapflights',
        'CJ': 'Affiliate',
        'CLICKTRIPZ':'Shop PPC Others',
        'SKYSCANNER' :'Shop PPC Skyscanner',
        'Kayak': 'Shop PPC Kayak'
        })

    df_dau_converison=df_dau_converison.reset_index()
    
    df_merged = (
    df_dau_converison.reset_index()   # channel becomes a column
    .merge(
        df_roi_section.reset_index(),         # channelgroup becomes a column
        left_on="channel",
        right_on="channel",
        how="left"
    ).set_index('channel')
    )

    df_roi_section=df_merged[['DAU','YoY','YoY_PW','conversion','conversion_YoY','conversion_YoY_PW','ROI','ROI_YoY','ROI_YoY_PW']]
    new_cols = [
        "DAU", "YoY", "YoY PW",
        "Conversion", "YoY", "YoY PW",
        "ROI", "YoY", "YoY PW",
    ]
    df_roi_section.columns = new_cols
    df_roi_section['DAU'] = df_roi_section['DAU'].apply(format_number)
    # df_roi_section['ROI'] = df_roi_section['ROI'].round(2)
    df_roi_section["ROI"] = df_roi_section["ROI"].apply(lambda x: "" if pd.isna(x) else f"{x:.2f}")

    df_roi_section=df_roi_section.reindex(['Direct', 'SEM Core', 'SEM Brand', 'Shop PPC Cheapflights','Shop PPC Google', 'Shop PPC Kayak', 'Shop PPC Skyscanner', 'Shop PPC Others','Affiliate','Total']).fillna('')
    
    return df_roi_section



