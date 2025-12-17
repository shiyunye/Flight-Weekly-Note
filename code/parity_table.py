import pandas as pd

def create_parity_table(
        format_percentage,
        df_direct_parity,
        round_to_nearest_10,
        df_meta_parity):

     # Helper function to safely extract values with a default value
    def safe_extract_value(df, condition, column, default=0):
        filtered = df.loc[condition, column]
        if not filtered.empty:
            return filtered.values[0]
        return default  # Return the default value if no rows match the condition


    df_parity=pd.DataFrame()
    df_parity.loc['Direct vs Expedia','Actual_pcln']=format_percentage(
        safe_extract_value(
            df_direct_parity,
            (df_direct_parity['period'] == 'one_week_ago') & (df_direct_parity['perspective'] == 'PRICELINE') ,
            'lw_parity'
        ) * 100
    )
    df_parity.loc['Direct vs Expedia','YoY (bps)_pcln']=round_to_nearest_10(
        safe_extract_value(
            df_direct_parity,
            (df_direct_parity['period'] == 'one_week_ago') & (df_direct_parity['perspective'] == 'PRICELINE') ,
            'parity_diff')*10000
            )
    df_parity.loc['Direct vs Expedia','YoY PW (bps)_pcln']=round_to_nearest_10(
        safe_extract_value(
            df_direct_parity,
            (df_direct_parity['period'] == 'two_weeks_ago') & (df_direct_parity['perspective'] == 'PRICELINE') ,
            'parity_diff')*10000
    )

    # Kayak Placement
    df_parity.loc['Kayak Placement', 'Actual_pcln'] = format_percentage(
        safe_extract_value(
            df_meta_parity,
            (df_meta_parity['period'] == 'one_week_ago') & (df_meta_parity['site'] == 'PRICELINE') & (df_meta_parity['meta'] == 'Kayak'),
            'lw_placement'
        ) * 100
    )
    df_parity.loc['Kayak Placement', 'YoY (bps)_pcln'] = round_to_nearest_10(
        safe_extract_value(
            df_meta_parity,
            (df_meta_parity['period'] == 'one_week_ago') & (df_meta_parity['site'] == 'PRICELINE') & (df_meta_parity['meta'] == 'Kayak'),
            'placement_diff'
        ) * 10000
    )
    df_parity.loc['Kayak Placement', 'YoY PW (bps)_pcln'] = round_to_nearest_10(
        safe_extract_value(
            df_meta_parity,
            (df_meta_parity['period'] == 'two_weeks_ago') & (df_meta_parity['site'] == 'PRICELINE') & (df_meta_parity['meta'] == 'Kayak'),
            'placement_diff'
        ) * 10000
    )

        # Kayak Placement expedia
    df_parity.loc['Kayak Placement', 'Actual_exp'] = format_percentage(
        safe_extract_value(
            df_meta_parity,
            (df_meta_parity['period'] == 'one_week_ago') & (df_meta_parity['site'] == 'EXPEDIA') & (df_meta_parity['meta'] == 'Kayak'),
            'lw_placement'
        ) * 100
    )
    df_parity.loc['Kayak Placement', 'YoY (bps)_exp'] = round_to_nearest_10(
        safe_extract_value(
            df_meta_parity,
            (df_meta_parity['period'] == 'one_week_ago') & (df_meta_parity['site'] == 'EXPEDIA') & (df_meta_parity['meta'] == 'Kayak'),
            'placement_diff'
        ) * 10000
    )
    df_parity.loc['Kayak Placement', 'YoY PW (bps)_exp'] = round_to_nearest_10(
        safe_extract_value(
            df_meta_parity,
            (df_meta_parity['period'] == 'two_weeks_ago') & (df_meta_parity['site'] == 'EXPEDIA') & (df_meta_parity['meta'] == 'Kayak'),
            'placement_diff'
        ) * 10000
    )


    # Skyscanner Placement
    df_parity.loc['Skyscanner Placement', 'Actual_pcln'] = format_percentage(
        safe_extract_value(
            df_meta_parity,
            (df_meta_parity['period'] == 'one_week_ago') & ( df_meta_parity['site'] == 'PRICELINE') &(df_meta_parity['meta'] == 'Skyscanner'),
            'lw_placement'
        ) * 100
    )
    df_parity.loc['Skyscanner Placement', 'YoY (bps)_pcln'] = round_to_nearest_10(
        safe_extract_value(
            df_meta_parity,
            (df_meta_parity['period'] == 'one_week_ago') & (df_meta_parity['site'] == 'PRICELINE') &(df_meta_parity['meta'] == 'Skyscanner'),
            'placement_diff'
        ) * 10000
    )
    df_parity.loc['Skyscanner Placement', 'YoY PW (bps)_pcln'] = round_to_nearest_10(
        safe_extract_value(
            df_meta_parity,
            (df_meta_parity['period'] == 'two_weeks_ago') &(df_meta_parity['site'] == 'PRICELINE') & (df_meta_parity['meta'] == 'Skyscanner'),
            'placement_diff'
        ) * 10000
    )

    # Skyscanner Placement Expedia
    df_parity.loc['Skyscanner Placement', 'Actual_exp'] = format_percentage(
        safe_extract_value(
            df_meta_parity,
            (df_meta_parity['period'] == 'one_week_ago') & (df_meta_parity['site'] == 'EXPEDIA') & (df_meta_parity['meta'] == 'Skyscanner'),
            'lw_placement'
        ) * 100
    )
    df_parity.loc['Skyscanner Placement', 'YoY (bps)_exp'] = round_to_nearest_10(
        safe_extract_value(
            df_meta_parity,
            (df_meta_parity['period'] == 'one_week_ago') & (df_meta_parity['site'] == 'EXPEDIA') & (df_meta_parity['meta'] == 'Skyscanner'),
            'placement_diff'
        ) * 10000
    )
    df_parity.loc['Skyscanner Placement', 'YoY PW (bps)_exp'] = round_to_nearest_10(
        safe_extract_value(
            df_meta_parity,
            (df_meta_parity['period'] == 'two_weeks_ago') & (df_meta_parity['site'] == 'EXPEDIA') & (df_meta_parity['meta'] == 'Skyscanner'),
            'placement_diff'
        ) * 10000
    )
   
    return df_parity
