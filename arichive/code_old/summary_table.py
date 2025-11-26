import pandas as pd

def create_summary_table(df_pricelince,format_number):

    df_summary = pd.DataFrame()

    df_grouped = df_pricelince.groupby(['wk_ending'])[[
        'net_tkts_cy', 'net_tkts_ly', 'gr_tkts_cy', 'gr_tkts_ly',
        'net_contribution_cy', 'net_contribution_ly', 
        'gross_contribution_cy', 'gross_contribution_ly',
        'normalized_net_tickets_cy', 'normalized_net_tickets_ly',
        'normalized_gross_tickets_cy', 'normalized_gross_tickets_ly'
    ]].sum().round(1).transpose()


    # Sort dates to determine the 'Prior Wk' and 'Reporting Wk'
    sorted_dates = sorted(df_grouped.columns,reverse=False)

    # print(dates)
    # # Rename the columns based on the earlier and later dates
    df_grouped.columns = [ 'Actual Prior Wk','Actual']

    # # Create an empty DataFrame to store YoY results
    df_YoY = pd.DataFrame()

    # # Calculate YoY for each relevant pair of CY and LY, round to 1 decimal, and format with '%'
    for metric in ['net_tkts', 'gr_tkts', 'net_contribution', 'gross_contribution', 'normalized_net_tickets', 'normalized_gross_tickets']:
        cy_col = f'{metric}_cy'
        ly_col = f'{metric}_ly'
        
        # Calculate YoY% and round to 1 decimal place
        YoY_percentage = ((df_grouped.loc[cy_col] - df_grouped.loc[ly_col]) / df_grouped.loc[ly_col]) * 100
        YoY_percentage_rounded = YoY_percentage.round(1)
        
        # Format the result with a '%' sign
        df_YoY[f'{metric}_cy'] = YoY_percentage_rounded.astype(str) + '%'

    # Transpose the YoY DataFrame for better readability
    df_YoY = df_YoY.transpose()
    # 

    # Add suffixes to distinguish columns from different DataFrames
    df_grouped = df_grouped.add_suffix('_grouped')
    df_YoY = df_YoY.add_suffix('_YoY')

    # Concatenate DataFrames and drop rows with NaN values
    df_summary = pd.concat([df_grouped, df_YoY], axis=1).dropna()


    df_summary = df_summary.drop(df_summary.columns[0], axis=1)
    # Rename specific columns
    df_summary = df_summary.rename(columns={
        'Actual_grouped': 'Actual',
        'Actual Prior Wk_YoY': 'Previous Week',
        'Actual_YoY': 'Reporting Week',
    })

    df_summary['Actual'] = df_summary['Actual'].apply(format_number)

    # Reorder the columns in the DataFrame
    new_order=['Actual','Reporting Week',	'Previous Week']
    df_summary = df_summary[new_order]

    return df_summary