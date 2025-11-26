import pandas as pd

def create_ytd_ly(df_weekly,df_gds_incentive,end_date,ytd_current_year):
    df_gds_incentive['net_gds_incentives']=df_gds_incentive['net_gds_incentives'].astype(int)

    total_priceline_ytd = (
        df_weekly.loc[
            (df_weekly['wk_ending'] <= end_date)
            & (df_weekly['wk_ending'] >= ytd_current_year)
            & (df_weekly['company'].str.contains('priceline', case=False, na=False))]
            [['net_tkts_cy', 'net_tkts_ly','gr_tkts_cy', 'gr_tkts_ly', 'net_contribution_cy', 'net_contribution_ly',
            'gross_contribution_cy', 'gross_contribution_ly','normalized_net_tickets_cy','normalized_net_tickets_ly',
            'normalized_gross_tickets_cy','normalized_gross_tickets_ly','net_contr_fee_cy','net_contr_fee_ly','gr_contr_fee_cy','gr_contr_fee_ly']]
        .sum()
        .to_frame()
        .T
    )
    total_priceline_ytd_flightonly = (
        df_weekly.loc[
            (df_weekly['brand'].str.contains('Priceline', case=False, na=False))
            & (df_weekly['offer_type'].str.contains('Flights', case=False, na=False))]
            [['net_tkts_cy','net_tkts_ly','net_contr_fee_cy','net_contr_fee_ly','gr_contr_fee_cy','gr_contr_fee_ly']]
        .sum()
        .to_frame()
        .T
    )

    total_priceline_ytd['company'] = 'Priceline YTD Total'

    total_priceline_ytd = total_priceline_ytd[
        ['company', 'net_tkts_cy', 'net_tkts_ly', 'gr_tkts_cy', 'gr_tkts_ly','net_contribution_cy'
        , 'net_contribution_ly', 'gross_contribution_cy', 'gross_contribution_ly',
        'normalized_net_tickets_cy','normalized_net_tickets_ly','normalized_gross_tickets_cy'
        ,'normalized_gross_tickets_ly', 'net_contr_fee_cy','net_contr_fee_ly','gr_contr_fee_cy','gr_contr_fee_ly']
    ]
    total_priceline_ytd['ytd_net'] = (
        total_priceline_ytd['net_tkts_cy'].fillna(0) / total_priceline_ytd['net_tkts_ly'].replace(0, pd.NA)
    ) - 1
    total_priceline_ytd['ytd_gr'] = (
        total_priceline_ytd['gr_tkts_cy'].fillna(0) / total_priceline_ytd['gr_tkts_ly'].replace(0, pd.NA)
    ) - 1

    total_priceline_ytd['ytd_netrev'] = (
        (total_priceline_ytd['net_contribution_cy'].fillna(0)) / (total_priceline_ytd['net_contribution_ly'].replace(0, pd.NA))
    ) - 1
    total_priceline_ytd['ytd_grrev'] = (
        (total_priceline_ytd['gross_contribution_cy']) / (total_priceline_ytd['gross_contribution_ly'].replace(0, pd.NA))
    ) - 1

    total_priceline_ytd['ytd_nornet'] = (
        total_priceline_ytd['normalized_net_tickets_cy'].fillna(0) / total_priceline_ytd['normalized_net_tickets_ly'].replace(0, pd.NA)
    ) - 1

    total_priceline_ytd['ytd_norgr'] = (
        total_priceline_ytd['normalized_gross_tickets_cy'].fillna(0) / total_priceline_ytd['normalized_gross_tickets_ly'].replace(0, pd.NA)
    ) - 1

    total_priceline_ytd['ytd_netconrfee'] = (
        (total_priceline_ytd['net_contr_fee_cy'].fillna(0)+df_gds_incentive['net_gds_incentives'][0]) / (total_priceline_ytd['net_contr_fee_ly'].replace(0, pd.NA)+df_gds_incentive['net_gds_incentives'][1])
    ) - 1

    total_priceline_ytd['ytd_grconrfee'] = (
        (total_priceline_ytd['gr_contr_fee_cy'].fillna(0)+df_gds_incentive['net_gds_incentives'][0]) / (total_priceline_ytd['gr_contr_fee_ly'].replace(0, pd.NA)+df_gds_incentive['net_gds_incentives'][1])
    ) - 1

    total_priceline_ytd['ytd_netconrfee_flightonly'] = (
        (total_priceline_ytd_flightonly['net_contr_fee_cy'].fillna(0)+df_gds_incentive['net_gds_incentives'][0]) / (total_priceline_ytd_flightonly['net_contr_fee_ly'].replace(0, pd.NA)+df_gds_incentive['net_gds_incentives'][1])
    ) - 1

    total_priceline_ytd['ytd_grconrfee_flightonly'] = (
        (total_priceline_ytd_flightonly['gr_contr_fee_cy'].fillna(0)+df_gds_incentive['net_gds_incentives'][0]) / (total_priceline_ytd_flightonly['gr_contr_fee_ly'].replace(0, pd.NA)+df_gds_incentive['net_gds_incentives'][1])
    ) - 1

    summary_data = [
    {
        'Measure': 'Net Tickets',
        'CY': total_priceline_ytd['net_tkts_cy'].iloc[0],
        'LY': total_priceline_ytd['net_tkts_ly'].iloc[0],
        'YTD': total_priceline_ytd['ytd_net'].iloc[0]*100
    },
    {
        'Measure': 'Gross Tickets',
        'CY': total_priceline_ytd['gr_tkts_cy'].iloc[0],
        'LY': total_priceline_ytd['gr_tkts_ly'].iloc[0],
        'YTD': total_priceline_ytd['ytd_gr'].iloc[0]*100
    },
    {
        'Measure': 'Net Revenue(net_contribution_cy)',
        'CY': total_priceline_ytd['net_contribution_cy'].iloc[0],
        'LY': total_priceline_ytd['net_contribution_ly'].iloc[0],
        'YTD': total_priceline_ytd['ytd_netrev'].iloc[0]*100
    },
    {
        'Measure': 'Gross Revenue(gross_contribution_cy)',
        'CY': total_priceline_ytd['gross_contribution_cy'].iloc[0],
        'LY': total_priceline_ytd['gross_contribution_ly'].iloc[0],
        'YTD': total_priceline_ytd['ytd_grrev'].iloc[0]*100
    },
    {
        'Measure': 'Normalized Net Tickets',
        'CY': total_priceline_ytd['normalized_net_tickets_cy'].iloc[0],
        'LY': total_priceline_ytd['normalized_net_tickets_ly'].iloc[0],
        'YTD': total_priceline_ytd['ytd_nornet'].iloc[0]*100
    },
    {
        'Measure': 'Normalized Gross Tickets',
        'CY': total_priceline_ytd['normalized_gross_tickets_cy'].iloc[0],
        'LY': total_priceline_ytd['normalized_gross_tickets_ly'].iloc[0],
        'YTD': total_priceline_ytd['ytd_norgr'].iloc[0]*100
    },
    {
        'Measure': 'Net Contribution+ Fee',
        'CY': total_priceline_ytd['net_contr_fee_cy'].iloc[0]
        # +df_gds_incentive['net_gds_incentives'][0]
        ,
        'LY': total_priceline_ytd['net_contr_fee_ly'].iloc[0]
        # +df_gds_incentive['net_gds_incentives'][1]
        ,
        'YTD': total_priceline_ytd['ytd_netconrfee'].iloc[0]*100
    },
    {
        'Measure': 'Gross Contribution + Fee',
        'CY': total_priceline_ytd['gr_contr_fee_cy'].iloc[0]
        # +df_gds_incentive['net_gds_incentives'][0]
        ,
        'LY': total_priceline_ytd['gr_contr_fee_ly'].iloc[0]
        # +df_gds_incentive['net_gds_incentives'][1]
        ,
        'YTD': total_priceline_ytd['ytd_grconrfee'].iloc[0]*100 
    },
    {
        'Measure': 'GDS Incentive',
        'CY': df_gds_incentive['net_gds_incentives'][0],
        'LY': df_gds_incentive['net_gds_incentives'][1],
        # 'YTD': total_priceline_ytd['ytd_gds_incentive'].iloc[0]*100
    },
    {
        'Measure': 'Net Cont + Fee + Incentives',
        'CY': total_priceline_ytd['net_contr_fee_cy'].iloc[0]+df_gds_incentive['net_gds_incentives'][0]
        ,
        'LY': total_priceline_ytd['net_contr_fee_ly'].iloc[0]+df_gds_incentive['net_gds_incentives'][1]
        ,
        'YTD': total_priceline_ytd['ytd_netconrfee'].iloc[0]*100
    },
    {
        'Measure': 'Gross Cont + Fee + Incentives',
        'CY': total_priceline_ytd['gr_contr_fee_cy'].iloc[0]+df_gds_incentive['net_gds_incentives'][0]
        ,
        'LY': total_priceline_ytd['gr_contr_fee_ly'].iloc[0]+df_gds_incentive['net_gds_incentives'][1]
        ,
        'YTD': total_priceline_ytd['ytd_grconrfee'].iloc[0]*100    
    },
    {
        'Measure': 'Net Cont + Fee + Incentives(Flight Only)',
        'CY': total_priceline_ytd_flightonly['net_contr_fee_cy'].iloc[0]+df_gds_incentive['net_gds_incentives'][0]
        ,
        'LY': total_priceline_ytd_flightonly['net_contr_fee_ly'].iloc[0]+df_gds_incentive['net_gds_incentives'][1]
        ,
        'YTD': total_priceline_ytd['ytd_netconrfee_flightonly'].iloc[0]*100
    },
    {
        'Measure': 'Gross Cont + Fee + Incentives(Flight Only)',
        'CY': total_priceline_ytd_flightonly['gr_contr_fee_cy'].iloc[0]+df_gds_incentive['net_gds_incentives'][0]
        ,
        'LY': total_priceline_ytd_flightonly['gr_contr_fee_ly'].iloc[0]+df_gds_incentive['net_gds_incentives'][1]
        ,
        'YTD': total_priceline_ytd['ytd_grconrfee_flightonly'].iloc[0]*100    
    }
    ]

    # Create DataFrame
    summary_table = pd.DataFrame(summary_data)

    # Format numeric columns
    summary_table['CY'] = summary_table['CY'].round(0)
    summary_table['LY'] = summary_table['LY'].round(0)
    summary_table['YTD'] = summary_table['YTD']
    summary_table

    return summary_table


def create_subsummary_table(df_pricelince,summary_table,df_pricelince_air,format_percentage,format_number):

    df_subsummary = pd.DataFrame()

    df_grouped = df_pricelince.groupby(['wk_ending'])[[
        'net_tkts_cy', 'net_tkts_ly', 'gr_tkts_cy', 'gr_tkts_ly',
        'normalized_net_tickets_cy', 'normalized_net_tickets_ly',
        'normalized_gross_tickets_cy', 'normalized_gross_tickets_ly'
    ]].sum().round(1).transpose()

    df_grouped_air = df_pricelince_air.groupby(['wk_ending'])[[
        'net_contribution_cy', 'net_contribution_ly', 
        'gross_contribution_cy', 'gross_contribution_ly'
    ]].sum().round(1).transpose()
    df_grouped_air
    

    df_grouped=pd.concat([df_grouped,df_grouped_air])

    df_grouped = df_grouped.reindex(['net_tkts_cy', 'net_tkts_ly', 'gr_tkts_cy', 'gr_tkts_ly', 'net_contribution_cy', 'net_contribution_ly', 
        'gross_contribution_cy', 'gross_contribution_ly',
        'normalized_net_tickets_cy', 'normalized_net_tickets_ly',
        'normalized_gross_tickets_cy', 'normalized_gross_tickets_ly']).fillna('')

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
    df_subsummary = pd.concat([df_grouped, df_YoY], axis=1).dropna()


    df_subsummary = df_subsummary.drop(df_subsummary.columns[0], axis=1)
    # Rename specific columns
    df_subsummary = df_subsummary.rename(columns={
        'Actual_grouped': 'Actual',
        'Actual Prior Wk_YoY': 'Previous Week',
        'Actual_YoY': 'Reporting Week',
    })

    df_subsummary['Actual'] = df_subsummary['Actual'].apply(format_number)

    # Reorder the columns in the DataFrame
    new_order=['Actual','Reporting Week',	'Previous Week']
    df_subsummary = df_subsummary[new_order]


    return df_subsummary