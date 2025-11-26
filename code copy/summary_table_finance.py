import pandas as pd

def create_ly_table(df_finance,df_subsummary,format_percentage,format_number):
    df_finance.rename(columns={'gross_cont_revenue_fee': 'gross_cont_revenue', 'net_cont_revenue_fee': 'net_cont_revenue'}, inplace=True)
    df_finance
    df_updated = df_finance[['Gross_Units', 'Net_Units', 'gross_cont_revenue', 'net_cont_revenue']].astype({
        'Gross_Units': 'int',
        'Net_Units': 'int',
        'gross_cont_revenue': 'int',
        'net_cont_revenue': 'int'
    })

    # If you need to keep the rest of the DataFrame unchanged, you can join the other columns back
    df_finance = df_finance.drop(['Gross_Units', 'Net_Units', 'gross_cont_revenue', 'net_cont_revenue'], axis=1).join(df_updated)
    df_finance


    # Calculate the percentages
    def calculate_percentage(current, previous):
        return (((current - previous) / previous) * 100).round(1)

    # Reporting Week
    reporting_week_actuals = df_finance[df_finance['Period'] == '1.Reporting Week'].iloc[0]
    reporting_week_ly = df_finance[df_finance['Period'] == '6.Reporting Week LY'].iloc[0]

    # Previous Week
    previous_week_actuals = df_finance[df_finance['Period'] == '2.Previous Week'].iloc[0]
    previous_week_ly = df_finance[df_finance['Period'] == '7.Previous Week LY'].iloc[0]

    #YTD
    ytd_actuals = df_finance[df_finance['Period'] == '5.YTD'].iloc[0]
    ytd_plan = df_finance[df_finance['Period'] == '5.YTD'].iloc[1]

    # Calculate metrics
    df_finance_ly = {
        'Measure': ['net_tkts_cy', 'gr_tkts_cy', 'net_contribution_cy'
                , 'gross_contribution_cy'
                ],
        'Actual': [
            reporting_week_actuals['Net_Units'],
            reporting_week_actuals['Gross_Units'],
            reporting_week_actuals['net_cont_revenue'],
            reporting_week_actuals['gross_cont_revenue']
        ],
        'Reporting Week': [
            calculate_percentage(reporting_week_actuals['Net_Units'], reporting_week_ly['Net_Units']),
            calculate_percentage(reporting_week_actuals['Gross_Units'], reporting_week_ly['Gross_Units']),
            calculate_percentage(reporting_week_actuals['net_cont_revenue'], reporting_week_ly['net_cont_revenue']),
            # None
            calculate_percentage(reporting_week_actuals['gross_cont_revenue'], reporting_week_ly['gross_cont_revenue'])
        ],
        'Previous Week': [
            calculate_percentage(previous_week_actuals['Net_Units'], previous_week_ly['Net_Units']),
            calculate_percentage(previous_week_actuals['Gross_Units'], previous_week_ly['Gross_Units']),
            calculate_percentage(previous_week_actuals['net_cont_revenue'], previous_week_ly['net_cont_revenue']),
            # None
            calculate_percentage(previous_week_actuals['gross_cont_revenue'], previous_week_ly['gross_cont_revenue'])
        ],
        # 'YTD': [
        #     calculate_percentage(ytd_actuals['Net_Units'], ytd_plan['Net_Units']),
        #     calculate_percentage(ytd_actuals['Gross_Units'], ytd_plan['Gross_Units']),
        #     calculate_percentage(ytd_actuals['net_cont_revenue'], ytd_plan['net_cont_revenue']),
        #     calculate_percentage(ytd_actuals['gross_cont_revenue'], ytd_plan['gross_cont_revenue'])
        # ]
    }
        
    df_finance_ly= pd.DataFrame(df_finance_ly)
    df_finance_ly['Actual']=df_finance_ly['Actual'].apply(format_number)
    df_finance_ly['Reporting Week']=df_finance_ly['Reporting Week'].apply(format_percentage)
    df_finance_ly['Previous Week']=df_finance_ly['Previous Week'].apply(format_percentage)
    df_finance_ly['YTD']= df_subsummary['YTD'][:-2]
    
    return df_finance_ly

def create_plan_table(df_finance,df_summary_ly,format_percentage,format_number):
    # Calculate the percentages
    def calculate_percentage(current, previous):
        return (((current - previous) / previous) * 100).round(1)

    # Reporting Week
    reporting_week_actuals = df_finance[df_finance['Period'] == '1.Reporting Week'].iloc[0]
    reporting_week_plan = df_finance[df_finance['Period'] == '1.Reporting Week'].iloc[1]

    # Previous Week
    previous_week_actuals = df_finance[df_finance['Period'] == '2.Previous Week'].iloc[0]
    previous_week_plan = df_finance[df_finance['Period'] == '2.Previous Week'].iloc[1]

    # MTD
    mtd_actuals = df_finance[df_finance['Period'] == '3.MTD'].iloc[0]
    mtd_plan = df_finance[df_finance['Period'] == '3.MTD'].iloc[1]


    qtd_actuals = df_finance[df_finance['Period'] == '4.QTD'].iloc[0]
    qtd_plan = df_finance[df_finance['Period'] == '4.QTD'].iloc[1]

    ytd_actuals = df_finance[df_finance['Period'] == '5.YTD'].iloc[0]
    ytd_plan = df_finance[df_finance['Period'] == '5.YTD'].iloc[1]

    # Calculate metrics
    finance_plan_data = {
        'Measure': ['net_tkts_cy', 'gr_tkts_cy', 'net_contribution_cy'
                #    , 'gross_contribution_cy'
                ],
        'Reporting Week': [
            calculate_percentage(reporting_week_actuals['Net_Units'], reporting_week_plan['Net_Units']),
            calculate_percentage(reporting_week_actuals['Gross_Units'], reporting_week_plan['Gross_Units']),
            calculate_percentage(reporting_week_actuals['net_cont_revenue'], reporting_week_plan['net_cont_revenue'])
            #,calculate_percentage(reporting_week_actuals['gross_cont_revenue'], reporting_week_plan['gross_cont_revenue'])
        ],
        'Previous Week': [
            calculate_percentage(previous_week_actuals['Net_Units'], previous_week_plan['Net_Units']),
            calculate_percentage(previous_week_actuals['Gross_Units'], previous_week_plan['Gross_Units']),
            calculate_percentage(previous_week_actuals['net_cont_revenue'], previous_week_plan['net_cont_revenue'])
            # ,calculate_percentage(previous_week_actuals['gross_cont_revenue'], previous_week_plan['gross_cont_revenue'])
        ],
        # 'MTD': [
        #     calculate_percentage(mtd_actuals['Net_Units'], mtd_plan['Net_Units']),
        #     calculate_percentage(mtd_actuals['Gross_Units'], mtd_plan['Gross_Units']),
        #     calculate_percentage(mtd_actuals['net_cont_revenue'], mtd_plan['net_cont_revenue'])
        #     #,calculate_percentage(mtd_actuals['gross_cont_revenue'], mtd_plan['gross_cont_revenue'])
        # ],
        # 'QTD': [
        #     calculate_percentage(qtd_actuals['Net_Units'], qtd_plan['Net_Units']),
        #     calculate_percentage(qtd_actuals['Gross_Units'], qtd_plan['Gross_Units']),
        #     calculate_percentage(qtd_actuals['net_cont_revenue'], qtd_plan['net_cont_revenue'])
        #     #,calculate_percentage(qtd_actuals['gross_cont_revenue'], qtd_plan['gross_cont_revenue'])
        # ],
        'YTD': [
            calculate_percentage(ytd_actuals['Net_Units'], ytd_plan['Net_Units']),
            calculate_percentage(ytd_actuals['Gross_Units'], ytd_plan['Gross_Units']),
            calculate_percentage(ytd_actuals['net_cont_revenue'], ytd_plan['net_cont_revenue'])
            # , calculate_percentage(ytd_actuals['gross_cont_revenue'], ytd_plan['gross_cont_revenue'])
        ]
    }

    df_finance_plan = pd.DataFrame(finance_plan_data)
    df_finance_plan['Reporting Week']=df_finance_plan['Reporting Week'].apply(format_percentage)
    df_finance_plan['Previous Week']=df_finance_plan['Previous Week'].apply(format_percentage)
    # df_finance_plan['MTD']=df_finance_plan['MTD'].apply(format_percentage)
    # df_finance_plan['QTD']=df_finance_plan['QTD'].apply(format_percentage)
    df_finance_plan['YTD']=df_finance_plan['YTD'].apply(format_percentage)
    df_finance_plan

    return df_finance_plan