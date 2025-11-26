import pandas as pd

def calculate_roi(roi, end_date, device=None, suffix=''):
    # Filter by end date and device (if specified)
    df = roi[roi['week_ending'] == end_date]
    if device:
        df = df[df['device'] == device]

    # Ensure numeric columns are float types to avoid Decimal-related errors
    numeric_columns = ['cost', 'ly_cost', 'contribution']
    for col in numeric_columns:
        df[col] = df[col].astype(float)

    # Group and calculate metrics
    result = df.groupby(['channel'])[['cost', 'ly_cost', 'contribution']].sum()
    result['ROI'] = (result['contribution'] / result['cost']).round(2)
    result['Cost YoY'] = (
        (result['cost'] / result['ly_cost'] - 1)
        .replace([float('inf'), -float('inf')], '')
        .fillna('')
        * 100
    ).round(1)

    # Add total row for 'cost', 'ly_cost', and 'contribution'
    total_values = {
        'cost': df['cost'].sum(),
        'ly_cost': df['ly_cost'].sum(),
        'contribution': df['contribution'].sum(),
    }
    total_values['ROI'] = (
        (total_values['contribution'] / total_values['cost']).round(2) if total_values['cost'] > 0 else ' '
    )
    total_values['Cost YoY'] = (
        ((total_values['cost'] / total_values['ly_cost'] - 1) * 100).round(1)
        if total_values['ly_cost'] > 0
        else ''
    )
    result.loc['Total'] = total_values

    # Add suffix for column names
    result = result.add_suffix(suffix)
    return result

def create_roi_table(roi, end_date, format_number):
    # Generate ROI tables for App, Desk/MWEB, and Total
    roi_app = calculate_roi(roi, end_date, device='App', suffix='_App')
    roi_desk_mweb = calculate_roi(roi, end_date, device='Desk/MWEB', suffix='_Desk/MWEB')
    roi_total = calculate_roi(roi, end_date, suffix='_Total')

    # Combine results
    df_roi = pd.concat([roi_app, roi_desk_mweb, roi_total], axis=1)

    # Reorder rows based on expected channels
    expected_channels = ['Direct', 'SEM Core', 'SEM Brand', 'Shop PPC', 'Affiliate', 'Email', 'Meta', 'Total']
    df_roi = df_roi.reindex(expected_channels).fillna('')

    # Add '%' to Cost YoY columns
    for col in ['Cost YoY_App', 'Cost YoY_Desk/MWEB', 'Cost YoY_Total']:
        df_roi[col] = df_roi[col].apply(lambda x: f"{x:.1f}%" if x != '' else x)

    # Select and organize final columns
    columns_order = [
        'cost_App', 'Cost YoY_App', 'ROI_App',
        'cost_Desk/MWEB', 'Cost YoY_Desk/MWEB', 'ROI_Desk/MWEB',
        'cost_Total', 'Cost YoY_Total', 'ROI_Total'
    ]
    df_roi = df_roi[columns_order]

    # Apply numeric formatting to cost columns
    numeric_columns = ['cost_App', 'cost_Desk/MWEB', 'cost_Total']
    for col in numeric_columns:
        df_roi[col] = pd.to_numeric(df_roi[col], errors='coerce').apply(format_number)
    
    df_roi.columns = [col.replace("cost", "Cost") for col in df_roi.columns]
    df_roi.columns = [col.replace("yoy", "YoY") for col in df_roi.columns]
    df_roi.columns = [col.replace("Roi", "ROI") for col in df_roi.columns]
    
    for col in df_roi.columns:
        if "Cost" in col:
            df_roi[col] = df_roi[col].apply(lambda x: f"${x:.1f}" if isinstance(x, (int, float)) else x)

    return df_roi





def calculate_conversion(dau,end_date, device=None, suffix=''):
    df = dau[(dau['week_ending'] == end_date)]
    
    if device:
        df = df[df['device'] == device]
    
    result = df.groupby(['channel'])[['conversions', 'daily_active_users']].sum().astype(int)
    
    # Calculate additional metrics
    result['DAU'] = (result['daily_active_users']).astype(int)
    result['Conversion'] = (result['conversions']*100/ result['daily_active_users']).round(2)

  
    total_values = {
        'conversions': df['conversions'].sum(),
        'DAU': df['daily_active_users'].sum()    }

    total_values['Conversion'] = (total_values['conversions'] / total_values['DAU']).round(2)
    result.loc['Total'] = total_values
    
    # Add suffix for clarity
    result = result.add_suffix(suffix)
    return result

def create_conversion_table(dau,end_date, format_number):
    # Calculate  dau metrics for App, Desk/MWEB, and Total
    conversion_app = calculate_conversion(dau, end_date,device='App', suffix='_App')
    conversion_desk_mweb = calculate_conversion(dau,end_date, device='Desk/MWEB', suffix='_Desk/MWEB')
    conversion_total = calculate_conversion(dau, end_date,suffix='_Total')

    # Combine the results
    df_conversion = pd.concat([conversion_app, conversion_desk_mweb, conversion_total] , axis=1)

    # Reorder and align the channels
    expected_channels = ['Direct', 'SEM Core', 'SEM Brand', 'Shop PPC', 'Affiliate', 'Email', 'Meta','Total']
    df_conversion = df_conversion.reindex(expected_channels).fillna('')


    df_conversion['Conversion_App'] = df_conversion['Conversion_App'].apply(lambda x: f"{x:.1f}%" if x != '' else x)
    df_conversion['Conversion_Desk/MWEB'] = df_conversion['Conversion_Desk/MWEB'].apply(lambda x: f"{x:.1f}%" if x != '' else x)
    df_conversion['Conversion_Total'] = df_conversion['Conversion_Total'].apply(lambda x: f"{x:.1f}%" if x != '' else x)
                                                                                        
    # Select and organize final columns for output
    columns_order = [
        'DAU_App', 'Conversion_App', 
        'DAU_Desk/MWEB', 'Conversion_Desk/MWEB', 
        'DAU_Total', 'Conversion_Total'
    ]
    df_conversion = df_conversion[columns_order]

    numeric_columns = ['DAU_Desk/MWEB','DAU_Total']
    for col in numeric_columns:
        # Convert to numeric, forcing non-convertible values to NaN
        df_conversion[col] = pd.to_numeric(df_conversion[col], errors='coerce')
        
        # Apply formatting only if the column is numeric
        if df_conversion[col].dtype in ['int64', 'float64']:
            df_conversion[col] = df_conversion[col].apply(format_number)
        else:
            print(f"Column {col} is not numeric and cannot be formatted.")
    return df_conversion