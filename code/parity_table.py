import pandas as pd

def create_parity_table(
        format_percentage,
        df_direct_parity,
        round_to_nearest_10,
        df_meta_parity):

    df_parity=pd.DataFrame()
    df_parity.loc['Direct vs Expedia','Actual_pcln']=format_percentage(df_direct_parity.iloc[1,3]*100)
    df_parity.loc['Direct vs Expedia','YoY (bps)_pcln']=round_to_nearest_10(df_direct_parity.iloc[1,-1]*10000)
    df_parity.loc['Direct vs Expedia','YoY PW (bps)_pcln']=round_to_nearest_10(df_direct_parity.iloc[0,-1]*10000)

    df_parity.loc['Kayak Placement','Actual_pcln']=format_percentage(df_meta_parity.iloc[0,4]*100)
    df_parity.loc['Kayak Placement','YoY (bps)_pcln']=round_to_nearest_10(df_meta_parity.iloc[0,-1]*10000)
    df_parity.loc['Kayak Placement','YoY PW (bps)_pcln']=round_to_nearest_10(df_meta_parity.iloc[4,-1]*10000)
    df_parity.loc['Skyscanner Placement','Actual_pcln']=format_percentage(df_meta_parity.iloc[1,4]*100)
    df_parity.loc['Skyscanner Placement','YoY (bps)_pcln']=round_to_nearest_10(df_meta_parity.iloc[1,-1]*10000)
    df_parity.loc['Skyscanner Placement','YoY PW (bps)_pcln']=round_to_nearest_10(df_meta_parity.iloc[5,-1]*10000)

    df_parity.loc['Kayak Placement','Actual_expe']=format_percentage(df_meta_parity.iloc[2,4]*100)
    df_parity.loc['Kayak Placement','YoY (bps)_expe']=round_to_nearest_10(df_meta_parity.iloc[2,-1]*10000)
    df_parity.loc['Kayak Placement','YoY PW (bps)_expe']=round_to_nearest_10(df_meta_parity.iloc[6,-1]*10000)
    df_parity.loc['Skyscanner Placement','Actual_expe']=format_percentage(df_meta_parity.iloc[3,4]*100)
    df_parity.loc['Skyscanner Placement','YoY (bps)_expe']=round_to_nearest_10(df_meta_parity.iloc[3,-1]*10000)
    df_parity.loc['Skyscanner Placement','YoY PW (bps)_expe']=round_to_nearest_10(df_meta_parity.iloc[7,-1]*10000)
   
    return df_parity
