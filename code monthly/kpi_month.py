import pandas as pd

def calculate_metrics(
    data, start_date, end_date, pp_start_date, pp_end_date, cwly_start_date, cwly_end_date,       
    pwly_start_date, pwly_end_date,  
    format_number, group_col, filters=None, suffix=""
):
    data = data.copy()

    data["trans_date"] = pd.to_datetime(data["trans_date"], errors="coerce")
    start_date    = pd.Timestamp(start_date)
    end_date      = pd.Timestamp(end_date)
    pp_start_date = pd.Timestamp(pp_start_date)
    pp_end_date   = pd.Timestamp(pp_end_date)
    cwly_start_date =  pd.Timestamp(cwly_start_date)
    cwly_end_date = pd.Timestamp(cwly_end_date)
    pwly_start_date= pd.Timestamp(pwly_start_date)
    pwly_end_date=pd.Timestamp(pwly_end_date)

    if filters:
        for col, val in filters.items():
            data = data[data[col] == val]

    current_mask = (data["trans_date"] >= start_date) & (data["trans_date"] <= end_date)
    prev_mask    = (data["trans_date"] >= pp_start_date) & (data["trans_date"] <= pp_end_date)
    current_ly_mask = (data["trans_date"] >= cwly_start_date) & (data["trans_date"] <= cwly_end_date)
    prev_ly_mask    = (data["trans_date"] >= pwly_start_date) & (data["trans_date"] <= pwly_end_date)

    grouped_current  = data.loc[current_mask].groupby(group_col, dropna=False)
    grouped_current_ly = data.loc[current_ly_mask].groupby(group_col, dropna=False)
    grouped_previous = data.loc[prev_mask].groupby(group_col, dropna=False)
    grouped_previous_ly= data.loc[prev_ly_mask].groupby(group_col, dropna=False)

    df = pd.DataFrame(index=sorted(data[group_col].dropna().unique()))
    df["Net Tickets"]      = grouped_current["net_tkts_cy"].sum()
    df["Net Tickets_cwly"] = grouped_current_ly["net_tkts_cy"].sum()
    df["Net Tickets_PW"]   = grouped_previous["net_tkts_cy"].sum()
    df["Net Tickets_PWly"] = grouped_previous_ly["net_tkts_cy"].sum()

    for c in ["Net Tickets", "Net Tickets_cwly", "Net Tickets_PW", "Net Tickets_PWly"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)

    df.loc["Total", ["Net Tickets", "Net Tickets_cwly", "Net Tickets_PW", "Net Tickets_PWly"]] = [
        df["Net Tickets"].sum(),
        df["Net Tickets_cwly"].sum(),
        df["Net Tickets_PW"].sum(),
        df["Net Tickets_PWly"].sum(),
    ]

    yoy = (df["Net Tickets"] / df["Net Tickets_cwly"].replace({0: pd.NA}) - 1) * 100
    yoy_pw = (df["Net Tickets_PW"] / df["Net Tickets_PWly"].replace({0: pd.NA}) - 1) * 100

    df["YoY"] = yoy.round(1).astype(str)+ "%"
    df["YoY PW"] = yoy_pw.round(1).astype(str)+ "%"

    df = df[["Net Tickets", "YoY", "YoY PW"]]
    df["Net Tickets"] = df["Net Tickets"].apply(format_number)

    return df.add_suffix(suffix)


def calculate_business_metrics(df_priceline, start_date, end_date,    pp_start_date,
    pp_end_date, 
    cwly_start_date, 
    cwly_end_date,       
    pwly_start_date, 
    pwly_end_date,       
    format_number):
    df_standalone = calculate_metrics(
        df_priceline, start_date, end_date, pp_start_date, pp_end_date,
        cwly_start_date, cwly_end_date,pwly_start_date, pwly_end_date, format_number,
        group_col="company", filters={"offer_type": "Flights Only"}, suffix="_standalone",
    )
    df_package = calculate_metrics(
        df_priceline, start_date, end_date, pp_start_date, pp_end_date,  cwly_start_date, cwly_end_date,pwly_start_date, pwly_end_date,format_number,
        group_col="company", filters={"offer_type": "Packages"}, suffix="_package",
    )
    df_total = calculate_metrics(
        df_priceline, start_date, end_date, pp_start_date, pp_end_date,  cwly_start_date, cwly_end_date,pwly_start_date, pwly_end_date,format_number,
        group_col="company", filters=None, suffix="_total",
    )

    df_business = pd.concat([df_standalone, df_package, df_total], axis=1)
    df_business.index = df_business.index.astype(str)
    df_business.index = df_business.index.str.replace("Priceline B2C", "B2C", regex=False)
    df_business.index = df_business.index.str.replace("Priceline B2B", "B2B", regex=False)
    return df_business.reindex(["B2C", "B2B", "Total"]).fillna("")


def calculate_carrier_metrics(df_pricelince_b2c_standalone, start_date, end_date, pp_start_date, pp_end_date,cwly_start_date, cwly_end_date,       
    pwly_start_date, pwly_end_date, format_number):
    df_retail = calculate_metrics(
        df_pricelince_b2c_standalone, start_date, end_date, pp_start_date, pp_end_date, cwly_start_date, cwly_end_date,pwly_start_date, pwly_end_date, format_number,
        "carrier", filters={"offer_method_code": "Retail (Disclosed)"}, suffix="_Retail",
    )
    df_opaque = calculate_metrics(
        df_pricelince_b2c_standalone, start_date, end_date, pp_start_date, pp_end_date,  cwly_start_date, cwly_end_date,pwly_start_date, pwly_end_date,format_number,
        "carrier", filters={"offer_method_code": "Opaque (Non-disclosed)"}, suffix="_Opaque",
    )
    df_total = calculate_metrics(
        df_pricelince_b2c_standalone, start_date, end_date, pp_start_date, pp_end_date, cwly_start_date, cwly_end_date,pwly_start_date, pwly_end_date,format_number,
        "carrier", suffix="_Total",
    )

    df_carrier = pd.concat([df_retail, df_opaque, df_total], axis=1)
    return df_carrier.reindex([
        "American Airlines (AA)", "Delta Air Lines (DL)", "United Airlines (UA)", "Southwest Airlines (WN)",
        "Spirit Airlines (NK)", "Frontier Airlines (F9)", "Alaska Airlines (AS)",
        "JetBlue Airways (B6)", "Other", "Total",
    ]).fillna("")


def calculate_channel_metrics(df_pricelince_b2c_standalone, start_date, end_date, pp_start_date, pp_end_date,cwly_start_date, cwly_end_date,       
    pwly_start_date, pwly_end_date,  format_number):
    df_app = calculate_metrics(
        df_pricelince_b2c_standalone, start_date, end_date, pp_start_date, pp_end_date, cwly_start_date, cwly_end_date,pwly_start_date, pwly_end_date, format_number,
        "search_channel_group", filters={"application": "App"}, suffix="_App",
    )
    df_desk_mweb = calculate_metrics(
        df_pricelince_b2c_standalone, start_date, end_date, pp_start_date, pp_end_date, cwly_start_date, cwly_end_date,pwly_start_date, pwly_end_date,format_number,
        "search_channel_group", filters={"application": "Desk/MWEB"}, suffix="_Desk/MWEB",
    )
    df_total = calculate_metrics(
        df_pricelince_b2c_standalone, start_date, end_date, pp_start_date, pp_end_date, cwly_start_date, cwly_end_date,pwly_start_date, pwly_end_date, format_number,
        "search_channel_group", suffix="_Total",
    )

    df_channel = pd.concat([df_app, df_desk_mweb, df_total], axis=1)
    return df_channel.reindex(["Direct", "Web Marketing", "Shop PPC", "Affiliate", "Total"]).fillna("")
