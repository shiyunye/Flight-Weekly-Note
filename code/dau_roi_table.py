import pandas as pd

def calculate_dau_conversion(
    dau_conversion: pd.DataFrame,
    format_percentage,            
    end_date,
    pp_end_date,
    cwly_date,
    pwly_date,
    begin_of_current_year,
    begin_of_last_year,
    ytd_last_year,               
):
    df = dau_conversion.copy()

    # --- normalize date types ---
    for c in ["week_ending", "actual_date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    end_date = pd.to_datetime(end_date)
    pp_end_date = pd.to_datetime(pp_end_date)
    cwly_date = pd.to_datetime(cwly_date)
    pwly_date = pd.to_datetime(pwly_date)
    begin_of_current_year = pd.to_datetime(begin_of_current_year)
    begin_of_last_year = pd.to_datetime(begin_of_last_year)
    ytd_last_year = pd.to_datetime(ytd_last_year)

    # --- helpers ---
    def safe_pct_change(cur, base):
        base = base.replace(0, pd.NA)
        return (cur / base - 1) * 100

    def sum_by_channel(mask, col):
        s = df.loc[mask].groupby("channel")[col].sum()
        # keep Int64 so NaN is allowed
        if pd.api.types.is_numeric_dtype(s):
            return s.astype("Int64") if col == "engaged_DAU" else s
        return s

    # --- build base frame with DAU / converted ---
    df_dau_converison = pd.DataFrame(index=sorted(df["channel"].dropna().unique()))

    # weekly masks
    m_cw   = df["week_ending"].eq(end_date)
    m_pw   = df["week_ending"].eq(pp_end_date)
    m_cwly = df["week_ending"].eq(cwly_date)
    m_pwly = df["week_ending"].eq(pwly_date)

    # ytd masks (use actual_date)
    m_ytd_cy = df["actual_date"].between(begin_of_current_year, end_date, inclusive="both")
    m_ytd_ly = df["actual_date"].between(begin_of_last_year, ytd_last_year, inclusive="both")

    # DAU
    df_dau_converison["DAU"]        = sum_by_channel(m_cw,   "engaged_DAU")
    df_dau_converison["DAU_cwly"]   = sum_by_channel(m_cwly, "engaged_DAU")
    df_dau_converison["DAU_pw"]     = sum_by_channel(m_pw,   "engaged_DAU")
    df_dau_converison["DAU_pwly"]   = sum_by_channel(m_pwly, "engaged_DAU")
    df_dau_converison["DAU_ytd"]    = sum_by_channel(m_ytd_cy, "engaged_DAU")
    df_dau_converison["DAU_ytd_ly"] = sum_by_channel(m_ytd_ly, "engaged_DAU")

    # converted
    df_dau_converison["converted"]      = sum_by_channel(m_cw,   "converted")
    df_dau_converison["converted_cwly"] = sum_by_channel(m_cwly, "converted")
    df_dau_converison["converted_pw"]   = sum_by_channel(m_pw,   "converted")
    df_dau_converison["converted_pwly"] = sum_by_channel(m_pwly, "converted")

    # --- totals row (after all base cols exist) ---
    base_cols = [
        "DAU","DAU_cwly","DAU_pw","DAU_pwly","DAU_ytd","DAU_ytd_ly",
        "converted","converted_cwly","converted_pw","converted_pwly",
    ]
    df_dau_converison.loc["Total", base_cols] = df_dau_converison[base_cols].sum(numeric_only=True)

    # --- YoY on converted ---
    df_dau_converison["converted YoY"]      = safe_pct_change(df_dau_converison["converted"],     df_dau_converison["converted_cwly"])
    df_dau_converison["converted YoY_PW"]   = safe_pct_change(df_dau_converison["converted_pw"],  df_dau_converison["converted_pwly"])

    # --- YoY on DAU ---
    df_dau_converison["DAU YoY"]      = safe_pct_change(df_dau_converison["DAU"],     df_dau_converison["DAU_cwly"])
    df_dau_converison["DAU YoY_PW"]   = safe_pct_change(df_dau_converison["DAU_pw"],  df_dau_converison["DAU_pwly"])
    df_dau_converison["DAU YoY_YTD"]  = safe_pct_change(df_dau_converison["DAU_ytd"], df_dau_converison["DAU_ytd_ly"])


    # --- conversion rates ---
    df_dau_converison["conversion"]       = df_dau_converison["converted"]      / df_dau_converison["DAU"].replace(0, pd.NA)
    df_dau_converison["conversion_cwly"]  = df_dau_converison["converted_cwly"] / df_dau_converison["DAU_cwly"].replace(0, pd.NA)
    df_dau_converison["conversion_pw"]    = df_dau_converison["converted_pw"]   / df_dau_converison["DAU_pw"].replace(0, pd.NA)
    df_dau_converison["conversion_pwly"]  = df_dau_converison["converted_pwly"] / df_dau_converison["DAU_pwly"].replace(0, pd.NA)

    # --- YoY on conversion ---
    df_dau_converison["conversion_YoY"] = safe_pct_change(df_dau_converison["conversion"],    df_dau_converison["conversion_cwly"])
    df_dau_converison["conversion_YoY_PW"] = safe_pct_change(df_dau_converison["conversion_pw"], df_dau_converison["conversion_pwly"])

    # --- formatting ---
    if callable(format_percentage):
        pct_cols = ["DAU YoY","DAU YoY_PW","DAU YoY_YTD","conversion_YoY","conversion_YoY_PW"]
        for c in pct_cols:
            df_dau_converison[c] = df_dau_converison[c].apply(format_percentage)
    df_dau_converison.index.name = "channel" 
    
    return df_dau_converison



def calculate_roi(df_roi, end_date, pp_end_date, cwly_date, pwly_date):
 

    df_roi = df_roi.copy()
    df_roi["week_ending"] = pd.to_datetime(df_roi["week_ending"], errors="coerce")

    end_date    = pd.to_datetime(end_date)
    pp_end_date = pd.to_datetime(pp_end_date)
    cwly_date   = pd.to_datetime(cwly_date)
    pwly_date   = pd.to_datetime(pwly_date)

    def sum_by_channel_roi(dt, col):
        m = df_roi["week_ending"].eq(dt)
        return df_roi.loc[m].groupby("channel")[col].sum()

    df_roi_section = pd.DataFrame()

    # contribution
    df_roi_section["contribution"]      = sum_by_channel_roi(end_date,     "contribution")
    df_roi_section["contribution_cwly"] = sum_by_channel_roi(cwly_date,    "contribution")
    df_roi_section["contribution_pw"]   = sum_by_channel_roi(pp_end_date,  "contribution")
    df_roi_section["contribution_pwly"] = sum_by_channel_roi(pwly_date,    "contribution")

    # cost
    df_roi_section["cost"]      = sum_by_channel_roi(end_date,     "cost")
    df_roi_section["cost_cwly"] = sum_by_channel_roi(cwly_date,    "cost")
    df_roi_section["cost_pw"]   = sum_by_channel_roi(pp_end_date,  "cost")
    df_roi_section["cost_pwly"] = sum_by_channel_roi(pwly_date,    "cost")

    # ROI
    df_roi_section["ROI"]      = df_roi_section["contribution"]      / df_roi_section["cost"].replace(0, pd.NA)
    df_roi_section["ROI_cwly"] = df_roi_section["contribution_cwly"] / df_roi_section["cost_cwly"].replace(0, pd.NA)
    df_roi_section["ROI_pw"]   = df_roi_section["contribution_pw"]   / df_roi_section["cost_pw"].replace(0, pd.NA)
    df_roi_section["ROI_pwly"] = df_roi_section["contribution_pwly"] / df_roi_section["cost_pwly"].replace(0, pd.NA)

    # ROI YoY (% numeric)
    df_roi_section["ROI_YoY"]    = (df_roi_section["ROI"]    / df_roi_section["ROI_cwly"].replace(0, pd.NA) - 1) * 100
    df_roi_section["ROI_YoY_PW"] = (df_roi_section["ROI_pw"] / df_roi_section["ROI_pwly"].replace(0, pd.NA) - 1) * 100

    # totals row
    df_roi_section.loc["Total", df_roi_section.columns] = df_roi_section.sum(numeric_only=True)

    return df_roi_section


def create_roi_table(df_roi_section, df_dau_converison, format_number):
    # rename ROI channels to match DAU conversion channels
    df_roi_section2 = df_roi_section.rename(index={
        "CHEAPFLIGHTS": "Shop PPC Cheapflights",
        "CJ": "Affiliate",
        "CLICKTRIPZ": "Shop PPC Others",
        "SKYSCANNER": "Shop PPC Skyscanner",
        "Kayak": "Shop PPC Kayak",
    })

    dau = df_dau_converison.reset_index()                 # has: channel + DAU/conv cols
    roi = df_roi_section2.reset_index().rename(columns={"index": "channel"})  # has: channel + ROI cols

    merged = dau.merge(roi, on="channel", how="left").set_index("channel")

    df_roi_v = merged[[
        "DAU", "DAU YoY", "DAU YoY_PW",
        "conversion", "conversion_YoY", "conversion_YoY_PW",
        "ROI", "ROI_YoY", "ROI_YoY_PW",
    ]].copy()

    # IMPORTANT: avoid duplicate col names
    df_roi_v.columns = [
        "DAU", "DAU_YoY", "DAU_YoY_PW",
        "Conversion", "Conversion_YoY", "Conversion_YoY_PW",
        "ROI", "ROI_YoY", "ROI_YoY_PW",
    ]

    # format DAU / ROI (keep conversion numeric for now)
    df_roi_v["DAU"] = df_roi_v["DAU"].apply(format_number)
    df_roi_v["ROI"] = df_roi_v["ROI"].apply(lambda x: "" if pd.isna(x) else f"{x:.2f}")

    order = [
        "Direct", "SEM Core", "SEM Brand",
        "Shop PPC Cheapflights", "Shop PPC Google", "Shop PPC Kayak",
        "Shop PPC Skyscanner", "Shop PPC Others",
        "Affiliate", "Total",
    ]
    df_roi_v = df_roi_v.reindex(order)

    return df_roi_v
