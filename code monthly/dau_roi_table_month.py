import pandas as pd

def calculate_dau_conversion(
    dau_conversion: pd.DataFrame,
    format_percentage,
    start_date,
    end_date,
    pp_start_date,
    pp_end_date,
    cwly_start_date,
    cwly_end_date,
    pwly_start_date,
    pwly_end_date,
    begin_of_current_year,
    begin_of_last_year,
    ytd_last_year
):
    df = dau_conversion.copy()

    # --- normalize date types ---
    for c in ["actual_date", "trans_date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    # normalize input dates
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    pp_start_date = pd.to_datetime(pp_start_date)
    pp_end_date = pd.to_datetime(pp_end_date)
    cwly_start_date = pd.to_datetime(cwly_start_date)
    cwly_end_date = pd.to_datetime(cwly_end_date)
    pwly_start_date = pd.to_datetime(pwly_start_date)
    pwly_end_date = pd.to_datetime(pwly_end_date)
    begin_of_current_year = pd.to_datetime(begin_of_current_year)
    begin_of_last_year = pd.to_datetime(begin_of_last_year)
    ytd_last_year = pd.to_datetime(ytd_last_year)

    # --- helpers ---
    def safe_pct_change(cur, base):
        base = base.replace(0, pd.NA)
        return (cur / base - 1) * 100

    def sum_by_channel(date_col, s, e, col):
        m = df[date_col].between(s, e, inclusive="both")
        srs = df.loc[m].groupby("channel")[col].sum()
        # keep NA-friendly ints for DAU
        if pd.api.types.is_numeric_dtype(srs) and col == "engaged_DAU":
            return srs.astype("Int64")
        return srs

    # --- base frame ---
    channels = sorted(df["channel"].dropna().unique())
    out = pd.DataFrame(index=channels)

    # ranges
    out["DAU"]        = sum_by_channel("actual_date",  start_date,      end_date,      "engaged_DAU")
    out["DAU_pw"]     = sum_by_channel("actual_date",  pp_start_date,   pp_end_date,   "engaged_DAU")
    out["DAU_cwly"]   = sum_by_channel("actual_date",  cwly_start_date, cwly_end_date, "engaged_DAU")
    out["DAU_pwly"]   = sum_by_channel("actual_date",  pwly_start_date, pwly_end_date, "engaged_DAU")

    # YTD (actual_date)
    out["DAU_ytd"]    = sum_by_channel("actual_date", begin_of_current_year, end_date,      "engaged_DAU")
    out["DAU_ytd_ly"] = sum_by_channel("actual_date", begin_of_last_year,    ytd_last_year, "engaged_DAU")

    # totals row
    base_cols = ["DAU","DAU_cwly","DAU_pw","DAU_pwly","DAU_ytd","DAU_ytd_ly"]
    out.loc["Total", base_cols] = out[base_cols].sum(numeric_only=True)

    # YoY
    out["DAU YoY"]     = safe_pct_change(out["DAU"],     out["DAU_cwly"])
    out["DAU YoY_PW"]  = safe_pct_change(out["DAU_pw"],  out["DAU_pwly"])
    out["DAU YoY_YTD"] = safe_pct_change(out["DAU_ytd"], out["DAU_ytd_ly"])

    # formatting
    if callable(format_percentage):
        for c in ["DAU YoY", "DAU YoY_PW", "DAU YoY_YTD"]:
            out[c] = out[c].apply(format_percentage)

    out.index.name = "channel"
    return out
