import pandas as pd

def build_plan_metrics_period_plan(
    df_finance: pd.DataFrame,
    start_date,
    end_date,
    pp_start_date,
    begin_of_current_year,
    begin_of_last_year,
    ytd_last_year,
    format_number=None,
    period_order=None,
    scenario_value="PLAN",
):
    df = df_finance.copy()

    # defaults
    if period_order is None:
        period_order = ["CW", "PW", "YTD", "YTD_LY"]

    # normalize dates
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    pp_start_date = pd.to_datetime(pp_start_date)
    begin_of_current_year = pd.to_datetime(begin_of_current_year)
    begin_of_last_year = pd.to_datetime(begin_of_last_year)
    ytd_last_year = pd.to_datetime(ytd_last_year)

    # normalize df date cols
    if "year_month" in df.columns:
        df["year_month"] = pd.to_datetime(df["year_month"], errors="coerce")
    if "trans_date" in df.columns:
        df["trans_date"] = pd.to_datetime(df["trans_date"], errors="coerce")

    # --- helper ---
    def grp_sum(mask, col):
        return (
            pd.to_numeric(df.loc[mask, col], errors="coerce")
              .groupby(df.loc[mask, "scenario"])
              .sum()
              .round(0)
        )

    # --- masks ---
    m_cw = df["year_month"] == start_date.normalize()
    m_pw = df["year_month"] == pp_start_date.normalize()

    m_ytd = (df["year_month"] <= end_date.normalize()) & (df["trans_date"] >= begin_of_current_year.normalize())
    m_ytd_ly = (df["year_month"] <= ytd_last_year.normalize()) & (df["trans_date"] >= begin_of_last_year.normalize())

    # --- metrics (Series) ---
    plan_net_cw = grp_sum(m_cw, "net_units")
    plan_net_pw = grp_sum(m_pw, "net_units")
    plan_net_ytd = grp_sum(m_ytd, "net_units")
    plan_net_ytd_ly = grp_sum(m_ytd_ly, "net_units")

    plan_gr_cw = grp_sum(m_cw, "gr_Units")
    plan_gr_pw = grp_sum(m_pw, "gr_Units")
    plan_gr_ytd = grp_sum(m_ytd, "gr_Units")
    plan_gr_ytd_ly = grp_sum(m_ytd_ly, "gr_Units")

    plan_grrev_cw = grp_sum(m_cw, "gr_cont_fee_GDS_Wex")
    plan_grrev_pw = grp_sum(m_pw, "gr_cont_fee_GDS_Wex")
    plan_grrev_ytd = grp_sum(m_ytd, "gr_cont_fee_GDS_Wex")
    plan_grrev_ytd_ly = grp_sum(m_ytd_ly, "gr_cont_fee_GDS_Wex")

    plan_netrev_cw = grp_sum(m_cw, "net_cont_fee_GDS_Wex")
    plan_netrev_pw = grp_sum(m_pw, "net_cont_fee_GDS_Wex")
    plan_netrev_ytd = grp_sum(m_ytd, "net_cont_fee_GDS_Wex")
    plan_netrev_ytd_ly = grp_sum(m_ytd_ly, "net_cont_fee_GDS_Wex")

    # 1) MultiIndex columns (metric, period)
    plan_metrics_mi = pd.concat(
        {
            ("net_tkts", "CW"): plan_net_cw,
            ("net_tkts", "PW"): plan_net_pw,
            ("net_tkts", "YTD"): plan_net_ytd,
            ("net_tkts", "YTD_LY"): plan_net_ytd_ly,

            ("gr_tkts", "CW"): plan_gr_cw,
            ("gr_tkts", "PW"): plan_gr_pw,
            ("gr_tkts", "YTD"): plan_gr_ytd,
            ("gr_tkts", "YTD_LY"): plan_gr_ytd_ly,

            ("gr_rev", "CW"): plan_grrev_cw,
            ("gr_rev", "PW"): plan_grrev_pw,
            ("gr_rev", "YTD"): plan_grrev_ytd,
            ("gr_rev", "YTD_LY"): plan_grrev_ytd_ly,

            ("net_rev", "CW"): plan_netrev_cw,
            ("net_rev", "PW"): plan_netrev_pw,
            ("net_rev", "YTD"): plan_netrev_ytd,
            ("net_rev", "YTD_LY"): plan_netrev_ytd_ly,
        },
        axis=1
    )
    plan_metrics_mi.index.name = "scenario"

    # 2) reshape to "period columns"
    plan_metrics_period_cols = (
        plan_metrics_mi
          .stack(0)  # stack metric level -> rows
          .reset_index()
          .rename(columns={"level_1": "metric"})
    )

    # optional formatting
    if callable(format_number):
        for c in period_order:
            if c in plan_metrics_period_cols.columns:
                plan_metrics_period_cols[c] = pd.to_numeric(plan_metrics_period_cols[c], errors="coerce").round(0)

    keep = ["scenario", "metric"] + [c for c in period_order if c in plan_metrics_period_cols.columns]
    plan_metrics_period = plan_metrics_period_cols[keep]

    # filter PLAN
    plan_metrics_period_plan = plan_metrics_period.loc[plan_metrics_period["scenario"].eq(scenario_value)]

    return plan_metrics_period_plan
