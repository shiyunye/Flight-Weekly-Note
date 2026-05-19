import pandas as pd

def create_finance_number_month(
    df_monthly,
    df_gds_incentive,
    start_date, end_date,
    pp_start_date, pp_end_date,
    cwly_start_date, cwly_end_date,
    pwly_start_date, pwly_end_date,
    begin_of_current_year,
    begin_of_last_year,
    ytd_last_year,
    format_percentage,
    format_number,
):
    w = df_monthly.copy()
    g = df_gds_incentive.copy()

    # --- normalize date types ---
    for d, cols in [(w, ["trans_date"]), (g, ["trans_date"])]:
        for c in cols:
            if c in d.columns:
                d[c] = pd.to_datetime(d[c], errors="coerce")

    # normalize inputs
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

    # ------
    def pct(cy, ly):
        if cy is None or pd.isna(cy) or ly is None or pd.isna(ly) or ly == 0:
            return pd.NA
        return (float(cy) / float(ly) - 1.0) * 100.0

    def sum_period(mask, cols, start=None, end=None):
        d = w.loc[mask]
        if start is not None:
            d = d[d["trans_date"] >= start]
        if end is not None:
            d = d[d["trans_date"] <= end]
        return d[cols].sum(numeric_only=True).reindex(cols, fill_value=0)

    def sum_inc(start=None, end=None):
        d = g
        if start is not None:
            d = d[d["trans_date"] >= start]
        if end is not None:
            d = d[d["trans_date"] <= end]
        return pd.to_numeric(d["net_gds_incentives"], errors="coerce").sum()

    # ---- masks ----
    m_all = w["company"].astype(str).str.contains("priceline", case=False, na=False)
    m_flt = (
        w["brand"].astype(str).str.contains("priceline", case=False, na=False)
        & w["offer_type"].astype(str).str.contains("flights", case=False, na=False)
    )

    # ---- columns ----
    ALL = [
        "net_tkts_cy","gr_tkts_cy",
        "net_contribution_cy","gross_contribution_cy",
        "normalized_net_tickets_cy","normalized_gross_tickets_cy",
        "net_contr_fee_cy","gr_contr_fee_cy",
        "net_wex_fee_cy","gr_wex_fee_cy"
    ]
    FLT = [
        "net_tkts_cy",
        "net_contr_fee_cy","gr_contr_fee_cy",
        "net_wex_fee_cy","gr_wex_fee_cy",
    ]

    # ---- totals (month ranges) ----
    cw_all   = sum_period(m_all, ALL, start=start_date, end=end_date)
    pw_all   = sum_period(m_all, ALL, start=pp_start_date, end=pp_end_date)
    cwly_all = sum_period(m_all, ALL, start=cwly_start_date, end=cwly_end_date)
    pwly_all = sum_period(m_all, ALL, start=pwly_start_date, end=pwly_end_date)

    ytd_cy_all = sum_period(m_all, ALL, start=begin_of_current_year, end=end_date)
    ytd_ly_all = sum_period(m_all, ALL, start=begin_of_last_year, end=ytd_last_year)

    cw_flt   = sum_period(m_flt, FLT, start=start_date, end=end_date)
    pw_flt   = sum_period(m_flt, FLT, start=pp_start_date, end=pp_end_date)
    cwly_flt = sum_period(m_flt, FLT, start=cwly_start_date, end=cwly_end_date)
    pwly_flt = sum_period(m_flt, FLT, start=pwly_start_date, end=pwly_end_date)

    ytd_cy_flt = sum_period(m_flt, FLT, start=begin_of_current_year, end=end_date)
    ytd_ly_flt = sum_period(m_flt, FLT, start=begin_of_last_year, end=ytd_last_year)

    # ---- incentives (month ranges) ----
    inc_cw   = sum_inc(start=start_date, end=end_date)
    inc_pw   = sum_inc(start=pp_start_date, end=pp_end_date)
    inc_cwly = sum_inc(start=cwly_start_date, end=cwly_end_date)
    inc_pwly = sum_inc(start=pwly_start_date, end=pwly_end_date)
    inc_cy   = sum_inc(start=begin_of_current_year, end=end_date)
    inc_ly   = sum_inc(start=begin_of_last_year, end=ytd_last_year)

    rows = []

    def add(label, cw, pw, cwly, pwly, rw, prevw, cy, ly):
        rows.append({
            "Measure": label,
            "CW": cw,
            "PW": pw,
            "CWly": cwly,
            "pWly": pwly,
            "Reporting Week": rw,
            "Previous Week": prevw,
            "CY": cy,
            "LY": ly,
            "YTD": pct(cy, ly),
        })

    # ---- simple measures ----
    simple = [
        ("Net Tickets", "net_tkts_cy", "net_tkts_ly"),
        ("Gross Tickets", "gr_tkts_cy", "gr_tkts_ly"),
        ("Net Revenue (net_contribution)", "net_contribution_cy", "net_contribution_ly"),
        ("Gross Revenue (gross_contribution)", "gross_contribution_cy", "gross_contribution_ly"),
        ("Normalized Net Tickets", "normalized_net_tickets_cy", "normalized_net_tickets_ly"),
        ("Normalized Gross Tickets", "normalized_gross_tickets_cy", "normalized_gross_tickets_ly"),
        ("Net Cont + Fee", "net_contr_fee_cy", "net_contr_fee_ly"),
        ("Gross Cont + Fee", "gr_contr_fee_cy", "gr_contr_fee_ly"),
    ]

    for label, cy_col, ly_col in simple:
        add(
            label,
            cw_all[cy_col],
            pw_all[cy_col],
            cwly_all[cy_col],
            pwly_all[cy_col],
            pct(cw_all[cy_col], cwly_all[cy_col]),
            pct(pw_all[cy_col], pwly_all[cy_col]),
            ytd_cy_all[cy_col],
            ytd_ly_all[cy_col],
        )

    # ---- GDS Incentive (YoY) ----
    add(
        "GDS Incentive",
        inc_cw,
        inc_pw,
        inc_cwly,
        inc_pwly,
        pct(inc_cw, inc_cwly),
        pct(inc_pw, inc_pwly),
        inc_cy,
        inc_ly,
    )

    # ---- VCC Rebate (net_wex_fee) ----
    add(
        "VCC Rebate (net_wex_fee)",
        cw_all["net_wex_fee_cy"],
        pw_all["net_wex_fee_cy"],
        cwly_all["net_wex_fee_cy"],
        pwly_all["net_wex_fee_cy"],
        pct(cw_all["net_wex_fee_cy"], cwly_all["net_wex_fee_cy"]),
        pct(pw_all["net_wex_fee_cy"], pwly_all["net_wex_fee_cy"]),
        ytd_cy_all["net_wex_fee_cy"],
        ytd_ly_all["net_wex_fee_cy"],
    )

    # ---- flight-only measures ----
    add(
        "Net Cont + Fee (Flight Only)",
        cw_flt["net_contr_fee_cy"],
        pw_flt["net_contr_fee_cy"],
        cwly_flt["net_contr_fee_cy"],
        pwly_flt["net_contr_fee_cy"],
        pct(cw_flt["net_contr_fee_cy"], cwly_flt["net_contr_fee_cy"]),
        pct(pw_flt["net_contr_fee_cy"], pwly_flt["net_contr_fee_cy"]),
        ytd_cy_flt["net_contr_fee_cy"],
        ytd_ly_flt["net_contr_fee_cy"],
    )

    add(
        "Gross Cont + Fee (Flight Only)",
        cw_flt["gr_contr_fee_cy"],
        pw_flt["gr_contr_fee_cy"],
        cwly_flt["gr_contr_fee_cy"],
        pwly_flt["gr_contr_fee_cy"],
        pct(cw_flt["gr_contr_fee_cy"], cwly_flt["gr_contr_fee_cy"]),
        pct(pw_flt["gr_contr_fee_cy"], pwly_flt["gr_contr_fee_cy"]),
        ytd_cy_flt["gr_contr_fee_cy"],
        ytd_ly_flt["gr_contr_fee_cy"],
    )

    # ---- flight-only (+ incentives + vcc rebate) ----
    add(
        "Net Cont + Fee + Incentives + vcc rebate(Flight Only)",
        cw_flt["net_contr_fee_cy"] + inc_cw + cw_flt["net_wex_fee_cy"],
        pw_flt["net_contr_fee_cy"] + inc_pw + pw_flt["net_wex_fee_cy"],
        cwly_flt["net_contr_fee_cy"] + inc_cwly + cwly_flt["net_wex_fee_cy"],
        pwly_flt["net_contr_fee_cy"] + inc_pwly + pwly_flt["net_wex_fee_cy"],
        pct(
            cw_flt["net_contr_fee_cy"] + inc_cw + cw_flt["net_wex_fee_cy"],
            cwly_flt["net_contr_fee_cy"] + inc_cwly + cwly_flt["net_wex_fee_cy"],
        ),
        pct(
            pw_flt["net_contr_fee_cy"] + inc_pw + pw_flt["net_wex_fee_cy"],
            pwly_flt["net_contr_fee_cy"] + inc_pwly + pwly_flt["net_wex_fee_cy"],
        ),
        ytd_cy_flt["net_contr_fee_cy"] + inc_cy + ytd_cy_flt["net_wex_fee_cy"],
        ytd_ly_flt["net_contr_fee_cy"] + inc_ly + ytd_ly_flt["net_wex_fee_cy"],
    )

    add(
        "Gross Cont + Fee + Incentives + vcc rebate(Flight Only)",
        cw_flt["gr_contr_fee_cy"] + inc_cw + cw_flt["gr_wex_fee_cy"],
        pw_flt["gr_contr_fee_cy"] + inc_pw + pw_flt["gr_wex_fee_cy"],
        cwly_flt["gr_contr_fee_cy"] + inc_cwly + cwly_flt["gr_wex_fee_cy"],
        pwly_flt["gr_contr_fee_cy"] + inc_pwly + pwly_flt["gr_wex_fee_cy"],
        pct(
            cw_flt["gr_contr_fee_cy"] + inc_cw + cw_flt["gr_wex_fee_cy"],
            cwly_flt["gr_contr_fee_cy"] + inc_cwly + cwly_flt["gr_wex_fee_cy"],
        ),
        pct(
            pw_flt["gr_contr_fee_cy"] + inc_pw + pw_flt["gr_wex_fee_cy"],
            pwly_flt["gr_contr_fee_cy"] + inc_pwly + pwly_flt["gr_wex_fee_cy"],
        ),
        ytd_cy_flt["gr_contr_fee_cy"] + inc_cy + ytd_cy_flt["gr_wex_fee_cy"],
        ytd_ly_flt["gr_contr_fee_cy"] + inc_ly + ytd_ly_flt["gr_wex_fee_cy"],
    )

    out = pd.DataFrame(rows)

    for c in ["CW", "PW", "CWly", "pWly", "CY", "LY"]:
        out[c] = pd.to_numeric(out[c], errors="coerce").round(0)

    return out
