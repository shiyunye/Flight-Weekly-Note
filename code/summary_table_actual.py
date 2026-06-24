import pandas as pd

def create_finance_number(
    df_weekly,
    df_gds_incentive,
    end_date,
    pp_end_date,
    begin_of_current_year,
    begin_of_last_year,
    ytd_last_year,
    cwly_date,
    pwly_date,
    format_percentage,
    format_number,
):

    w = df_weekly.copy()
    g = df_gds_incentive.copy()

    def pct(cy, ly):
        if cy is None or pd.isna(cy) or ly is None or pd.isna(ly) or ly == 0:
            return pd.NA
        return (float(cy) / float(ly) - 1.0) * 100.0

    def sum_weekly(mask, cols, wk=None, start=None, end=None):
        d = w.loc[mask]
        if wk is not None:    d = d[d["wk_ending"] == wk]
        if start is not None: d = d[d["trans_date"] >= start]
        if end is not None:   d = d[d["trans_date"] <= end]
        return d[cols].sum(numeric_only=True).reindex(cols, fill_value=0)

    def sum_inc(wk=None, start=None, end=None):
        d = g
        if wk is not None:     d = d[d["wk_ending"] == wk]
        if start is not None:  d = d[d["trans_date"] >= start]
        if end is not None:    d = d[d["trans_date"] <= end]
        return d["net_gds_incentives"].sum()

    # ---- masks ----
    m_all = w["company"].astype(str).str.contains("priceline", case=False, na=False)
    m_flt = (
        w["brand"].astype(str).str.contains("priceline", case=False, na=False)
        & w["offer_type"].astype(str).str.contains("flights", case=False, na=False)
    )

    # ---- columns ----
    ALL = [
        "net_tkts_cy","net_tkts_ly","gr_tkts_cy","gr_tkts_ly",
        "net_contribution_cy","net_contribution_ly","gross_contribution_cy","gross_contribution_ly",
        "normalized_net_tickets_cy","normalized_net_tickets_ly","normalized_gross_tickets_cy","normalized_gross_tickets_ly",
        "net_contr_fee_cy","net_contr_fee_ly","gr_contr_fee_cy","gr_contr_fee_ly",
        "net_wex_fee_cy","gr_wex_fee_cy","net_wex_fee_ly","gr_wex_fee_ly",
    ]
    FLT = [
        "net_tkts_cy","net_tkts_ly",
        "net_contr_fee_cy","net_contr_fee_ly","gr_contr_fee_cy","gr_contr_fee_ly",
        "net_wex_fee_cy","gr_wex_fee_cy","net_wex_fee_ly","gr_wex_fee_ly",
    ]

    # ---- totals ----
    cw_all  = sum_weekly(m_all, ALL, wk=end_date)
    pw_all  = sum_weekly(m_all, ALL, wk=pp_end_date)
    ytd_cy_all = sum_weekly(m_all, ALL, start=begin_of_current_year, end=end_date)
    ytd_ly_all = sum_weekly(m_all, ALL, start=begin_of_last_year, end=ytd_last_year)

    cw_flt  = sum_weekly(m_flt, FLT, wk=end_date)
    pw_flt  = sum_weekly(m_flt, FLT, wk=pp_end_date)
    ytd_cy_flt = sum_weekly(m_flt, FLT, start=begin_of_current_year, end=end_date)
    ytd_ly_flt = sum_weekly(m_flt, FLT, start=begin_of_last_year, end=ytd_last_year)

    inc_cw = sum_inc(wk=end_date)
    inc_cwly= sum_inc(wk=cwly_date)
    inc_pwly= sum_inc(wk=pwly_date)
    inc_pw = sum_inc(wk=pp_end_date)
    inc_cy = sum_inc(start=begin_of_current_year, end=end_date)
    inc_ly = sum_inc(start=begin_of_last_year, end=ytd_last_year)

    rows = []

    def add(label, cw, pw, rw, prevw, cy, ly):
        rows.append({
            "Measure": label,
            "CW": cw,
            "PW": pw,
            "Reporting Week": rw,
            "Previous Week": prevw,
            "CY": cy,
            "LY": ly,
            "YTD": pct(cy, ly),
        })

    # ---- simple measures (same pattern) ----
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
            pct(cw_all[cy_col], cw_all[ly_col]),
            pct(pw_all[cy_col], pw_all[ly_col]),
            ytd_cy_all[cy_col],
            ytd_ly_all[cy_col],
        )

    # ---- incentive + vcc ----
    rows.append({
        "Measure": "GDS Incentive",
        "CW": inc_cw, "PW": inc_pw,
        "Reporting Week": pct(inc_cw, inc_cwly), "Previous Week": pct(inc_pw, inc_pwly),
        "CY": inc_cy, "LY": inc_ly,
        "YTD": pct(inc_cy, inc_ly),
    })

    # FIX: net_wex logic (use CY vs LY, not CY vs PW-LY; and YTD use the right fields)
    rows.append({
        "Measure": "VCC Rebate (net_wex_fee)",
        "CW": cw_all["net_wex_fee_cy"], "PW": pw_all["net_wex_fee_cy"],
        "Reporting Week": pct(cw_all["net_wex_fee_cy"], cw_all["net_wex_fee_ly"]),
        "Previous Week": pct(pw_all["net_wex_fee_cy"], pw_all["net_wex_fee_ly"]),
        "CY": ytd_cy_all["net_wex_fee_cy"], "LY": ytd_ly_all["net_wex_fee_ly"],
        "YTD": pct(ytd_cy_all["net_wex_fee_cy"], ytd_ly_all["net_wex_fee_ly"]),
    })

    # ---- + incentives (YTD uses incentives in both CY/LY) ----
    add(
        "Net Cont + Fee (Flight Only)",
        cw_flt["net_contr_fee_cy"],
        pw_flt["net_contr_fee_cy"],
        pct(cw_flt["net_contr_fee_cy"], cw_flt["net_contr_fee_ly"]),
        pct(pw_flt["net_contr_fee_cy"], pw_flt["net_contr_fee_ly"]),
        ytd_cy_flt["net_contr_fee_cy"],
        ytd_ly_flt["net_contr_fee_ly"],
    )

    add(
        "Gross Cont + Fee (Flight Only)",
        cw_flt["gr_contr_fee_cy"],
        pw_flt["gr_contr_fee_cy"],
        pct(cw_flt["gr_contr_fee_cy"], cw_flt["gr_contr_fee_ly"]),
        pct(pw_flt["gr_contr_fee_cy"], pw_flt["gr_contr_fee_ly"]),
        ytd_cy_flt["gr_contr_fee_cy"],
        ytd_ly_flt["gr_contr_fee_ly"],
    )

    # ---- flight-only (+ incentives + vcc rebate) ----
    add(
        "Net Cont + Fee + Incentives + vcc rebate(Flight Only)",
        cw_flt["net_contr_fee_cy"] + inc_cw + cw_flt["net_wex_fee_cy"],
        pw_flt["net_contr_fee_cy"] + inc_pw + pw_flt["net_wex_fee_cy"],
        pct(
            cw_flt["net_contr_fee_cy"] + inc_cw + cw_flt["net_wex_fee_cy"],
            cw_flt["net_contr_fee_ly"] + inc_cwly + cw_flt["net_wex_fee_ly"],
        ),
        pct(
            pw_flt["net_contr_fee_cy"] + inc_pw + pw_flt["net_wex_fee_cy"],
            pw_flt["net_contr_fee_ly"] + inc_pwly + pw_flt["net_wex_fee_ly"],
        ),
        ytd_cy_flt["net_contr_fee_cy"] + inc_cy + ytd_cy_flt["net_wex_fee_cy"],
        ytd_ly_flt["net_contr_fee_cy"] + inc_ly + ytd_ly_flt["net_wex_fee_cy"],
    )

    add(
        "Gross Cont + Fee + Incentives + vcc rebate(Flight Only)",
        cw_flt["gr_contr_fee_cy"] + inc_cw + cw_flt["gr_wex_fee_cy"],
        pw_flt["gr_contr_fee_cy"] + inc_pw + pw_flt["gr_wex_fee_cy"],
        pct(
            cw_flt["gr_contr_fee_cy"] + inc_cw + cw_flt["gr_wex_fee_cy"],
            cw_flt["gr_contr_fee_ly"] + inc_cwly + cw_flt["gr_wex_fee_ly"],
        ),
        pct(
            pw_flt["gr_contr_fee_cy"] + inc_pw + pw_flt["gr_wex_fee_cy"],
            pw_flt["gr_contr_fee_ly"] + inc_pwly + pw_flt["gr_wex_fee_ly"],
        ),
        ytd_cy_flt["gr_contr_fee_cy"] + inc_cy + ytd_cy_flt["gr_wex_fee_cy"],
        ytd_ly_flt["gr_contr_fee_cy"] + inc_ly + ytd_ly_flt["gr_wex_fee_cy"],
    )

    out = pd.DataFrame(rows)

    # ---- formatting----
    num_cols = ["CW", "PW", "CY", "LY"]
    pct_cols = ["Reporting Week", "Previous Week", "YTD"]

    for c in num_cols:
        out[c] = pd.to_numeric(out[c], errors="coerce").round(0)

    for c in pct_cols:
        out[c] = pd.to_numeric(out[c], errors="coerce")

    if callable(format_number):
        for c in num_cols:
            out[c] = out[c].round(0)

    if callable(format_percentage):
        for c in pct_cols:
            out[c] = out[c].apply(format_percentage)

    return out
