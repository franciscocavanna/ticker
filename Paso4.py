#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Paso4 - Scorecard & Valuation (con comparación multi-período en la hoja Scorecard)
----------------------------------------------------------------------------------
- Calcula Scorecard del último año (F-Score, Altman Z/Z', ROIC, WACC, DCF, Upside, Decision).
- Agrega un bloque "Compare_Periods (últimos 5)" en la MISMA hoja Scorecard con:
  ROIC %, WACC %, Spread, F-Score, Altman Z/Z', FCF, Gross Margin %, Net Margin %.
- También exporta hojas Timeline_Ratios, Timeline_Scores, Timeline_FCF, Timeline_Base.
- Incluye paneles DCF: DCF_WACC_vs_TermG y DCF_WACC_vs_G5y.

Uso:
    pip install --upgrade pandas numpy openpyxl yfinance
    python Paso4.py --key GLOB_key.xlsx --full GLOB_full.xlsx --ticker GLOB --out GLOB_scorecard.xlsx
"""

import argparse
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import numpy as np
import pandas as pd

try:
    import yfinance as yf
except Exception:
    yf = None


# ------------------ Utilidades ------------------
def _read_sheet(path: Path, name: str, index_col=0) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=name, index_col=index_col)
    except ValueError:
        df = pd.read_excel(path, sheet_name=name[:31], index_col=index_col)
    except Exception:
        return pd.DataFrame()
    if df is None or df.empty:
        return pd.DataFrame()
    cols = []
    for c in df.columns:
        try:
            cols.append(pd.to_datetime(c).strftime("%Y-%m"))
        except Exception:
            cols.append(str(c))
    df.columns = cols
    df.index = [str(i).strip() for i in df.index]
    return df


def _latest_col(df: pd.DataFrame) -> Optional[str]:
    if df is None or df.empty:
        return None
    try:
        return sorted(df.columns, key=lambda x: pd.to_datetime(x, errors="coerce"))[-1]
    except Exception:
        return df.columns[-1]


def _prev_col(df: pd.DataFrame, col: str) -> Optional[str]:
    if df is None or df.empty or col not in df.columns:
        return None
    try:
        cols = sorted(df.columns, key=lambda x: pd.to_datetime(x, errors="coerce"))
    except Exception:
        cols = list(df.columns)
    i = cols.index(col)
    return cols[i-1] if i > 0 else None


def _to_sorted_cols(df: pd.DataFrame) -> list:
    try:
        return sorted(df.columns, key=lambda x: pd.to_datetime(x, errors="coerce"))
    except Exception:
        return list(df.columns)


def _safe(v):
    try:
        return float(v)
    except Exception:
        return np.nan


def _row(df: pd.DataFrame, name: str) -> pd.Series:
    return df.loc[name] if (df is not None and name in df.index) else pd.Series(dtype="float64")


def _safe_div(a, b):
    return (a / b).replace([np.inf, -np.inf], np.nan)


def _pick_from_full(full_path: Optional[Path], sheet: str, candidates: List[str]) -> pd.Series:
    if full_path is None:
        return pd.Series(dtype="float64")
    df = _read_sheet(full_path, sheet)
    if df.empty:
        return pd.Series(dtype="float64")
    idx = {str(i).lower(): i for i in df.index}
    for c in candidates:
        if c.lower() in idx:
            return pd.to_numeric(df.loc[idx[c.lower()]], errors="coerce")
    for c in candidates:
        for low, real in idx.items():
            if c.lower() in low:
                return pd.to_numeric(df.loc[real], errors="coerce")
    return pd.Series(dtype="float64")


# ------------------ F-Score (último año) ------------------
def piotroski_f_score_last(key_inc_a: pd.DataFrame, key_bal_a: pd.DataFrame, key_cf_a: pd.DataFrame,
                           full_bal_a: Optional[Path], full_cf_a: Optional[Path]) -> Tuple[float, Dict[str, int]]:
    col = _latest_col(key_inc_a)
    prev = _prev_col(key_inc_a, col)
    if not col or not prev:
        return np.nan, {}

    sales = _row(key_inc_a, "Revenue")
    net = _row(key_inc_a, "Net Income")
    assets = _row(key_bal_a, "Total Assets")
    ocf = _row(key_cf_a, "Operating Cash Flow")
    total_debt = _row(key_bal_a, "Total Debt")

    cur_assets = _pick_from_full(full_bal_a, "Balance Sheet (Annual)", ["Total Current Assets", "Current Assets"])
    cur_liab   = _pick_from_full(full_bal_a, "Balance Sheet (Annual)", ["Total Current Liabilities", "Current Liabilities"])
    issued_stock = _pick_from_full(full_cf_a, "Cash Flow (Annual)", [
        "Issuance Of Stock", "Common Stock Issued",
        "Sale Purchase Of Stock", "Sale Purchase Of Common And Preferred Stock"
    ])

    s: Dict[str, int] = {}
    # ROA > 0
    roa_t = _safe(net.get(col, np.nan)) / _safe(assets.get(col, np.nan))
    roa_p = _safe(net.get(prev, np.nan)) / _safe(assets.get(prev, np.nan))
    s["ROA > 0"] = int(roa_t > 0)
    # CFO > 0
    s["CFO > 0"] = int(_safe(ocf.get(col, np.nan)) > 0)
    # ΔROA > 0
    s["ΔROA > 0"] = int(roa_t > roa_p)
    # Accruals
    s["Accruals (CFO > NI)"] = int(_safe(ocf.get(col, np.nan)) > _safe(net.get(col, np.nan)))
    # Deuda ↓
    s["↓ Leverage (Debt)"] = int((_safe(total_debt.get(col, np.nan)) - _safe(total_debt.get(prev, np.nan))) < 0)
    # Current Ratio ↑
    try:
        cr_t = _safe(cur_assets.get(col, np.nan)) / _safe(cur_liab.get(col, np.nan))
        cr_p = _safe(cur_assets.get(prev, np.nan)) / _safe(cur_liab.get(prev, np.nan))
        s["↑ Current Ratio"] = int(cr_t > cr_p and cr_t > 0 and cr_p > 0)
    except Exception:
        s["↑ Current Ratio"] = 0
    # No new shares
    try:
        s["No new shares"] = int(_safe(issued_stock.get(col, 0)) <= 0)
    except Exception:
        s["No new shares"] = 0
    # Márgenes ↑
    try:
        gp = _row(key_inc_a, "Gross Profit")
        gm_t = _safe(gp.get(col, np.nan)) / _safe(sales.get(col, np.nan))
        gm_p = _safe(gp.get(prev, np.nan)) / _safe(sales.get(prev, np.nan))
        s["↑ Gross Margin"] = int(gm_t > gm_p)
    except Exception:
        s["↑ Gross Margin"] = 0
    # Asset Turnover ↑
    try:
        at_t = _safe(sales.get(col, np.nan)) / _safe(assets.get(col, np.nan))
        at_p = _safe(sales.get(prev, np.nan)) / _safe(assets.get(prev, np.nan))
        s["↑ Asset Turnover"] = int(at_t > at_p)
    except Exception:
        s["↑ Asset Turnover"] = 0

    return float(sum(s.values())), s


# ------------------ Altman Z/Z' (último año) ------------------
def altman_z_last(full_bal_a: Path, key_inc_a: pd.DataFrame, ticker: Optional[str]) -> Tuple[float, Dict[str, float]]:
    bal = _read_sheet(full_bal_a, "Balance Sheet (Annual)")
    if bal.empty:
        return np.nan, {}

    def pick(name_list: List[str]) -> pd.Series:
        idx = {str(i).lower(): i for i in bal.index}
        for n in name_list:
            if n.lower() in idx:
                return pd.to_numeric(bal.loc[idx[n.lower()]], errors="coerce")
        for n in name_list:
            for low, real in idx.items():
                if n.lower() in low:
                    return pd.to_numeric(bal.loc[real], errors="coerce")
        return pd.Series(dtype="float64")

    col = _latest_col(bal)
    total_assets = pick(["Total Assets"])
    total_liab   = pick(["Total Liab", "Total Liabilities"])
    cur_assets   = pick(["Total Current Assets", "Current Assets"])
    cur_liab     = pick(["Total Current Liabilities", "Current Liabilities"])
    wc = cur_assets - cur_liab
    retained = pick(["Retained Earnings", "Accumulated Retained Earnings Deficit", "Retained Earnings Accumulated Deficit"])
    ebit = _row(key_inc_a, "Operating Income (EBIT)")
    sales = _row(key_inc_a, "Revenue")

    # Market Cap
    mcap = np.nan
    if ticker and yf is not None:
        try:
            tk = yf.Ticker(ticker)
            px = float(getattr(tk, "fast_info", {}).get("last_price") or tk.history(period="1d")["Close"].iloc[-1])
            sh = float((tk.get_shares_full() or pd.Series([np.nan])).iloc[-1]) if hasattr(tk, "get_shares_full") else float(tk.info.get("sharesOutstanding", np.nan))
            mcap = px * sh if np.isfinite(px) and np.isfinite(sh) else np.nan
        except Exception:
            pass

    A = _safe(wc.get(col, np.nan)) / _safe(total_assets.get(col, np.nan))
    B = _safe(retained.get(col, np.nan)) / _safe(total_assets.get(col, np.nan))
    C = _safe(ebit.get(col, np.nan)) / _safe(total_assets.get(col, np.nan))
    E = _safe(sales.get(col, np.nan)) / _safe(total_assets.get(col, np.nan))

    if np.isfinite(mcap):
        D = _safe(mcap) / _safe(total_liab.get(col, np.nan))
        Z = 1.2*A + 1.4*B + 3.3*C + 0.6*D + 1.0*E
        return Z, {"A": A, "B": B, "C": C, "D": D, "E": E}
    else:
        book_equity = pick(["Total Stockholder Equity", "Total Shareholder Equity", "Stockholders Equity"])
        Dp = _safe(book_equity.get(col, np.nan)) / _safe(total_liab.get(col, np.nan))
        Zp = 0.717*A + 0.847*B + 3.107*C + 0.420*Dp + 0.998*E
        return Zp, {"A": A, "B": B, "C": C, "D': book_equity/TL": Dp, "E": E}


# ------------------ ROIC/WACC/DCF (último año) ------------------
def roic_wacc_and_dcf_last(key_inc_a: pd.DataFrame, key_bal_a: pd.DataFrame, key_cf_a: pd.DataFrame,
                           ticker: Optional[str], full_path: Optional[Path],
                           risk_free: float = 0.045, mrp: float = 0.05, term_growth: float = 0.02, years: int = 5
                           ) -> Tuple[Dict[str, float], Dict[str, float]]:
    col = _latest_col(key_inc_a)
    prev = _prev_col(key_inc_a, col)

    ebit = _row(key_inc_a, "Operating Income (EBIT)")
    tax_exp = _row(key_inc_a, "Tax Expense")
    pretax = _row(key_inc_a, "Pretax Income")
    try:
        tax_rate = max(0.0, min(0.40, _safe(tax_exp.get(col, np.nan)) / _safe(pretax.get(col, np.nan))))
    except Exception:
        tax_rate = 0.25
    nopat = _safe(ebit.get(col, np.nan)) * (1 - (tax_rate if np.isfinite(tax_rate) else 0.25))

    assets = _row(key_bal_a, "Total Assets")
    cash = _row(key_bal_a, "Cash & Equivalents")
    invested_capital = _safe(assets.get(col, np.nan)) - _safe(cash.get(col, np.nan))
    roic = (nopat / invested_capital)*100 if invested_capital and np.isfinite(invested_capital) and invested_capital != 0 else np.nan

    beta = np.nan
    price = np.nan
    shares = np.nan
    if ticker and yf is not None:
        try:
            tk = yf.Ticker(ticker)
            price = float(getattr(tk, "fast_info", {}).get("last_price") or tk.history(period="1d")["Close"].iloc[-1])
            beta = float(tk.info.get("beta", np.nan))
            shares = float((tk.get_shares_full() or pd.Series([np.nan])).iloc[-1]) if hasattr(tk, "get_shares_full") else float(tk.info.get("sharesOutstanding", np.nan))
        except Exception:
            pass

    if (not np.isfinite(shares)) and full_path is not None:
        sh_series = _pick_from_full(full_path, "Income Statement (Annual)", [
            "Diluted Average Shares", "Basic Average Shares",
            "Weighted Average Shares", "Weighted Average Shares Diluted",
            "Weighted Average Shares Outstanding Diluted",
        ])
        if not sh_series.empty:
            val = sh_series.get(col, np.nan)
            if not np.isfinite(val):
                nz = sh_series.dropna()
                val = nz.iloc[-1] if not nz.empty else np.nan
            shares = float(val) if np.isfinite(val) else shares

    ke = risk_free + (beta if np.isfinite(beta) else 1.0)*mrp

    total_debt = _row(key_bal_a, "Total Debt")
    interest = _pick_from_full(full_path, "Income Statement (Annual)", [
        "Interest Expense", "Interest Expense Non Operating", "Non Operating Interest Expense"
    ])
    try:
        kd = abs(_safe(interest.get(col, np.nan))) / max(_safe(total_debt.get(col, np.nan)), 1e-9)
    except Exception:
        kd = 0.06
    kd_after_tax = kd * (1 - (tax_rate if np.isfinite(tax_rate) else 0.25))

    equity_value = (price*shares) if np.isfinite(price) and np.isfinite(shares) else _safe(_row(key_bal_a, "Total Equity").get(col, np.nan))
    debt_value = _safe(total_debt.get(col, np.nan))
    ev_total = equity_value + debt_value if np.isfinite(equity_value) and np.isfinite(debt_value) else np.nan
    we = equity_value / ev_total if np.isfinite(ev_total) and ev_total != 0 else 0.6
    wd = debt_value / ev_total if np.isfinite(ev_total) and ev_total != 0 else 0.4
    wacc = we*ke + wd*kd_after_tax

    fcf = _row(key_cf_a, "Free Cash Flow (OCF - CapEx)")
    fcf0 = _safe(fcf.get(col, np.nan))

    rev = _row(key_inc_a, "Revenue")
    r2 = _safe(rev.get(col, np.nan))
    r1 = _safe(rev.get(prev, np.nan))
    g = (r2/r1 - 1) if np.isfinite(r2) and np.isfinite(r1) and r1 != 0 else 0.04
    g = float(np.clip(g, -0.05, 0.10))

    disc = (1 + wacc) if np.isfinite(wacc) and wacc > 0 else 1.0
    pv = 0.0
    f = fcf0
    for t in range(1, years+1):
        f *= (1 + g)
        pv += f / (disc**t)
    tv = f * (1 + term_growth) / (wacc - term_growth) if (np.isfinite(wacc) and wacc > term_growth) else np.nan
    if np.isfinite(tv):
        pv += tv / (disc**years)

    net_debt = _safe(total_debt.get(col, np.nan)) - _safe(cash.get(col, np.nan))
    equity_dcf = pv - net_debt if np.isfinite(pv) and np.isfinite(net_debt) else np.nan
    fv_per_share = (equity_dcf / shares) if (np.isfinite(equity_dcf) and np.isfinite(shares) and shares > 0) else np.nan
    upside = ((fv_per_share / price) - 1)*100 if (np.isfinite(fv_per_share) and np.isfinite(price) and price > 0) else np.nan

    metrics = {
        "ROIC %": roic,
        "WACC %": wacc*100 if np.isfinite(wacc) else np.nan,
        "ROIC - WACC (pp)": (roic - (wacc*100)) if (np.isfinite(roic) and np.isfinite(wacc)) else np.nan,
        "Price": price,
        "Shares": shares,
        "Fair Value / sh (DCF)": fv_per_share,
        "Upside % vs Price": upside,
        "Assumed g (5y)": g*100,
        "Terminal g": term_growth*100,
        "Ke % (CAPM)": ke*100,
        "Kd % after tax": kd_after_tax*100,
        "FCF last": fcf0,
        "Net Debt": net_debt
    }
    inputs = {
        "FCF last": fcf0,
        "Debt": debt_value,
        "Cash": _safe(cash.get(col, np.nan)),
        "Equity book": _safe(_row(key_bal_a, "Total Equity").get(col, np.nan)),
        "Tax rate": (tax_rate*100) if np.isfinite(tax_rate) else np.nan
    }
    return metrics, inputs


# ---------- Sensibilidades DCF ----------
def _dcf_value_per_share(fcf0, g5, term_g, wacc, years, net_debt, shares):
    if not all(np.isfinite([fcf0, wacc, net_debt])) or not np.isfinite(shares) or shares <= 0 or wacc <= term_g or wacc <= 0:
        return np.nan
    disc = 1 + wacc
    pv = 0.0
    f = fcf0
    for t in range(1, years+1):
        f *= (1 + g5)
        pv += f / (disc**t)
    tv = f * (1 + term_g) / (wacc - term_g)
    pv += tv / (disc**years)
    equity = pv - net_debt
    return equity / shares if np.isfinite(equity) else np.nan


def make_dcf_panels(roic_val: Dict[str, float], price: float, years: int = 5):
    fcf0 = roic_val.get("FCF last")
    net_debt = roic_val.get("Net Debt")
    shares = roic_val.get("Shares")
    base_wacc = roic_val.get("WACC %", np.nan) / 100.0
    base_g5 = roic_val.get("Assumed g (5y)", np.nan) / 100.0
    base_term = roic_val.get("Terminal g", np.nan) / 100.0

    waccs = [max(0.02, base_wacc + d) for d in [-0.02, -0.01, 0.0, 0.01, 0.02]]
    terms = [max(0.00, base_term + d) for d in [-0.01, -0.005, 0.0, 0.005, 0.01]]
    g5s = [max(-0.02, base_g5 + d) for d in [-0.02, -0.01, 0.0, 0.01, 0.02]]

    grid1 = pd.DataFrame(index=[f"{w*100:.1f}%" for w in waccs], columns=[f"{t*100:.1f}%" for t in terms])
    for w in waccs:
        for t in terms:
            grid1.loc[f"{w*100:.1f}%", f"{t*100:.1f}%"] = _dcf_value_per_share(fcf0, base_g5, t, w, years, net_debt, shares)

    grid2 = pd.DataFrame(index=[f"{w*100:.1f}%" for w in waccs], columns=[f"{g*100:.1f}%" for g in g5s])
    for w in waccs:
        for g in g5s:
            grid2.loc[f"{w*100:.1f}%", f"{g*100:.1f}%"] = _dcf_value_per_share(fcf0, g, base_term, w, years, net_debt, shares)

    if np.isfinite(price) and price > 0:
        grid1["Upside vs Price"] = (grid1.iloc[:, -1] / price - 1) * 100
        grid2["Upside vs Price"] = (grid2.iloc[:, -1] / price - 1) * 100

    return grid1, grid2


# ------------------ Timeline por año ------------------
def build_timeline(key_path: Path, full_path: Optional[Path], ticker: Optional[str],
                   risk_free: float = 0.045, mrp: float = 0.05) -> Dict[str, pd.DataFrame]:

    inc = _read_sheet(key_path, "Income (Annual)")
    bal = _read_sheet(key_path, "Balance (Annual)")
    cf  = _read_sheet(key_path, "Cash Flow (Annual)")
    if inc.empty or bal.empty or cf.empty:
        raise RuntimeError("Faltan hojas Annual en el KEY (Income/Balance/Cash Flow).")

    cols = _to_sorted_cols(inc)

    revenue = _row(inc, "Revenue")
    ebit    = _row(inc, "Operating Income (EBIT)")
    net_inc = _row(inc, "Net Income")
    assets  = _row(bal, "Total Assets")
    equity  = _row(bal, "Total Equity")
    cash    = _row(bal, "Cash & Equivalents")
    debt    = _row(bal, "Total Debt")
    gross   = _row(inc, "Gross Profit")
    pretax  = _row(inc, "Pretax Income")
    taxexp  = _row(inc, "Tax Expense")
    ocf     = _row(cf,  "Operating Cash Flow")
    fcf     = _row(cf,  "Free Cash Flow (OCF - CapEx)")

    cur_assets = _pick_from_full(full_path, "Balance Sheet (Annual)", ["Total Current Assets", "Current Assets"])
    cur_liab   = _pick_from_full(full_path, "Balance Sheet (Annual)", ["Total Current Liabilities", "Current Liabilities"])
    retained   = _pick_from_full(full_path, "Balance Sheet (Annual)", ["Retained Earnings", "Accumulated Retained Earnings Deficit", "Retained Earnings Accumulated Deficit"])
    interest   = _pick_from_full(full_path, "Income Statement (Annual)", ["Interest Expense", "Interest Expense Non Operating", "Non Operating Interest Expense"]).abs()

    beta = np.nan
    if ticker and yf is not None:
        try:
            tk = yf.Ticker(ticker)
            beta = float(tk.info.get("beta", np.nan))
        except Exception:
            pass
    ke = risk_free + (beta if np.isfinite(beta) else 1.0) * mrp

    ratios_rows = {"ROIC %": [], "WACC %": [], "ROIC - WACC (pp)": [], "Gross Margin %": [], "Net Margin %": []}
    scores_rows = {"F-Score": [], "Altman Z / Z'": []}
    fcf_rows    = {"FCF": []}
    base_rows   = {"Revenue": [], "EBIT": [], "Net Income": [], "Total Assets": [], "Total Equity": [],
                   "Cash & Eq": [], "Total Debt": [], "Current Assets": [], "Current Liabilities": [],
                   "Retained Earnings": [], "Interest Expense (abs)": []}

    gm = _safe_div(gross, revenue) * 100.0
    nm = _safe_div(net_inc, revenue) * 100.0

    for i, c in enumerate(cols):
        rev_t = _safe(revenue.get(c, np.nan))
        ebit_t = _safe(ebit.get(c, np.nan))
        net_t = _safe(net_inc.get(c, np.nan))
        assets_t = _safe(assets.get(c, np.nan))
        equity_t = _safe(equity.get(c, np.nan))
        cash_t = _safe(cash.get(c, np.nan))
        debt_t = _safe(debt.get(c, np.nan))
        ca_t = _safe(cur_assets.get(c, np.nan))
        cl_t = _safe(cur_liab.get(c, np.nan))
        reta_t = _safe(retained.get(c, np.nan))
        int_t = _safe(interest.get(c, np.nan))

        try:
            tx = max(0.0, min(0.40, _safe(taxexp.get(c, np.nan)) / _safe(pretax.get(c, np.nan))))
        except Exception:
            tx = 0.25

        ic_t = assets_t - cash_t
        roic_t = (ebit_t * (1 - (tx if np.isfinite(tx) else 0.25)) / ic_t)*100 if (np.isfinite(ebit_t) and np.isfinite(ic_t) and ic_t not in [0, np.nan]) else np.nan

        kd_t = (abs(int_t) / max(debt_t, 1e-9)) if (np.isfinite(int_t) and np.isfinite(debt_t) and debt_t != 0) else np.nan
        kd_after_tax_t = kd_t * (1 - (tx if np.isfinite(tx) else 0.25)) if np.isfinite(kd_t) else np.nan

        ev_t = (equity_t if np.isfinite(equity_t) else 0.0) + (debt_t if np.isfinite(debt_t) else 0.0)
        we_t = (equity_t / ev_t) if (np.isfinite(ev_t) and ev_t != 0) else 0.6
        wd_t = (debt_t / ev_t) if (np.isfinite(ev_t) and ev_t != 0) else 0.4
        wacc_t = we_t * ke + wd_t * (kd_after_tax_t if np.isfinite(kd_after_tax_t) else 0.06*(1-0.25))
        wacc_t_pct = wacc_t * 100 if np.isfinite(wacc_t) else np.nan

        ratios_rows["ROIC %"].append(roic_t)
        ratios_rows["WACC %"].append(wacc_t_pct)
        ratios_rows["ROIC - WACC (pp)"].append(roic_t - wacc_t_pct if (np.isfinite(roic_t) and np.isfinite(wacc_t_pct)) else np.nan)
        ratios_rows["Gross Margin %"].append(_safe(gm.get(c, np.nan)))
        ratios_rows["Net Margin %"].append(_safe(nm.get(c, np.nan)))

        fcf_rows["FCF"].append(_safe(fcf.get(c, np.nan)))

        base_rows["Revenue"].append(rev_t)
        base_rows["EBIT"].append(ebit_t)
        base_rows["Net Income"].append(net_t)
        base_rows["Total Assets"].append(assets_t)
        base_rows["Total Equity"].append(equity_t)
        base_rows["Cash & Eq"].append(cash_t)
        base_rows["Total Debt"].append(debt_t)
        base_rows["Current Assets"].append(ca_t)
        base_rows["Current Liabilities"].append(cl_t)
        base_rows["Retained Earnings"].append(reta_t)
        base_rows["Interest Expense (abs)"].append(int_t)

        # F-Score_t (requiere t-1)
        if i == 0:
            scores_rows["F-Score"].append(np.nan)
        else:
            c_prev = cols[i-1]
            s = 0
            roa_t = net_t / assets_t if (np.isfinite(net_t) and np.isfinite(assets_t) and assets_t != 0) else np.nan
            net_p = _safe(net_inc.get(c_prev, np.nan))
            assets_p = _safe(assets.get(c_prev, np.nan))
            roa_p = net_p / assets_p if (np.isfinite(net_p) and np.isfinite(assets_p) and assets_p != 0) else np.nan
            s += int(np.isfinite(roa_t) and roa_t > 0)
            ocf_t = _safe(ocf.get(c, np.nan))
            s += int(np.isfinite(ocf_t) and ocf_t > 0)
            s += int(np.isfinite(roa_t) and np.isfinite(roa_p) and roa_t > roa_p)
            s += int(np.isfinite(ocf_t) and np.isfinite(net_t) and ocf_t > net_t)
            debt_p = _safe(debt.get(c_prev, np.nan))
            s += int(np.isfinite(debt_t) and np.isfinite(debt_p) and (debt_t - debt_p) < 0)
            ca_p = _safe(cur_assets.get(c_prev, np.nan))
            cl_p = _safe(cur_liab.get(c_prev, np.nan))
            cr_t = (ca_t / cl_t) if (np.isfinite(ca_t) and np.isfinite(cl_t) and cl_t != 0) else np.nan
            cr_p = (ca_p / cl_p) if (np.isfinite(ca_p) and np.isfinite(cl_p) and cl_p != 0) else np.nan
            s += int(np.isfinite(cr_t) and np.isfinite(cr_p) and cr_t > cr_p)
            issued = _pick_from_full(full_path, "Cash Flow (Annual)", [
                "Issuance Of Stock", "Common Stock Issued",
                "Sale Purchase Of Stock", "Sale Purchase Of Common And Preferred Stock"
            ])
            iss_t = _safe(issued.get(c, np.nan)) if not issued.empty else np.nan
            s += int(np.isfinite(iss_t) and iss_t <= 0)
            gm_t = _safe(gm.get(c, np.nan))
            gm_p = _safe(gm.get(c_prev, np.nan))
            s += int(np.isfinite(gm_t) and np.isfinite(gm_p) and gm_t > gm_p)
            at_t = _safe(rev_t / assets_t) if (np.isfinite(rev_t) and np.isfinite(assets_t) and assets_t != 0) else np.nan
            at_p = _safe(_safe(revenue.get(c_prev, np.nan)) / _safe(assets.get(c_prev, np.nan))) if (np.isfinite(_safe(revenue.get(c_prev, np.nan))) and np.isfinite(_safe(assets.get(c_prev, np.nan))) and _safe(assets.get(c_prev, np.nan)) != 0) else np.nan
            s += int(np.isfinite(at_t) and np.isfinite(at_p) and at_t > at_p)
            scores_rows["F-Score"].append(s)

        # Altman por año: Z clásico si se puede, si no Z'
        A = _safe(ca_t - cl_t) / assets_t if (np.isfinite(ca_t) and np.isfinite(cl_t) and np.isfinite(assets_t) and assets_t != 0) else np.nan
        B = _safe(reta_t) / assets_t if (np.isfinite(reta_t) and np.isfinite(assets_t) and assets_t != 0) else np.nan
        C = _safe(ebit_t) / assets_t if (np.isfinite(ebit_t) and np.isfinite(assets_t) and assets_t != 0) else np.nan
        E = _safe(rev_t) / assets_t if (np.isfinite(rev_t) and np.isfinite(assets_t) and assets_t != 0) else np.nan
        # Usamos Z' (privadas) por default — el clásico requiere market cap histórico
        Dp = equity_t / debt_t if (np.isfinite(equity_t) and np.isfinite(debt_t) and debt_t != 0) else np.nan
        z_val = 0.717*A + 0.847*B + 3.107*C + 0.420*Dp + 0.998*E
        scores_rows["Altman Z / Z'"].append(z_val)

    ratios_df = pd.DataFrame(ratios_rows, index=cols).T
    scores_df = pd.DataFrame(scores_rows, index=cols).T
    fcf_df    = pd.DataFrame(fcf_rows, index=cols).T
    base_df   = pd.DataFrame(base_rows, index=cols).T
    return {"Timeline_Ratios": ratios_df, "Timeline_Scores": scores_df, "Timeline_FCF": fcf_df, "Timeline_Base": base_df}


# ------------------ Pipeline ------------------
def build_scorecard_with_comparison(key_path: Path, full_path: Optional[Path], ticker: Optional[str]) -> Dict[str, pd.DataFrame]:
    inc_a = _read_sheet(key_path, "Income (Annual)")
    bal_a = _read_sheet(key_path, "Balance (Annual)")
    cf_a  = _read_sheet(key_path, "Cash Flow (Annual)")

    # métricas último año
    fscore_last, fdetail = piotroski_f_score_last(inc_a, bal_a, cf_a, full_path, full_path)
    z_last, zdetail = altman_z_last(full_path, inc_a, ticker) if full_path else (np.nan, {})
    roic_val, inputs = roic_wacc_and_dcf_last(inc_a, bal_a, cf_a, ticker, full_path)

    spread = roic_val.get("ROIC - WACC (pp)", np.nan)
    upside = roic_val.get("Upside % vs Price", np.nan)

    decision = "HOLD"
    reason = []
    if np.isfinite(spread) and spread > 2:
        reason.append("ROIC > WACC +2pp")
    if np.isfinite(fscore_last) and fscore_last >= 7:
        reason.append("F-Score ≥ 7")
    if np.isfinite(z_last) and z_last >= 3:
        reason.append("Z-Score seguro (≥3)")
    if np.isfinite(upside) and upside > 20:
        reason.append("Upside DCF > 20%")

    positives = sum([
        int(np.isfinite(spread) and spread > 2),
        int(np.isfinite(fscore_last) and fscore_last >= 7),
        int(np.isfinite(z_last) and z_last >= 3),
        int(np.isfinite(upside) and upside > 20),
    ])
    if positives >= 3:
        decision = "BUY"
    elif positives == 0:
        decision = "AVOID"

    scorecard = pd.DataFrame({
        "Metric": [
            "Piotroski F-Score (0-9)",
            "Altman Z-Score",
            "ROIC %",
            "WACC %",
            "ROIC - WACC (pp)",
            "Fair Value / sh (DCF)",
            "Upside % vs Price",
            "Decision",
            "Reasons"
        ],
        "Value": [
            fscore_last, z_last,
            roic_val.get("ROIC %"),
            roic_val.get("WACC %"),
            spread,
            roic_val.get("Fair Value / sh (DCF)"),
            upside,
            decision,
            "; ".join(reason) if reason else "—"
        ]
    })

    # timeline completo
    tl = build_timeline(key_path, full_path, ticker)

    # armo mini comparación de últimos 5 períodos para pegar a Scorecard
    def last5(df: pd.DataFrame, row_name: str) -> pd.Series:
        if df.empty or row_name not in df.index:
            return pd.Series(dtype="float64")
        cols = list(df.columns)
        cols = cols[-5:] if len(cols) > 5 else cols
        s = df.loc[row_name, cols]
        return pd.to_numeric(s, errors="coerce")

    compare = pd.DataFrame({
        "ROIC %": last5(tl["Timeline_Ratios"], "ROIC %"),
        "WACC %": last5(tl["Timeline_Ratios"], "WACC %"),
        "Spread (ROIC-WACC)": last5(tl["Timeline_Ratios"], "ROIC - WACC (pp)"),
        "F-Score": last5(tl["Timeline_Scores"], "F-Score"),
        "Altman Z/Z'": last5(tl["Timeline_Scores"], "Altman Z / Z'"),
        "FCF": last5(tl["Timeline_FCF"], "FCF"),
        "Gross Margin %": last5(tl["Timeline_Ratios"], "Gross Margin %"),
        "Net Margin %": last5(tl["Timeline_Ratios"], "Net Margin %"),
    }).T  # filas= métricas, columnas = años

    # paneles DCF
    grid1, grid2 = make_dcf_panels(roic_val, roic_val.get("Price", np.nan))

    # detalles/inputs
    f_det = pd.DataFrame([fdetail]).T.reset_index()
    f_det.columns = ["Piotroski Signal", "Flag (1/0)"]
    z_det = pd.DataFrame([zdetail]).T.reset_index()
    z_det.columns = ["Altman Component", "Value"]
    roic_det = pd.DataFrame([roic_val]).T.reset_index()
    roic_det.columns = ["Metric", "Value"]
    inputs_df = pd.DataFrame([inputs]).T.reset_index()
    inputs_df.columns = ["Input", "Value"]

    return {
        "Scorecard": scorecard,
        "Compare_Periods": compare,
        "Piotroski_Detail": f_det,
        "Altman_Detail": z_det,
        "Valuation_Detail": roic_det,
        "Inputs": inputs_df,
        "Timeline_Ratios": tl["Timeline_Ratios"],
        "Timeline_Scores": tl["Timeline_Scores"],
        "Timeline_FCF": tl["Timeline_FCF"],
        "Timeline_Base": tl["Timeline_Base"],
        "DCF_WACC_vs_TermG": grid1,
        "DCF_WACC_vs_G5y": grid2,
    }


# ------------------ CLI ------------------
def main():
    import argparse
    from pathlib import Path
    import pandas as pd

    parser = argparse.ArgumentParser(description="Scorecard & Valuation con comparación multi-período.")
    parser.add_argument("--key", required=True, help="Excel KEY (Paso 2).")
    parser.add_argument("--full", required=True, help="Excel FULL (Paso 1).")
    parser.add_argument("--ticker", required=False, help="Ticker para precio/beta/shares (yfinance).")
    parser.add_argument("--out", "-o", default=None, help="Excel de salida (default: <key>_score.xlsx)")
    args = parser.parse_args()

    key_path = Path(args.key)
    full_path = Path(args.full)
    out_path = Path(args.out) if args.out else key_path.with_name(key_path.stem.replace("_key", "") + "_score.xlsx")

    # Construye todas las hojas (debe existir build_scorecard_with_comparison)
    sheets = build_scorecard_with_comparison(key_path, full_path, args.ticker)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # 1) Scorecard (bloque principal de decisión del último año)
        sheets["Scorecard"].to_excel(writer, sheet_name="Scorecard", index=False)

        # 2) Título + tabla de comparación dentro de la MISMA hoja "Scorecard"
        startrow = sheets["Scorecard"].shape[0] + 2  # deja una línea en blanco
        title_df = pd.DataFrame([["Compare_Periods (últimos 5)"]])
        title_df.to_excel(writer, sheet_name="Scorecard", index=False, header=False, startrow=startrow)

        startrow += 2
        # La comparación tiene métricas en filas y años en columnas → mantener index=True
        sheets["Compare_Periods"].to_excel(writer, sheet_name="Scorecard", startrow=startrow)

        # 3) Resto de hojas auxiliares
        for name in [
            "Piotroski_Detail", "Altman_Detail", "Valuation_Detail", "Inputs",
            "Timeline_Ratios", "Timeline_Scores", "Timeline_FCF", "Timeline_Base",
            "DCF_WACC_vs_TermG", "DCF_WACC_vs_G5y"
        ]:
            if name in sheets:
                sheets[name].to_excel(writer, sheet_name=name[:31], index=True)

    print(f"✔ Scorecard creado con comparación en la misma hoja: {out_path}")

if __name__ == "__main__":
    main()
