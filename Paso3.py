#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Paso3 - Cálculo de Ratios (Gold) – versión robustecida con diagnóstico
----------------------------------------------------------------------
Lee el Excel FULL (Paso 1) y calcula ratios por período (annual o quarterly).
Además, escribe una hoja "Diagnostics" indicando qué filas se usaron o faltan.

Uso:
    pip install pandas openpyxl numpy
    python Paso3.py --in KO_full.xlsx --out KO_ratios.xlsx --freq annual
"""

import argparse
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import re
import numpy as np
import pandas as pd


# ------------------- Lectura y normalización -------------------
def _read_sheet(path: Path, name: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=name, index_col=0)
    except ValueError:
        df = pd.read_excel(path, sheet_name=name[:31], index_col=0)
    except Exception:
        return pd.DataFrame()
    if df is None or df.empty:
        return pd.DataFrame()
    if "Mensaje" in df.columns and len(df.columns) == 1:
        return pd.DataFrame()

    # columnas a fechas AAAA-MM cuando se pueda
    cols = []
    for c in df.columns:
        try:
            cols.append(pd.to_datetime(c).strftime("%Y-%m"))
        except Exception:
            cols.append(str(c))
    df.columns = cols
    # índice como str
    df.index = [str(i).strip() for i in df.index]
    return df


def _norm_text(s: str) -> str:
    """Normaliza texto para matching flexible (minúsculas, sin puntuación, sin múlt. espacios, &→and, sin stopwords menores)."""
    if s is None:
        return ""
    s = str(s).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    # unificar algunos sinónimos comunes
    syn = [
        ("shareholder", "stockholder"),
        ("liablities", "liabilities"),
        ("liability", "liabilities"),
        ("asset", "assets"),
        ("equivalent", "equivalents"),
        ("receivable", "receivables"),
        ("payable", "payables"),
        ("short term", "shortterm"),
        ("long term", "longterm"),
        ("carrying value", ""),  # ruido frecuente en Yahoo
        ("reported", ""),        # ruido
        ("continuing", ""),      # ruido en cash flows
    ]
    for a, b in syn:
        s = s.replace(a, b)
    return s


def _find_best_label(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    """Busca la mejor fila en df que matchee alguno de los alias (con normalización y fallback contains)."""
    if df is None or df.empty:
        return None

    # mapa normalizado → etiqueta real
    idx_map = {_norm_text(ix): ix for ix in df.index}
    idx_keys = list(idx_map.keys())

    # 1) exact por alias normalizado
    for a in aliases:
        na = _norm_text(a)
        if na in idx_map:
            return idx_map[na]

    # 2) contains (alias dentro de índice)
    for a in aliases:
        na = _norm_text(a)
        for k in idx_keys:
            if na and na in k:
                return idx_map[k]

    # 3) contains al revés (índice dentro del alias)
    for a in aliases:
        na = _norm_text(a)
        for k in idx_keys:
            if k and k in na:
                return idx_map[k]

    return None


def _pick_series(df: pd.DataFrame, aliases: List[str]) -> Tuple[pd.Series, Optional[str]]:
    """Devuelve (serie, etiqueta_usada). Serie vacía si no se encontró."""
    label = _find_best_label(df, aliases)
    if label is None:
        return pd.Series(dtype="float64"), None
    s = pd.to_numeric(df.loc[label], errors="coerce")
    return s, label


def _safe_div(a: pd.Series, b: pd.Series) -> pd.Series:
    return (a / b).replace([np.inf, -np.inf], np.nan)


def _avg_two_periods(series: pd.Series) -> pd.Series:
    """Promedio de dos períodos contiguos, ordenando por fecha real."""
    if series is None or series.empty:
        return pd.Series(dtype="float64")
    s = series.copy()
    try:
        s.index = pd.to_datetime(s.index, errors="coerce")
        s = s.sort_index()
        avg = (s + s.shift(1)) / 2
        avg.index = [dt.strftime("%Y-%m") if pd.notna(dt) else "" for dt in avg.index]
        return avg
    except Exception:
        return (s + s.shift(1)) / 2


# ------------------- Cálculo de ratios -------------------
def compute_ratios(full_excel: Path, freq: str = "annual") -> Tuple[Dict[str, pd.DataFrame], pd.DataFrame]:
    # Hojas a usar
    if freq.lower().startswith("q"):
        inc = _read_sheet(full_excel, "Income Statement (Quarterly)")
        bal = _read_sheet(full_excel, "Balance Sheet (Quarterly)")
    else:
        inc = _read_sheet(full_excel, "Income Statement (Annual)")
        bal = _read_sheet(full_excel, "Balance Sheet (Annual)")

    # Aliases (ampliados)
    REVENUE = ["Total Revenue", "Revenue"]
    COGS = ["Cost Of Revenue", "Cost Of Goods Sold", "Cost of Revenue"]
    GROSS_PROFIT = ["Gross Profit"]
    EBIT = ["Operating Income", "Ebit", "Ebit (Operating Income)"]
    NET_INCOME = ["Net Income", "Net Income Applicable To Common Shares", "Net Income Common Stockholders"]

    INT_EXP = ["Interest Expense", "Interest Expense Non Operating", "Non Operating Interest Expense"]
    TAX_EXP = ["Income Tax Expense", "Provision For Income Taxes", "Tax Provision"]
    PRETAX = ["Income Before Tax", "Earnings Before Tax", "Pretax Income"]

    TOTAL_ASSETS = ["Total Assets"]
    TOTAL_LIAB = ["Total Liab", "Total Liabilities"]
    TOTAL_EQUITY = ["Total Stockholder Equity", "Total Shareholder Equity", "Stockholders Equity"]

    CUR_ASSETS = ["Total Current Assets", "Current Assets", "Total current assets"]
    CUR_LIAB = ["Total Current Liabilities", "Current Liabilities", "Total current liabilities"]

    INVENTORY = ["Inventory", "Total Inventory"]
    AR = ["Net Receivables", "Accounts Receivable", "Trade And Other Receivables", "Receivables, net"]
    AP = ["Accounts Payable", "Trade And Other Payables", "Accounts payable, trade and other"]

    CASH_EQ = [
        "Cash And Cash Equivalents",
        "Cash And Cash Equivalents Including Restricted Cash",
        "Cash and cash equivalents",
        "Cash and cash equivalents, at carrying value",
        "Cash, cash equivalents and shortterm investments",
        "Cash, cash equivalents and short-term investments",
        "Cash",
    ]

    # Extraer series + etiqueta usada (para diagnóstico)
    used = []

    def get(df, aliases, name):
        s, lbl = _pick_series(df, aliases)
        used.append((name, lbl if lbl else "NOT FOUND"))
        return s

    sales = get(inc, REVENUE, "Revenue")
    cogs = get(inc, COGS, "COGS")
    gross = get(inc, GROSS_PROFIT, "Gross Profit")
    ebit = get(inc, EBIT, "EBIT")
    net = get(inc, NET_INCOME, "Net Income")
    int_exp = get(inc, INT_EXP, "Interest Expense").abs()
    tax_exp = get(inc, TAX_EXP, "Tax Expense")
    pretax = get(inc, PRETAX, "Pretax Income")

    assets = get(bal, TOTAL_ASSETS, "Total Assets")
    liab = get(bal, TOTAL_LIAB, "Total Liabilities")
    equity = get(bal, TOTAL_EQUITY, "Equity")
    cur_assets = get(bal, CUR_ASSETS, "Current Assets")
    cur_liab = get(bal, CUR_LIAB, "Current Liabilities")
    inventory = get(bal, INVENTORY, "Inventory")
    ar = get(bal, AR, "Accounts Receivable")
    ap = get(bal, AP, "Accounts Payable")
    cash = get(bal, CASH_EQ, "Cash & Equivalents")

    # ------------- Ratios -------------
    gross_margin = _safe_div(gross, sales) * 100
    op_margin = _safe_div(ebit, sales) * 100
    net_margin = _safe_div(net, sales) * 100

    inv_for_quick = inventory.fillna(0) if not inventory.empty else pd.Series(0.0, index=cur_assets.index)
    quick_ratio = _safe_div(cur_assets - inv_for_quick, cur_liab)
    current_ratio = _safe_div(cur_assets, cur_liab)
    defensive_interval = _safe_div(cash, cur_liab)

    inv_avg = _avg_two_periods(inventory) if not inventory.empty else pd.Series(dtype="float64")
    inv_turnover = _safe_div(cogs, inv_avg)
    days_inventory = _safe_div(365.0, inv_turnover)
    dso = _safe_div(ar, sales) * 365.0
    dpo = _safe_div(ap, cogs) * 365.0
    ccc = days_inventory + dso - dpo

    short_term_leverage = _safe_div(cur_liab, equity)
    long_term_leverage = _safe_div(liab - cur_liab, equity)
    debt_to_equity = _safe_div(liab, equity)
    financial_leverage = _safe_div(assets, equity)
    interest_coverage = _safe_div(ebit, int_exp.replace(0, np.nan))

    roa = _safe_div(net, assets) * 100
    roe = _safe_div(net, equity) * 100

    eff_tax_rate = _safe_div(tax_exp, pretax).clip(lower=0.0, upper=0.6)
    nopat = ebit * (1 - eff_tax_rate.fillna(0.25))
    invested_capital = (liab + equity) - cash
    roic = _safe_div(nopat, invested_capital) * 100

    dupont_margin = _safe_div(net, sales)
    dupont_turnover = _safe_div(sales, assets)
    dupont_leverage = _safe_div(assets, equity)
    dupont_roe = dupont_margin * dupont_turnover * dupont_leverage * 100

    def sort_cols(df: pd.DataFrame) -> pd.DataFrame:
        try:
            cols = sorted(df.columns, key=lambda x: pd.to_datetime(x, errors="coerce"))
            return df.loc[:, cols]
        except Exception:
            return df

    out = {
        "Liquidity": sort_cols(pd.DataFrame({
            "Current Ratio": current_ratio,
            "Quick Ratio": quick_ratio,
            "Defensive Interval (Cash/CL)": defensive_interval,
        }).T),
        "Activity": sort_cols(pd.DataFrame({
            "Asset Turnover": dupont_turnover,
            "Inventory Turnover": inv_turnover,
            "Days Inventory": days_inventory,
            "DSO (Days Sales Outstanding)": dso,
            "DPO (Days Payables Outstanding)": dpo,
            "Cash Conversion Cycle": ccc,
        }).T),
        "Leverage": sort_cols(pd.DataFrame({
            "Short-Term Leverage (CL/Equity)": short_term_leverage,
            "Long-Term Leverage (NCL/Equity)": long_term_leverage,
            "Debt-to-Equity (TL/Equity)": debt_to_equity,
            "Financial Leverage (Assets/Equity)": financial_leverage,
            "Interest Coverage (EBIT/Interest)": interest_coverage,
        }).T),
        "Profitability": sort_cols(pd.DataFrame({
            "Gross Margin %": gross_margin,
            "Operating Margin %": op_margin,
            "Net Margin %": net_margin,
            "ROA %": roa,
            "ROE %": roe,
            "ROIC %": roic,
            "DuPont ROE %": dupont_roe,
        }).T),
    }

    # ------------- Diagnostics sheet -------------
    diag_rows = []
    for name, lbl in used:
        diag_rows.append({"Metric base": name, "Row used": lbl})
    # también listamos todas las filas disponibles en esas hojas
    def list_index(df):
        return "; ".join(list(df.index)[:200]) if df is not None and not df.empty else "EMPTY"

    diag_rows.append({"Metric base": "-- Available (Income idx) --", "Row used": list_index(inc)})
    diag_rows.append({"Metric base": "-- Available (Balance idx) --", "Row used": list_index(bal)})
    diagnostics = pd.DataFrame(diag_rows)

    return out, diagnostics


# ------------------- CLI -------------------
def main():
    ap = argparse.ArgumentParser(description="Calcula ratios financieros desde un Excel FULL (Paso 1).")
    ap.add_argument("--in", "-i", required=True, help="Ruta al Excel COMPLETO (salida del Paso 1).")
    ap.add_argument("--out", "-o", default=None, help="Excel de salida (default: <input>_ratios.xlsx)")
    ap.add_argument("--freq", default="annual", choices=["annual", "quarterly"], help="Frecuencia de hojas a usar.")
    args = ap.parse_args()

    in_path = Path(args.__dict__["in"])
    out_path = Path(args.out) if args.out else in_path.with_name(in_path.stem + "_ratios.xlsx")

    sheets, diagnostics = compute_ratios(in_path, freq=args.freq)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            sheet = name[:31]
            df.to_excel(writer, sheet_name=sheet)
        diagnostics.to_excel(writer, sheet_name="Diagnostics", index=False)

    print(f"✔ Archivo Excel de ratios creado: {out_path}")
    print("   Revisa la pestaña 'Diagnostics' para ver qué filas se usaron/missing.")


if __name__ == "__main__":
    main()
