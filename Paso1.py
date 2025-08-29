#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
01_fetch_full_financials.py
---------------------------------
Descarga desde Yahoo Finance (vía yfinance) los ESTADOS COMPLETOS para un ticker:
- Income Statement (anual y trimestral)
- Balance Sheet (anual y trimestral)
- Cash Flow (anual y trimestral)

Uso:
    pip install yfinance pandas openpyxl

    # ejemplo 1: pasar el ticker directo
    python 01_fetch_full_financials.py --ticker GLOB --out glob_full.xlsx

    # ejemplo 2: sin pasar ticker, lo pedirá por consola
    python 01_fetch_full_financials.py
"""

import argparse
from pathlib import Path
import sys
import pandas as pd

try:
    import yfinance as yf
except ImportError:
    print("Instala dependencias: pip install yfinance pandas openpyxl", file=sys.stderr)
    raise


def _normalize(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    out.columns = [c.strftime("%Y-%m-%d") if hasattr(c, "strftime") else str(c) for c in out.columns]
    out.index = [str(i).strip() for i in out.index]
    return out


def fetch_all(ticker: str) -> dict:
    tk = yf.Ticker(ticker)

    # cashflow puede venir como cashflow/cash_flow
    cf_a = getattr(tk, "cashflow", None)
    if (cf_a is None) or (hasattr(cf_a, "empty") and cf_a.empty):
        cf_a = getattr(tk, "cash_flow", None)

    cf_q = getattr(tk, "quarterly_cashflow", None)
    if (cf_q is None) or (hasattr(cf_q, "empty") and cf_q.empty):
        cf_q = getattr(tk, "quarterly_cash_flow", None)

    data = {
        "Income Statement (Annual)": _normalize(tk.financials),
        "Income Statement (Quarterly)": _normalize(tk.quarterly_financials),
        "Balance Sheet (Annual)": _normalize(tk.balance_sheet),
        "Balance Sheet (Quarterly)": _normalize(tk.quarterly_balance_sheet),
        "Cash Flow (Annual)": _normalize(cf_a),
        "Cash Flow (Quarterly)": _normalize(cf_q),
    }
    return data


def write_excel(dfs: dict, out_path: Path) -> None:
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name, df in dfs.items():
            sheet = name[:31]
            if df is None or df.empty:
                pd.DataFrame({"Mensaje": [f"Sin datos para {name}"]}).to_excel(writer, index=False, sheet_name=sheet)
            else:
                df.to_excel(writer, sheet_name=sheet)
    print(f"✔ Archivo Excel creado: {out_path}")


def main():
    ap = argparse.ArgumentParser(description="Descarga estados financieros COMPLETOS desde Yahoo Finance para un ticker.")
    ap.add_argument("--ticker", "-t", required=False, help="Ticker (ej: AAPL, NVDA, GLOB, YPF)")
    ap.add_argument("--out", "-o", default=None, help="Excel de salida. Default: <ticker>_full.xlsx")
    args = ap.parse_args()

    ticker = (args.ticker or "").strip().upper()
    if not ticker:
        try:
            ticker = input("👉 Ingresá el ticker de la empresa (ej: AAPL, GLOB, NVDA, YPF): ").strip().upper()
        except EOFError:
            ticker = ""

    if not ticker:
        print("⚠ No se especificó ticker. Abortando.")
        return

    out_path = Path(args.out) if args.out else Path(f"{ticker}_full.xlsx")

    dfs = fetch_all(ticker)
    if not any([not (v is None or v.empty) for v in dfs.values()]):
        print("⚠ No se encontraron datos. Verifica el ticker o intenta más tarde.", file=sys.stderr)

    write_excel(dfs, out_path)


if __name__ == "__main__":
    main()
