#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Paso2 - Curado de métricas clave (Silver)
-----------------------------------------
Lee el Excel "full" del Paso 1 y crea un Excel con lo esencial:

- Income: Revenue, Cost of Revenue, Gross Profit, Operating Income (EBIT),
          Net Income, Gross/Operating/Net Margin %
- Balance: Cash & Equivalents, Total Debt, Total Assets, Total Liabilities, Total Equity
- Cash Flow: Operating Cash Flow, Capital Expenditures (CapEx),
             Free Cash Flow (OCF - CapEx), Financing Cash Flow

Uso (ejemplos):
    pip install pandas openpyxl
    python Paso2.py --in GLOB_full.xlsx --out GLOB_key.xlsx
    # si no pasás --out, genera <input>_key.xlsx
"""

import argparse
from pathlib import Path
from typing import Dict, List, Optional
import pandas as pd


# -------------------- Helpers --------------------
def _read_sheet(path: Path, name: str) -> pd.DataFrame:
    """Lee una hoja por nombre (o truncada a 31 chars) y normaliza columnas/índice."""
    try:
        df = pd.read_excel(path, sheet_name=name, index_col=0)
    except ValueError:
        df = pd.read_excel(path, sheet_name=name[:31], index_col=0)
    except Exception:
        return pd.DataFrame()

    if df is None or df.empty:
        return pd.DataFrame()

    # Si era una hoja "vacía" con el mensaje
    if "Mensaje" in df.columns and len(df.columns) == 1:
        return pd.DataFrame()

    # Columnas como fechas AAAA-MM si corresponde
    cols = []
    for c in df.columns:
        try:
            cols.append(pd.to_datetime(c).strftime("%Y-%m"))
        except Exception:
            cols.append(str(c))
    df.columns = cols

    # Índice como string limpio
    df.index = [str(i).strip() for i in df.index]
    return df


def _find_first(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """Devuelve el label del índice que mejor matchea algún candidato (case-insensitive)."""
    if df is None or df.empty:
        return None
    idx_lower = {str(ix).lower(): ix for ix in df.index}

    # exact match
    for c in candidates:
        c_low = c.lower()
        if c_low in idx_lower:
            return idx_lower[c_low]

    # contains fallback
    for c in candidates:
        c_low = c.lower()
        for low, real in idx_lower.items():
            if c_low in low:
                return real
    return None


def _pick_rows(df: pd.DataFrame, spec: Dict[str, List[str]]) -> pd.DataFrame:
    """Construye un DF con solo las filas pedidas (o NaN si no existen)."""
    if df is None or df.empty:
        return pd.DataFrame()
    rows = {}
    for nice, candidates in spec.items():
        ix = _find_first(df, candidates)
        if ix is not None and ix in df.index:
            rows[nice] = df.loc[ix]
        else:
            rows[nice] = pd.Series([pd.NA] * len(df.columns), index=df.columns)
    out = pd.DataFrame(rows).T
    out.index.name = "Metric"
    return out


# -------------------- Core --------------------
def curate(input_excel: Path) -> Dict[str, pd.DataFrame]:
    # Leer hojas del "full"
    fin_a = _read_sheet(input_excel, "Income Statement (Annual)")
    fin_q = _read_sheet(input_excel, "Income Statement (Quarterly)")
    bal_a = _read_sheet(input_excel, "Balance Sheet (Annual)")
    bal_q = _read_sheet(input_excel, "Balance Sheet (Quarterly)")
    cf_a = _read_sheet(input_excel, "Cash Flow (Annual)")
    cf_q = _read_sheet(input_excel, "Cash Flow (Quarterly)")

    # Income & Balance: alias robustos
    INCOME_ITEMS = {
        "Revenue": ["Total Revenue", "Revenue"],
        "Cost of Revenue": ["Cost Of Revenue", "Cost Of Goods Sold", "Cost of Revenue"],
        "Gross Profit": ["Gross Profit"],
        "Operating Income (EBIT)": ["Operating Income", "Ebit", "Ebit (Operating Income)"],
        "Net Income": ["Net Income", "Net Income Applicable To Common Shares", "Net Income Common Stockholders"],
    }
    BALANCE_ITEMS = {
        "Cash & Equivalents": [
            "Cash And Cash Equivalents",
            "Cash And Cash Equivalents Including Restricted Cash",
            "Cash",
        ],
        "Total Debt": [
            "Total Debt",
            "Short Long Term Debt Total",
            "Long Term Debt",
            "Long Term Debt And Capital Lease Obligation",
        ],
        "Total Assets": ["Total Assets"],
        "Total Liabilities": ["Total Liab", "Total Liabilities"],
        "Total Equity": ["Total Stockholder Equity", "Total Shareholder Equity", "Stockholders Equity"],
    }

    # Cash Flow: incluye las variantes que viste (singular/plural/reportado)
    CASHFLOW_ITEMS = {
        "Operating Cash Flow": [
            "Operating Cash Flow",
            "Total Cash From Operating Activities",
            "Net Cash Provided By Operating Activities",
            "Net Cash Provided By (Used In) Operating Activities",
            "Net Cash Provided By Used In Operating Activities",
            "Cash Provided By Operating Activities",
            "Net Cash From Operating Activities",
            "Cash Flow From Continuing Operating Activities",
        ],
        "Capital Expenditures (CapEx)": [
            "Capital Expenditure",                # singular
            "Capital Expenditure Reported",       # variante reportada
            "Capital Expenditures",               # plural
            "Investments In Property Plant And Equipment",
            "Purchase Of Property Plant And Equipment",
            "Additions To Property Plant And Equipment",
        ],
        "Financing Cash Flow": [
            "Financing Cash Flow",
            "Total Cash From Financing Activities",
            "Net Cash Provided By (Used In) Financing Activities",
            "Net Cash Provided By Used In Financing Activities",
            "Net Cash From Financing Activities",
            "Cash Provided By Financing Activities",
            "Cash Flow From Continuing Financing Activities",
        ],
    }

    # Selección de filas clave
    ia = _pick_rows(fin_a, INCOME_ITEMS)
    iq = _pick_rows(fin_q, INCOME_ITEMS)
    ba = _pick_rows(bal_a, BALANCE_ITEMS)
    bq = _pick_rows(bal_q, BALANCE_ITEMS)
    ca = _pick_rows(cf_a, CASHFLOW_ITEMS)
    cq = _pick_rows(cf_q, CASHFLOW_ITEMS)

    # Derivadas
    def add_income_margins(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df
        out = df.copy()
        if "Revenue" in out.index:
            rev = out.loc["Revenue"]
            gross = out.loc["Gross Profit"] if "Gross Profit" in out.index else pd.Series([pd.NA] * len(out.columns), index=out.columns)
            op = out.loc["Operating Income (EBIT)"] if "Operating Income (EBIT)" in out.index else pd.Series([pd.NA] * len(out.columns), index=out.columns)
            net = out.loc["Net Income"] if "Net Income" in out.index else pd.Series([pd.NA] * len(out.columns), index=out.columns)
            with pd.option_context("mode.use_inf_as_na", True):
                out.loc["Gross Margin %"] = (gross / rev) * 100
                out.loc["Operating Margin %"] = (op / rev) * 100
                out.loc["Net Margin %"] = (net / rev) * 100
        return out

    def add_fcf(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df
        out = df.copy()
        if "Operating Cash Flow" in out.index and "Capital Expenditures (CapEx)" in out.index:
            out.loc["Free Cash Flow (OCF - CapEx)"] = out.loc["Operating Cash Flow"] - out.loc["Capital Expenditures (CapEx)"]
        else:
            out.loc["Free Cash Flow (OCF - CapEx)"] = pd.Series([pd.NA] * len(out.columns), index=out.columns)
        return out

    ia = add_income_margins(ia)
    iq = add_income_margins(iq)
    ca = add_fcf(ca)
    cq = add_fcf(cq)

    # Orden amigable
    def reorder(df: pd.DataFrame, order: List[str]) -> pd.DataFrame:
        if df.empty:
            return df
        present = [r for r in order if r in df.index]
        rest = [r for r in df.index if r not in present]
        return df.loc[present + rest]

    INCOME_ORDER = [
        "Revenue",
        "Cost of Revenue",
        "Gross Profit",
        "Operating Income (EBIT)",
        "Net Income",
        "Gross Margin %",
        "Operating Margin %",
        "Net Margin %",
    ]
    BALANCE_ORDER = [
        "Cash & Equivalents",
        "Total Debt",
        "Total Assets",
        "Total Liabilities",
        "Total Equity",
    ]
    CASH_ORDER = [
        "Operating Cash Flow",
        "Capital Expenditures (CapEx)",
        "Free Cash Flow (OCF - CapEx)",
        "Financing Cash Flow",
    ]

    ia = reorder(ia, INCOME_ORDER)
    iq = reorder(iq, INCOME_ORDER)
    ba = reorder(ba, BALANCE_ORDER)
    bq = reorder(bq, BALANCE_ORDER)
    ca = reorder(ca, CASH_ORDER)
    cq = reorder(cq, CASH_ORDER)

    return {
        "Income (Annual)": ia,
        "Income (Quarterly)": iq,
        "Balance (Annual)": ba,
        "Balance (Quarterly)": bq,
        "Cash Flow (Annual)": ca,
        "Cash Flow (Quarterly)": cq,
    }


def main():
    ap = argparse.ArgumentParser(description="Curar métricas clave desde un Excel FULL (Paso 1).")
    ap.add_argument("--in", "-i", required=True, help="Ruta al Excel FULL (salida del Paso 1).")
    ap.add_argument("--out", "-o", default=None, help="Excel de salida (default: <input>_key.xlsx)")
    args = ap.parse_args()

    in_path = Path(args.__dict__["in"])
    out_path = Path(args.out) if args.out else in_path.with_name(in_path.stem + "_key.xlsx")

    sheets = curate(in_path)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            sheet = name[:31]
            if df.empty:
                pd.DataFrame({"Mensaje": [f"Sin datos para {name}"]}).to_excel(writer, index=False, sheet_name=sheet)
            else:
                fmt = df.copy()
                for r in ["Gross Margin %", "Operating Margin %", "Net Margin %"]:
                    if r in fmt.index:
                        fmt.loc[r] = fmt.loc[r].astype(float)
                fmt.to_excel(writer, sheet_name=sheet)

    print(f"✔ Archivo Excel curado creado: {out_path}")


if __name__ == "__main__":
    main()
