#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
compare_scorecards.py
Lee <T>_scorecard.xlsx (hoja 'Scorecard') por cada ticker y arma un Excel
comparativo con las métricas clave.

Uso:
  python compare_scorecards.py --tickers KO,PEP,MNST
  python compare_scorecards.py --dir . --tickers GLOB,KO --out Scorecards_Compare.xlsx
  python compare_scorecards.py --tickers KO,PEP --no_rank
"""

import argparse
from pathlib import Path
import pandas as pd
import numpy as np
from datetime import datetime

# ---- helpers ----
def _safe_float(x):
    try:
        return float(x)
    except Exception:
        return np.nan

def _read_scorecard_book(base: Path, ticker: str) -> pd.DataFrame:
    path = base / f"{ticker}_scorecard.xlsx"
    if not path.exists():
        print(f"ADVERTENCIA: no se encontró {path.name}")
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, sheet_name="Scorecard")
        # algunos Paso4 escriben columnas Metric/Value
        if "Metric" in df.columns and "Value" in df.columns:
            df = df[["Metric", "Value"]]
        return df
    except Exception as e:
        print(f"ADVERTENCIA: no pude leer 'Scorecard' de {path.name}: {e}")
        return pd.DataFrame()

# mapeo flexible de nombres (por si cambió la etiqueta)
ALIASES = {
    "F-Score": [
        "Piotroski F-Score (0-9)", "Piotroski F-Score", "F-Score", "Piotroski"
    ],
    "Altman Z": [
        "Altman Z-Score", "Altman Z", "Z-Score"
    ],
    "ROIC %": [
        "ROIC %", "ROIC", "ROIC (%)"
    ],
    "WACC %": [
        "WACC %", "WACC", "WACC (%)"
    ],
    "Spread (pp)": [
        "ROIC - WACC (pp)", "ROIC minus WACC (pp)", "Spread (ROIC-WACC)"
    ],
    "FV/sh (DCF)": [
        "Fair Value / sh (DCF)", "Fair Value per Share (DCF)", "FV/sh"
    ],
    "Upside %": [
        "Upside % vs Price", "Upside (%)", "Upside"
    ],
    # opcionales si estuvieran
    "Price": ["Price", "Precio"],
    "Shares": ["Shares", "Acciones"],
}

def _pick_metric(row_map: dict, names: list):
    for n in names:
        if n in row_map and pd.notna(row_map[n]):
            return row_map[n]
    return np.nan

def extract_metrics(df: pd.DataFrame) -> dict:
    """
    df: hoja Scorecard con columnas ['Metric','Value'] o similar.
    devuelve dict { metric_name: value }
    """
    if df is None or df.empty:
        return {}
    if "Metric" in df.columns and "Value" in df.columns:
        m = dict(zip(df["Metric"].astype(str), df["Value"]))
    else:
        # fallback: primera col = metric, segunda = value
        m = dict(zip(df.iloc[:,0].astype(str), df.iloc[:,1]))
    out = {}
    for target, aliases in ALIASES.items():
        out[target] = _safe_float(_pick_metric(m, aliases))
    return out

def main():
    ap = argparse.ArgumentParser(description="Comparar scorecards de varios tickers.")
    ap.add_argument("--tickers", required=True, help="Lista separada por comas. Ej: KO,PEP,MNST")
    ap.add_argument("--dir", default=".", help="Carpeta donde están los <T>_scorecard.xlsx")
    ap.add_argument("--out", default="Scorecards_Compare.xlsx", help="Nombre del Excel de salida")
    ap.add_argument("--no_rank", action="store_true", help="No agregar columnas de ranking")
    args = ap.parse_args()

    base = Path(args.dir).resolve()
    tickers = [t.strip().upper() for t in args.tickers.split(",") if t.strip()]

    rows = []
    raw_tabs = {}  # opcional: para volcar la hoja Scorecard cruda por ticker

    for t in tickers:
        sc_df = _read_scorecard_book(base, t)
        raw_tabs[t] = sc_df if not sc_df.empty else pd.DataFrame()
        metrics = extract_metrics(sc_df) if not sc_df.empty else {}
        if not metrics:
            print(f"ADVERTENCIA: {t} sin métricas legibles en 'Scorecard'.")
        row = {"Ticker": t}
        row.update(metrics)
        rows.append(row)

    comp = pd.DataFrame(rows).set_index("Ticker")

    # Ranks (si no se desactivan)
    if not args.no_rank and not comp.empty:
        def rank_desc(col):
            if col in comp.columns:
                comp[f"Rank_{col}"] = comp[col].rank(ascending=False, method="min")
        def rank_asc(col):
            if col in comp.columns:
                comp[f"Rank_{col}"] = comp[col].rank(ascending=True, method="min")

        # donde "más alto = mejor"
        for c in ["Spread (pp)", "ROIC %", "F-Score", "Altman Z", "Upside %"]:
            rank_desc(c)
        # donde "más bajo = mejor" (ninguno directo del scorecard salvo que agregues Debt/Equity)
        # rank_asc("Debt/Equity")

        # ordenar por Spread y F-Score si existen
        sort_cols = [c for c in ["Spread (pp)", "F-Score", "Altman Z"] if c in comp.columns]
        if sort_cols:
            comp = comp.sort_values(by=sort_cols, ascending=False)

    # Escribir Excel
    out_path = base / args.out
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # hoja principal
        comp.to_excel(writer, sheet_name="Overview")

        # hoja notas
        notes = pd.DataFrame({
            "Field": ["Generated", "Directory", "Tickers"],
            "Value": [datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                      str(base),
                      ", ".join(tickers)]
        })
        notes.to_excel(writer, sheet_name="Notes", index=False)

        # hojas crudas por si querés ver el detalle
        for t, df in raw_tabs.items():
            if not df.empty:
                df.to_excel(writer, sheet_name=f"{t}_Scorecard"[:31], index=False)

    print(f"OK: comparativo creado -> {out_path}")

if __name__ == "__main__":
    main()
