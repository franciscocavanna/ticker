#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Runthepipeline.py
Orquesta Paso1 → Paso2 → Paso3 → Paso4 para 1+ tickers y (opcional) ejecuta Paso5_v2
y crea comparativo de scorecards.

Uso:
  python Runthepipeline.py --tickers KO,PEP,MNST --use-v2 --v2-out Scorecards_v2.xlsx --v2-full-out Scorecards_v2_full.xlsx
"""

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path
from datetime import datetime

import pandas as pd
import numpy as np

# ================== util impresión/subprocesos ==================
def _safe_print(txt: str):
    try:
        print(txt)
    except Exception:
        try:
            print(txt.encode("utf-8", errors="backslashreplace").decode("utf-8", errors="ignore"))
        except Exception:
            print(str(txt).encode("ascii", errors="replace").decode())

def _run(cmd: list, workdir: Path, env_utf8: dict) -> int:
    _safe_print("→ " + " ".join(map(str, cmd)))
    res = subprocess.run(cmd, cwd=str(workdir), capture_output=True, text=True, env=env_utf8)
    if res.stdout: _safe_print(res.stdout.strip())
    if res.stderr: _safe_print(res.stderr.strip())
    return res.returncode

def _run_or_die(cmd: list, workdir: Path, env_utf8: dict):
    code = _run(cmd, workdir, env_utf8)
    if code != 0:
        raise SystemExit(code)

def _run_with_fallback(candidates: list, workdir: Path, env_utf8: dict):
    """Prueba comandos en orden hasta que uno devuelva 0."""
    last = None
    for cmd in candidates:
        code = _run(cmd, workdir, env_utf8)
        if code == 0:
            return
        last = code
    raise SystemExit(last or 1)

# ================== comparativo de scorecards (sector-aware) ==================
ALIASES = {
    "F-Score": ["Piotroski F-Score (0-9)", "Piotroski F-Score", "F-Score", "Piotroski"],
    "Altman Z": ["Altman Z-Score", "Altman Z", "Z-Score"],
    "ROIC %": ["ROIC %", "ROIC", "ROIC (%)"],
    "WACC %": ["WACC %", "WACC", "WACC (%)"],
    "Spread (pp)": ["ROIC - WACC (pp)", "ROIC minus WACC (pp)", "Spread (ROIC-WACC)"],
    "FV/sh (DCF)": ["Fair Value / sh (DCF)", "Fair Value per Share (DCF)", "FV/sh"],
    "Upside %": ["Upside % vs Price", "Upside (%)", "Upside"],
    "Price": ["Price", "Precio"],
    "Shares": ["Shares", "Acciones"],
}

def _safe_float(x):
    try:
        return float(x)
    except Exception:
        return np.nan

def _read_scorecard(path: Path) -> pd.DataFrame:
    if not path.exists():
        _safe_print(f"ADVERTENCIA: no se encontró {path.name}")
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, sheet_name="Scorecard")
        if "Metric" in df.columns and "Value" in df.columns:
            return df[["Metric", "Value"]]
        return df
    except Exception as e:
        _safe_print(f"ADVERTENCIA: no pude leer 'Scorecard' de {path.name}: {e}")
        return pd.DataFrame()

def _pick_metric(mapping: dict, names: list):
    for n in names:
        if n in mapping and pd.notna(mapping[n]):
            return mapping[n]
    return np.nan

def _extract_metrics(df: pd.DataFrame) -> dict:
    if df is None or df.empty:
        return {}
    if "Metric" in df.columns and "Value" in df.columns:
        m = dict(zip(df["Metric"].astype(str), df["Value"]))
    else:
        m = dict(zip(df.iloc[:,0].astype(str), df.iloc[:,1]))
    out = {}
    for target, aliases in ALIASES.items():
        out[target] = _safe_float(_pick_metric(m, aliases))
    return out

def build_scorecards_compare(base: Path, tickers: list, out_name: str, add_ranks: bool = True) -> Path:
    try:
        import yfinance as yf
    except Exception:
        yf = None

    rows, raw_tabs = [], {}
    for t in tickers:
        sc_path = base / f"{t}_scorecard.xlsx"
        df = _read_scorecard(sc_path)
        raw_tabs[t] = df if not df.empty else pd.DataFrame()
        metrics = _extract_metrics(df) if not df.empty else {}

        sector, industry = "Unknown", "Unknown"
        if yf is not None:
            try:
                info = yf.Ticker(t).info or {}
                sector = info.get("sector") or sector
                industry = info.get("industry") or industry
            except Exception:
                pass

        row = {"Ticker": t, "Sector": sector, "Industry": industry}
        row.update(metrics)
        rows.append(row)

    comp = pd.DataFrame(rows).set_index("Ticker")

    if add_ranks and not comp.empty:
        def rank_desc(col):
            if col in comp.columns:
                comp[f"Rank_{col}"] = comp[col].rank(ascending=False, method="min")
        for c in ["Spread (pp)", "ROIC %", "F-Score", "Altman Z", "Upside %"]:
            rank_desc(c)
        sort_cols = [c for c in ["Spread (pp)", "F-Score", "Altman Z"] if c in comp.columns]
        if sort_cols:
            comp = comp.sort_values(by=sort_cols, ascending=False)

    if "Sector" in comp.columns:
        bysec = comp.groupby("Sector")
        def add_vs_sector(col, higher_is_better=True):
            if col not in comp.columns: return
            med = bysec[col].transform("median")
            comp[f"{col} vs SecMed"] = comp[col] - med
            asc = not higher_is_better
            comp[f"SecRank_{col}"] = bysec[col].rank(ascending=asc, method="min")

        add_vs_sector("Spread (pp)", higher_is_better=True)
        add_vs_sector("ROIC %", higher_is_better=True)
        add_vs_sector("WACC %", higher_is_better=False)
        add_vs_sector("F-Score", higher_is_better=True)
        add_vs_sector("Altman Z", higher_is_better=True)
        add_vs_sector("Upside %", higher_is_better=True)

    out_path = base / out_name
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        comp.to_excel(writer, sheet_name="Overview")
        for t, df in raw_tabs.items():
            if not df.empty:
                df.to_excel(writer, sheet_name=f"{t}_Scorecard"[:31], index=False)
    return out_path

# ================== main pipeline ==================
def main():
    ap = argparse.ArgumentParser(description="Orquestador multi-ticker (Pasos 1→4) + comparativo + Paso5_v2 opcional.")
    ap.add_argument("--tickers", help="Lista separada por comas (ej: KO,PEP,MNST). Si se omite, se piden por consola.")
    ap.add_argument("--dir", default=".", help="Carpeta con Paso1..Paso4 y salidas Excel.")
    ap.add_argument("--sleep", type=float, default=2.0, help="Segundos de espera entre tickers (default 2).")
    ap.add_argument("--skip", default="", help="Pasos a saltar, ej: 2,4")
    ap.add_argument("--no-compare", dest="compare", action="store_false", help="No generar comparativo de scorecards.")
    ap.add_argument("--compare-out", default="Scorecards_Compare.xlsx", help="Nombre del comparativo.")
    ap.add_argument("--use-v2", action="store_true", help="Ejecuta Paso5_v2 al final")
    ap.add_argument("--v2-out", default="Scorecards_v2.xlsx", help="Salida Scorecards v2")
    ap.add_argument("--v2-full-out", default="Scorecards_v2_full.xlsx", help="Salida Scorecards v2 FULL")
    ap.set_defaults(compare=True)
    args = ap.parse_args()

    tickers_arg = args.tickers or input("Ingresá los tickers separados por coma (ej: KO, PEP, MNST): ").strip()
    if not tickers_arg:
        print("No se recibieron tickers. Saliendo.")
        return
    tickers = [t.strip().upper() for t in tickers_arg.split(",") if t.strip()]
    skip = {int(x) for x in args.skip.split(",") if x.strip().isdigit()}
    workdir = Path(args.dir).resolve()

    env_utf8 = os.environ.copy()
    env_utf8["PYTHONIOENCODING"] = "utf-8"
    env_utf8["PYTHONUTF8"] = "1"

    _safe_print(f"Working dir: {workdir}")

    succeeded = []
    for t in tickers:
        _safe_print(f"\n========== {t} ==========")
        full = workdir / f"{t}_full.xlsx"
        key  = workdir / f"{t}_key.xlsx"
        ratios = workdir / f"{t}_ratios.xlsx"
        score = workdir / f"{t}_scorecard.xlsx"

        try:
            if 1 in skip:
                _safe_print("… saltando Paso 1")
            else:
                _run_or_die([sys.executable, "Paso1.py", "--ticker", t, "--out", str(full)], workdir, env_utf8)

            if 2 in skip:
                _safe_print("… saltando Paso 2")
            else:
                _run_or_die([sys.executable, "Paso2.py", "--in", str(full), "--out", str(key)], workdir, env_utf8)

            if 3 in skip:
                _safe_print("… saltando Paso 3")
            else:
                _run_with_fallback([
                    [sys.executable, "Paso3.py", "--in", str(full), "--out", str(ratios)],
                    [sys.executable, "Paso3.py", "--full", str(full), "--key", str(key), "--out", str(ratios)]
                ], workdir, env_utf8)

            if 4 in skip:
                _safe_print("… saltando Paso 4")
            else:
                _run_or_die([sys.executable, "Paso4.py", "--key", str(key), "--full", str(full), "--ticker", t, "--out", str(score)], workdir, env_utf8)

            succeeded.append(t)
        except SystemExit as e:
            _safe_print(f"✖ {t}: se interrumpió con código {e.code}. Continuo con el siguiente…")
        except Exception as e:
            _safe_print(f"✖ {t}: error inesperado: {e}. Continuo con el siguiente…")

        time.sleep(args.sleep)

    _safe_print("\nOK: pipeline terminado.")

    if args.compare and len(succeeded) >= 2:
        try:
            out_path = build_scorecards_compare(workdir, succeeded, args.compare_out, add_ranks=True)
            _safe_print(f"✔ Comparativo de scorecards creado: {out_path}")
        except Exception as e:
            _safe_print(f"ADVERTENCIA: no se pudo crear el comparativo: {e}")
    elif args.compare:
        _safe_print("Nota: comparativo no generado (necesita al menos 2 tickers con scorecard).")

    # === Paso 5 v2 opcional ===
    if args.use_v2:
        try:
            t_list = succeeded if succeeded else tickers
            t_arg = ",".join(t_list) if not isinstance(t_list, str) else t_list
            script_path = Path(__file__).parent / "Paso5_v2.py"
            subprocess.check_call([
                sys.executable, str(script_path),
                "--tickers", t_arg,
                "--out", args.v2_out,
                "--full-out", args.v2_full_out
            ])
            _safe_print(f"✔ Scorecards v2: {args.v2_out} | FULL: {args.v2_full_out}")
        except Exception as e:
            _safe_print(f"ADVERTENCIA: Paso5_v2 no pudo ejecutarse: {e}")

if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
