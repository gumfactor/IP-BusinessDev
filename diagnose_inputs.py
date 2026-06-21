#!/usr/bin/env python3
"""
diagnose_inputs.py — Run this before anything else.

Point it at your downloaded CIPO / TSX files and it will print:
  - Every column name found in each file
  - The first 3 rows of each file
  - File size and row count

This tells us the exact column names we need to hardcode into
build_prospect_list.py so it doesn't guess wrong.

Usage examples:

  # Inspect a single file:
  python diagnose_inputs.py somefile.csv

  # Inspect a whole folder of files (e.g. unzipped CIPO download):
  python diagnose_inputs.py path/to/cipo_folder/

  # Inspect several specific files:
  python diagnose_inputs.py TM_Main.csv TM_Interested_Party.csv PT_applicant.csv

"""

import sys
import os
from pathlib import Path

# ── Try to import pandas; give a clear error if it's missing ──────────────────
try:
    import pandas as pd
except ImportError:
    print("\nERROR: pandas is not installed. Run:")
    print("    pip install pandas openpyxl")
    sys.exit(1)


SUPPORTED_EXTENSIONS = {".csv", ".txt", ".xlsx", ".xls"}


def inspect_file(path: Path) -> None:
    size_mb = path.stat().st_size / (1024 * 1024)
    print(f"\n{'='*70}")
    print(f"FILE: {path.name}  ({size_mb:.1f} MB)")
    print(f"{'='*70}")

    try:
        ext = path.suffix.lower()
        if ext in (".xlsx", ".xls"):
            df = pd.read_excel(path, nrows=200)
        elif ext in (".csv", ".txt"):
            # Try pipe-delimited first (CIPO IP Horizons format), then auto-detect.
            # sep=None auto-detection fails on CIPO files because French names contain
            # commas (e.g. "HAMILTON, W.") which fool the sniffer into picking comma.
            df = None
            for enc in ("utf-8", "latin-1", "cp1252", "utf-16", "utf-16-le"):
                for sep, engine in (("|" , "c"), (None, "python")):
                    try:
                        candidate = pd.read_csv(
                            path,
                            nrows=200,
                            encoding=enc,
                            on_bad_lines="skip",
                            sep=sep,
                            engine=engine,
                        )
                        # Prefer the parse that gives the most columns
                        if df is None or len(candidate.columns) > len(df.columns):
                            df = candidate
                            if len(df.columns) > 5:
                                break   # good enough
                    except (UnicodeDecodeError, Exception):
                        continue
                if df is not None and len(df.columns) > 5:
                    break
            if df is None:
                print("  Could not decode this file with any encoding tried.")
                return
        else:
            print(f"  Skipping unsupported extension: {ext}")
            return
    except Exception as e:
        print(f"  ERROR loading file: {e}")
        return

    print(f"\nCOLUMNS ({len(df.columns)} total):")
    for i, col in enumerate(df.columns, 1):
        sample_values = df[col].dropna().head(3).tolist()
        sample_str = " | ".join(str(v)[:60] for v in sample_values)
        print(f"  {i:3}. {col!r:<45} e.g. {sample_str}")

    print(f"\nFIRST 3 ROWS (truncated to 80 chars per cell):")
    display = df.head(3).astype(str).apply(lambda c: c.str[:80])
    for i, row in display.iterrows():
        print(f"\n  Row {i}:")
        for col, val in row.items():
            if val not in ("nan", "None", ""):
                print(f"    {col}: {val}")

    # Estimate full row count without loading everything
    print(f"\nROW COUNT ESTIMATE: ", end="")
    try:
        if ext in (".csv", ".txt"):
            with open(path, "rb") as f:
                lines = sum(1 for _ in f)
            print(f"~{lines - 1:,} rows (line count minus header)")
        else:
            full = pd.read_excel(path)
            print(f"{len(full):,} rows")
    except Exception as e:
        print(f"(could not count: {e})")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    paths_to_check: list[Path] = []

    for arg in sys.argv[1:]:
        p = Path(arg)
        if not p.exists():
            print(f"WARNING: '{arg}' does not exist — skipping.")
            continue
        if p.is_dir():
            found = [f for f in sorted(p.iterdir())
                     if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS]
            if not found:
                print(f"WARNING: No CSV/TXT/XLSX files found in '{arg}'.")
            paths_to_check.extend(found)
        else:
            paths_to_check.append(p)

    if not paths_to_check:
        print("No valid files found. Check your paths and try again.")
        sys.exit(1)

    print(f"\nInspecting {len(paths_to_check)} file(s)...\n")
    for p in paths_to_check:
        inspect_file(p)

    print(f"\n{'='*70}")
    print("Done. Copy and paste this output back to Claude.")
    print(f"{'='*70}\n")


if __name__ == "__main__":
    main()
