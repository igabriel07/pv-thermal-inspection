#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
tiff_to_pixel_temperatures_csv.py

Διαβάζει ένα thermal TIFF (π.χ. DJI M3T .tiff που έχει float32 θερμοκρασίες)
και γράφει CSV με θερμοκρασία ανά pixel.

Modes:
1) long (default): 1 γραμμή ανά pixel: row,col,temp
   -> τεράστιο αρχείο για 640x512 (~327,680 γραμμές) αλλά πολύ πρακτικό για GIS/analysis.

2) wide: 1 γραμμή ανά row, 640 στήλες (ή όσο το πλάτος):
   -> πιο "matrix-like", αλλά επίσης μεγάλο. Πιο εύκολο να το ανοίξεις σε numpy/pandas.

3) long_with_xy: 1 γραμμή ανά pixel με x,y (όπου x=col, y=row)

ΠΡΟΣΟΧΗ:
- Αν το TIFF δεν είναι float (π.χ. uint16 radiometric), το script θα κάνει export τις raw τιμές.
  (Δεν κάνει calibration χωρίς extra metadata/planck params.)

Απαιτήσεις:
  pip install tifffile numpy

Χρήση:
  python tiff_to_pixel_temperatures_csv.py input.tiff
  python tiff_to_pixel_temperatures_csv.py input.tiff --out temps.csv
  python tiff_to_pixel_temperatures_csv.py input.tiff --mode wide
  python tiff_to_pixel_temperatures_csv.py input.tiff --sample 10   (κρατάει κάθε 10ο pixel για μικρότερο CSV)
"""

import argparse
import csv
import os
from typing import Optional, Tuple

import numpy as np

try:
    import tifffile
except Exception as e:
    tifffile = None


def infer_default_out(path: str, mode: str) -> str:
    base = os.path.splitext(path)[0]
    suffix = ".pixel_temps.long.csv" if mode.startswith("long") else ".pixel_temps.wide.csv"
    return base + suffix


def read_array(path: str) -> np.ndarray:
    if tifffile is None:
        raise RuntimeError("tifffile not available. Install with: pip install tifffile")
    arr = tifffile.imread(path)
    if arr.ndim != 2:
        raise ValueError(f"Expected 2D image, got shape={arr.shape}")
    return arr


def write_long_csv(
    arr: np.ndarray,
    out_csv: str,
    include_xy: bool = False,
    sample_step: int = 1,
    nan_as_empty: bool = True,
) -> Tuple[int, int]:
    """
    Writes row,col,temp (and optionally x,y) for each pixel.
    Returns (rows_written, cols_per_row)
    """
    h, w = arr.shape
    sample_step = max(1, int(sample_step))

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f)

        if include_xy:
            wr.writerow(["x", "y", "row", "col", "temp"])
        else:
            wr.writerow(["row", "col", "temp"])

        written = 0
        # iterate rows
        for r in range(0, h, sample_step):
            row_vals = arr[r]
            for c in range(0, w, sample_step):
                v = row_vals[c]
                if isinstance(v, np.generic):
                    v = v.item()

                if nan_as_empty and isinstance(v, float) and (np.isnan(v) or np.isinf(v)):
                    v_out = ""
                else:
                    v_out = v

                if include_xy:
                    wr.writerow([c, r, r, c, v_out])
                else:
                    wr.writerow([r, c, v_out])
                written += 1

    return written, w


def write_wide_csv(
    arr: np.ndarray,
    out_csv: str,
    sample_step: int = 1,
    nan_as_empty: bool = True,
    header: bool = True,
) -> Tuple[int, int]:
    """
    Writes one row per image row, values across columns.
    Returns (rows_written, cols_written)
    """
    h, w = arr.shape
    sample_step = max(1, int(sample_step))

    cols = list(range(0, w, sample_step))

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f)

        if header:
            wr.writerow(["row"] + [f"c{c}" for c in cols])

        rows_written = 0
        for r in range(0, h, sample_step):
            vals = []
            row_vals = arr[r]
            for c in cols:
                v = row_vals[c]
                if isinstance(v, np.generic):
                    v = v.item()
                if nan_as_empty and isinstance(v, float) and (np.isnan(v) or np.isinf(v)):
                    vals.append("")
                else:
                    vals.append(v)
            wr.writerow([r] + vals)
            rows_written += 1

    return rows_written, len(cols)


def main():
    ap = argparse.ArgumentParser(description="Export per-pixel temperatures from a 2D TIFF to CSV.")
    ap.add_argument("tiff", help="Input .tif/.tiff file (2D)")
    ap.add_argument("--out", default=None, help="Output CSV path")
    ap.add_argument(
        "--mode",
        default="long",
        choices=["long", "long_with_xy", "wide"],
        help="CSV format: long (row,col,temp), long_with_xy (x,y,row,col,temp), wide (row + many columns)",
    )
    ap.add_argument(
        "--sample",
        type=int,
        default=1,
        help="Keep every Nth pixel in rows/cols (1=all pixels, 2=every 2nd pixel, etc.)",
    )
    ap.add_argument(
        "--nan-empty",
        action="store_true",
        help="Write NaN/Inf as empty cell (default: on). Use --no-nan-empty to keep literal values.",
    )
    ap.add_argument(
        "--no-nan-empty",
        action="store_true",
        help="Keep NaN/Inf as literal values (overrides --nan-empty).",
    )
    args = ap.parse_args()

    if not os.path.isfile(args.tiff):
        raise SystemExit(f"File not found: {args.tiff}")

    arr = read_array(args.tiff)

    # Determine NaN handling
    nan_as_empty = True
    if args.no_nan_empty:
        nan_as_empty = False
    elif args.nan_empty:
        nan_as_empty = True

    out_csv = args.out or infer_default_out(args.tiff, args.mode)

    if args.mode == "wide":
        rows, cols = write_wide_csv(arr, out_csv, sample_step=args.sample, nan_as_empty=nan_as_empty)
        print(f"Wrote CSV (wide): {out_csv}")
        print(f"Rows written: {rows}, Cols per row (sampled): {cols}")
    else:
        include_xy = (args.mode == "long_with_xy")
        written, w = write_long_csv(arr, out_csv, include_xy=include_xy, sample_step=args.sample, nan_as_empty=nan_as_empty)
        print(f"Wrote CSV (long): {out_csv}")
        print(f"Pixels written: {written} (sample step={args.sample}), Image width={w}")

    print(f"Input dtype: {arr.dtype}, shape: {arr.shape}")
    try:
        a = arr.astype(np.float64)
        mn, av, mx = float(np.nanmin(a)), float(np.nanmean(a)), float(np.nanmax(a))
        print(f"Stats (min/mean/max): {mn:.3f} / {av:.3f} / {mx:.3f}")
    except Exception:
        pass


if __name__ == "__main__":
    main()
