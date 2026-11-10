"""Merge monthly sales CSV exports, clean them and build a formatted Excel report.

Usage:
    python report.py --input input --output sales_report.xlsx
"""
import argparse
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")


# Exports mix ISO dates (2026-08-01) with Turkish day-first dates (01.08.2026, 01/08/2026).
# A single "dayfirst" parse would silently turn 2026-08-01 into 8 January, so each format is explicit.
DATE_FORMATS = ["%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"]


def parse_dates(s: pd.Series) -> pd.Series:
    out = pd.Series(pd.NaT, index=s.index)
    for fmt in DATE_FORMATS:
        missing = out.isna()
        out[missing] = pd.to_datetime(s[missing].str.strip(), format=fmt, errors="coerce")
    return out


def load_and_clean(folder: Path) -> tuple[pd.DataFrame, dict]:
    files = sorted(folder.glob("*.csv"))
    if not files:
        raise SystemExit(f"No CSV files found in {folder}")
    raw = pd.concat([pd.read_csv(f, sep=";", dtype=str) for f in files], ignore_index=True)
    log = {"files": len(files), "raw_rows": len(raw)}

    df = raw.copy()
    df["Branch"] = df["Branch"].str.strip().str.title()
    df["Unit Price"] = df["Unit Price"].str.replace(",", ".", regex=False).astype(float)
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce")
    df["Date"] = parse_dates(df["Date"])
    log["unparseable_dates_dropped"] = int(df["Date"].isna().sum())
    df = df.dropna(subset=["Date"])

    log["missing_qty_dropped"] = int(df["Qty"].isna().sum())
    df = df.dropna(subset=["Qty"])
    before = len(df)
    # same receipt number exported twice = duplicate; identical sales on the same day are kept
    df = df.drop_duplicates(subset="Receipt No")
    log["duplicates_removed"] = before - len(df)

    df["Qty"] = df["Qty"].astype(int)
    df["Revenue"] = df["Qty"] * df["Unit Price"]
    df["Month"] = df["Date"].dt.to_period("M").astype(str)
    log["clean_rows"] = len(df)
    return df.sort_values("Date"), log


def style_sheet(ws, money_cols=()):
    for cell in ws[1]:
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for col in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col) + 3
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(width, 40)
    for col_letter in money_cols:
        for cell in ws[col_letter][1:]:
            cell.number_format = '#,##0.00 "TL"'
    ws.freeze_panes = "A2"


def build_report(df: pd.DataFrame, log: dict, out: Path):
    by_month = df.pivot_table(index="Month", columns="Branch", values="Revenue", aggfunc="sum").round(2)
    by_product = (df.groupby("Product").agg(Units=("Qty", "sum"), Revenue=("Revenue", "sum"))
                  .sort_values("Revenue", ascending=False).round(2))

    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        by_month.to_excel(xw, sheet_name="Monthly by Branch")
        by_product.to_excel(xw, sheet_name="Products")
        df.drop(columns="Month").to_excel(xw, sheet_name="Clean Data", index=False)
        pd.DataFrame(list(log.items()), columns=["Check", "Value"]).to_excel(
            xw, sheet_name="Data Quality", index=False)

    wb = load_workbook(out)
    ws = wb["Monthly by Branch"]
    style_sheet(ws, money_cols=[get_column_letter(i) for i in range(2, ws.max_column + 1)])
    chart = LineChart()
    chart.title, chart.y_axis.title = "Monthly revenue by branch", "TL"
    chart.add_data(Reference(ws, min_col=2, max_col=ws.max_column, min_row=1, max_row=ws.max_row),
                   titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ws.max_row))
    chart.width, chart.height = 18, 9
    ws.add_chart(chart, f"A{ws.max_row + 3}")

    ws = wb["Products"]
    style_sheet(ws, money_cols=["C"])
    bar = BarChart()
    bar.title = "Revenue by product"
    bar.add_data(Reference(ws, min_col=3, min_row=1, max_row=ws.max_row), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ws.max_row))
    bar.width, bar.height = 18, 9
    ws.add_chart(bar, "E2")

    style_sheet(wb["Clean Data"], money_cols=["F", "G"])
    for cell in wb["Clean Data"]["B"][1:]:
        cell.number_format = "dd.mm.yyyy"
    style_sheet(wb["Data Quality"])
    wb.save(out)


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--input", default="input", type=Path)
    p.add_argument("--output", default="sales_report.xlsx", type=Path)
    a = p.parse_args()
    df, log = load_and_clean(a.input)
    build_report(df, log, a.output)
    print(f"Report written to {a.output}")
    for k, v in log.items():
        print(f"  {k}: {v}")


if __name__ == "__main__":
    main()
