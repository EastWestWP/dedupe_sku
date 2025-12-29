import sys
from pathlib import Path
import pandas as pd

SKU_COL = "Variant SKU"       # change if needed
PRICE_COL = "Variant Price"   # change if needed

def die(msg: str, code: int = 1):
    print(msg)
    input("Press Enter to close...")
    raise SystemExit(code)

def main():
    if len(sys.argv) < 2:
        die("Usage: python dedupe_sku_keep_max_price.py <input.xlsx>")

    input_path = Path(sys.argv[1]).expanduser()

    if not input_path.exists():
        die(f"File not found: {input_path}")

    if input_path.suffix.lower() not in [".xlsx", ".xlsm", ".xls"]:
        die("Input must be an Excel file (.xlsx/.xlsm/.xls)")

    try:
        df = pd.read_excel(input_path, dtype={SKU_COL: str})
    except Exception as e:
        die(f"Failed to read Excel: {e}")

    if SKU_COL not in df.columns:
        die(f"Missing column: '{SKU_COL}'")
    if PRICE_COL not in df.columns:
        die(f"Missing column: '{PRICE_COL}'")

    # Parse price to numeric
    df[PRICE_COL] = pd.to_numeric(df[PRICE_COL], errors="coerce")

    # Sort so max price for each SKU is first
    df_sorted = df.sort_values([SKU_COL, PRICE_COL], ascending=[True, False], kind="mergesort")

    # Keep highest-price row per SKU
    deduped = df_sorted.drop_duplicates(subset=[SKU_COL], keep="first")

    removed = df_sorted.loc[~df_sorted.index.isin(deduped.index)]

    out_base = input_path.with_suffix("")
    output_file = Path(str(out_base) + "_deduped.xlsx")
    removed_file = Path(str(out_base) + "_removed_duplicates.xlsx")

    try:
        deduped.to_excel(output_file, index=False)
        removed.to_excel(removed_file, index=False)
    except Exception as e:
        die(f"Failed to write Excel output: {e}")

    print("Done.")
    print(f"Saved cleaned file: {output_file}")
    print(f"Saved removed rows: {removed_file}")
    input("Press Enter to close...")

if __name__ == "__main__":
    main()
