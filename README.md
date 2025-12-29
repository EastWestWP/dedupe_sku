# SKU Deduper (Keep Highest Price)

This tool cleans a store product Excel export by removing duplicate SKUs.
**Rule:** if multiple rows share the same SKU, keep the row with the **highest price**.

It also prevents Excel from showing large numeric IDs in scientific notation by forcing all **ID-like columns** (any header containing `id`) to be written as **TEXT**.

The cleaned output is **sorted by `Title`** (to match typical product exports).

Outputs:
- `<input>_deduped.xlsx` (cleaned file)
- `<input>_removed_duplicates.xlsx` (rows removed as duplicates)

---

## Clone the repo

```bash
git clone https://github.com/EastWestWP/dedupe_sku.git
cd dedupe_sku
```
---

## Quick Start (Windows)

1. Put these files in one folder:
   - `run_dedupe_venv.bat`
   - `dedupe_sku_keep_max_price.py`
   - `requirements.txt`
   - your Excel export (e.g. `products.xlsx`)

2. Double-click `run_dedupe_venv.bat`
3. Enter the Excel filename or full path

On first run it will create `./venv` and install dependencies automatically.

---

## Notes on IDs and Scientific Notation

If your original Excel shows values like `7.5928E+12`, Excel is only displaying the number in scientific notation.
This tool writes ID-like columns as TEXT so the output shows full digits (e.g. `7592857632974`) and does not truncate.

By default, any column whose header contains `id` is treated as text (plus a small allowlist: `ID`, `Variant ID`, etc.).

---

## Output Sorting

By default the cleaned file is sorted by the `Title` column.

If your file uses a different column name, update `TITLE_COL` at the top of `dedupe_sku_keep_max_price.py`.

---

## Column Names Assumed

- SKU column: `Variant SKU`
- Price column: `Variant Price`
- Title column: `Title`

If your export uses different headers, edit the top of `dedupe_sku_keep_max_price.py`.
