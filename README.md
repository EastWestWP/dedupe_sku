# SKU Deduper (Keep Highest Price)

This tool cleans a store product Excel export by removing duplicate SKUs.
**Rule:** if multiple rows share the same SKU, keep the row with the **highest price**.

Outputs:
- `<input>_deduped.xlsx` (cleaned file)
- `<input>_removed_duplicates.xlsx` (rows removed as duplicates)

---

## Folder Structure

Put these files in the same folder:

```
dedupe_sku/
  run_dedupe_venv.bat
  dedupe_sku_keep_max_price.py
  requirements.txt
  products.xlsx   (your input file)
```

---

## Requirements

- Windows 10/11
- Python 3.x installed (must have a working `python.exe`)
  - If `python` is not on PATH, the `.bat` will ask you to paste the full path to `python.exe`

---

## Install / Setup (One-time)

No manual install needed.

On first run, the script will:
1. Create a virtual environment in `./venv`
2. Install dependencies into that venv

Dependencies are defined in `requirements.txt`.

---

## How to Run (One click)

1. Double-click: `run_dedupe_venv.bat`
2. Enter your Excel filename (example: `products.xlsx`)  
   or paste a full path (example: `E:\Exports\products.xlsx`)
3. Wait for it to finish
4. Check the folder for output files:
   - `products_deduped.xlsx`
   - `products_removed_duplicates.xlsx`

---

## If the BAT closes immediately

This `.bat` includes `pause` and prints errors.  
If you still need more details:
1. Press `Win + R`
2. Type `cmd` and press Enter
3. Drag `run_dedupe_venv.bat` into the window and press Enter
4. Copy/paste the output to debug

---

## Common Issues

### 1) Python launcher / PATH issues
Symptom:
- Error like: `system cannot find the file specified`

Fix:
- Reinstall/repair Python from python.org, or paste the correct full path to `python.exe` when prompted.

### 2) Missing columns error (SKU / Price)
Symptom:
- Script prints: `Missing column: 'Variant SKU'` or `Missing column: 'Variant Price'`

Fix:
- Update the column names in `dedupe_sku_keep_max_price.py`:
  ```python
  SKU_COL = "Variant SKU"
  PRICE_COL = "Variant Price"
  ```

---

## Column Names Assumed

By default the script expects:
- SKU column: `Variant SKU`
- Price column: `Variant Price`

You can change them at the top of `dedupe_sku_keep_max_price.py`.

---

## Notes on Tie-Breakers

If duplicate SKUs have the **same** price, the script keeps the first row after sorting
(stable sort). If you need a different tie-breaker (e.g., keep row with the most inventory,
latest updated date, non-empty image, etc.), update the sorting logic.
