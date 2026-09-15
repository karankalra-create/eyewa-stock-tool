# Eyewa Stock Upload Builder — KSA

Turns the raw WMS inventory snapshot into ready-to-upload stock files for Trendyol,
6th Street, Noon, Namshi and Amazon. Upload one file, pick your platforms, download
the filled templates in each platform's own format.

## What it does

1. **Keeps only sellable stock.** A row survives when `BLOCKED` is FALSE, `LOCATION TYPE`
   is BULK or PICK_FACE, and `LOCATION` is not one of the 17 non-sellable bins (damaged,
   missing, expired, refurb, quarantine, 3PL and so on). Locations are matched
   case-insensitively, because the WMS stores the same bin in more than one case
   (`Default`/`default`, `WH-EXP`/`wh-exp`).
2. **Merges duplicate SKUs.** Quantities are summed across bins. SKUs are matched
   case-insensitively too — the WMS holds the same prescription-lens SKU (SVAE, SVAP,
   SVAB, SVNB, SVNO, SVAU) in both cases, and treating them as separate silently drops units.
3. **Works out sellable quantity** as `TOTAL` = Available − Reserved, floored at zero.
   Reserved units are already committed to open orders.
4. **Holds back each SKU's threshold** from the master file: `MAX(sellable − threshold, 0)`.
5. **Writes the platform template** using only the SKUs on that platform's master sheet.
   A listed SKU with no stock is written as `0` rather than omitted, so each upload is a
   complete refresh. Prices, warehouse codes, processing times, headers, dropdown sheets
   and the Amazon example row are left exactly as the platform supplied them.

## Platform mapping

| Platform | SKU source | Quantity column | Output |
|---|---|---|---|
| Trendyol | master `Trendyol` sheet → eyewa SKU | `Stock` (col D) | `.xlsx` |
| 6th Street | master `6th Street` sheet → eyewa SKU | `Count` | `.csv` |
| Noon | master `Noon` sheet, `partner_sku` is the eyewa SKU | `stock_gross` (col E) | `.xlsx` |
| Namshi | master `Namshi` sheet, `partner_sku` is the eyewa SKU | `stock_gross` (col D) | `.xlsx` |
| Amazon | master `amazon` sheet → eyewa SKU | `Quantity (SA)` (col C) | `.xlsm` |

Thresholds are read per SKU from the master file's `Threshold` column, so updating the
master is enough to change them — no code change needed.

## Running it locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Updating the master item file

Two options:

- **Permanent:** replace `reference/master_file.xlsx` and commit. Everyone gets it.
- **One-off:** use the *Advanced* uploader in the sidebar to override the built-in master
  for a single run.

The master file needs five sheets — `Trendyol`, `6th Street`, `amazon`, `Namshi`, `Noon`.
Each carries the platform SKU in the first column, then `Brand`, `Category` and `Threshold`;
Trendyol, 6th Street and amazon also carry an `eyewa SKUs` column.

## Replacing a platform template

Drop the new file into `reference/templates/` under the same name. If the platform moves its
quantity column or renames its sheet, update that platform's entry in `PLATFORMS` in
`stock_engine.py` — `sheet`, `sku_col`, `qty_col` and `first_row` are all that define it.

## Known data trap — 6th Street barcodes

The 6th Street CSV export mangles 13-digit barcode SKUs into scientific notation
(`8.80954E+12`), which is lossy: eleven distinct barcodes collapse to one string. This tool
sidesteps it by building the 6th Street file from the master sheet, which holds the true
barcodes, so the output never inherits the corruption.

## Layout

```
app.py                        Streamlit UI
stock_engine.py               all the logic — filters, thresholds, template writing
reference/master_file.xlsx    master item list, one sheet per platform
reference/templates/          the five platform templates, used as the output format
```
