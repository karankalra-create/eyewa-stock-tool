# Eyewa Stock Upload Builder — KSA & UAE

Turns the raw WMS inventory snapshot into ready-to-upload stock files for Trendyol,
6th Street, Noon, Namshi and Amazon — for both KSA and UAE. Upload one or both regions'
raw files, pick your platforms, download the filled templates in each platform's own
format, one click.

## What it does

1. **Keeps only sellable stock.** A row survives when `BLOCKED` is FALSE, `LOCATION TYPE`
   is BULK or PICK_FACE, and `LOCATION` is not one of that region's own non-sellable bins
   (damaged, missing, expired, refurb, quarantine, 3PL and so on — 17 bins for KSA, a much
   longer racking/cage/quarantine list for UAE). Locations are matched case-insensitively,
   because the WMS stores the same bin in more than one case (`Default`/`default`,
   `WH-EXP`/`wh-exp`).
2. **Merges duplicate SKUs.** Quantities are summed across bins. SKUs are matched
   case-insensitively too — the WMS holds the same prescription-lens SKU (SVAE, SVAP,
   SVAB, SVNB, SVNO, SVAU) in both cases, and treating them as separate silently drops units.
3. **Works out sellable quantity** as `TOTAL` = Available − Reserved, floored at zero.
   Reserved units are already committed to open orders.
4. **Holds back each SKU's threshold** from that region's master file:
   `MAX(sellable − threshold, 0)`.
5. **Writes the platform template** using only the SKUs on that platform's master sheet,
   filling *that region's own template file* (KSA and UAE each have their own — see
   Regions below). A listed SKU with no stock is written as `0` rather than omitted, so
   each upload is a complete refresh. Prices, warehouse codes, processing times, headers,
   dropdown sheets and the Amazon example row are left exactly as the platform supplied them.

## Regions

Each region is fully self-contained: its own excluded-locations list, its own master
item file, and its own copy of every platform template (because Noon, Namshi and Amazon
tag stock to a specific warehouse/project/marketplace-country inside the file itself —
using the wrong region's template would upload the right quantities to the wrong
country's listings).

| | KSA | UAE |
|---|---|---|
| Master file | `reference/master_ksa.xlsx` | `reference/master_uae.xlsx` |
| Templates | `reference/templates/ksa/` | `reference/templates/uae/` |
| Excluded locations | inline in `stock_engine.py` (17 bins) | `reference/excluded_locations_uae.txt` (~800 codes) |
| Noon / Namshi sheet | `Noon KSA` / `KSA` | `Noon UAE` / `Final_UAE` |

Upload the KSA raw file, the UAE raw file, or both — the tool only builds outputs for
the regions you give it a file for. Each region's downloads are labelled
`<Platform>_<Region>_stock_<date>.<ext>`, e.g. `Trendyol_UAE_stock_2026-09-21.xlsx`.

## Platform mapping

| Platform | SKU source | Quantity column | Output |
|---|---|---|---|
| Trendyol | master `Trendyol` sheet → eyewa SKU | `Stock` (col D) | `.xlsx` |
| 6th Street | master `6th Street` sheet → eyewa SKU | `Count` | `.csv` |
| Noon | master `Noon` sheet, `partner_sku` is the eyewa SKU | `stock_gross` (col E) | `.xlsx` |
| Namshi | master `Namshi` sheet, `partner_sku` is the eyewa SKU | `stock_gross` (col D) | `.xlsx` |
| Amazon | master `amazon` sheet → eyewa SKU | `Quantity (SA)` / `Quantity (AE)` (col C) | `.xlsm` |

Column positions are the same in both regions' templates — only the sheet name (Noon,
Namshi) and the embedded warehouse/country tags differ. Thresholds are read per SKU from
each region's master file `Threshold` column, so updating a master file is enough to
change them — no code change needed.

## Running it locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Updating a master item file

Two options, per region:

- **Permanent:** replace `reference/master_ksa.xlsx` or `reference/master_uae.xlsx` and
  commit. Everyone gets it.
- **One-off:** use that region's *Advanced* uploader in the sidebar to override the
  built-in master for a single run.

The master file needs five sheets — `Trendyol`, `6th Street`, `amazon`, `Namshi`, `Noon`
— the same sheet names in both regions. Each carries the platform SKU in the first
column, then `Brand`, `Category` and `Threshold`; Trendyol, 6th Street and amazon also
carry an `eyewa SKUs` column.

## Replacing a platform template

Drop the new file into `reference/templates/ksa/` or `reference/templates/uae/` under the
same name. If the platform moves its quantity column, update that platform's entry in
`PLATFORMS` in `stock_engine.py` — `sku_col`, `qty_col` and `first_row` are all that
define it (shared across regions). If a region's template renames its worksheet, update
that region's `sheet_overrides` in `REGIONS` instead of touching `PLATFORMS`.

## Adding another region

1. Add its excluded-locations to a new `reference/excluded_locations_<region>.txt` (or
   inline in `stock_engine.py` if it's short, like KSA's).
2. Drop its master file at `reference/master_<region>.xlsx` and its five templates under
   `reference/templates/<region>/`.
3. Add an entry to `REGIONS` in `stock_engine.py` (and to `REGION_ORDER`) pointing at them,
   with `sheet_overrides` for any platform whose worksheet is named differently there.
4. The sidebar picks up the new region automatically — no `app.py` changes needed.

## Known data trap — 6th Street barcodes

The 6th Street CSV export mangles 13-digit barcode SKUs into scientific notation
(`8.80954E+12`), which is lossy: eleven distinct barcodes collapse to one string. This tool
sidesteps it by building the 6th Street file from the master sheet, which holds the true
barcodes, so the output never inherits the corruption — the `Sku` column in the generated
CSV always holds clean plain digits (e.g. `6290360089133`), never scientific notation or a
trailing `.0`.

That said, if you open the generated CSV by double-clicking it in Excel, Excel will
auto-convert those same long digit strings back into scientific notation for **display only**
— this is a well-known Excel CSV quirk and has no CSV-only fix, since plain CSV carries no
per-cell formatting. It does not change the underlying file, and it does not affect what
6th Street's importer reads (it parses the raw text, not Excel's rendering). To inspect the
file without triggering this, use Excel's **Data → From Text/CSV** and set the `Sku` column
type to **Text** before loading, instead of opening it directly.

## Layout

```
app.py                              Streamlit UI
stock_engine.py                     all the logic — regions, filters, thresholds, template writing
reference/master_ksa.xlsx           KSA master item list, one sheet per platform
reference/master_uae.xlsx           UAE master item list, one sheet per platform
reference/excluded_locations_uae.txt UAE's non-sellable bin list (KSA's is inline, it's short)
reference/templates/ksa/            KSA's five platform templates
reference/templates/uae/            UAE's five platform templates
```
