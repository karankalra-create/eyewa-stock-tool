"""
Eyewa KSA marketplace stock upload engine.

Pure logic, no Streamlit imports, so it can be unit-tested or driven from a script.

Pipeline
--------
raw WMS inventory snapshot
  -> sellable filter (blocked / location type / excluded locations)
  -> merge duplicate SKUs (case-insensitive)
  -> per platform: master-file SKU list + per-SKU threshold
  -> filled platform template, byte-for-byte in the platform's own format
"""

from __future__ import annotations

import io
import os
import re
from dataclasses import dataclass, field

import openpyxl
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
REFERENCE_DIR = os.path.join(HERE, "reference")
TEMPLATE_DIR = os.path.join(REFERENCE_DIR, "templates")
DEFAULT_MASTER = os.path.join(REFERENCE_DIR, "master_file.xlsx")

# --------------------------------------------------------------------------
# Sellable-stock rules
# --------------------------------------------------------------------------

#: Only these location types hold sellable stock.
SELLABLE_LOCATION_TYPES = {"BULK", "PICK_FACE"}

#: Locations that never hold sellable stock, regardless of location type.
#: Matched case-insensitively -- the WMS holds case variants of the same bin
#: (Default/default, MIS-RIY/mis-riy, WH-EXP/wh-exp, ...).
EXCLUDED_LOCATIONS = {
    "DEFAULT", "MIS-RIY", "OF-07-5-5", "OF-07-7-9", "QSCRIY",
    "RF-04-1-1", "RF-04-3-1", "RF-05-2-1", "RF-05-3-1",
    "WH-EXP", "WH-MISS", "WH-OTH-DAM", "WH-RE-REFURB", "WMS-DF",
    "WB-1-1-1", "WH-3PL", "WH-SUP-DAM",
}


def _norm(value) -> str:
    """Normalise a SKU / location for matching: trimmed, lower-cased."""
    return str(value).strip().lower()


def _norm_header(value) -> str:
    """Normalise a column header so 'Available Qty', 'AVAILABLE QTY' etc. all match."""
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())


# --------------------------------------------------------------------------
# Platform configuration
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Platform:
    key: str                     # internal id
    label: str                   # shown in the UI
    master_sheet: str            # sheet name in the master file
    template: str                # bundled template filename
    out_ext: str                 # extension of the produced file
    sheet: str | None = None     # worksheet that holds the upload rows
    sku_col: int | None = None   # 1-based column holding the platform SKU
    qty_col: int | None = None   # 1-based column that receives the quantity
    raw_col: int | None = None   # optional helper column receiving pre-threshold qty
    first_row: int = 2           # first data row
    #: master column holding the eyewa SKU; None means the platform SKU *is* the eyewa SKU
    eyewa_col: str | None = "eyewa SKUs"


PLATFORMS: dict[str, Platform] = {
    "trendyol": Platform(
        key="trendyol", label="Trendyol", master_sheet="Trendyol",
        template="trendyol_template.xlsx", out_ext=".xlsx",
        sheet="Stock & Price Update", sku_col=1, qty_col=4, first_row=2,
    ),
    "sixth_street": Platform(
        key="sixth_street", label="6th Street", master_sheet="6th Street",
        template="sixth_street_template.csv", out_ext=".csv",
    ),
    "noon": Platform(
        key="noon", label="Noon", master_sheet="Noon",
        template="noon_template.xlsx", out_ext=".xlsx",
        sheet="Noon KSA", sku_col=3, qty_col=5, raw_col=8, first_row=2,
        eyewa_col=None,
    ),
    "namshi": Platform(
        key="namshi", label="Namshi", master_sheet="Namshi",
        template="namshi_template.xlsx", out_ext=".xlsx",
        sheet="KSA", sku_col=2, qty_col=4, raw_col=6, first_row=2,
        eyewa_col=None,
    ),
    "amazon": Platform(
        key="amazon", label="Amazon", master_sheet="amazon",
        template="amazon_template.xlsm", out_ext=".xlsm",
        sheet="Template", sku_col=1, qty_col=3, first_row=7,
    ),
}

PLATFORM_ORDER = ["trendyol", "sixth_street", "noon", "namshi", "amazon"]


# --------------------------------------------------------------------------
# Step 1 -- raw stock file -> sellable qty per SKU
# --------------------------------------------------------------------------

REQUIRED_RAW_COLUMNS = {
    "blocked": "BLOCKED",
    "location": "LOCATION",
    "locationtype": "LOCATION TYPE",
    "sku": "SKU",
}


class StockFileError(ValueError):
    """Raised when the uploaded raw stock file cannot be interpreted."""


def read_raw_stock(file) -> pd.DataFrame:
    """Read an uploaded raw WMS snapshot (.csv / .xlsx / .xlsm) into a DataFrame."""
    name = getattr(file, "name", str(file)).lower()
    if name.endswith((".xlsx", ".xlsm", ".xls")):
        df = pd.read_excel(file, dtype=str)
    else:
        df = pd.read_csv(file, dtype=str)
    if df.empty:
        raise StockFileError("The stock file is empty.")
    return df


@dataclass
class SellableResult:
    qty: dict[str, int]                 # normalised SKU -> sellable units
    rows_in: int = 0
    rows_after_blocked: int = 0
    rows_after_loctype: int = 0
    rows_after_location: int = 0
    skus_out: int = 0
    units_available: int = 0
    units_reserved: int = 0
    units_sellable: int = 0
    merged_case_variants: int = 0
    qty_basis: str = ""


def build_sellable(df: pd.DataFrame) -> SellableResult:
    """Apply the sellable filter and merge duplicate SKUs into one row each."""
    lookup = {_norm_header(c): c for c in df.columns}
    missing = [
        label for key, label in REQUIRED_RAW_COLUMNS.items() if key not in lookup
    ]
    if missing:
        raise StockFileError(
            "The stock file is missing these columns: "
            + ", ".join(missing)
            + ". Found: "
            + ", ".join(str(c) for c in df.columns)
        )

    col = {k: lookup[k] for k in REQUIRED_RAW_COLUMNS}
    has_total = "total" in lookup
    has_avail = "availableqty" in lookup
    has_resv = "reservedqty" in lookup

    if not (has_total or has_avail):
        raise StockFileError(
            "The stock file needs either a TOTAL column, or AVAILABLE QTY "
            "(and ideally RESERVED QTY) so sellable units can be worked out."
        )

    work = df.copy()
    res = SellableResult(qty={}, rows_in=len(work))

    # --- the three sellable rules -----------------------------------------
    work = work[work[col["blocked"]].astype(str).str.strip().str.upper() == "FALSE"]
    res.rows_after_blocked = len(work)

    work = work[
        work[col["locationtype"]].astype(str).str.strip().str.upper()
        .isin(SELLABLE_LOCATION_TYPES)
    ]
    res.rows_after_loctype = len(work)

    work = work[
        ~work[col["location"]].astype(str).str.strip().str.upper()
        .isin(EXCLUDED_LOCATIONS)
    ]
    res.rows_after_location = len(work)

    if work.empty:
        raise StockFileError(
            "No sellable rows survived the filter. Check that BLOCKED, "
            "LOCATION TYPE and LOCATION hold the values this tool expects."
        )

    # --- quantity basis: Total = Available - Reserved ----------------------
    def num(series):
        return pd.to_numeric(series, errors="coerce").fillna(0).astype(int)

    if has_avail:
        avail = num(work[lookup["availableqty"]])
        resv = num(work[lookup["reservedqty"]]) if has_resv else 0
        res.units_available = int(avail.sum())
        res.units_reserved = int(resv.sum()) if has_resv else 0
    if has_total:
        qty = num(work[lookup["total"]])
        res.qty_basis = "TOTAL column (Available - Reserved)"
    else:
        qty = avail - resv
        res.qty_basis = "AVAILABLE QTY - RESERVED QTY"
    qty = qty.clip(lower=0)

    # --- merge duplicate SKUs, case-insensitively --------------------------
    keys = work[col["sku"]].astype(str).str.strip().str.lower()
    res.merged_case_variants = int(
        work[col["sku"]].astype(str).str.strip().nunique() - keys.nunique()
    )
    grouped = qty.groupby(keys).sum()

    res.qty = {str(k): int(v) for k, v in grouped.items()}
    res.skus_out = len(res.qty)
    res.units_sellable = int(sum(res.qty.values()))
    return res


# --------------------------------------------------------------------------
# Step 2 -- master file
# --------------------------------------------------------------------------


def load_master(source=None) -> dict[str, pd.DataFrame]:
    """Load the master item file, one DataFrame per platform sheet."""
    source = source or DEFAULT_MASTER
    out: dict[str, pd.DataFrame] = {}
    for key in PLATFORM_ORDER:
        p = PLATFORMS[key]
        try:
            df = pd.read_excel(source, sheet_name=p.master_sheet, dtype=str)
        except ValueError as exc:
            raise StockFileError(
                f"The master file has no '{p.master_sheet}' sheet (needed for {p.label})."
            ) from exc
        df = df.dropna(how="all")
        if "Threshold" not in df.columns:
            raise StockFileError(
                f"The '{p.master_sheet}' sheet has no Threshold column."
            )
        out[key] = df
        if hasattr(source, "seek"):
            source.seek(0)
    return out


@dataclass
class PlatformPlan:
    """What each platform SKU should end up with."""
    platform: Platform
    rows: dict[str, int] = field(default_factory=dict)   # norm platform sku -> final qty
    raw_rows: dict[str, int] = field(default_factory=dict)  # norm platform sku -> pre-threshold
    skus: int = 0
    matched: int = 0
    units_before: int = 0
    units_after: int = 0
    live_skus: int = 0


def plan_platform(key: str, master: pd.DataFrame, sellable: dict[str, int]) -> PlatformPlan:
    """Work out the final quantity for every SKU the master file lists."""
    p = PLATFORMS[key]
    plan = PlatformPlan(platform=p)

    sku_col = master.columns[0]
    eyewa_col = p.eyewa_col if p.eyewa_col in master.columns else sku_col

    for _, row in master.iterrows():
        platform_sku = row[sku_col]
        if pd.isna(platform_sku):
            continue
        eyewa_sku = row[eyewa_col] if not pd.isna(row[eyewa_col]) else platform_sku
        raw = sellable.get(_norm(eyewa_sku), 0)
        try:
            threshold = int(float(row["Threshold"]))
        except (TypeError, ValueError):
            threshold = 0
        final = max(raw - threshold, 0)

        k = _norm(platform_sku)
        plan.rows[k] = final
        plan.raw_rows[k] = raw
        plan.skus += 1
        plan.matched += 1 if _norm(eyewa_sku) in sellable else 0
        plan.units_before += raw
        plan.units_after += final
        plan.live_skus += 1 if final > 0 else 0

    return plan


# --------------------------------------------------------------------------
# Step 3 -- write the platform template, format untouched
# --------------------------------------------------------------------------


def _template_path(p: Platform) -> str:
    return os.path.join(TEMPLATE_DIR, p.template)


def _build_csv(plan: PlatformPlan, master: pd.DataFrame) -> bytes:
    """6th Street: a plain two-column Sku,Count file."""
    sku_col = master.columns[0]
    rows = []
    for _, row in master.iterrows():
        sku = row[sku_col]
        if pd.isna(sku):
            continue
        sku = str(sku).strip()
        rows.append({"Sku": sku, "Count": plan.rows.get(_norm(sku), 0)})
    buf = io.StringIO()
    pd.DataFrame(rows, columns=["Sku", "Count"]).to_csv(buf, index=False)
    return buf.getvalue().encode("utf-8")


def _build_workbook(plan: PlatformPlan) -> bytes:
    """
    Open the real platform template, drop rows the master file does not list,
    and write the quantity into the platform's own quantity column. Every other
    cell, sheet and piece of formatting is left exactly as the platform sent it.
    """
    p = plan.platform
    keep_vba = p.template.endswith(".xlsm")
    wb = openpyxl.load_workbook(_template_path(p), keep_vba=keep_vba)
    ws = wb[p.sheet]

    drop: list[int] = []
    for r in range(p.first_row, ws.max_row + 1):
        sku = ws.cell(r, p.sku_col).value
        if sku is None or str(sku).strip() == "":
            continue
        k = _norm(sku)
        if k not in plan.rows:
            drop.append(r)
            continue
        ws.cell(r, p.qty_col).value = plan.rows[k]
        if p.raw_col:
            ws.cell(r, p.raw_col).value = plan.raw_rows.get(k, 0)

    # delete bottom-up so earlier indices stay valid; group runs for speed
    for r in sorted(drop, reverse=True):
        ws.delete_rows(r)

    buf = io.BytesIO()
    wb.save(buf)
    wb.close()
    return buf.getvalue()


def build_output(key: str, plan: PlatformPlan, master: pd.DataFrame) -> bytes:
    p = PLATFORMS[key]
    if p.out_ext == ".csv":
        return _build_csv(plan, master)
    return _build_workbook(plan)


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------


def run(raw_df: pd.DataFrame, platform_keys: list[str], master_source=None):
    """Full pipeline. Returns (SellableResult, {key: (PlatformPlan, bytes)})."""
    sellable = build_sellable(raw_df)
    master = load_master(master_source)
    results: dict[str, tuple[PlatformPlan, bytes]] = {}
    for key in platform_keys:
        plan = plan_platform(key, master[key], sellable.qty)
        results[key] = (plan, build_output(key, plan, master[key]))
    return sellable, results


def sellable_csv(sellable: SellableResult) -> bytes:
    df = pd.DataFrame(
        sorted(sellable.qty.items()), columns=["SKU", "SELLABLE QTY"]
    )
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    return buf.getvalue().encode("utf-8")
