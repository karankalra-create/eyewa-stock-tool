"""Eyewa marketplace stock upload builder — Streamlit front end. Multi-region (KSA, UAE)."""

from __future__ import annotations

import io
import zipfile
from datetime import date

import pandas as pd
import streamlit as st

import stock_engine as E

st.set_page_config(
    page_title="Eyewa Stock Upload Builder",
    page_icon="📦",
    layout="wide",
)

# --------------------------------------------------------------------------
# styling
# --------------------------------------------------------------------------
st.markdown(
    """
    <style>
      .block-container {padding-top: 2.4rem; max-width: 1180px;}
      h1 {font-size: 1.85rem !important; font-weight: 650; letter-spacing: -0.02em;}
      .lede {color: #5b6472; font-size: 0.95rem; margin: -0.5rem 0 1.6rem 0; max-width: 62ch;}
      .stDownloadButton button {width: 100%;}
      div[data-testid="stMetricValue"] {font-size: 1.45rem;}
      .rule-box {background:#f6f7f9; border:1px solid #e6e8ec; border-radius:10px;
                 padding:0.9rem 1.1rem; font-size:0.86rem; color:#404a58; line-height:1.55;}
      .region-tag {display:inline-block; background:#eef2ff; color:#3730a3; font-weight:600;
                   font-size:0.78rem; padding:0.15rem 0.6rem; border-radius:999px; margin-bottom:0.6rem;}
      @media (prefers-color-scheme: dark) {
        .lede {color:#9aa4b2;}
        .rule-box {background:#1b1f27; border-color:#2b313c; color:#c2cad6;}
        .region-tag {background:#1e2247; color:#a5b4fc;}
      }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Stock Upload Builder")
st.markdown(
    '<p class="lede">Drop in the raw warehouse stock file for each region. This applies that '
    "region's own sellable-stock rules, merges duplicate SKUs, holds back each SKU's threshold "
    "and fills every platform template in its own native format — ready to upload as-is.</p>",
    unsafe_allow_html=True,
)

# --------------------------------------------------------------------------
# sidebar — inputs
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("1 · Raw stock files")
    st.caption("Upload one or both — only the regions you provide a file for get built.")
    raw_files: dict[str, object] = {}
    for rk in E.REGION_ORDER:
        region = E.REGIONS[rk]
        f = st.file_uploader(
            f"{region.label} — WMS inventory snapshot",
            type=["csv", "xlsx", "xlsm"],
            key=f"raw_{rk}",
            help="The export with BLOCKED, LOCATION, LOCATION TYPE, SKU and TOTAL columns.",
        )
        if f is not None:
            raw_files[rk] = f

    st.header("2 · Platforms")
    chosen: list[str] = []
    for key in E.PLATFORM_ORDER:
        if st.checkbox(E.PLATFORMS[key].label, value=True, key=f"cb_{key}"):
            chosen.append(key)

    with st.expander("Advanced"):
        st.caption(
            "Each region's master item file is built into this tool. Upload a newer one "
            "here to override it for this run — same five sheets, same columns."
        )
        master_overrides: dict[str, object] = {}
        for rk in E.REGION_ORDER:
            region = E.REGIONS[rk]
            mf = st.file_uploader(
                f"{region.label} master item file (optional)",
                type=["xlsx"],
                key=f"master_{rk}",
            )
            if mf is not None:
                master_overrides[rk] = mf

    run_clicked = st.button(
        "Build upload files", type="primary", width='stretch',
        disabled=not (raw_files and chosen),
    )

# --------------------------------------------------------------------------
# rules panel
# --------------------------------------------------------------------------
with st.expander("What this tool does to your file", expanded=not raw_files):
    excl_counts = " · ".join(
        f"{E.REGIONS[rk].label}: {len(E.REGIONS[rk].excluded_locations)} bins"
        for rk in E.REGION_ORDER
    )
    st.markdown(
        f"""
        <div class="rule-box">
        <b>1 · Keep only sellable stock.</b> A row survives when <code>BLOCKED</code> is FALSE,
        <code>LOCATION TYPE</code> is BULK or PICK_FACE, and <code>LOCATION</code> is not one of
        that region's non-sellable bins (damaged, missing, expired, quarantine, 3PL and so on) —
        {excl_counts}. Locations are matched case-insensitively, because the WMS stores the same
        bin in more than one case.<br><br>
        <b>2 · One row per SKU.</b> Quantities are summed across bins. SKUs are matched
        case-insensitively here too — the WMS holds the same prescription-lens SKU in both cases,
        and keeping them apart silently loses units.<br><br>
        <b>3 · Sellable quantity.</b> <code>TOTAL</code> = Available − Reserved, floored at zero.
        Reserved units are already committed to open orders.<br><br>
        <b>4 · Threshold held back.</b> Each SKU's own threshold from that region's master file is
        deducted: final = MAX(sellable − threshold, 0).<br><br>
        <b>5 · Master file decides the SKU list.</b> Only SKUs listed on that platform's master
        sheet are written, using each region's own template file. A listed SKU with no stock is
        sent as 0 rather than left out, so every upload is a complete refresh.
        </div>
        """,
        unsafe_allow_html=True,
    )

# --------------------------------------------------------------------------
# run
# --------------------------------------------------------------------------
if run_clicked:
    out: dict[str, tuple] = {}
    try:
        for rk, f in raw_files.items():
            region = E.REGIONS[rk]
            with st.spinner(f"Reading {region.label} stock file…"):
                raw_df = E.read_raw_stock(f)
            with st.spinner(f"Applying {region.label} rules and filling templates…"):
                sellable, results = E.run(raw_df, chosen, rk, master_overrides.get(rk))
            out[rk] = (sellable, results, f.name)
        st.session_state["out"] = out
    except E.StockFileError as exc:
        st.error(str(exc))
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Something went wrong reading that file: {exc}")
        st.stop()

if "out" in st.session_state:
    out = st.session_state["out"]
    stamp = date.today().isoformat()

    st.success(
        f"Built upload files for {len(out)} region(s): "
        + ", ".join(E.REGIONS[rk].label for rk in out)
    )

    all_zbuf = io.BytesIO()
    with zipfile.ZipFile(all_zbuf, "w", zipfile.ZIP_DEFLATED) as all_zip:
        for rk in E.REGION_ORDER:
            if rk not in out:
                continue
            region = E.REGIONS[rk]
            sellable, results, src_name = out[rk]

            st.divider()
            st.markdown(f'<span class="region-tag">{region.label}</span>', unsafe_allow_html=True)
            st.subheader(f"{region.label} · Sellable stock")
            c = st.columns(4)
            c[0].metric("Rows in file", f"{sellable.rows_in:,}")
            c[1].metric("Sellable rows", f"{sellable.rows_after_location:,}")
            c[2].metric("Unique SKUs", f"{sellable.skus_out:,}")
            c[3].metric("Sellable units", f"{sellable.units_sellable:,}")

            detail = [
                f"Source: {src_name}",
                f"Blocked filter: {sellable.rows_in:,} → {sellable.rows_after_blocked:,}",
                f"Location type: → {sellable.rows_after_loctype:,}",
                f"Excluded bins: → {sellable.rows_after_location:,}",
                f"Quantity basis: {sellable.qty_basis}",
            ]
            if sellable.merged_case_variants:
                detail.append(
                    f"Merged {sellable.merged_case_variants:,} case-variant SKUs into their counterparts"
                )
            st.caption("  ·  ".join(detail))

            st.subheader(f"{region.label} · Per platform")
            table = pd.DataFrame(
                [
                    {
                        "Platform": plan.platform.label,
                        "SKUs in master": plan.skus,
                        "With stock": plan.matched,
                        "Units before threshold": plan.units_before,
                        "Units uploaded": plan.units_after,
                        "Held back": plan.units_before - plan.units_after,
                        "SKUs live": plan.live_skus,
                        "SKUs at zero": plan.skus - plan.live_skus,
                    }
                    for plan, _ in results.values()
                ]
            )
            st.dataframe(
                table, hide_index=True, width='stretch',
                column_config={
                    c: st.column_config.NumberColumn(format="localized")
                    for c in table.columns if c != "Platform"
                },
            )

            st.subheader(f"{region.label} · Download")
            cols = st.columns(min(len(results), 3))
            for i, (key, (plan, payload)) in enumerate(results.items()):
                p = plan.platform
                fname = f"{p.label.replace(' ', '_')}_{region.label}_stock_{stamp}{p.out_ext}"
                cols[i % len(cols)].download_button(
                    f"⬇  {p.label}", payload, file_name=fname,
                    key=f"dl_{rk}_{key}", width="stretch",
                )
                all_zip.writestr(fname, payload)

            if "sixth_street" in results:
                st.caption(
                    "⚠️ The 6th Street file has 13-digit barcode SKUs (e.g. 6290360089133). "
                    "The data in the CSV is correct plain digits — but if you double-click it "
                    "open in Excel, Excel auto-converts long digit strings to scientific "
                    "notation (6.29036E+12) for display only. This is cosmetic and doesn't "
                    "affect what gets uploaded. To check the file safely, use Excel's "
                    "**Data → From Text/CSV** and set the Sku column type to **Text** before "
                    "loading, rather than opening it directly."
                )

            all_zip.writestr(f"sellable_stock_{region.label}_{stamp}.csv", E.sellable_csv(sellable))

    st.divider()
    d = st.columns(1)
    d[0].download_button(
        "⬇  Everything, all regions (.zip)", all_zbuf.getvalue(),
        file_name=f"stock_uploads_{stamp}.zip", mime="application/zip", type="primary",
        width='stretch',
    )
else:
    st.info("Upload at least one region's raw stock file in the sidebar to begin.")
