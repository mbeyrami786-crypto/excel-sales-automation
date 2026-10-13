"""Streamlit front-end for the Excel sales automation project.

Run with:  streamlit run app.py

It reuses ``core.py``, so the downloaded workbook is byte-for-byte the same
report the CLI (``python main3.py``) produces.
"""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from core import DataValidationError, clean_sales_data, money, summarize, write_report

MAX_UPLOAD_MB = 10

st.set_page_config(page_title="AI Sales Automation", layout="wide")

st.title("AI Sales Automation Dashboard")
st.write(
    "Upload a sales Excel file (`Product`, `Region`, `Quantity`, `Price`) "
    "and get a fully styled dashboard report."
)


def read_raw(data: bytes) -> pd.DataFrame:
    """Read the uploaded workbook as-is (preview + cleaning input)."""
    return pd.read_excel(io.BytesIO(data))


uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"], accept_multiple_files=False)

if uploaded_file is None:
    st.info(
        "No file selected. Expected columns: "
        "`Product`, `Region`, `Quantity`, `Price`. "
        f"Files above {MAX_UPLOAD_MB} MB are rejected."
    )
    st.stop()

if uploaded_file.size > MAX_UPLOAD_MB * 1024 * 1024:
    st.error(f"File is larger than {MAX_UPLOAD_MB} MB.")
    st.stop()

raw_bytes = uploaded_file.getvalue()

try:
    raw_df = read_raw(raw_bytes)
    # same validation/cleaning rules as the CLI
    df = clean_sales_data(raw_df)
except DataValidationError as exc:
    st.error(str(exc))
    st.stop()
except Exception as exc:  # noqa: BLE001 - never show a raw traceback to the user
    st.error(f"Could not process the file: {exc}")
    st.stop()

summary = summarize(df)

st.success(f"File loaded — {summary.transactions:,} valid rows "
           f"(out of {len(raw_df):,} uploaded).")

st.subheader("Preview Data")
st.dataframe(raw_df.head(20), use_container_width=True)

st.divider()
st.subheader("Business Dashboard")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Revenue", money(summary.total_revenue))
col2.metric("Total Quantity", f"{summary.total_quantity:,.0f}")
col3.metric("Avg Unit Price", money(summary.average_unit_price))
col4.metric("Transactions", f"{summary.transactions:,}")

col5, col6, col7, col8 = st.columns(4)
col5.metric("Best Product", summary.best_product, help=f"Revenue: {money(summary.best_product_revenue)}")
col6.metric("Best Region", summary.best_region, help=f"Revenue: {money(summary.best_region_revenue)}")
col7.metric("Weighted Price", money(summary.weighted_price))
col8.metric("Rows Discarded", f"{len(raw_df) - summary.transactions:,}")

st.divider()

left, right = st.columns(2)
with left:
    st.subheader("Sales By Product")
    st.bar_chart(summary.product_sales)
with right:
    st.subheader("Sales By Region")
    st.bar_chart(summary.region_sales)

st.subheader("Top 5")
top_l, top_r = st.columns(2)
with top_l:
    st.dataframe(
        summary.top_products().rename("Revenue").rename_axis("Product").reset_index(),
        use_container_width=True,
        hide_index=True,
    )
with top_r:
    st.dataframe(
        summary.top_regions().rename("Revenue").rename_axis("Region").reset_index(),
        use_container_width=True,
        hide_index=True,
    )

# ---- report download ------------------------------------------------------
# Built in memory: no shared file on disk, safe for several users at once.
report_bytes = write_report(df, summary, None)

st.divider()
st.subheader("Download")
st.download_button(
    label="Download full Excel report (with dashboard & charts)",
    data=report_bytes,
    file_name="sales_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
