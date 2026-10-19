"""FX Deal ID Cleanup - Streamlit app."""
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from cleanup_logic import DEFAULT_DEAL_IN_ID, CleanupError, process_workbook

st.set_page_config(page_title="FX Deal ID Cleanup", page_icon="💱", layout="centered")

st.title("💱 FX Deal ID Cleanup")
st.write(
    "Upload the workbook and get back the same file with the "
    "**Aspect Internal Trades Uploader** (Sheet3) filled in from the "
    "**Risk Month x Deal Net EUR Exposure** grid (Sheet2)."
)

with st.expander("What does this do?"):
    st.markdown(
        """
- Reads every Deal ID × Month cell on **Sheet2** and keeps the nonzero numeric ones.
- Writes one row per combination to **Sheet3**:
  - **Deal In ID** → the fixed internal deal (default 627080, editable below)
  - **Deal Out ID** → the Deal ID from Sheet2
  - **Trade Date** → the date chosen below (defaults to today)
  - **Conv Date** → the 10th of that month, or the previous Friday if the 10th is a weekend
  - **Quantity** → the value from Sheet2
- Currency, Exchange Rate, TP IN/OUT, Hedge Request ID and Unique ID columns are never touched.
- Your original file is not changed; you download a new `<name>_output` file.
        """
    )

today_ist = datetime.now(ZoneInfo("Asia/Kolkata")).date()

with st.sidebar:
    st.header("Settings")
    deal_in_id = st.number_input(
        "Deal In ID (main internal deal)",
        value=DEFAULT_DEAL_IN_ID,
        step=1,
        format="%d",
        help="All residual exposure flows into this deal.",
    )
    trade_date = st.date_input("Trade Date", value=today_ist, format="DD/MM/YYYY")

uploaded = st.file_uploader(
    "Drag and drop the Excel file here, or click to browse",
    type=["xlsx", "xlsm"],
)

if uploaded is None:
    st.info("Waiting for a file…")
    st.stop()

try:
    with st.spinner("Processing…"):
        result = process_workbook(
            uploaded.getvalue(),
            uploaded.name,
            trade_date=trade_date,
            deal_in_id=int(deal_in_id),
        )
except CleanupError as exc:
    st.error(str(exc))
    st.stop()
except Exception as exc:  # anything unexpected
    st.error(f"Something unexpected went wrong: {exc}")
    st.stop()

if result.skipped_text_cells:
    n = len(result.skipped_text_cells)
    st.warning(
        f"{n} cell(s) on Sheet2 contained text instead of a number (e.g. a literal "
        "'####') and were NOT included. Fix them at the source (re-copy as real values) "
        "and upload again."
    )
    st.dataframe(
        pd.DataFrame(result.skipped_text_cells, columns=["Deal ID", "Month", "Cell text"]),
        hide_index=True,
        use_container_width=True,
    )

if not result.rows:
    st.error("No qualifying rows were found – nothing was written.")
    st.stop()

deal_count = len({r[0] for r in result.rows})
st.success(
    f"Found {len(result.rows)} Deal ID / Month combinations across {deal_count} deals. "
    f"Written to Sheet3 rows {result.first_row}–{result.last_row}."
)

st.download_button(
    label=f"⬇️ Download {result.output_filename}",
    data=result.output_bytes,
    file_name=result.output_filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.subheader("Preview")
preview = pd.DataFrame(
    [
        {
            "Deal In ID": result.deal_in_id,
            "Deal Out ID": d,
            "Trade Date": t.strftime("%d-%b-%Y"),
            "Conv Date": c.strftime("%d-%b-%Y"),
            "Quantity": q,
        }
        for d, t, c, q in result.rows
    ]
)
st.dataframe(
    preview.style.format({"Quantity": "{:,.2f}"}),
    hide_index=True,
    use_container_width=True,
)
