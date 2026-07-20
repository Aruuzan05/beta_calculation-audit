"""
Beta Calculation Audit — web app (Streamlit).

Deployment layer only. All the actual beta / fundamentals logic lives in
weighted_beta_v2.py and is reused here via build_audit_workbook().

The portfolio manager opens the URL, enters the shared password, uploads a
Holdings.xlsx, and downloads the generated Beta_Calculation_Audit.xlsx.

Run locally:   streamlit run app.py
"""

import io
import contextlib
from datetime import datetime

import pandas as pd
import streamlit as st

from weighted_beta_v2 import build_audit_workbook

st.set_page_config(page_title="Beta Calculation Audit", page_icon="📈")


# ---------------------------------------------------------------------------
# 1. Password gate  (single shared password, stored in Streamlit secrets)
# ---------------------------------------------------------------------------
def check_password() -> bool:
    expected = st.secrets.get("APP_PASSWORD")
    if not expected:
        st.error("APP_PASSWORD is not set. Add it to .streamlit/secrets.toml "
                 "(local) or the app's Secrets box (cloud).")
        return False

    if st.session_state.get("authenticated"):
        return True

    pw = st.text_input("Password", type="password")
    if pw == "":
        st.stop()
    if pw == expected:
        st.session_state["authenticated"] = True
        return True
    st.error("Incorrect password.")
    st.stop()


# ---------------------------------------------------------------------------
# 2. Main UI
# ---------------------------------------------------------------------------
st.title("📈 Beta Calculation Audit")
st.caption("Upload a holdings file → generate the 45-day beta audit workbook. "
           "Benchmark: 2800.HK (Hang Seng ETF).")

if not check_password():
    st.stop()

uploaded = st.file_uploader(
    "Upload holdings (.xlsx)",
    type=["xlsx"],
    help="Must contain the columns: Ticker, Local Currency, Weighting.",
)

REQUIRED_COLS = {"Ticker", "Local Currency", "Weighting"}

if uploaded is not None:
    try:
        holdings = pd.read_excel(io.BytesIO(uploaded.getvalue()))
    except Exception as e:
        st.error(f"Could not read that file as Excel: {e}")
        st.stop()

    missing = REQUIRED_COLS - set(holdings.columns)
    if missing:
        st.error(f"Holdings file is missing required column(s): {', '.join(sorted(missing))}")
        st.stop()

    st.success(f"Loaded {len(holdings)} holdings.")
    with st.expander("Preview holdings"):
        st.dataframe(holdings, use_container_width=True)

    if st.button("Generate audit workbook", type="primary"):
        log_buffer = io.StringIO()
        try:
            with st.spinner("Fetching prices & fundamentals and building the workbook… "
                            "this can take a minute for large portfolios."):
                # Capture the engine's print() output so we can show it as a log.
                with contextlib.redirect_stdout(log_buffer):
                    wb = build_audit_workbook(holdings)
                out = io.BytesIO()
                wb.save(out)
        except Exception as e:
            st.error(f"Generation failed: {e}")
            with st.expander("Run log"):
                st.code(log_buffer.getvalue() or "(no output)")
            st.stop()

        stamp = datetime.today().strftime("%Y%m%d_%H%M")
        st.success("Done. Download your audit workbook below.")
        st.download_button(
            "⬇️ Download Beta_Calculation_Audit.xlsx",
            data=out.getvalue(),
            file_name=f"Beta_Calculation_Audit_{stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        with st.expander("Run log"):
            st.code(log_buffer.getvalue() or "(no output)")
else:
    st.info("Waiting for a holdings file…")
