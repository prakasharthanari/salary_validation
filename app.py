"""
Net Avenue - Monthly Payroll Validation Tool (Streamlit UI)

Automates the accounts team's monthly checks:
  1. CTC cross-check (current vs previous month)
  2. Salary calculation check (Basic/HRA/Conveyance/Special Allowance/Bonus/
     Gross/PF/ESI/PT recomputed from CTC + Salary Days, vs what was reported)
  3. Bank details check (Account Number & IFSC vs Bank Master)

All the actual logic lives in payroll_check.py so it can be tested without
launching Streamlit. This file is just the page.
"""
import base64
from pathlib import Path

import streamlit as st

import payroll_check as pc

BRAND_PINK = "#F41276"
BRAND_PINK_DARK = "#C90E60"
BRAND_INK = "#1A1A1A"
LOGO_PATH = Path(__file__).parent / "cbazaar_logo.jpg"


def check_password():
    try:
        secret = st.secrets.get("APP_PASSWORD", None)
    except Exception:
        secret = None
    if not secret:
        return True
    if st.session_state.get("authed"):
        return True
    pw = st.text_input("Password", type="password")
    if pw == secret:
        st.session_state["authed"] = True
        st.rerun()
    if pw:
        st.error("Incorrect password.")
    return False


def inject_css():
    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap');
        html, body, [class*="css"] {{ font-family: 'Poppins', sans-serif; }}
        .block-container {{ padding-top: 1.5rem; max-width: 900px; }}
        .app-header {{
            display: flex; align-items: center; gap: 1rem;
            padding-bottom: 0.75rem; border-bottom: 3px solid {BRAND_PINK};
            margin-bottom: 1.5rem;
        }}
        .app-header img {{ height: 46px; }}
        .app-header .title-block h1 {{
            font-size: 1.5rem; font-weight: 700; color: {BRAND_INK}; margin: 0; line-height: 1.2;
        }}
        .app-header .title-block p {{ font-size: 0.88rem; color: #6b6b6b; margin: 0; }}
        h2, h3 {{ color: {BRAND_INK}; font-weight: 600; }}
        div.stButton > button, div.stDownloadButton > button {{
            border-radius: 8px; font-weight: 600; border: none;
        }}
        div.stButton > button[kind="primary"] {{ background-color: {BRAND_PINK}; }}
        div.stButton > button[kind="primary"]:hover {{ background-color: {BRAND_PINK_DARK}; }}
        div.stDownloadButton > button {{
            background-color: white; color: {BRAND_PINK}; border: 1.5px solid {BRAND_PINK};
        }}
        div.stDownloadButton > button:hover {{ background-color: {BRAND_PINK}; color: white; }}
        footer {{visibility: hidden;}}
        #MainMenu {{visibility: hidden;}}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header():
    if LOGO_PATH.exists():
        logo_b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode()
        logo_html = f'<img src="data:image/jpeg;base64,{logo_b64}" alt="Cbazaar logo">'
    else:
        logo_html = ""
    st.markdown(
        f"""
        <div class="app-header">
            {logo_html}
            <div class="title-block">
                <h1>Monthly Payroll Validation</h1>
                <p>Net Avenue Technologies &middot; Accounts Tools</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


st.set_page_config(page_title="Net Avenue Payroll Validation", page_icon="\U0001F4CB", layout="centered")
inject_css()

if not check_password():
    render_header()
    st.stop()

render_header()

st.markdown(
    "Upload this month's Pay Register and CTC Master, last month's CTC Master "
    "(for the CTC cross-check), and the Bank Master export -- get back a single "
    "exceptions report instead of checking everything by hand."
)

with st.container(border=True):
    st.markdown("##### 1. This month's files")
    c1, c2 = st.columns(2)
    payreg_file = c1.file_uploader("Pay Register (.xlsx)", type=["xlsx"], key="payreg")
    payreg_pw = c1.text_input("Password (if protected)", type="password", key="payreg_pw")
    ctc_file = c2.file_uploader("CTC Master (.xlsx)", type=["xlsx"], key="ctc")
    ctc_pw = c2.text_input("Password (if protected)", type="password", key="ctc_pw")

    st.markdown("##### 2. Reference files")
    c3, c4 = st.columns(2)
    prev_ctc_file = c3.file_uploader("Last month's CTC Master (.xlsx)", type=["xlsx"], key="prev_ctc")
    prev_ctc_pw = c3.text_input("Password (if protected)", type="password", key="prev_ctc_pw")
    bank_master_file = c4.file_uploader("Bank Master export (.xlsx)", type=["xlsx"], key="bank_master")

    run = st.button("Run Validation", type="primary")

if run:
    if not (payreg_file and ctc_file and prev_ctc_file):
        st.error("Pay Register, CTC Master, and last month's CTC Master are all required.")
        st.stop()

    try:
        with st.spinner("Reading files..."):
            payreg_bytes = pc.decrypt_if_needed(payreg_file.getvalue(), payreg_pw or None)
            ctc_bytes = pc.decrypt_if_needed(ctc_file.getvalue(), ctc_pw or None)
            prev_ctc_bytes = pc.decrypt_if_needed(prev_ctc_file.getvalue(), prev_ctc_pw or None)

            salary_df, bank_df = pc.load_payregister(payreg_bytes)
            salary_df = salary_df[salary_df[pc.find_col(salary_df.columns, "emp", "id")].notna()]
            ctc_df = pc.load_ctc_master(ctc_bytes)
            prev_ctc_df = pc.load_ctc_master(prev_ctc_bytes)
            bank_master_df = pc.load_bank_master(bank_master_file.getvalue()) if bank_master_file else None

        with st.spinner("Running checks..."):
            results = pc.run_checks(salary_df, bank_df, ctc_df, prev_ctc_df, bank_master_df)

        st.session_state["results"] = results
    except ValueError as e:
        st.error(str(e))
        st.stop()
    except Exception as e:
        st.error(f"Something went wrong while processing the files: {e}")
        st.stop()

if st.session_state.get("results"):
    results = st.session_state["results"]

    with st.container(border=True):
        st.markdown("##### Results")
        st.caption(f"Days in month used for proration (auto-detected): **{int(results.days_in_month)}**")
        m1, m2, m3 = st.columns(3)
        m1.metric("CTC changes", len(results.ctc_changes))
        m2.metric("Salary exceptions", len(results.salary_exceptions))
        m3.metric("Bank mismatches", len(results.bank_mismatches))

        for w in results.warnings:
            st.warning(w)

        if len(results.not_in_ctc):
            st.warning(
                f"{len(results.not_in_ctc)} employee(s) in the Pay Register were not found "
                "in the CTC Master -- see the report for details."
            )

        report_bytes = pc.write_report(results)
        st.download_button(
            "\u2b07 Download Exceptions Report",
            data=report_bytes,
            file_name="Payroll_Validation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if len(results.ctc_changes):
        with st.expander(f"CTC Changes ({len(results.ctc_changes)}) -- needs HR approval"):
            st.dataframe(results.ctc_changes, use_container_width=True)

    if len(results.salary_exceptions):
        with st.expander(f"Salary Exceptions ({len(results.salary_exceptions)})"):
            st.dataframe(results.salary_exceptions, use_container_width=True)

    if len(results.bank_mismatches):
        with st.expander(f"Bank Mismatches ({len(results.bank_mismatches)})"):
            st.dataframe(results.bank_mismatches, use_container_width=True)
