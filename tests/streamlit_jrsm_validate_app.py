"""Simple Streamlit UI for local JRSM validation runs.

Run with:
    streamlit run tests/streamlit_jrsm_validate_app.py
"""

from __future__ import annotations

import json
from pathlib import Path
import sys
from tempfile import NamedTemporaryFile

import streamlit as st

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from espen_sql_api.jap_validation.response_builder import build_validation_response
from espen_sql_api.jap_validation.validators.jrsm import validate_jrsm_workbook


def _save_uploaded_workbook(uploaded_file) -> Path:
    suffix = Path(uploaded_file.name).suffix or ".xlsx"
    with NamedTemporaryFile(delete=False, suffix=suffix) as handle:
        handle.write(uploaded_file.getbuffer())
        return Path(handle.name)


def _download_filename(uploaded_name: str) -> str:
    stem = Path(uploaded_name).stem or "jrsm-validation"
    return f"{stem}.validation.json"


def main() -> None:
    st.set_page_config(page_title="JRSM Validate", layout="wide")
    st.title("JRSM Validate")
    st.caption("Local JRSM workbook validator for manual review.")

    with st.sidebar:
        st.header("Inputs")
        country = st.text_input("Country", value="Rwanda")
        year = st.number_input("Year for request of medicine", min_value=1900, max_value=2100, value=2026, step=1)
        findings_mode = st.selectbox("Findings mode", options=["exceptions_only", "full"], index=0)
        findings_cap_enabled = st.checkbox("Apply findings cap", value=False)
        findings_cap = st.number_input("Findings cap", min_value=1, max_value=5000, value=100, step=1, disabled=not findings_cap_enabled)

    workbook = st.file_uploader("Drag in a JRSM workbook", type=["xlsx", "xlsm"])
    validate_clicked = st.button("Validate", type="primary", disabled=workbook is None)

    if not validate_clicked:
        st.info("Upload a workbook and click Validate.")
        return

    if workbook is None:
        st.error("Please upload a workbook.")
        return

    temp_path = _save_uploaded_workbook(workbook)
    try:
        with st.spinner("Running validation..."):
            validation_result = validate_jrsm_workbook(
                workbook_path=temp_path,
                country=country,
                year_for_request_of_medicine=int(year),
                metadata={"source": "streamlit_local"},
            )
            response = build_validation_response(
                file_reference=f"local::{workbook.name}",
                validation_result=validation_result,
                findings_mode=findings_mode,
                findings_cap=int(findings_cap) if findings_cap_enabled else None,
            )
    finally:
        temp_path.unlink(missing_ok=True)

    st.subheader("Summary")
    st.write(response.summary_text)

    summary = response.executive_summary
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Errors", summary.error_count)
    col2.metric("Warnings", summary.warn_count)
    col3.metric("Info", summary.info_count)
    col4.metric("Total Findings", summary.total_findings)

    if response.findings_truncated:
        st.warning(
            f"Returned {response.findings_returned_count} of {response.findings_total_count} findings "
            f"because the findings cap is active."
        )

    if summary.major_blockers:
        st.subheader("Major Blockers")
        for blocker in summary.major_blockers:
            st.write(f"- {blocker}")

    st.subheader("Response Blocks")
    for block in response.response_text_blocks:
        with st.expander(block.title, expanded=True):
            st.write(block.body)

    st.subheader("Findings")
    finding_rows = [
        {
            "issue_id": finding.issue_id,
            "severity": finding.severity,
            "rule_id": finding.rule_id,
            "sheet": finding.sheet or "",
            "cell": finding.cell or "",
            "message": finding.message,
            "recommendation": finding.recommendation,
        }
        for finding in response.findings
    ]
    st.dataframe(finding_rows, use_container_width=True, hide_index=True)

    payload = response.model_dump(mode="json", exclude_none=True)
    st.subheader("Raw JSON")
    st.json(payload)
    st.download_button(
        label="Download validation JSON",
        data=json.dumps(payload, indent=2),
        file_name=_download_filename(workbook.name),
        mime="application/json",
    )


if __name__ == "__main__":
    main()
