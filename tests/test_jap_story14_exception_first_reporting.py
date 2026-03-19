"""Story 14 tests for exception-first reporting and focused JRSM output."""

from pathlib import Path

from espen_sql_api.jap_validation.models import Finding, ValidationEngineResult
from espen_sql_api.jap_validation.response_builder import build_validation_response
from espen_sql_api.jap_validation.validators.jrsm import validate_jrsm_workbook


def test_build_validation_response_defaults_to_exception_only_mode():
    result = ValidationEngineResult(
        validation_outcome="fail",
        findings=[
            Finding(
                issue_id="F-0001",
                rule_id="JRSM.TEMPLATE.SHEET_SET",
                severity="info",
                message="Workbook sheet set matches expected JRSM template sheet names.",
                recommendation="No action required.",
            ),
            Finding(
                issue_id="F-0002",
                rule_id="JRSM.COUNTRY_INFO.REQUIRED_OUTPUT_COLUMNS_V_TO_AD",
                severity="error",
                message="COUNTRY_INFO column AB (Oncho epidemiological surveys planned for the year) is blank across active IU rows 10:39.",
                recommendation="Populate the column.",
                sheet="COUNTRY_INFO",
                cell="AB10:AB39",
                actual="SYSTEMATICALLY_BLANK",
            ),
        ],
    )

    response = build_validation_response(file_reference="upl_test", validation_result=result)

    assert len(response.findings) == 1
    assert response.findings[0].rule_id == "JRSM.COUNTRY_INFO.REQUIRED_OUTPUT_COLUMNS_V_TO_AD"
    assert response.executive_summary.info_count == 0
    assert response.sheet_summaries
    assert next(summary for summary in response.sheet_summaries if summary.sheet == "COUNTRY_INFO").status == "issues_found"
    assert next(summary for summary in response.sheet_summaries if summary.sheet == "INTRO").status == "passed"


def test_build_validation_response_full_mode_keeps_info_findings():
    result = ValidationEngineResult(
        validation_outcome="fail",
        findings=[
            Finding(
                issue_id="F-0001",
                rule_id="JRSM.TEMPLATE.SHEET_SET",
                severity="info",
                message="Workbook sheet set matches expected JRSM template sheet names.",
                recommendation="No action required.",
            ),
            Finding(
                issue_id="F-0002",
                rule_id="JRSM.COUNTRY_INFO.REQUIRED_OUTPUT_COLUMNS_V_TO_AD",
                severity="error",
                message="COUNTRY_INFO column AB (Oncho epidemiological surveys planned for the year) is blank across active IU rows 10:39.",
                recommendation="Populate the column.",
                sheet="COUNTRY_INFO",
                cell="AB10:AB39",
                actual="SYSTEMATICALLY_BLANK",
            ),
        ],
    )

    response = build_validation_response(
        file_reference="upl_test",
        validation_result=result,
        findings_mode="full",
        findings_cap=None,
    )

    assert len(response.findings) == 2
    assert response.executive_summary.info_count == 1
    assert next(summary for summary in response.sheet_summaries if summary.sheet == "COUNTRY_INFO").error_count == 1


def test_build_validation_response_parse_failure_marks_sheets_not_evaluated():
    result = ValidationEngineResult(
        validation_outcome="fail",
        findings=[
            Finding(
                issue_id="F-0001",
                rule_id="JRSM.WORKBOOK.FILE_PARSE",
                severity="error",
                message="Workbook parse failed.",
                recommendation="Upload a valid workbook.",
                actual="WORKBOOK_PARSE_FAILED: bad zip",
            )
        ],
    )

    response = build_validation_response(file_reference="upl_test", validation_result=result)

    assert response.sheet_summaries
    assert all(summary.status == "not_evaluated" for summary in response.sheet_summaries)
    assert all(summary.error_count == 0 and summary.warn_count == 0 for summary in response.sheet_summaries)


def test_rwanda_workbook_exception_first_output_is_contextualized_without_default_truncation():
    workbook_path = Path("docs/Rwanda_JRSM_2026_PC_ESPEN_22112025.V2 xlsm.xlsm")
    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )

    response = build_validation_response(file_reference="upl_rwanda", validation_result=result)

    assert response.findings_truncated is None
    assert response.findings_total_count is None
    assert response.findings_returned_count is None
    assert len(response.findings) == response.executive_summary.total_findings
    assert len(response.findings) > 100
    assert all(finding.severity != "info" for finding in response.findings)
    assert not any(
        finding.rule_id == "JRSM.ENDEMICITY.ACTUAL_SHEET_STATE_MISMATCH" and finding.sheet == "DATA_POLICY"
        for finding in response.findings
    )
    assert any(
        "Oncho" in finding.message and "Epidemiological surveys planned for the year" in finding.message
        for finding in response.findings
        if finding.rule_id == "JRSM.COUNTRY_INFO.REQUIRED_OUTPUT_COLUMNS_V_TO_AD"
    )
    assert any(
        "Target population for STH / PreSAC / Rounds" in finding.message
        for finding in response.findings
        if finding.rule_id == "JRSM.ALB_MBD.REQUIRED_COLUMNS_G_J_Q_T"
    )


def test_rwanda_workbook_explicit_findings_cap_truncates_output():
    workbook_path = Path("docs/Rwanda_JRSM_2026_PC_ESPEN_22112025.V2 xlsm.xlsm")
    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )

    response = build_validation_response(
        file_reference="upl_rwanda",
        validation_result=result,
        findings_cap=100,
    )

    assert response.findings_truncated is True
    assert response.findings_total_count is not None
    assert response.findings_returned_count == 100
    assert response.findings_total_count > response.findings_returned_count
