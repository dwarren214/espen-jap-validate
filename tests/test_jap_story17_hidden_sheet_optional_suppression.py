"""Story 17 tests for hidden sheet and optional-field false-positive suppression."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from espen_sql_api.jap_validation.models import Finding, ValidationEngineResult
from espen_sql_api.jap_validation.response_builder import build_validation_response
from espen_sql_api.jap_validation.validators.jrsm import validate_jrsm_workbook

EXPECTED_VERSION_MARKER = "Joint request for selected PC medicines v.4.4"
REQUIRED_SHEETS = [
    "INTRO",
    "COUNTRY_INFO",
    "DEC",
    "ALB_MBD",
    "PZQ",
    "IVM",
    "IVM+",
    "SUMMARY",
    "SHIPMENT",
    "DATA_POLICY",
]
REQUIRED_INTRO_FIELDS = [
    "E37",
    "E38",
    "E39",
    "E40",
    "E41",
    "E42",
    "E43",
    "E44",
    "E45",
    "E48",
    "E49",
    "E50",
    "E51",
]


def _build_story17_workbook(
    path: Path,
    *,
    endemicity_values: dict[str, str] | None = None,
    output_header_columns: list[str] | None = None,
    hidden_columns: set[str] | None = None,
    sheet_state_overrides: dict[str, str] | None = None,
    cell_overrides: dict[str, object] | None = None,
    n_value: int = 2,
) -> None:
    endemicity_values = endemicity_values or {
        "E37": "Non-endemic",
        "E39": "Non-endemic",
        "E41": "Non-endemic",
        "E43": "Non-endemic",
    }
    output_header_columns = output_header_columns or []
    workbook = Workbook()
    intro_sheet = workbook.active
    intro_sheet.title = "INTRO"
    for sheet_name in REQUIRED_SHEETS:
        if sheet_name != "INTRO":
            workbook.create_sheet(title=sheet_name)

    intro_sheet["B2"] = EXPECTED_VERSION_MARKER
    intro_sheet["E33"] = "Rwanda"
    intro_sheet["E35"] = 2026
    for cell in REQUIRED_INTRO_FIELDS:
        intro_sheet[cell] = "value"
    intro_sheet["E37"] = endemicity_values["E37"]
    intro_sheet["E39"] = endemicity_values["E39"]
    intro_sheet["E41"] = endemicity_values["E41"]
    intro_sheet["E43"] = endemicity_values["E43"]
    intro_sheet["E45"] = n_value

    has_lf = endemicity_values["E37"].strip().casefold() != "non-endemic"
    has_oncho = endemicity_values["E39"].strip().casefold() != "non-endemic"
    has_sth = endemicity_values["E41"].strip().casefold() != "non-endemic"
    has_sch = endemicity_values["E43"].strip().casefold() != "non-endemic"
    workbook["DEC"].sheet_state = "visible" if (has_lf and not has_oncho) else "hidden"
    workbook["IVM"].sheet_state = "visible" if has_oncho else "hidden"
    workbook["IVM+"].sheet_state = "hidden"
    workbook["ALB_MBD"].sheet_state = "visible" if (has_lf or has_sth) else "hidden"
    workbook["PZQ"].sheet_state = "visible" if has_sch else "hidden"
    workbook["DATA_POLICY"].sheet_state = "hidden"
    for sheet_name, sheet_state in (sheet_state_overrides or {}).items():
        workbook[sheet_name].sheet_state = sheet_state

    country_info = workbook["COUNTRY_INFO"]
    for column in output_header_columns:
        country_info[f"{column}7"] = f"Output {column}"
    for column in hidden_columns or set():
        country_info.column_dimensions[column].hidden = True

    for row in range(10, 10 + n_value):
        country_info[f"B{row}"] = f"admin-{row}"
        country_info[f"C{row}"] = f"district-{row}"
        country_info[f"D{row}"] = f"iu-{row}"
        if has_lf:
            country_info[f"H{row}"] = "LF_CODE"
        if has_oncho:
            country_info[f"I{row}"] = "ONCHO_CODE"
        if has_sth:
            country_info[f"J{row}"] = "STH_CODE"
        if has_sch:
            country_info[f"K{row}"] = "SCH_CODE"

    for sheet_cell, value in (cell_overrides or {}).items():
        sheet_name, cell_address = sheet_cell.split("!")
        workbook[sheet_name][cell_address] = value

    workbook.save(path)


def test_hidden_pzq_module_emits_no_content_or_not_applicable_findings(tmp_path: Path):
    workbook_path = tmp_path / "hidden_pzq.xlsx"
    _build_story17_workbook(
        workbook_path,
        endemicity_values={
            "E37": "Endemic",
            "E39": "Endemic",
            "E41": "Endemic",
            "E43": "Non-endemic",
        },
        sheet_state_overrides={"PZQ": "hidden"},
        cell_overrides={"PZQ!G10": "manual override", "PZQ!I10": None},
    )

    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )

    assert result.context["expected_sheet_visibility"]["PZQ"] is False
    assert result.context["reviewable_sheet_visibility"]["PZQ"] is False
    assert not [finding for finding in result.findings if finding.sheet == "PZQ"]


def test_hidden_alb_mbd_and_ivm_modules_emit_no_content_findings(tmp_path: Path):
    workbook_path = tmp_path / "hidden_modules.xlsx"
    _build_story17_workbook(
        workbook_path,
        sheet_state_overrides={"ALB_MBD": "hidden", "IVM": "hidden"},
        cell_overrides={
            "ALB_MBD!A10": "manual override",
            "ALB_MBD!G10": None,
            "IVM!A10": "manual override",
            "IVM!M10": None,
        },
    )

    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )

    assert result.context["reviewable_sheet_visibility"]["ALB_MBD"] is False
    assert result.context["reviewable_sheet_visibility"]["IVM"] is False
    assert not [finding for finding in result.findings if finding.sheet in {"ALB_MBD", "IVM"}]


def test_formula_governance_still_reports_visible_reviewable_module_override(tmp_path: Path):
    workbook_path = tmp_path / "visible_pzq_override.xlsx"
    _build_story17_workbook(
        workbook_path,
        endemicity_values={
            "E37": "Non-endemic",
            "E39": "Non-endemic",
            "E41": "Non-endemic",
            "E43": "Endemic",
        },
        sheet_state_overrides={"PZQ": "visible"},
        cell_overrides={"PZQ!E10": "=IF(1=1,123,456)"},
    )

    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )

    assert result.context["reviewable_sheet_visibility"]["PZQ"] is True
    assert any(
        finding.rule_id == "JRSM.FORMULA.GOVERNED_FORMULA_MATCH"
        and finding.sheet == "PZQ"
        and finding.cell == "E10"
        for finding in result.findings
    )


def test_country_info_hidden_z_column_is_not_reported_as_required_blank(tmp_path: Path):
    workbook_path = tmp_path / "hidden_z.xlsx"
    _build_story17_workbook(
        workbook_path,
        output_header_columns=["Z"],
        hidden_columns={"Z"},
    )

    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )

    assert not any(
        finding.rule_id == "JRSM.COUNTRY_INFO.REQUIRED_OUTPUT_COLUMNS_V_TO_AD"
        and (finding.cell or "").startswith("Z")
        for finding in result.findings
    )
    assert "Z" not in result.context["country_info_required_output_columns"]


def test_country_info_survey_planning_columns_are_not_blanket_required(tmp_path: Path):
    workbook_path = tmp_path / "optional_survey_columns.xlsx"
    _build_story17_workbook(
        workbook_path,
        output_header_columns=["AA", "AB", "AC", "AD"],
    )

    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )

    assert not any(
        finding.rule_id == "JRSM.COUNTRY_INFO.REQUIRED_OUTPUT_COLUMNS_V_TO_AD"
        and (finding.cell or "")[:2] in {"AA", "AB", "AC", "AD"}
        for finding in result.findings
    )
    assert not ({"AA", "AB", "AC", "AD"} & set(result.context["country_info_required_output_columns"]))


def test_hidden_non_applicable_module_summary_is_not_evaluated():
    result = ValidationEngineResult(
        validation_outcome="pass",
        findings=[
            Finding(
                issue_id="F-0001",
                rule_id="JRSM.TEMPLATE.SHEET_SET",
                severity="info",
                message="Workbook sheet set matches expected JRSM template sheet names.",
                recommendation="No action required.",
            )
        ],
        context={
            "reviewable_sheet_visibility": {
                "INTRO": True,
                "COUNTRY_INFO": True,
                "DEC": False,
                "ALB_MBD": True,
                "PZQ": False,
                "IVM": True,
                "IVM+": False,
                "SUMMARY": True,
                "SHIPMENT": True,
                "DATA_POLICY": False,
            }
        },
    )

    response = build_validation_response(file_reference="upl_story17", validation_result=result)

    pzq_summary = next(summary for summary in response.sheet_summaries if summary.sheet == "PZQ")
    assert pzq_summary.status == "not_evaluated"
