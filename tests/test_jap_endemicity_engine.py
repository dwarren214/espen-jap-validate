"""Story 6 tests for endemicity visibility and IU-window derivation."""

from pathlib import Path

import pytest
from openpyxl import Workbook

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


def _build_workbook(
    path: Path,
    *,
    endemicity_values: dict[str, str],
    n_value: object = 3,
    ida_rows: set[int] | None = None,
    align_sheet_state_with_expected: bool = True,
    sheet_state_overrides: dict[str, str] | None = None,
) -> None:
    workbook = Workbook()
    intro_sheet = workbook.active
    intro_sheet.title = "INTRO"
    for sheet_name in REQUIRED_SHEETS:
        if sheet_name == "INTRO":
            continue
        workbook.create_sheet(title=sheet_name)

    # Story 4/5 baseline-required Intro cells.
    intro_sheet["B2"] = EXPECTED_VERSION_MARKER
    intro_sheet["E33"] = "Rwanda"
    intro_sheet["E35"] = 2026
    intro_sheet["E37"] = endemicity_values["E37"]
    intro_sheet["E38"] = "value_E38"
    intro_sheet["E39"] = endemicity_values["E39"]
    intro_sheet["E40"] = "value_E40"
    intro_sheet["E41"] = endemicity_values["E41"]
    intro_sheet["E42"] = "value_E42"
    intro_sheet["E43"] = endemicity_values["E43"]
    intro_sheet["E44"] = "value_E44"
    intro_sheet["E45"] = n_value
    intro_sheet["E48"] = "value_E48"
    intro_sheet["E49"] = "value_E49"
    intro_sheet["E50"] = "value_E50"
    intro_sheet["E51"] = "value_E51"

    country_info_sheet = workbook["COUNTRY_INFO"]
    ida_rows = ida_rows or set()
    for row in ida_rows:
        country_info_sheet[f"Z{row}"] = "ida"

    if align_sheet_state_with_expected:
        has_lf = str(endemicity_values["E37"]).strip().casefold() != "non-endemic"
        has_oncho = str(endemicity_values["E39"]).strip().casefold() != "non-endemic"
        has_sth = str(endemicity_values["E41"]).strip().casefold() != "non-endemic"
        has_sch = str(endemicity_values["E43"]).strip().casefold() != "non-endemic"
        n_int = int(n_value) if isinstance(n_value, int) else None
        has_ida = False
        if n_int is not None:
            active_rows = set(range(10, 10 + n_int))
            has_ida = any(row in active_rows for row in ida_rows)

        expected_visibility = {
            "INTRO": "visible",
            "COUNTRY_INFO": "visible",
            "SUMMARY": "visible",
            "SHIPMENT": "visible",
            "DEC": "visible" if (has_lf and not has_oncho) else "hidden",
            "IVM": "visible" if has_oncho else "hidden",
            "IVM+": "visible" if has_ida else "hidden",
            "ALB_MBD": "visible" if (has_lf or has_sth) else "hidden",
            "PZQ": "visible" if has_sch else "hidden",
            "DATA_POLICY": "hidden",
        }
        for sheet_name, state in expected_visibility.items():
            workbook[sheet_name].sheet_state = state

    for sheet_name, state in (sheet_state_overrides or {}).items():
        workbook[sheet_name].sheet_state = state

    workbook.save(path)


@pytest.mark.parametrize(
    ("profile_name", "endemicity_values", "ida_rows", "expected_visibility"),
    [
        (
            "lf_only",
            {"E37": "Endemic", "E39": "Non-endemic", "E41": "Non-endemic", "E43": "Non-endemic"},
            set(),
            {"DEC": True, "IVM": False, "IVM+": False, "ALB_MBD": True, "PZQ": False},
        ),
        (
            "oncho_only",
            {"E37": "Non-endemic", "E39": "Endemic", "E41": "Non-endemic", "E43": "Non-endemic"},
            set(),
            {"DEC": False, "IVM": True, "IVM+": False, "ALB_MBD": False, "PZQ": False},
        ),
        (
            "sth_only",
            {"E37": "Non-endemic", "E39": "Non-endemic", "E41": "Endemic", "E43": "Non-endemic"},
            set(),
            {"DEC": False, "IVM": False, "IVM+": False, "ALB_MBD": True, "PZQ": False},
        ),
        (
            "sch_only",
            {"E37": "Non-endemic", "E39": "Non-endemic", "E41": "Non-endemic", "E43": "Endemic"},
            set(),
            {"DEC": False, "IVM": False, "IVM+": False, "ALB_MBD": False, "PZQ": True},
        ),
        (
            "mixed",
            {"E37": "Endemic", "E39": "Endemic", "E41": "Endemic", "E43": "Endemic"},
            {10},
            {"DEC": False, "IVM": True, "IVM+": True, "ALB_MBD": True, "PZQ": True},
        ),
    ],
)
def test_endemicity_profiles_expected_visibility(
    tmp_path: Path,
    profile_name: str,
    endemicity_values: dict[str, str],
    ida_rows: set[int],
    expected_visibility: dict[str, bool],
):
    workbook_path = tmp_path / f"{profile_name}.xlsx"
    _build_workbook(
        workbook_path,
        endemicity_values=endemicity_values,
        n_value=3,
        ida_rows=ida_rows,
        align_sheet_state_with_expected=True,
    )

    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )
    for sheet_name, expected in expected_visibility.items():
        assert result.context["expected_sheet_visibility"][sheet_name] is expected

    mismatch_findings = [
        finding
        for finding in result.findings
        if finding.rule_id == "JRSM.ENDEMICITY.ACTUAL_SHEET_STATE_MISMATCH"
    ]
    assert not mismatch_findings


def test_ida_detection_uses_only_active_iu_rows(tmp_path: Path):
    workbook_path = tmp_path / "ida_outside_window.xlsx"
    _build_workbook(
        workbook_path,
        endemicity_values={"E37": "Endemic", "E39": "Endemic", "E41": "Endemic", "E43": "Endemic"},
        n_value=2,
        ida_rows={15},  # outside active rows 10..11
        align_sheet_state_with_expected=False,
    )
    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )
    assert result.context["has_ida"] is False
    ida_findings = [finding for finding in result.findings if finding.rule_id == "JRSM.ENDEMICITY.IDA_DETECTION"]
    assert ida_findings
    assert ida_findings[-1].actual == "IDA_NOT_DETECTED"


def test_visibility_mismatch_and_data_policy_visible_warn(tmp_path: Path):
    workbook_path = tmp_path / "visibility_mismatch.xlsx"
    _build_workbook(
        workbook_path,
        endemicity_values={"E37": "Endemic", "E39": "Non-endemic", "E41": "Non-endemic", "E43": "Non-endemic"},
        n_value=3,
        ida_rows=set(),
        align_sheet_state_with_expected=True,
        sheet_state_overrides={
            "DEC": "hidden",      # expected visible
            "IVM": "visible",     # expected hidden
            "DATA_POLICY": "visible",  # expected hidden
        },
    )
    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )
    mismatch_sheets = {
        finding.sheet
        for finding in result.findings
        if finding.rule_id == "JRSM.ENDEMICITY.ACTUAL_SHEET_STATE_MISMATCH"
    }
    assert {"DEC", "IVM", "DATA_POLICY"} <= mismatch_sheets


def test_n_validation_and_iu_window_derivation(tmp_path: Path):
    valid_workbook_path = tmp_path / "n_valid.xlsx"
    _build_workbook(
        valid_workbook_path,
        endemicity_values={"E37": "Endemic", "E39": "Non-endemic", "E41": "Non-endemic", "E43": "Non-endemic"},
        n_value=4,
        ida_rows=set(),
        align_sheet_state_with_expected=True,
    )
    valid_result = validate_jrsm_workbook(
        workbook_path=valid_workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )
    assert valid_result.context["n"] == 4
    assert valid_result.context["iu_window_start_row"] == 10
    assert valid_result.context["iu_window_end_row"] == 13
    assert any(
        finding.rule_id == "JRSM.IU_WINDOW.N_VALID" and finding.severity == "info"
        for finding in valid_result.findings
    )
    assert any(
        finding.rule_id == "JRSM.IU_WINDOW.RANGE_ENFORCED" and finding.severity == "info"
        for finding in valid_result.findings
    )

    invalid_workbook_path = tmp_path / "n_invalid.xlsx"
    _build_workbook(
        invalid_workbook_path,
        endemicity_values={"E37": "Endemic", "E39": "Non-endemic", "E41": "Non-endemic", "E43": "Non-endemic"},
        n_value="not_a_number",
        ida_rows=set(),
        align_sheet_state_with_expected=False,
    )
    invalid_result = validate_jrsm_workbook(
        workbook_path=invalid_workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )
    assert invalid_result.context["n"] is None
    assert invalid_result.context["iu_window_end_row"] is None
    assert any(
        finding.rule_id == "JRSM.IU_WINDOW.N_VALID" and finding.severity == "error"
        for finding in invalid_result.findings
    )
    assert any(
        finding.rule_id == "JRSM.IU_WINDOW.RANGE_ENFORCED" and finding.severity == "warn"
        for finding in invalid_result.findings
    )
