"""Story 7 tests for COUNTRY_INFO row completeness checks."""

from pathlib import Path

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


def _build_workbook(
    path: Path,
    *,
    endemicity_values: dict[str, str],
    n_value: int = 2,
    ida_rows: set[int] | None = None,
    blank_cells: set[str] | None = None,
    output_header_columns: list[str] | None = None,
    extra_columns: dict[str, str] | None = None,
) -> None:
    workbook = Workbook()
    intro_sheet = workbook.active
    intro_sheet.title = "INTRO"
    for sheet_name in REQUIRED_SHEETS:
        if sheet_name == "INTRO":
            continue
        workbook.create_sheet(title=sheet_name)

    # Intro required baseline for Story 4/5/6.
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
    ida_rows = ida_rows or set()
    has_ida = any(10 <= row <= (9 + n_value) for row in ida_rows)

    # Align sheet states to Story 6 expected profile to avoid unrelated mismatch warnings.
    workbook["DEC"].sheet_state = "visible" if (has_lf and not has_oncho) else "hidden"
    workbook["IVM"].sheet_state = "visible" if has_oncho else "hidden"
    workbook["IVM+"].sheet_state = "visible" if has_ida else "hidden"
    workbook["ALB_MBD"].sheet_state = "visible" if (has_lf or has_sth) else "hidden"
    workbook["PZQ"].sheet_state = "visible" if has_sch else "hidden"
    workbook["DATA_POLICY"].sheet_state = "hidden"

    country_info = workbook["COUNTRY_INFO"]
    output_header_columns = output_header_columns or []
    for column in output_header_columns:
        country_info[f"{column}7"] = f"Output {column}"

    for row in ida_rows:
        country_info[f"Z{row}"] = "ida"

    active_rows = range(10, 10 + n_value)
    for row in active_rows:
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

        for column in output_header_columns:
            country_info[f"{column}{row}"] = f"value-{column}-{row}"

    for cell in (blank_cells or set()):
        if "!" in cell:
            sheet_name, cell_ref = cell.split("!")
            workbook[sheet_name][cell_ref] = "   "
        else:
            country_info[cell] = "   "

    for cell, value in (extra_columns or {}).items():
        country_info[cell] = value

    workbook.save(path)


def test_country_info_success_has_no_story7_findings(tmp_path: Path):
    workbook_path = tmp_path / "success.xlsx"
    _build_workbook(
        workbook_path,
        endemicity_values={"E37": "Endemic", "E39": "Endemic", "E41": "Endemic", "E43": "Endemic"},
        n_value=2,
        output_header_columns=["V", "W", "X"],
    )
    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )
    story7_rule_ids = {
        "JRSM.COUNTRY_INFO.REQUIRED_ID_COLUMNS_B_TO_D",
        "JRSM.COUNTRY_INFO.ENDEMICITY_COLUMNS_H_TO_K",
        "JRSM.COUNTRY_INFO.REQUIRED_OUTPUT_COLUMNS_V_TO_AD",
    }
    assert not [finding for finding in result.findings if finding.rule_id in story7_rule_ids]


def test_missing_bcd_emits_required_id_findings(tmp_path: Path):
    workbook_path = tmp_path / "missing_bcd.xlsx"
    _build_workbook(
        workbook_path,
        endemicity_values={"E37": "Non-endemic", "E39": "Non-endemic", "E41": "Non-endemic", "E43": "Non-endemic"},
        n_value=2,
        output_header_columns=["V"],
        blank_cells={"B10", "C10", "D11"},
    )
    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )
    id_findings = [
        finding
        for finding in result.findings
        if finding.rule_id == "JRSM.COUNTRY_INFO.REQUIRED_ID_COLUMNS_B_TO_D"
    ]
    cells = {finding.cell for finding in id_findings}
    assert {"B10", "C10", "D11"} <= cells


def test_endemicity_gated_columns_h_to_k(tmp_path: Path):
    workbook_path = tmp_path / "endemicity_h_to_k.xlsx"
    _build_workbook(
        workbook_path,
        endemicity_values={"E37": "Non-endemic", "E39": "Endemic", "E41": "Non-endemic", "E43": "Non-endemic"},
        n_value=2,
        output_header_columns=["V"],
        blank_cells={"I10"},
    )
    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )
    endemicity_findings = [
        finding
        for finding in result.findings
        if finding.rule_id == "JRSM.COUNTRY_INFO.ENDEMICITY_COLUMNS_H_TO_K"
    ]
    assert any(finding.cell == "I10" for finding in endemicity_findings)
    assert not any(finding.cell in {"H10", "J10", "K10"} for finding in endemicity_findings)


def test_required_output_columns_v_to_ad_checks_and_aggregate(tmp_path: Path):
    workbook_path = tmp_path / "output_columns.xlsx"
    _build_workbook(
        workbook_path,
        endemicity_values={"E37": "Non-endemic", "E39": "Non-endemic", "E41": "Non-endemic", "E43": "Non-endemic"},
        n_value=2,
        output_header_columns=["V", "W"],
        blank_cells={"V10", "V11", "W11"},
    )
    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )
    output_findings = [
        finding
        for finding in result.findings
        if finding.rule_id == "JRSM.COUNTRY_INFO.REQUIRED_OUTPUT_COLUMNS_V_TO_AD"
    ]
    assert any(finding.cell == "W11" for finding in output_findings)
    assert any(finding.cell == "V10:V11" for finding in output_findings)


def test_extra_columns_ae_af_are_tolerated(tmp_path: Path):
    workbook_path = tmp_path / "extra_columns.xlsx"
    _build_workbook(
        workbook_path,
        endemicity_values={"E37": "Non-endemic", "E39": "Non-endemic", "E41": "Non-endemic", "E43": "Non-endemic"},
        n_value=2,
        output_header_columns=["V"],
        extra_columns={"AE8": "IU_ID", "AF8": "IU_NAME"},
    )
    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )
    assert not any(
        (finding.cell or "").startswith("AE") or (finding.cell or "").startswith("AF")
        for finding in result.findings
    )
