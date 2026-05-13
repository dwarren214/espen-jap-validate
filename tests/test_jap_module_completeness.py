"""Story 8 tests for medicine-module active-row completeness checks."""

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
REQUIRED_INTRO_CELLS = [
    "B2",
    "E33",
    "E35",
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
MODULE_REQUIRED_COLUMNS = {
    "ALB_MBD": ["G", "J", "Q", "T"],
    "PZQ": ["I", "J", "K", "L", "Q"],
    "IVM": ["M"],
}


def _build_workbook(
    path: Path,
    *,
    endemicity_values: dict[str, str],
    n_value: int = 2,
    blank_cells: set[str] | None = None,
    sheet_state_overrides: dict[str, str] | None = None,
) -> None:
    workbook = Workbook()
    intro_sheet = workbook.active
    intro_sheet.title = "INTRO"
    for sheet_name in REQUIRED_SHEETS:
        if sheet_name == "INTRO":
            continue
        workbook.create_sheet(title=sheet_name)

    # Story 4/5 baseline intro values.
    for cell_address in REQUIRED_INTRO_CELLS:
        intro_sheet[cell_address] = f"value_{cell_address}"
    intro_sheet["B2"] = EXPECTED_VERSION_MARKER
    intro_sheet["E33"] = "Rwanda"
    intro_sheet["E35"] = 2026
    intro_sheet["E37"] = endemicity_values["E37"]
    intro_sheet["E39"] = endemicity_values["E39"]
    intro_sheet["E41"] = endemicity_values["E41"]
    intro_sheet["E43"] = endemicity_values["E43"]
    intro_sheet["E45"] = n_value
    intro_sheet["E48"] = "value_E48"
    intro_sheet["E49"] = "value_E49"
    intro_sheet["E50"] = "value_E50"
    intro_sheet["E51"] = "value_E51"

    has_lf = endemicity_values["E37"].strip().casefold() != "non-endemic"
    has_oncho = endemicity_values["E39"].strip().casefold() != "non-endemic"
    has_sth = endemicity_values["E41"].strip().casefold() != "non-endemic"
    has_sch = endemicity_values["E43"].strip().casefold() != "non-endemic"

    # Keep sheet states aligned with Story 6 expected profile to avoid mismatch noise.
    workbook["DEC"].sheet_state = "visible" if (has_lf and not has_oncho) else "hidden"
    workbook["IVM"].sheet_state = "visible" if has_oncho else "hidden"
    workbook["IVM+"].sheet_state = "hidden"
    workbook["ALB_MBD"].sheet_state = "visible" if (has_lf or has_sth) else "hidden"
    workbook["PZQ"].sheet_state = "visible" if has_sch else "hidden"
    workbook["DATA_POLICY"].sheet_state = "hidden"

    for sheet_name, state in (sheet_state_overrides or {}).items():
        workbook[sheet_name].sheet_state = state

    active_rows = range(10, 10 + n_value)

    # Satisfy Story 7 COUNTRY_INFO prerequisites.
    country_info = workbook["COUNTRY_INFO"]
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

    # Pre-populate module required columns so tests can blank specific cells.
    for module_sheet, columns in MODULE_REQUIRED_COLUMNS.items():
        sheet = workbook[module_sheet]
        for row in active_rows:
            for column in columns:
                sheet[f"{column}{row}"] = f"{module_sheet}_{column}_{row}"

    for sheet_cell in (blank_cells or set()):
        sheet_name, cell_address = sheet_cell.split("!")
        workbook[sheet_name][cell_address] = "   "

    workbook.save(path)


def test_alb_mbd_required_columns_missing(tmp_path: Path):
    workbook_path = tmp_path / "alb_mbd_missing.xlsx"
    _build_workbook(
        workbook_path,
        endemicity_values={"E37": "Endemic", "E39": "Non-endemic", "E41": "Non-endemic", "E43": "Non-endemic"},
        n_value=2,
        blank_cells={"ALB_MBD!G10", "ALB_MBD!T11"},
    )
    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )
    alb_findings = [f for f in result.findings if f.rule_id == "JRSM.ALB_MBD.REQUIRED_COLUMNS_G_J_Q_T"]
    assert {"G10", "T11"} <= {finding.cell for finding in alb_findings}


def test_pzq_required_columns_missing(tmp_path: Path):
    workbook_path = tmp_path / "pzq_missing.xlsx"
    _build_workbook(
        workbook_path,
        endemicity_values={"E37": "Non-endemic", "E39": "Non-endemic", "E41": "Non-endemic", "E43": "Endemic"},
        n_value=2,
        blank_cells={"PZQ!I10", "PZQ!Q11"},
    )
    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )
    pzq_findings = [f for f in result.findings if f.rule_id == "JRSM.PZQ.REQUIRED_COLUMNS_I_TO_L_AND_Q"]
    assert {"I10", "Q11"} <= {finding.cell for finding in pzq_findings}


def test_ivm_required_column_m_missing(tmp_path: Path):
    workbook_path = tmp_path / "ivm_missing.xlsx"
    _build_workbook(
        workbook_path,
        endemicity_values={"E37": "Non-endemic", "E39": "Endemic", "E41": "Non-endemic", "E43": "Non-endemic"},
        n_value=2,
        blank_cells={"IVM!M10"},
    )
    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )
    ivm_findings = [f for f in result.findings if f.rule_id == "JRSM.IVM.REQUIRED_COLUMN_M"]
    assert any(finding.cell == "M10" for finding in ivm_findings)


def test_not_applicable_modules_skip_required_checks_without_sheet_findings(tmp_path: Path):
    workbook_path = tmp_path / "not_applicable.xlsx"
    _build_workbook(
        workbook_path,
        endemicity_values={"E37": "Non-endemic", "E39": "Non-endemic", "E41": "Non-endemic", "E43": "Non-endemic"},
        n_value=2,
        blank_cells={"ALB_MBD!G10", "PZQ!I10", "IVM!M10"},
    )
    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )

    required_rule_ids = {
        "JRSM.ALB_MBD.REQUIRED_COLUMNS_G_J_Q_T",
        "JRSM.PZQ.REQUIRED_COLUMNS_I_TO_L_AND_Q",
        "JRSM.IVM.REQUIRED_COLUMN_M",
    }
    assert not any(finding.rule_id in required_rule_ids for finding in result.findings)

    assert not any(finding.sheet in {"ALB_MBD", "PZQ", "IVM"} for finding in result.findings)


def test_module_checks_apply_only_to_active_rows_not_headers(tmp_path: Path):
    workbook_path = tmp_path / "active_rows_only.xlsx"
    _build_workbook(
        workbook_path,
        endemicity_values={"E37": "Endemic", "E39": "Non-endemic", "E41": "Non-endemic", "E43": "Non-endemic"},
        n_value=1,
        blank_cells={"ALB_MBD!G8", "ALB_MBD!J9", "ALB_MBD!Q7", "ALB_MBD!T6"},
    )
    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )
    alb_findings = [f for f in result.findings if f.rule_id == "JRSM.ALB_MBD.REQUIRED_COLUMNS_G_J_Q_T"]
    assert not alb_findings


def test_all_blank_required_column_emits_aggregate_with_header_label(tmp_path: Path):
    workbook_path = tmp_path / "all_blank_column.xlsx"
    blank_cells = {f"ALB_MBD!G{row}" for row in (10, 11)}
    _build_workbook(
        workbook_path,
        endemicity_values={"E37": "Endemic", "E39": "Non-endemic", "E41": "Non-endemic", "E43": "Non-endemic"},
        n_value=2,
        blank_cells=blank_cells,
    )
    result = validate_jrsm_workbook(
        workbook_path=workbook_path,
        country="Rwanda",
        year_for_request_of_medicine=2026,
    )
    alb_findings = [f for f in result.findings if f.rule_id == "JRSM.ALB_MBD.REQUIRED_COLUMNS_G_J_Q_T"]
    aggregate_finding = next(finding for finding in alb_findings if finding.cell == "G10:G11")
    assert "Target population for STH / PreSAC / Rounds" in aggregate_finding.message
    assert not any(finding.cell in {"G10", "G11"} for finding in alb_findings)
