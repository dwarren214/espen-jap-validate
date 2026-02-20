"""JRSM deterministic validator for Story 4 workbook/template checks."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import re
from typing import Any

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import column_index_from_string, coordinate_to_tuple, get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from ..models import Finding, ValidationEngineResult
from ..specs_loader import load_jrsm_cell_color_map, load_jrsm_formula_spec

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
EXPECTED_VERSION_MARKER = "Joint request for selected PC medicines"
MODULE_VISIBILITY_SHEETS = ["DEC", "IVM", "IVM+", "ALB_MBD", "PZQ"]
CORE_EXPECTED_VISIBLE_SHEETS = ["INTRO", "COUNTRY_INFO", "SUMMARY", "SHIPMENT"]
REQUIRED_INTRO_FIELD_RANGES = [("E37", "E45"), ("E48", "E51")]
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


def _expand_cell_range(start: str, end: str) -> list[str]:
    start_row, start_col = coordinate_to_tuple(start)
    end_row, end_col = coordinate_to_tuple(end)
    if start_col != end_col:
        raise ValueError(f"Only single-column ranges are supported: {start}:{end}")

    column_letter = start[0]
    return [f"{column_letter}{row}" for row in range(start_row, end_row + 1)]


REQUIRED_INTRO_FIELD_CELLS = [
    cell
    for range_start, range_end in REQUIRED_INTRO_FIELD_RANGES
    for cell in _expand_cell_range(range_start, range_end)
]
REQUIRED_INTRO_VALUE_CELLS = ["E37", "E39", "E41", "E43", "E45", "E48", "E49", "E50", "E51"]
MODULE_COMPLETENESS_SPECS = [
    (
        "ALB_MBD",
        ["G", "J", "Q", "T"],
        "JRSM.ALB_MBD.REQUIRED_COLUMNS_G_J_Q_T",
        "JRSM.ALB_MBD.NOT_APPLICABLE",
    ),
    (
        "PZQ",
        ["I", "J", "K", "L", "Q"],
        "JRSM.PZQ.REQUIRED_COLUMNS_I_TO_L_AND_Q",
        "JRSM.PZQ.NOT_APPLICABLE",
    ),
    (
        "IVM",
        ["M"],
        "JRSM.IVM.REQUIRED_COLUMN_M",
        "JRSM.IVM.NOT_APPLICABLE",
    ),
]
SUMMARY_SHIPMENT_REFERENCE_DOC_PATH = Path(__file__).resolve(strict=True).parents[4] / "docs/jrsm_summary_shipment_validation_reference.md"
SUMMARY_SHIPMENT_ANCHOR_CELLS = {
    "SUMMARY": ["A1", "G12", "B61", "C61"],
    "SHIPMENT": ["A1", "C4", "H4"],
}


def _has_cell_address(sheet: Worksheet, cell_address: str) -> bool:
    if coordinate_to_tuple(cell_address) in sheet._cells:
        return True
    for merged_range in sheet.merged_cells.ranges:
        if cell_address in merged_range:
            return True
    return False


def _build_finding(
    findings: list[Finding],
    *,
    rule_id: str,
    severity: str,
    message: str,
    recommendation: str,
    sheet: str | None = None,
    cell: str | None = None,
    expected: str | None = None,
    actual: str | None = None,
) -> None:
    issue_id = f"F-{len(findings) + 1:04d}"
    findings.append(
        Finding(
            issue_id=issue_id,
            rule_id=rule_id,
            severity=severity,  # type: ignore[arg-type]
            sheet=sheet,
            cell=cell,
            expected=expected,
            actual=actual,
            message=message,
            recommendation=recommendation,
        )
    )


def _derive_outcome(findings: list[Finding]) -> str:
    return "fail" if any(finding.severity == "error" for finding in findings) else "pass"


def _normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).split()).casefold()


def _normalize_year(value: Any) -> int | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value) if value.is_integer() else None
    if isinstance(value, (date, datetime)):
        return value.year

    text_value = " ".join(str(value).split()).replace(",", "")
    if not text_value:
        return None
    try:
        numeric_value = float(text_value)
    except ValueError:
        return None
    if not numeric_value.is_integer():
        return None
    return int(numeric_value)


def _normalize_positive_int(value: Any) -> int | None:
    normalized = _normalize_year(value)
    if normalized is None:
        return None
    return normalized if normalized > 0 else None


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False


def _header_text(sheet: Worksheet, column_letter: str) -> str:
    values = [sheet[f"{column_letter}{row}"].value for row in (6, 7, 8)]
    return " ".join(str(value).strip() for value in values if not _is_blank(value))


def _is_output_column_expected(
    *,
    header_text: str,
    has_lf: bool,
    has_oncho: bool,
    has_sth: bool,
    has_sch: bool,
    has_ida: bool | None,
) -> bool:
    if not header_text:
        return False

    header_lower = header_text.casefold()
    keyword_map = {
        "lf": has_lf,
        "lymphatic": has_lf,
        "filaria": has_lf,
        "oncho": has_oncho,
        "onchocer": has_oncho,
        "sth": has_sth,
        "soil-transmitted": has_sth,
        "soil transmitted": has_sth,
        "sch": has_sch,
        "schisto": has_sch,
        "ida": bool(has_ida),
    }

    matched_any = False
    matched_expected = False
    for keyword, enabled in keyword_map.items():
        if keyword in header_lower:
            matched_any = True
            matched_expected = matched_expected or enabled
    if matched_any:
        return matched_expected
    return True


def _is_formula_value(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("=")


def _normalize_formula(formula: str) -> str:
    # Normalize for deterministic comparison across spacing/case/$-anchor variants.
    normalized = re.sub(r"\s+", "", formula).upper()
    return normalized.replace("$", "")


def _translate_formula_to_target(*, formula: str, source_cell: str, target_cell: str) -> str:
    if source_cell == target_cell:
        return formula
    try:
        return Translator(formula, origin=source_cell).translate_formula(target_cell)
    except Exception:
        return formula


def _format_actual_value(value: Any) -> str:
    if _is_blank(value):
        return "BLANK_OR_MISSING"
    return str(value)


def _governance_severity(governance_class: str) -> str:
    return "error" if governance_class == "must_not_change" else "warn"


def _literal_override_rule(governance_class: str) -> str:
    if governance_class == "must_not_change":
        return "JRSM.FORMULA.PROTECTED_CELL_ENFORCEMENT"
    return "JRSM.FORMULA.EDITABLE_CELL_OVERRIDE_DETECTED"


def _governance_recommendation(governance_class: str) -> str:
    if governance_class == "must_not_change":
        return "Restore the canonical template formula for this protected governed cell."
    return "Review editable formula override and restore canonical formula if override is unintended."


def _is_anchor_addressable(sheet: Worksheet, cell_address: str) -> bool:
    try:
        coordinate_to_tuple(cell_address)
        _ = sheet[cell_address]
    except Exception:
        return False
    return True


def _evaluate_story9_formula_governance(
    *,
    workbook: Workbook,
    findings: list[Finding],
    iu_start_row: int,
    iu_end_row: int | None,
) -> None:
    formula_spec, formula_spec_error = load_jrsm_formula_spec()
    cell_color_map, cell_color_map_error = load_jrsm_cell_color_map()
    if formula_spec is None:
        _build_finding(
            findings,
            rule_id="JRSM.FORMULA.SPEC_NOT_CONFIGURED",
            severity="info",
            message="JRSM formula governance spec is not configured; Story 9 enforcement skipped.",
            recommendation="Provide src/espen_sql_api/jap_validation/specs/jrsm_formula_spec.json for Story 9 enforcement.",
            expected="Loadable JSON formula governance spec",
            actual=formula_spec_error or "UNKNOWN_ERROR",
        )

    if cell_color_map is None:
        _build_finding(
            findings,
            rule_id="JRSM.CELL_COLOR.SPEC_NOT_CONFIGURED",
            severity="info",
            message="JRSM cell-color map is not configured; continuing with governance-class defaults.",
            recommendation="Provide src/espen_sql_api/jap_validation/specs/jrsm_cell_color_map.yaml.",
            expected="Loadable YAML cell-color map",
            actual=cell_color_map_error or "UNKNOWN_ERROR",
        )
    else:
        _ = cell_color_map

    if formula_spec is None:
        return

    def evaluate_cell(
        *,
        sheet_name: str,
        cell_address: str,
        governance_class: str,
        expected_formula: str,
    ) -> None:
        if sheet_name not in workbook.sheetnames:
            return

        sheet = workbook[sheet_name]
        actual_value = sheet[cell_address].value
        severity = _governance_severity(governance_class)
        recommendation = _governance_recommendation(governance_class)

        if not _is_formula_value(actual_value):
            actual_text = _format_actual_value(actual_value)
            _build_finding(
                findings,
                rule_id="JRSM.FORMULA.GOVERNED_FORMULA_PRESENT",
                severity=severity,
                message="Governed cell is missing a formula.",
                recommendation=recommendation,
                sheet=sheet_name,
                cell=cell_address,
                expected="Formula present in governed cell",
                actual=actual_text,
            )
            _build_finding(
                findings,
                rule_id=_literal_override_rule(governance_class),
                severity=severity,
                message="Literal value detected in formula-governed cell.",
                recommendation=recommendation,
                sheet=sheet_name,
                cell=cell_address,
                expected=expected_formula,
                actual=actual_text,
            )
            return

        actual_formula = str(actual_value)
        if _normalize_formula(actual_formula) == _normalize_formula(expected_formula):
            return

        _build_finding(
            findings,
            rule_id="JRSM.FORMULA.GOVERNED_FORMULA_MATCH",
            severity=severity,
            message="Governed cell formula differs from canonical template formula.",
            recommendation=recommendation,
            sheet=sheet_name,
            cell=cell_address,
            expected=expected_formula,
            actual=actual_formula,
        )

    iu_sheet_order = formula_spec.get("iu_sheet_order", [])
    if isinstance(iu_sheet_order, list) and iu_end_row is not None:
        iu_sheets = formula_spec.get("iu_sheets", {})
        if isinstance(iu_sheets, dict):
            for sheet_name in iu_sheet_order:
                if not isinstance(sheet_name, str):
                    continue
                sheet_spec = iu_sheets.get(sheet_name)
                if not isinstance(sheet_spec, dict):
                    continue

                canonical_anchor_row = sheet_spec.get("canonical_anchor_row", iu_start_row)
                if not isinstance(canonical_anchor_row, int):
                    canonical_anchor_row = iu_start_row
                canonical_formulas_by_column = sheet_spec.get("canonical_formulas_by_column", {})
                if not isinstance(canonical_formulas_by_column, dict):
                    canonical_formulas_by_column = {}

                for governance_class, columns_key in [
                    ("must_not_change", "must_not_change_columns"),
                    ("editable_override_warn", "editable_override_warn_columns"),
                ]:
                    columns = sheet_spec.get(columns_key, [])
                    if not isinstance(columns, list):
                        continue

                    for column in columns:
                        if not isinstance(column, str):
                            continue
                        canonical_formula = canonical_formulas_by_column.get(column)
                        if not isinstance(canonical_formula, str) or not canonical_formula.startswith("="):
                            continue

                        source_cell = f"{column}{canonical_anchor_row}"
                        for row in range(iu_start_row, iu_end_row + 1):
                            target_cell = f"{column}{row}"
                            expected_formula = _translate_formula_to_target(
                                formula=canonical_formula,
                                source_cell=source_cell,
                                target_cell=target_cell,
                            )
                            evaluate_cell(
                                sheet_name=sheet_name,
                                cell_address=target_cell,
                                governance_class=governance_class,
                                expected_formula=expected_formula,
                            )

    non_iu_sheet_order = formula_spec.get("non_iu_sheet_order", [])
    if isinstance(non_iu_sheet_order, list):
        non_iu_sheets = formula_spec.get("non_iu_sheets", {})
        if isinstance(non_iu_sheets, dict):
            for sheet_name in non_iu_sheet_order:
                if not isinstance(sheet_name, str):
                    continue
                sheet_spec = non_iu_sheets.get(sheet_name)
                if not isinstance(sheet_spec, dict):
                    continue
                canonical_formulas_by_cell = sheet_spec.get("canonical_formulas_by_cell", {})
                if not isinstance(canonical_formulas_by_cell, dict):
                    canonical_formulas_by_cell = {}

                for governance_class, cells_key in [
                    ("must_not_change", "must_not_change_cells"),
                    ("editable_override_warn", "editable_override_warn_cells"),
                ]:
                    cells = sheet_spec.get(cells_key, [])
                    if not isinstance(cells, list):
                        continue

                    for cell_address in cells:
                        if not isinstance(cell_address, str):
                            continue
                        canonical_formula = canonical_formulas_by_cell.get(cell_address)
                        if not isinstance(canonical_formula, str) or not canonical_formula.startswith("="):
                            continue
                        evaluate_cell(
                            sheet_name=sheet_name,
                            cell_address=cell_address,
                            governance_class=governance_class,
                            expected_formula=canonical_formula,
                        )


def _evaluate_story10_summary_shipment_placeholder(
    *,
    workbook: Workbook,
    findings: list[Finding],
) -> None:
    if not SUMMARY_SHIPMENT_REFERENCE_DOC_PATH.exists():
        _build_finding(
            findings,
            rule_id="JRSM.SUMMARY_SHIPMENT.SPEC_NOT_CONFIGURED",
            severity="info",
            message="Detailed SUMMARY/SHIPMENT validation reference is not configured; running placeholder checks only.",
            recommendation="Add docs/jrsm_summary_shipment_validation_reference.md to enable detailed Story 10+ rules.",
            expected=str(SUMMARY_SHIPMENT_REFERENCE_DOC_PATH),
            actual="MISSING_REFERENCE_DOC",
        )

    for sheet_name in ["SUMMARY", "SHIPMENT"]:
        anchor_cells = SUMMARY_SHIPMENT_ANCHOR_CELLS[sheet_name]
        if sheet_name not in workbook.sheetnames:
            _build_finding(
                findings,
                rule_id="JRSM.SUMMARY_SHIPMENT.BASIC_ACCESS_CHECK",
                severity="warn",
                message=f"{sheet_name} basic-access check failed: required sheet is missing.",
                recommendation="Regenerate workbook from standard template and ensure required sheet exists.",
                sheet=sheet_name,
                expected=f"Sheet present with addressable anchors: {', '.join(anchor_cells)}",
                actual="MISSING_SHEET",
            )
            continue

        try:
            sheet = workbook[sheet_name]
        except Exception as exc:
            _build_finding(
                findings,
                rule_id="JRSM.SUMMARY_SHIPMENT.BASIC_ACCESS_CHECK",
                severity="warn",
                message=f"{sheet_name} basic-access check failed: sheet could not be read.",
                recommendation="Verify workbook integrity and resubmit.",
                sheet=sheet_name,
                expected=f"Readable sheet with addressable anchors: {', '.join(anchor_cells)}",
                actual=f"READ_ACCESS_ERROR: {exc}",
            )
            continue

        missing_anchors = [cell for cell in anchor_cells if not _is_anchor_addressable(sheet, cell)]
        if missing_anchors:
            _build_finding(
                findings,
                rule_id="JRSM.SUMMARY_SHIPMENT.BASIC_ACCESS_CHECK",
                severity="warn",
                message=f"{sheet_name} basic-access check failed: anchor cells are not addressable.",
                recommendation="Use a standard JRSM workbook template so placeholder anchor cells are present.",
                sheet=sheet_name,
                expected=f"Addressable anchors: {', '.join(anchor_cells)}",
                actual=", ".join([f"{sheet_name}!{cell}" for cell in missing_anchors]),
            )
            continue

        _build_finding(
            findings,
            rule_id="JRSM.SUMMARY_SHIPMENT.BASIC_ACCESS_CHECK",
            severity="info",
            message=f"{sheet_name} basic-access placeholder checks passed.",
            recommendation="No action required until detailed SUMMARY/SHIPMENT rule reference is configured.",
            sheet=sheet_name,
            expected=f"Sheet present and anchors addressable: {', '.join(anchor_cells)}",
            actual="ACCESS_OK",
        )


def validate_jrsm_workbook(
    *,
    workbook_path: Path,
    country: str,
    year_for_request_of_medicine: int,
    metadata: dict[str, Any] | None = None,
) -> ValidationEngineResult:
    """Run Story 4 workbook parsing + template conformity checks."""
    _ = (country, year_for_request_of_medicine, metadata)
    findings: list[Finding] = []
    context: dict[str, Any] = {}

    try:
        workbook = load_workbook(filename=workbook_path, data_only=False)
    except Exception as exc:
        _build_finding(
            findings,
            rule_id="JRSM.WORKBOOK.FILE_PARSE",
            severity="error",
            message="Workbook parse failed.",
            recommendation="Upload a valid JRSM .xlsx or .xlsm workbook.",
            expected="Loadable workbook in formula-preserving mode (data_only=False).",
            actual=f"WORKBOOK_PARSE_FAILED: {exc}",
        )
        return ValidationEngineResult(
            status="completed",
            validation_outcome="fail",
            findings=findings,
            context=context,
        )

    missing_sheets = [sheet_name for sheet_name in REQUIRED_SHEETS if sheet_name not in workbook.sheetnames]
    if missing_sheets:
        missing_text = ", ".join(missing_sheets)
        _build_finding(
            findings,
            rule_id="JRSM.WORKBOOK.REQUIRED_SHEETS",
            severity="error",
            message=f"Workbook is missing required sheets: {missing_text}.",
            recommendation="Regenerate workbook from the JRSM template to restore required sheets.",
            expected=", ".join(REQUIRED_SHEETS),
            actual=missing_text,
        )
        _build_finding(
            findings,
            rule_id="JRSM.TEMPLATE.SHEET_SET",
            severity="warn",
            message="Workbook sheet set does not match expected JRSM template sheet set.",
            recommendation="Use a standard JRSM workbook generated from INTRO > Generate new form.",
            expected=", ".join(REQUIRED_SHEETS),
            actual=", ".join(workbook.sheetnames),
        )
    else:
        _build_finding(
            findings,
            rule_id="JRSM.TEMPLATE.SHEET_SET",
            severity="info",
            message="Workbook sheet set matches expected JRSM template sheet names.",
            recommendation="No action required.",
            expected=", ".join(REQUIRED_SHEETS),
            actual=", ".join(workbook.sheetnames),
        )

    intro_sheet = workbook["INTRO"] if "INTRO" in workbook.sheetnames else None
    if intro_sheet is None:
        _build_finding(
            findings,
            rule_id="JRSM.TEMPLATE.REQUIRED_CELLS",
            severity="error",
            message="Cannot validate required INTRO cells because INTRO sheet is missing.",
            recommendation="Provide a workbook containing the INTRO sheet.",
            sheet="INTRO",
            expected=", ".join([f"INTRO!{address}" for address in REQUIRED_INTRO_CELLS]),
            actual="INTRO sheet missing",
        )
        _build_finding(
            findings,
            rule_id="JRSM.TEMPLATE.VERSION_MARKER",
            severity="warn",
            message="Cannot evaluate template version marker because INTRO sheet is missing.",
            recommendation="Provide a workbook with INTRO!B2 populated by template generation.",
            sheet="INTRO",
            cell="B2",
            expected=EXPECTED_VERSION_MARKER,
            actual="INTRO sheet missing",
        )
    else:
        missing_cells = [cell_address for cell_address in REQUIRED_INTRO_CELLS if not _has_cell_address(intro_sheet, cell_address)]
        if missing_cells:
            _build_finding(
                findings,
                rule_id="JRSM.TEMPLATE.REQUIRED_CELLS",
                severity="error",
                message=f"Required INTRO cells are not addressable: {', '.join(missing_cells)}.",
                recommendation="Use a standard JRSM template workbook with required template cell layout.",
                sheet="INTRO",
                expected=", ".join([f"INTRO!{address}" for address in REQUIRED_INTRO_CELLS]),
                actual=", ".join([f"INTRO!{address}" for address in missing_cells]),
            )

        b2_value = intro_sheet["B2"].value if _has_cell_address(intro_sheet, "B2") else None
        b2_text = str(b2_value).strip() if b2_value is not None else ""
        if EXPECTED_VERSION_MARKER.lower() in b2_text.lower():
            _build_finding(
                findings,
                rule_id="JRSM.TEMPLATE.VERSION_MARKER",
                severity="info",
                message="Template version marker matches expected JRSM signature.",
                recommendation="No action required.",
                sheet="INTRO",
                cell="B2",
                expected=EXPECTED_VERSION_MARKER,
                actual=b2_text,
            )
        else:
            _build_finding(
                findings,
                rule_id="JRSM.TEMPLATE.VERSION_MARKER",
                severity="warn",
                message="Template version marker does not match expected JRSM signature.",
                recommendation="Confirm workbook was generated from the JRSM v4.4 template lineage.",
                sheet="INTRO",
                cell="B2",
                expected=EXPECTED_VERSION_MARKER,
                actual=b2_text or "BLANK_OR_MISSING",
            )

        # Story 5: Request consistency checks against INTRO.
        e33_value = intro_sheet["E33"].value if _has_cell_address(intro_sheet, "E33") else None
        if _normalize_text(country) != _normalize_text(e33_value):
            _build_finding(
                findings,
                rule_id="JRSM.INTRO.COUNTRY_MATCH",
                severity="error",
                message="Country mismatch between request and workbook.",
                recommendation="Correct request country or update INTRO!E33 to the expected country name.",
                sheet="INTRO",
                cell="E33",
                expected=country,
                actual=str(e33_value).strip() if e33_value is not None else "BLANK_OR_MISSING",
            )

        e35_value = intro_sheet["E35"].value if _has_cell_address(intro_sheet, "E35") else None
        request_year = _normalize_year(year_for_request_of_medicine)
        workbook_year = _normalize_year(e35_value)
        if request_year != workbook_year:
            _build_finding(
                findings,
                rule_id="JRSM.INTRO.YEAR_MATCH",
                severity="error",
                message="Request year does not match INTRO!E35.",
                recommendation="Correct request year or update INTRO!E35 so they match numerically.",
                sheet="INTRO",
                cell="E35",
                expected=str(request_year) if request_year is not None else str(year_for_request_of_medicine),
                actual=str(workbook_year) if workbook_year is not None else str(e35_value).strip() or "BLANK_OR_MISSING",
            )

        # Story 5 required-value checks apply to driver input cells (not spacer cells).
        for cell_address in REQUIRED_INTRO_VALUE_CELLS:
            cell_value = intro_sheet[cell_address].value if _has_cell_address(intro_sheet, cell_address) else None
            if _is_blank(cell_value):
                _build_finding(
                    findings,
                    rule_id="JRSM.INTRO.REQUIRED_FIELDS_PRESENT",
                    severity="error",
                    message="Required INTRO field is blank.",
                    recommendation="Populate this required INTRO field and rerun validation.",
                    sheet="INTRO",
                    cell=cell_address,
                    expected="Non-empty value",
                    actual="BLANK_OR_MISSING",
                )

        # Story 6: Endemicity/visibility derivation + IU row-window engine.
        e37_value = intro_sheet["E37"].value if _has_cell_address(intro_sheet, "E37") else None
        e39_value = intro_sheet["E39"].value if _has_cell_address(intro_sheet, "E39") else None
        e41_value = intro_sheet["E41"].value if _has_cell_address(intro_sheet, "E41") else None
        e43_value = intro_sheet["E43"].value if _has_cell_address(intro_sheet, "E43") else None
        has_lf = _normalize_text(e37_value) != "non-endemic"
        has_oncho = _normalize_text(e39_value) != "non-endemic"
        has_sth = _normalize_text(e41_value) != "non-endemic"
        has_sch = _normalize_text(e43_value) != "non-endemic"

        n_raw_value = intro_sheet["E45"].value if _has_cell_address(intro_sheet, "E45") else None
        n_value = _normalize_positive_int(n_raw_value)
        iu_start_row = 10
        iu_end_row: int | None = None
        if n_value is None:
            _build_finding(
                findings,
                rule_id="JRSM.IU_WINDOW.N_VALID",
                severity="error",
                message="INTRO!E45 is not a valid positive integer for active IU count.",
                recommendation="Populate INTRO!E45 with a positive integer number of administrative units.",
                sheet="INTRO",
                cell="E45",
                expected="Positive integer",
                actual=str(n_raw_value).strip() if n_raw_value is not None else "BLANK_OR_MISSING",
            )
            _build_finding(
                findings,
                rule_id="JRSM.IU_WINDOW.RANGE_ENFORCED",
                severity="warn",
                message="Active IU row window could not be derived because N is invalid.",
                recommendation="Correct INTRO!E45, then rerun validation to derive active IU row range.",
                expected="Rows 10..(9+N)",
                actual="UNAVAILABLE",
            )
        else:
            iu_end_row = 9 + n_value
            _build_finding(
                findings,
                rule_id="JRSM.IU_WINDOW.N_VALID",
                severity="info",
                message="INTRO!E45 contains a valid IU row count.",
                recommendation="No action required.",
                sheet="INTRO",
                cell="E45",
                expected="Positive integer",
                actual=str(n_value),
            )
            _build_finding(
                findings,
                rule_id="JRSM.IU_WINDOW.RANGE_ENFORCED",
                severity="info",
                message="Active IU row window derived from INTRO!E45.",
                recommendation="Use this IU window for row-level deterministic checks.",
                expected="Rows 10..(9+N)",
                actual=f"Rows {iu_start_row}..{iu_end_row}",
            )

        has_ida = False
        has_ida_known = False
        country_info_sheet = workbook["COUNTRY_INFO"] if "COUNTRY_INFO" in workbook.sheetnames else None
        if n_value is None:
            _build_finding(
                findings,
                rule_id="JRSM.ENDEMICITY.IDA_DETECTION",
                severity="warn",
                message="IDA detection skipped because active IU range is unavailable.",
                recommendation="Correct INTRO!E45 so IDA detection can scan COUNTRY_INFO column Z.",
                expected="Any non-empty value in COUNTRY_INFO!Z10:Z(9+N)",
                actual="NOT_EVALUATED",
            )
        elif country_info_sheet is None:
            _build_finding(
                findings,
                rule_id="JRSM.ENDEMICITY.IDA_DETECTION",
                severity="warn",
                message="IDA detection skipped because COUNTRY_INFO sheet is missing.",
                recommendation="Provide a workbook with COUNTRY_INFO sheet present.",
                expected="COUNTRY_INFO!Z10:Z(9+N) available",
                actual="COUNTRY_INFO missing",
            )
        else:
            has_ida_known = True
            assert iu_end_row is not None
            for row in range(iu_start_row, iu_end_row + 1):
                z_value = country_info_sheet[f"Z{row}"].value
                if not _is_blank(z_value):
                    has_ida = True
                    break
            _build_finding(
                findings,
                rule_id="JRSM.ENDEMICITY.IDA_DETECTION",
                severity="info",
                message="IDA criteria detection evaluated from COUNTRY_INFO active IU rows.",
                recommendation="No action required.",
                expected=f"Scan COUNTRY_INFO!Z{iu_start_row}:Z{iu_end_row} for any non-empty value",
                actual="IDA_DETECTED" if has_ida else "IDA_NOT_DETECTED",
            )

        expected_sheet_visibility = {sheet_name: True for sheet_name in CORE_EXPECTED_VISIBLE_SHEETS}
        expected_sheet_visibility.update(
            {
                "DEC": has_lf and not has_oncho,
                "IVM": has_oncho,
                "IVM+": has_ida if has_ida_known else False,
                "ALB_MBD": has_lf or has_sth,
                "PZQ": has_sch,
                "DATA_POLICY": False,
            }
        )
        expected_visible_module_count = sum(1 for sheet_name in MODULE_VISIBILITY_SHEETS if expected_sheet_visibility[sheet_name])
        _build_finding(
            findings,
            rule_id="JRSM.ENDEMICITY.EXPECTED_MODULE_VISIBILITY",
            severity="info",
            message="Expected module visibility profile was derived from INTRO endemicity and IDA signals.",
            recommendation="Use this profile as the baseline for sheet-state mismatch checks.",
            expected=", ".join(
                [f"{sheet}={'visible' if expected_sheet_visibility[sheet] else 'hidden'}" for sheet in MODULE_VISIBILITY_SHEETS]
            ),
            actual=(
                f"has_lf={has_lf}, has_oncho={has_oncho}, has_sth={has_sth}, "
                f"has_sch={has_sch}, has_ida={has_ida if has_ida_known else 'unknown'}"
            ),
        )

        for sheet_name, expected_visible in expected_sheet_visibility.items():
            if sheet_name not in workbook.sheetnames:
                continue
            actual_state = workbook[sheet_name].sheet_state
            actual_visible = actual_state == "visible"
            if actual_visible == expected_visible:
                continue

            _build_finding(
                findings,
                rule_id="JRSM.ENDEMICITY.ACTUAL_SHEET_STATE_MISMATCH",
                severity="warn",
                message=f"Sheet visibility mismatch for {sheet_name}.",
                recommendation="Review sheet visibility against expected endemicity-driven profile.",
                sheet=sheet_name,
                expected="visible" if expected_visible else "hidden",
                actual=actual_state,
            )

        context = {
            "has_lf": has_lf,
            "has_oncho": has_oncho,
            "has_sth": has_sth,
            "has_sch": has_sch,
            "has_ida": has_ida if has_ida_known else None,
            "n": n_value,
            "iu_window_start_row": iu_start_row if n_value is not None else None,
            "iu_window_end_row": iu_end_row,
            "expected_visible_module_count": expected_visible_module_count,
            "expected_sheet_visibility": expected_sheet_visibility,
        }

        # Story 7: COUNTRY_INFO active-row completeness checks.
        if country_info_sheet is not None and iu_end_row is not None:
            active_rows = list(range(iu_start_row, iu_end_row + 1))

            # Required country ID columns B:D for every active IU row.
            for row in active_rows:
                for column in ["B", "C", "D"]:
                    cell_address = f"{column}{row}"
                    if _is_blank(country_info_sheet[cell_address].value):
                        _build_finding(
                            findings,
                            rule_id="JRSM.COUNTRY_INFO.REQUIRED_ID_COLUMNS_B_TO_D",
                            severity="error",
                            message="COUNTRY_INFO active IU row is missing required ID field.",
                            recommendation="Populate COUNTRY_INFO columns B, C, and D for each active IU row.",
                            sheet="COUNTRY_INFO",
                            cell=cell_address,
                            expected="Non-empty value",
                            actual="BLANK_OR_MISSING",
                        )

            # Endemicity-driven H:K checks.
            endemicity_required_columns = {
                "H": has_lf,
                "I": has_oncho,
                "J": has_sth,
                "K": has_sch,
            }
            for column, is_required in endemicity_required_columns.items():
                if not is_required:
                    continue

                blank_count = 0
                for row in active_rows:
                    cell_address = f"{column}{row}"
                    if _is_blank(country_info_sheet[cell_address].value):
                        blank_count += 1
                        _build_finding(
                            findings,
                            rule_id="JRSM.COUNTRY_INFO.ENDEMICITY_COLUMNS_H_TO_K",
                            severity="error",
                            message="COUNTRY_INFO endemicity-required disease column is blank for active IU row.",
                            recommendation="Populate required endemicity code columns H, I, J, K for active rows based on endemicity profile.",
                            sheet="COUNTRY_INFO",
                            cell=cell_address,
                            expected="Non-empty value",
                            actual="BLANK_OR_MISSING",
                        )

                if blank_count == len(active_rows) and _header_text(country_info_sheet, column):
                    _build_finding(
                        findings,
                        rule_id="JRSM.COUNTRY_INFO.ENDEMICITY_COLUMNS_H_TO_K",
                        severity="error",
                        message=f"COUNTRY_INFO column {column} is systematically blank across active IU rows.",
                        recommendation="Populate the endemicity-required column for active IU rows or correct endemicity configuration.",
                        sheet="COUNTRY_INFO",
                        cell=f"{column}{iu_start_row}:{column}{iu_end_row}",
                        expected="At least one non-empty value in active IU window",
                        actual="SYSTEMATICALLY_BLANK",
                    )

            # Required output block V:AD when header-populated and configuration-expected.
            required_output_columns: list[str] = []
            for column_index in range(column_index_from_string("V"), column_index_from_string("AD") + 1):
                column = get_column_letter(column_index)
                header = _header_text(country_info_sheet, column)
                if not _is_output_column_expected(
                    header_text=header,
                    has_lf=has_lf,
                    has_oncho=has_oncho,
                    has_sth=has_sth,
                    has_sch=has_sch,
                    has_ida=has_ida if has_ida_known else None,
                ):
                    continue
                required_output_columns.append(column)

                blank_count = 0
                for row in active_rows:
                    cell_address = f"{column}{row}"
                    if _is_blank(country_info_sheet[cell_address].value):
                        blank_count += 1
                        _build_finding(
                            findings,
                            rule_id="JRSM.COUNTRY_INFO.REQUIRED_OUTPUT_COLUMNS_V_TO_AD",
                            severity="error",
                            message="COUNTRY_INFO required output column is blank for active IU row.",
                            recommendation="Populate required COUNTRY_INFO output columns V:AD for active IU rows when header-populated and expected.",
                            sheet="COUNTRY_INFO",
                            cell=cell_address,
                            expected="Non-empty value",
                            actual="BLANK_OR_MISSING",
                        )

                if blank_count == len(active_rows):
                    _build_finding(
                        findings,
                        rule_id="JRSM.COUNTRY_INFO.REQUIRED_OUTPUT_COLUMNS_V_TO_AD",
                        severity="error",
                        message=f"COUNTRY_INFO required output column {column} is systematically blank across active IU rows.",
                        recommendation="Populate the required output column or verify whether this header should be present for current configuration.",
                        sheet="COUNTRY_INFO",
                        cell=f"{column}{iu_start_row}:{column}{iu_end_row}",
                        expected="At least one non-empty value in active IU window",
                        actual="SYSTEMATICALLY_BLANK",
                    )

            context["country_info_required_output_columns"] = required_output_columns

        # Story 8: medicine module row completeness checks on active IU rows.
        if iu_end_row is not None:
            active_rows = list(range(iu_start_row, iu_end_row + 1))
            for (
                module_sheet_name,
                required_columns,
                required_rule_id,
                not_applicable_rule_id,
            ) in MODULE_COMPLETENESS_SPECS:
                if module_sheet_name not in workbook.sheetnames:
                    continue

                module_sheet = workbook[module_sheet_name]
                expected_visible = bool(expected_sheet_visibility.get(module_sheet_name, False))
                actual_visible = module_sheet.sheet_state == "visible"

                if not expected_visible:
                    if not actual_visible:
                        _build_finding(
                            findings,
                            rule_id=not_applicable_rule_id,
                            severity="info",
                            message=f"{module_sheet_name} module is not applicable for this endemicity profile; row checks skipped.",
                            recommendation="No action required.",
                            sheet=module_sheet_name,
                            expected="hidden",
                            actual=module_sheet.sheet_state,
                        )
                    continue

                for row in active_rows:
                    for column in required_columns:
                        cell_address = f"{column}{row}"
                        if _is_blank(module_sheet[cell_address].value):
                            _build_finding(
                                findings,
                                rule_id=required_rule_id,
                                severity="error",
                                message=f"{module_sheet_name} active IU row is missing required module value.",
                                recommendation=f"Populate required columns {', '.join(required_columns)} for active IU rows in {module_sheet_name}.",
                                sheet=module_sheet_name,
                                cell=cell_address,
                                expected="Non-empty value",
                                actual="BLANK_OR_MISSING",
                            )

        # Story 9: submitter-day formula governance checks.
        _evaluate_story9_formula_governance(
            workbook=workbook,
            findings=findings,
            iu_start_row=iu_start_row,
            iu_end_row=iu_end_row,
        )
        _evaluate_story10_summary_shipment_placeholder(
            workbook=workbook,
            findings=findings,
        )

    return ValidationEngineResult(
        status="completed",
        validation_outcome=_derive_outcome(findings),
        findings=findings,
        context=context,
    )
