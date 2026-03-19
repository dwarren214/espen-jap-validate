"""Response composition helpers for JAP validation outputs."""

from __future__ import annotations

from uuid import uuid4

from .models import (
    ExecutiveSummary,
    Finding,
    ResponseTextBlock,
    SheetValidationSummary,
    ValidateJRSMResponse,
    ValidationEngineResult,
)
from .validators.jrsm import REQUIRED_SHEETS

ROUTINE_INFO_RULE_IDS = {
    "JRSM.TEMPLATE.SHEET_SET",
    "JRSM.TEMPLATE.VERSION_MARKER",
    "JRSM.IU_WINDOW.N_VALID",
    "JRSM.IU_WINDOW.RANGE_ENFORCED",
    "JRSM.ENDEMICITY.IDA_DETECTION",
    "JRSM.ENDEMICITY.EXPECTED_MODULE_VISIBILITY",
    "JRSM.SUMMARY_SHIPMENT.BASIC_ACCESS_CHECK",
    "JRSM.SUMMARY_SHIPMENT.REFERENCE_APPLIED",
    "JRSM.ALB_MBD.NOT_APPLICABLE",
    "JRSM.PZQ.NOT_APPLICABLE",
    "JRSM.IVM.NOT_APPLICABLE",
}
EXCEPTION_RELEVANT_INFO_RULE_IDS = {
    "JRSM.FORMULA.SPEC_NOT_CONFIGURED",
    "JRSM.CELL_COLOR.SPEC_NOT_CONFIGURED",
    "JRSM.SUMMARY_SHIPMENT.SPEC_NOT_CONFIGURED",
}


def _dedupe_messages(messages: list[str], *, limit: int = 3) -> list[str]:
    seen: set[str] = set()
    deduped: list[str] = []
    for message in messages:
        normalized = " ".join(message.split())
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        deduped.append(message)
        if len(deduped) == limit:
            break
    return deduped


def _filter_findings(findings: list[Finding], *, findings_mode: str) -> list[Finding]:
    if findings_mode == "full":
        return list(findings)

    filtered: list[Finding] = []
    for finding in findings:
        if finding.severity in {"error", "warn"}:
            filtered.append(finding)
            continue
        if finding.rule_id in EXCEPTION_RELEVANT_INFO_RULE_IDS:
            filtered.append(finding)
            continue
        if finding.rule_id in ROUTINE_INFO_RULE_IDS:
            continue
    return filtered


def _build_executive_summary(findings: list[Finding]) -> ExecutiveSummary:
    error_count = sum(1 for finding in findings if finding.severity == "error")
    warn_count = sum(1 for finding in findings if finding.severity == "warn")
    info_count = sum(1 for finding in findings if finding.severity == "info")

    major_blockers = _dedupe_messages([finding.message for finding in findings if finding.severity == "error"])
    if major_blockers:
        blockers_text = "; ".join(major_blockers)
    else:
        blockers_text = "No major blockers identified."

    return ExecutiveSummary(
        total_findings=len(findings),
        error_count=error_count,
        warn_count=warn_count,
        info_count=info_count,
        major_blockers=major_blockers,
        major_blockers_text=blockers_text,
    )


def _build_summary_text(summary: ExecutiveSummary) -> str:
    base = (
        f"Validation completed with {summary.error_count} errors, "
        f"{summary.warn_count} warnings, and {summary.info_count} info findings."
    )
    if summary.major_blockers:
        return f"{base} Major blockers: {summary.major_blockers_text}"
    return base


def _build_text_blocks(findings: list[Finding], summary: ExecutiveSummary) -> list[ResponseTextBlock]:
    if summary.major_blockers:
        top_issues_body = " ".join(
            [f"{index}. {message}" for index, message in enumerate(summary.major_blockers, start=1)]
        )
    elif findings:
        fallback_messages = _dedupe_messages([finding.message for finding in findings])
        top_issues_body = " ".join(
            [f"{index}. {message}" for index, message in enumerate(fallback_messages, start=1)]
        )
    else:
        top_issues_body = "No findings were produced by the validation run."

    if summary.error_count > 0:
        next_steps_body = "Resolve error findings and rerun validation."
    elif summary.warn_count > 0:
        next_steps_body = "Review warning findings and rerun validation as needed."
    else:
        next_steps_body = "Placeholder validation completed; proceed with full rule implementation stories."

    return [
        ResponseTextBlock(title="Top Issues", body=top_issues_body),
        ResponseTextBlock(title="Next Steps", body=next_steps_body),
    ]


def _parse_missing_required_sheets(findings: list[Finding]) -> set[str]:
    missing_sheets: set[str] = set()
    for finding in findings:
        if finding.rule_id != "JRSM.WORKBOOK.REQUIRED_SHEETS" or not finding.actual:
            continue
        missing_sheets.update(sheet_name.strip() for sheet_name in finding.actual.split(",") if sheet_name.strip())
    return missing_sheets


def _build_sheet_summaries(findings: list[Finding]) -> list[SheetValidationSummary]:
    if any(finding.rule_id == "JRSM.WORKBOOK.FILE_PARSE" and finding.severity == "error" for finding in findings):
        return [
            SheetValidationSummary(sheet=sheet_name, status="not_evaluated", error_count=0, warn_count=0)
            for sheet_name in REQUIRED_SHEETS
        ]

    counts_by_sheet = {
        sheet_name: {"error_count": 0, "warn_count": 0}
        for sheet_name in REQUIRED_SHEETS
    }
    missing_required_sheets = _parse_missing_required_sheets(findings)

    for missing_sheet in missing_required_sheets:
        if missing_sheet in counts_by_sheet:
            counts_by_sheet[missing_sheet]["error_count"] = max(counts_by_sheet[missing_sheet]["error_count"], 1)

    for finding in findings:
        if finding.sheet not in counts_by_sheet:
            continue
        if finding.severity == "error":
            counts_by_sheet[finding.sheet]["error_count"] += 1
        elif finding.severity == "warn":
            counts_by_sheet[finding.sheet]["warn_count"] += 1

    sheet_summaries: list[SheetValidationSummary] = []
    for sheet_name in REQUIRED_SHEETS:
        error_count = counts_by_sheet[sheet_name]["error_count"]
        warn_count = counts_by_sheet[sheet_name]["warn_count"]
        status = "issues_found" if error_count or warn_count else "passed"
        sheet_summaries.append(
            SheetValidationSummary(
                sheet=sheet_name,
                status=status,
                error_count=error_count,
                warn_count=warn_count,
            )
        )
    return sheet_summaries


def build_validation_response(
    *,
    file_reference: str,
    validation_result: ValidationEngineResult,
    findings_mode: str = "exceptions_only",
    findings_cap: int | None = None,
) -> ValidateJRSMResponse:
    """Build validate endpoint response envelope from validator result."""
    filtered_findings = _filter_findings(validation_result.findings, findings_mode=findings_mode)
    findings_total_count = len(filtered_findings)
    effective_cap = findings_cap
    truncated = effective_cap is not None and findings_total_count > effective_cap
    findings = list(filtered_findings)
    if truncated:
        findings = findings[:effective_cap]

    summary = _build_executive_summary(filtered_findings)
    response = ValidateJRSMResponse(
        run_id=f"run_{uuid4().hex}",
        file_reference=file_reference,
        status=validation_result.status,
        validation_outcome=validation_result.validation_outcome,
        summary_text=_build_summary_text(summary),
        executive_summary=summary,
        sheet_summaries=_build_sheet_summaries(validation_result.findings),
        findings=findings,
        response_text_blocks=_build_text_blocks(filtered_findings, summary),
        findings_truncated=True if truncated else None,
        findings_returned_count=len(findings) if truncated else None,
        findings_total_count=findings_total_count if truncated else None,
    )
    return response
