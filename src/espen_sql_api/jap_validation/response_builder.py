"""Response composition helpers for JAP validation outputs."""

from __future__ import annotations

from uuid import uuid4

from .models import ExecutiveSummary, Finding, ResponseTextBlock, ValidateJRSMResponse, ValidationEngineResult


def _build_executive_summary(findings: list[Finding]) -> ExecutiveSummary:
    error_count = sum(1 for finding in findings if finding.severity == "error")
    warn_count = sum(1 for finding in findings if finding.severity == "warn")
    info_count = sum(1 for finding in findings if finding.severity == "info")

    major_blockers = [finding.message for finding in findings if finding.severity == "error"][:3]
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
        top_issues_body = " ".join(
            [f"{index}. {finding.message}" for index, finding in enumerate(findings[:3], start=1)]
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


def build_validation_response(
    *,
    file_reference: str,
    validation_result: ValidationEngineResult,
    findings_cap: int | None = None,
) -> ValidateJRSMResponse:
    """Build validate endpoint response envelope from validator result."""
    findings = list(validation_result.findings)
    findings_total_count = len(findings)
    truncated = findings_cap is not None and findings_total_count > findings_cap
    if truncated:
        findings = findings[:findings_cap]

    summary = _build_executive_summary(findings)
    response = ValidateJRSMResponse(
        run_id=f"run_{uuid4().hex}",
        file_reference=file_reference,
        status=validation_result.status,
        validation_outcome=validation_result.validation_outcome,
        summary_text=_build_summary_text(summary),
        executive_summary=summary,
        findings=findings,
        response_text_blocks=_build_text_blocks(findings, summary),
        findings_truncated=True if truncated else None,
        findings_returned_count=len(findings) if truncated else None,
        findings_total_count=findings_total_count if truncated else None,
    )
    return response
