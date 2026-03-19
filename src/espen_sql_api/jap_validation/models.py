"""Data models for JAP validation upload + validate contracts."""

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field


class UploadResponse(BaseModel):
    """Public response contract for POST /upload."""

    file_reference: str
    file_name: str
    content_type: str
    size_bytes: int
    storage_backend: str
    uploaded_at: str
    expires_at: str
    ttl_seconds: int


class UploadMetadata(UploadResponse):
    """Internal metadata sidecar shape stored for file-reference lifecycle."""

    model_config = ConfigDict(extra="ignore")

    stored_file_name: str
    source: str | None = None


class ValidateJRSMRequest(BaseModel):
    """Request contract for POST /validate/jrsm."""

    file_reference: str
    country: str
    year_for_request_of_medicine: int
    metadata: dict[str, Any] | None = None


class ExecutiveSummary(BaseModel):
    """Validation result summary counts and blockers."""

    total_findings: int
    error_count: int
    warn_count: int
    info_count: int
    major_blockers: list[str]
    major_blockers_text: str


class Finding(BaseModel):
    """Deterministic structured finding entry."""

    issue_id: str
    rule_id: str
    severity: Literal["error", "warn", "info"]
    sheet: str | None = None
    cell: str | None = None
    expected: str | None = None
    actual: str | None = None
    message: str
    recommendation: str


class ResponseTextBlock(BaseModel):
    """Text block for OCS rendering."""

    title: str
    body: str


class SheetValidationSummary(BaseModel):
    """High-level per-sheet validation rollup."""

    sheet: str
    status: Literal["passed", "issues_found", "not_evaluated"]
    error_count: int
    warn_count: int


class ValidateJRSMResponse(BaseModel):
    """Response envelope for POST /validate/jrsm."""

    run_id: str
    file_reference: str
    status: str
    validation_outcome: str
    summary_text: str
    executive_summary: ExecutiveSummary
    sheet_summaries: list[SheetValidationSummary] = Field(default_factory=list)
    findings: list[Finding]
    response_text_blocks: list[ResponseTextBlock]
    findings_truncated: bool | None = None
    findings_returned_count: int | None = None
    findings_total_count: int | None = None


class ValidationEngineResult(BaseModel):
    """Internal result shape returned by placeholder JRSM validator."""

    status: str = "completed"
    validation_outcome: str = "pass"
    findings: list[Finding] = Field(default_factory=list)
    context: dict[str, Any] = Field(default_factory=dict)
