# JRSM SUMMARY/SHIPMENT Detailed Validation Reference

Document owner: Engineering  
Status: Draft for Story 10b implementation  
Template baseline: `docs/JRSM-template-all-sheets.xlsx`  
Calibration submission: `docs/Rwanda_JRSM_2026_PC_ESPEN_22112025.V2 xlsm.xlsm`  
Date: 2026-02-20

## 1) Purpose

Define deterministic validation rules for `SUMMARY` and `SHIPMENT` beyond Story 10 placeholder checks.

This document is the authoritative reference for Story 10b behavior and is intended to be applied only when present and version-approved.

## 2) Canonical Baseline and Bounds

1. Template version marker:
   - `INTRO!B2` contains `Joint request for selected PC medicines v.4.4`
2. Canonical logical sheet bounds:
   - `SUMMARY`: `A1:I98`
   - `SHIPMENT`: `A1:J57`
3. Validation for this story must ignore used-range inflation beyond those bounds (for example, workbooks showing row/column max at `1000/26`).
4. Merged-cell handling:
   - treat only top-left cell of each merged range as authoritative.

## 3) Rule IDs and Severity Model

1. `JRSM.SUMMARY_SHIPMENT.REFERENCE_APPLIED`
   - `info`
   - emitted once when this reference is successfully applied.
2. `JRSM.SUMMARY.FORMULA_REQUIRED`
3. `JRSM.SUMMARY.FORMULA_MATCH`
4. `JRSM.SHIPMENT.FORMULA_REQUIRED`
5. `JRSM.SHIPMENT.FORMULA_MATCH`
   - severity mapping for formula checks:
   - if cell is also governed by Story 9, reuse Story 9 governance severity (`must_not_change` => `error`, `editable_override_warn` => `warn`)
   - if cell is not governed by Story 9 (`SUMMARY!C6`, `SUMMARY!H6`), use `warn`.
6. `JRSM.SUMMARY_SHIPMENT.TABLE_BLOCK_REQUIRED`
   - severity from block definition table in Section 5.

## 4) Canonical Formula Checks

Formula normalization should follow `docs/JRSM-template-formula-spec.md` Section 4, with one additional normalization step:

1. strip workbook-scoping prefixes on externalized references before compare:
   - example: `=IF([1]INTRO!$E$43=...)` should normalize to `=IF(INTRO!$E$43=...)`.

### 4.1 SUMMARY formula cells

Validate formula presence and normalized match for:

1. `C6`, `H6`
2. `G12`, `G13`, `G14`, `G16`, `G17`, `G18`, `G19`, `G22`, `G23`, `G24`, `G26`, `G27`, `G28`
3. `B61`, `C61`, `B62`, `C62`, `B63`, `C63`, `B64`, `C64`

Canonical formula source:

1. `docs/JRSM-template-formula-spec.md` Section `Sheet: SUMMARY`

### 4.2 SHIPMENT formula cells

Validate formula presence and normalized match for:

1. `C4`, `H4`

Canonical formula source:

1. `docs/JRSM-template-formula-spec.md` Section `Sheet: SHIPMENT`

## 5) Non-Formula Table Block Requirements

`TABLE_BLOCK_REQUIRED` findings must include:

1. `sheet`
2. `cell` (single anchor or block range)
3. `expected` (block requirement text)
4. `actual` (missing/invalid cell list)
5. `message` with block identifier

### 5.1 SUMMARY blocks

| Block ID | Trigger | Required cells | Severity |
|---|---|---|---|
| `SUMMARY.BLOCK.INTERVENTION_DATES` | For each row in `43:48`, if corresponding medicine request `G` cell is `> 0` (`43->12`, `44->14`, `45->16`, `46->18`, `47->22`, `48->26`) | `C:D` (1st round month/year) and `G:H` (delivery month/year) must be non-empty | `warn` |
| `SUMMARY.BLOCK.FUNDING_BY_DISEASE` | `C61:C64 > 0` for row | `D`, `E`, and `H` on same row must be non-empty | `warn` |
| `SUMMARY.BLOCK.CONTACT_METADATA` | Always | `B72`, `C72`, `E72`, `G72`, `H72` must be non-empty | `warn` |
| `SUMMARY.BLOCK.AUTHORIZATION` | Always | `D84` and `H85` must be non-empty | `warn` |

Notes:

1. `D:E:F` medicine quantity inputs (`rows 12..28`) are intentionally not hard-required in this version because request configurations may legitimately leave rows blank.
2. Rows with no trigger condition should be skipped without finding.

### 5.2 SHIPMENT blocks

| Block ID | Trigger | Required cells | Severity |
|---|---|---|---|
| `SHIPMENT.BLOCK.CONSIGNEE_SET_1` | `B8` non-empty | `C9`, `C11`, `C12`, `C14`, `C15`, `F9`, `F11`, `F12`, `F14`, `F15` | `error` |
| `SHIPMENT.BLOCK.CONSIGNEE_SET_2` | `B17` non-empty | `C18`, `C20`, `C21`, `C23`, `C24`, `F18`, `F20`, `F21`, `F23`, `F24` | `error` |
| `SHIPMENT.BLOCK.CONSIGNEE_SET_3` | `B26` non-empty | `C27`, `C29`, `C30`, `C32`, `C33`, `F27`, `F29`, `F30`, `F32`, `F33` | `error` |
| `SHIPMENT.BLOCK.IMPORT_REQUIREMENTS` | Always | `H38`, `H40`, `H42`, `H43` must be non-empty | `error` |
| `SHIPMENT.BLOCK.IMPORT_PERMIT_LEAD_TIME` | `H38` equals `Yes` (case-insensitive) | `G39` numeric and `H39` non-empty | `error` |
| `SHIPMENT.BLOCK.IMPORT_DOCUMENTS` | Any of `H38`, `H40`, `H42`, `H43` equals `Yes` | `B44` non-empty | `warn` |

Notes:

1. `B48` additional information is optional and should not produce required-field findings.
2. `Yes/No` fields should accept normalized variants (`yes`, `YES`, trimmed).

## 6) Deduplication with Story 9

To prevent duplicate findings on overlap cells:

1. Overlap set:
   - `SUMMARY`: `G12`, `G13`, `G14`, `G16`, `G17`, `G18`, `G19`, `G22`, `G23`, `G24`, `G26`, `G27`, `G28`, `B61`, `C61`, `B62`, `C62`, `B63`, `C63`, `B64`, `C64`
   - `SHIPMENT`: `C4`, `H4`
2. If Story 9 already emitted equivalent formula-presence or formula-mismatch finding for the same cell and condition, Story 10b must suppress duplicate emission.
3. Story 10b should still validate and emit findings for non-overlap formula cells:
   - `SUMMARY!C6`
   - `SUMMARY!H6`

## 7) Deterministic Execution Order

1. Run Story 10 placeholder checks first (for continuity and explicit missing-reference signaling).
2. If this reference is present and loadable:
   - emit `JRSM.SUMMARY_SHIPMENT.REFERENCE_APPLIED` (`info`)
   - execute Section 4 formula checks
   - execute Section 5 table-block checks
   - apply Section 6 deduplication.

## 8) Calibration Notes (Non-Normative)

Observed against submitted workbook `docs/Rwanda_JRSM_2026_PC_ESPEN_22112025.V2 xlsm.xlsm`:

1. `SUMMARY` canonical formula cells `G24` and `G28` are missing formulas in submission.
2. `SHIPMENT` canonical formula cells (`C4`, `H4`) match template formulas.
3. Significant valid user-entered data exists in `SUMMARY` and `SHIPMENT` non-formula blocks, supporting block-level required checks in Section 5.
