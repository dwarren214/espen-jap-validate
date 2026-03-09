# JRSM Test Fixtures

This directory documents the committed fixture strategy for deterministic JRSM regression tests.

## Canonical baseline fixture

The primary golden workbook fixture is:

- `docs/JRSM-template-all-sheets.xlsx`

Tests that need a valid full-template workbook should copy that file into a temporary test directory and mutate only the cells needed for the scenario under test.

## Committed byte fixture

The committed corrupt-workbook byte fixture is:

- `tests/fixtures/jrsm/corrupt_workbook.bin`

Use it for parse-failure tests that need stable non-workbook bytes.

## Variant strategy

Most negative fixtures are generated inside tests from either:

1. `docs/JRSM-template-all-sheets.xlsx`, or
2. a small in-test workbook builder using `openpyxl`

This keeps committed binary fixture sprawl low while still making the baseline source explicit.

Typical generated variants include:

1. missing required INTRO fields
2. mismatched request country/year
3. tampered governed formulas
4. missing `COUNTRY_INFO` required columns
5. missing `ALB_MBD`/`PZQ`/`IVM` required columns
6. `SUMMARY`/`SHIPMENT` rule failures

## Test-location policy

1. committed deterministic fixtures belong in `tests/fixtures/`
2. exploratory/manual-only files belong in `tests_local/`
3. if a `tests_local/` asset becomes required for regression, promote it into `tests/fixtures/`
