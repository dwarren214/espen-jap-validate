"""Utilities for loading JRSM formula-governance specification artifacts."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

SPEC_DIR = Path(__file__).resolve(strict=True).parent / "specs"
JRSM_FORMULA_SPEC_PATH = SPEC_DIR / "jrsm_formula_spec.json"
JRSM_CELL_COLOR_MAP_PATH = SPEC_DIR / "jrsm_cell_color_map.yaml"


def _load_json_file(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    if not isinstance(payload, dict):
        raise ValueError(f"JSON payload at {path} must be an object.")
    return payload


def _load_yaml_file(path: Path) -> dict[str, Any]:
    try:
        import yaml
    except ImportError as exc:  # pragma: no cover - exercised if yaml is unavailable.
        raise RuntimeError("PyYAML is not installed.") from exc

    with path.open("r", encoding="utf-8") as handle:
        payload = yaml.safe_load(handle) or {}
    if not isinstance(payload, dict):
        raise ValueError(f"YAML payload at {path} must be a mapping.")
    return payload


def load_jrsm_formula_spec(path: Path | None = None) -> tuple[dict[str, Any] | None, str | None]:
    """Load JRSM formula governance JSON spec and return (payload, error)."""
    spec_path = path or JRSM_FORMULA_SPEC_PATH
    try:
        return _load_json_file(spec_path), None
    except FileNotFoundError:
        return None, f"Missing spec file at {spec_path}."
    except Exception as exc:  # pragma: no cover - defensive branch.
        return None, f"Failed to parse formula spec at {spec_path}: {exc}"


def load_jrsm_cell_color_map(path: Path | None = None) -> tuple[dict[str, Any] | None, str | None]:
    """Load JRSM cell-color YAML map and return (payload, error)."""
    spec_path = path or JRSM_CELL_COLOR_MAP_PATH
    try:
        return _load_yaml_file(spec_path), None
    except FileNotFoundError:
        return None, f"Missing spec file at {spec_path}."
    except Exception as exc:  # pragma: no cover - defensive branch.
        return None, f"Failed to parse cell-color spec at {spec_path}: {exc}"
