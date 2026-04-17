from __future__ import annotations

import json
from pathlib import Path


_INITIAL_VERSION = (0, 1, 1)
_STATE_PATH = Path(__file__).resolve().parent.parent / ".version_state.json"

def _normalize_version(major: int, minor: int, patch: int) -> tuple[int, int, int]:
    major = max(0, int(major))
    minor = max(0, int(minor))
    patch = max(0, int(patch))
    if patch >= 10:
        minor += patch // 10
        patch %= 10
    if minor >= 10:
        major += minor // 10
        minor %= 10
    return major, minor, patch


def _load_state() -> dict[str, object]:
    if not _STATE_PATH.exists():
        return {}
    try:
        return json.loads(_STATE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_state(state: dict[str, object]) -> None:
    try:
        _STATE_PATH.write_text(json.dumps(state, indent=2), encoding="utf-8")
    except Exception:
        pass


def _resolve_version() -> str:
    state = _load_state()
    forced_build_version = str(state.get("build_version", "")).strip() if state else ""
    if forced_build_version:
        return forced_build_version.replace(" (Beta)", "")

    major = int(state.get("major", _INITIAL_VERSION[0]))
    minor = int(state.get("minor", _INITIAL_VERSION[1]))
    patch = int(state.get("patch", _INITIAL_VERSION[2]))
    major, minor, patch = _normalize_version(major, minor, patch)
    # Keep persisted state in canonical x.y.z format where y/z are 0..9.
    _save_state(
        {
            "major": major,
            "minor": minor,
            "patch": patch,
            "fingerprint": str(state.get("fingerprint", "")),
        }
    )

    return f"v{major}.{minor}.{patch}"


APP_VERSION = _resolve_version()
