from __future__ import annotations

import json
from hashlib import sha1
from pathlib import Path


_INITIAL_VERSION = (0, 1, 1)
_STATE_PATH = Path(__file__).resolve().parent.parent / ".version_state.json"


def _compute_source_fingerprint() -> str:
    package_dir = Path(__file__).resolve().parent
    digest = sha1()
    for path in sorted(package_dir.rglob("*.py")):
        if "__pycache__" in path.parts:
            continue
        rel = path.relative_to(package_dir).as_posix()
        stat = path.stat()
        digest.update(rel.encode("utf-8"))
        digest.update(str(stat.st_mtime_ns).encode("ascii"))
        digest.update(str(stat.st_size).encode("ascii"))
    return digest.hexdigest()


def _next_version(major: int, minor: int, patch: int) -> tuple[int, int, int]:
    patch += 1
    if patch >= 10:
        patch = 0
        minor += 1
        if minor >= 10:
            minor = 0
            major += 1
    return major, minor, patch


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
    fingerprint = _compute_source_fingerprint()
    state = _load_state()

    if not state:
        major, minor, patch = _INITIAL_VERSION
        _save_state(
            {
                "major": major,
                "minor": minor,
                "patch": patch,
                "fingerprint": fingerprint,
            }
        )
        return f"v{major}.{minor}.{patch} (Beta)"

    major = int(state.get("major", _INITIAL_VERSION[0]))
    minor = int(state.get("minor", _INITIAL_VERSION[1]))
    patch = int(state.get("patch", _INITIAL_VERSION[2]))
    major, minor, patch = _normalize_version(major, minor, patch)
    previous_fingerprint = str(state.get("fingerprint", ""))

    if previous_fingerprint != fingerprint:
        major, minor, patch = _next_version(major, minor, patch)
        state = {
            "major": major,
            "minor": minor,
            "patch": patch,
            "fingerprint": fingerprint,
        }
        _save_state(state)
    else:
        # Keep persisted state in canonical x.y.z format where y/z are 0..9.
        _save_state(
            {
                "major": major,
                "minor": minor,
                "patch": patch,
                "fingerprint": previous_fingerprint,
            }
        )

    return f"v{major}.{minor}.{patch} (Beta)"


APP_VERSION = _resolve_version()
