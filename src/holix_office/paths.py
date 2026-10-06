"""Workspace jail for office documents. Paths stay inside one directory."""

from __future__ import annotations

import os
from pathlib import Path

OFFICE_SUFFIXES = frozenset({".docx", ".xlsx", ".pptx", ".odt", ".ods", ".odp"})


class OfficePathError(ValueError):
    """A path is empty, the wrong type, or outside the workspace."""


def is_office_path(path: str) -> bool:
    return Path(path).suffix.lower() in OFFICE_SUFFIXES


def workspace_root() -> Path:
    """Directory the tools may read and write.

    ``HOLIX_OFFICE_WORKSPACE`` is the public setting. Studio also sets
    ``HOLIX_STUDIO_WORKSPACE_ROOT`` to the profile workspace. Otherwise the
    process working directory is the workspace.
    """
    for key in ("HOLIX_OFFICE_WORKSPACE", "HOLIX_STUDIO_WORKSPACE_ROOT"):
        raw = (os.getenv(key) or "").strip()
        if raw:
            return Path(raw).expanduser().resolve()
    return Path.cwd().resolve()


def resolve_inside(root: Path, rel_path: str) -> Path:
    """Resolve a workspace-relative path and reject anything outside ``root``."""
    base = root.expanduser().resolve()
    text = (rel_path or "").strip().replace("\\", "/")
    if not text or text == ".":
        raise OfficePathError("Path is empty")
    if text.startswith("/") or (len(text) >= 2 and text[1] == ":"):
        raise OfficePathError("Path must stay inside the workspace")
    parts = [part for part in text.split("/") if part and part != "."]
    if not parts or any(part == ".." for part in parts):
        raise OfficePathError("Path must stay inside the workspace")
    resolved = (base.joinpath(*parts)).resolve()
    if resolved != base and base not in resolved.parents:
        raise OfficePathError("Path escapes the workspace")
    return resolved
