"""Workspace jail."""

from pathlib import Path

import pytest

from holix_office.paths import OfficePathError, resolve_inside, workspace_root


def test_relative_path_stays_inside(tmp_path: Path) -> None:
    target = resolve_inside(tmp_path, "reports/plan.docx")
    assert target == (tmp_path / "reports" / "plan.docx").resolve()


@pytest.mark.parametrize("rel", ["../secret.docx", "/etc/passwd", "reports/../../x.docx", ""])
def test_escape_is_rejected(tmp_path: Path, rel: str) -> None:
    with pytest.raises(OfficePathError):
        resolve_inside(tmp_path, rel)


def test_env_selects_the_workspace(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setenv("HOLIX_OFFICE_WORKSPACE", str(tmp_path))
    monkeypatch.setenv("HOLIX_STUDIO_WORKSPACE_ROOT", "/tmp")
    assert workspace_root() == tmp_path.resolve()
