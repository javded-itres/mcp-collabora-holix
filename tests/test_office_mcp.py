"""Office document MCP: zip edits, profile registration, editor refresh."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from holix_office.edit import (
    OfficeEditError,
    append_office_paragraph,
    format_office_text,
    list_office_files,
    read_office_text,
    replace_office_text,
)
from holix_office.notify import notify_url

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _docx(path: Path, body: str) -> None:
    document = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}">
  <w:body>{body}<w:sectPr/></w:body>
</w:document>
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types></Types>")
        archive.writestr("word/document.xml", document)


def test_replace_contiguous_and_split_runs(tmp_path: Path) -> None:
    path = tmp_path / "note.docx"
    _docx(
        path,
        "<w:p><w:r><w:t>Hello</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Hel</w:t></w:r><w:r><w:t>lo</w:t></w:r></w:p>",
    )
    assert replace_office_text(path, "Hello", "Hi") == 2
    assert read_office_text(path)["text"] == "Hi\nHi"
    xml = zipfile.ZipFile(path).read("word/document.xml")
    assert b"ns0" not in xml
    assert b"w:t" in xml
    assert replace_office_text(path, "missing", "x") == 0


def test_append_paragraph_and_skip_hidden(tmp_path: Path) -> None:
    path = tmp_path / "docs" / "note.docx"
    _docx(path, "<w:p><w:r><w:t>Hello</w:t></w:r></w:p>")
    _docx(tmp_path / ".git" / "secret.docx", "<w:p><w:r><w:t>no</w:t></w:r></w:p>")
    append_office_paragraph(path, "Tail")
    assert read_office_text(path)["text"] == "Hello\nTail"
    assert list_office_files(tmp_path) == ["docs/note.docx"]


def test_replace_split_shared_string(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    xml = """<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><r><t>Hel</t></r><r><t>lo</t></r></si>
</sst>
"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/sharedStrings.xml", xml)
    assert replace_office_text(path, "Hello", "Hi") == 1
    assert read_office_text(path)["text"] == "Hi"


def test_format_phrase_without_touching_neighbors(tmp_path: Path) -> None:
    path = tmp_path / "note.docx"
    _docx(
        path,
        "<w:p><w:r><w:t>Hello Holix Agent</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Hol</w:t></w:r><w:r><w:t>ix Agent</w:t></w:r></w:p>",
    )
    with pytest.raises(OfficeEditError):
        format_office_text(path, "Holix Agent", color="fff")
    assert format_office_text(path, "Holix Agent", bold=True, color="FFFFFF", fill="000000") == 2
    xml = zipfile.ZipFile(path).read("word/document.xml").decode()
    assert "ns0" not in xml
    assert xml.count('w:fill="000000"') == 1
    assert xml.count('w:val="FFFFFF"') == 3
    assert read_office_text(path)["text"] == "Hello Holix Agent\nHolix Agent"
    root = __import__("xml.etree.ElementTree", fromlist=["ElementTree"]).fromstring(xml)
    fills = []
    for para in root.iter(f"{{{W}}}p"):
        shade = para.find(f"{{{W}}}pPr/{{{W}}}shd")
        fills.append(None if shade is None else shade.get(f"{{{W}}}fill"))
    assert fills == [None, "000000"]
    hello = [
        run
        for run in root.iter(f"{{{W}}}r")
        if "".join(node.text or "" for node in run.iter(f"{{{W}}}t")) == "Hello "
    ]
    assert len(hello) == 1
    assert hello[0].find(f"{{{W}}}b") is None


def test_format_odt_wraps_the_phrase(tmp_path: Path) -> None:
    path = tmp_path / "note.odt"
    content = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:text><text:p>Hello Holix</text:p></office:text></office:body>
</office:document-content>
"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("content.xml", content)
    assert format_office_text(path, "Holix", bold=True, color="FFFFFF") == 1
    xml = zipfile.ZipFile(path).read("content.xml").decode()
    assert "Hello" in xml
    assert 'fo:font-weight="bold"' in xml
    assert "ns0" not in xml


def test_format_notifies_only_on_a_match(tmp_path: Path, monkeypatch) -> None:
    root = tmp_path / "ws"
    _docx(root / "a.docx", "<w:p><w:r><w:t>Title</w:t></w:r></w:p>")
    monkeypatch.setenv("HOLIX_PROFILE", "ada")
    monkeypatch.setenv("HOLIX_STUDIO_WORKSPACE_ROOT", str(root))
    calls: list[tuple[str, str]] = []

    def _notify(profile: str, rel: str) -> bool:
        calls.append((profile, rel))
        return True

    monkeypatch.setattr("holix_office.notify.notify_editor", _notify)
    from holix_office.ops import office_format

    missed = office_format("a.docx", "missing", bold=True)
    assert missed["formatted"] == 0
    assert missed["editor_refresh"] is False
    hit = office_format("a.docx", "Title", bold=True, fill="000000")
    assert hit["formatted"] == 1
    assert hit["editor_refresh"] is True
    assert calls == [("ada", "a.docx")]


def test_notify_url_rewrites_docker_host() -> None:
    assert (
        notify_url("http://host.docker.internal:8788/studio")
        == "http://127.0.0.1:8788/studio/api/office/changed"
    )
    assert (
        notify_url("http://127.0.0.1:8788/studio")
        == "http://127.0.0.1:8788/studio/api/office/changed"
    )


def test_replace_notifies_only_when_text_changes(tmp_path: Path, monkeypatch) -> None:
    root = tmp_path / "ws"
    _docx(root / "a.docx", "<w:p><w:r><w:t>Hello</w:t></w:r></w:p>")
    monkeypatch.setenv("HOLIX_PROFILE", "ada")
    monkeypatch.setenv("HOLIX_STUDIO_WORKSPACE_ROOT", str(root))
    calls: list[tuple[str, str]] = []

    def _notify(profile: str, rel: str) -> bool:
        calls.append((profile, rel))
        return True

    monkeypatch.setattr("holix_office.notify.notify_editor", _notify)
    from holix_office.ops import office_replace

    missed = office_replace("a.docx", "missing", "x")
    assert missed["replaced"] == 0
    assert missed["editor_refresh"] is False
    assert calls == []
    hit = office_replace("a.docx", "Hello", "Hi")
    assert hit["replaced"] == 1
    assert hit["editor_refresh"] is True
    assert calls == [("ada", "a.docx")]

