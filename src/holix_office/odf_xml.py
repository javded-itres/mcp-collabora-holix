"""Small ODF content.xml edits shared by sheets, slides, tables, and media."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile
from collections.abc import Callable
from pathlib import Path

from holix_office.edit import (
    OfficeEditError,
    _clone_info,
    _prepare,
    _read_zip,
    _serialize,
    _write_zip,
)

DRAW = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
SVG = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
XLINK = "http://www.w3.org/1999/xlink"
OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
MANIFEST = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
_SHEET_REF = re.compile(
    r"(?:'([^']+)'|([A-Za-z_][\w. ]*))!"
    r"(\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?)"
)
_LOCAL_REF = re.compile(
    r"(?<![\w.\[])(\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?)"
)


def read_content(path: Path) -> ET.Element:
    members = _read_zip(path)
    for _info, data in members:
        if _info.filename == "content.xml":
            return _prepare(data)
    raise OfficeEditError("content.xml is missing")


def edit_content(
    path: Path,
    change: Callable[[ET.Element], None],
    additions: list[tuple[str, str, bytes]] | None = None,
) -> None:
    """Rewrite content.xml. additions are (zip path, mime, bytes) plus a manifest entry."""
    members = _read_zip(path)
    names = {info.filename for info, _data in members}
    updated: list[tuple[zipfile.ZipInfo, bytes]] = []
    wrote = False
    extra = list(additions or [])
    for info, data in members:
        if info.filename == "content.xml":
            root = _prepare(data)
            change(root)
            updated.append((_clone_info(info), _serialize(root)))
            wrote = True
            continue
        if info.filename == "META-INF/manifest.xml" and extra:
            updated.append((_clone_info(info), _manifest(data, extra)))
            continue
        updated.append((info, data))
    if not wrote:
        raise OfficeEditError("content.xml is missing")
    for name, _mime, blob in extra:
        if name in names:
            continue
        info = zipfile.ZipInfo(name)
        info.compress_type = zipfile.ZIP_DEFLATED
        info.date_time = (2026, 1, 1, 0, 0, 0)
        updated.append((info, blob))
    _write_zip(path, updated)


def media_frame(href: str, mime: str, width_cm: float, height_cm: float, *, video: bool) -> ET.Element:
    box = ET.Element(f"{{{DRAW}}}frame")
    box.set(f"{{{SVG}}}x", "1cm")
    box.set(f"{{{SVG}}}y", "1cm")
    box.set(f"{{{SVG}}}width", _cm(width_cm))
    box.set(f"{{{SVG}}}height", _cm(height_cm))
    node = ET.SubElement(box, f"{{{DRAW}}}plugin" if video else f"{{{DRAW}}}image")
    if video:
        node.set(f"{{{DRAW}}}mime-type", mime)
    node.set(f"{{{XLINK}}}href", href)
    node.set(f"{{{XLINK}}}type", "simple")
    node.set(f"{{{XLINK}}}show", "embed")
    node.set(f"{{{XLINK}}}actuate", "onLoad")
    return box


def body(root: ET.Element, local: str) -> ET.Element:
    tag = f"{{{OFFICE}}}{local}"
    found = next((el for el in root.iter() if el.tag == tag), None)
    if found is None:
        raise OfficeEditError(f"Document has no {local} body")
    return found


def pages(root: ET.Element) -> list[ET.Element]:
    return [el for el in root.iter() if el.tag == f"{{{DRAW}}}page"]


def sheet(root: ET.Element, name: str) -> ET.Element:
    tables = [el for el in root.iter() if el.tag == f"{{{TABLE}}}table"]
    if not name:
        if tables:
            return tables[0]
        table = ET.SubElement(body(root, "spreadsheet"), f"{{{TABLE}}}table")
        table.set(f"{{{TABLE}}}name", "Sheet1")
        return table
    for table in tables:
        if table.get(f"{{{TABLE}}}name") == name:
            return table
    raise OfficeEditError(f"Sheet {name} was not found")


def add_sheet(root: ET.Element, name: str) -> None:
    tables = [el for el in root.iter() if el.tag == f"{{{TABLE}}}table"]
    if any(el.get(f"{{{TABLE}}}name") == name for el in tables):
        raise OfficeEditError(f"Sheet {name} already exists")
    table = ET.SubElement(body(root, "spreadsheet"), f"{{{TABLE}}}table")
    table.set(f"{{{TABLE}}}name", name)


def parse_a1(addr: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Za-z]{1,3})([1-9][0-9]*)", (addr or "").strip())
    if not match:
        raise OfficeEditError("Cell must look like A1")
    col = 0
    for char in match.group(1).upper():
        col = col * 26 + (ord(char) - 64)
    return col - 1, int(match.group(2)) - 1


def put_cell(table: ET.Element, addr: str, value: str, formula: str = "", link: str = "") -> None:
    col, row = parse_a1(addr)
    cell = _cell_at(_row_at(table, row), col)
    for child in list(cell):
        if _local(child.tag) == "p":
            cell.remove(child)
    if formula:
        cell.set(f"{{{TABLE}}}formula", ods_formula(formula))
        cell.set(f"{{{OFFICE}}}value-type", "float")
    else:
        cell.attrib.pop(f"{{{TABLE}}}formula", None)
        cell.set(f"{{{OFFICE}}}value-type", "string")
    paragraph = ET.SubElement(cell, f"{{{TEXT}}}p")
    if link:
        anchor = ET.SubElement(paragraph, f"{{{TEXT}}}a")
        anchor.set(f"{{{XLINK}}}href", link)
        anchor.set(f"{{{XLINK}}}type", "simple")
        anchor.text = value or link
    else:
        paragraph.text = value


def read_sheets(root: ET.Element, limit: int = 400) -> list[dict[str, object]]:
    sheets = []
    for table in root.iter():
        if table.tag != f"{{{TABLE}}}table":
            continue
        cells: list[dict[str, str]] = []
        row_index = 0
        for row in list(table):
            if row.tag != f"{{{TABLE}}}table-row":
                continue
            row_repeat = int(row.get(f"{{{TABLE}}}number-rows-repeated") or "1")
            col_index = 0
            for cell in list(row):
                if _local(cell.tag) not in {"table-cell", "covered-table-cell"}:
                    continue
                col_repeat = int(cell.get(f"{{{TABLE}}}number-columns-repeated") or "1")
                text = "".join(cell.itertext()).strip()
                formula = cell.get(f"{{{TABLE}}}formula") or ""
                link = ""
                for anchor in cell.iter():
                    if anchor.tag == f"{{{TEXT}}}a":
                        link = anchor.get(f"{{{XLINK}}}href") or ""
                if (text or formula or link) and len(cells) < limit:
                    item = {"cell": _a1(col_index, row_index), "value": text}
                    if formula:
                        item["formula"] = formula
                    if link:
                        item["link"] = link
                    cells.append(item)
                col_index += col_repeat
            row_index += row_repeat
            if row_repeat > 1000:
                break
        sheets.append({"name": table.get(f"{{{TABLE}}}name") or "Sheet", "cells": cells})
    return sheets


def ods_formula(formula: str) -> str:
    """Turn an Excel-like formula into an ODF formula."""
    text = (formula or "").strip()
    if not text:
        raise OfficeEditError("Formula is empty")
    if text.lower().startswith("of:"):
        return text
    if text.startswith("="):
        text = text[1:]

    def sheet_ref(match: re.Match[str]) -> str:
        name = match.group(1) or match.group(2)
        return _bracket(name, match.group(3))

    text = _SHEET_REF.sub(sheet_ref, text)
    text = _LOCAL_REF.sub(lambda match: _bracket("", match.group(1)), text)
    return "of:=" + text


def text_box(text: str, x: str, y: str, width: str, height: str) -> ET.Element:
    box = ET.Element(f"{{{DRAW}}}frame")
    box.set(f"{{{SVG}}}x", x)
    box.set(f"{{{SVG}}}y", y)
    box.set(f"{{{SVG}}}width", width)
    box.set(f"{{{SVG}}}height", height)
    holder = ET.SubElement(box, f"{{{DRAW}}}text-box")
    for line in str(text or "").splitlines() or [""]:
        paragraph = ET.SubElement(holder, f"{{{TEXT}}}p")
        paragraph.text = line
    return box


def _manifest(payload: bytes, files: list[tuple[str, str, bytes]]) -> bytes:
    root = _prepare(payload)
    known = {el.get(f"{{{MANIFEST}}}full-path") for el in root}
    for name, mime, _blob in files:
        if name in known:
            continue
        entry = ET.SubElement(root, f"{{{MANIFEST}}}file-entry")
        entry.set(f"{{{MANIFEST}}}full-path", name)
        entry.set(f"{{{MANIFEST}}}media-type", mime)
    return _serialize(root)


def _row_at(table: ET.Element, index: int) -> ET.Element:
    return _split_repeated(
        table,
        index,
        tag="table-row",
        repeat_attr="number-rows-repeated",
        factory=f"{{{TABLE}}}table-row",
    )


def _cell_at(row: ET.Element, index: int) -> ET.Element:
    return _split_repeated(
        row,
        index,
        tag="table-cell",
        repeat_attr="number-columns-repeated",
        factory=f"{{{TABLE}}}table-cell",
        also="covered-table-cell",
    )


def _split_repeated(parent, index, *, tag, repeat_attr, factory, also=""):
    logical = 0
    key = f"{{{TABLE}}}{repeat_attr}"
    names = {tag, also} if also else {tag}
    for el in list(parent):
        if _local(el.tag) not in names:
            continue
        repeat = int(el.get(key) or "1")
        if index < logical + repeat:
            offset = index - logical
            if repeat == 1:
                return el
            el.attrib.pop(key, None)
            at = list(parent).index(el)
            if offset:
                before = ET.Element(el.tag)
                if offset > 1:
                    before.set(key, str(offset))
                parent.insert(at, before)
                at += 1
            rest = repeat - offset - 1
            if rest:
                after = ET.Element(el.tag)
                if rest > 1:
                    after.set(key, str(rest))
                parent.insert(at + 1, after)
            return el
        logical += repeat
    while logical <= index:
        el = ET.SubElement(parent, factory)
        logical += 1
    return el


def _bracket(sheet: str, ref: str) -> str:
    clean = ref.replace("$", "")
    if ":" in clean:
        start, end = clean.split(":", 1)
        if sheet:
            return f"[{sheet}.{start}:.{end}]"
        return f"[.{start}:.{end}]"
    if sheet:
        return f"[{sheet}.{clean}]"
    return f"[.{clean}]"


def _a1(col: int, row: int) -> str:
    label = ""
    number = col + 1
    while number:
        number, rem = divmod(number - 1, 26)
        label = chr(65 + rem) + label
    return f"{label}{row + 1}"


def _cm(value: float) -> str:
    text = str(int(value)) if float(value).is_integer() else str(value)
    return f"{text}cm"


def _local(tag: str) -> str:
    if tag.startswith("{"):
        return tag.rsplit("}", 1)[-1]
    return tag
