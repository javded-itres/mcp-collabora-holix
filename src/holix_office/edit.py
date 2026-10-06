"""Read and edit OOXML / ODF packages with the stdlib (zip + XML).

Text that sits in one element is replaced in place. Text split across runs is
joined inside that paragraph, then written back into the first text node.
"""

from __future__ import annotations

import os
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from holix_office.paths import OFFICE_SUFFIXES

MAX_OFFICE_BYTES = 40_000_000
_LIST_CAP = 200
_LIST_SCAN_CAP = 5000
_SKIP_DIRS = {"node_modules", ".venv", "venv", "__pycache__", ".holix"}

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
_OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
_STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
_FO = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_HEX_COLOR = re.compile(r"^#?([0-9A-Fa-f]{6})$")

_PARA_TAGS = {
    f"{{{_W}}}p",
    f"{{{_A}}}p",
    f"{{{_X}}}si",
    f"{{{_X}}}is",
    f"{{{_TEXT}}}p",
    f"{{{_TEXT}}}h",
}
_TEXT_LOCALS = {"t", "span"}
_XMLNS_PREFIX = re.compile(r'xmlns:([A-Za-z_][\w.-]*)="([^"]+)"')
_XMLNS_DEFAULT = re.compile(r'\sxmlns="([^"]+)"')


class OfficeEditError(Exception):
    """The office file could not be read or updated."""


def list_office_files(root: Path, *, limit: int = _LIST_CAP) -> list[str]:
    """Workspace-relative office paths, hidden directories skipped."""
    base = root.expanduser().resolve()
    if not base.is_dir():
        return []
    found: list[str] = []
    for dirpath, dirnames, filenames in os.walk(base):
        dirnames[:] = [
            name
            for name in dirnames
            if name not in _SKIP_DIRS and not name.startswith(".")
        ]
        for name in filenames:
            if name.startswith("."):
                continue
            if Path(name).suffix.lower() not in OFFICE_SUFFIXES:
                continue
            rel = (Path(dirpath) / name).resolve().relative_to(base).as_posix()
            found.append(rel)
            if len(found) >= _LIST_SCAN_CAP:
                return sorted(found)[:limit]
    return sorted(found)[:limit]


def read_office_text(path: Path, *, max_chars: int = 40000) -> dict[str, object]:
    path = _require_office_file(path)
    limit = max(1, min(int(max_chars), 100_000))
    paragraphs = _paragraphs(path)
    text = "\n".join(paragraphs)
    truncated = len(text) > limit
    if truncated:
        text = text[:limit]
    return {
        "ok": True,
        "path": path.name,
        "text": text,
        "truncated": truncated,
        "paragraphs": len(paragraphs),
    }


def replace_office_text(path: Path, old: str, new: str) -> int:
    """Replace every occurrence. Returns how many were replaced."""
    path = _require_office_file(path)
    needle = _require_needle(old)
    replacement = "" if new is None else str(new)
    if len(replacement) > 200_000:
        raise OfficeEditError("Replacement text is too long")
    members = _read_zip(path)
    total = 0
    changed = False
    updated: list[tuple[zipfile.ZipInfo, bytes]] = []
    names = [info.filename for info, _data in members]
    targets = set(_editable_parts(path.suffix.lower(), names))
    for info, data in members:
        if info.filename not in targets:
            updated.append((info, data))
            continue
        try:
            count, payload = _replace_in_xml(data, needle, replacement)
        except ET.ParseError:
            updated.append((info, data))
            continue
        total += count
        if count:
            changed = True
            updated.append((_clone_info(info), payload))
        else:
            updated.append((info, data))
    if changed:
        _write_zip(path, updated)
    return total


def append_office_paragraph(path: Path, text: str) -> None:
    """Add one paragraph at the end of a docx or odt."""
    path = _require_office_file(path)
    suffix = path.suffix.lower()
    body = str(text or "")
    if not body.strip():
        raise OfficeEditError("Text is empty")
    if suffix not in {".docx", ".odt"}:
        raise OfficeEditError("Appending a paragraph supports docx and odt")
    members = _read_zip(path)
    part = "word/document.xml" if suffix == ".docx" else "content.xml"
    wrote = False
    updated: list[tuple[zipfile.ZipInfo, bytes]] = []
    for info, data in members:
        if info.filename != part:
            updated.append((info, data))
            continue
        payload = _append_in_xml(data, body, kind=suffix)
        updated.append((_clone_info(info), payload))
        wrote = True
    if not wrote:
        raise OfficeEditError(f"{part} is missing")
    _write_zip(path, updated)


def format_office_text(
    path: Path,
    text: str,
    *,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    color: str | None = None,
    fill: str | None = None,
    font: str | None = None,
    size: str | float | None = None,
    style: str | None = None,
) -> int:
    """Apply run formatting to matches. Shade a paragraph only when its text is the phrase."""
    path = _require_office_file(path)
    needle = _require_needle(text)
    font_name = (font or "").strip() or None
    style_name = (style or "").strip() or None
    size_pt = _font_size(size)
    if (
        bold is None
        and italic is None
        and underline is None
        and not color
        and not fill
        and not font_name
        and size_pt is None
        and not style_name
    ):
        raise OfficeEditError(
            "Set bold, italic, underline, color, fill, font, size, or style"
        )
    color_hex = _hex_color(color, "color") if color else None
    fill_hex = _hex_color(fill, "fill") if fill else None
    suffix = path.suffix.lower()
    styled = 0
    run_format = any(
        (
            bold is not None,
            italic is not None,
            underline is not None,
            color_hex,
            fill_hex,
            font_name,
            size_pt is not None,
        )
    )
    if style_name:
        if suffix not in {".docx", ".odt"} and not run_format:
            raise OfficeEditError("Named styles are available for docx and odt")
        if suffix in {".docx", ".odt"}:
            from holix_office.rich import apply_named_style

            styled = apply_named_style(path, needle, style_name)
            if not run_format:
                return styled
    if suffix == ".xlsx":
        from holix_office.rich import format_xlsx_cells

        return format_xlsx_cells(
            path,
            needle,
            bold=bold,
            italic=italic,
            underline=underline,
            color=color_hex,
            fill=fill_hex,
            font=font_name,
            size_pt=size_pt,
        )
    if suffix not in {".docx", ".odt", ".pptx", ".ods", ".odp"}:
        raise OfficeEditError("Formatting supports docx, xlsx, pptx, odt, ods, and odp")
    members = _read_zip(path)
    total = 0
    changed = False
    updated: list[tuple[zipfile.ZipInfo, bytes]] = []
    names = [info.filename for info, _data in members]
    targets = set(_editable_parts(suffix, names))
    for info, data in members:
        if info.filename not in targets:
            updated.append((info, data))
            continue
        try:
            count, payload = _format_in_xml(
                data,
                needle,
                kind=suffix,
                bold=bold,
                italic=italic,
                underline=underline,
                color=color_hex,
                fill=fill_hex,
                font=font_name,
                size_pt=size_pt,
            )
        except ET.ParseError:
            updated.append((info, data))
            continue
        total += count
        if count:
            changed = True
            updated.append((_clone_info(info), payload))
        else:
            updated.append((info, data))
    if changed:
        _write_zip(path, updated)
    return total or styled


def _font_size(value: str | float | None) -> float | None:
    if value is None or value == "":
        return None
    text = str(value).strip().lower().removesuffix("pt")
    try:
        size = float(text)
    except ValueError as exc:
        raise OfficeEditError("size must be a number of points, for example 14") from exc
    if size <= 0 or size > 500:
        raise OfficeEditError("size is out of range")
    return size


def _hex_color(value: str | None, label: str) -> str:
    match = _HEX_COLOR.fullmatch(str(value or "").strip())
    if not match:
        raise OfficeEditError(f"{label} must be a hex color like FFFFFF")
    return match.group(1).upper()


def _require_office_file(path: Path) -> Path:
    resolved = path.expanduser().resolve()
    if resolved.suffix.lower() not in OFFICE_SUFFIXES:
        raise OfficeEditError("Not an office document")
    if not resolved.is_file():
        raise OfficeEditError("File not found")
    if resolved.stat().st_size > MAX_OFFICE_BYTES:
        raise OfficeEditError("File is too large")
    return resolved


def _require_needle(old: str) -> str:
    needle = "" if old is None else str(old)
    if needle == "":
        raise OfficeEditError("Text to replace is empty")
    if len(needle) > 200_000:
        raise OfficeEditError("Text to replace is too long")
    return needle


def _editable_parts(suffix: str, names: list[str]) -> list[str]:
    if suffix == ".docx":
        keep = {
            "word/document.xml",
            "word/footnotes.xml",
            "word/endnotes.xml",
            "word/comments.xml",
        }
        return [
            name
            for name in names
            if name in keep
            or name.startswith("word/header")
            or name.startswith("word/footer")
        ]
    if suffix == ".pptx":
        return [
            name
            for name in names
            if name.endswith(".xml")
            and (
                name.startswith("ppt/slides/slide")
                or name.startswith("ppt/notesSlides/")
            )
        ]
    if suffix == ".xlsx":
        return [
            name
            for name in names
            if name == "xl/sharedStrings.xml"
            or (name.startswith("xl/worksheets/sheet") and name.endswith(".xml"))
        ]
    if suffix in {".odt", ".ods", ".odp"}:
        return [name for name in names if name == "content.xml"]
    return []


def _read_zip(path: Path) -> list[tuple[zipfile.ZipInfo, bytes]]:
    try:
        with zipfile.ZipFile(path) as archive:
            return [
                (info, archive.read(info.filename))
                for info in archive.infolist()
                if not info.is_dir()
            ]
    except zipfile.BadZipFile as exc:
        raise OfficeEditError("File is not an office package") from exc


def _clone_info(info: zipfile.ZipInfo) -> zipfile.ZipInfo:
    cloned = zipfile.ZipInfo(filename=info.filename, date_time=info.date_time)
    cloned.compress_type = info.compress_type
    cloned.external_attr = info.external_attr
    cloned.flag_bits = info.flag_bits
    return cloned


def _write_zip(path: Path, members: list[tuple[zipfile.ZipInfo, bytes]]) -> None:
    fd, tmp_name = tempfile.mkstemp(prefix=".office-", suffix=".tmp", dir=path.parent)
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        with zipfile.ZipFile(tmp, "w") as archive:
            for info, data in members:
                archive.writestr(info, data)
        if tmp.stat().st_size > MAX_OFFICE_BYTES:
            raise OfficeEditError("Edited file is too large")
        os.replace(tmp, path)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise


def _prepare(payload: bytes) -> ET.Element:
    text = payload.decode("utf-8-sig")
    for prefix, uri in _XMLNS_PREFIX.findall(text):
        ET.register_namespace(prefix, uri)
    defaults = _XMLNS_DEFAULT.findall(text)
    if defaults:
        ET.register_namespace("", defaults[0])
    for prefix, uri in (
        ("w", _W),
        ("a", _A),
        ("x", _X),
        ("text", _TEXT),
        ("office", _OFFICE),
        ("style", _STYLE),
        ("fo", _FO),
        ("draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"),
        ("table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0"),
        ("presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"),
        ("svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"),
        ("xlink", "http://www.w3.org/1999/xlink"),
        ("manifest", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"),
    ):
        ET.register_namespace(prefix, uri)
    return ET.fromstring(text)


def _serialize(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _local(tag: str) -> str:
    if tag.startswith("{"):
        return tag.rsplit("}", 1)[-1]
    return tag


def _parent_map(root: ET.Element) -> dict[ET.Element, ET.Element]:
    return {child: parent for parent in root.iter() for child in list(parent)}


def _top_paragraphs(root: ET.Element) -> list[ET.Element]:
    parents = _parent_map(root)
    found: list[ET.Element] = []
    for el in root.iter():
        if el.tag not in _PARA_TAGS:
            continue
        cursor = parents.get(el)
        nested = False
        while cursor is not None:
            if cursor.tag in _PARA_TAGS:
                nested = True
                break
            cursor = parents.get(cursor)
        if not nested:
            found.append(el)
    return found


def _pieces(para: ET.Element) -> list[tuple[ET.Element, str]]:
    pieces: list[tuple[ET.Element, str]] = []

    def walk(el: ET.Element) -> None:
        if el.text:
            pieces.append((el, "text"))
        for child in list(el):
            walk(child)
            if child.tail:
                pieces.append((child, "tail"))

    walk(para)
    return pieces


def _piece_value(el: ET.Element, attr: str) -> str:
    return el.text if attr == "text" else (el.tail or "")


def _set_piece(el: ET.Element, attr: str, value: str) -> None:
    if attr == "text":
        el.text = value
    else:
        el.tail = value


def _replace_direct(root: ET.Element, old: str, new: str) -> int:
    count = 0
    for el in root.iter():
        if el.text and old in el.text:
            count += el.text.count(old)
            el.text = el.text.replace(old, new)
        if el.tail and old in el.tail:
            count += el.tail.count(old)
            el.tail = el.tail.replace(old, new)
    return count


def _replace_split_paragraphs(root: ET.Element, old: str, new: str) -> int:
    count = 0
    for para in _top_paragraphs(root):
        pieces = _pieces(para)
        if not pieces:
            continue
        joined = "".join(_piece_value(el, attr) for el, attr in pieces)
        if old not in joined:
            continue
        found = joined.count(old)
        updated = joined.replace(old, new)
        first_el, first_attr = pieces[0]
        _set_piece(first_el, first_attr, updated)
        for el, attr in pieces[1:]:
            _set_piece(el, attr, "")
        count += found
    return count


def _format_in_xml(
    payload: bytes,
    needle: str,
    *,
    kind: str,
    bold: bool | None,
    italic: bool | None,
    underline: bool | None,
    color: str | None,
    fill: str | None,
    font: str | None,
    size_pt: float | None,
) -> tuple[int, bytes]:
    root = _prepare(payload)
    if kind == ".docx":
        count = _format_docx(
            root,
            needle,
            bold=bold,
            italic=italic,
            underline=underline,
            color=color,
            fill=fill,
            font=font,
            size_pt=size_pt,
        )
    elif kind == ".pptx":
        from holix_office.rich import format_pptx_tree

        count = format_pptx_tree(
            root,
            needle,
            bold=bold,
            italic=italic,
            underline=underline,
            color=color,
            fill=fill,
            font=font,
            size_pt=size_pt,
        )
    else:
        count = _format_odt(
            root,
            needle,
            bold=bold,
            italic=italic,
            underline=underline,
            color=color,
            fill=fill,
            font=font,
            size_pt=size_pt,
        )
    if not count:
        return 0, payload
    return count, _serialize(root)


def _w(local: str) -> str:
    return f"{{{_W}}}{local}"


def _docx_text_nodes(para: ET.Element) -> list[tuple[ET.Element, ET.Element]]:
    parents = _parent_map(para)
    nodes: list[tuple[ET.Element, ET.Element]] = []
    for el in para.iter():
        if el.tag != _w("t") or not el.text:
            continue
        cursor: ET.Element | None = el
        run: ET.Element | None = None
        while cursor is not None and cursor is not para:
            if cursor.tag == _w("r"):
                run = cursor
                break
            cursor = parents.get(cursor)
        if run is None:
            continue
        nodes.append((run, el))
    return nodes


def _parent_of(root: ET.Element, node: ET.Element) -> ET.Element | None:
    for parent in root.iter():
        if node in list(parent):
            return parent
    return None


def _preserve_space(node: ET.Element) -> None:
    text = node.text or ""
    if text[:1].isspace() or text[-1:].isspace():
        node.set(_XML_SPACE, "preserve")


def _run_like(run: ET.Element, source: ET.Element, text: str) -> ET.Element:
    clone = ET.Element(run.tag, run.attrib)
    props = run.find(_w("rPr"))
    if props is not None:
        clone.append(ET.fromstring(ET.tostring(props)))
    node = ET.Element(source.tag, source.attrib)
    node.text = text
    _preserve_space(node)
    clone.append(node)
    return clone


def _isolate_span(
    para: ET.Element,
    run: ET.Element,
    node: ET.Element,
    local_start: int,
    local_end: int,
) -> ET.Element:
    text = node.text or ""
    local_start = max(0, min(local_start, len(text)))
    local_end = max(local_start, min(local_end, len(text)))
    if local_start == 0 and local_end == len(text):
        return run
    before, mid, after = text[:local_start], text[local_start:local_end], text[local_end:]
    parent = _parent_of(para, run)
    if parent is None or not mid:
        return run
    index = list(parent).index(run)
    node.text = mid
    _preserve_space(node)
    if before:
        parent.insert(index, _run_like(run, node, before))
        index += 1
    if after:
        parent.insert(index + 1, _run_like(run, node, after))
    return run


def _set_toggle(props: ET.Element, local: str, on: bool) -> None:
    tag = _w(local)
    node = props.find(tag)
    if node is None:
        node = ET.SubElement(props, tag)
    if on:
        node.attrib.pop(_w("val"), None)
    else:
        node.set(_w("val"), "0")


def _style_run(
    run: ET.Element,
    *,
    bold: bool | None,
    italic: bool | None,
    underline: bool | None = None,
    color: str | None = None,
    font: str | None = None,
    size_pt: float | None = None,
) -> None:
    props = run.find(_w("rPr"))
    if props is None:
        props = ET.Element(_w("rPr"))
        run.insert(0, props)
    if bold is not None:
        _set_toggle(props, "b", bold)
        _set_toggle(props, "bCs", bold)
    if italic is not None:
        _set_toggle(props, "i", italic)
        _set_toggle(props, "iCs", italic)
    if color:
        node = props.find(_w("color"))
        if node is None:
            node = ET.SubElement(props, _w("color"))
        node.set(_w("val"), color)
    if underline is not None:
        line = props.find(_w("u"))
        if line is None:
            line = ET.SubElement(props, _w("u"))
        line.set(_w("val"), "single" if underline else "none")
    if size_pt:
        half = str(int(round(size_pt * 2)))
        for local in ("sz", "szCs"):
            sized = props.find(_w(local))
            if sized is None:
                sized = ET.SubElement(props, _w(local))
            sized.set(_w("val"), half)
    if font:
        fonts = props.find(_w("rFonts"))
        if fonts is None:
            fonts = ET.SubElement(props, _w("rFonts"))
        for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
            fonts.set(_w(attr), font)


def _paragraph_is_phrase(text: str, needle: str) -> bool:
    """Shading is for a heading that is the phrase, not every mention of it."""
    phrase = needle.strip()
    return bool(phrase) and text.strip() == phrase


def _shade_paragraph(para: ET.Element, fill: str) -> None:
    props = para.find(_w("pPr"))
    if props is None:
        props = ET.Element(_w("pPr"))
        para.insert(0, props)
    shade = props.find(_w("shd"))
    if shade is None:
        shade = ET.SubElement(props, _w("shd"))
    shade.set(_w("val"), "clear")
    shade.set(_w("color"), "auto")
    shade.set(_w("fill"), fill)


def _style_range(
    para: ET.Element,
    nodes: list[tuple[ET.Element, ET.Element]],
    start: int,
    end: int,
    *,
    bold: bool | None,
    italic: bool | None,
    underline: bool | None,
    color: str | None,
    font: str | None,
    size_pt: float | None,
) -> None:
    if bold is None and italic is None and underline is None and not color and not font and not size_pt:
        return
    pos = 0
    overlapping: list[tuple[ET.Element, ET.Element, int, int]] = []
    for run, node in nodes:
        text = node.text or ""
        node_end = pos + len(text)
        if node_end > start and pos < end:
            overlapping.append((run, node, max(0, start - pos), min(len(text), end - pos)))
        pos = node_end
    for run, node, local_start, local_end in reversed(overlapping):
        target = _isolate_span(para, run, node, local_start, local_end)
        _style_run(
            target,
            bold=bold,
            italic=italic,
            underline=underline,
            color=color,
            font=font,
            size_pt=size_pt,
        )


def _format_docx(
    root: ET.Element,
    needle: str,
    *,
    bold: bool | None,
    italic: bool | None,
    underline: bool | None = None,
    color: str | None,
    fill: str | None,
    font: str | None = None,
    size_pt: float | None = None,
) -> int:
    hits = 0
    for para in _top_paragraphs(root):
        if para.tag != _w("p"):
            continue
        joined = "".join(node.text or "" for _run, node in _docx_text_nodes(para))
        shade = bool(fill) and _paragraph_is_phrase(joined, needle)
        found = 0
        guard = 0
        while guard < 10000:
            nodes = _docx_text_nodes(para)
            joined = "".join(node.text or "" for _run, node in nodes)
            start = _mask_right_matches(joined, needle, found).rfind(needle)
            if start < 0:
                break
            _style_range(
                para,
                nodes,
                start,
                start + len(needle),
                bold=bold,
                italic=italic,
                underline=underline,
                color=color,
                font=font,
                size_pt=size_pt,
            )
            found += 1
            guard += 1
        if found and shade:
            _shade_paragraph(para, fill)
        hits += found
    return hits


def _mask_right_matches(text: str, needle: str, skip: int) -> str:
    """Blank the rightmost `skip` matches so the next search finds a new one."""
    chars = list(text)
    hidden = 0
    cursor = len(text)
    while hidden < skip:
        start = text.rfind(needle, 0, cursor)
        if start < 0:
            break
        for index in range(start, start + len(needle)):
            chars[index] = "\x00"
        cursor = start
        hidden += 1
    return "".join(chars)


def _format_odt(
    root: ET.Element,
    needle: str,
    *,
    bold: bool | None,
    italic: bool | None,
    underline: bool | None = None,
    color: str | None,
    fill: str | None,
    font: str | None = None,
    size_pt: float | None = None,
) -> int:
    hits = 0
    styles = _odt_styles(root)
    serial = len(list(styles))
    for para in _top_paragraphs(root):
        if para.tag not in {f"{{{_TEXT}}}p", f"{{{_TEXT}}}h"}:
            continue
        joined = "".join(_piece_value(el, attr) for el, attr in _pieces(para))
        if needle not in joined:
            continue
        shade = bool(fill) and _paragraph_is_phrase(joined, needle)
        hosts = [para, *[el for el in para.iter() if el is not para and el.text and needle in el.text]]
        wrapped = 0
        seen: set[int] = set()
        for host in hosts:
            if id(host) in seen or not host.text or needle not in host.text:
                continue
            seen.add(id(host))
            serial += 1
            name = f"HolixText{serial}"
            _odt_text_style(
                styles,
                name,
                bold=bold,
                italic=italic,
                underline=underline,
                color=color,
                font=font,
                size_pt=size_pt,
            )
            wrapped += _wrap_text(host, needle, name)
        if wrapped == 0:
            serial += 1
            name = f"HolixPara{serial}"
            _odt_paragraph_style(
                styles,
                name,
                bold=bold,
                italic=italic,
                underline=underline,
                color=color,
                fill=fill if shade else None,
                font=font,
                size_pt=size_pt,
                parent=para.get(f"{{{_TEXT}}}style-name"),
            )
            para.set(f"{{{_TEXT}}}style-name", name)
            wrapped = joined.count(needle)
        elif shade:
            serial += 1
            name = f"HolixPara{serial}"
            _odt_paragraph_style(
                styles,
                name,
                bold=None,
                italic=None,
                color=None,
                fill=fill,
                parent=para.get(f"{{{_TEXT}}}style-name"),
            )
            para.set(f"{{{_TEXT}}}style-name", name)
        hits += wrapped
    return hits


def _odt_styles(root: ET.Element) -> ET.Element:
    styles = root.find(f"{{{_OFFICE}}}automatic-styles")
    if styles is None:
        styles = ET.Element(f"{{{_OFFICE}}}automatic-styles")
        root.insert(0, styles)
    return styles


def _odt_font_props(
    props: ET.Element,
    *,
    underline: bool | None,
    font: str | None,
    size_pt: float | None,
) -> None:
    if font:
        props.set(f"{{{_FO}}}font-family", font)
    if size_pt:
        text = str(int(size_pt)) if float(size_pt).is_integer() else str(size_pt)
        props.set(f"{{{_FO}}}font-size", f"{text}pt")
    if underline is True:
        props.set(f"{{{_STYLE}}}text-underline-style", "solid")
    elif underline is False:
        props.set(f"{{{_STYLE}}}text-underline-style", "none")


def _odt_text_style(
    styles: ET.Element,
    name: str,
    *,
    bold: bool | None,
    italic: bool | None,
    color: str | None,
    underline: bool | None = None,
    font: str | None = None,
    size_pt: float | None = None,
) -> None:
    style = ET.SubElement(styles, f"{{{_STYLE}}}style")
    style.set(f"{{{_STYLE}}}name", name)
    style.set(f"{{{_STYLE}}}family", "text")
    props = ET.SubElement(style, f"{{{_STYLE}}}text-properties")
    if bold is True:
        props.set(f"{{{_FO}}}font-weight", "bold")
    elif bold is False:
        props.set(f"{{{_FO}}}font-weight", "normal")
    if italic is True:
        props.set(f"{{{_FO}}}font-style", "italic")
    elif italic is False:
        props.set(f"{{{_FO}}}font-style", "normal")
    if color:
        props.set(f"{{{_FO}}}color", f"#{color}")
    _odt_font_props(props, underline=underline, font=font, size_pt=size_pt)


def _odt_paragraph_style(
    styles: ET.Element,
    name: str,
    *,
    bold: bool | None,
    italic: bool | None,
    color: str | None,
    fill: str | None,
    parent: str | None,
    underline: bool | None = None,
    font: str | None = None,
    size_pt: float | None = None,
) -> None:
    style = ET.SubElement(styles, f"{{{_STYLE}}}style")
    style.set(f"{{{_STYLE}}}name", name)
    style.set(f"{{{_STYLE}}}family", "paragraph")
    if parent:
        style.set(f"{{{_STYLE}}}parent-style-name", parent)
    if fill:
        paragraph = ET.SubElement(style, f"{{{_STYLE}}}paragraph-properties")
        paragraph.set(f"{{{_FO}}}background-color", f"#{fill}")
    if (
        bold is not None
        or italic is not None
        or color
        or font
        or size_pt
        or underline is not None
    ):
        _odt_text_props(
            style,
            bold=bold,
            italic=italic,
            color=color,
            underline=underline,
            font=font,
            size_pt=size_pt,
        )


def _odt_text_props(
    style: ET.Element,
    *,
    bold: bool | None,
    italic: bool | None,
    color: str | None,
    underline: bool | None = None,
    font: str | None = None,
    size_pt: float | None = None,
) -> None:
    props = ET.SubElement(style, f"{{{_STYLE}}}text-properties")
    if bold is True:
        props.set(f"{{{_FO}}}font-weight", "bold")
    elif bold is False:
        props.set(f"{{{_FO}}}font-weight", "normal")
    if italic is True:
        props.set(f"{{{_FO}}}font-style", "italic")
    elif italic is False:
        props.set(f"{{{_FO}}}font-style", "normal")
    if color:
        props.set(f"{{{_FO}}}color", f"#{color}")
    _odt_font_props(props, underline=underline, font=font, size_pt=size_pt)


def _wrap_text(host: ET.Element, needle: str, style_name: str) -> int:
    source = host.text or ""
    parts = source.split(needle)
    count = len(parts) - 1
    if count <= 0:
        return 0
    host.text = parts[0]
    for offset, tail in enumerate(parts[1:]):
        span = ET.Element(f"{{{_TEXT}}}span")
        span.set(f"{{{_TEXT}}}style-name", style_name)
        span.text = needle
        span.tail = tail
        host.insert(offset, span)
    return count


def _replace_in_xml(payload: bytes, old: str, new: str) -> tuple[int, bytes]:
    root = _prepare(payload)
    count = _replace_direct(root, old, new)
    count += _replace_split_paragraphs(root, old, new)
    if not count:
        return 0, payload
    return count, _serialize(root)


def _append_in_xml(payload: bytes, text: str, *, kind: str) -> bytes:
    root = _prepare(payload)
    if kind == ".docx":
        body = root.find(f"{{{_W}}}body")
        if body is None:
            raise OfficeEditError("Document has no body")
        paragraph = ET.Element(f"{{{_W}}}p")
        run = ET.SubElement(paragraph, f"{{{_W}}}r")
        node = ET.SubElement(run, f"{{{_W}}}t")
        node.set(_XML_SPACE, "preserve")
        node.text = text
        children = list(body)
        if children and _local(children[-1].tag) == "sectPr":
            body.insert(len(children) - 1, paragraph)
        else:
            body.append(paragraph)
    else:
        target = next((el for el in root.iter() if el.tag == f"{{{_OFFICE}}}text"), None)
        if target is None:
            raise OfficeEditError("Document has no text body")
        paragraph = ET.Element(f"{{{_TEXT}}}p")
        paragraph.text = text
        target.append(paragraph)
    return _serialize(root)


def _paragraphs(path: Path) -> list[str]:
    members = _read_zip(path)
    names = [info.filename for info, _data in members]
    targets = set(_editable_parts(path.suffix.lower(), names))
    lines: list[str] = []
    for info, data in members:
        if info.filename not in targets:
            continue
        try:
            root = _prepare(data)
        except ET.ParseError:
            continue
        paras = _top_paragraphs(root)
        if paras:
            for para in paras:
                text = "".join(_piece_value(el, attr) for el, attr in _pieces(para)).strip()
                if text:
                    lines.append(text)
            continue
        loose = []
        for el in root.iter():
            if _local(el.tag) in _TEXT_LOCALS and el.text and el.text.strip():
                loose.append(el.text.strip())
        lines.extend(loose)
    return lines
