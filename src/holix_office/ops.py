"""Office MCP operations against the profile workspace the editor opens."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from holix_office.edit import (
    OfficeEditError,
    append_office_paragraph,
    format_office_text,
    list_office_files,
    read_office_text,
    replace_office_text,
)


def _profile() -> str:
    return (os.getenv("HOLIX_PROFILE") or "default").strip() or "default"


def _root() -> Path:
    from holix_office.paths import workspace_root

    return workspace_root()


def _resolve(rel_path: str) -> Path:
    from holix_office.paths import OfficePathError, is_office_path, resolve_inside

    rel = (rel_path or "").strip()
    if not rel or not is_office_path(rel):
        raise OfficeEditError("Path must be a docx, xlsx, pptx, odt, ods, or odp file")
    try:
        return resolve_inside(_root(), rel)
    except OfficePathError as exc:
        raise OfficeEditError(str(exc)) from exc


def _rel(path: Path) -> str:
    return path.resolve().relative_to(_root()).as_posix()


def office_list() -> dict[str, Any]:
    root = _root()
    files = list_office_files(root)
    return {"ok": True, "files": files, "count": len(files)}


def office_read(path: str, max_chars: int = 40000) -> dict[str, Any]:
    resolved = _resolve(path)
    payload = read_office_text(resolved, max_chars=max_chars)
    payload["path"] = _rel(resolved)
    return payload


def office_replace(path: str, old: str, new: str) -> dict[str, Any]:
    resolved = _resolve(path)
    replaced = replace_office_text(resolved, old, new)
    rel = _rel(resolved)
    refreshed = False
    if replaced:
        from holix_office.notify import notify_editor

        refreshed = notify_editor(_profile(), rel)
    return {
        "ok": True,
        "path": rel,
        "replaced": replaced,
        "editor_refresh": refreshed,
    }


def office_format(
    path: str,
    text: str,
    *,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    color: str = "",
    fill: str = "",
    font: str = "",
    size: str = "",
    style: str = "",
) -> dict[str, Any]:
    resolved = _resolve(path)
    formatted = format_office_text(
        resolved,
        text,
        bold=bold,
        italic=italic,
        underline=underline,
        color=color.strip() or None,
        fill=fill.strip() or None,
        font=font,
        size=size,
        style=style,
    )
    rel = _rel(resolved)
    refreshed = False
    if formatted:
        from holix_office.notify import notify_editor

        refreshed = notify_editor(_profile(), rel)
    return {
        "ok": True,
        "path": rel,
        "formatted": formatted,
        "editor_refresh": refreshed,
    }


def office_media(
    path: str,
    media: str,
    *,
    slide: int = 1,
    width_cm: float = 12,
    height_cm: float = 6.8,
) -> dict[str, Any]:
    from holix_office.media import insert_office_media

    resolved = _resolve(path)
    source = _resolve_any(media)
    payload = insert_office_media(
        resolved,
        source,
        slide=slide,
        width_cm=width_cm,
        height_cm=height_cm,
    )
    return _refresh(resolved, payload)


def office_slides(
    path: str,
    action: str,
    *,
    slide: int = 1,
    title: str = "",
    body: str = "",
    layout: str = "title_body",
    notes: str = "",
) -> dict[str, Any]:
    from holix_office.slides import slide_office

    resolved = _resolve(path)
    payload = slide_office(
        resolved,
        action,
        slide=slide,
        title=title,
        body_text=body,
        layout=layout,
        notes=notes,
    )
    return _refresh(resolved, payload, changed=action.strip().lower() != "list")


def office_sheets(
    path: str,
    action: str,
    *,
    sheet: str = "",
    cell: str = "",
    value: str = "",
    formula: str = "",
    link: str = "",
    range_ref: str = "",
) -> dict[str, Any]:
    from holix_office.sheets import sheet_office

    resolved = _resolve(path)
    payload = sheet_office(
        resolved,
        action,
        sheet_name=sheet,
        cell=cell,
        value=value,
        formula=formula,
        link=link,
        range_ref=range_ref,
    )
    quiet = action.strip().lower() in {"list", "read"}
    return _refresh(resolved, payload, changed=not quiet)


def office_chart(
    path: str,
    *,
    chart_type: str = "bar",
    sheet: str = "",
    data_range: str = "",
    categories_range: str = "",
    anchor: str = "E2",
    title: str = "",
    slide: int = 1,
    categories: str = "",
    series_name: str = "",
    values: str = "",
) -> dict[str, Any]:
    from holix_office.charts import chart_office

    resolved = _resolve(path)
    payload = chart_office(
        resolved,
        chart_type=chart_type,
        sheet_name=sheet,
        data_range=data_range,
        categories_range=categories_range,
        anchor=anchor,
        title=title,
        slide=slide,
        categories=categories,
        series_name=series_name,
        values=values,
    )
    return _refresh(resolved, payload)


def office_table(
    path: str,
    action: str,
    *,
    rows_json: str = "",
    table: int = 1,
    row: int = 0,
    col: int = 0,
    value: str = "",
    slide: int = 1,
    header: bool = True,
) -> dict[str, Any]:
    from holix_office.tables import parse_rows, table_office

    resolved = _resolve(path)
    rows = parse_rows(rows_json) if action.strip().lower() == "add" else None
    payload = table_office(
        resolved,
        action,
        rows=rows,
        table=table,
        row=row,
        col=col,
        value=value,
        slide=slide,
        header=header,
    )
    return _refresh(resolved, payload, changed=action.strip().lower() != "read")


def _resolve_any(rel_path: str) -> Path:
    from holix_office.paths import OfficePathError, resolve_inside

    rel = (rel_path or "").strip()
    if not rel:
        raise OfficeEditError("Media path is empty")
    try:
        return resolve_inside(_root(), rel)
    except OfficePathError as exc:
        raise OfficeEditError(str(exc)) from exc


def _refresh(path: Path, payload: dict[str, Any], *, changed: bool = True) -> dict[str, Any]:
    rel = _rel(path)
    payload["path"] = rel
    payload["editor_refresh"] = False
    if changed:
        from holix_office.notify import notify_editor

        payload["editor_refresh"] = notify_editor(_profile(), rel)
    return payload


def office_append(path: str, text: str) -> dict[str, Any]:
    resolved = _resolve(path)
    append_office_paragraph(resolved, text)
    rel = _rel(resolved)
    from holix_office.notify import notify_editor

    return {
        "ok": True,
        "path": rel,
        "appended": True,
        "editor_refresh": notify_editor(_profile(), rel),
    }
