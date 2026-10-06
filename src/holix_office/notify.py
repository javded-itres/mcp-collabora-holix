"""Tell the Studio process that an office file changed.

The MCP server is a separate stdio process, so it cannot touch the browser
websocket. It posts to Studio with the WOPI secret. The secret stays in
office.env and is never copied into the profile MCP config.
"""

from __future__ import annotations

import logging
import os
from urllib.parse import urlparse

logger = logging.getLogger(__name__)


def notify_url(wopi_base_url: str) -> str:
    """Studio notify URL as seen from the host that runs the MCP process."""
    raw = (wopi_base_url or "").strip()
    if not raw:
        return ""
    parsed = urlparse(raw)
    host = parsed.hostname or ""
    if parsed.scheme not in {"http", "https"} or not host:
        return ""
    if host in {"host.docker.internal", "0.0.0.0"}:
        host = "127.0.0.1"
    if ":" in host:
        host = f"[{host}]"
    port = f":{parsed.port}" if parsed.port else ""
    path = (parsed.path or "").rstrip("/") + "/api/office/changed"
    return f"{parsed.scheme}://{host}{port}{path}"


def _studio_notify_target() -> tuple[str, str]:
    """Studio keeps the WOPI secret in office.env. Other clients skip refresh."""
    try:
        from holix_studio.application.office_settings import load_office_runtime
    except Exception:
        return "", ""
    try:
        runtime = load_office_runtime()
    except Exception:
        logger.debug("office runtime for editor refresh failed", exc_info=True)
        return "", ""
    secret = str(runtime.get("secret") or "")
    url = notify_url(str(runtime.get("wopi_base_url") or ""))
    return secret, url


def notify_editor(profile: str, rel_path: str) -> bool:
    """Return True when Studio accepted the change notice."""
    name = (profile or os.getenv("HOLIX_PROFILE") or "").strip()
    rel = (rel_path or "").strip().replace("\\", "/").lstrip("/")
    if not name or not rel:
        return False
    secret, url = _studio_notify_target()
    if not secret or not url:
        return False
    try:
        import httpx

        response = httpx.post(
            url,
            json={"profile": name, "path": rel},
            headers={"X-Office-Notify": secret},
            timeout=2.0,
        )
    except Exception:
        logger.debug("office editor refresh failed", exc_info=True)
        return False
    return 200 <= response.status_code < 300
