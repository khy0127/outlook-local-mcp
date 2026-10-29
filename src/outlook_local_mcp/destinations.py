"""
destinations.py

Where saved attachments go. Instead of a fixed Downloads folder, files land
in folders Claude Desktop / Cowork can already open:

- by default, '<Cowork user files folder>/Outlook Attachments'
  (coworkUserFilesPath in claude_desktop_config.json);
- or a named destination: a Cowork project (space) folder or any folder the
  user has trusted in Cowork, e.g. destination='Project Alpha';
- or any absolute path.

Everything is read from Claude Desktop's own config on each call, so new
Cowork projects show up without restarting this server.
"""

from __future__ import annotations

import glob
import json
import os
from typing import Any

APPDATA_CLAUDE = os.path.join(os.environ.get("APPDATA", ""), "Claude")
DESKTOP_CONFIG = os.path.join(APPDATA_CLAUDE, "claude_desktop_config.json")
DEFAULT_SUBFOLDER = "Outlook Attachments"


class DestinationError(ValueError):
    pass


def _load_json(path: str) -> Any:
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return None


def default_dir() -> str:
    override = os.environ.get("OUTLOOK_MCP_ATTACHMENT_DIR")
    if override:
        return override
    config = _load_json(DESKTOP_CONFIG) or {}
    base = config.get("coworkUserFilesPath") or os.path.join(os.path.expanduser("~"), "Claude")
    return os.path.join(base, DEFAULT_SUBFOLDER)


def known_destinations() -> list[dict[str, str]]:
    """Named folders: Cowork projects first, then folders trusted in Cowork."""
    found: list[tuple[str, str, str]] = []  # (name, path, kind)

    for spaces_file in glob.glob(os.path.join(APPDATA_CLAUDE, "local-agent-mode-sessions", "*", "*", "spaces.json")):
        data = _load_json(spaces_file) or {}
        for space in data.get("spaces", []):
            folders = [f.get("path") if isinstance(f, dict) else f for f in space.get("folders", [])]
            folders = [f for f in folders if f]
            for i, path in enumerate(folders):
                name = space.get("name") or os.path.basename(path)
                found.append((name if i == 0 else f"{name} / {os.path.basename(path)}", path, "cowork project"))

    prefs = (_load_json(DESKTOP_CONFIG) or {}).get("preferences", {})
    for path in prefs.get("localAgentModeTrustedFolders", []):
        found.append((os.path.basename(path.rstrip("\\/")) or path, path, "trusted folder"))

    out: list[dict[str, str]] = []
    seen_paths: set[str] = set()
    seen_names: set[str] = set()
    for name, path, kind in found:
        key = os.path.normcase(os.path.normpath(path))
        if key in seen_paths or not os.path.isdir(path):
            continue
        seen_paths.add(key)
        if name.lower() in seen_names:  # e.g. Desktop\Project Alpha vs OneDrive\...\Project Alpha
            parent = os.path.basename(os.path.dirname(path.rstrip("\\/")))
            name = f"{name} ({parent})"
        seen_names.add(name.lower())
        out.append({"name": name, "path": path, "kind": kind})
    return out


def resolve(destination: str | None, subfolder: str | None) -> tuple[str, str]:
    """Return (directory, label). The directory is created if missing."""
    if not destination:
        base, label = default_dir(), "default (Cowork files)"
        sub = subfolder or ""
    else:
        dest = destination.strip()
        if os.path.isabs(dest):
            base, label = dest, dest
        else:
            options = known_destinations()
            exact = [d for d in options if d["name"].lower() == dest.lower()]
            partial = [d for d in options if dest.lower() in d["name"].lower()]
            matches = exact or partial
            if len(matches) != 1:
                names = ", ".join(d["name"] for d in (matches or options))
                problem = "matches several" if matches else "matches no"
                raise DestinationError(f"'{destination}' {problem} known destination. Choose one of: {names}")
            base, label = matches[0]["path"], matches[0]["name"]
        sub = DEFAULT_SUBFOLDER if subfolder is None else subfolder

    sub = (sub or "").strip().strip("\\/")
    if sub in (".", ""):
        directory = base
    else:
        parts = [p for p in sub.replace("\\", "/").split("/") if p]
        if any(p in ("..", ".") or ":" in p for p in parts):
            raise DestinationError("subfolder must be a plain relative path inside the destination.")
        directory = os.path.join(base, *parts)
    os.makedirs(directory, exist_ok=True)
    return directory, label
