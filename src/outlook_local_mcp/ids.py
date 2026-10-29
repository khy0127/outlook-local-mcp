"""
ids.py

Short email ids ('m1k3') in place of 140-character EntryIDs.

Claude Desktop starts one server process per chat / Cowork session, so the
mapping lives in a small SQLite file shared by all of them: an id handed out
by one process works in every other one, and across restarts. The file holds
only id <-> EntryID pairs, never mail content. Ids are never reused; a move
made through this server repoints the id at the moved email.
"""

from __future__ import annotations

import os
import re
import sqlite3
import threading

DB_PATH = os.environ.get("OUTLOOK_MCP_ID_DB") or os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "outlook-local-mcp", "ids.sqlite3"
)
HANDLE_RE = re.compile(r"^m([0-9a-z]+)$")
_OLD_FORMAT_RE = re.compile(r"^[a-z]{2}\.\d+$")  # per-process ids from before ids were shared
_DIGITS = "0123456789abcdefghijklmnopqrstuvwxyz"

_lock = threading.Lock()
_conn: sqlite3.Connection | None = None


class UnknownId(KeyError):
    pass


def _db() -> sqlite3.Connection:
    global _conn
    if _conn is None:
        os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
        conn = sqlite3.connect(DB_PATH, timeout=15, isolation_level=None, check_same_thread=False)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("CREATE TABLE IF NOT EXISTS handles (n INTEGER PRIMARY KEY AUTOINCREMENT, entry TEXT NOT NULL UNIQUE)")
        conn.execute("CREATE TABLE IF NOT EXISTS retired (entry TEXT PRIMARY KEY, n INTEGER NOT NULL)")
        _conn = conn
    return _conn


def _encode(n: int) -> str:
    out = ""
    while True:
        n, r = divmod(n, 36)
        out = _DIGITS[r] + out
        if n == 0:
            return "m" + out


def _decode(handle: str) -> int:
    return int(HANDLE_RE.match(handle).group(1), 36)


def looks_like_id(value: str) -> bool:
    value = value.strip()
    return bool(HANDLE_RE.match(value) or _OLD_FORMAT_RE.match(value))


def handle(entry_id: str) -> str:
    with _lock:
        db = _db()
        db.execute("INSERT OR IGNORE INTO handles(entry) VALUES (?)", (entry_id,))
        return _encode(db.execute("SELECT n FROM handles WHERE entry = ?", (entry_id,)).fetchone()[0])


def entry(item_id: str) -> str:
    """Short id -> EntryID. Anything that isn't a short id is assumed to be a raw EntryID."""
    item_id = item_id.strip()
    if _OLD_FORMAT_RE.match(item_id):
        raise UnknownId(f"'{item_id}' is an id from an older version of this server. List or search again.")
    if not HANDLE_RE.match(item_id):
        return item_id
    with _lock:
        row = _db().execute("SELECT entry FROM handles WHERE n = ?", (_decode(item_id),)).fetchone()
    if row is None:
        raise UnknownId(f"Unknown id '{item_id}'. List or search again to get current ids.")
    return row[0]


def rebind(item_id_or_entry: str, new_entry: str) -> str:
    """After a move: point the same short id at the email's new EntryID."""
    with _lock:
        db = _db()
        if HANDLE_RE.match(item_id_or_entry):
            n = _decode(item_id_or_entry)
        else:
            row = db.execute("SELECT n FROM handles WHERE entry = ?", (item_id_or_entry,)).fetchone()
            n = row[0] if row else None
        if n is None:
            db.execute("INSERT OR IGNORE INTO handles(entry) VALUES (?)", (new_entry,))
            return _encode(db.execute("SELECT n FROM handles WHERE entry = ?", (new_entry,)).fetchone()[0])
        db.execute("BEGIN IMMEDIATE")
        try:
            old = db.execute("SELECT entry FROM handles WHERE n = ?", (n,)).fetchone()
            if old:
                db.execute("INSERT OR REPLACE INTO retired(entry, n) VALUES (?, ?)", (old[0], n))
            db.execute("DELETE FROM handles WHERE entry = ? AND n != ?", (new_entry, n))
            db.execute("UPDATE handles SET entry = ? WHERE n = ?", (new_entry, n))
            db.execute("COMMIT")
        except Exception:
            db.execute("ROLLBACK")
            raise
        return _encode(n)


def is_retired(entry_id: str) -> bool:
    with _lock:
        return _db().execute("SELECT 1 FROM retired WHERE entry = ?", (entry_id,)).fetchone() is not None
