"""
outlook_client.py

Operations on classic desktop Outlook (OUTLOOK.EXE) through the COM object
model. Derived from r-rohit/outlook-mcp-readonly (MIT), extended with
guarded move/archive, metadata updates and draft creation.

SAFETY MODEL
------------
1. Nothing in this package can send mail or delete anything. The COM
   methods that could (Send, Delete, PermanentlyDelete, Folder.MoveTo,
   Display, ExecuteMso, SendAndReceive, ...) are never referenced;
   tests/test_safety.py parses the source and fails if one appears, or if
   an attribute is looked up dynamically (getattr with a non-literal name).
2. The only call to ``.Move(`` is in ``_guarded_move``. It refuses any
   target that is, or sits under, Deleted Items, Junk Email, Outbox, Sync
   Issues or Recoverable Items, and any folder that does not hold mail.
   Moved mail therefore always stays in a normal, visible folder.
3. Drafts are created with ``.Save()`` only. ``Reply()``, ``ReplyAll()`` and
   ``Forward()`` merely return a new unsent item. ``update_draft`` refuses
   anything that is not an unsent message inside a Drafts folder.
4. Every write is appended to a local JSONL audit log that records the
   source folder of each move, so any move can be reversed.
"""

from __future__ import annotations

import datetime as _dt
import html as _html
import json
import os
import re
import time
from typing import Any

import win32com.client

from . import destinations, ids

# Outlook constants, hardcoded so we don't depend on the makepy cache.
OL_FOLDER_DELETED = 3
OL_FOLDER_OUTBOX = 4
OL_FOLDER_SENT = 5
OL_FOLDER_INBOX = 6
OL_FOLDER_CALENDAR = 9
OL_FOLDER_DRAFTS = 16
OL_FOLDER_CONFLICTS = 19
OL_FOLDER_SYNC_ISSUES = 20
OL_FOLDER_LOCAL_FAILURES = 21
OL_FOLDER_SERVER_FAILURES = 22
OL_FOLDER_JUNK = 23

OL_MAIL_ITEM = 0  # CreateItem type and Folder.DefaultItemType for mail folders
OL_CLASS_FOLDER = 2
OL_CLASS_MAIL = 43
OL_CLASS_MEETING = {53, 54, 55, 56, 57}  # request / cancellation / responses
OL_FORMAT_HTML = 2
OL_MARK_NO_DATE = 4
OL_MARK_COMPLETE = 5

PR_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
PR_FLAG_STATUS = "http://schemas.microsoft.com/mapi/proptag/0x10900003"
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".emz", ".wmz", ".tif", ".tiff"}
# Tables hand out short-term EntryIDs in cached mode; this column gives the durable one.
PR_LONGTERM_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x66700102"

MAX_ITEMS_HARD_CAP = 500
MAX_BATCH = 200

FOLDER_ALIASES = {
    "inbox": OL_FOLDER_INBOX,
    "drafts": OL_FOLDER_DRAFTS,
    "sent": OL_FOLDER_SENT,
    "sent items": OL_FOLDER_SENT,
    "deleted": OL_FOLDER_DELETED,
    "deleted items": OL_FOLDER_DELETED,
    "junk": OL_FOLDER_JUNK,
    "junk email": OL_FOLDER_JUNK,
    "outbox": OL_FOLDER_OUTBOX,
}
ARCHIVE_NAMES = ("archive", "存档", "归档", "封存", "archives", "archiv")

# Folders mail may never be moved into (checked against the target and all
# of its ancestors, by EntryID per store and by name as a backstop).
_FORBIDDEN_TARGET_DEFAULTS = (
    OL_FOLDER_DELETED,
    OL_FOLDER_OUTBOX,
    OL_FOLDER_JUNK,
    OL_FOLDER_SYNC_ISSUES,
    OL_FOLDER_CONFLICTS,
    OL_FOLDER_LOCAL_FAILURES,
    OL_FOLDER_SERVER_FAILURES,
)
_FORBIDDEN_TARGET_NAMES = {
    "deleted items", "已删除邮件", "已删除的邮件", "recoverable items",
    "junk email", "junk e-mail", "垃圾邮件", "垃圾电子邮件",
    "outbox", "发件箱", "sync issues", "同步问题",
}
# Housekeeping folders hidden from listings and "all folders" searches.
_HIDDEN_FOLDER_NAMES = {
    "conversation action settings", "quick step settings", "yammer root",
    "social activity notifications", "sync issues", "同步问题", "rss feeds",
}

AUDIT_LOG_PATH = os.environ.get("OUTLOOK_MCP_AUDIT_LOG") or os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "outlook-local-mcp", "audit.jsonl"
)


class OutlookError(RuntimeError):
    """A problem worth reporting to the model as a plain message."""


# --------------------------------------------------------------------------
# Connection and lookups
# --------------------------------------------------------------------------


def _app():
    try:
        return win32com.client.Dispatch("Outlook.Application")
    except Exception as exc:  # noqa: BLE001
        raise OutlookError(
            "Could not attach to Outlook via COM. Make sure classic desktop Outlook "
            "(OUTLOOK.EXE) is installed and running. The 'new Outlook' (olk.exe) has no COM support."
        ) from exc


def _ns():
    return _app().GetNamespace("MAPI")


def _same(ns, a, b) -> bool:
    try:
        return bool(ns.CompareEntryIDs(a.EntryID, b.EntryID))
    except Exception:  # noqa: BLE001
        return False


# Short ids ('m1k3') handed to the model instead of 140-character EntryIDs; shared by all
# server processes via ids.py, and kept pointing at an email when we move it.
_handle = ids.handle


def _entry_id(item_id: str) -> str:
    try:
        return ids.entry(item_id)
    except ids.UnknownId as exc:
        raise OutlookError(exc.args[0]) from exc


def _get_item(ns, item_id: str):
    entry = _entry_id(item_id)
    try:
        return ns.GetItemFromID(entry)
    except Exception:  # noqa: BLE001
        pass
    for store in ns.Stores:
        try:
            return ns.GetItemFromID(entry, store.StoreID)
        except Exception:  # noqa: BLE001
            continue
    raise OutlookError(f"Item '{item_id}' not found. It may have been moved in Outlook; search for it again.")


def _child(folder, name: str):
    wanted = name.strip().lower()
    names = []
    for f in folder.Folders:
        if f.Name.lower() == wanted:
            return f
        names.append(f.Name)
    raise OutlookError(
        f"No folder '{name}' under '{_folder_display(folder)}'. Available: {', '.join(names[:40]) or '(none)'}"
    )


def _default_store_name(ns) -> str:
    return ns.DefaultStore.GetRootFolder().Name


def _archive_folder(ns):
    override = os.environ.get("OUTLOOK_MCP_ARCHIVE_FOLDER")
    if override:
        return _resolve_folder(ns, override, allow_archive_alias=False)
    root = ns.DefaultStore.GetRootFolder()
    for f in root.Folders:
        if f.Name.lower() in ARCHIVE_NAMES:
            return f
    raise OutlookError(
        "No Archive folder found at the top of the default mailbox. Set the environment variable "
        "OUTLOOK_MCP_ARCHIVE_FOLDER to the folder path to archive into."
    )


def _resolve_folder(ns, path: str | None, allow_archive_alias: bool = True):
    """Resolve a folder path.

    - '' / None / 'Inbox'           -> default Inbox
    - 'Inbox/Projects/2026'         -> path in the default mailbox; the first segment may be
                                       an alias (inbox, drafts, sent, deleted, junk, archive)
    - '//Online Archive - x/Archive' -> path in another store (also accepts '\\\\Store\\...')
    """
    if not path or not path.strip():
        return ns.GetDefaultFolder(OL_FOLDER_INBOX)
    raw = path.strip().replace("\\", "/")
    parts = [p for p in raw.split("/") if p.strip()]
    if not parts:
        return ns.GetDefaultFolder(OL_FOLDER_INBOX)

    if raw.startswith("//"):
        store_name = parts[0].lower()
        current = None
        for top in ns.Folders:
            if top.Name.lower() == store_name:
                current = top
                break
        if current is None:
            raise OutlookError(
                f"No store named '{parts[0]}'. Available: {', '.join(f.Name for f in ns.Folders)}"
            )
        rest = parts[1:]
    else:
        first = parts[0].lower()
        if first in FOLDER_ALIASES:
            current, rest = ns.GetDefaultFolder(FOLDER_ALIASES[first]), parts[1:]
        elif allow_archive_alias and first in ARCHIVE_NAMES:
            current, rest = _archive_folder(ns), parts[1:]
        else:
            current, rest = ns.DefaultStore.GetRootFolder(), parts

    for part in rest:
        current = _child(current, part)
    return current


def _folder_display(folder) -> str:
    """'Inbox/Sub' for the default mailbox, '//Store/Folder' for other stores."""
    try:
        segs = [s for s in folder.FolderPath.split("\\") if s]
    except Exception:  # noqa: BLE001
        return folder.Name
    try:
        default = folder.Session.DefaultStore.GetRootFolder().Name
    except Exception:  # noqa: BLE001
        default = None
    if segs and segs[0] == default:
        return "/".join(segs[1:]) or segs[0]
    return "//" + "/".join(segs)


# --------------------------------------------------------------------------
# Formatting helpers
# --------------------------------------------------------------------------


def _fmt_time(value) -> str | None:
    """Outlook returns local wall-clock times (pywin32 may label them UTC); print them as-is."""
    if value is None:
        return None
    try:
        if value.year >= 4500:  # Outlook's "no date" sentinel (4501-01-01)
            return None
        return value.strftime("%Y-%m-%d %H:%M")
    except Exception:  # noqa: BLE001
        return str(value)


def _naive(value) -> _dt.datetime:
    try:
        return _dt.datetime(value.year, value.month, value.day, value.hour, value.minute, value.second)
    except Exception:  # noqa: BLE001
        return _dt.datetime.min


def _is_mail_like(item) -> bool:
    try:
        cls = item.Class
    except Exception:  # noqa: BLE001
        return False
    return cls == OL_CLASS_MAIL or cls in OL_CLASS_MEETING


def _sender_email(item) -> str:
    try:
        smtp = item.PropertyAccessor.GetProperty(PR_SENDER_SMTP)
        if smtp:
            return smtp
    except Exception:  # noqa: BLE001
        pass
    try:
        return item.SenderEmailAddress or ""
    except Exception:  # noqa: BLE001
        return ""


def _importance(value) -> str:
    return {0: "low", 1: "normal", 2: "high"}.get(value, "normal")


def _from(name: str | None, email: str | None) -> str:
    name, email = name or "", email or ""
    if not email or email.startswith("/") or email.lower() == name.lower():
        return name or email
    return f"{name} <{email}>" if name else email


def _compact(entry_id, subject, name, email, received, unread, attachments, importance, categories,
             flag_status, folder_path) -> dict[str, Any]:
    """Summary shape shared by Table rows and COM items; false/normal values are left out to save tokens."""
    data: dict[str, Any] = {
        "id": _handle(entry_id),
        "subject": subject or "",
        "from": _from(name, email),
        "received": _fmt_time(received),
    }
    if unread:
        data["unread"] = True
    if attachments:
        data["attachments"] = True
    if importance in (0, 2):
        data["importance"] = _importance(importance)
    if categories:
        data["categories"] = categories
    flag = {1: "complete", 2: "flagged"}.get(flag_status)
    if flag:
        data["flag"] = flag
    if folder_path is not None:
        data["folder"] = folder_path
    return data


def _summary(item, folder_path: str | None = None) -> dict[str, Any]:
    try:
        has_attachments = item.Attachments.Count > 0
    except Exception:  # noqa: BLE001
        has_attachments = False
    try:
        flag_status = item.FlagStatus
    except Exception:  # noqa: BLE001
        flag_status = 0
    return _compact(
        item.EntryID, getattr(item, "Subject", ""), getattr(item, "SenderName", ""), _sender_email(item),
        getattr(item, "ReceivedTime", None), getattr(item, "UnRead", False), has_attachments,
        getattr(item, "Importance", 1), getattr(item, "Categories", ""), flag_status, folder_path,
    )


def _truncate(text: str, max_chars: int) -> str:
    if max_chars and len(text) > max_chars:
        return text[:max_chars] + f"\n\n[... truncated, {len(text) - max_chars} more characters]"
    return text


def _audit(action: str, **details: Any) -> None:
    record = {"time": _dt.datetime.now().isoformat(timespec="seconds"), "action": action, **details}
    try:
        os.makedirs(os.path.dirname(AUDIT_LOG_PATH), exist_ok=True)
        with open(AUDIT_LOG_PATH, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(record, ensure_ascii=False) + "\n")
    except OSError:
        pass


# --------------------------------------------------------------------------
# Reading
# --------------------------------------------------------------------------


def _iter_mail_folders(folder, recursive: bool):
    yield folder
    if not recursive:
        return
    for sub in folder.Folders:
        try:
            if sub.DefaultItemType != OL_MAIL_ITEM or sub.Name.lower() in _HIDDEN_FOLDER_NAMES:
                continue
        except Exception:  # noqa: BLE001
            continue
        yield from _iter_mail_folders(sub, True)


def list_folders(store: str = "default", max_depth: int = 10) -> list[dict[str, Any]]:
    ns = _ns()
    if store == "all":
        roots = []
        for s in ns.Stores:
            try:
                if s.ExchangeStoreType == 2:  # public folders: huge and slow
                    continue
                roots.append(s.GetRootFolder())
            except Exception:  # noqa: BLE001
                continue
    else:
        roots = [ns.DefaultStore.GetRootFolder()]

    results: list[dict[str, Any]] = []

    def walk(folder, depth: int) -> None:
        for sub in folder.Folders:
            try:
                if sub.DefaultItemType != OL_MAIL_ITEM or sub.Name.lower() in _HIDDEN_FOLDER_NAMES:
                    continue
                results.append({
                    "path": _folder_display(sub),
                    "items": sub.Items.Count,
                    "unread": sub.UnReadItemCount,
                })
                if depth < max_depth:
                    walk(sub, depth + 1)
            except Exception:  # noqa: BLE001
                continue

    for root in roots:
        walk(root, 1)
    return results


def _dasl_text(value: str) -> str:
    return value.replace("'", "''")


def _dasl_date(value: str, end_of_day: bool) -> str:
    """Local 'YYYY-MM-DD[ HH:MM]' -> UTC literal for a DASL comparison."""
    value = value.strip()
    try:
        if len(value) <= 10:
            local = _dt.datetime.strptime(value, "%Y-%m-%d")
            if end_of_day:
                local += _dt.timedelta(days=1)
        else:
            local = _dt.datetime.strptime(value.replace("T", " ")[:16], "%Y-%m-%d %H:%M")
    except ValueError as exc:
        raise OutlookError(f"Bad date '{value}'. Use YYYY-MM-DD or YYYY-MM-DD HH:MM.") from exc
    utc = local.astimezone().astimezone(_dt.timezone.utc)
    return utc.strftime("'%m/%d/%Y %I:%M %p'")


def _build_filter(
    query: str | None,
    search_body: bool,
    subject: str | None,
    sender: str | None,
    recipient: str | None,
    received_after: str | None,
    received_before: str | None,
    unread_only: bool,
    has_attachments: bool | None,
) -> str | None:
    """DASL filter. `query` is matched via the Windows Search index (ci_phrasematch on the body):
    sub-second, but it misses many Chinese words. Pass query=None for the exact two-stage search,
    which applies the other criteria here and checks the text in Python."""
    clauses = []
    if query:
        q = _dasl_text(query)
        parts = [f"\"urn:schemas:httpmail:subject\" LIKE '%{q}%'"]
        if search_body:
            parts.append(f"\"urn:schemas:httpmail:textdescription\" ci_phrasematch '{q}'")
        clauses.append("(" + " OR ".join(parts) + ")")
    if subject:
        clauses.append(f"\"urn:schemas:httpmail:subject\" LIKE '%{_dasl_text(subject)}%'")
    if sender:
        s = _dasl_text(sender)
        clauses.append(
            "("
            f"\"urn:schemas:httpmail:fromname\" LIKE '%{s}%' OR "
            f"\"urn:schemas:httpmail:fromemail\" LIKE '%{s}%' OR "
            f"\"{PR_SENDER_SMTP}\" LIKE '%{s}%'"
            ")"
        )
    if recipient:
        r = _dasl_text(recipient)
        clauses.append(
            f"(\"urn:schemas:httpmail:displayto\" LIKE '%{r}%' OR \"urn:schemas:httpmail:displaycc\" LIKE '%{r}%')"
        )
    if received_after:
        clauses.append(f"\"urn:schemas:httpmail:datereceived\" >= {_dasl_date(received_after, False)}")
    if received_before:
        clauses.append(f"\"urn:schemas:httpmail:datereceived\" < {_dasl_date(received_before, True)}")
    if unread_only:
        clauses.append("\"urn:schemas:httpmail:read\" = 0")
    if has_attachments is not None:
        clauses.append(f"\"urn:schemas:httpmail:hasattachment\" = {1 if has_attachments else 0}")
    if not clauses:
        return None
    return "@SQL=" + " AND ".join(clauses)


# Folder.GetTable() returns many rows' columns in one cross-process call, 20-100x faster than
# reading each property off each item. Column order matters: see _Row below.
_TABLE_COLUMNS = (
    "Subject", "SenderName", "ReceivedTime", "UnRead", "Importance", "Categories", "MessageClass",
    PR_SENDER_SMTP, "urn:schemas:httpmail:hasattachment", PR_FLAG_STATUS, PR_LONGTERM_ENTRYID, "EntryID",
)
_TABLE_BATCH = 500


def _is_mail_class(message_class: str | None) -> bool:
    c = message_class or ""
    return c == "IPM.Note" or c.startswith(("IPM.Note.", "IPM.Schedule.Meeting"))


class _Row:
    """One mail row from a folder Table, kept until it is turned into a summary."""

    __slots__ = ("values", "folder_path", "received")

    def __init__(self, values, folder_path: str):
        self.values = values
        self.folder_path = folder_path
        self.received = _naive(values[2])

    @property
    def subject(self) -> str:
        return self.values[0] or ""

    @property
    def entry_id(self) -> str:
        long_term = self.values[10]
        return bytes(long_term).hex().upper() if long_term is not None else self.values[11]

    def summary(self) -> dict[str, Any]:
        v = self.values
        return _compact(self.entry_id, v[0], v[1], v[7], v[2], v[3], v[8], v[4], v[5], v[9], self.folder_path)


def _table_rows(folder, dasl: str | None, limit: int, offset: int = 0) -> list[_Row]:
    """Mail rows of one folder, newest first, after an optional DASL filter."""
    table = folder.GetTable(dasl) if dasl else folder.GetTable()
    table.Columns.RemoveAll()
    for col in _TABLE_COLUMNS:
        table.Columns.Add(col)
    table.Sort("[ReceivedTime]", True)
    path = _folder_display(folder)
    out: list[_Row] = []
    skipped = 0
    while len(out) < limit and not table.EndOfTable:
        batch = table.GetArray(_TABLE_BATCH)
        if not batch:
            break
        for values in batch:
            if not _is_mail_class(values[6]):
                continue
            if skipped < offset:
                skipped += 1
                continue
            out.append(_Row(values, path))
            if len(out) >= limit:
                break
    return out


def list_emails(folder: str = "Inbox", limit: int = 25, offset: int = 0, unread_only: bool = False) -> dict[str, Any]:
    limit = max(1, min(limit, MAX_ITEMS_HARD_CAP))
    ns = _ns()
    f = _resolve_folder(ns, folder)
    dasl = "@SQL=\"urn:schemas:httpmail:read\" = 0" if unread_only else None
    rows = _table_rows(f, dasl, limit, max(0, offset))
    return {
        "folder": _folder_display(f),
        "total": f.Items.Count,
        "unread": f.UnReadItemCount,
        "emails": [r.summary() for r in rows],
    }


_CJK_RE = re.compile(r"[぀-ヿ㐀-䶿一-鿿가-힯豈-﫿]")
EXACT_SCAN_DEFAULT = 1500
EXACT_SCAN_MAX = 10000
EXACT_TIME_BUDGET_SECONDS = 90


def _search_folders(ns, folder: str, include_subfolders: bool) -> list:
    if folder.strip().lower() in ("all", "*"):
        root = ns.DefaultStore.GetRootFolder()
        return [f for sub in root.Folders
                if sub.DefaultItemType == OL_MAIL_ITEM and sub.Name.lower() not in _HIDDEN_FOLDER_NAMES
                for f in _iter_mail_folders(sub, True)]
    return list(_iter_mail_folders(_resolve_folder(ns, folder), include_subfolders))


def _gather(folders, dasl: str | None, per_folder: int, errors: list[str]) -> list[_Row]:
    rows: list[_Row] = []
    for f in folders:
        try:
            rows.extend(_table_rows(f, dasl, per_folder))
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{_folder_display(f)}: {exc}")
    rows.sort(key=lambda r: r.received, reverse=True)
    return rows


def _exact_scan(ns, candidates: list[_Row], needle: str, search_body: bool, limit: int) -> tuple[list[_Row], int, bool]:
    """Stage 2: check the text itself, newest first, stopping at `limit` hits or the time budget.

    The subject comes with the row; only rows whose subject doesn't match need the body,
    which costs one Outlook round trip each (~40 ms).
    """
    needle = needle.lower()
    hits: list[_Row] = []
    scanned = 0
    deadline = time.monotonic() + EXACT_TIME_BUDGET_SECONDS
    for row in candidates:
        if len(hits) >= limit:
            break
        if time.monotonic() > deadline:
            return hits, scanned, True
        scanned += 1
        if needle in row.subject.lower():
            hits.append(row)
            continue
        if not search_body:
            continue
        try:
            body = ns.GetItemFromID(row.entry_id).Body or ""
        except Exception:  # noqa: BLE001
            continue
        if needle in body.lower():
            hits.append(row)
    return hits, scanned, False


def search_emails(
    query: str | None = None,
    folder: str = "Inbox",
    include_subfolders: bool = False,
    search_body: bool = True,
    subject: str | None = None,
    sender: str | None = None,
    recipient: str | None = None,
    received_after: str | None = None,
    received_before: str | None = None,
    unread_only: bool = False,
    has_attachments: bool | None = None,
    match_mode: str = "auto",
    max_scan: int = EXACT_SCAN_DEFAULT,
    limit: int = 25,
) -> dict[str, Any]:
    if match_mode not in ("auto", "index", "exact"):
        raise OutlookError("match_mode must be 'auto', 'index' or 'exact'.")
    limit = max(1, min(limit, MAX_ITEMS_HARD_CAP))
    ns = _ns()
    folders = _search_folders(ns, folder, include_subfolders)
    errors: list[str] = []

    exact = bool(query) and (
        match_mode == "exact"
        or (match_mode == "auto" and (bool(_CJK_RE.search(query)) or not ns.DefaultStore.IsInstantSearchEnabled))
    )
    result: dict[str, Any] = {"folders_searched": len(folders), "match": "exact" if exact else "index"}

    if not exact:
        dasl = _build_filter(query, search_body, subject, sender, recipient,
                             received_after, received_before, unread_only, has_attachments)
        rows = _gather(folders, dasl, limit, errors)
        hits = rows[:limit]
        if len(rows) > limit:
            result["note"] = "More matches exist; narrow the search or raise limit."
    else:
        # Stage 1: every criterion except the text, via the fast Table filter.
        max_scan = max(1, min(max_scan, EXACT_SCAN_MAX))
        dasl = _build_filter(None, search_body, subject, sender, recipient,
                             received_after, received_before, unread_only, has_attachments)
        candidates = _gather(folders, dasl, max_scan, errors)
        candidates = candidates[:max_scan]
        # Stage 2: exact substring check in subject/body.
        hits, scanned, timed_out = _exact_scan(ns, candidates, query, search_body, limit)
        result["scanned"] = scanned
        covered_all = scanned == len(candidates) and len(candidates) < max_scan
        if len(hits) >= limit:
            result["note"] = "Limit reached; older matches may exist."
        elif not covered_all and candidates:
            oldest = _fmt_time(candidates[min(scanned, len(candidates)) - 1].received) if scanned else None
            reason = "time budget" if timed_out else f"max_scan={max_scan}"
            result["note"] = (
                f"Exact check stopped at {reason}; emails older than {oldest} were not checked. "
                "Narrow with received_after/received_before, sender or folder, or raise max_scan. "
                "Indexed matches from older emails are added below."
            )
            # Fall back to the index for the part the exact scan did not reach.
            seen = {r.entry_id for r in hits}
            idx = _build_filter(query, search_body, subject, sender, recipient,
                                received_after, received_before, unread_only, has_attachments)
            extra = [r for r in _gather(folders, idx, limit, errors) if r.entry_id not in seen]
            hits = sorted(hits + extra, key=lambda r: r.received, reverse=True)[:limit]

    result["returned"] = len(hits)
    result["emails"] = [r.summary() for r in hits]
    if errors:
        result["errors"] = errors[:10]
    return result


def _recipients(item) -> dict[str, list[str]]:
    """{'to': ['Name <smtp>', ...], 'cc': [...], 'bcc': [...]}, empty groups left out."""
    out: dict[str, list[str]] = {}
    kinds = {1: "to", 2: "cc", 3: "bcc"}
    try:
        for r in item.Recipients:
            try:
                email = r.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
            except Exception:  # noqa: BLE001
                email = getattr(r, "Address", "") or ""
            out.setdefault(kinds.get(r.Type, "other"), []).append(_from(r.Name, email))
    except Exception:  # noqa: BLE001
        pass
    return out


def _is_inline(att) -> bool:
    """Signature logos and pasted pictures: hidden, or an image referenced by Content-ID."""
    try:
        pa = att.PropertyAccessor
        try:
            if pa.GetProperty(PR_ATTACHMENT_HIDDEN):
                return True
        except Exception:  # noqa: BLE001
            pass
        try:
            cid = pa.GetProperty(PR_ATTACH_CONTENT_ID)
        except Exception:  # noqa: BLE001
            cid = ""
        return bool(cid) and os.path.splitext(att.FileName or "")[1].lower() in _IMAGE_EXTS
    except Exception:  # noqa: BLE001
        return False


def _attachments(item) -> list[dict[str, Any]]:
    out = []
    try:
        for att in item.Attachments:
            entry: dict[str, Any] = {"index": att.Index, "name": att.FileName, "kb": round(att.Size / 1024)}
            if _is_inline(att):
                entry["inline"] = True
            out.append(entry)
    except Exception:  # noqa: BLE001
        pass
    return out


def _tidy_body(text: str) -> str:
    """Outlook bodies are padded with '\\r\\n \\r\\n' runs; collapse them to save tokens."""
    text = text.replace("\r\n", "\n").replace("\xa0", " ")
    text = re.sub(r"[ \t]+\n", "\n", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def get_email(item_id: str, max_body_chars: int = 20000, include_html: bool = False) -> dict[str, Any]:
    """Reading through COM does not mark the message as read."""
    ns = _ns()
    item = _get_item(ns, item_id)
    data = _summary(item, _folder_display(item.Parent))
    data.update(_recipients(item))
    data["sent"] = _fmt_time(getattr(item, "SentOn", None))
    attachments = _attachments(item)
    if attachments:
        data["attachment_list"] = attachments
    data["body"] = _truncate(_tidy_body(getattr(item, "Body", "") or ""), max_body_chars)
    if include_html:
        data["html_body"] = _truncate(getattr(item, "HTMLBody", "") or "", max_body_chars)
    if item.Class == OL_CLASS_MAIL and not item.Sent:
        data["draft"] = True
    return data


def get_thread(item_id: str, limit: int = 50) -> dict[str, Any]:
    ns = _ns()
    item = _get_item(ns, item_id)
    conv = None
    try:
        conv = item.GetConversation()
    except Exception:  # noqa: BLE001
        pass
    if conv is None:
        raise OutlookError("Conversation view is not available for this item's store.")
    # Walk real items rather than conv.GetTable(): in cached mode the table hands out
    # short-term EntryIDs that GetItemFromID cannot open.
    found: list[tuple[_dt.datetime, dict[str, Any]]] = []

    def walk(items) -> None:
        for child in items:
            if len(found) >= limit:
                return
            try:
                if _is_mail_like(child):
                    found.append((_naive(child.ReceivedTime), _summary(child, _folder_display(child.Parent))))
                walk(conv.GetChildren(child))
            except Exception:  # noqa: BLE001
                continue

    walk(conv.GetRootItems())
    found.sort(key=lambda r: r[0])
    return {"conversation_topic": getattr(item, "ConversationTopic", ""), "messages": [r[1] for r in found]}


def _unique_path(directory: str, filename: str) -> str:
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", os.path.basename(filename)).strip() or "attachment"
    base, ext = os.path.splitext(safe)
    candidate = os.path.join(directory, safe)
    n = 1
    while os.path.exists(candidate):
        candidate = os.path.join(directory, f"{base} ({n}){ext}")
        n += 1
    return candidate


def save_attachments(
    item_id: str,
    attachment_indexes: list[int] | None = None,
    destination: str | None = None,
    subfolder: str | None = None,
    include_inline: bool = False,
) -> dict[str, Any]:
    """Write attachments to local disk. Never overwrites an existing file.

    No indexes -> every real attachment (inline signature images skipped unless include_inline).
    """
    try:
        directory, label = destinations.resolve(destination, subfolder)
    except destinations.DestinationError as exc:
        raise OutlookError(str(exc)) from exc
    ns = _ns()
    item = _get_item(ns, item_id)
    wanted = set(attachment_indexes or [])
    saved, skipped = [], []
    for att in item.Attachments:
        if wanted and att.Index not in wanted:
            continue
        if not wanted and not include_inline and _is_inline(att):
            skipped.append(att.FileName)
            continue
        dest = _unique_path(directory, att.FileName)
        att.SaveAsFile(dest)
        saved.append({"name": os.path.basename(dest), "kb": round(att.Size / 1024)})
    if wanted - {a.Index for a in item.Attachments}:
        raise OutlookError(f"No attachment with index {sorted(wanted)}; see attachment_list from get_email.")
    if saved:
        _audit("save_attachments", id=item.EntryID, subject=item.Subject, folder=directory,
               files=[s["name"] for s in saved])
    result: dict[str, Any] = {"destination": label, "folder": directory, "saved": saved}
    if skipped:
        result["skipped_inline_images"] = skipped
    return result


def list_save_destinations() -> dict[str, Any]:
    return {"default": destinations.default_dir(), "named": destinations.known_destinations()}


def list_calendar_events(days_ahead: int = 7, days_back: int = 0) -> list[dict[str, Any]]:
    ns = _ns()
    items = ns.GetDefaultFolder(OL_FOLDER_CALENDAR).Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")
    now = _dt.datetime.now()
    start = (now - _dt.timedelta(days=max(0, days_back))).replace(hour=0, minute=0)
    end = (now + _dt.timedelta(days=max(0, days_ahead))).replace(hour=23, minute=59)
    restricted = items.Restrict(
        f"[Start] >= '{start:%m/%d/%Y %I:%M %p}' AND [Start] <= '{end:%m/%d/%Y %I:%M %p}'"
    )
    results = []
    for item in restricted:
        if len(results) >= MAX_ITEMS_HARD_CAP:
            break
        try:
            results.append({
                "subject": item.Subject,
                "start": _fmt_time(item.Start),
                "end": _fmt_time(item.End),
                "location": item.Location,
                "organizer": item.Organizer,
                "is_recurring": bool(item.IsRecurring),
                "all_day": bool(item.AllDayEvent),
            })
        except Exception:  # noqa: BLE001
            continue
    return results


def list_recent_actions(limit: int = 20) -> list[dict[str, Any]]:
    try:
        with open(AUDIT_LOG_PATH, encoding="utf-8") as fh:
            lines = fh.readlines()
    except OSError:
        return []
    out = []
    for line in lines[-max(1, limit):]:
        try:
            record = json.loads(line)
        except ValueError:
            continue
        # The log keeps full EntryIDs; hand back short ids that work with the other tools.
        record.pop("old_id", None)
        if "new_id" in record:
            record["id"] = record.pop("new_id")
        if record.get("id"):
            record["id"] = ("(moved again later; see newer entries)" if ids.is_retired(record["id"])
                            else _handle(record["id"]))
        out.append(record)
    return out


# --------------------------------------------------------------------------
# Moving and archiving (guarded)
# --------------------------------------------------------------------------


def _assert_move_target_allowed(ns, target) -> None:
    if target.DefaultItemType != OL_MAIL_ITEM:
        raise OutlookError(f"'{_folder_display(target)}' is not a mail folder.")
    forbidden = []
    try:
        for kind in _FORBIDDEN_TARGET_DEFAULTS:
            try:
                forbidden.append(target.Store.GetDefaultFolder(kind))
            except Exception:  # noqa: BLE001
                continue
    except Exception:  # noqa: BLE001
        pass
    current = target
    while current is not None and getattr(current, "Class", None) == OL_CLASS_FOLDER:
        if current.Name.strip().lower() in _FORBIDDEN_TARGET_NAMES or any(_same(ns, current, f) for f in forbidden):
            raise OutlookError(
                f"Refused: moving mail into '{_folder_display(target)}' is not allowed by this server "
                "(Deleted Items, Junk, Outbox and Sync Issues are off-limits so nothing can be purged or sent)."
            )
        try:
            current = current.Parent
        except Exception:  # noqa: BLE001
            break


def _guarded_move(ns, item, target):
    """The single place in this package that moves an item."""
    _assert_move_target_allowed(ns, target)
    return item.Move(target)


def move_emails(item_ids: list[str], target_folder: str, action: str = "move") -> dict[str, Any]:
    if not item_ids:
        raise OutlookError("No item ids given.")
    if len(item_ids) > MAX_BATCH:
        raise OutlookError(f"At most {MAX_BATCH} items per call.")
    ns = _ns()
    target = _archive_folder(ns) if target_folder == "__archive__" else _resolve_folder(ns, target_folder)
    _assert_move_target_allowed(ns, target)
    target_path = _folder_display(target)
    moved, failed = [], []
    for item_id in item_ids:
        try:
            item = _get_item(ns, item_id)
            source_path = _folder_display(item.Parent)
            if _same(ns, item.Parent, target):
                moved.append({"id": _handle(item.EntryID), "subject": item.Subject, "note": "already there"})
                continue
            subject, old_entry = item.Subject, item.EntryID
            new_entry = _guarded_move(ns, item, target).EntryID
            # Keep the same short id pointing at the moved email.
            handle = ids.rebind(item_id if ids.HANDLE_RE.match(item_id.strip()) else old_entry, new_entry)
            moved.append({"id": handle, "subject": subject, "from": source_path})
            _audit(action, subject=subject, from_folder=source_path, to_folder=target_path,
                   old_id=old_entry, new_id=new_entry)
        except Exception as exc:  # noqa: BLE001
            failed.append({"id": item_id, "error": str(exc)})
    result: dict[str, Any] = {"target": target_path, "moved": moved}
    if failed:
        result["failed"] = failed
    return result


def create_folder(parent_folder: str, name: str) -> dict[str, Any]:
    name = name.strip()
    if not name or "/" in name or "\\" in name:
        raise OutlookError("Folder name must be non-empty and contain no slashes.")
    ns = _ns()
    parent = _resolve_folder(ns, parent_folder)
    _assert_move_target_allowed(ns, parent)
    for f in parent.Folders:
        if f.Name.lower() == name.lower():
            return {"path": _folder_display(f), "note": "already exists"}
    new = parent.Folders.Add(name)
    path = _folder_display(new)
    _audit("create_folder", path=path)
    return {"path": path}


# --------------------------------------------------------------------------
# Status: read/unread, flags, categories
# --------------------------------------------------------------------------


def _split_categories(value: str) -> list[str]:
    return [c.strip() for c in re.split(r"[,;]", value or "") if c.strip()]


def set_email_status(
    item_ids: list[str],
    read: bool | None = None,
    flag: str | None = None,
    add_categories: list[str] | None = None,
    remove_categories: list[str] | None = None,
) -> dict[str, Any]:
    if not item_ids:
        raise OutlookError("No item ids given.")
    if len(item_ids) > MAX_BATCH:
        raise OutlookError(f"At most {MAX_BATCH} items per call.")
    if flag not in (None, "flagged", "complete", "clear"):
        raise OutlookError("flag must be 'flagged', 'complete' or 'clear'.")
    ns = _ns()
    done, failed = [], []
    for item_id in item_ids:
        try:
            item = _get_item(ns, item_id)
            if read is not None:
                item.UnRead = not read
            if flag == "flagged":
                item.MarkAsTask(OL_MARK_NO_DATE)
            elif flag == "complete":
                item.MarkAsTask(OL_MARK_COMPLETE)
            elif flag == "clear":
                item.ClearTaskFlag()
            if add_categories or remove_categories:
                cats = _split_categories(item.Categories)
                for c in add_categories or []:
                    if c not in cats:
                        cats.append(c)
                drop = {c.lower() for c in remove_categories or []}
                item.Categories = ", ".join(c for c in cats if c.lower() not in drop)
            item.Save()
            done.append(item_id)
            _audit("set_status", subject=item.Subject, id=item.EntryID, read=read, flag=flag,
                   add_categories=add_categories, remove_categories=remove_categories)
        except Exception as exc:  # noqa: BLE001
            failed.append({"id": item_id, "error": str(exc)})
    result: dict[str, Any] = {"updated": len(done)}
    if failed:
        result["failed"] = failed
    return result


# --------------------------------------------------------------------------
# Drafts (saved, never sent)
# --------------------------------------------------------------------------

_MARKER_START = "<!--mcp-text-start-->"
_MARKER_END = "<!--mcp-text-end-->"


def _text_to_html(text: str) -> str:
    return _html.escape(text).replace("\r\n", "\n").replace("\n", "<br>\n")


def _join(addresses: list[str] | str | None) -> str | None:
    if addresses is None:
        return None
    if isinstance(addresses, str):
        return addresses
    return "; ".join(a.strip() for a in addresses if a and a.strip())


def _apply_headers(mail, to=None, cc=None, bcc=None, subject=None) -> None:
    if to is not None:
        mail.To = _join(to)
    if cc is not None:
        mail.CC = _join(cc)
    if bcc is not None:
        mail.BCC = _join(bcc)
    if subject is not None:
        mail.Subject = subject
    try:
        mail.Recipients.ResolveAll()
    except Exception:  # noqa: BLE001
        pass


def _add_attachments(mail, paths: list[str] | None) -> None:
    for p in paths or []:
        full = os.path.abspath(os.path.expanduser(p))
        if not os.path.isfile(full):
            raise OutlookError(f"Attachment not found: {full}")
        mail.Attachments.Add(full)


def _insert_text(mail, body: str, body_format: str) -> None:
    """Put new text above whatever the item already contains (quoted thread, forward header)."""
    if mail.BodyFormat == OL_FORMAT_HTML:
        fragment = body if body_format == "html" else _text_to_html(body)
        block = f"{_MARKER_START}<div>{fragment}</div><br>{_MARKER_END}"
        current = mail.HTMLBody or ""
        match = re.search(r"<body[^>]*>", current, re.IGNORECASE)
        mail.HTMLBody = (current[: match.end()] + block + current[match.end():]) if match else block + current
    else:
        mail.Body = body + "\r\n\r\n" + (mail.Body or "")


def _draft_result(mail, action: str, **extra: Any) -> dict[str, Any]:
    result = {
        "id": _handle(mail.EntryID),
        "subject": mail.Subject,
        "to": mail.To,
        "cc": mail.CC,
        "folder": _folder_display(mail.Parent),
        "status": "Saved as draft. NOT sent - review and send it yourself in Outlook.",
        **extra,
    }
    _audit(action, id=mail.EntryID, subject=mail.Subject, to=mail.To, cc=mail.CC)
    return result


def create_draft(
    to: list[str] | None = None,
    subject: str = "",
    body: str = "",
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    body_format: str = "text",
    attachments: list[str] | None = None,
) -> dict[str, Any]:
    mail = _app().CreateItem(OL_MAIL_ITEM)
    _apply_headers(mail, to, cc, bcc, subject)
    fragment = body if body_format == "html" else _text_to_html(body)
    mail.HTMLBody = f"<html><body>{_MARKER_START}<div>{fragment}</div><br>{_MARKER_END}</body></html>"
    _add_attachments(mail, attachments)
    mail.Save()
    return _draft_result(mail, "create_draft")


def create_reply_draft(
    item_id: str,
    body: str,
    reply_all: bool = False,
    body_format: str = "text",
    add_cc: list[str] | None = None,
    attachments: list[str] | None = None,
) -> dict[str, Any]:
    ns = _ns()
    original = _get_item(ns, item_id)
    reply = original.ReplyAll() if reply_all else original.Reply()
    if add_cc:
        existing = reply.CC or ""
        reply.CC = "; ".join(x for x in [existing, _join(add_cc)] if x)
        reply.Recipients.ResolveAll()
    _insert_text(reply, body, body_format)
    _add_attachments(reply, attachments)
    reply.Save()
    return _draft_result(reply, "create_reply_draft", in_reply_to=original.Subject)


def create_forward_draft(
    item_id: str,
    to: list[str],
    body: str = "",
    cc: list[str] | None = None,
    body_format: str = "text",
    attachments: list[str] | None = None,
) -> dict[str, Any]:
    ns = _ns()
    original = _get_item(ns, item_id)
    fwd = original.Forward()
    _apply_headers(fwd, to, cc)
    if body:
        _insert_text(fwd, body, body_format)
    _add_attachments(fwd, attachments)
    fwd.Save()
    return _draft_result(fwd, "create_forward_draft", forwarding=original.Subject)


def update_draft(
    item_id: str,
    to: list[str] | None = None,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    subject: str | None = None,
    body: str | None = None,
    body_format: str = "text",
    replace_entire_body: bool = False,
    add_attachments: list[str] | None = None,
) -> dict[str, Any]:
    ns = _ns()
    mail = _get_item(ns, item_id)
    if mail.Class != OL_CLASS_MAIL or mail.Sent:
        raise OutlookError("Refused: this item is not an unsent draft.")
    try:
        drafts = mail.Parent.Store.GetDefaultFolder(OL_FOLDER_DRAFTS)
    except Exception:  # noqa: BLE001
        drafts = ns.GetDefaultFolder(OL_FOLDER_DRAFTS)
    if not _same(ns, mail.Parent, drafts):
        raise OutlookError("Refused: only items in the Drafts folder can be edited.")

    _apply_headers(mail, to, cc, bcc, subject)
    note = None
    if body is not None:
        if mail.BodyFormat == OL_FORMAT_HTML:
            fragment = body if body_format == "html" else _text_to_html(body)
            current = mail.HTMLBody or ""
            start, end = current.find(_MARKER_START), current.find(_MARKER_END)
            if not replace_entire_body and start != -1 and end > start:
                block = f"{_MARKER_START}<div>{fragment}</div><br>"
                mail.HTMLBody = current[:start] + block + current[end:]
                note = "Replaced the text written by this server; quoted thread kept."
            elif replace_entire_body:
                mail.HTMLBody = f"<html><body>{_MARKER_START}{fragment}{_MARKER_END}</body></html>"
                note = "Replaced the entire body."
            else:
                _insert_text(mail, body, body_format)
                note = ("Could not find the earlier inserted text (Outlook may have rewritten it), so the new "
                        "text was added on top. Set replace_entire_body=true to overwrite everything instead.")
        else:
            if replace_entire_body:
                mail.Body = body
            else:
                _insert_text(mail, body, body_format)
                note = "Plain-text draft: new text added on top. Set replace_entire_body=true to overwrite."
    _add_attachments(mail, add_attachments)
    mail.Save()
    extra = {"note": note} if note else {}
    return _draft_result(mail, "update_draft", **extra)
