"""
draft_html.py

Find, replace, delete and format text inside a draft's HTML body, in pure
Python. (Outlook's Word editor only writes edits back when the draft window
is on screen, so it can't be used from a background server.)

The HTML is split into tags and text nodes. Text is matched on what a reader
sees: entities decoded, any run of whitespace (including a paragraph break
between two tags) matching any whitespace in the search text, and tags between
words ignored. An edit is applied per text node, so it never breaks the markup:
a replacement goes into the first text node the match touches; formatting wraps
each touched piece of text in its own <span style="...">.
"""

from __future__ import annotations

import html as _html
import re
from dataclasses import dataclass

# Tags whose content is not visible text.
_OPAQUE = r"<!--.*?-->|<head\b.*?</head\s*>|<style\b.*?</style\s*>|<script\b.*?</script\s*>|<xml\b.*?</xml\s*>"
_TOKEN_RE = re.compile(rf"({_OPAQUE}|<[^>]*>)", re.IGNORECASE | re.DOTALL)
# Tags that end a line for the reader: text on either side of them is not joined.
_BREAK_RE = re.compile(r"^</?(p|div|br|li|ul|ol|tr|td|th|table|h[1-6]|blockquote)\b", re.IGNORECASE)
_WS = " \t\r\n\xa0　"

HIGHLIGHTS = {"yellow", "green", "cyan", "magenta", "red", "blue", "gray", "lightgray", "orange", "pink"}


class EditError(ValueError):
    pass


@dataclass
class _Node:
    source: str
    text: str | None  # decoded text for a text node, None for a tag


def _parse(html: str) -> list[_Node]:
    nodes = []
    for part in _TOKEN_RE.split(html):
        if not part:
            continue
        if part.startswith("<"):
            nodes.append(_Node(part, None))
        else:
            nodes.append(_Node(part, _html.unescape(part)))
    return nodes


def _haystack(nodes: list[_Node], lo: int, hi: int):
    """Visible text of nodes[lo:hi] with whitespace collapsed, and for each character
    the (node index, start, end) of the decoded text it stands for (None for a line break)."""
    chars: list[str] = []
    spans: list[tuple[int, int, int] | None] = []
    for idx in range(lo, hi):
        node = nodes[idx]
        if node.text is None:
            if _BREAK_RE.match(node.source) and chars and chars[-1] != " ":
                chars.append(" ")
                spans.append(None)
            continue
        t = node.text
        i = 0
        while i < len(t):
            if t[i] in _WS:
                j = i
                while j < len(t) and t[j] in _WS:
                    j += 1
                if chars and chars[-1] == " ":
                    # Merge into the space already there; widen it if it maps to real text.
                    if spans[-1] is not None and spans[-1][0] == idx:
                        spans[-1] = (idx, spans[-1][1], j)
                    elif spans[-1] is None:
                        spans[-1] = (idx, i, j)
                else:
                    chars.append(" ")
                    spans.append((idx, i, j))
                i = j
            else:
                chars.append(t[i])
                spans.append((idx, i, i + 1))
                i += 1
    return "".join(chars), spans


def _normalize_needle(needle: str) -> str:
    return re.sub(rf"[{_WS}]+", " ", needle).strip()


def _scope(nodes: list[_Node], start_marker: str | None, end_marker: str | None) -> tuple[int, int]:
    if start_marker and end_marker:
        starts = [i for i, n in enumerate(nodes) if n.text is None and n.source == start_marker]
        ends = [i for i, n in enumerate(nodes) if n.text is None and n.source == end_marker]
        if starts and ends and ends[0] > starts[0]:
            return starts[0] + 1, ends[0]
    body = next((i for i, n in enumerate(nodes) if n.text is None and re.match(r"<body\b", n.source, re.I)), -1)
    return body + 1, len(nodes)


def _matches(hay: str, needle: str, match_case: bool) -> list[tuple[int, int]]:
    if not needle:
        raise EditError("The text to find is empty.")
    h, n = (hay, needle) if match_case else (hay.lower(), needle.lower())
    out, pos = [], 0
    while True:
        i = h.find(n, pos)
        if i < 0:
            return out
        out.append((i, i + len(n)))
        pos = i + len(n)


def _pieces(spans, i: int, j: int) -> dict[int, tuple[int, int]]:
    """Per text node, the decoded range a match [i, j) covers."""
    per: dict[int, tuple[int, int]] = {}
    for s in spans[i:j]:
        if s is None:
            continue
        node, a, b = s
        lo, hi = per.get(node, (a, b))
        per[node] = (min(lo, a), max(hi, b))
    return per


def _select(found: list, occurrence: str | int) -> list:
    if occurrence == "all":
        return found
    if occurrence == "first":
        return found[:1]
    if occurrence == "last":
        return found[-1:]
    try:
        k = int(occurrence)
    except (TypeError, ValueError) as exc:
        raise EditError("occurrence must be 'all', 'first', 'last' or a number (1 = first).") from exc
    if not 1 <= k <= len(found):
        raise EditError(f"Only {len(found)} occurrence(s) found; occurrence={k} is out of range.")
    return [found[k - 1]]


def _closest_hint(hay: str, needle: str) -> str:
    """Longest prefix of the needle that does occur, to show where the text stops matching."""
    low, n = hay.lower(), needle.lower()
    best = 0
    for k in range(len(n), 0, -1):
        if n[:k] in low:
            best = k
            break
    if best < min(3, len(n)):
        return ""
    at = low.find(n[:best])
    return f" The draft matches up to '{needle[:best]}' and then reads '{hay[at + best: at + best + 40]}'."


def _rebuild(nodes: list[_Node], edits: dict[int, list[tuple[int, int, str, str]]]) -> str:
    """edits[node] = [(start, end, kind, payload)]: kind 'replace' (payload = new text) or
    'wrap' (payload = style)."""
    out = []
    for idx, node in enumerate(nodes):
        if idx not in edits:
            out.append(node.source)
            continue
        text, pos, parts = node.text, 0, []
        for a, b, kind, payload in sorted(edits[idx]):
            if a < pos:
                raise EditError("Two matches overlap; make the text to find more specific.")
            parts.append(_html.escape(text[pos:a], quote=False))
            if kind == "replace":
                parts.append(_text_html(payload))
            else:
                parts.append(f'<span style="{payload}">{_html.escape(text[a:b], quote=False)}</span>')
            pos = b
        parts.append(_html.escape(text[pos:], quote=False))
        out.append("".join(parts))
    return "".join(out)


def _text_html(text: str) -> str:
    return _html.escape(text, quote=False).replace("\r\n", "\n").replace("\n", "<br>")


def _find_all(html, find, match_case, marker_start, marker_end, whole_body):
    nodes = _parse(html)
    lo, hi = _scope(nodes, None if whole_body else marker_start, None if whole_body else marker_end)
    hay, spans = _haystack(nodes, lo, hi)
    needle = _normalize_needle(find)
    found = _matches(hay, needle, match_case)
    return nodes, hay, spans, needle, found


def _not_found(hay: str, needle: str, scoped: bool) -> EditError:
    where = ("in the text this server wrote; pass whole_body=true to search the whole draft, "
             "including any quoted thread" if scoped else "in the draft")
    return EditError(f"'{needle}' not found {where}.{_closest_hint(hay, needle)}")


def replace_text(html: str, find: str, replacement: str, occurrence: str | int = "all",
                 match_case: bool = False, marker_start: str | None = None, marker_end: str | None = None,
                 whole_body: bool = False) -> tuple[str, int]:
    """Replace (or, with replacement='', delete) visible text. Returns (new html, count)."""
    nodes, hay, spans, needle, found = _find_all(html, find, match_case, marker_start, marker_end, whole_body)
    if not found:
        raise _not_found(hay, needle, bool(marker_start) and not whole_body and marker_start in html)
    chosen = _select(found, occurrence)
    edits: dict[int, list] = {}
    for i, j in chosen:
        per = _pieces(spans, i, j)
        for k, (node, (a, b)) in enumerate(sorted(per.items())):
            edits.setdefault(node, []).append((a, b, "replace", replacement if k == 0 else ""))
    return _rebuild(nodes, edits), len(chosen)


def _style(bold, italic, underline, strikethrough, color, highlight, font_size, font_name) -> str:
    rules = []
    if bold is not None:
        rules.append(f"font-weight:{'bold' if bold else 'normal'}")
    if italic is not None:
        rules.append(f"font-style:{'italic' if italic else 'normal'}")
    decorations = [d for d, on in (("underline", underline), ("line-through", strikethrough)) if on]
    if underline is False and strikethrough is False:
        decorations = ["none"]
    if decorations:
        rules.append("text-decoration:" + " ".join(decorations))
    for name, value in (("color", color), ("background", highlight)):
        if value:
            if not re.fullmatch(r"#[0-9a-fA-F]{3}([0-9a-fA-F]{3})?|[a-zA-Z]{3,20}", value):
                raise EditError(f"{name} must be a colour name or #RRGGBB.")
            rules.append(f"{name}:{value}")
    if font_size is not None:
        if not 6 <= float(font_size) <= 72:
            raise EditError("font_size must be between 6 and 72 (points).")
        rules.append(f"font-size:{float(font_size):g}pt")
    if font_name:
        if not re.fullmatch(r"[\w \-一-鿿]{1,40}", font_name):
            raise EditError("font_name must be a plain font name.")
        rules.append(f"font-family:'{font_name}'")
    if not rules:
        raise EditError("No formatting given: set bold, italic, underline, strikethrough, color, "
                        "highlight, font_size or font_name.")
    return ";".join(rules)


def format_text(html: str, find: str, occurrence: str | int = "all", match_case: bool = False,
                bold: bool | None = None, italic: bool | None = None, underline: bool | None = None,
                strikethrough: bool | None = None, color: str | None = None, highlight: str | None = None,
                font_size: float | None = None, font_name: str | None = None,
                marker_start: str | None = None, marker_end: str | None = None,
                whole_body: bool = False) -> tuple[str, int]:
    """Apply character formatting to visible text. Returns (new html, count)."""
    style = _style(bold, italic, underline, strikethrough, color, highlight, font_size, font_name)
    nodes, hay, spans, needle, found = _find_all(html, find, match_case, marker_start, marker_end, whole_body)
    if not found:
        raise _not_found(hay, needle, bool(marker_start) and not whole_body and marker_start in html)
    chosen = _select(found, occurrence)
    edits: dict[int, list] = {}
    for i, j in chosen:
        for node, (a, b) in _pieces(spans, i, j).items():
            text = nodes[node].text
            while a < b and text[a] in _WS:  # don't wrap leading/trailing whitespace
                a += 1
            while b > a and text[b - 1] in _WS:
                b -= 1
            if a < b:
                edits.setdefault(node, []).append((a, b, "wrap", style))
    return _rebuild(nodes, edits), len(chosen)


def visible_text(html: str) -> str:
    nodes = _parse(html)
    lo, hi = _scope(nodes, None, None)
    return _haystack(nodes, lo, hi)[0].strip()
