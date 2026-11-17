"""
Offline tests for draft_html (no Outlook needed).

Run:  .venv\\Scripts\\python.exe tests\\test_draft_html.py
"""

from __future__ import annotations

import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1] / "src"))

from outlook_local_mcp import draft_html as dh  # noqa: E402

S, E = "<!--mcp-text-start-->", "<!--mcp-text-end-->"
REPLY = (
    "<html><head><style>p {margin:0}</style></head><body lang=EN-US>"
    f"{S}<div>Dear Ms. Wang,<br>\nPlease find the <b>draft</b> attached.&nbsp; Thanks.</div><br>{E}"
    "<div><p class=MsoNormal>From: Someone<o:p></o:p></p>"
    "<p class=MsoNormal>Please find the attached schedule.</p></div></body></html>"
)


def check(name: str, cond: bool) -> None:
    print(("ok   " if cond else "FAIL ") + name)
    if not cond:
        raise SystemExit(1)


# Replace across a tag (<b>) and an entity, inside the server's own block only.
out, n = dh.replace_text(REPLY, "Please find the draft attached.", "I attach the revised draft.",
                         marker_start=S, marker_end=E)
check("replace across tag", n == 1 and "I attach the revised draft." in dh.visible_text(out))
check("quoted thread untouched", "Please find the attached schedule." in out)
check("markup still balanced", out.count("<b>") == out.count("</b>"))

# Deleting = replacing with ''.
out, n = dh.replace_text(REPLY, "Thanks.", "", marker_start=S, marker_end=E)
check("delete text", "Thanks" not in dh.visible_text(out) and n == 1)

# Whitespace in the search matches a <br> + newline line break.
out, n = dh.replace_text(REPLY, "Wang, Please", "Wang, please", match_case=True, marker_start=S, marker_end=E)
check("match across line break", n == 1)

# Scope: the quoted thread is only reachable with whole_body.
try:
    dh.replace_text(REPLY, "attached schedule", "x", marker_start=S, marker_end=E)
    check("scoped search refuses quoted text", False)
except dh.EditError as exc:
    check("scoped search refuses quoted text", "whole_body" in str(exc))
out, n = dh.replace_text(REPLY, "attached schedule", "enclosed schedule", marker_start=S, marker_end=E,
                         whole_body=True)
check("whole_body reaches quoted text", n == 1 and "enclosed schedule" in out)

# Formatting wraps each touched text node in its own span.
out, n = dh.format_text(REPLY, "the draft attached", bold=True, color="#C00000", marker_start=S, marker_end=E)
check("format across tag", n == 1 and out.count('<span style="font-weight:bold;color:#C00000">') == 3)
check("format keeps text", "the draft attached" in dh.visible_text(out).replace("  ", " "))
out, _ = dh.format_text(REPLY, "Thanks.", highlight="yellow", font_size=12, marker_start=S, marker_end=E)
check("highlight + size", 'background:yellow;font-size:12pt">Thanks.</span>' in out)

# Occurrence selection and errors.
two = "<html><body><p>alpha beta alpha</p></body></html>"
out, n = dh.replace_text(two, "alpha", "gamma", occurrence=2)
check("second occurrence only", dh.visible_text(out) == "alpha beta gamma")
out, n = dh.replace_text(two, "ALPHA", "x")
check("case-insensitive by default, all occurrences", n == 2)
try:
    dh.replace_text(two, "alpha bet x", "y")
except dh.EditError as exc:
    check("not-found hint", "matches up to 'alpha bet" in str(exc))
try:
    dh.format_text(two, "alpha", color="red;position:fixed")
except dh.EditError:
    check("style injection refused", True)
check("escaping", "&lt;script&gt;" in dh.replace_text(two, "beta", "<script>")[0])
print("all draft_html tests passed")
