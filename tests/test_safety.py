"""
Static safety audit: parse every module of the package and fail if any
code could send mail, delete items, or move mail outside the guard.

Run:  .venv\\Scripts\\python.exe tests\\test_safety.py
"""

from __future__ import annotations

import ast
import pathlib
import sys

PACKAGE = pathlib.Path(__file__).resolve().parents[1] / "src" / "outlook_local_mcp"

# COM members that can send, delete, purge, or open UI/commands. None may be referenced at all.
FORBIDDEN = {
    "Send", "SendAndReceive", "Delete", "PermanentlyDelete", "Purge", "Empty",
    "MoveTo", "CopyTo", "Display", "ShowDialog", "Execute", "ExecuteMso",
    "ClearConversation", "Remove", "SaveAs", "Close",
    "Invoke", "InvokeTypes", "_oleobj_", "Dynamic", "DeleteAfterSubmit",
    "DeferredDeliveryTime", "AutoForwarded",
}
# Allowed, but only inside these functions.
RESTRICTED = {
    "Move": {"_guarded_move"},
    "Reply": {"create_reply_draft"},
    "ReplyAll": {"create_reply_draft"},
    "Forward": {"create_forward_draft"},
    # Table.Columns.RemoveAll only clears the column set of a read-only table view.
    "RemoveAll": {"_table_rows"},
}
DYNAMIC_ATTR_FUNCS = {"getattr", "setattr", "delattr", "hasattr"}


def audit_file(path: pathlib.Path) -> list[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    problems: list[str] = []
    parents: dict[ast.AST, ast.AST] = {}
    for node in ast.walk(tree):
        for child in ast.iter_child_nodes(node):
            parents[child] = node

    def enclosing_function(node: ast.AST) -> str | None:
        while node in parents:
            node = parents[node]
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                return node.name
        return None

    def check_name(name: str, node: ast.AST) -> None:
        where = f"{path.name}:{node.lineno}"
        if name in FORBIDDEN:
            problems.append(f"{where}: forbidden member '{name}'")
        elif name in RESTRICTED and enclosing_function(node) not in RESTRICTED[name]:
            problems.append(f"{where}: '{name}' used outside {sorted(RESTRICTED[name])}")

    for node in ast.walk(tree):
        if isinstance(node, ast.Attribute):
            check_name(node.attr, node)
        elif isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in DYNAMIC_ATTR_FUNCS:
            if len(node.args) < 2 or not isinstance(node.args[1], ast.Constant):
                problems.append(f"{path.name}:{node.lineno}: {node.func.id}() with a non-literal name")
            else:
                check_name(str(node.args[1].value), node)
        elif isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in {"eval", "exec", "__import__"}:
            problems.append(f"{path.name}:{node.lineno}: {node.func.id}() is not allowed")
    return problems


def main() -> int:
    files = sorted(PACKAGE.glob("*.py"))
    assert files, f"no sources under {PACKAGE}"
    problems = [p for f in files for p in audit_file(f)]
    moves = sum(
        1 for f in files for n in ast.walk(ast.parse(f.read_text(encoding="utf-8")))
        if isinstance(n, ast.Attribute) and n.attr == "Move"
    )
    if moves != 1:
        problems.append(f"expected exactly one '.Move' reference, found {moves}")
    if problems:
        print("SAFETY AUDIT FAILED:")
        for p in problems:
            print("  " + p)
        return 1
    print(f"Safety audit passed ({len(files)} files): no send/delete members, one guarded Move.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
