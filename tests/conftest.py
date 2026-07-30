"""Expose pure helpers from the entry-point script without executing it.

Importing `run_amzn_account_health` outright would demand real credentials and then
start the scheduler, so the pieces under test are lifted out of the source via AST.
"""
from __future__ import annotations

import ast
from pathlib import Path

import pytest

SCRIPT = Path(__file__).resolve().parent.parent / "run_amzn_account_health.py"
WANTED = {"COUNT_UNAVAILABLE", "_premium_shipping_values"}


def _load_isolated() -> dict:
    """Compile only the wanted top-level names out of the entry-point script.

    Returns:
        Namespace holding the extracted constants and functions.
    """
    tree = ast.parse(SCRIPT.read_text(encoding="utf-8"))
    kept: list[ast.stmt] = [
        node
        for node in tree.body
        if (isinstance(node, ast.FunctionDef) and node.name in WANTED)
        or (
            isinstance(node, ast.Assign)
            and any(isinstance(t, ast.Name) and t.id in WANTED for t in node.targets)
        )
    ]
    namespace: dict = {}
    exec(compile(ast.Module(body=kept, type_ignores=[]), str(SCRIPT), "exec"), namespace)
    return namespace


@pytest.fixture(scope="session")
def premium_shipping_values():
    """The `_premium_shipping_values` function, loaded without side effects."""
    return _load_isolated()["_premium_shipping_values"]


@pytest.fixture(scope="session")
def count_unavailable() -> str:
    """The placeholder written into cells Amazon no longer supplies."""
    return _load_isolated()["COUNT_UNAVAILABLE"]
