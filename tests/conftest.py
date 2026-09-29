"""Expose pure helpers from the entry-point script without executing it.

Importing `run_amzn_account_health` outright would demand real credentials and then
start the scheduler, so the pieces under test are lifted out of the source via AST.
"""
from __future__ import annotations

import ast
import logging
import os
import signal
import threading
from pathlib import Path

import pytest
import pywintypes
import win32api
import win32con

from seller_automation_utils import WorkbookRefreshError

SCRIPT = Path(__file__).resolve().parent.parent / "run_amzn_account_health.py"
WANTED = {
    "COUNT_UNAVAILABLE",
    "_premium_shipping_values",
    "MACRO_TIMEOUT_SEC",
    "_macro_failure",
    "_run_checked_macro",
    "_end_excel",
    "_kill_pid",
    "_pin_pid",
}


def _load_isolated() -> dict:
    """Compile only the wanted top-level names out of the entry-point script.

    The Excel helpers need a handful of module-level names; the real objects are
    supplied here instead of running the script's imports, which reach config files.

    Returns:
        Namespace holding the extracted constants and functions.
    """
    tree = ast.parse(SCRIPT.read_text(encoding="utf-8"))
    kept: list[ast.stmt] = [
        node
        for node in tree.body
        if (isinstance(node, ast.FunctionDef) and node.name in WANTED)
        or (
            isinstance(node, (ast.Assign, ast.AnnAssign))
            and any(
                isinstance(t, ast.Name) and t.id in WANTED
                for t in (node.targets if isinstance(node, ast.Assign) else [node.target])
            )
        )
    ]
    namespace: dict = {
        "log": logging.getLogger("amzn_account_health.tests"),
        "os": os,
        "signal": signal,
        "threading": threading,
        "pywintypes": pywintypes,
        "win32api": win32api,
        "win32con": win32con,
        "WorkbookRefreshError": WorkbookRefreshError,
    }
    exec(compile(ast.Module(body=kept, type_ignores=[]), str(SCRIPT), "exec"), namespace)
    return namespace


@pytest.fixture
def excel_helpers() -> dict:
    """A fresh namespace with the Excel helpers, so a test may patch it freely."""
    return _load_isolated()


@pytest.fixture(scope="session")
def script_tree() -> ast.Module:
    """The parsed entry-point script, for structural refusal checks."""
    return ast.parse(SCRIPT.read_text(encoding="utf-8"))


@pytest.fixture(scope="session")
def premium_shipping_values():
    """The `_premium_shipping_values` function, loaded without side effects."""
    return _load_isolated()["_premium_shipping_values"]


@pytest.fixture(scope="session")
def count_unavailable() -> str:
    """The placeholder written into cells Amazon no longer supplies."""
    return _load_isolated()["COUNT_UNAVAILABLE"]
