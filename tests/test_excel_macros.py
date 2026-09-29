"""Tests for the checked, time-bounded chart macros and for ending this run's own Excel.

`modUtilities.deleteCharts` and `modUtilities.resizeCharts` used to end in a `MsgBox` on
error. Excel runs hidden, so that modal waits on a window nobody can see: the macro call
never returns, the run neither finishes nor crashes, and `handle_crash` never fires. Both
are now `Function ... As String` returning `""` or a reason, and the Python raises on a
reason. An old `Sub` in the workbook returns `None`, which is still success, so this is
safe before the workbook module is re-pasted.

Since library 1.8.8 `handle_crash` only ends Excel started through `excel_utils`; this
script opens its own `xw.App`, so it must end that Excel itself, on every path.

No real Excel is involved: the workbook and app are in-memory fakes.
"""
from __future__ import annotations

import ast
import os
import re
import signal
import subprocess
import sys
import threading
from types import SimpleNamespace

import pytest
import pywintypes
import win32event

from seller_automation_utils import WorkbookRefreshError

from conftest import SCRIPT

REASON = "deleteCharts failed - error 1004: Delete method of Range class failed"


def _com_error(message: str = "Exception occurred.") -> pywintypes.com_error:
    return pywintypes.com_error(-2147352567, message, None, None)


class _FakeWorkbook:
    def __init__(self, result: object = "", error: BaseException | None = None, block: threading.Event | None = None):
        self.result = result
        self.error = error
        self.block = block
        self.macro_names: list[str] = []

    def macro(self, name: str):
        self.macro_names.append(name)

        def _run():
            if self.block is not None:
                # A modal on the hidden Excel: only killing the process releases the call,
                # and the COM call then fails because the server is gone.
                self.block.wait(10)
                raise pywintypes.com_error(-2147023174, "The RPC server is unavailable.", None, None)
            if self.error is not None:
                raise self.error
            return self.result

        return _run


class _FakeExcel:
    """Minimal ``xw.App`` recording how it was ended."""

    def __init__(self, quit_error: BaseException | None = None, kill_error: BaseException | None = None):
        self.pid = 4242
        self.quit_error = quit_error
        self.kill_error = kill_error
        self.display_alerts = True
        self.screen_updating = True
        self.alerts_at_quit: bool | None = None
        self.quit_called = False
        self.kill_calls = 0

    def quit(self):
        self.quit_called = True
        self.alerts_at_quit = self.display_alerts
        if self.quit_error is not None:
            raise self.quit_error

    def kill(self):
        self.kill_calls += 1
        if self.kill_error is not None:
            raise self.kill_error


@pytest.mark.parametrize(
    ("result", "expected"),
    [
        pytest.param("", None, id="hardened-success"),
        pytest.param(None, None, id="legacy-sub-returns-none"),
        pytest.param(REASON, REASON, id="reason"),
        pytest.param(True, "returned an unexpected bool", id="bool-is-not-a-contract-value"),
        pytest.param(0, "returned an unexpected int", id="number-is-not-a-contract-value"),
    ],
)
def test_macro_failure_reads_the_contract(excel_helpers, result, expected):
    failure = excel_helpers["_macro_failure"](result)

    if expected is None:
        assert failure is None
    else:
        assert expected in failure


@pytest.mark.parametrize("result", ["", None])
def test_a_clean_macro_returns_and_turns_alerts_back_off(excel_helpers, result):
    """The macros' Cleanup sets DisplayAlerts = True; a later failed save would then prompt."""
    excel, wb = _FakeExcel(), _FakeWorkbook(result)

    excel_helpers["_run_checked_macro"](excel, wb, "modUtilities.deleteCharts")

    assert wb.macro_names == ["modUtilities.deleteCharts"]
    assert excel.display_alerts is False
    assert excel.kill_calls == 0


@pytest.mark.parametrize("macro", ["modUtilities.deleteCharts", "modUtilities.resizeCharts"])
def test_a_reported_reason_raises(excel_helpers, macro):
    excel, wb = _FakeExcel(), _FakeWorkbook(REASON)

    with pytest.raises(WorkbookRefreshError) as excinfo:
        excel_helpers["_run_checked_macro"](excel, wb, macro)

    assert macro in str(excinfo.value)
    assert REASON in str(excinfo.value)
    assert excel.display_alerts is False


def test_a_com_error_from_the_macro_propagates_unchanged(excel_helpers):
    error = _com_error()
    excel, wb = _FakeExcel(), _FakeWorkbook(error=error)

    with pytest.raises(pywintypes.com_error) as excinfo:
        excel_helpers["_run_checked_macro"](excel, wb, "modUtilities.resizeCharts")

    assert excinfo.value is error
    assert excel.kill_calls == 0


def test_a_stuck_macro_kills_its_own_excel_and_says_so(excel_helpers):
    """A COM call already waiting on a modal can only be freed by killing that Excel.

    The kill runs on the watchdog thread by the pid read up front: touching the xlwings
    App from that thread would be a COM call into the very Excel that is stuck.
    """
    released = threading.Event()
    killed: list[int] = []
    excel, wb = _FakeExcel(), _FakeWorkbook(block=released)
    excel_helpers["MACRO_TIMEOUT_SEC"] = 0.2
    excel_helpers["_kill_pid"] = lambda pid: (killed.append(pid), released.set())

    with pytest.raises(WorkbookRefreshError, match=r"modUtilities\.deleteCharts.*exceeded 0\.2s"):
        excel_helpers["_run_checked_macro"](excel, wb, "modUtilities.deleteCharts")

    assert killed == [4242]
    assert excel.kill_calls == 0


def test_a_clean_macro_never_fires_the_watchdog(excel_helpers):
    killed: list[int] = []
    excel_helpers["_kill_pid"] = killed.append
    excel_helpers["MACRO_TIMEOUT_SEC"] = 0.2

    excel_helpers["_run_checked_macro"](_FakeExcel(), _FakeWorkbook(""), "modUtilities.resizeCharts")
    threading.Event().wait(0.4)

    assert killed == []


class _TimerAlreadyFiring:
    """A ``threading.Timer`` whose callback had already started when ``cancel()`` came.

    ``cancel()`` cannot stop it; the callback finishes on its own thread, which is what
    ``join()`` waits for.
    """

    def __init__(self, _interval, function):
        self.function = function
        self.daemon = False

    def start(self):
        pass

    def cancel(self):
        pass

    def join(self):
        self.function()


def test_a_timeout_firing_as_the_macro_returns_is_still_reported(excel_helpers):
    """Without the join, the macro's result is trusted while its Excel is being killed."""
    killed: list[int] = []
    excel_helpers["_kill_pid"] = killed.append
    excel_helpers["threading"] = SimpleNamespace(Timer=_TimerAlreadyFiring, Event=threading.Event)

    with pytest.raises(WorkbookRefreshError, match=r"modUtilities\.resizeCharts exceeded"):
        excel_helpers["_run_checked_macro"](_FakeExcel(), _FakeWorkbook(""), "modUtilities.resizeCharts")

    assert killed == [4242]


def test_a_pinned_pid_still_names_the_exited_process(excel_helpers):
    """While pinned, the pid cannot be reused, so a late kill fails instead of hitting another process."""
    child = subprocess.Popen([sys.executable, "-c", "import sys; sys.stdin.read()"], stdin=subprocess.PIPE)
    pin = excel_helpers["_pin_pid"](child.pid)
    try:
        child.stdin.close()
        child.wait(10)

        assert win32event.WaitForSingleObject(pin, 0) == win32event.WAIT_OBJECT_0
        with pytest.raises(OSError):
            os.kill(child.pid, signal.SIGTERM)
        excel_helpers["_kill_pid"](child.pid)
    finally:
        pin.Close()


def test_pin_pid_returns_none_when_the_pid_cannot_be_opened(excel_helpers):
    assert excel_helpers["_pin_pid"](0) is None


def test_kill_pid_targets_exactly_one_pid(excel_helpers, monkeypatch):
    calls: list[tuple[int, int]] = []
    monkeypatch.setattr(excel_helpers["os"], "kill", lambda pid, sig: calls.append((pid, sig)))

    excel_helpers["_kill_pid"](4242)

    assert calls == [(4242, excel_helpers["signal"].SIGTERM)]


def test_kill_pid_ignores_a_process_that_already_exited(excel_helpers, monkeypatch):
    def gone(pid, sig):
        raise OSError(87, "The parameter is incorrect")

    monkeypatch.setattr(excel_helpers["os"], "kill", gone)

    excel_helpers["_kill_pid"](4242)


def test_the_macro_time_bound_is_finite_and_generous(excel_helpers):
    # 3x the slowest open+deleteCharts on record (23s) is under the 300s floor.
    assert 300 <= excel_helpers["MACRO_TIMEOUT_SEC"] <= 900


@pytest.mark.parametrize(
    ("quit_error", "kill_error"),
    [
        pytest.param(None, None, id="clean-quit"),
        pytest.param(_com_error(), None, id="quit-fails-kill-still-runs"),
        pytest.param(None, OSError("process already gone"), id="kill-after-exit-is-harmless"),
    ],
)
def test_end_excel_always_kills_its_own_pid(excel_helpers, quit_error, kill_error):
    excel = _FakeExcel(quit_error, kill_error)

    excel_helpers["_end_excel"](excel)

    assert excel.quit_called and excel.kill_calls == 1
    assert excel.alerts_at_quit is False, "a Save prompt on a hidden Excel would hang the quit"


def _main(script_tree: ast.Module) -> ast.FunctionDef:
    return next(n for n in script_tree.body if isinstance(n, ast.FunctionDef) and n.name == "main")


def _calls(node: ast.AST) -> list[str]:
    return [ast.unparse(n.func) for n in ast.walk(node) if isinstance(n, ast.Call)]


def test_main_never_calls_a_macro_unchecked(script_tree):
    """Refusal guard: every macro goes through ``_run_checked_macro``, never ``wb.macro(...)()``."""
    direct = [c for c in _calls(_main(script_tree)) if c.endswith(".macro")]

    assert direct == []
    assert _calls(_main(script_tree)).count("_run_checked_macro") == 2


def test_main_ends_its_excel_in_finally(script_tree):
    """The crash path must end this Excel too; ``handle_crash`` no longer does it."""
    try_node = next(n for n in _main(script_tree).body if isinstance(n, ast.Try))
    finally_calls = [c for stmt in try_node.finalbody for c in _calls(stmt)]

    assert "_end_excel" in finally_calls
    assert "excel.quit" not in _calls(_main(script_tree)), "quit without kill can leave Excel behind"


def test_main_holds_its_excel_pid_until_the_last_kill(script_tree):
    """The pin is taken straight after launch and released only after ``_end_excel``."""
    main = _main(script_tree)
    try_node = next(n for n in main.body if isinstance(n, ast.Try))
    body = "\n".join(ast.unparse(s) for s in try_node.body)
    finally_src = "\n".join(ast.unparse(s) for s in try_node.finalbody)

    assert body.index("excel = xw.App(visible=False)") < body.index("excel_pin = _pin_pid(excel.pid)") < body.index("_run_checked_macro")
    assert finally_src.index("_end_excel(excel)") < finally_src.index("excel_pin.Close()")


def test_module_has_no_machine_wide_kill():
    source = SCRIPT.read_text(encoding="utf-8")

    assert "kill_app" not in source
    assert not re.search(r"taskkill|/im\b|pkill|killall", source, re.IGNORECASE)
