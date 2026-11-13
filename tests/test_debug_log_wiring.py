import builtins
import json
import os
import threading

import simple_project_manager as spm
from app.debug_log import DebugLog


def test_set_debug_creates_log_and_reports_enabled(tmp_path, monkeypatch):
    monkeypatch.setattr(spm, "app_dir", lambda: str(tmp_path))
    api = spm.Api()

    assert api.set_debug(True) == {"ok": True, "enabled": True}
    logs = list(tmp_path.glob("Debug_Log_*.txt"))
    assert len(logs) == 1
    api.log("Saved project")
    assert "Saved project" in logs[0].read_text(encoding="utf-8")


def test_log_when_off_writes_nothing(tmp_path, monkeypatch):
    monkeypatch.setattr(spm, "app_dir", lambda: str(tmp_path))
    api = spm.Api()

    api.log("x")

    assert list(tmp_path.glob("Debug_Log_*.txt")) == []


def test_set_debug_failure_reports_disabled_and_warns_before_window(tmp_path, monkeypatch):
    monkeypatch.setattr(spm, "app_dir", lambda: str(tmp_path))
    monkeypatch.setattr(DebugLog, "_write_new_file", lambda self, path: None)
    api = spm.Api()

    assert api.set_debug(True) == {"ok": False, "enabled": False}
    assert len(api.get_state()["log_warnings"]) == 1
    assert api.get_state()["log_warnings"] == []


def test_warning_with_window_is_pushed_to_page(tmp_path, monkeypatch):
    monkeypatch.setattr(spm, "app_dir", lambda: str(tmp_path))
    event = threading.Event()
    scripts = []

    class Window:
        def evaluate_js(self, script):
            scripts.append(script)
            event.set()

    api = spm.Api()
    api.set_window(Window())
    api._on_debug_log_warning("Debug log: test")

    assert event.wait(2)
    assert "onDebugLogWarning" in scripts[0]
    assert json.dumps("Debug log: test") in scripts[0]
    assert "false" in scripts[0]


def test_failed_write_turns_logging_off_and_warns(tmp_path, monkeypatch):
    monkeypatch.setattr(spm, "app_dir", lambda: str(tmp_path))
    api = spm.Api()
    api.set_debug(True)
    active_path = os.path.normcase(os.fspath(api._debug_log._path))
    real_open = builtins.open

    def fake_open(file, mode="r", *args, **kwargs):
        if "a" in mode and os.path.normcase(os.fspath(file)) == active_path:
            raise OSError(5, "simulated disk failure")
        return real_open(file, mode, *args, **kwargs)

    monkeypatch.setattr(builtins, "open", fake_open)
    api.log("x")

    assert api._debug_log.is_enabled() is False
    warnings = api.get_state()["log_warnings"]
    assert len(warnings) == 1
    assert "Logging turned off" in warnings[0]


def test_get_state_has_empty_log_warnings_by_default(tmp_path, monkeypatch):
    monkeypatch.setattr(spm, "app_dir", lambda: str(tmp_path))

    assert spm.Api().get_state()["log_warnings"] == []
