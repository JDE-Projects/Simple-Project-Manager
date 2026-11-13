"""Debug log lines carry only an action and its outcome: never a file path,
typed text, or an error's own message (which can quote either)."""

import re
import urllib.request
from pathlib import Path

import pytest

import simple_project_manager as spm

FOLDER_MARKER = "ClientFolderMarker"
TYPED_MARKER = "TypedNameMarker"


class FakeWindow:
    def __init__(self, choice):
        self.choice = choice

    def create_file_dialog(self, *args, **kwargs):
        return self.choice


@pytest.fixture
def project_path(tmp_path, monkeypatch):
    monkeypatch.setattr(spm, "app_dir", lambda: str(tmp_path))
    folder = tmp_path / FOLDER_MARKER
    folder.mkdir()
    return folder / f"{TYPED_MARKER}.xlsx"


def logging_api(choice):
    api = spm.Api()
    api.set_window(FakeWindow(choice))
    api.logged = []
    api.log = api.logged.append
    return api


def items():
    return [{"id": "p1", "type": "Phase", "name": TYPED_MARKER}]


def assert_private(lines):
    assert lines, "expected at least one log line"
    for line in lines:
        assert FOLDER_MARKER not in line
        assert TYPED_MARKER not in line


def test_save_and_open_log_no_path_or_typed_text(project_path):
    api = logging_api(str(project_path))

    assert api.save_as(items())["ok"] is True
    assert api.open_project()["ok"] is True

    assert_private(api.logged)
    assert "Saved project (1 rows)" in api.logged
    assert "Opened project (1 rows)" in api.logged


def test_save_failure_logs_error_kind_not_message(project_path, monkeypatch):
    api = logging_api(str(project_path))

    def fail(path, rows, **kwargs):
        raise ValueError(f"could not write {path} with {TYPED_MARKER}")

    monkeypatch.setattr(api, "_write_xlsx", fail)

    assert api.save_as(items())["ok"] is False
    assert api.logged == ["Save failed: ValueError"]


def test_open_failure_logs_error_kind_not_path(project_path, monkeypatch):
    api = logging_api(str(project_path))

    def fail(path):
        raise FileNotFoundError(2, "The system cannot find the file specified", path)

    monkeypatch.setattr(api, "_read_xlsx", fail)

    assert api.open_project()["ok"] is False
    assert api.logged == [
        "Open failed: FileNotFoundError: The system cannot find the file specified"]


def test_recovery_write_failure_logs_error_kind_not_message(project_path, monkeypatch):
    api = logging_api(None)

    def fail(path, rows, **kwargs):
        raise ValueError(f"{path} {TYPED_MARKER}")

    monkeypatch.setattr(api, "_write_xlsx", fail)
    api.write_recovery(items())

    assert api.logged == ["Recovery write failed: ValueError"]


def test_update_check_failure_logs_error_kind_not_message(monkeypatch):
    api = logging_api(None)

    def fail(*args, **kwargs):
        raise ValueError(f"proxy said {TYPED_MARKER}")

    monkeypatch.setattr(urllib.request, "urlopen", fail)
    api.check_update()

    assert api.logged == ["check_update failed: ValueError"]


def test_no_log_call_interpolates_a_path_url_or_raw_error():
    """Guard for future edits: a log line may not drop a path, address, or
    raw exception straight into its text."""
    source = Path(spm.__file__).read_text(encoding="utf-8")
    calls = re.findall(r"self\.log\((.*?)\)\s*$", source, flags=re.MULTILINE)
    assert len(calls) > 15
    # The settings-file helpers log through _log_quietly; their calls span
    # lines, so they are matched across line breaks.
    pref_calls = re.findall(r"_log_quietly\(log,(.*?)\)\s*$", source,
                            flags=re.MULTILINE | re.DOTALL)
    assert len(pref_calls) >= 4
    banned = re.compile(r"\{(e|move_error|path|aside|tmp|url|sel|"
                        r"self\._filename|str\(e\))\}")
    offenders = [call for call in calls + pref_calls if banned.search(call)]
    assert offenders == []
