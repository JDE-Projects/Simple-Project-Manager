"""Saving swaps a finished temporary file into place and never harms the old file."""

import os

import pytest

import simple_project_manager as spm

REAL_REPLACE = os.replace
SAVE_FAILED = "Couldn't save the project file."
FOLDER_MESSAGE = "Couldn't save: this folder doesn't allow new files."


def item(name):
    return {"id": name, "type": "Phase", "name": name}


def temp_files(folder):
    return sorted(p.name for p in folder.glob(".spm_*.tmp"))


@pytest.fixture
def api(tmp_path, monkeypatch):
    monkeypatch.setattr(spm, "app_dir", lambda: str(tmp_path))
    monkeypatch.setattr(spm.time, "sleep", lambda s: None)
    a = spm.Api()
    a.logged = []
    a.log = a.logged.append
    return a


@pytest.fixture
def target(api, tmp_path):
    """An existing project file holding one item."""
    path = tmp_path / "project.xlsx"
    api._write_xlsx(str(path), [item("Original")])
    return path


def raise_os_error(*args, **kwargs):
    raise OSError("simulated failure")


def test_normal_save_round_trips_and_leaves_no_temp_file(api, tmp_path):
    path = tmp_path / "new.xlsx"
    api._filename = str(path)

    result = api.save_project([item("One"), item("Two")])

    assert result["ok"] is True
    items, _ = api._read_xlsx(str(path))
    assert [i["name"] for i in items] == ["One", "Two"]
    assert temp_files(tmp_path) == []


def test_failed_workbook_write_keeps_the_existing_file(api, target, tmp_path, monkeypatch):
    before = target.read_bytes()
    monkeypatch.setattr(spm.Workbook, "save", raise_os_error)

    result = api._do_save(str(target), [item("New")])

    assert result["ok"] is False
    assert result["error"] == SAVE_FAILED
    assert target.read_bytes() == before
    assert temp_files(tmp_path) == []


def test_failed_disk_flush_keeps_the_existing_file(api, target, tmp_path, monkeypatch):
    before = target.read_bytes()
    monkeypatch.setattr(spm.os, "fsync", raise_os_error)

    result = api._do_save(str(target), [item("New")])

    assert result["ok"] is False
    assert target.read_bytes() == before
    assert temp_files(tmp_path) == []


def test_refused_swap_reports_excel_and_keeps_the_file(api, target, tmp_path, monkeypatch):
    before = target.read_bytes()
    calls = []

    def refuse(src, dst):
        calls.append(src)
        raise PermissionError(13, "denied")

    monkeypatch.setattr(spm.os, "replace", refuse)

    result = api._do_save(str(target), [item("New")])

    assert result["ok"] is False
    assert "open in Excel" in result["error"]
    assert target.read_bytes() == before
    assert temp_files(tmp_path) == []
    assert any(line.startswith("Save swap refused: PermissionError") for line in api.logged)
    assert len(calls) == spm.SWAP_RETRIES + 1


def test_swap_succeeds_after_brief_refusal(api, target, monkeypatch):
    before = target.read_bytes()
    calls = []

    def flaky(src, dst):
        calls.append(src)
        if len(calls) <= 2:
            raise PermissionError(13, "denied")
        return REAL_REPLACE(src, dst)

    monkeypatch.setattr(spm.os, "replace", flaky)

    result = api._do_save(str(target), [item("Changed")])

    assert result["ok"] is True
    assert len(calls) == 3
    assert target.read_bytes() != before
    items, _ = api._read_xlsx(str(target))
    assert [i["name"] for i in items] == ["Changed"]


def test_other_swap_errors_are_not_retried(api, target, tmp_path, monkeypatch):
    before = target.read_bytes()
    calls = []

    def fail(src, dst):
        calls.append(src)
        raise OSError("simulated failure")

    monkeypatch.setattr(spm.os, "replace", fail)

    result = api._do_save(str(target), [item("New")])

    assert len(calls) == 1
    assert result["ok"] is False
    assert result["error"] == SAVE_FAILED
    assert target.read_bytes() == before
    assert temp_files(tmp_path) == []


def test_folder_refusing_new_files_is_reported(api, target, monkeypatch):
    before = target.read_bytes()

    def refuse(*args, **kwargs):
        raise PermissionError(13, "denied")

    monkeypatch.setattr(spm.tempfile, "mkstemp", refuse)

    result = api._do_save(str(target), [item("New")])

    assert result == {"ok": False, "error": FOLDER_MESSAGE}
    assert target.read_bytes() == before


def test_temp_file_that_cannot_be_removed_is_logged_and_original_error_kept(
        api, target, monkeypatch):
    real_remove = os.remove

    def remove(p):
        if str(p).endswith(".tmp"):
            raise PermissionError(13, "denied")
        return real_remove(p)

    monkeypatch.setattr(spm.Workbook, "save", raise_os_error)
    monkeypatch.setattr(spm.os, "remove", remove)

    result = api._do_save(str(target), [item("New")])

    assert result["ok"] is False
    assert result["error"] == SAVE_FAILED
    assert "Temporary save file not removed: PermissionError" in api.logged


@pytest.mark.skipif(os.name != "nt", reason="needs Windows file locking")
def test_file_held_open_by_another_program_is_refused(api, target, tmp_path):
    before = target.read_bytes()

    with open(target, "rb"):
        result = api._do_save(str(target), [item("New")])

    assert result["ok"] is False
    assert "open in Excel" in result["error"]
    assert target.read_bytes() == before
    assert temp_files(tmp_path) == []


def test_recovery_write_failure_keeps_the_previous_copy(api, tmp_path, monkeypatch):
    assert api.write_recovery([item("Earlier")]) == {"ok": True}
    recovery = tmp_path / ".spm_recovery.xlsx"
    before = recovery.read_bytes()
    monkeypatch.setattr(spm.Workbook, "save", raise_os_error)

    result = api.write_recovery([item("Later")])

    assert result == {"ok": False}
    assert recovery.read_bytes() == before
    assert temp_files(tmp_path) == []
