"""The settings file is written safely and a damaged one is kept aside."""

import json
import os

import pytest

import simple_project_manager as spm

FOLDER_MARKER = "PrefsFolderMarker"


@pytest.fixture
def app_folder(tmp_path, monkeypatch):
    folder = tmp_path / FOLDER_MARKER
    folder.mkdir()
    monkeypatch.setattr(spm, "app_dir", lambda: str(folder))
    return folder


def pref_file(folder):
    return folder / "simple_project_manager.pref"


def leftovers(folder):
    return [p.name for p in folder.iterdir() if p.name.endswith(".tmp")]


def test_save_then_load_round_trips(app_folder):
    prefs = {"theme": "light", "window": {"x": 1, "y": 2, "width": 3, "height": 4}}
    assert spm.save_prefs(prefs) is True
    assert spm.load_prefs() == prefs
    assert leftovers(app_folder) == []


def test_missing_file_is_first_run_and_logs_nothing(app_folder):
    logged = []
    assert spm.load_prefs(logged.append) == {}
    assert logged == []


def test_failed_write_leaves_existing_settings_unchanged(app_folder, monkeypatch):
    original = {"theme": "light"}
    pref_file(app_folder).write_text(json.dumps(original), encoding="utf-8")

    def broken_dump(obj, f):
        f.write('{"theme": "da')   # partial write, then the crash
        raise OSError(28, "No space left on device", str(app_folder))

    monkeypatch.setattr(spm.json, "dump", broken_dump)
    logged = []

    assert spm.save_prefs({"theme": "dark"}, logged.append) is False
    monkeypatch.undo()
    monkeypatch.setattr(spm, "app_dir", lambda: str(app_folder))

    assert spm.load_prefs() == original
    assert leftovers(app_folder) == []
    assert logged == ["Could not save settings: OSError: No space left on device"]


def test_failed_swap_leaves_existing_settings_unchanged(app_folder, monkeypatch):
    original = {"theme": "light"}
    pref_file(app_folder).write_text(json.dumps(original), encoding="utf-8")

    def refuse(src, dst):
        raise PermissionError(13, "Access is denied", str(dst))

    monkeypatch.setattr(spm.os, "replace", refuse)
    logged = []
    assert spm.save_prefs({"theme": "dark"}, logged.append) is False
    monkeypatch.undo()

    assert json.loads(pref_file(app_folder).read_text(encoding="utf-8")) == original
    assert leftovers(app_folder) == []
    assert logged == ["Could not save settings: PermissionError: Access is denied"]


@pytest.mark.parametrize("content", ['{"theme": "da', "[1, 2]", "\xff\xfe junk"])
def test_damaged_file_is_kept_aside_and_logged(app_folder, content):
    path = pref_file(app_folder)
    path.write_bytes(content.encode("latin-1"))
    logged = []

    assert spm.load_prefs(logged.append) == {}

    assert not path.exists()
    kept = [p for p in app_folder.iterdir()
            if p.name.startswith("simple_project_manager.pref.corrupt-")]
    assert len(kept) == 1
    assert kept[0].read_bytes() == content.encode("latin-1")
    assert len(logged) == 1
    assert logged[0].startswith("Settings file unreadable, kept aside: ")
    assert FOLDER_MARKER not in logged[0]


def test_next_save_after_damage_starts_fresh_without_touching_kept_copy(app_folder):
    pref_file(app_folder).write_text("not json", encoding="utf-8")
    prefs = spm.load_prefs()
    prefs["theme"] = "dark"
    assert spm.save_prefs(prefs) is True

    assert spm.load_prefs() == {"theme": "dark"}
    kept = [p for p in app_folder.iterdir() if ".corrupt-" in p.name]
    assert len(kept) == 1
    assert kept[0].read_text(encoding="utf-8") == "not json"


def test_unopenable_file_is_left_in_place(app_folder, monkeypatch):
    path = pref_file(app_folder)
    path.write_text('{"theme": "light"}', encoding="utf-8")
    real_open = open

    def locked_open(file, mode="r", *args, **kwargs):
        if os.fspath(file) == str(path):
            raise PermissionError(13, "Access is denied", str(file))
        return real_open(file, mode, *args, **kwargs)

    monkeypatch.setattr("builtins.open", locked_open)
    logged = []
    assert spm.load_prefs(logged.append) == {}
    monkeypatch.undo()

    assert path.exists()
    assert [p.name for p in app_folder.iterdir()] == [path.name]
    assert logged == ["Settings file not read: PermissionError: Access is denied"]


def test_save_theme_reports_failure(app_folder, monkeypatch):
    monkeypatch.setattr(spm, "save_prefs", lambda prefs, log=None: False)
    api = spm.Api()
    api.logged = []
    api.log = api.logged.append

    assert api.save_theme("light") == {"ok": False}
    assert "Could not save theme pref: save_prefs failed" in api.logged
