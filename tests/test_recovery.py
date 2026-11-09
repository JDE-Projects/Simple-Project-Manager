"""What Save does after the launch-time "Restore it?" prompt is answered."""

import pytest

import simple_project_manager as spm


class FakeWindow:
    """Stands in for the app window and records each Save dialog it is asked for."""

    def __init__(self, choice=None):
        self.choice = choice
        self.dialogs = 0

    def create_file_dialog(self, *args, **kwargs):
        self.dialogs += 1
        return self.choice


def item(name):
    return {"id": name, "type": "Phase", "name": name}


@pytest.fixture
def old_file(tmp_path, monkeypatch):
    """An existing project file, plus a recovery copy left by a previous session."""
    monkeypatch.setattr(spm, "app_dir", lambda: str(tmp_path))
    path = tmp_path / "old.xlsx"
    previous = spm.Api()
    previous._write_xlsx(str(path), [item("Saved on disk")])
    previous._filename = str(path)
    previous.write_recovery([item("Unsaved work")])
    return path


def launch(window):
    api = spm.Api()
    api.set_window(window)
    return api


def test_recovery_offer_names_the_source_file(old_file):
    rec = launch(FakeWindow()).check_recovery()

    assert rec["exists"] is True
    assert rec["filename"] == str(old_file)
    assert [i["name"] for i in rec["items"]] == ["Unsaved work"]


def test_declined_restore_save_asks_where_and_leaves_old_file_alone(old_file):
    window = FakeWindow(choice=None)
    api = launch(window)
    before = old_file.read_bytes()

    api.check_recovery()
    api.clear_recovery()
    result = api.save_project([item("Typed after declining")])

    assert window.dialogs == 1
    assert result == {"cancelled": True}
    assert old_file.read_bytes() == before


def test_declined_restore_deletes_the_recovery_copy(old_file):
    api = launch(FakeWindow())

    api.check_recovery()
    api.clear_recovery()

    assert api.check_recovery() == {"exists": False}


def test_declined_restore_save_writes_to_the_chosen_file(old_file, tmp_path):
    chosen = tmp_path / "new.xlsx"
    window = FakeWindow(choice=str(chosen))
    api = launch(window)
    before = old_file.read_bytes()

    api.check_recovery()
    api.clear_recovery()
    result = api.save_project([item("Typed after declining")])

    assert result == {"ok": True, "filename": str(chosen)}
    assert chosen.exists()
    assert old_file.read_bytes() == before


def test_accepted_restore_saves_to_the_source_file_without_a_dialog(old_file):
    window = FakeWindow()
    api = launch(window)

    api.check_recovery()
    first = api.save_project([item("Restored work")])
    second = api.save_project([item("Edited again")])

    assert window.dialogs == 0
    assert first == {"ok": True, "filename": str(old_file)}
    assert second == {"ok": True, "filename": str(old_file)}
    items, _ = api._read_xlsx(str(old_file))
    assert [i["name"] for i in items] == ["Edited again"]


@pytest.fixture
def undeletable_recovery(old_file, tmp_path, monkeypatch):
    """Make deleting the recovery copy fail, as when another program holds it."""
    real_remove = spm.os.remove

    def refuse_recovery(path):
        if str(path).endswith(".spm_recovery.xlsx"):
            raise PermissionError("held by another program")
        real_remove(path)

    monkeypatch.setattr(spm.os, "remove", refuse_recovery)
    return tmp_path / ".spm_recovery.xlsx"


def logging_api(window):
    api = launch(window)
    api.logged = []
    api.log = api.logged.append
    return api


def assert_reported(api, result):
    assert result["warn"] == spm.RECOVERY_DELETE_WARNING
    assert "Recovery delete failed: PermissionError" in api.logged


def test_failed_delete_after_save_keeps_the_save_and_reports(old_file, undeletable_recovery):
    api = logging_api(FakeWindow())
    api._filename = str(old_file)

    result = api.save_project([item("Fresh save")])

    assert result["ok"] is True
    assert_reported(api, result)
    assert undeletable_recovery.exists()
    items, _ = api._read_xlsx(str(old_file))
    assert [i["name"] for i in items] == ["Fresh save"]


def test_failed_delete_on_new_project_reports(undeletable_recovery):
    api = logging_api(FakeWindow())

    result = api.new_project()

    assert result["ok"] is True
    assert_reported(api, result)


def test_failed_delete_on_open_reports(old_file, undeletable_recovery):
    api = logging_api(FakeWindow(choice=str(old_file)))

    result = api.open_project()

    assert result["ok"] is True
    assert_reported(api, result)


def test_failed_delete_on_open_keeps_the_workbook_warning(undeletable_recovery, tmp_path):
    foreign = tmp_path / "foreign.xlsx"
    spm.Workbook().save(foreign)
    api = logging_api(FakeWindow(choice=str(foreign)))

    result = api.open_project()

    assert "no Simple Project Manager marker" in result["warn"]
    assert spm.RECOVERY_DELETE_WARNING in result["warn"]


def test_failed_delete_on_decline_reports_and_still_forgets_the_file(undeletable_recovery):
    api = logging_api(FakeWindow())
    api.check_recovery()

    result = api.clear_recovery()

    assert result["ok"] is False
    assert_reported(api, result)
    assert api._filename == ""


def test_failed_delete_of_unreadable_recovery_copy_reports(undeletable_recovery):
    undeletable_recovery.write_bytes(b"not a workbook")
    api = logging_api(FakeWindow())

    result = api.check_recovery()

    assert result["exists"] is False
    assert_reported(api, result)


def test_successful_delete_reports_nothing(old_file):
    api = logging_api(FakeWindow())
    api._filename = str(old_file)

    assert "warn" not in api.save_project([item("Fresh save")])
    assert not [line for line in api.logged if "Recovery delete failed" in line]
