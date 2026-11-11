"""What Save does after the launch-time "Restore it?" prompt is answered."""

import threading

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

    assert result["ok"] is True
    assert result["filename"] == str(chosen)
    assert result["recovery_gen"] > 0
    assert chosen.exists()
    assert old_file.read_bytes() == before


def test_accepted_restore_saves_to_the_source_file_without_a_dialog(old_file):
    window = FakeWindow()
    api = launch(window)

    api.check_recovery()
    first = api.save_project([item("Restored work")])
    second = api.save_project([item("Edited again")])

    assert window.dialogs == 0
    assert first["ok"] is True
    assert first["filename"] == str(old_file)
    assert second["ok"] is True
    assert second["filename"] == str(old_file)
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


def unreadable_setup(tmp_path, monkeypatch, data=b"not a workbook"):
    monkeypatch.setattr(spm, "app_dir", lambda: str(tmp_path))
    recovery = tmp_path / ".spm_recovery.xlsx"
    recovery.write_bytes(data)
    return recovery, tmp_path / ".spm_recovery_unreadable.xlsx"


def test_unreadable_recovery_is_kept_and_reported(tmp_path, monkeypatch):
    recovery, kept = unreadable_setup(tmp_path, monkeypatch)
    api = logging_api(FakeWindow())

    result = api.check_recovery()

    assert result["exists"] is False
    assert result["warn"] == spm.RECOVERY_UNREADABLE_KEPT
    assert any(line.startswith("Recovery read failed") for line in api.logged)
    assert kept.read_bytes() == b"not a workbook"
    assert not recovery.exists()
    assert api._filename == ""


def test_kept_unreadable_copy_is_not_offered_again(tmp_path, monkeypatch):
    unreadable_setup(tmp_path, monkeypatch)
    api = logging_api(FakeWindow())
    api.check_recovery()

    assert api.check_recovery() == {"exists": False}


def test_new_unreadable_copy_replaces_the_kept_one(tmp_path, monkeypatch):
    _, kept = unreadable_setup(tmp_path, monkeypatch, b"new bytes")
    kept.write_bytes(b"older bytes")
    api = logging_api(FakeWindow())

    api.check_recovery()

    assert kept.read_bytes() == b"new bytes"


def test_valid_copy_that_cannot_be_read_is_kept_intact(old_file, tmp_path, monkeypatch):
    original = (tmp_path / ".spm_recovery.xlsx").read_bytes()
    api = logging_api(FakeWindow())

    def refuse(path):
        raise PermissionError("held by another program")

    monkeypatch.setattr(api, "_read_xlsx", refuse)

    result = api.check_recovery()

    assert result["warn"] == spm.RECOVERY_UNREADABLE_KEPT
    assert (tmp_path / ".spm_recovery_unreadable.xlsx").read_bytes() == original


def test_unreadable_copy_that_cannot_be_moved_is_reported(tmp_path, monkeypatch):
    recovery, _ = unreadable_setup(tmp_path, monkeypatch)
    api = logging_api(FakeWindow())

    def refuse(src, dst):
        raise PermissionError("held by another program")

    monkeypatch.setattr(spm.os, "replace", refuse)

    result = api.check_recovery()

    assert result["warn"] == spm.RECOVERY_UNREADABLE_STUCK
    assert "Recovery set-aside failed: PermissionError" in api.logged
    assert recovery.read_bytes() == b"not a workbook"


def test_successful_delete_reports_nothing(old_file):
    api = logging_api(FakeWindow())
    api._filename = str(old_file)

    assert "warn" not in api.save_project([item("Fresh save")])
    assert not [line for line in api.logged if "Recovery delete failed" in line]


def test_write_with_old_gen_after_clear_is_discarded(old_file, tmp_path):
    api = launch(FakeWindow())
    old_gen = api.get_state()["recovery_gen"]

    api.write_recovery([item("Before clear")], old_gen)
    api.new_project()
    result = api.write_recovery([item("Stale")], old_gen)

    assert result == {"ok": True, "stale": True}
    assert not (tmp_path / ".spm_recovery.xlsx").exists()


def test_write_with_current_gen_after_clear_writes(old_file, tmp_path):
    api = launch(FakeWindow())

    result = api.new_project()
    write = api.write_recovery([item("Current")], result["recovery_gen"])

    assert write == {"ok": True}
    assert (tmp_path / ".spm_recovery.xlsx").exists()


def test_save_result_carries_new_gen(old_file, tmp_path):
    api = launch(FakeWindow(choice=str(old_file)))
    before = api.get_state()["recovery_gen"]

    saved = api._do_save(str(old_file), [item("Saved")])
    cleared = api.clear_recovery()
    new = api.new_project()
    opened = api.open_project()

    assert saved["recovery_gen"] > before
    assert cleared["recovery_gen"] > saved["recovery_gen"]
    assert new["recovery_gen"] > cleared["recovery_gen"]
    assert opened["recovery_gen"] > new["recovery_gen"]


def test_write_started_before_clear_cannot_survive_it(old_file, tmp_path, monkeypatch):
    api = launch(FakeWindow())
    entered = threading.Event()
    release = threading.Event()
    real_write = api._write_xlsx

    def blocked_write(path, items, source=""):
        if path == str(tmp_path / ".spm_recovery.xlsx"):
            entered.set()
            assert release.wait(5), "recovery write was not released"
        return real_write(path, items, source)

    monkeypatch.setattr(api, "_write_xlsx", blocked_write)
    gen = api.get_state()["recovery_gen"]
    writer = threading.Thread(target=api.write_recovery, args=([item("Racing")], gen))
    clearer = threading.Thread(target=api.new_project)
    writer.start()
    assert entered.wait(5), "recovery write did not start"
    clearer.start()
    # Give the clear time to delete the copy while the write is still blocked,
    # so the test fails if the two are not made to take turns.
    clearer.join(0.5)
    release.set()
    writer.join(5)
    clearer.join(5)

    assert not writer.is_alive(), "recovery write thread did not finish"
    assert not clearer.is_alive(), "recovery clear thread did not finish"
    assert not (tmp_path / ".spm_recovery.xlsx").exists()


def test_failed_delete_still_bumps_gen(undeletable_recovery):
    api = logging_api(FakeWindow())
    old_gen = api.get_state()["recovery_gen"]

    result = api.new_project()
    stale = api.write_recovery([item("Stale")], old_gen)

    assert result["recovery_gen"] > old_gen
    assert stale == {"ok": True, "stale": True}


def test_write_without_gen_still_writes(tmp_path, monkeypatch):
    monkeypatch.setattr(spm, "app_dir", lambda: str(tmp_path))
    api = launch(FakeWindow())

    assert api.write_recovery([item("Compatible")]) == {"ok": True}
    assert (tmp_path / ".spm_recovery.xlsx").exists()
