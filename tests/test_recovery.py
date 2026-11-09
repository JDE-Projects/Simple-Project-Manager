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
