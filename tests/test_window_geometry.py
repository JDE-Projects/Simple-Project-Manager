import os
from types import SimpleNamespace

import simple_project_manager as spm


class FakeFunction:
    def __init__(self, callback):
        self.callback = callback

    def __call__(self, *args):
        return self.callback(*args)


class FakeUser32:
    def __init__(self, windows):
        self.windows = windows
        self.EnumWindows = FakeFunction(self._enum_windows)
        self.GetWindowThreadProcessId = FakeFunction(self._get_pid)
        self.GetWindowTextLengthW = FakeFunction(self._get_text_length)
        self.GetWindowTextW = FakeFunction(self._get_text)
        self.IsWindowVisible = FakeFunction(self._is_visible)

    def _enum_windows(self, callback, lparam):
        for hwnd in self.windows:
            if not callback(hwnd, lparam):
                break
        return True

    def _get_pid(self, hwnd, pid):
        pid._obj.value = self.windows[hwnd]["pid"]
        return 1

    def _get_text_length(self, hwnd):
        return len(self.windows[hwnd]["title"])

    def _get_text(self, hwnd, buf, length):
        buf.value = self.windows[hwnd]["title"]
        return len(buf.value)

    def _is_visible(self, hwnd):
        return self.windows[hwnd]["visible"]


def test_own_window_handle_returns_matching_visible_window_for_this_process(monkeypatch):
    own_pid = os.getpid()
    user32 = FakeUser32({
        1: {"pid": own_pid + 1, "title": "Simple Project Manager", "visible": True},
        2: {"pid": own_pid, "title": "Other window", "visible": True},
        3: {"pid": own_pid, "title": "Simple Project Manager", "visible": False},
        4: {"pid": own_pid, "title": "Simple Project Manager", "visible": True},
    })
    monkeypatch.setattr(spm.ctypes, "windll", SimpleNamespace(user32=user32))
    monkeypatch.setattr(spm.ctypes, "WINFUNCTYPE", lambda *args: lambda callback: callback)

    assert spm._own_window_handle("Simple Project Manager") == 4


def test_own_window_handle_quietly_returns_none_when_win32_lookup_fails(monkeypatch):
    class BrokenWindll:
        @property
        def user32(self):
            raise OSError("unavailable")

    monkeypatch.setattr(spm.ctypes, "windll", BrokenWindll())

    assert spm._own_window_handle("Simple Project Manager") is None
