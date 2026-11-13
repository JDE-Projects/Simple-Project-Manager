"""open_url opens only https addresses on exactly jde-projects.com."""

import webbrowser

import pytest

import simple_project_manager as spm


@pytest.fixture
def opened(monkeypatch):
    """Records every address handed to the system browser instead of opening it."""
    calls = []
    monkeypatch.setattr(webbrowser, "open", calls.append)
    return calls


@pytest.fixture
def api():
    api = spm.Api()
    api.logged = []
    api.log = api.logged.append
    return api


@pytest.mark.parametrize("url", [
    "https://jde-projects.com",
    "https://jde-projects.com/",
    "https://jde-projects.com/support/#commercial-licensing",
    "https://jde-projects.com/privacy/?app=spm",
])
def test_website_addresses_open(api, opened, url):
    assert api.open_url(url) == {"ok": True}
    assert opened == [url]
    assert api.logged == []


@pytest.mark.parametrize("url", [
    # wrong scheme
    "http://jde-projects.com",
    "file:///C:/Windows/System32/calc.exe",
    "javascript:alert(1)",
    "ftp://jde-projects.com",
    # wrong or lookalike host
    "https://github.com/JDE-Projects/Simple-Project-Manager",
    "https://jde-projects.com.evil.example",
    "https://evil.example/jde-projects.com",
    "https://www.jde-projects.com",
    "https://evil-jde-projects.com",
    "https://jde-projects.com@evil.example",
    "https://user@jde-projects.com",
    "https://jde-projects.com:8443",
    "https://JDE-Projects.com",
    # not a usable address
    "",
    "jde-projects.com",
    "https://[jde-projects.com",
    None,
    123,
])
def test_other_addresses_are_refused_without_opening(api, opened, url):
    assert api.open_url(url) == {"ok": False}
    assert opened == []
    assert api.logged == ["open_url refused: address is not on jde-projects.com"]
