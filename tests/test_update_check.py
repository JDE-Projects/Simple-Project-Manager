import errno
import json
import socket
import ssl
import urllib.error

import simple_project_manager as spm


def _http_error(code):
    return urllib.error.HTTPError("https://api.github.com/x", code, "msg", {}, None)


def _url_error(reason):
    return urllib.error.URLError(reason)


def test_http_error_403():
    reason = spm._update_error_reason(_http_error(403))
    assert "rate-limiting" in reason


def test_http_error_404():
    reason = spm._update_error_reason(_http_error(404))
    assert reason == "No published release was found."


def test_http_error_5xx():
    reason = spm._update_error_reason(_http_error(503))
    assert reason == "GitHub is having trouble on its end (HTTP 503)."


def test_http_error_other_code():
    reason = spm._update_error_reason(_http_error(418))
    assert reason == "GitHub returned an error (HTTP 418)."


def test_json_decode_error():
    exc = json.JSONDecodeError("Expecting value", "<html>", 0)
    reason = spm._update_error_reason(exc)
    assert "unexpected" in reason


def test_ssl_cert_verification_error():
    exc = _url_error(ssl.SSLCertVerificationError("certificate verify failed"))
    reason = spm._update_error_reason(exc)
    assert "certificate could not be verified" in reason


def test_ssl_eof_error():
    exc = _url_error(ssl.SSLEOFError("EOF occurred"))
    reason = spm._update_error_reason(exc)
    assert reason == "The secure connection was cut off during the handshake with GitHub."


def test_ssl_zero_return_error():
    exc = _url_error(ssl.SSLZeroReturnError("zero return"))
    reason = spm._update_error_reason(exc)
    assert reason == "The secure connection was cut off during the handshake with GitHub."


def test_generic_ssl_error():
    exc = _url_error(ssl.SSLError("generic ssl failure"))
    reason = spm._update_error_reason(exc)
    assert reason == "The secure connection to GitHub failed."


def test_socket_gaierror():
    exc = _url_error(socket.gaierror("name resolution failed"))
    reason = spm._update_error_reason(exc)
    assert "could not be looked up" in reason


def test_socket_timeout():
    exc = _url_error(socket.timeout("timed out"))
    reason = spm._update_error_reason(exc)
    assert reason == "GitHub didn't respond in time."


def test_timeout_error():
    exc = _url_error(TimeoutError("timed out"))
    reason = spm._update_error_reason(exc)
    assert reason == "GitHub didn't respond in time."


def test_connection_refused_error():
    exc = _url_error(ConnectionRefusedError("connection refused"))
    reason = spm._update_error_reason(exc)
    assert "refused or reset" in reason


def test_connection_reset_error():
    exc = _url_error(ConnectionResetError("connection reset"))
    reason = spm._update_error_reason(exc)
    assert "refused or reset" in reason


def test_os_error_network_unreachable():
    cause = OSError("network is unreachable")
    cause.errno = errno.ENETUNREACH
    exc = _url_error(cause)
    reason = spm._update_error_reason(exc)
    assert reason == "No network connection."


def test_plain_url_error_fallback():
    exc = _url_error("some unclassified reason string")
    reason = spm._update_error_reason(exc)
    assert reason == "Couldn't reach GitHub. Check the internet connection."


def test_generic_exception_fallback():
    exc = ValueError("something odd happened")
    reason = spm._update_error_reason(exc)
    assert reason == "ValueError: something odd happened"


def test_generic_exception_fallback_truncated():
    exc = ValueError("x" * 200)
    reason = spm._update_error_reason(exc)
    assert len(reason) <= 120
    assert reason.endswith("...")
    assert reason.startswith("ValueError: ")
