"""Market-run failure classification (src/vendor_intel/batch/failure_classifier.py)."""
from __future__ import annotations

import httpx

from vendor_intel.batch.failure_classifier import classify_failure


def test_network_errors_are_transient():
    assert classify_failure(ConnectionError("Connection reset by peer")) == "transient"
    assert classify_failure(TimeoutError("Read timed out")) == "transient"
    assert classify_failure(OSError("[WinError 10060] connection attempt failed")) == "transient"


def test_rate_limit_and_5xx_are_transient():
    assert classify_failure(RuntimeError("http_429: rate limit exceeded")) == "transient"
    assert classify_failure(RuntimeError("http_503: Service Unavailable")) == "transient"
    assert classify_failure(RuntimeError("DDGSException: No results found.")) == "transient"


def test_programming_and_config_errors_are_permanent():
    assert classify_failure(ValueError("Provide a query or phase1_plan_path")) == "permanent"
    assert classify_failure(FileNotFoundError("seeds file not found: x.txt")) == "permanent"
    assert classify_failure(KeyError("industry")) == "permanent"


def test_unknown_error_defaults_to_transient():
    class WeirdCustomError(Exception):
        pass

    assert classify_failure(WeirdCustomError("something unexpected")) == "transient"


def test_real_httpx_status_error_is_transient_when_429():
    request = httpx.Request("POST", "https://api.example.com/v1/chat")
    response = httpx.Response(429, request=request, text="rate limit exceeded")
    exc = httpx.HTTPStatusError("429 rate limited", request=request, response=response)
    assert classify_failure(exc) == "transient"
