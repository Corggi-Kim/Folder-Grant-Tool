import json
from unittest.mock import patch

import pytest

from test_client.api_client import ApiClientError, FolderGrantApiClient


class FakeResponse:
    status = 200

    def __init__(self, data: dict):
        self.data = data

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def read(self) -> bytes:
        return json.dumps(self.data).encode()


def test_ready_sends_api_key_header() -> None:
    with patch("test_client.api_client.urlopen", return_value=FakeResponse({"status": "ready"})) as open_url:
        response = FolderGrantApiClient("http://server:8000/", "secret").ready()

    assert response.data == {"status": "ready"}
    request = open_url.call_args.args[0]
    assert request.get_header("X-api-key") == "secret"
    assert request.full_url == "http://server:8000/ready"


def test_protected_request_requires_api_key() -> None:
    with pytest.raises(ApiClientError, match="API Key"):
        FolderGrantApiClient("http://server:8000", "").ready()


def test_register_serializes_unicode_payload_as_json() -> None:
    payload = {"requested_by": "테스터"}
    with patch("test_client.api_client.urlopen", return_value=FakeResponse({"job_id": "1"})) as open_url:
        FolderGrantApiClient("http://server:8000", "secret").register(payload)

    request = open_url.call_args.args[0]
    assert json.loads(request.data.decode("utf-8")) == payload
