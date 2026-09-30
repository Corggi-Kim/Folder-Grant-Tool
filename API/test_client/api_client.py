import json
from dataclasses import dataclass
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import quote
from urllib.request import Request, urlopen


@dataclass(frozen=True)
class ApiResponse:
    status_code: int
    data: Any


class ApiClientError(RuntimeError):
    def __init__(self, message: str, *, status_code: int | None = None, detail: Any = None):
        self.status_code = status_code
        self.detail = detail
        super().__init__(message)


class FolderGrantApiClient:
    def __init__(self, base_url: str, api_key: str, timeout_seconds: int = 30):
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key.strip()
        self.timeout_seconds = timeout_seconds

    def health(self) -> ApiResponse:
        return self._request("GET", "/health", authenticated=False)

    def ready(self) -> ApiResponse:
        return self._request("GET", "/ready")

    def preview(self, payload: dict) -> ApiResponse:
        return self._request("POST", "/api/v1/access-jobs/preview", payload)

    def register(self, payload: dict) -> ApiResponse:
        return self._request("POST", "/api/v1/access-jobs", payload)

    def get_job(self, job_id: str) -> ApiResponse:
        return self._request("GET", f"/api/v1/access-jobs/{quote(job_id, safe='')}")

    def cancel(self, job_id: str, requested_by: str) -> ApiResponse:
        return self._request(
            "POST",
            f"/api/v1/access-jobs/{quote(job_id, safe='')}/cancel",
            {"requested_by": requested_by},
        )

    def _request(
        self,
        method: str,
        path: str,
        payload: dict | None = None,
        *,
        authenticated: bool = True,
    ) -> ApiResponse:
        if authenticated and not self.api_key:
            raise ApiClientError("API Key를 입력하세요.")
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8") if payload is not None else None
        headers = {"Accept": "application/json"}
        if body is not None:
            headers["Content-Type"] = "application/json; charset=utf-8"
        if authenticated:
            headers["X-API-Key"] = self.api_key
        request = Request(self.base_url + path, data=body, headers=headers, method=method)
        try:
            with urlopen(request, timeout=self.timeout_seconds) as response:
                raw = response.read().decode("utf-8")
                return ApiResponse(response.status, json.loads(raw) if raw else None)
        except HTTPError as exc:
            raw = exc.read().decode("utf-8", errors="replace")
            try:
                detail = json.loads(raw)
            except json.JSONDecodeError:
                detail = raw
            raise ApiClientError(
                _error_message(detail, f"HTTP {exc.code}"),
                status_code=exc.code,
                detail=detail,
            ) from exc
        except URLError as exc:
            raise ApiClientError(f"API 서버에 연결할 수 없습니다: {exc.reason}") from exc
        except TimeoutError as exc:
            raise ApiClientError("API 요청 시간이 초과되었습니다.") from exc


def _error_message(detail: Any, fallback: str) -> str:
    if isinstance(detail, dict):
        value = detail.get("detail")
        if isinstance(value, dict):
            return str(value.get("message") or value.get("code") or fallback)
        if isinstance(value, list):
            messages = [str(item.get("msg", item)) if isinstance(item, dict) else str(item) for item in value]
            return "; ".join(messages)
        if value:
            return str(value)
    return fallback
