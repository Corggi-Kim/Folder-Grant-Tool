from hmac import compare_digest

from fastapi import HTTPException, Request, Security, status
from fastapi.security import APIKeyHeader

api_key_header = APIKeyHeader(name="X-API-Key", auto_error=False)


def require_api_key(request: Request, supplied_key: str | None = Security(api_key_header)) -> None:
    settings = request.app.state.settings
    if not supplied_key or not compare_digest(supplied_key, settings.api_key):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail={"code": "INVALID_API_KEY", "message": "유효한 API Key가 필요합니다."},
        )
