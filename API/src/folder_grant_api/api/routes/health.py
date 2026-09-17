from fastapi import APIRouter, HTTPException, Request

from folder_grant_api.infrastructure.database import check_database

router = APIRouter(tags=["health"])


@router.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok", "service": "folder-grant-api"}


@router.get("/ready")
def ready(request: Request) -> dict[str, str]:
    try:
        check_database(request.app.state.database_engine)
    except Exception as exc:
        raise HTTPException(status_code=503, detail="database unavailable") from exc
    return {"status": "ready", "database": "ok"}
