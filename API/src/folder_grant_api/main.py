from contextlib import asynccontextmanager

from fastapi import FastAPI

from folder_grant_api import __version__
from folder_grant_api.api.routes import health, preview
from folder_grant_api.infrastructure.database import create_database
from folder_grant_api.settings import get_settings


@asynccontextmanager
async def lifespan(app: FastAPI):
    settings = get_settings()
    engine, session_factory = create_database(settings)
    app.state.database_engine = engine
    app.state.session_factory = session_factory
    yield
    engine.dispose()


def create_app() -> FastAPI:
    app = FastAPI(
        title="Folder Grant API",
        version=__version__,
        description="BUS와 테스트 클라이언트가 사용하는 폴더 권한 처리 API",
        lifespan=lifespan,
    )
    app.include_router(health.router)
    app.include_router(preview.router)
    return app


app = create_app()
