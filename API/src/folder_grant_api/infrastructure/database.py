from datetime import datetime, timezone

from sqlalchemy import DateTime, Integer, String, create_engine, text
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, sessionmaker

from folder_grant_api.settings import Settings


class Base(DeclarativeBase):
    pass


class ServiceMetadata(Base):
    __tablename__ = "service_metadata"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    key: Mapped[str] = mapped_column(String(100), unique=True, nullable=False)
    value: Mapped[str] = mapped_column(String(500), nullable=False)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
        nullable=False,
    )


def create_database(settings: Settings):
    # Import model mappings before create_all so all tables are registered.
    from folder_grant_api.infrastructure import models  # noqa: F401

    settings.ensure_local_directories()
    connect_args = {"check_same_thread": False} if settings.database_url.startswith("sqlite") else {}
    engine = create_engine(settings.database_url, connect_args=connect_args)
    Base.metadata.create_all(engine)
    return engine, sessionmaker(bind=engine, expire_on_commit=False)


def check_database(engine) -> bool:
    with engine.connect() as connection:
        connection.execute(text("SELECT 1"))
    return True
