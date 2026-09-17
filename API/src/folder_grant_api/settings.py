from functools import lru_cache
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    environment: str = "development"
    database_url: str = "sqlite:///./data/folder_grant.db"
    api_key: str = "change-me"
    share_root: str = r"\\LSK_S010\Study Folder\{proj_seg}\{lv2}\{lv3}"
    closed_root: str = r"\\192.168.1.95\Study_Closed"
    closed_archive_root: str = r"\\192.168.1.95\Study_Archive"

    model_config = SettingsConfigDict(
        env_file=".env",
        env_prefix="FGT_",
        extra="ignore",
    )

    def ensure_local_directories(self) -> None:
        if self.database_url.startswith("sqlite:///./"):
            relative_path = self.database_url.removeprefix("sqlite:///./")
            Path(relative_path).parent.mkdir(parents=True, exist_ok=True)


@lru_cache
def get_settings() -> Settings:
    return Settings()
