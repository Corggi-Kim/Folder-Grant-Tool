from functools import lru_cache
from pathlib import Path
from typing import Literal

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    environment: str = "development"
    database_url: str = "sqlite:///./data/folder_grant.db"
    api_key: str = "change-me"
    executor_mode: Literal["mock", "powershell"] = "mock"
    powershell_path: str = "powershell.exe"
    powershell_timeout_seconds: int = 300
    template_root: str = r"\\LSK_S010\Study folder\_Template"
    group_ou_path: str = r"OU=Group Project Folder,OU=0.Management Object Group,OU=lskglobal,DC=lskglobal,DC=com"
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

    def validate_runtime_safety(self) -> None:
        if self.executor_mode == "powershell" and self.api_key == "change-me":
            raise ValueError("PowerShell 모드에서는 기본 API Key를 사용할 수 없습니다.")
        if self.powershell_timeout_seconds < 1:
            raise ValueError("PowerShell timeout은 1초 이상이어야 합니다.")


@lru_cache
def get_settings() -> Settings:
    return Settings()
