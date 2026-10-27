"""Read-only live checks are separate from plans and executed Job results."""
from datetime import datetime
from typing import Literal
from pydantic import BaseModel, Field


class ValidationCheck(BaseModel):
    id: str
    label: str
    target: str
    status: Literal['passed', 'failed', 'warning', 'unknown', 'not_applicable']
    message: str


class GroupPreparation(BaseModel):
    group_name: str
    root_path: str
    isolated_path: str


class ValidationResponse(BaseModel):
    group_preparation: GroupPreparation | None = None
    request_id: str
    executor_mode: Literal['mock', 'powershell']
    validation_status: Literal['passed', 'failed', 'warning', 'unknown']
    checked_at: datetime
    checked_by: str | None = None
    checks: list[ValidationCheck]
    warnings: list[str] = Field(default_factory=list)
