from datetime import datetime
import re
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator

from folder_grant_api.domain.operation_plan import Operation, ProjectStatus


class AccessJobRequest(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    request_id: str = Field(min_length=1, max_length=100)
    source: Literal["test-client", "bus", "swagger"]
    operation: Operation
    project_status: ProjectStatus
    employee_id: str = Field(min_length=1, max_length=50, pattern=r"^[A-Za-z0-9._-]+$")
    project_code: str = Field(min_length=4, max_length=20)
    level2: str | None = Field(default=None, max_length=50)
    level3: str | None = Field(default=None, max_length=50)
    role: str | None = Field(default=None, max_length=100)
    requested_by: str = Field(min_length=1, max_length=100)

    @field_validator("level3")
    @classmethod
    def reject_path_separators(cls, value: str | None) -> str | None:
        allowed = {"ARS", "CO", "DM", "ER", "MW", "PM", "PV", "RA", "SSU", "STAT", "STAT_IDMC", "ETC"}
        if value and value.upper() not in allowed:
            raise ValueError("허용되지 않는 Level3입니다.")
        return value

    @field_validator("level2")
    @classmethod
    def validate_level2(cls, value: str | None) -> str | None:
        if value and value.lower() not in {"study", "isolated"}:
            raise ValueError("Level2는 Study 또는 Isolated여야 합니다.")
        return value

    @field_validator("project_code")
    @classmethod
    def validate_project_code(cls, value: str) -> str:
        if not re.fullmatch(r"(?:\d{4}|\d{5}(?:A\d+|-\d+)?)", value, flags=re.IGNORECASE):
            raise ValueError("프로젝트 코드는 4~5자리 숫자와 선택적인 A 접미사 형식이어야 합니다.")
        return value


class PlanStepResponse(BaseModel):
    type: str
    description: str
    target: str
    permission: str | None = None
    metadata: dict = Field(default_factory=dict)


class PreviewResponse(BaseModel):
    valid: bool = True
    request_id: str
    canonical_project_code: str
    target_group: str | None
    target_paths: list[str]
    steps: list[PlanStepResponse]
    warnings: list[str]


class ErrorDetail(BaseModel):
    code: str
    message: str


class ErrorResponse(BaseModel):
    detail: ErrorDetail


class JobStepResponse(BaseModel):
    order: int
    type: str
    status: str
    target: str
    description: str
    details: dict
    result: dict | None
    error_code: str | None
    error_message: str | None
    started_at: datetime | None
    finished_at: datetime | None


class JobResponse(BaseModel):
    job_id: str
    request_id: str
    replayed: bool = False
    status: str
    executor_mode: Literal["mock", "powershell"]
    operation: str
    project_status: str
    requested_by: str
    cancel_requested: bool
    result: dict | None
    error_code: str | None
    error_message: str | None
    created_at: datetime
    started_at: datetime | None
    finished_at: datetime | None
    steps: list[JobStepResponse]


class CancelJobRequest(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    requested_by: str = Field(min_length=1, max_length=100)
