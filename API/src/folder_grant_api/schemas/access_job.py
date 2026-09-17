from datetime import datetime
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator

from folder_grant_api.domain.operation_plan import Operation, ProjectStatus


class AccessJobRequest(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    request_id: str = Field(min_length=1, max_length=100)
    source: Literal["test-client", "bus", "swagger"]
    operation: Operation
    project_status: ProjectStatus
    employee_id: str = Field(min_length=1, max_length=50)
    project_code: str = Field(min_length=1, max_length=50)
    level2: str | None = Field(default=None, max_length=50)
    level3: str | None = Field(default=None, max_length=50)
    role: str | None = Field(default=None, max_length=100)
    requested_by: str = Field(min_length=1, max_length=100)

    @field_validator("level3")
    @classmethod
    def reject_path_separators(cls, value: str | None) -> str | None:
        if value and ("\\" in value or "/" in value):
            raise ValueError("Level3에는 경로 구분자를 사용할 수 없습니다.")
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
    executor_mode: Literal["mock"] = "mock"
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
