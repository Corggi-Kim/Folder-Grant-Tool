import re
from datetime import datetime
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator


class ProjectJobRequest(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    request_id: str = Field(min_length=1, max_length=100)
    source: Literal["test-client", "bus", "swagger"]
    project_code: str = Field(min_length=4, max_length=20)
    project_name: str = Field(default="", max_length=300)
    requested_by: str = Field(min_length=1, max_length=100)

    @field_validator("project_code")
    @classmethod
    def validate_project_code(cls, value: str) -> str:
        if not re.fullmatch(r"(?:\d{4}|\d{5}(?:A\d+|-\d+)?)", value, flags=re.IGNORECASE):
            raise ValueError("프로젝트 코드는 4~5자리 숫자와 선택적인 A 접미사 형식이어야 합니다.")
        return value


class ProjectPlanStepResponse(BaseModel):
    type: str
    description: str
    target: str
    metadata: dict = Field(default_factory=dict)


class ProjectPreviewResponse(BaseModel):
    valid: bool = True
    request_id: str
    canonical_project_code: str
    group_name: str
    root_path: str
    study_all_path: str
    template_path: str
    steps: list[ProjectPlanStepResponse]
    warnings: list[str]


class ProjectJobStepResponse(BaseModel):
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


class ProjectJobResponse(BaseModel):
    job_id: str
    request_id: str
    replayed: bool = False
    status: str
    executor_mode: Literal["mock", "powershell"]
    project_code: str
    project_name: str
    requested_by: str
    cancel_requested: bool
    result: dict | None
    error_code: str | None
    error_message: str | None
    created_at: datetime
    started_at: datetime | None
    finished_at: datetime | None
    steps: list[ProjectJobStepResponse]


class CancelProjectJobRequest(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    requested_by: str = Field(min_length=1, max_length=100)
