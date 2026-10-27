from fastapi import APIRouter, Depends, Request
from folder_grant_api.api.security import require_api_key
from folder_grant_api.api.routes.preview import create_plan
from folder_grant_api.api.routes.project_jobs import create_project_plan
from folder_grant_api.schemas.access_job import AccessJobRequest
from folder_grant_api.schemas.project_job import ProjectJobRequest
from folder_grant_api.schemas.validation import ValidationResponse
from folder_grant_api.services.validation_service import access_probes, project_probes, validate_resources, validate_access_resources

router = APIRouter(tags=['validation'], dependencies=[Depends(require_api_key)])


@router.post('/api/v1/access-jobs/validate', response_model=ValidationResponse)
def validate_access(payload: AccessJobRequest, request: Request):
    settings = request.app.state.settings
    plan = create_plan(payload, settings)
    return validate_access_resources(payload, plan, settings)


@router.post('/api/v1/project-jobs/validate', response_model=ValidationResponse)
def validate_project(payload: ProjectJobRequest, request: Request):
    settings = request.app.state.settings
    plan = create_project_plan(payload, settings)
    return validate_resources(payload.request_id, project_probes(payload, plan, settings), settings)
