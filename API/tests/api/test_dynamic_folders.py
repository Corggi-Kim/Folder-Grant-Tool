from unittest.mock import patch

import pytest
from test_api import build_client, valid_payload
from folder_grant_api.schemas.access_job import AccessJobRequest
from folder_grant_api.settings import Settings
from folder_grant_api.api.routes.preview import create_plan
from folder_grant_api.domain.operation_plan import build_operation_plan, Operation, ProjectStatus, PlanValidationError


@pytest.mark.parametrize('folder',['DM_IWRS','DM_CUSTOM','Project Documents','한글폴더','straße'])
@pytest.mark.parametrize('operation',['grant','revoke'])
def test_unlisted_folder_uses_existing_general_folder_rules(folder,operation):
    payload=AccessJobRequest(**dict(valid_payload(),level2='Isolated',level3=folder,role=None,operation=operation))
    plan=create_plan(payload,Settings())
    assert plan.target_paths==['\\\\LSK_S010\\Study Folder\\26012\\Isolated\\'+folder]
    assert [s.type for s in plan.steps]==(['add_ad_group_member','grant_folder_acl'] if operation=='grant' else ['remove_ad_group_member','revoke_folder_acl'])
    assert plan.steps[1].metadata['principal']=='12345@lskglobal.com'
    if operation=='grant': assert plan.steps[1].permission=='modify'


@pytest.mark.parametrize('folder',['','.', '..', '../DM',r'DM\CO',r'\\server\share','C:DM','DM:stream',
    '*','DM?','DM[1]','DM|CO','DM<CO','DM>CO','DM"CO','DM\x00','DM\nCO','DM.',
    'CON','con.txt','LPT1','COM¹','CONIN$','x'*51])
def test_schema_and_direct_plan_reject_unsafe_folder_names(folder):
    with pytest.raises(ValueError):AccessJobRequest(**dict(valid_payload(),level3=folder))
    with pytest.raises(PlanValidationError):
        build_operation_plan(operation=Operation.GRANT,project_status=ProjectStatus.PROGRESS,
            employee_id='E1',project_code='26051',level2='Isolated',level3=folder,role=None,
            share_root=r'\\LSK_S010\Study Folder\{proj_seg}\{lv2}\{lv3}',closed_root='CLOSED',archive_root='ARCHIVE')


@pytest.mark.parametrize('path',['/preview','/validate',''])
def test_api_accepts_dynamic_folder_and_keeps_path_within_project(tmp_path,path):
    with build_client(tmp_path) as client:
        response=client.post('/api/v1/access-jobs'+path,json=dict(valid_payload(),level2='Isolated',level3='DM_IWRS',role=None))
        assert response.status_code in (200,202),response.text
        body=response.json()
        if path=='/preview':assert body['target_paths']==[r'\\LSK_S010\Study Folder\26012\Isolated\DM_IWRS']
        if path=='/validate':assert body['validation_status']=='unknown'  # Mock never invents live success.


@pytest.mark.parametrize('path',['/preview','/validate','','/prepare-group'])
def test_api_rejects_traversal_before_any_powershell_or_job(tmp_path,path):
    data=dict(valid_payload(),level2='Isolated',level3=r'..\DM_IWRS')
    if path=='/prepare-group':data['confirm_group_creation']=True
    with build_client(tmp_path,executor_mode='powershell') as client,patch('folder_grant_api.services.validation_service.PowerShellRunner.run') as runner:
        assert client.post('/api/v1/access-jobs'+path,json=data).status_code==422
        runner.assert_not_called()


@pytest.mark.parametrize('folder',['dm','stat'])
def test_replay_of_legacy_lowercase_payload_remains_idempotent(tmp_path,folder):
    from folder_grant_api.infrastructure.models import JobRecord
    data=dict(valid_payload(),level3=folder,role=None)
    with build_client(tmp_path) as client:
        job=client.post('/api/v1/access-jobs',json=data).json()
        with client.app.state.session_factory() as db:
            stored=db.get(JobRecord,job['job_id'])
            stored.request_payload=dict(stored.request_payload,level3=folder)
            db.commit()
        replay=client.post('/api/v1/access-jobs',json=data)
        assert replay.status_code==200,replay.text
        assert replay.json()['job_id']==job['job_id']
        assert replay.json()['replayed']


def test_closed_requests_keep_legacy_empty_unused_level3(tmp_path):
    data=dict(valid_payload(),project_status='closed',level2=None,level3='',role=None)
    with build_client(tmp_path) as client:
        first=client.post('/api/v1/access-jobs',json=data)
        assert first.status_code==202,first.text
        replay=client.post('/api/v1/access-jobs',json=data)
        assert replay.status_code==200
        assert replay.json()['job_id']==first.json()['job_id']


def test_dynamic_folder_does_not_bypass_existing_role_rules():
    payload=AccessJobRequest(**dict(valid_payload(),level2='Isolated',level3='DM_IWRS',role='Manager'))
    from fastapi import HTTPException
    with pytest.raises(HTTPException) as error:create_plan(payload,Settings())
    assert error.value.detail['code']=='INVALID_ROLE'
