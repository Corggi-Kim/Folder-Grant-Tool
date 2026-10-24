import base64
import json
import re
from unittest.mock import patch
import pytest
from sqlalchemy import select, func
from folder_grant_api.infrastructure.models import JobRecord
from folder_grant_api.infrastructure.powershell import PowerShellExecutionError, PowerShellResult
from test_api import build_client, valid_payload, valid_project_payload


def result_for_script(command, override=None):
    encoded=re.search(r"FromBase64String\('([^']+)'\)",command).group(1)
    spec=json.loads(base64.b64decode(encoded))
    checks=[dict(id=p['id'],label=p['label'],target=p['target'],status='passed',message='exists') for p in spec['probes']]
    for check in checks:
        if override and check['id'] in override:
            check.update(status=override[check['id']])
    return PowerShellResult(0,json.dumps(dict(checks=checks,checked_by='DOMAIN\\service')),'')


@pytest.mark.parametrize('kind,payload',[('access',valid_payload()),('project',valid_project_payload())])
def test_validate_auth_mock_and_no_job(tmp_path,kind,payload):
    with build_client(tmp_path) as client, patch('folder_grant_api.services.validation_service.PowerShellRunner.run') as runner:
        path=f'/api/v1/{kind}-jobs/validate'
        client.headers.pop('X-API-Key')
        assert client.post(path,json=payload).status_code==401
        client.headers['X-API-Key']='test-api-key'
        body=client.post(path,json=payload).json()
        assert body['validation_status']=='unknown'
        assert 'Mock' in body['checks'][0]['message']
        runner.assert_not_called()
        with client.app.state.session_factory() as db:
            assert db.scalar(select(func.count()).select_from(JobRecord))==0


@pytest.mark.parametrize('kind,payload',[('access',valid_payload()),('project',valid_project_payload())])
def test_live_validation_uses_read_only_commands_and_leaves_database_empty(tmp_path,kind,payload):
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.services.validation_service.PowerShellRunner.run',side_effect=result_for_script) as runner:
            body=client.post(f'/api/v1/{kind}-jobs/validate',json=payload).json()
        assert body['validation_status']=='passed'
        assert body['checked_by']=='DOMAIN\\service'
        command=runner.call_args.args[0]
        assert not re.search(r'(?i)\b(Add-AD|Remove-AD|New-AD|Set-Acl|icacls|robocopy)',command)
        with client.app.state.session_factory() as db:
            assert db.scalar(select(func.count()).select_from(JobRecord))==0


@pytest.mark.parametrize('status',['failed','unknown','warning'])
def test_missing_or_unreadable_resource_is_not_passed(tmp_path,status):
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.services.validation_service.PowerShellRunner.run',side_effect=lambda s:result_for_script(s,{'level3':status})):
            body=client.post('/api/v1/access-jobs/validate',json=valid_payload()).json()
        assert body['validation_status']==status


@pytest.mark.parametrize('value',[PowerShellExecutionError('POWERSHELL_TIMEOUT','timeout'),PowerShellResult(0,'{"checks":[]}',''),PowerShellResult(0,'bad json','')])
def test_timeout_or_invalid_result_is_unknown(tmp_path,value):
    with build_client(tmp_path,executor_mode='powershell') as client:
        options={'side_effect':value} if isinstance(value,Exception) else {'return_value':value}
        with patch('folder_grant_api.services.validation_service.PowerShellRunner.run',**options):
            body=client.post('/api/v1/access-jobs/validate',json=valid_payload()).json()
        assert body['validation_status']=='unknown'
        assert body['checks'][0]['status']=='unknown'


def test_plan_preview_stays_offline_and_invalid_input_does_not_query(tmp_path):
    with build_client(tmp_path,executor_mode='powershell') as client, patch('folder_grant_api.services.validation_service.PowerShellRunner.run') as runner:
        assert client.post('/api/v1/access-jobs/preview',json=valid_payload()).status_code==200
        payload=valid_payload();payload['role']='invalid'
        assert client.post('/api/v1/access-jobs/validate',json=payload).status_code==422
        runner.assert_not_called()
