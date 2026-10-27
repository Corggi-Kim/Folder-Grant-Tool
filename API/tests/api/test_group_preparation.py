import json
from unittest.mock import patch
import pytest
from test_api import build_client, valid_payload
from test_validation import result_for_script
from folder_grant_api.infrastructure.powershell import PowerShellResult


def payload():
    return dict(valid_payload(), level2='Isolated', role=None, request_id='PREP-001', confirm_group_creation=True)


def ready_script(command, changes=None):
    raw=json.loads(result_for_script(command).stdout)
    for c in raw['checks']:
        if c['id']=='group': c.update(status='failed',message='대상 없음')
        if c['id']=='membership': c.update(status='unknown',message='선행 대상이 없거나 확인되지 않아 검사하지 못했습니다.')
        if changes and c['id'] in changes: c.update(changes[c['id']])
    return PowerShellResult(0,json.dumps(raw),'')


@pytest.mark.parametrize('level3',['STAT','ETC'])
def test_offer_and_prepare_three_steps_replay_and_collision(tmp_path,level3):
    data=dict(payload(),level3=level3)
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.services.validation_service.PowerShellRunner.run', side_effect=lambda command: ready_script(command) if 'FromBase64String' in command else PowerShellResult(0,'ok','')) as worker:
            check=client.post('/api/v1/access-jobs/validate',json={k:v for k,v in data.items() if k!='confirm_group_creation'})
            offer=check.json()['group_preparation']
            assert offer['group_name'].endswith(' Isolated')
            result=client.post('/api/v1/access-jobs/prepare-group',json=data)
        assert result.status_code==202, result.text
        body=result.json()
        assert [s['type'] for s in body['steps']]==['prepare_isolated_group','grant_isolated_base_acl','grant_isolated_base_acl']
        assert worker.call_count==5
        commands=''.join(c.args[0] for c in worker.call_args_list if 'FromBase64String' not in c.args[0])
        for forbidden in ['Add-ADGroupMember','robocopy','New-Item','/T','/C','(CI)','(OI)','@lsk']:
            assert forbidden not in commands
        assert "':RX'" in commands
        assert 'Start-Sleep -Seconds 15' in commands
        with patch('folder_grant_api.services.validation_service.PowerShellRunner.run') as reader:
            replay=client.post('/api/v1/access-jobs/prepare-group',json=data)
            assert replay.status_code==200 and replay.json()['replayed']
            reader.assert_not_called()
        assert client.post('/api/v1/access-jobs',json={k:v for k,v in data.items() if k!='confirm_group_creation'}).status_code==409
        assert client.post(f"/api/v1/access-jobs/{body['job_id']}/retry",json={'requested_by':'tester'}).status_code==409
        stored=client.get(f"/api/v1/access-jobs/{body['job_id']}").json()
        assert stored['result']['job_kind']=='group_preparation'


@pytest.mark.parametrize('changes',[
    {'group':{'status':'unknown','message':'denied'}},
    {'group':{'status':'failed','message':'보안 그룹이 아닙니다.'}},
    {'ou':{'status':'warning'}}, {'ou':{'status':'not_applicable'}}, {'user':{'status':'failed'}},
    {'preparation_root':{'status':'failed'}}, {'preparation_isolated':{'status':'failed'}},
    {'level3':{'status':'failed'}}, {'acl':{'status':'unknown'}},
    {'membership':{'status':'unknown','message':'other error'}},
])
def test_offer_exclusions_and_registration_gate(tmp_path,changes):
    data=payload()
    with build_client(tmp_path,executor_mode='powershell') as client, patch('folder_grant_api.services.validation_service.PowerShellRunner.run',side_effect=lambda s: ready_script(s,changes)):
        result=client.post('/api/v1/access-jobs/validate',json={k:v for k,v in data.items() if k!='confirm_group_creation'})
        assert result.json()['group_preparation'] is None
        assert client.post('/api/v1/access-jobs/prepare-group',json=data).status_code==409


@pytest.mark.parametrize('change',[{'confirm_group_creation':False},{'confirm_group_creation':None},{'operation':'revoke'},{'level2':'Study'},{'role':'invalid'}])
def test_bad_consent_or_scope_never_queries_or_mutates(tmp_path,change):
    with build_client(tmp_path,executor_mode='powershell') as client, patch('folder_grant_api.services.validation_service.PowerShellRunner.run') as runner:
        assert client.post('/api/v1/access-jobs/prepare-group',json=dict(payload(),**change)).status_code==422
        runner.assert_not_called()


def test_missing_consent_mock_and_auth(tmp_path):
    with build_client(tmp_path) as client, patch('folder_grant_api.services.validation_service.PowerShellRunner.run') as runner:
        data=payload();del data['confirm_group_creation']
        assert client.post('/api/v1/access-jobs/prepare-group',json=data).status_code==422
        assert client.post('/api/v1/access-jobs/prepare-group',json=payload()).status_code==422
        assert client.post('/api/v1/access-jobs/prepare-group',json=payload(),headers={'X-API-Key':'bad'}).status_code==401
        runner.assert_not_called()


def test_normal_request_id_cannot_be_reused_for_preparation(tmp_path):
    data=payload()
    with build_client(tmp_path,executor_mode='powershell') as client, patch('folder_grant_api.services.validation_service.PowerShellRunner.run',return_value=PowerShellResult(0,'ok','')):
        normal={k:v for k,v in data.items() if k!='confirm_group_creation'}
        assert client.post('/api/v1/access-jobs',json=normal).status_code==202
        with patch('folder_grant_api.services.validation_service.validate_access_resources') as validation:
            assert client.post('/api/v1/access-jobs/prepare-group',json=data).status_code==409
            validation.assert_not_called()
