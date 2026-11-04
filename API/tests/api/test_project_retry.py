import os
import shutil
from unittest.mock import patch

import pytest

from test_api import build_client, valid_payload, valid_project_payload
from folder_grant_api.infrastructure.models import JobRecord
from folder_grant_api.infrastructure.powershell import PowerShellCancelled, PowerShellExecutionError, PowerShellResult


@pytest.mark.parametrize('failure,first_success,expected_status', [
    (PowerShellExecutionError('POWERSHELL_FAILED', 'group failed'),False,'failed'),
    (PowerShellExecutionError('POWERSHELL_FAILED', 'ACL failed'),True,'partially_succeeded'),
    (PowerShellCancelled(),True,'cancelled'),
])
def test_project_retry_preserves_history_successes_and_original_request(tmp_path,failure,first_success,expected_status):
    payload = valid_project_payload()
    with build_client(tmp_path,executor_mode='powershell') as client:
        outcomes = [PowerShellResult(0,'group exists',''),failure] if first_success else [failure]
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',side_effect=outcomes):
            old = client.post('/api/v1/project-jobs',json=payload).json()
        endpoint = '/api/v1/project-jobs/'+old['job_id']+'/retry'
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',return_value=PowerShellResult(0,'recovered','')):
            retry = client.post(endpoint,json={'requested_by':'operator'})
        assert retry.status_code == 202
        new = retry.json()
        assert new['job_id'] != old['job_id']
        assert new['original_request_id'] == payload['request_id']
        assert new['project_code'] == payload['project_code']
        assert new['project_name'] == payload['project_name']
        replay = client.post(endpoint,json={'requested_by':'operator'})
        assert replay.status_code == 200 and replay.json()['job_id'] == new['job_id']
        assert client.get('/api/v1/project-jobs/'+old['job_id']).json()['status'] == expected_status
        latest = client.get('/api/v1/project-jobs/by-request/'+payload['request_id']).json()
        assert latest['job_id'] == new['job_id'] and latest['status'] == 'succeeded'
        if first_success:
            assert latest['steps'][0]['result']['reused_from_job_id'] == old['job_id']
            assert latest['steps'][0]['result']['stdout'] == 'group exists'
        repeated = client.post('/api/v1/project-jobs',json=payload)
        assert repeated.status_code == 200 and repeated.json()['job_id'] == new['job_id']
        changed = {**payload,'project_name':'different'}
        assert client.post('/api/v1/project-jobs',json=changed).status_code == 409
        assert client.post('/api/v1/project-jobs/'+new['job_id']+'/retry',json={'requested_by':'operator'}).status_code == 409


def test_queued_project_cancellation_can_retry_without_execution_result(tmp_path):
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.api.routes.project_jobs.execute_job'):
            old = client.post('/api/v1/project-jobs',json=valid_project_payload()).json()
        endpoint = '/api/v1/project-jobs/'+old['job_id']
        assert client.post(endpoint+'/retry',json={'requested_by':'operator'}).status_code == 409
        assert client.post(endpoint+'/cancel',json={'requested_by':'operator'}).json()['status'] == 'cancelled'
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',return_value=PowerShellResult(0,'done','')):
            retry = client.post(endpoint+'/retry',json={'requested_by':'operator'})
        assert retry.status_code == 202 and retry.json()['job_id'] != old['job_id']


def test_project_retry_auth_validation_kind_and_mock_rejection(tmp_path):
    with build_client(tmp_path) as client:
        old = client.post('/api/v1/project-jobs',json=valid_project_payload()).json()
        endpoint = '/api/v1/project-jobs/'+old['job_id']+'/retry'
        assert client.post(endpoint,json={'requested_by':'operator'},headers={'X-API-Key':'wrong'}).status_code == 401
        assert client.post(endpoint,json={}).status_code == 422
        assert client.post(endpoint,json={'requested_by':'operator'}).status_code == 409
        access = client.post('/api/v1/access-jobs',json=valid_payload()).json()
        wrong_kind = client.post('/api/v1/project-jobs/'+access['job_id']+'/retry',json={'requested_by':'operator'})
        assert wrong_kind.status_code == 404
        assert client.post('/api/v1/project-jobs/missing/retry',json={'requested_by':'operator'}).status_code == 404


@pytest.mark.parametrize('status,error_code,result_mode', [
    ('running',None,'powershell'),('cancel_requested',None,'powershell'),
    ('failed','POWERSHELL_STOP_FAILED','powershell'),
    ('cancelled','POWERSHELL_STOP_FAILED','powershell'),
    ('failed',None,'mock'),('succeeded',None,'powershell'),
])
def test_project_retry_rejects_unsafe_or_completed_attempts(tmp_path,status,error_code,result_mode):
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.api.routes.project_jobs.execute_job'):
            old = client.post('/api/v1/project-jobs',json=valid_project_payload()).json()
        with client.app.state.session_factory() as session:
            job = session.get(JobRecord,old['job_id'])
            job.status = status
            job.result_payload = {'executor_mode':result_mode}
            job.error_code = error_code
            session.commit()
        retry = client.post('/api/v1/project-jobs/'+old['job_id']+'/retry',json={'requested_by':'operator'})
        assert retry.status_code == 409
        assert retry.json()['detail']['code'] == 'JOB_NOT_RETRYABLE'


def test_project_retry_chain_returns_latest_after_restart(tmp_path):
    payload = valid_project_payload()
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',side_effect=PowerShellCancelled()):
            old = client.post('/api/v1/project-jobs',json=payload).json()
            first = client.post('/api/v1/project-jobs/'+old['job_id']+'/retry',json={'requested_by':'operator'}).json()
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',return_value=PowerShellResult(0,'done','')):
            second = client.post('/api/v1/project-jobs/'+first['job_id']+'/retry',json={'requested_by':'operator'}).json()
        assert second['job_id'] != first['job_id']
    with build_client(tmp_path,executor_mode='powershell') as client:
        stale = client.post('/api/v1/project-jobs/'+old['job_id']+'/retry',json={'requested_by':'operator'})
        assert stale.status_code == 200 and stale.json()['job_id'] == second['job_id']
        assert stale.json()['original_request_id'] == payload['request_id']
        assert client.get('/api/v1/project-jobs/by-request/'+payload['request_id']).json()['job_id'] == second['job_id']


def test_incomplete_template_copy_marker_survives_queued_retry_cancellation(tmp_path):
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',side_effect=[
            PowerShellResult(0,'group ready',''),PowerShellCancelled()
        ]):
            original = client.post('/api/v1/project-jobs',json=valid_project_payload()).json()
        with patch('folder_grant_api.api.routes.project_jobs.execute_job'):
            first = client.post('/api/v1/project-jobs/'+original['job_id']+'/retry',json={'requested_by':'operator'}).json()
        assert first['steps'][1]['details']['_template_copy_incomplete'] is True
        assert client.post('/api/v1/project-jobs/'+first['job_id']+'/cancel',json={'requested_by':'operator'}).json()['status'] == 'cancelled'
        with patch('folder_grant_api.api.routes.project_jobs.execute_job'):
            second = client.post('/api/v1/project-jobs/'+first['job_id']+'/retry',json={'requested_by':'operator'})
        assert second.status_code == 202
        assert second.json()['steps'][1]['details']['_template_copy_incomplete'] is True


def test_successful_template_copy_clears_incomplete_marker_on_full_restart(tmp_path):
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',side_effect=[
            PowerShellResult(0,'group ready',''),PowerShellCancelled()
        ]):
            original = client.post('/api/v1/project-jobs',json=valid_project_payload()).json()
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',return_value=PowerShellResult(0,'root recreated','')):
            recovered = client.post('/api/v1/project-jobs/'+original['job_id']+'/retry',json={'requested_by':'operator'})
        assert recovered.status_code == 202
        assert recovered.json()['steps'][1]['details']['_template_copy_incomplete'] is True
        with patch('folder_grant_api.api.routes.project_jobs.execute_job'):
            restart = client.post('/api/v1/project-jobs/'+recovered.json()['job_id']+'/retry',json={'requested_by':'operator','restart_all':True})
        assert restart.status_code == 202
        assert '_template_copy_incomplete' not in restart.json()['steps'][1]['details']


def test_interrupted_template_retry_surfaces_repair_error_and_preserves_existing_files(tmp_path,monkeypatch):
    executable = os.environ.get('FGT_TEST_POWERSHELL') or shutil.which('pwsh')
    if not executable:
        pytest.skip('Set FGT_TEST_POWERSHELL for actual interpreter tests')
    for key in ('XDG_CACHE_HOME','XDG_CONFIG_HOME','XDG_DATA_HOME'):
        monkeypatch.setenv(key,str(tmp_path))
    root = tmp_path/'project'
    root.mkdir()
    existing = root/'existing.txt'
    existing.write_text('operator content')
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',side_effect=[
            PowerShellResult(0,'group ready',''),PowerShellCancelled()
        ]):
            original = client.post('/api/v1/project-jobs',json=valid_project_payload()).json()
        with client.app.state.session_factory() as session:
            job = session.get(JobRecord,original['job_id'])
            job.steps[1].target = str(root)
            session.commit()
        client.app.state.settings.powershell_path = executable
        retry = client.post('/api/v1/project-jobs/'+original['job_id']+'/retry',json={'requested_by':'operator'})
        assert retry.status_code == 202
        failed = client.get('/api/v1/project-jobs/'+retry.json()['job_id']).json()
        assert failed['status'] == 'partially_succeeded'
        assert failed['error_code'] == 'PROJECT_TEMPLATE_REPAIR_REQUIRED'
        assert '기존 폴더 내용을 확인·복구해야 하므로 자동 복사를 중단했습니다.' in failed['error_message']
        assert all(step['status'] == 'skipped' for step in failed['steps'][2:])
        assert existing.read_text() == 'operator content'
        assert list(root.iterdir()) == [existing]
