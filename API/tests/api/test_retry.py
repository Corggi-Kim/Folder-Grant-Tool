from unittest.mock import patch
from test_api import build_client, valid_payload
from folder_grant_api.infrastructure.powershell import PowerShellExecutionError, PowerShellResult


def fail_acl(client):
    payload = valid_payload(); payload['role'] = ''
    with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',
               side_effect=[PowerShellResult(0,'group done',''),PowerShellExecutionError('POWERSHELL_FAILED','missing folder')]):
        job = client.post('/api/v1/access-jobs',json=payload).json()
    failed = client.get('/api/v1/access-jobs/'+job['job_id']).json()
    assert failed['status']=='partially_succeeded'
    return payload,failed


def test_retry_creates_one_new_attempt_keeps_history_and_recovers_by_original_request(tmp_path):
    with build_client(tmp_path,executor_mode='powershell') as client:
        payload,old = fail_acl(client)
        endpoint = '/api/v1/access-jobs/'+old['job_id']+'/retry'
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',return_value=PowerShellResult(0,'folder fixed','')) as run:
            response = client.post(endpoint,json={'requested_by':'retry-operator'})
            assert response.status_code==202
            new = response.json()
            assert new['job_id']!=old['job_id']
            assert new['original_request_id']==payload['request_id']
            assert run.call_count==1
            replay = client.post(endpoint,json={'requested_by':'retry-operator'})
            assert replay.status_code==200
            assert replay.json()['job_id']==new['job_id']
            assert run.call_count==1
        assert client.get('/api/v1/access-jobs/'+old['job_id']).json()['status']=='partially_succeeded'
        latest = client.get('/api/v1/access-jobs/by-request/'+payload['request_id']).json()
        assert latest['job_id']==new['job_id']
        assert latest['status']=='succeeded'
        assert all(s['status']=='succeeded' for s in latest['steps'])
        assert latest['steps'][0]['result']['reused_from_job_id']==old['job_id']
        repeated = client.post('/api/v1/access-jobs',json=payload).json()
        assert repeated['job_id']==new['job_id'] and repeated['replayed']
        assert client.post('/api/v1/access-jobs/'+new['job_id']+'/retry',json={'requested_by':'operator'}).status_code==409


def test_retry_requires_auth_and_rejects_mock_jobs(tmp_path):
    with build_client(tmp_path) as client:
        job = client.post('/api/v1/access-jobs',json=valid_payload()).json()
        endpoint = '/api/v1/access-jobs/'+job['job_id']+'/retry'
        assert client.post(endpoint,json={'requested_by':'operator'},headers={'X-API-Key':'wrong'}).status_code==401
        assert client.post(endpoint,json={'requested_by':'operator'}).status_code==409
        assert client.post(endpoint,json={}).status_code==422


def test_retry_chain_survives_restart_and_stale_retry_request_returns_latest_attempt(tmp_path):
    with build_client(tmp_path,executor_mode='powershell') as client:
        payload,old = fail_acl(client)
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',side_effect=PowerShellExecutionError('POWERSHELL_FAILED','still missing')):
            first = client.post('/api/v1/access-jobs/'+old['job_id']+'/retry',json={'requested_by':'operator'}).json()
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',return_value=PowerShellResult(0,'fixed','')):
            second = client.post('/api/v1/access-jobs/'+first['job_id']+'/retry',json={'requested_by':'operator'}).json()
        replay = client.post('/api/v1/access-jobs/'+old['job_id']+'/retry',json={'requested_by':'operator'}).json()
        assert replay['job_id']==second['job_id']
    with build_client(tmp_path,executor_mode='powershell') as client:
        latest = client.get('/api/v1/access-jobs/by-request/'+payload['request_id']).json()
        assert latest['job_id']==second['job_id'] and latest['status']=='succeeded'


def test_queued_and_cancelled_jobs_cannot_be_retried(tmp_path):
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.api.routes.access_jobs.execute_job'):
            job = client.post('/api/v1/access-jobs',json=valid_payload()).json()
        endpoint = '/api/v1/access-jobs/'+job['job_id']+'/retry'
        assert client.post(endpoint,json={'requested_by':'operator'}).status_code==409
        client.post('/api/v1/access-jobs/'+job['job_id']+'/cancel',json={'requested_by':'operator'})
        assert client.post(endpoint,json={'requested_by':'operator'}).status_code==409
