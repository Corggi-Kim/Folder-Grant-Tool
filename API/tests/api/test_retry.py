from unittest.mock import patch
import pytest
from test_api import build_client, valid_payload
from folder_grant_api.infrastructure.models import JobRecord, utc_now
from folder_grant_api.infrastructure.powershell import PowerShellCancelled, PowerShellExecutionError, PowerShellResult


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


def test_queued_job_must_be_cancelled_before_retry_and_preserves_original(tmp_path):
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.api.routes.access_jobs.execute_job'):
            job = client.post('/api/v1/access-jobs',json=valid_payload()).json()
        endpoint = '/api/v1/access-jobs/'+job['job_id']+'/retry'
        assert client.post(endpoint,json={'requested_by':'operator'}).status_code==409
        cancelled = client.post('/api/v1/access-jobs/'+job['job_id']+'/cancel',json={'requested_by':'operator'}).json()
        assert cancelled['status'] == 'cancelled'
        assert cancelled['result'] is None and cancelled['started_at'] is None
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run', return_value=PowerShellResult(0, 'done', '')):
            retry = client.post(endpoint,json={'requested_by':'operator'})
        assert retry.status_code == 202
        assert retry.json()['job_id'] != job['job_id']
        assert client.get('/api/v1/access-jobs/'+job['job_id']).json()['status'] == 'cancelled'
        assert client.post(endpoint,json={'requested_by':'operator'}).json()['job_id'] == retry.json()['job_id']
        assert client.post('/api/v1/access-jobs',json=valid_payload()).json()['job_id'] == retry.json()['job_id']


def test_cancelled_access_retry_reuses_successful_steps(tmp_path):
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run', side_effect=[
            PowerShellResult(0, 'membership complete', ''), PowerShellCancelled()
        ]):
            old = client.post('/api/v1/access-jobs',json=valid_payload()).json()
        assert client.get('/api/v1/access-jobs/'+old['job_id']).json()['status'] == 'cancelled'
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',return_value=PowerShellResult(0,'recovered','')):
            retry = client.post('/api/v1/access-jobs/'+old['job_id']+'/retry',json={'requested_by':'operator'})
        assert retry.status_code == 202
        latest = client.get('/api/v1/access-jobs/by-request/TEST-001').json()
        assert latest['status'] == 'succeeded'
        assert latest['steps'][0]['result']['stdout'] == 'membership complete'
        assert latest['steps'][0]['result']['reused_from_job_id'] == old['job_id']


@pytest.mark.parametrize('status,result_mode,started,error_code', [
    ('running','powershell',True,None),
    ('cancel_requested','powershell',True,None),
    ('failed','powershell',True,'POWERSHELL_STOP_FAILED'),
    ('cancelled','powershell',True,'POWERSHELL_STOP_FAILED'),
    ('cancelled',None,True,None),
    ('failed','mock',True,None),
    ('succeeded','powershell',True,None),
    ('simulated','mock',True,None),
])
def test_access_retry_rejects_unsafe_or_completed_attempts(tmp_path,status,result_mode,started,error_code):
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.api.routes.access_jobs.execute_job'):
            old = client.post('/api/v1/access-jobs',json=valid_payload()).json()
        with client.app.state.session_factory() as session:
            job = session.get(JobRecord,old['job_id'])
            job.status = status
            job.started_at = utc_now() if started else None
            job.result_payload = {'executor_mode':result_mode} if result_mode else None
            job.error_code = error_code
            session.commit()
        response = client.post('/api/v1/access-jobs/'+old['job_id']+'/retry',json={'requested_by':'operator'})
        assert response.status_code == 409
        assert response.json()['detail']['code'] == 'JOB_NOT_RETRYABLE'


@pytest.mark.parametrize('job_kind', ['access', 'project'])
def test_retry_rejects_unconfirmed_step_stop_and_current_mock_mode(tmp_path,job_kind):
    from test_api import valid_project_payload
    payload = valid_payload() if job_kind == 'access' else valid_project_payload()
    path = '/api/v1/'+job_kind+'-jobs'
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.api.routes.'+job_kind+'_jobs.execute_job'):
            old = client.post(path,json=payload).json()
        with client.app.state.session_factory() as session:
            job = session.get(JobRecord,old['job_id'])
            job.status = 'failed'
            job.result_payload = {'executor_mode':'powershell'}
            job.steps[0].error_code = 'POWERSHELL_STOP_FAILED'
            session.commit()
        assert client.post(path+'/'+old['job_id']+'/retry',json={'requested_by':'operator'}).status_code == 409
        with client.app.state.session_factory() as session:
            job = session.get(JobRecord,old['job_id'])
            job.steps[0].error_code = None
            session.commit()
    with build_client(tmp_path) as client:
        assert client.post(path+'/'+old['job_id']+'/retry',json={'requested_by':'operator'}).status_code == 409


@pytest.mark.parametrize('restart_all', [False, True])
def test_retry_does_not_reexecute_group_preparation(tmp_path,restart_all):
    with build_client(tmp_path,executor_mode='powershell') as client:
        payload,old = fail_acl(client)
        with client.app.state.session_factory() as session:
            job = session.get(JobRecord,old['job_id'])
            job.request_payload = {**job.request_payload,'confirm_group_creation':True}
            session.commit()
        assert client.post('/api/v1/access-jobs/'+old['job_id']+'/retry',json={'requested_by':'operator','restart_all':restart_all}).status_code == 409


@pytest.mark.parametrize('job_kind', ['access', 'project'])
def test_successful_job_requires_explicit_full_restart_and_queues_every_step(tmp_path,job_kind):
    from test_api import valid_project_payload
    payload = valid_payload() if job_kind == 'access' else valid_project_payload()
    path = '/api/v1/'+job_kind+'-jobs'
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',return_value=PowerShellResult(0,'first execution','')):
            original = client.post(path,json=payload).json()
        old = client.get(path+'/'+original['job_id']).json()
        assert old['status'] == 'succeeded'
        assert client.post(path,json=payload).json()['job_id'] == old['job_id']
        endpoint = path+'/'+old['job_id']+'/retry'
        assert client.post(endpoint,json={'requested_by':'operator'}).status_code == 409
        with patch('folder_grant_api.api.routes.'+job_kind+'_jobs.execute_job'):
            response = client.post(endpoint,json={'requested_by':'operator','restart_all':True})
            assert response.status_code == 202
            new = response.json()
            assert new['job_id'] != old['job_id']
            assert new['original_request_id'] == payload['request_id']
            assert all(step['status'] == 'queued' and step['result'] is None for step in new['steps'])
            assert all(step['started_at'] is None and step['finished_at'] is None for step in new['steps'])
            replay = client.post(endpoint,json={'requested_by':'operator','restart_all':True})
            assert replay.status_code == 200 and replay.json()['job_id'] == new['job_id']
        assert client.get(path+'/'+old['job_id']).json() == old
        assert client.get(path+'/by-request/'+payload['request_id']).json()['job_id'] == new['job_id']


@pytest.mark.parametrize('job_kind', ['access', 'project'])
@pytest.mark.parametrize('first_restart_all', [False, True])
def test_retry_child_mode_conflict_is_rejected(tmp_path,job_kind,first_restart_all):
    from test_api import valid_project_payload
    payload = valid_payload() if job_kind == 'access' else valid_project_payload()
    path = '/api/v1/'+job_kind+'-jobs'
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',side_effect=PowerShellExecutionError('POWERSHELL_FAILED','failed')):
            old = client.post(path,json=payload).json()
        endpoint = path+'/'+old['job_id']+'/retry'
        with patch('folder_grant_api.api.routes.'+job_kind+'_jobs.execute_job'):
            first = client.post(endpoint,json={'requested_by':'operator','restart_all':first_restart_all})
        assert first.status_code == 202
        replay = client.post(endpoint,json={'requested_by':'operator','restart_all':first_restart_all})
        assert replay.status_code == 200 and replay.json()['job_id'] == first.json()['job_id']
        conflict = client.post(endpoint,json={'requested_by':'operator','restart_all':not first_restart_all})
        assert conflict.status_code == 409
        assert conflict.json()['detail']['code'] == 'REQUEST_ID_CONFLICT'
        assert client.get(path+'/by-request/'+payload['request_id']).json()['job_id'] == first.json()['job_id']


@pytest.mark.parametrize('job_kind', ['access', 'project'])
def test_normal_retry_after_full_restart_does_not_inherit_full_restart_marker(tmp_path,job_kind):
    from test_api import valid_project_payload
    payload = valid_payload() if job_kind == 'access' else valid_project_payload()
    path = '/api/v1/'+job_kind+'-jobs'
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',return_value=PowerShellResult(0,'complete','')):
            old = client.post(path,json=payload).json()
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',side_effect=[
            PowerShellResult(0,'restarted group',''),PowerShellExecutionError('POWERSHELL_FAILED','failed')
        ]):
            first = client.post(path+'/'+old['job_id']+'/retry',json={'requested_by':'operator','restart_all':True})
        assert first.status_code == 202
        endpoint = path+'/'+first.json()['job_id']+'/retry'
        with patch('folder_grant_api.api.routes.'+job_kind+'_jobs.execute_job'):
            second = client.post(endpoint,json={'requested_by':'operator'})
        assert second.status_code == 202
        assert second.json()['steps'][0]['status'] == 'succeeded'
        assert second.json()['steps'][0]['result']['stdout'] == 'restarted group'
        assert second.json()['steps'][0]['result']['reused_from_job_id'] == first.json()['job_id']
        assert client.post(endpoint,json={'requested_by':'operator'}).status_code == 200
        assert client.post(endpoint,json={'requested_by':'operator','restart_all':True}).status_code == 409


@pytest.mark.parametrize('job_kind', ['access', 'project'])
@pytest.mark.parametrize('status,error_code,mode', [
    ('queued',None,'powershell'),('running',None,'powershell'),
    ('cancel_requested',None,'powershell'),('failed','POWERSHELL_STOP_FAILED','powershell'),
    ('succeeded',None,'mock'),('simulated',None,'mock'),
])
def test_restart_all_keeps_execution_safety_guards(tmp_path,job_kind,status,error_code,mode):
    from test_api import valid_project_payload
    path = '/api/v1/'+job_kind+'-jobs'
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.api.routes.'+job_kind+'_jobs.execute_job'):
            old = client.post(path,json=valid_payload() if job_kind == 'access' else valid_project_payload()).json()
        with client.app.state.session_factory() as session:
            job = session.get(JobRecord,old['job_id'])
            job.status = status
            job.error_code = error_code
            job.result_payload = {'executor_mode':mode}
            session.commit()
        response = client.post(path+'/'+old['job_id']+'/retry',json={'requested_by':'operator','restart_all':True})
        assert response.status_code == 409
        assert response.json()['detail']['code'] == 'JOB_NOT_RETRYABLE'


@pytest.mark.parametrize('job_kind', ['access', 'project'])
def test_cancelling_queued_retry_preserves_successful_steps_for_next_retry(tmp_path,job_kind):
    from test_api import valid_project_payload
    payload = valid_payload() if job_kind == 'access' else valid_project_payload()
    path = '/api/v1/'+job_kind+'-jobs'
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',side_effect=[
            PowerShellResult(0,'membership preserved',''),PowerShellExecutionError('POWERSHELL_FAILED','failed')
        ]):
            original = client.post(path,json=payload).json()
        with patch('folder_grant_api.api.routes.'+job_kind+'_jobs.execute_job'):
            first = client.post(path+'/'+original['job_id']+'/retry',json={'requested_by':'operator'}).json()
        assert first['steps'][0]['status'] == 'succeeded'
        cancelled = client.post(path+'/'+first['job_id']+'/cancel',json={'requested_by':'operator'}).json()
        assert cancelled['status'] == 'cancelled' and cancelled['result'] is None
        assert cancelled['steps'][0] == first['steps'][0]
        assert all(step['status'] == 'cancelled' for step in cancelled['steps'][1:])
        with patch('folder_grant_api.api.routes.'+job_kind+'_jobs.execute_job'):
            second = client.post(path+'/'+first['job_id']+'/retry',json={'requested_by':'operator'})
        assert second.status_code == 202
        assert second.json()['steps'][0]['status'] == 'succeeded'
        assert second.json()['steps'][0]['result']['stdout'] == 'membership preserved'
        assert second.json()['steps'][0]['result']['reused_from_job_id'] == first['job_id']
        assert all(step['status'] == 'queued' for step in second.json()['steps'][1:])
        assert client.get(path+'/'+original['job_id']).json()['status'] == 'partially_succeeded'
        assert client.get(path+'/'+first['job_id']).json()['steps'][0] == first['steps'][0]


@pytest.mark.parametrize('job_kind', ['access', 'project'])
def test_cancelled_mixed_steps_without_trusted_retry_parent_cannot_retry(tmp_path,job_kind):
    from test_api import valid_project_payload
    path = '/api/v1/'+job_kind+'-jobs'
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.api.routes.'+job_kind+'_jobs.execute_job'):
            old = client.post(path,json=valid_payload() if job_kind == 'access' else valid_project_payload()).json()
        with client.app.state.session_factory() as session:
            job = session.get(JobRecord,old['job_id'])
            job.status = 'cancelled'
            job.request_id = 'RETRY-missing-parent'
            job.request_payload = {**job.request_payload,'_retry_of_job_id':'missing-parent',
                                   '_original_request_id':old['request_id']}
            for step in job.steps:
                step.status = 'cancelled'
            job.steps[0].status = 'succeeded'
            job.steps[0].result = {'executor_mode':'powershell','reused_from_job_id':'missing-parent'}
            session.commit()
        retry = client.post(path+'/'+old['job_id']+'/retry',json={'requested_by':'operator'})
        assert retry.status_code == 409


@pytest.mark.parametrize('job_kind', ['access', 'project'])
def test_cancelled_queued_retry_requires_powershell_evidence_for_reused_success(tmp_path,job_kind):
    from test_api import valid_project_payload
    path = '/api/v1/'+job_kind+'-jobs'
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',side_effect=[
            PowerShellResult(0,'membership complete',''),PowerShellExecutionError('POWERSHELL_FAILED','failed')
        ]):
            original = client.post(path,json=valid_payload() if job_kind == 'access' else valid_project_payload()).json()
        with patch('folder_grant_api.api.routes.'+job_kind+'_jobs.execute_job'):
            child = client.post(path+'/'+original['job_id']+'/retry',json={'requested_by':'operator'}).json()
        client.post(path+'/'+child['job_id']+'/cancel',json={'requested_by':'operator'})
        with client.app.state.session_factory() as session:
            job = session.get(JobRecord,child['job_id'])
            job.steps[0].result = {**job.steps[0].result,'executor_mode':'mock'}
            session.commit()
        retry = client.post(path+'/'+child['job_id']+'/retry',json={'requested_by':'operator'})
        assert retry.status_code == 409
