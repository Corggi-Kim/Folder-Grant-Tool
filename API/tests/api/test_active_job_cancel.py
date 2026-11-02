import threading
from concurrent.futures import ThreadPoolExecutor
from unittest.mock import patch

import pytest
from sqlalchemy import select
from test_api import build_client, valid_payload, valid_project_payload
from folder_grant_api.infrastructure.models import AuditLogRecord, JobRecord
from folder_grant_api.infrastructure.powershell import PowerShellCancelled, PowerShellResult
from folder_grant_api.workers.powershell_worker import execute_powershell_job


@pytest.mark.parametrize('kind',['access','project'])
def test_authenticated_cancel_interrupts_running_step_and_preserves_prior_success(tmp_path,kind):
    endpoint='/api/v1/'+('access-jobs' if kind=='access' else 'project-jobs')
    route='access_jobs' if kind=='access' else 'project_jobs'
    with build_client(tmp_path,executor_mode='powershell') as client:
        data=valid_payload() if kind=='access' else valid_project_payload()
        with patch('folder_grant_api.api.routes.'+route+'.execute_job'):
            created=client.post(endpoint,json=data).json()
        job_id=created['job_id'];started=threading.Event();release=threading.Event();observed=[];calls=[]
        def run(runner,command):
            calls.append(command)
            if len(calls)==1: return PowerShellResult(0,'first step completed','')
            started.set()
            assert release.wait(5)
            observed.append(bool(runner.cancel_check and runner.cancel_check()))
            raise PowerShellCancelled()
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',autospec=True,side_effect=run):
            with ThreadPoolExecutor() as pool:
                future=pool.submit(execute_powershell_job,client.app.state.session_factory,client.app.state.settings,job_id)
                try:
                    assert started.wait(3)
                    denied=client.post(endpoint+'/'+job_id+'/cancel',json={'requested_by':'operator'},headers={'X-API-Key':'wrong'})
                    assert denied.status_code==401
                    response=client.post(endpoint+'/'+job_id+'/cancel',json={'requested_by':'operator'})
                    assert response.status_code==200
                finally:
                    release.set()
                future.result(timeout=5)
        final=client.get(endpoint+'/'+job_id).json()
        assert observed==[True]
        assert final['status']=='cancelled'
        assert final['cancel_requested']
        assert final['steps'][0]['status']=='succeeded'
        assert final['steps'][1]['status']=='cancelled'
        assert all(s['status']=='cancelled' for s in final['steps'][2:])
        assert len(calls)==2
        assert final['result']['partial_changes_possible']
        with client.app.state.session_factory() as db:
            events=list(db.scalars(select(AuditLogRecord.event_type).where(AuditLogRecord.job_id==job_id)))
        assert 'job.cancelled' in events
        assert 'job.succeeded' not in events


def test_cancel_during_final_step_is_not_overwritten_by_success(tmp_path):
    with build_client(tmp_path,executor_mode='powershell') as client:
        data=valid_payload();data['role']=''
        with patch('folder_grant_api.api.routes.access_jobs.execute_job'):
            job=client.post('/api/v1/access-jobs',json=data).json()
        calls=[]
        def run(command):
            calls.append(command)
            if len(calls)==len(job['steps']):
                response=client.post('/api/v1/access-jobs/'+job['job_id']+'/cancel',json={'requested_by':'operator'})
                assert response.status_code==200
            return PowerShellResult(0,'completed','')
        with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',side_effect=run):
            execute_powershell_job(client.app.state.session_factory,client.app.state.settings,job['job_id'])
        final=client.get('/api/v1/access-jobs/'+job['job_id']).json()
        assert final['status']=='cancelled'
        assert all(s['status']=='succeeded' for s in final['steps'])


@pytest.mark.skipif(__import__('os').name=='nt',reason='POSIX executable fixture; Windows Job Object tested separately')
@pytest.mark.parametrize('kind',['access','project','preparation'])
def test_cancel_endpoint_stops_real_worker_process_tree(tmp_path,fake_powershell,kind):
    import time
    from types import SimpleNamespace
    executable,marker=fake_powershell
    route='project_jobs' if kind=='project' else 'access_jobs'
    endpoint='/api/v1/'+('project-jobs' if kind=='project' else 'access-jobs')
    with build_client(tmp_path,executor_mode='powershell') as client:
        client.app.state.settings.powershell_path=str(executable)
        client.app.state.settings.powershell_timeout_seconds=10
        data=valid_project_payload() if kind=='project' else valid_payload()
        register=endpoint
        if kind=='preparation':
            data.update(level2='Isolated',role=None,confirm_group_creation=True)
            register+='/prepare-group'
        ready=SimpleNamespace(group_preparation=SimpleNamespace(group_name='LSK 26-012 Isolated',root_path='ROOT',isolated_path='ISOLATED'))
        with patch('folder_grant_api.api.routes.'+route+'.execute_job'),patch('folder_grant_api.api.routes.access_jobs.validate_access_resources',return_value=ready,create=True):
            # prepare-group imports its validator locally.
            with patch('folder_grant_api.services.validation_service.validate_access_resources',return_value=ready):
                response=client.post(register,json=data)
        assert response.status_code==202,response.text
        job_id=response.json()['job_id']
        with ThreadPoolExecutor() as pool:
            future=pool.submit(execute_powershell_job,client.app.state.session_factory,client.app.state.settings,job_id)
            try:
                deadline=time.monotonic()+5
                while not marker.exists() and time.monotonic()<deadline: threading.Event().wait(.02)
                assert marker.exists()
                before=time.monotonic()
                assert client.post(endpoint+'/'+job_id+'/cancel',json={'requested_by':'operator'}).status_code==200
                future.result(timeout=5)
                assert time.monotonic()-before<5
            finally:
                client.post(endpoint+'/'+job_id+'/cancel',json={'requested_by':'cleanup'})
        final=client.get(endpoint+'/'+job_id).json()
        assert final['status']=='cancelled'
        assert all(s['status']=='cancelled' for s in final['steps'])
        last=marker.read_text();threading.Event().wait(.15)
        assert marker.read_text()==last
        if kind=='preparation': assert final['result']['job_kind']=='group_preparation'


def test_cancel_wins_when_committed_immediately_before_success_update(tmp_path):
    from sqlalchemy import event
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.api.routes.access_jobs.execute_job'):
            job=client.post('/api/v1/access-jobs',json=valid_payload()).json()
        injected=[]
        def before_update(conn,cursor,statement,parameters,context,executemany):
            if statement.startswith('UPDATE jobs SET') and 'succeeded' in parameters and not injected:
                injected.append(True)
                response=client.post('/api/v1/access-jobs/'+job['job_id']+'/cancel',json={'requested_by':'race'})
                assert response.status_code==200
        engine=client.app.state.database_engine
        event.listen(engine,'before_cursor_execute',before_update)
        try:
            with patch('folder_grant_api.workers.powershell_worker.PowerShellRunner.run',return_value=PowerShellResult(0,'OK','')):
                execute_powershell_job(client.app.state.session_factory,client.app.state.settings,job['job_id'])
        finally:
            event.remove(engine,'before_cursor_execute',before_update)
        final=client.get('/api/v1/access-jobs/'+job['job_id']).json()
        assert injected==[True]
        assert final['status']=='cancelled'
        assert final['cancel_requested']


def test_stale_cancel_cannot_overwrite_already_completed_job(tmp_path):
    from sqlalchemy import update
    from folder_grant_api.infrastructure.job_repository import JobRepository
    with build_client(tmp_path,executor_mode='powershell') as client:
        with patch('folder_grant_api.api.routes.access_jobs.execute_job'):
            job=client.post('/api/v1/access-jobs',json=valid_payload()).json()
        with client.app.state.session_factory() as stale:
            repository=JobRepository(stale);loaded=repository.get(job['job_id'])
            with client.app.state.session_factory() as winner:
                winner.execute(update(JobRecord).where(JobRecord.id==job['job_id']).values(status='succeeded'))
                winner.commit()
            repository.request_cancel(loaded,'late operator')
        final=client.get('/api/v1/access-jobs/'+job['job_id']).json()
        assert final['status']=='succeeded'
        assert not final['cancel_requested']
