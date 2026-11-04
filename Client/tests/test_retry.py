import unittest
from unittest.mock import Mock, patch
import test_ui_design as design


class RetryTests(unittest.TestCase):
    setUp = design.DesignTests.setUp

    def failed_row(self):
        return self.fixture.row(result=dict(job_id='failed-job',status='partially_succeeded',executor_mode='powershell'))

    def test_selected_execute_confirms_retry_and_keeps_bus_request_identity(self):
        w = self.window; row = self.failed_row(); request_id = w._row_metadata(row)['request_id']
        w._start_next_job = Mock()
        with patch.object(self.m.QMessageBox,'question',return_value=self.m.QMessageBox.Yes):
            w.run_execute()
        self.assertEqual(w.run_queue[0][3],'failed-job')
        self.assertEqual(w.run_queue[0][1]['request_id'],request_id)
        self.assertEqual(w._row_metadata(row)['source'],'bus')
        w._set_running_ui(False)

    def test_declining_confirmation_does_not_register_or_clear_failure(self):
        w = self.window; row = self.failed_row(); w._start_next_job = Mock()
        with patch.object(self.m.QMessageBox,'question',return_value=self.m.QMessageBox.No):
            w.run_execute()
        w._start_next_job.assert_not_called()
        self.assertEqual(w._row_metadata(row)['result']['job_id'],'failed-job')

    def test_worker_submits_retry_endpoint_and_polls_new_job_only(self):
        api = Mock(); api.request.side_effect = [dict(job_id='new-job',status='queued'),dict(job_id='new-job',status='succeeded',executor_mode='powershell')]
        worker = self.m.ApiJobWorker(api,'access',dict(request_id='BUS-original',requested_by='operator'),retry_job_id='failed-job')
        worker.poll_interval = .001
        self.assertEqual(worker.execute()['job_id'],'new-job')
        self.assertEqual(api.request.call_args_list[0].args,('POST','/api/v1/access-jobs/failed-job/retry',{'requested_by':'operator'}))
        self.assertEqual(api.request.call_args_list[1].args,('GET','/api/v1/access-jobs/new-job'))

    def test_lost_retry_response_recovers_without_creating_another_job(self):
        api = Mock(); api.request.side_effect = [self.m.ApiClientError('lost',503),dict(job_id='new-job',status='succeeded',executor_mode='powershell')]
        worker = self.m.ApiJobWorker(api,'access',dict(request_id='BUS-original',requested_by='operator'),retry_job_id='failed-job')
        self.assertEqual(worker.execute()['job_id'],'new-job')
        self.assertEqual(api.request.call_args_list[1].args,('GET','/api/v1/access-jobs/by-request/BUS-original'))
        self.assertEqual(sum(call.args[0]=='POST' for call in api.request.call_args_list),1)

    def test_retry_recovery_rejects_original_failed_job_as_unconfirmed(self):
        api = Mock(); api.request.side_effect = [self.m.ApiClientError('lost',503),dict(job_id='failed-job',status='partially_succeeded')]
        worker = self.m.ApiJobWorker(api,'access',dict(request_id='BUS-original',requested_by='operator'),retry_job_id='failed-job')
        with self.assertRaises(self.m.ApiClientError): worker.execute()

    def test_retry_preview_never_submits_a_retry_or_real_job(self):
        api = Mock(); api.request.return_value = {'valid':True}
        worker = self.m.ApiJobWorker(api,'access',dict(request_id='BUS-original',requested_by='operator'),True,retry_job_id='failed-job')
        self.assertEqual(worker.execute()['status'],'preview')
        self.assertEqual(api.request.call_args.args[1],'/api/v1/access-jobs/preview')
        self.assertEqual(api.request.call_count,1)

    def test_old_server_requires_update_and_never_falls_back_to_a_new_request_id(self):
        api = Mock(); api.request.side_effect = self.m.ApiClientError('not found',404)
        worker = self.m.ApiJobWorker(api,'access',dict(request_id='BUS-original',requested_by='operator'),retry_job_id='failed-job')
        with self.assertRaisesRegex(self.m.ApiClientError,'업데이트'): worker.execute()
        self.assertEqual(api.request.call_count,1)

    def test_bus_completion_refresh_accepts_latest_retry_linked_to_original_request(self):
        w = self.window
        row = self.fixture.row()
        request_id = w._row_metadata(row)['request_id']
        result = dict(job_id='retry-success',request_id='RETRY-failed-job',original_request_id=request_id,
                      status='succeeded',executor_mode='powershell')
        w._api_client().request = Mock(return_value=result)
        w.run_complete(); self.fixture.wait_for_api()
        self.assertEqual([item['row'] for item in self.fixture.completed],[row])
        self.assertEqual(w._row_metadata(row)['result']['job_id'],'retry-success')

    def test_preview_of_failed_row_keeps_the_retry_candidate(self):
        w = self.window; row = self.failed_row(); w.current_row = row
        w.auto_complete_after_add = False
        w._on_api_result(dict(status='preview',valid=True))
        w._start_next_job = Mock()
        with patch.object(self.m.QMessageBox,'question',return_value=self.m.QMessageBox.Yes):
            w.run_execute()
        self.assertEqual(w.run_queue[0][3],'failed-job')

    def test_cancelled_row_is_a_new_retry_and_completed_rows_are_excluded(self):
        w=self.window;row=self.fixture.row(result=dict(job_id='stopped',status='cancelled',executor_mode='powershell'))
        done=self.fixture.row(result=dict(job_id='done',status='succeeded',executor_mode='powershell'))
        w._confirm_completed_restart=Mock(return_value=False)
        w._start_next_job=Mock()
        with patch.object(self.m.QMessageBox,'question',return_value=self.m.QMessageBox.Yes):w.run_execute()
        self.assertEqual([entry[0] for entry in w.run_queue],[row])
        self.assertEqual(w.run_queue[0][3],'stopped')
        self.assertEqual(w._row_metadata(done)['result']['job_id'],'done')
        w._set_running_ui(False)

    def test_explicit_execution_can_retry_replayed_cancelled_result_after_reload(self):
        original=dict(job_id='stopped',request_id='BUS-original',original_request_id='BUS-original',
                      status='cancelled',executor_mode='powershell',replayed=True)
        api=Mock();api.request.side_effect=[original,dict(job_id='new',status='queued'),dict(job_id='new',status='succeeded',executor_mode='powershell')]
        worker=self.m.ApiJobWorker(api,'access',dict(request_id='BUS-original',requested_by='operator'),retry_replayed=True)
        worker.poll_interval=.001
        self.assertEqual(worker.execute()['job_id'],'new')
        self.assertEqual([c.args[1] for c in api.request.call_args_list],
            ['/api/v1/access-jobs','/api/v1/access-jobs/stopped/retry','/api/v1/access-jobs/new'])

    def test_retrying_stale_parent_uses_replayed_latest_cancelled_child_once(self):
        latest=dict(job_id='stopped-child',request_id='RETRY-old',original_request_id='BUS-original',
                    status='cancelled',executor_mode='powershell',replayed=True)
        api=Mock();api.request.side_effect=[latest,dict(job_id='fresh',status='succeeded',executor_mode='powershell')]
        worker=self.m.ApiJobWorker(api,'access',dict(request_id='BUS-original',requested_by='operator'),retry_job_id='old',retry_replayed=True)
        self.assertEqual(worker.execute()['job_id'],'fresh')
        self.assertEqual([c.args[1] for c in api.request.call_args_list],
            ['/api/v1/access-jobs/old/retry','/api/v1/access-jobs/stopped-child/retry'])

    def test_network_recovery_does_not_implicitly_retry_a_terminal_attempt(self):
        api=Mock();api.request.side_effect=[self.m.ApiClientError('lost',503),dict(job_id='unknown',status='cancelled',executor_mode='powershell')]
        worker=self.m.ApiJobWorker(api,'access',dict(request_id='BUS-original',requested_by='operator'),retry_replayed=True)
        self.assertEqual(worker.execute()['status'],'cancelled')
        self.assertEqual(sum(c.args[0]=='POST' for c in api.request.call_args_list),1)

    def test_replayed_success_and_active_jobs_never_spawn_new_attempts(self):
        for status in ('running','succeeded'):
            with self.subTest(status=status):
                api=Mock();api.request.side_effect=[dict(job_id='old',status=status,replayed=True,executor_mode='powershell'),dict(job_id='old',status='succeeded',executor_mode='powershell')]
                worker=self.m.ApiJobWorker(api,'access',dict(request_id='BUS-original',requested_by='operator'),retry_replayed=True)
                worker.poll_interval=.001;self.assertEqual(worker.execute()['status'],'succeeded')
                self.assertEqual(sum(c.args[0]=='POST' for c in api.request.call_args_list),1)

    def test_cancelled_project_retry_uses_same_bus_identity_and_new_project_job(self):
        item=dict(proj='26012',name='Project',source='bus',request_id='BUS-original',
                  result=dict(job_id='stopped',status='cancelled',executor_mode='powershell'))
        api=Mock();api.request.return_value=dict(job_id='fresh',status='succeeded',executor_mode='powershell')
        worker=self.m.CreateWorker([item],{},api,auto_complete=False)
        results=[];worker.itemResult.connect(lambda row,result:results.append(result));worker.run()
        self.assertEqual(api.request.call_args_list[0].args,
            ('POST','/api/v1/project-jobs/stopped/retry',{'requested_by':'client-operator'}))
        self.assertEqual(results[0]['job_id'],'fresh')

    def test_completed_restart_requires_approval_and_keeps_bus_history(self):
        w=self.window
        row=self.fixture.row(result=dict(job_id='done',status='succeeded',executor_mode='powershell'))
        meta=w._row_metadata(row);meta['bus_done']=True;w._set_row_metadata(row,meta)
        w._confirm_completed_restart=Mock(return_value=True);w._start_next_job=Mock()
        w.run_execute()
        w._confirm_completed_restart.assert_called_once_with(1)
        self.assertEqual(w.run_queue[0][3:5],('done',True))
        self.assertTrue(w._row_metadata(row)['bus_done'])
        w._set_running_ui(False)

    def test_force_restart_sends_full_reset_only_after_approval(self):
        api=Mock();api.request.return_value=dict(job_id='fresh',status='succeeded',executor_mode='powershell')
        worker=self.m.ApiJobWorker(api,'access',dict(request_id='BUS-original',requested_by='op'),retry_job_id='done',restart_all=True)
        worker.execute()
        self.assertEqual(api.request.call_args.args,('POST','/api/v1/access-jobs/done/retry',{'requested_by':'op','restart_all':True}))

    def test_project_force_restart_does_not_repeat_bus_completion(self):
        item=dict(proj='26012',name='Project',source='bus',request_id='BUS-original',bus_done=True,restart_all=True,
                  result=dict(job_id='done',status='succeeded',executor_mode='powershell'))
        api=Mock();api.request.return_value=dict(job_id='fresh',status='succeeded',executor_mode='powershell')
        worker=self.m.CreateWorker([item],{'id':'op','pw':'secret'},api,auto_complete=True)
        worker._bus_click_process=Mock();worker.run()
        self.assertEqual(api.request.call_args.args[1],'/api/v1/project-jobs/done/retry')
        self.assertTrue(api.request.call_args.args[2]['restart_all'])
        worker._bus_click_process.assert_not_called()

    def test_reloaded_completed_job_waits_for_user_choice(self):
        for approved in (False,True):
            with self.subTest(approved=approved):
                api=Mock();api.request.side_effect=[dict(job_id='done',request_id='BUS-original',status='succeeded',executor_mode='powershell',replayed=True),
                                                    dict(job_id='fresh',status='succeeded',executor_mode='powershell')]
                worker=self.m.ApiJobWorker(api,'access',dict(request_id='BUS-original',requested_by='op'),retry_replayed=True)
                prompts=[]
                def answer(job):
                    prompts.append(job['job_id']);worker.respond_to_completed_restart(approved)
                worker.completedRestartRequested.connect(answer)
                self.assertEqual(worker.execute()['job_id'],'fresh' if approved else 'done')
                self.assertEqual(prompts,['done'])
                self.assertEqual(api.request.call_count,2 if approved else 1)
                if approved: self.assertTrue(api.request.call_args.args[2]['restart_all'])

    def test_stop_before_completed_confirmation_never_retries(self):
        api=Mock();api.request.return_value=dict(job_id='done',request_id='BUS-original',status='succeeded',executor_mode='powershell',replayed=True)
        worker=self.m.ApiJobWorker(api,'access',dict(request_id='BUS-original',requested_by='op'),retry_replayed=True)
        worker.completedRestartRequested.connect(lambda job:worker.stop())
        self.assertTrue(worker.execute()['client_stopped'])
        self.assertEqual(api.request.call_count,1)

    def test_new_project_completed_choice_controls_registration_and_preserves_history(self):
        import time
        for approved in (False,True):
            with self.subTest(approved=approved):
                w=self.window;v=self.m.NewItemsViewer(w);w._newdlg=v
                v.set_data(['프로젝트코드','프로젝트명'],[['26012','X']])
                item=v.tbl.item(0,v._hidx['프로젝트코드']);meta=dict(item.data(self.m.Qt.UserRole))
                meta.update(bus_done=True,result=dict(job_id='done',status='succeeded',executor_mode='powershell'))
                item.setData(self.m.Qt.UserRole,meta)
                api=Mock();api.api_key='test';api.request.return_value=dict(job_id='fresh',status='failed',executor_mode='powershell')
                w._api_client=Mock(return_value=api);w._confirm_completed_restart=Mock(return_value=approved)
                with patch.object(self.m.CreateWorker,'_bus_click_process') as bus:
                    v._on_create_clicked()
                    deadline=time.monotonic()+3
                    while v.worker_thread is not None and time.monotonic()<deadline:
                        self.app.processEvents();time.sleep(.01)
                    self.assertIsNone(v.worker_thread);bus.assert_not_called()
                if approved:
                    self.assertTrue(api.request.call_args.args[2]['restart_all'])
                    self.assertEqual(item.data(self.m.Qt.UserRole)['result']['status'],'failed')
                    self.assertIn('실패',v.tbl.item(0,v._status_col).text())
                else: api.request.assert_not_called()
                self.assertTrue(item.data(self.m.Qt.UserRole)['bus_done'])
                v.close();w._newdlg=None

    def test_prior_bus_completion_does_not_turn_latest_failure_into_full_restart(self):
        w=self.window
        row=self.fixture.row(result=dict(job_id='latest-failed',status='partially_succeeded',executor_mode='powershell'))
        meta=w._row_metadata(row);meta['bus_done']=True;w._set_row_metadata(row,meta)
        w._confirm_completed_restart=Mock(return_value=True);w._start_next_job=Mock()
        with patch.object(self.m.QMessageBox,'question',return_value=self.m.QMessageBox.Yes):w.run_execute()
        w._confirm_completed_restart.assert_not_called()
        self.assertEqual(w.run_queue[0][3:5],('latest-failed',False))
        self.assertTrue(w._row_metadata(row)['bus_done'])
        w._set_running_ui(False)

    def test_project_latest_failure_resumes_despite_previous_bus_completion(self):
        import time
        w=self.window;v=self.m.NewItemsViewer(w);w._newdlg=v
        v.set_data(['프로젝트코드','프로젝트명'],[['26012','X']])
        item=v.tbl.item(0,v._hidx['프로젝트코드']);meta=dict(item.data(self.m.Qt.UserRole))
        meta.update(bus_done=True,result=dict(job_id='latest-failed',status='partially_succeeded',executor_mode='powershell'))
        item.setData(self.m.Qt.UserRole,meta)
        api=Mock();api.api_key='test';api.request.return_value=dict(job_id='fresh',status='succeeded',executor_mode='powershell')
        w._api_client=Mock(return_value=api);w._confirm_completed_restart=Mock(return_value=True)
        with patch.object(self.m.CreateWorker,'_bus_click_process') as bus:
            v._on_create_clicked()
            deadline=time.monotonic()+3
            while v.worker_thread is not None and time.monotonic()<deadline:
                self.app.processEvents();time.sleep(.01)
            self.assertIsNone(v.worker_thread);bus.assert_not_called()
        w._confirm_completed_restart.assert_not_called()
        self.assertEqual(api.request.call_args.args[1],'/api/v1/project-jobs/latest-failed/retry')
        self.assertNotIn('restart_all',api.request.call_args.args[2])
        self.assertTrue(item.data(self.m.Qt.UserRole)['bus_done'])
        v.close();w._newdlg=None

    def test_group_preparation_and_unconfirmed_stop_never_auto_retry(self):
        for change in (dict(result={'job_kind':'group_preparation'}),dict(error_code='POWERSHELL_STOP_FAILED'),
                       dict(steps=[dict(error_code='POWERSHELL_STOP_FAILED')])):
            with self.subTest(change=change):
                api=Mock();api.request.return_value=dict(job_id='old',request_id='BUS-original',status='failed',executor_mode='powershell',replayed=True,**change)
                worker=self.m.ApiJobWorker(api,'access',dict(request_id='BUS-original',requested_by='op'),retry_replayed=True)
                self.assertEqual(worker.execute()['job_id'],'old')
                self.assertEqual(api.request.call_count,1)

    def test_declining_reloaded_completion_does_not_complete_bus(self):
        api=Mock();api.request.return_value=dict(job_id='done',request_id='BUS-original',status='succeeded',executor_mode='powershell',replayed=True)
        worker=self.m.ApiJobWorker(api,'access',dict(request_id='BUS-original',requested_by='op'),retry_replayed=True)
        worker.completedRestartRequested.connect(lambda job:worker.respond_to_completed_restart(False))
        result=worker.execute()
        self.assertFalse(self.m.can_complete_bus(result,'bus'))
        self.assertEqual(api.request.call_count,1)
