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
