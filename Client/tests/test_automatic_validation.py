"""Execution gate: read-only preflight must pass before registration or retry."""
import unittest
from unittest.mock import Mock, patch
import test_ui_design as design
from test_validation import validation


class AutomaticValidationTests(unittest.TestCase):
    setUp=design.DesignTests.setUp

    def worker(self,kind='access',retry=None):
        api=Mock();api.api_key='test'
        worker=self.m.ApiJobWorker(api,kind,dict(request_id='BUS-X',requested_by='op'),retry_job_id=retry,preflight=True)
        return worker,api

    def test_access_project_and_retry_register_only_after_passed_validation(self):
        for kind,retry in [('access',None),('project',None),('access','old')]:
            with self.subTest(kind=kind,retry=retry):
                worker,api=self.worker(kind,retry)
                api.request.side_effect=[validation(),dict(job_id='J1',status='succeeded',executor_mode='powershell')]
                events=[];worker.progress.connect(events.append)
                result=worker.execute()
                self.assertEqual(result['status'],'succeeded')
                paths=[call.args[1] for call in api.request.call_args_list]
                self.assertEqual(paths,[f'/api/v1/{kind}-jobs/validate',f'/api/v1/{kind}-jobs'+('/old/retry' if retry else '')])
                self.assertEqual(events[0]['status'],'validating')
                self.assertTrue(any(event.get('auto_validation') for event in events))

    def test_failed_warning_unknown_mock_invalid_response_and_old_server_never_register(self):
        responses=[validation(state=state) for state in ['failed','warning','unknown']]
        responses += [dict(validation(),executor_mode='mock'),validation('wrong'),dict(valid=True),
                      self.m.ApiClientError('missing',404),self.m.ApiClientError('offline')]
        for response in responses:
            with self.subTest(response=response):
                worker,api=self.worker()
                if isinstance(response,Exception):api.request.side_effect=response
                else:api.request.return_value=response
                result=worker.execute()
                self.assertTrue(result['preflight_blocked'])
                self.assertEqual(result['status'],'validated')
                self.assertFalse(self.m.can_complete_bus(result,'bus'))
                api.request.assert_called_once_with('POST','/api/v1/access-jobs/validate',worker.payload)

    def test_stop_during_or_immediately_after_validation_never_registers_or_cancels(self):
        for stop_on_report in [False,True]:
            worker,api=self.worker()
            def response(*args):
                if not stop_on_report:worker.stop()
                return validation()
            api.request.side_effect=response
            worker.progress.connect(lambda report:worker.stop() if stop_on_report and report.get('auto_validation') else None)
            result=worker.execute()
            self.assertTrue(result['preflight_blocked'])
            self.assertTrue(result['client_stopped'])
            self.assertEqual(api.request.call_count,1)

    def test_access_ui_validates_each_row_logs_results_and_preserves_blocked_prior_job(self):
        w=self.window;w.chk_auto_complete.setChecked(False)
        first=self.fixture.row(result=dict(job_id='old',status='partially_succeeded',executor_mode='powershell'))
        second=self.fixture.row('test-client')
        first_id=w._row_metadata(first)['request_id']
        api=Mock();api.api_key='test';calls=[]
        def request(method,path,payload=None):
            calls.append((method,path,payload))
            if path.endswith('/validate'):
                return validation(payload['request_id'],'failed' if payload['request_id']==first_id else 'passed')
            return dict(job_id='new',status='succeeded',executor_mode='powershell')
        api.request.side_effect=request;w._api_client=Mock(return_value=api)
        with patch.object(self.m.QMessageBox,'question',return_value=self.m.QMessageBox.Yes):
            w.btn_run_execute.click();self.fixture.wait_for_api()
        self.assertEqual(w._row_metadata(first)['result']['job_id'],'old')
        self.assertIn('실행 보류',w.table.item(first,w.COL_STATUS).text())
        self.assertEqual(w._row_metadata(second)['result']['job_id'],'new')
        self.assertEqual(len(calls),3)
        self.assertTrue(calls[0][1].endswith('/validate'))
        self.assertTrue(calls[1][1].endswith('/validate'))
        self.assertEqual(calls[2][1],'/api/v1/access-jobs')
        self.assertIn('Level3 폴더',w.log.toPlainText())
        self.assertIn('실행 전 검증',w.log.toPlainText())
        self.assertFalse(self.fixture.completed)
        self.assertTrue(w.btn_preview.isHidden())

    def test_project_execution_registers_directly_without_preflight(self):
        for status,mode in [('succeeded','powershell'),('simulated','mock'),('failed','powershell')]:
            with self.subTest(status=status):
                api=Mock();api.request.return_value=dict(request_id='BUS-X',job_id='PROJECT',status=status,executor_mode=mode)
                worker=self.m.CreateWorker([dict(proj='26012',name='X',source='bus',request_id='BUS-X')],{},api,auto_complete=False)
                worker._bus_click_process=Mock();results=[]
                worker.itemResult.connect(lambda row,result:results.append(result))
                worker.run()
                api.request.assert_called_once()
                self.assertEqual(api.request.call_args.args[:2],('POST','/api/v1/project-jobs'))
                self.assertEqual(results[0]['job_id'],'PROJECT')
                self.assertNotIn('preflight_blocked',results[0])
                worker._bus_click_process.assert_not_called()
        v=self.m.NewItemsViewer(self.window);self.addCleanup(v.close)
        self.assertTrue(v.btn_preview.isHidden())
