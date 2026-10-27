import os
from unittest.mock import Mock, patch
import unittest
import test_ui_design as design


def validation(request_id='BUS-X', state='passed'):
    return dict(request_id=request_id,validation_status=state,executor_mode='powershell',checked_at='2026-10-23T00:00:00Z',
                checks=[dict(id='user',label='대상 사용자',target='E123',status='passed',message='존재'),
                        dict(id='level3',label='Level3 폴더',target=r'\\server\share\Study\DM',status=state,message='폴더 확인')],
                warnings=['권한은 변경하지 않습니다.'])


class ValidationTests(unittest.TestCase):
    setUp=design.DesignTests.setUp

    def test_running_access_and_project_steps_animate_and_stop(self):
        w=self.window
        row=self.fixture.row()
        w.table.setCurrentCell(row,w.COL_PROJ)
        w.current_row=row
        w.show();self.app.processEvents()
        w.table.item(row,w.COL_STATUS).setText('요청 검증 중')
        w._set_row_metadata(row,dict(w._row_metadata(row),live_result=dict(status='validating')))
        w._refresh_detail()
        self.assertEqual(w.step_icons[0].property('state'),'running')
        self.assertTrue(w.step_icons[0].spinner.timer.isActive())
        w.table.item(row,w.COL_STATUS).setText('API 실행중')
        w._on_api_progress(dict(job_id='live',status='running',steps=[dict(type='add_ad_group_member',status='running')]))
        self.assertEqual(w.step_icons[1].property('state'),'running')
        self.assertTrue(w.step_icons[1].spinner.timer.isActive())
        w._on_api_progress(dict(job_id='live',status='running',steps=[dict(type='add_ad_group_member',status='succeeded'),dict(type='grant_folder_acl',status='running')]))
        self.assertFalse(w.step_icons[1].spinner.timer.isActive())
        self.assertTrue(w.step_icons[2].spinner.timer.isActive())
        v=self.m.NewItemsViewer(w);v.show();self.app.processEvents()
        v._refresh_project_process(dict(project_action='running'),dict(job_id='p',steps=[dict(type='ensure_project_folder',status='running')]))
        self.assertEqual(v.project_step_icons[1].property('state'),'running')
        self.assertTrue(v.project_step_icons[1].spinner.timer.isActive())
        v._refresh_project_process({},dict(job_id='p',steps=[dict(type='ensure_project_folder',status='succeeded')]))
        self.assertFalse(v.project_step_icons[1].spinner.timer.isActive())
        v.close()

    def test_live_worker_never_registers_or_uses_preview_or_recovers_by_request(self):
        api=Mock();api.request.return_value=validation()
        worker=self.m.ApiJobWorker(api,'access',dict(request_id='BUS-X'),True,retry_job_id='old',validate_only=True)
        result=worker.execute()
        self.assertEqual(result['status'],'validated')
        api.request.assert_called_once_with('POST','/api/v1/access-jobs/validate',dict(request_id='BUS-X'))
        self.assertFalse(self.m.can_complete_bus(result,'bus'))

    def test_old_server_or_wrong_response_never_falls_back_to_execution(self):
        for response in (self.m.ApiClientError('missing',404),validation('wrong'),dict(valid=True)):
            api=Mock()
            if isinstance(response,Exception):api.request.side_effect=response
            else:api.request.return_value=response
            worker=self.m.ApiJobWorker(api,'access',dict(request_id='BUS-X'),validate_only=True)
            with self.assertRaises(self.m.ApiClientError):worker.execute()
            self.assertEqual(api.request.call_count,1)

    def test_access_button_logs_each_check_and_preserves_actual_failure(self):
        w=self.window
        row=self.fixture.row(result=dict(job_id='old',status='partially_succeeded',executor_mode='powershell'))
        before=w._row_metadata(row)
        w.table.setCurrentCell(-1,-1)
        api=Mock();api.api_key='key';api.request.return_value=validation(before['request_id'],'failed')
        w._api_client=Mock(return_value=api)
        w.btn_preview.click();self.fixture.wait_for_api()
        self.assertEqual(w.table.currentRow(),row)
        meta=w._row_metadata(row)
        self.assertEqual(meta['result'],before['result'])
        self.assertEqual(meta['request_id'],before['request_id'])
        self.assertFalse(self.fixture.completed)
        self.assertIn('요청 검증 실패',w.log.toPlainText())
        self.assertIn('Level3 폴더',w.log.toPlainText())
        self.assertNotIn('"checks"',w.log.toPlainText())
        with open(os.path.join(self.m.LOG_DIR,'access.log'),encoding='utf-8') as f:
            self.assertIn('"checks"',f.read())
        self.assertEqual(w.step_icons[0].property('state'),'failed')

    def test_project_validation_network_error_preserves_cached_success(self):
        worker=self.m.CreateWorker([dict(proj='26012',name='X',source='bus',request_id='BUS-X',result=dict(job_id='old',status='succeeded',executor_mode='powershell'))],{},Mock(),True,validate_only=True)
        worker.client.request.side_effect=self.m.ApiClientError('update API',404)
        results=[];worker.itemResult.connect(lambda row,result:results.append(result))
        worker._bus_click_process=Mock()
        worker.run()
        self.assertEqual(results[0]['status'],'validated')
        self.assertEqual(results[0]['validation_status'],'unknown')
        worker._bus_click_process.assert_not_called()
        v=self.m.NewItemsViewer(self.window);v._manual_add('26012','X')
        v._on_project_result(0,dict(job_id='old',status='succeeded',executor_mode='powershell'))
        v._on_project_result(0,results[0])
        self.assertEqual(v.tbl.item(0,v._hidx['프로젝트코드']).data(self.m.Qt.UserRole)['result']['job_id'],'old')
        v.close()

    def test_validation_flag_cannot_reuse_success_for_bus_completion(self):
        api=Mock();api.request.return_value=validation()
        worker=self.m.CreateWorker([dict(proj='26012',name='X',source='bus',request_id='BUS-X',result=dict(status='succeeded',executor_mode='powershell'))],{},api,validate_only=True)
        worker._bus_click_process=Mock()
        results=[];worker.itemResult.connect(lambda row,result:results.append(result))
        worker.run()
        api.request.assert_called_once()
        self.assertEqual(results[0]['status'],'validated')
        worker._bus_click_process.assert_not_called()
        check=self.m.ApiJobWorker(api,'access',dict(request_id='BUS-X'),validate_only=True)
        check.stop()
        self.assertEqual(check.execute()['status'],'validated')

    def test_timer_restarts_after_finish_and_freezes_during_idle(self):
        w=self.window
        with patch.object(self.m.time,'monotonic',return_value=100):
            w._set_progress_status('요청 1',running=True)
        with patch.object(self.m.time,'monotonic',return_value=112):
            self.assertEqual(w._format_elapsed(),'00:12')
            w._set_running_ui(False)
        with patch.object(self.m.time,'monotonic',return_value=999):
            self.assertEqual(w._format_elapsed(),'00:12')
            w._set_progress_status('요청 2',running=True)
            self.assertEqual(w._format_elapsed(),'00:00')
        w._finish_progress_status()

    def test_replayed_running_job_does_not_display_historical_age(self):
        w=self.window;row=self.fixture.row(result=dict(job_id='old',status='running',started_at='2020-01-01T00:00:00Z',executor_mode='powershell'))
        w.table.item(row,w.COL_STATUS).setText('API 실행중')
        w.table.setCurrentCell(row,w.COL_PROJ)
        with patch.object(self.m.time,'monotonic',return_value=100):w._set_progress_status('작업 중')
        with patch.object(self.m.time,'monotonic',return_value=105):w._refresh_detail()
        self.assertIn('이번 요청 경과 00:05',w.status_label.text())
        w._finish_progress_status()
