import time
import unittest
from unittest.mock import Mock, patch
import test_ui_design as design


class ProjectCompletionTests(unittest.TestCase):
    setUp=design.DesignTests.setUp

    def worker(self, result=None):
        api=Mock();api.api_key='test-key'
        api.request.return_value=result or dict(request_id='BUS-X',job_id='J1',project_code='26012',project_name='X',status='succeeded',executor_mode='powershell')
        worker=self.m.CreateWorker([dict(proj='26-012',name='X',source='bus',request_id='BUS-X',row=0,
            result=dict(status='succeeded',executor_mode='powershell',job_id='cached'))],dict(id='operator',pw='pw'),api,completion_only=True)
        worker._bus_click_process=Mock(return_value=(True,'완료'))
        results=[];worker.itemResult.connect(lambda row,result:results.append(result))
        return worker,api,results

    def test_recheck_only_gets_latest_project_result_then_completes(self):
        worker,api,results=self.worker();worker.run()
        api.request.assert_called_once_with('GET','/api/v1/project-jobs/by-request/BUS-X')
        worker._bus_click_process.assert_called_once_with('26-012')
        self.assertTrue(results[0]['bus_done'])
        self.assertEqual(results[0]['job_id'],'J1')

    def test_recheck_never_uses_cached_success_after_lookup_failure(self):
        for response in (self.m.ApiClientError('offline'),self.m.ApiClientError('not found',404)):
            worker,api,results=self.worker();api.request.side_effect=response
            worker.run();worker._bus_click_process.assert_not_called()
            self.assertTrue(results[0]['completion_check_failed'])
            self.assertEqual(api.request.call_count,1)

    def test_non_success_or_mismatched_project_never_completes(self):
        for change in [dict(status=s) for s in ('running','queued','failed','partially_succeeded','cancelled','simulated')]+[
            dict(executor_mode='mock'),dict(request_id='OTHER'),dict(project_code='26013'),dict(project_name='different'),dict(job_id=None)]:
            worker,api,results=self.worker();api.request.return_value.update(change)
            worker.run();worker._bus_click_process.assert_not_called()
            self.assertFalse(results[0].get('bus_done'))
            self.assertEqual(api.request.call_count,1)

    def test_manual_done_missing_request_or_stop_never_completes(self):
        for change in (dict(source='test-client'),dict(bus_done=True),dict(request_id='')):
            worker,api,results=self.worker();worker.items[0].update(change)
            worker.run();api.request.assert_not_called();worker._bus_click_process.assert_not_called()
        worker,api,results=self.worker()
        original=api.request.return_value
        api.request.side_effect=lambda *args:(worker.stop(),original)[1]
        worker.run();worker._bus_click_process.assert_not_called()
        worker,api,results=self.worker();worker.stop();worker.run();api.request.assert_not_called()

    def test_missing_credentials_or_bus_error_stays_pending(self):
        worker,api,results=self.worker();worker.creds={};worker.run()
        worker._bus_click_process.assert_not_called();self.assertFalse(results[0]['bus_done'])
        worker,api,results=self.worker();worker._bus_click_process.return_value=(False,'BUS 오류');worker.run()
        self.assertFalse(results[0]['bus_done']);self.assertIn('BUS 오류',results[0]['message'])

    def test_button_lifecycle_filters_manual_and_preserves_result_on_error(self):
        w=self.window;v=self.m.NewItemsViewer(w);w._newdlg=v
        v.set_data(['프로젝트코드','프로젝트명'],[['26012','X']]);v._manual_add('26013','Manual')
        col=v._hidx['프로젝트코드'];meta=v.tbl.item(0,col).data(self.m.Qt.UserRole)
        meta.update(result=dict(job_id='cached',status='succeeded',executor_mode='powershell'))
        v.tbl.item(0,col).setData(self.m.Qt.UserRole,meta)
        api=Mock();api.api_key='test-key';api.request.side_effect=self.m.ApiClientError('offline')
        w._api_client=Mock(return_value=api)
        v.chk_auto_complete.setChecked(False)
        with patch.object(self.m.CreateWorker,'_bus_click_process') as bus:
            v.btn_complete.click()
            self.assertFalse(v.btn_complete.isEnabled())
            deadline=time.monotonic()+3
            while v.worker_thread is not None and time.monotonic()<deadline:
                self.app.processEvents();time.sleep(.01)
            self.assertIsNone(v.worker_thread);bus.assert_not_called()
        self.assertEqual(api.request.call_count,1)
        self.assertEqual(v.tbl.item(0,col).data(self.m.Qt.UserRole)['result']['job_id'],'cached')
        self.assertIn('확인 불가',v.tbl.item(0,v._status_col).text())
        self.assertTrue(v.btn_complete.isEnabled())

    def test_both_tabs_align_actions_lists_and_log_boundaries(self):
        w=self.window;w.ensure_bus_session_async=lambda *args:None
        w.open_new_viewer();v=w._newdlg
        v.set_data(['프로젝트코드','프로젝트명'],[['26012','X']])
        w.resize(1060,900);w.show();self.app.processEvents()
        self.assertEqual(w.chk_auto_complete.text(),v.chk_auto_complete.text())
        positions=[]
        for tab,owner,table in ((0,w,w.table),(1,v,v.tbl)):
            w.tabs.setCurrentIndex(tab)
            for _ in range(3):self.app.processEvents()
            w._refresh_ui();v._refresh_selection();self.app.processEvents()
            pos=lambda widget:widget.mapTo(w,self.m.QPoint(0,0))
            positions.append((pos(owner.selection_label).x(),pos(owner.chk_auto_complete).x(),
                              pos(table).y(),table.height(),pos(w.logs_panel).y(),pos(owner.selection_label).y()))
        self.assertEqual(positions[0],positions[1])
        w._set_running_ui(True);self.assertFalse(v.btn_complete.isEnabled());w._set_running_ui(False)
