import json
import os
import tempfile
import time
import threading
import unittest
from unittest.mock import patch, Mock
from client_module import load_client

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')


class UiFlowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.m = load_client()
        cls.app = cls.m.QApplication.instance() or cls.m.QApplication([])

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.patches = []
        for key in ('DL_DIR','LOG_DIR','CONF_DIR','DEBUG_DIR'):
            p = patch.object(self.m, key, self.tmp.name)
            p.start(); self.patches.append(p)
        config = os.path.join(self.tmp.name,'api_client.json')
        with open(config, 'w') as f:
            json.dump(dict(base_url='http://localhost:8000', api_key='test'), f)
        for key, value in [('API_CONFIG_FILE',config),('CONF_FILE',os.path.join(self.tmp.name,'login.json'))]:
            p = patch.object(self.m,key,value); p.start(); self.patches.append(p)
        for name in ('refresh_notifications','_show_initial_api_settings'):
            p = patch.object(self.m.AccessManager,name); p.start(); self.patches.append(p)
        for name in ('information','warning','critical'):
            p = patch.object(self.m.QMessageBox,name); p.start(); self.patches.append(p)
        self.window = self.m.AccessManager()
        self.window.trigger_session_process.disconnect()
        self.completed = []
        self.window.trigger_session_process.connect(self.completed.extend)
        self.window._ensure_bus_session = Mock(return_value=True)
        self.refresh_results = {}
        self.api_calls = []
        self.install_refresh_api()

    def install_refresh_api(self):
        from urllib.parse import unquote
        api = self.m.FolderGrantApiClient('http://localhost:8000','test')
        def request(method,path,payload=None):
            self.api_calls.append((method,path,payload))
            self.assertEqual(method,'GET','Completion recovery must never register or rerun jobs')
            request_id = unquote(path.rsplit('/',1)[-1])
            override = self.refresh_results.get(request_id)
            if isinstance(override,Exception):
                raise override
            if override is not None:
                return override
            for row in range(self.window.table.rowCount()):
                meta = self.window._row_metadata(row)
                if meta.get('request_id') == request_id:
                    result = meta.get('result',{})
                    if result:
                        return dict(result,request_id=request_id,job_id=result.get('job_id','J1'))
            raise self.m.ApiClientError('not found',404)
        api.request = request
        self.window._api_client = Mock(return_value=api)

    def wait_for_api(self):
        deadline = time.monotonic()+3
        while self.window.api_thread is not None and time.monotonic()<deadline:
            self.app.processEvents(); time.sleep(0.01)
        self.app.processEvents()
        self.assertIsNone(self.window.api_thread)

    def test_complete_recovers_server_success_after_client_lookup_error(self):
        self.install_refresh_api()
        row = self.row()
        request_id = self.window._row_metadata(row)['request_id']
        self.refresh_results[request_id] = dict(request_id=request_id,job_id='late-job',
                                               status='succeeded',executor_mode='powershell')
        self.window.current_row = row
        self.window._on_api_error('API 조회 시간 초과')
        self.window.run_complete()
        self.wait_for_api()
        self.assertEqual([x['row'] for x in self.completed],[row])
        self.assertEqual(self.window._row_metadata(row)['result']['job_id'],'late-job')
        self.assertTrue(self.api_calls)

    def test_complete_does_not_trust_cached_success_when_lookup_fails(self):
        self.install_refresh_api()
        row = self.row(result=dict(status='succeeded',executor_mode='powershell'))
        request_id = self.window._row_metadata(row)['request_id']
        self.refresh_results[request_id] = self.m.ApiClientError('API 연결 실패')
        self.window.run_complete()
        self.wait_for_api()
        self.assertFalse(self.completed)
        self.assertTrue(self.api_calls)

    def test_complete_refresh_keeps_running_and_timeout_jobs_pending(self):
        self.install_refresh_api()
        for status in ('running','failed','partially_succeeded','simulated'):
            row = self.row()
            request_id = self.window._row_metadata(row)['request_id']
            self.refresh_results[request_id] = dict(request_id=request_id,job_id=f'job-{row}',
                status=status,executor_mode='powershell' if status!='simulated' else 'mock',
                error_code='POWERSHELL_TIMEOUT' if status in {'failed','partially_succeeded'} else None)
        self.window.run_complete()
        self.wait_for_api()
        self.assertEqual(len(self.api_calls),4)
        self.assertFalse(self.completed)
        self.assertIn('running',self.window._row_metadata(0)['result']['status'])

    def test_complete_recovers_running_job_on_a_later_click(self):
        row = self.row()
        request_id = self.window._row_metadata(row)['request_id']
        self.refresh_results[request_id] = dict(request_id=request_id,job_id='long-job',
                                               status='running',executor_mode='powershell')
        self.window.run_complete()
        self.wait_for_api()
        self.assertFalse(self.completed)
        self.refresh_results[request_id]['status'] = 'succeeded'
        self.window.run_complete()
        self.wait_for_api()
        self.assertEqual([x['row'] for x in self.completed],[row])
        self.assertEqual(len(self.api_calls),2)

    def test_complete_rejects_response_for_another_request(self):
        row = self.row(result=dict(status='succeeded',executor_mode='powershell'))
        request_id = self.window._row_metadata(row)['request_id']
        self.refresh_results[request_id] = dict(request_id='OTHER',job_id='wrong-job',
                                               status='succeeded',executor_mode='powershell')
        self.window.run_complete()
        self.wait_for_api()
        self.assertFalse(self.completed)
        self.assertFalse(self.window._row_metadata(row)['result'])

    def test_stop_during_completion_refresh_never_completes_bus(self):
        row = self.row(result=dict(status='succeeded',executor_mode='powershell'))
        api = self.window._api_client()
        request = api.request
        def delayed_request(*args):
            time.sleep(0.1)
            return request(*args)
        api.request = delayed_request
        self.window.run_complete()
        self.m.QTimer.singleShot(10,self.window._stop_all)
        self.wait_for_api()
        self.assertTrue(self.window.stop_requested)
        self.assertFalse(self.completed)

    def tearDown(self):
        self.window.close()
        deadline = time.monotonic()+3
        while (self.window.session_thread.isRunning() or self.window.watch_thread.isRunning()) and time.monotonic()<deadline:
            self.app.processEvents()
            time.sleep(0.01)
        self.assertFalse(self.window.session_thread.isRunning())
        self.assertFalse(self.window.watch_thread.isRunning())
        for p in reversed(self.patches): p.stop()
        self.tmp.cleanup()

    def row(self, source='bus', result=None):
        w = self.window
        w.add_table_row('권한부여','E123','26-012','Study','DM','','')
        r = w.table.rowCount()-1
        meta = w._row_metadata(r)
        meta.update(source=source, result=result or {}, bus_done=False)
        w._set_row_metadata(r,meta)
        return r

    def test_manual_completion_requires_bus_real_success(self):
        good = dict(status='succeeded',executor_mode='powershell')
        self.row('test-client',good)
        self.row('bus',dict(status='simulated',executor_mode='mock'))
        valid = self.row('bus',good)
        self.window.run_complete()
        self.wait_for_api()
        self.assertEqual([x['row'] for x in self.completed], [valid])

    def test_auto_completion_blocks_mock_and_stopped(self):
        w = self.window
        row = self.row()
        for result, stopped in [(dict(status='simulated',executor_mode='mock'),False),
                                (dict(status='succeeded',executor_mode='powershell'),True)]:
            w.current_row = row; w.stop_requested = stopped
            w._on_api_result(result)
        self.assertFalse(self.completed)

    def test_real_success_auto_and_bus_only_retry(self):
        w = self.window
        row = self.row()
        w.current_row = row
        w._on_api_result(dict(status='succeeded', executor_mode='powershell', job_id='J1'))
        self.assertEqual(len(self.completed),1)
        w._on_session_processed([dict(row=row,ok=False,msg='BUS failed')])
        self.completed.clear()
        w.run_complete()
        self.wait_for_api()
        self.assertEqual([x['row'] for x in self.completed],[row])
        self.assertEqual(w._row_metadata(row)['result']['job_id'],'J1')

    def test_editing_bus_request_removes_completion_authority(self):
        row = self.row(result=dict(status='succeeded',executor_mode='powershell'))
        self.window.table.item(row,self.window.COL_PROJ).setText('26-013')
        meta = self.window._row_metadata(row)
        self.assertEqual(meta['source'],'test-client')
        self.assertFalse(meta.get('result'))

    def test_project_mock_preview_and_manual_never_complete_bus(self):
        for result,source in [(dict(status='simulated',executor_mode='mock'),'bus'),
                              (dict(status='preview'),'bus'),
                              (dict(status='succeeded',executor_mode='powershell'),'test-client')]:
            worker = self.m.CreateWorker([dict(proj='26-012',name='Test',source=source,request_id='P1')],
                                         {}, self.m.FolderGrantApiClient('http://localhost:8000','key'))
            worker._bus_click_process = Mock(return_value=(True,'ok'))
            with patch.object(self.m.ApiJobWorker,'execute',return_value=result):
                worker.run()
            worker._bus_click_process.assert_not_called()

    def test_project_bus_failure_retries_completion_without_api(self):
        good = dict(status='succeeded',executor_mode='powershell',job_id='P1')
        worker = self.m.CreateWorker([dict(proj='26-012',name='Test',source='bus',
                                         request_id='P1',result=good,bus_done=False)],
                                    dict(id='operator',pw='password'),
                                    self.m.FolderGrantApiClient('http://localhost:8000','key'))
        worker._bus_click_process = Mock(return_value=(True,'ok'))
        with patch.object(self.m.ApiJobWorker,'execute') as api:
            worker.run()
        api.assert_not_called()
        worker._bus_click_process.assert_called_once_with('26-012')

    def test_stopped_project_does_not_start_browser_for_completion(self):
        worker = self.m.CreateWorker([],{},self.m.FolderGrantApiClient('http://localhost:8000','key'))
        worker.stop()
        with patch.object(self.m.webdriver,'Chrome',side_effect=AssertionError('browser started')) as browser:
            ok,message = worker._bus_click_process('26-012')
        self.assertFalse(ok)
        browser.assert_not_called()

    def test_excel_file_provenance_and_repeated_bus_download(self):
        file = os.path.join(self.tmp.name,'requests.xlsx')
        from openpyxl import Workbook
        book = Workbook()
        book.active.append(['대상자사번','프로젝트코드','폴더Level2','폴더Level3','요청번호'])
        book.active.append(['E1','26-012','Study','DM','BUS-ORIGINAL-1'])
        book.active.append(['E1','26-012','Study','DM','BUS-ORIGINAL-2'])
        book.save(file)
        w = self.window
        w.load_excel(file)
        self.assertEqual(w._row_metadata(0)['source'],'test-client')
        w.load_excel(file,source='bus')
        ids = [w._row_metadata(row)['request_id'] for row in range(2)]
        self.assertNotEqual(ids[0],ids[1])
        w.load_excel(file,source='bus')
        self.assertEqual(ids,[w._row_metadata(row)['request_id'] for row in range(2)])

    def test_project_viewer_manual_request_and_worker_lifecycle(self):
        w = self.window
        viewer = self.m.NewItemsViewer(w)
        w._newdlg = viewer
        viewer._manual_add('26-012','Test')
        item = viewer.tbl.item(0,viewer._hidx['프로젝트코드'])
        self.assertEqual(item.data(self.m.Qt.UserRole)['source'],'test-client')
        with patch.object(self.m.ApiJobWorker,'execute',return_value=dict(status='simulated',executor_mode='mock')):
            viewer._on_create_clicked()
            deadline = time.monotonic()+3
            while viewer.worker_thread is not None and time.monotonic()<deadline:
                self.app.processEvents(); time.sleep(0.01)
        self.assertIsNone(viewer.worker_thread)
        self.assertEqual(item.data(self.m.Qt.UserRole)['result']['status'],'simulated')
        viewer.close()

    def test_new_project_viewer_is_singleton(self):
        w = self.window
        w.ensure_bus_session_async = lambda purpose,callback:callback(True)
        w._load_new_items = Mock()
        w.open_new_viewer()
        first = w._newdlg
        w.open_new_viewer()
        self.assertIs(w._newdlg,first)
        self.assertEqual(sum(isinstance(x,self.m.NewItemsViewer) for x in w.children()),1)
        first.close()

    def test_stop_in_project_startup_does_not_register(self):
        w = self.window
        viewer = self.m.NewItemsViewer(w)
        w._newdlg = viewer
        viewer._manual_add('26-012','Test')
        original_busy = viewer._set_busy
        def busy(on,message=''):
            original_busy(on,message)
            if on:
                w._stop_all()
        viewer._set_busy = busy
        with patch.object(self.m.ApiJobWorker,'execute') as api:
            viewer._on_create_clicked()
            deadline = time.monotonic()+3
            while viewer.worker_thread is not None and time.monotonic()<deadline:
                self.app.processEvents(); time.sleep(0.01)
        api.assert_not_called()
        viewer.close()

    def test_manual_project_accepts_alternate_bus_name_header(self):
        viewer = self.m.NewItemsViewer(self.window)
        viewer.set_data(['프로젝트코드','과제명'],[])
        viewer._manual_add('26-012','Test')
        self.assertEqual(viewer.tbl.item(0,viewer._hidx['과제명']).text(),'Test')
        viewer.close()

    def test_close_waits_for_api_thread_without_blocking(self):
        w = self.window
        worker = Mock(); thread = Mock(); thread.isRunning.return_value = True
        w.api_worker = worker; w.api_thread = thread
        event = Mock()
        w.closeEvent(event)
        event.ignore.assert_called_once()
        worker.stop.assert_called_once()
        thread.wait.assert_not_called()
        w.api_worker = None; w.api_thread = None

    def test_qt_job_threads_run_sequentially_and_release(self):
        w = self.window
        self.row('test-client'); self.row('test-client')
        w.chk_auto_complete.setChecked(False)
        seen = []
        gui_thread = threading.get_ident()
        def execute(worker):
            seen.append(threading.get_ident())
            time.sleep(0.02)
            return dict(status='simulated',executor_mode='mock',job_id='J1')
        with patch.object(self.m.ApiJobWorker,'execute',execute):
            w.run_execute()
            deadline = time.monotonic()+3
            while (w.api_thread is not None or w.run_queue) and time.monotonic()<deadline:
                self.app.processEvents(); time.sleep(0.01)
        self.assertEqual(len(seen),2)
        self.assertTrue(all(t != gui_thread for t in seen))
        self.assertIsNone(w.api_thread)
        self.assertEqual(w.done_jobs,2)

    def test_bus_session_preparation_cannot_start_next_api_job(self):
        w = self.window
        row = self.row()
        w.current_row = row
        w.api_thread = Mock()
        w.run_queue = [(row,{},False)]
        w._start_next_job = Mock()
        def prepare(purpose):
            w._api_thread_finished()
            self.assertTrue(w._waiting_for_bus)
            return True
        w._ensure_bus_session = prepare
        w._on_api_result(dict(status='succeeded',executor_mode='powershell'))
        w._start_next_job.assert_not_called()
        w.run_queue = []
        w._waiting_for_bus = False

    def test_manual_completion_holds_controls_during_bus_preparation(self):
        w = self.window
        self.row(result=dict(status='succeeded',executor_mode='powershell'))
        def prepare(purpose):
            w._on_session_busy(False)
            self.assertTrue(w._waiting_for_bus)
            self.assertFalse(w.btn_run_execute.isEnabled())
            return True
        w._ensure_bus_session = prepare
        w.run_complete()
        self.wait_for_api()
        self.assertEqual(len(self.completed),1)
