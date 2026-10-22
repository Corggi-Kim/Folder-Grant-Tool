"""Startup sequencing and failure recovery without real API/BUS access."""
import time
import unittest
from unittest.mock import Mock, patch
import test_ui_design as design


class StartupTests(unittest.TestCase):
    setUp = design.DesignTests.setUp
    show = design.DesignTests.show
    # Reuse only the fixture lifecycle, not the inherited design cases.
    def make_dialog(self):
        dialog = self.m.StartupLoadingDialog(self.window)
        self.addCleanup(dialog.hide)
        return dialog

    def wait(self,predicate):
        end = time.monotonic()+3
        while not predicate() and time.monotonic()<end:
            self.app.processEvents(); time.sleep(.005)
        self.app.processEvents()
        self.assertTrue(predicate())

    def test_startup_checks_api_then_bus_then_both_lists_before_showing_main(self):
        w = self.window; w.creds.update(id="operator",pw="password"); dialog = self.make_dialog()
        order = []
        api = Mock()
        api.health.side_effect = lambda: order.append('health') or {}
        api.ready.side_effect = lambda: order.append('ready') or {'executor_mode':'mock'}
        w._api_client = Mock(return_value=api)
        def bus(purpose,callback):
            order.append('bus'); callback(True)
        w.ensure_bus_session_async = bus
        w._request_and_import = lambda:order.append('access')
        w._load_new_items = lambda:order.append('project')
        dialog.show(); dialog.start()
        self.wait(lambda:dialog.stage==2 and dialog.api_thread is None)
        self.assertEqual(order,['health','ready','bus','access'])
        self.assertFalse(w.isVisible())
        self.assertTrue(dialog.spinners[2].timer.isActive())
        w.startup_result.emit('access',True,'')
        self.assertEqual(order[-1],'project')
        self.assertFalse(w.isVisible())
        w.startup_result.emit('project',True,'')
        self.assertTrue(w.isVisible()); self.assertFalse(dialog.isVisible())
        self.assertFalse(dialog.active)
        api.request.assert_not_called()

    def test_failed_startup_can_retry_or_open_main_and_ignores_late_results(self):
        w = self.window; w.creds.update(id="operator",pw="password"); dialog = self.make_dialog()
        dialog.show(); dialog.active = True; w._startup_active = True
        dialog._stage(2); dialog.fail('조회 실패')
        self.assertFalse(dialog.active)
        self.assertFalse(dialog.retry.isHidden())
        self.assertFalse(dialog.spinners[2].timer.isActive())
        self.assertIn('조회하지 못했습니다',w.table.empty_label.text())
        w._load_new_items = Mock()
        w.startup_result.emit('access',True,'')
        w._load_new_items.assert_not_called()
        dialog.enter.click(); self.assertTrue(w.isVisible())

    def test_api_connection_error_does_not_start_bus_or_show_raw_response(self):
        w = self.window; w.creds.update(id="operator",pw="password"); dialog = self.make_dialog()
        api = Mock(); api.health.side_effect = RuntimeError('raw network failure')
        w._api_client = Mock(return_value=api)
        w.ensure_bus_session_async = Mock()
        dialog.show(); dialog.start()
        self.wait(lambda:dialog.api_thread is None)
        self.assertFalse(dialog.active)
        self.assertIn('API 연결에 실패',dialog.message.text())
        w.ensure_bus_session_async.assert_not_called()

    def test_zero_badges_hidden_and_new_counts_do_not_replace_loaded_rows(self):
        w = self.window
        w._on_counts_ready({})
        self.assertTrue(all(b.isHidden() for b in w.tab_count_badges))
        row = self.fixture.row(); w._mark_list_loaded('access',1)
        before = w._row_values(row)
        w._on_counts_ready({'진행-부여':2})
        self.assertFalse(w.tab_count_badges[0].isHidden())
        self.assertFalse(w.new_request_notice.isHidden())
        self.assertEqual(w._row_values(row),before)
        w._mark_list_loaded('access',2)
        self.assertTrue(w.new_request_notice.isHidden())

    def test_empty_list_and_pending_fetch_have_distinct_messages(self):
        w = self.window; w._ensure_project_viewer()
        for kind,table in [('access',w.table),('project',w._newdlg.tbl)]:
            table.set_empty_message('요청을 조회하고 있습니다…')
            self.assertIn('조회하고',table.empty_label.text())
            w._mark_list_loaded(kind,0)
            self.assertEqual(table.empty_label.text(),'대기 중인 요청이 없습니다.')
            self.assertFalse(table.empty_label.isHidden())

    def test_preloaded_project_tab_does_not_download_again(self):
        w = self.window; w._ensure_project_viewer()
        w._mark_list_loaded('project',0)
        w.ensure_bus_session_async = Mock(); w._load_new_items = Mock()
        w.tabs.setCurrentIndex(1)
        w.ensure_bus_session_async.assert_not_called(); w._load_new_items.assert_not_called()

    def test_download_error_is_handled_inside_startup_without_modal(self):
        w = self.window; w.creds.update(id="operator",pw="password"); dialog = self.make_dialog()
        dialog.active = True; w._startup_active = True; dialog._stage(3)
        with patch.object(self.m.QMessageBox,'critical') as critical:
            w._on_new_downloaded('','download failed')
        critical.assert_not_called()
        self.assertFalse(dialog.active)



    def test_spinner_changes_angle_without_blocking_gui(self):
        dialog = self.make_dialog(); dialog.show(); dialog._stage(0)
        spinner = dialog.spinners[0]; before = spinner.angle
        self.wait(lambda:spinner.angle!=before)
        dialog.hide()
        self.assertFalse(spinner.timer.isActive())

    def test_permission_query_uses_async_login_callback(self):
        w = self.window
        callbacks = []
        w.ensure_bus_session_async = lambda purpose,callback:callbacks.append(callback)
        calls = []
        w.trigger_session_download.connect(lambda:calls.append('download'))
        w._request_and_import()
        self.assertFalse(calls)
        callbacks[0](True)
        self.assertEqual(calls,['download'])

    def test_startup_retries_after_api_failure(self):
        w = self.window; w.creds.update(id='operator',pw='password'); dialog = self.make_dialog()
        api = Mock(); api.health.side_effect = RuntimeError('offline')
        w._api_client = Mock(return_value=api)
        dialog.show(); dialog.start(); self.wait(lambda:dialog.api_thread is None)
        self.assertFalse(dialog.active)
        api.health.side_effect = None; api.health.return_value = {}; api.ready.return_value = {}
        w.ensure_bus_session_async = lambda purpose,callback:callback(True)
        w._request_and_import = Mock()
        dialog.start(); self.wait(lambda:dialog.api_thread is None)
        self.assertTrue(dialog.active); self.assertEqual(dialog.stage,2)
        w._request_and_import.assert_called_once()
        dialog.open_main()

    def test_access_download_with_zero_rows_is_a_successful_empty_result(self):
        from pathlib import Path
        from openpyxl import Workbook
        w = self.window
        path = Path(self.fixture.tmp.name)/'empty.xlsx'
        workbook = Workbook(); workbook.active.append(['대상자사번','프로젝트코드','폴더Level2','폴더Level3'])
        workbook.save(path); workbook.close()
        outcomes = []; w.startup_result.connect(lambda *args:outcomes.append(args))
        w._on_session_downloaded(str(path),'')
        self.assertEqual(outcomes[-1],('access',True,''))
        self.assertEqual(w.table.rowCount(),0)
        self.assertEqual(w.table.empty_label.text(),'대기 중인 요청이 없습니다.')
        self.assertTrue(w.tab_count_badges[0].isHidden())

    def test_access_download_parse_failure_does_not_complete_startup(self):
        w = self.window
        outcomes = []; w.startup_result.connect(lambda *args:outcomes.append(args))
        w._on_session_downloaded('/missing.xlsx','')
        self.assertEqual(outcomes[-1][0:2],('access',False))
        self.assertFalse(w._list_loaded['access'])

    def test_project_download_waits_for_previous_session_operation_to_finish(self):
        w = self.window; dialog = self.make_dialog()
        dialog.active = True; dialog._stage(2)
        w.session.is_busy = Mock(return_value=True)
        w._load_new_items = Mock()
        dialog._loaded('access',True,'')
        w._load_new_items.assert_not_called()
        w.session.is_busy.return_value = False
        self.wait(lambda:w._load_new_items.called)
        dialog.open_main()
