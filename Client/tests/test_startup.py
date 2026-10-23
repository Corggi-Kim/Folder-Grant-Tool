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

    def test_loading_lists_does_not_activate_notifications(self):
        w = self.window
        row = self.fixture.row(); before = w._row_values(row)
        w._mark_list_loaded('access',1)
        self.assertEqual(w._row_values(row),before)
        self.assertFalse(hasattr(w,'notify_timer'))
        self.assertFalse(hasattr(w,'tab_count_badges'))

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
        self.assertFalse(hasattr(w,'tab_count_badges'))

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

    def test_startup_missing_credentials_uses_login_dialog_and_preserves_settings(self):
        w = self.window; w.creds = {'id':'','pw':'','debug':True,'remove_fail_tol':7}
        dialog = self.make_dialog(); dialog.active = True; dialog.stage = 0
        w._open_settings = Mock(); w.ensure_bus_session_async = Mock()
        login = Mock(); login.exec_.return_value = self.m.QDialog.Accepted
        login.values.return_value = ('operator','secret with space ')
        login.remember.isChecked.return_value = False
        with patch.object(self.m,'BusLoginDialog',return_value=login):
            dialog._api_ready({'ready':{}})
        w._open_settings.assert_not_called()
        self.assertEqual(w.creds['id'],'operator')
        self.assertEqual(w.creds['pw'],'secret with space ')
        self.assertTrue(w.creds['debug'])
        self.assertEqual(w.creds['remove_fail_tol'],7)
        w.ensure_bus_session_async.assert_called_once()
        from pathlib import Path
        path = Path(self.m.CONF_FILE)
        self.assertNotIn('secret with space ',path.read_text(encoding='utf-8'))

    def test_cancelled_startup_login_does_not_start_bus(self):
        w = self.window; w.creds = {'id':'','pw':''}; dialog = self.make_dialog()
        dialog.active = True; dialog.stage = 0; w.ensure_bus_session_async = Mock()
        login = Mock(); login.exec_.return_value = self.m.QDialog.Rejected
        with patch.object(self.m,'BusLoginDialog',return_value=login):
            dialog._api_ready({'ready':{}})
        w.ensure_bus_session_async.assert_not_called()
        self.assertFalse(dialog.active)
        self.assertIn('취소',dialog.message.text())

    def test_login_dialog_has_only_credential_inputs_and_masks_password(self):
        login = self.m.BusLoginDialog('saved-user',self.window)
        self.assertEqual(login.username.text(),'saved-user')
        self.assertEqual(login.password.echoMode(),self.m.QLineEdit.Password)
        self.assertEqual(login.password.text(),'')
        self.assertFalse(login.login_button.isEnabled())
        login.password.setText('secret ')
        self.assertTrue(login.login_button.isEnabled())
        self.assertEqual(login.values(),('saved-user','secret '))
        self.assertEqual(len(login.findChildren(self.m.QLineEdit)),2)
        self.assertEqual(len(login.findChildren(self.m.QCheckBox)),1)
        self.assertFalse(login.remember.isChecked())
        self.assertFalse(login.findChildren(self.m.QSpinBox))
        login.close()

    def test_failed_bus_login_prompts_for_credentials_on_retry(self):
        w = self.window; w.creds = {'id':'wrong-user','pw':'wrong-password'}
        dialog = self.make_dialog(); dialog.active = True; dialog.stage = 1
        dialog._bus_ready(False)
        dialog.active = True; dialog.stage = 0; w.ensure_bus_session_async = Mock()
        login = Mock(); login.exec_.return_value = self.m.QDialog.Accepted
        login.values.return_value = ('correct-user','correct-password')
        login.remember.isChecked.return_value = False
        with patch.object(self.m,'BusLoginDialog',return_value=login) as factory:
            dialog._api_ready({'ready':{}})
        factory.assert_called_once()
        self.assertEqual(w.creds['id'],'correct-user')
        w.ensure_bus_session_async.assert_called_once()

    def test_startup_login_save_choice_uses_existing_settings_behavior_and_preserves_other_settings(self):
        import json
        w = self.window; w.creds = {'id':'','pw':''}
        with open(self.m.CONF_FILE,'w',encoding='utf-8') as handle:
            json.dump({'theme':'dark','debug':True,'remove_fail_tol':9},handle)
        dialog = self.make_dialog(); dialog.active = True; dialog.stage = 0
        callbacks = []; w.ensure_bus_session_async = lambda purpose,callback:callbacks.append(callback)
        w._request_and_import = Mock()
        login = Mock(); login.exec_.return_value = self.m.QDialog.Accepted
        login.values.return_value = ('operator','saved-secret'); login.remember.isChecked.return_value = True
        with patch.object(self.m,'BusLoginDialog',return_value=login):
            dialog._api_ready({'ready':{}})
        with open(self.m.CONF_FILE,encoding='utf-8') as handle:
            self.assertEqual(json.load(handle)['pw'],'saved-secret')
        callbacks[0](True)
        with open(self.m.CONF_FILE,encoding='utf-8') as handle: saved = json.load(handle)
        self.assertEqual(saved['pw'],'saved-secret')
        self.assertEqual(saved['id'],'operator')
        self.assertEqual(saved['theme'],'dark'); self.assertTrue(saved['debug'])
        self.assertEqual(saved['remove_fail_tol'],9)
        self.assertNotIn('saved-secret',w.log.toPlainText())

    def test_startup_login_without_save_removes_old_saved_credentials_after_success(self):
        import json
        w = self.window; w.creds = {'id':'old','pw':'old-password'}
        with open(self.m.CONF_FILE,'w',encoding='utf-8') as handle:
            json.dump({'id':'old','pw':'old-password','theme':'dark'},handle)
        dialog = self.make_dialog(); dialog._login_failed = True; dialog.active = True; dialog.stage = 0
        w.ensure_bus_session_async = lambda purpose,callback:callback(True)
        w._request_and_import = Mock()
        login = Mock(); login.exec_.return_value = self.m.QDialog.Accepted
        login.values.return_value = ('new-user','session-password'); login.remember.isChecked.return_value = False
        with patch.object(self.m,'BusLoginDialog',return_value=login): dialog._api_ready({'ready':{}})
        with open(self.m.CONF_FILE,encoding='utf-8') as handle: saved = json.load(handle)
        self.assertNotIn('id',saved); self.assertNotIn('pw',saved)
        self.assertEqual(saved['theme'],'dark')
        self.assertEqual(w.creds['pw'],'session-password')

    def test_login_save_choice_matches_settings_even_when_authentication_fails(self):
        from pathlib import Path
        w = self.window; w.creds = {'id':'','pw':''}
        dialog = self.make_dialog(); dialog.active = True; dialog.stage = 0
        w.ensure_bus_session_async = lambda purpose,callback:callback(False)
        login = Mock(); login.exec_.return_value = self.m.QDialog.Accepted
        login.values.return_value = ('bad-user','bad-password'); login.remember.isChecked.return_value = True
        with patch.object(self.m,'BusLoginDialog',return_value=login): dialog._api_ready({'ready':{}})
        self.assertIn('bad-password',Path(self.m.CONF_FILE).read_text(encoding='utf-8'))
        self.assertNotIn('bad-password',w.log.toPlainText())

    def test_changed_credentials_never_reuse_an_authenticated_driver(self):
        manager = self.m.BusSessionManager(self.fixture.tmp.name)
        old_driver = Mock(); manager.driver = old_driver; manager._ready = True
        manager._authenticated_credentials = ('old-user','old-password')
        manager.set_creds('new-user','new-password')
        self.assertFalse(manager.is_ready())
        manager.MAX_INIT_RETRY = 1
        outcomes = []; manager.readyChanged.connect(lambda ok,msg:outcomes.append(ok))
        with patch.object(self.m.webdriver,'Chrome',side_effect=RuntimeError('login failed')):
            manager.start()
        old_driver.quit.assert_called_once()
        self.assertEqual(outcomes,[False])
        self.assertFalse(manager.is_busy())

    def test_login_retries_emit_failure_only_after_all_attempts_finish(self):
        manager = self.m.BusSessionManager(self.fixture.tmp.name); manager.set_creds('user','password')
        manager.INIT_RETRY_BASE_DELAY = 0
        outcomes = []; manager.readyChanged.connect(lambda ok,msg:outcomes.append(ok))
        def fail(**kwargs):
            self.assertTrue(manager.is_busy())
            raise RuntimeError('login failed')
        with patch.object(self.m.webdriver,'Chrome',side_effect=fail) as chrome:
            manager.start()
        self.assertEqual(chrome.call_count,manager.MAX_INIT_RETRY)
        self.assertEqual(outcomes,[False])
        self.assertFalse(manager.is_busy())

    def test_general_settings_does_not_recheck_save_for_session_only_credentials(self):
        w = self.window; w.creds = {'id':'session-user','pw':'session-password'}
        w._remember_credentials = False
        settings = Mock(); settings.exec_.return_value = self.m.QDialog.Rejected
        with patch.object(self.m,'SettingsDialog',return_value=settings) as factory:
            w._open_settings()
        self.assertFalse(factory.call_args.kwargs['remembered'])
