"""Rendered layout and interaction contracts for the approved desktop mockup."""
import unittest
import json
from unittest.mock import Mock
import test_ui_flow as fixtures


class DesignTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.UiFlowTests()
        self.fixture.setUpClass()
        self.fixture.setUp()
        self.addCleanup(self.fixture.tearDown)
        self.window = self.fixture.window
        self.app = self.fixture.app
        self.m = self.fixture.m

    def show(self):
        self.window.show()
        self.app.processEvents()

    def test_version_and_build_are_visible_in_main_header(self):
        self.assertEqual(self.m.APP_VERSION,'1.1.0')
        self.assertEqual(self.m.APP_BUILD,'2026-10-22')
        self.assertEqual(self.window.release_label.text(),'V1.1.0 Build:20261022')
        self.show()
        w = self.window
        self.assertGreater(w.release_label.mapTo(w,self.m.QPoint(0,0)).y(),w.btn_stop.mapTo(w,self.m.QPoint(0,0)).y()+w.btn_stop.height())
        self.assertEqual(w.release_label.alignment(),self.m.Qt.AlignRight|self.m.Qt.AlignVCenter)

    def test_approved_layout_has_right_inspector_and_bottom_actions(self):
        self.show()
        w = self.window
        pos = lambda widget:widget.mapTo(w,self.m.QPoint(0,0))
        self.assertLess(pos(w.connection_bar).y(),pos(w.table).y())
        self.assertGreater(pos(w.detail_panel).x(),pos(w.table).x()+w.table.width()-5)
        self.assertGreater(pos(w.logs_panel).y(),pos(w.table).y()+w.table.height()-5)
        self.assertGreater(pos(w.actions_bar).y(),pos(w.logs_panel).y()+w.logs_panel.height()-5)
        self.assertEqual(w.btn_preview.text(),'미리보기')
        self.assertEqual(w.btn_run_execute.text(),'선택 실행')
        self.assertEqual(w.btn_run_complete.text(),'BUS 완료 재확인')

    def test_inspector_shows_real_row_role_without_changing_payload_data(self):
        row = self.fixture.row()
        w = self.window
        w.table.item(row,w.COL_ROLE).setText('Manager')
        w.table.setCurrentCell(row,w.COL_PROJ)
        self.show()
        self.assertIn('26-012',w.detail_project.text())
        self.assertIn('E123',w.detail_employee.text())
        self.assertIn('Manager',w.detail_role.text())
        self.assertEqual(w._row_values(row)['lv2'],'Study')
        self.assertEqual(w._row_values(row)['lv3'],'DM')

    def test_preview_and_execute_route_to_same_existing_job_flow(self):
        w = self.window
        w.run_execute = Mock()
        w.btn_preview.click()
        self.assertTrue(w.chk_dry.isChecked())
        w.run_execute.assert_called_once()
        w.run_execute.reset_mock()
        w.btn_run_execute.click()
        self.assertFalse(w.chk_dry.isChecked())
        w.run_execute.assert_called_once()

    def test_connection_strip_does_not_invent_live_connectivity_or_executor_mode(self):
        w = self.window
        self.assertIn('미확인',w.api_status_label.text())
        self.assertFalse(hasattr(w,'execution_mode_label'))
        self.assertIsNone(w._executor_mode)
        w.current_row = self.fixture.row()
        w._on_api_progress(dict(job_id='J',status='running',executor_mode='mock',steps=[]))
        self.assertIn('연결됨',w.api_status_label.text())
        self.assertEqual(w._executor_mode,'mock')

    def test_both_palettes_and_log_collapse_preserve_state(self):
        w = self.window
        self.show()
        for mode in ('light','dark'):
            w.apply_theme(mode)
            self.assertEqual(w.current_theme,mode)
            self.assertIn(self.m.THEMES[mode]['bg'],w.styleSheet())
        w.log_toggle.click()
        self.assertTrue(w.log.isHidden())
        w.log_toggle.click()
        self.assertFalse(w.log.isHidden())

    def test_selection_summary_tracks_checkboxes(self):
        self.fixture.row(); self.fixture.row()
        self.show()
        self.assertIn('2',self.window.selection_label.text())
        self.window._get_checkbox(0).setChecked(False)
        self.app.processEvents()
        self.assertIn('1',self.window.selection_label.text())

    def test_project_workspace_is_embedded_and_returns_to_requests(self):
        w = self.window
        w.ensure_bus_session_async = lambda purpose,callback:None
        w.open_new_viewer()
        self.assertEqual(w.tabs.currentIndex(),1)
        self.assertIs(w._newdlg.host,w)
        self.assertFalse(w._newdlg.isWindow())
        w.apply_theme('dark')
        self.assertEqual(w._newdlg.styleSheet(),'')
        w.tabs.setCurrentIndex(0)
        self.assertEqual(w.tabs.currentIndex(),0)
        self.assertIsNotNone(w._newdlg)

    def test_busy_controls_include_new_preview_and_workspace_navigation(self):
        w = self.window
        w._set_running_ui(True)
        self.assertFalse(w.btn_preview.isEnabled())
        self.assertFalse(w.tabs.isEnabled())
        self.assertFalse(w.btn_run_execute.isEnabled())
        self.assertTrue(w.btn_stop.isEnabled())
        w._set_running_ui(False)
        self.assertTrue(w.tabs.isEnabled())

    def test_detail_steps_follow_server_results(self):
        row = self.fixture.row(result=dict(job_id='J1',status='running',executor_mode='powershell',
              steps=[dict(type='add_group_member',status='succeeded'),dict(type='grant_acl',status='running')]))
        self.window.table.setCurrentCell(row,self.window.COL_PROJ)
        self.show()
        self.assertEqual([icon.property('state') for icon in self.window.step_icons],
                         ['succeeded','succeeded','running','queued'])

    def test_optional_columns_keep_existing_fields_editable(self):
        w = self.window
        self.assertTrue(w.table.isColumnHidden(w.COL_ROLE))
        w.columns_action.setChecked(True)
        self.assertFalse(w.table.isColumnHidden(w.COL_ROLE))
        self.assertFalse(w.table.isColumnHidden(w.COL_LV3))
        self.assertEqual(w.table.horizontalHeaderItem(w.COL_LV2).text(),'Level2')

    def test_theme_switch_persists_choice_in_both_directions(self):
        w = self.window
        for mode in ('dark','light'):
            w._toggle_theme()
            self.assertEqual(w.current_theme,mode)
            with open(self.m.CONF_FILE,encoding='utf-8') as handle:
                self.assertEqual(json.load(handle)['theme'],mode)

    def test_execution_clears_previous_live_progress_before_retry(self):
        row = self.fixture.row()
        w = self.window
        meta = w._row_metadata(row)
        meta['live_result'] = dict(status='running',steps=[dict(type='grant_acl',status='succeeded')])
        w._set_row_metadata(row,meta)
        w._start_next_job = Mock()
        w.run_execute()
        self.assertFalse(w._row_metadata(row).get('live_result'))

    def test_project_results_refresh_header_mode(self):
        w = self.window
        viewer = self.m.NewItemsViewer(w)
        viewer.set_data(['프로젝트코드','프로젝트명'],[['26-012','Test']])
        viewer._on_project_result(0,dict(status='succeeded',executor_mode='powershell'))
        self.assertEqual(w._executor_mode,'powershell')
        self.assertIn('연결됨',w.api_status_label.text())
        viewer.close()

    def test_workspace_tab_metrics_use_the_same_font_as_the_painted_labels(self):
        self.show()
        tabs = self.window.tabs
        self.assertEqual(tabs.font().pixelSize(),15)
        self.assertEqual(tabs.elideMode(),self.m.Qt.ElideNone)
        for index in range(tabs.count()):
            with self.subTest(label=tabs.tabText(index)):
                needed = tabs.fontMetrics().horizontalAdvance(tabs.tabText(index))+44
                self.assertGreaterEqual(tabs.tabRect(index).width(),needed)

    def test_project_workspace_matches_request_toolbar_and_bottom_actions(self):
        w = self.window
        w.ensure_bus_session_async = lambda purpose,callback:None
        w.open_new_viewer()
        self.show()
        v = w._newdlg
        pos = lambda widget:widget.mapTo(w,self.m.QPoint(0,0))
        self.assertEqual(v.btn_refresh.text(),'BUS 신규조회')
        self.assertEqual(v.btn_manual.text(),'수동 추가')
        self.assertEqual(v.btn_create.text(),'선택실행')
        self.assertFalse(hasattr(v,'btn_close'))
        self.assertLess(pos(v.toolbar).y(),pos(v.tbl).y())
        self.assertGreater(pos(v.actions_bar).y(),pos(w.logs_panel).y()+w.logs_panel.height()-5)
        self.assertTrue(w.table.verticalHeader().isHidden()==False)
        self.assertFalse(v.tbl.verticalHeader().isHidden())
        self.assertFalse(v.btn_stop.isEnabled())
        self.assertEqual(v.btn_preview.text(),'미리보기')
        self.assertFalse(hasattr(v,'chk_dry'))
        self.assertLess(pos(v.btn_preview).x(),pos(v.btn_create).x())
        w.tabs.setCurrentIndex(0)
        self.assertTrue(v.actions_bar.isHidden())

    def test_toolbar_button_order_and_primary_lookup_actions(self):
        w = self.window
        self.assertFalse(hasattr(w,'theme_switch'))
        self.assertEqual(w.btn_manual.text(),'수동 입력')
        self.assertNotEqual(w.btn_manual.objectName(),'primaryButton')
        self.assertEqual(w.btn_request.objectName(),'primaryButton')
        self.show()
        self.assertLess(w.btn_manual.x(),w.btn_file.x())
        self.assertLess(w.btn_file.x(),w.btn_request.x())
        w.ensure_bus_session_async = lambda purpose,callback:None
        w.open_new_viewer()
        self.app.processEvents()
        v = w._newdlg
        self.assertLess(v.btn_manual.x(),v.btn_refresh.x())
        self.assertEqual(v.btn_refresh.objectName(),'primaryButton')
        self.assertNotEqual(v.btn_manual.objectName(),'primaryButton')

    def test_horizontal_headers_have_no_vertical_separator(self):
        for mode in ('light','dark'):
            self.window.apply_theme(mode)
            style = self.window.styleSheet().split('QHeaderView::section {',1)[1].split('}',1)[0]
            self.assertNotIn('border-right:1px',style)
            self.assertIn('border-bottom:1px',style)

    def test_project_summary_keeps_full_data_in_selected_request_detail(self):
        w = self.window
        v = self.m.NewItemsViewer(w)
        v.set_data(['프로젝트코드','프로젝트명','신청자','요청일'],[['26-012','연구 프로젝트','홍길동','2026-10-22']])
        self.assertTrue(v.tbl.isColumnHidden(v._hidx['신청자']))
        self.assertFalse(v.tbl.isColumnHidden(v._hidx['프로젝트명']))
        v.tbl.setCurrentCell(0,v._hidx['프로젝트코드'])
        self.assertIn('홍길동',v.detail_values['신청자'].text())
        self.assertIn('26-012',v.detail_values['프로젝트코드'].text())
        self.assertEqual(v.tbl.item(0,v._status_col).text(),'실행 대기')
        v._on_project_result(0,dict(status='succeeded',executor_mode='powershell',bus_done=True))
        self.assertEqual(v.tbl.item(0,v._status_col).text(),'BUS 완료')
        self.assertIn('BUS 완료',v.detail_result.text())
        v._manual_add('26-019','수동 연구')
        self.assertEqual(v.detail_values['프로젝트코드'].text(),'26-019')

    def test_notification_and_badge_entry_points_are_disabled(self):
        w = self.window
        self.assertFalse(hasattr(w,'watch_session'))
        self.assertFalse(hasattr(w,'watch_thread'))
        self.assertFalse(hasattr(w,'notify_timer'))
        self.assertFalse(hasattr(w,'tab_count_badges'))
        self.assertNotIn('BUS 요청 알림',[a.text() for a in w.settings_button.menu().actions()])
        w._ensure_project_viewer()
        self.assertFalse(hasattr(w,'new_request_notice'))
        self.assertFalse(hasattr(w._newdlg,'new_request_notice'))
        self.assertTrue(w.session_thread.isRunning())

    def test_screen_log_is_summary_and_file_log_keeps_detail(self):
        from pathlib import Path
        w = self.window
        w.log.clear()
        detail = json.dumps({'steps':[{'target':'detail-only-target'}],'api_key':'test'})
        w._log(detail,summary='프로젝트 26-012: 권한 적용 성공')
        self.assertIn('권한 적용 성공',w.log.toPlainText())
        self.assertNotIn('detail-only-target',w.log.toPlainText())
        path = Path(self.fixture.tmp.name)/'access.log'
        saved = path.read_text(encoding='utf-8')
        self.assertIn('detail-only-target',saved)
        self.assertNotIn('"test"',saved)
        w._clear_log()
        w._log('두 번째 작업')
        self.assertIn('detail-only-target',path.read_text(encoding='utf-8'))
        self.assertIn('두 번째 작업',path.read_text(encoding='utf-8'))
        self.assertEqual(list(Path(self.fixture.tmp.name).glob('access_*.log')),[])

    def test_api_result_logs_a_readable_summary_without_raw_schema(self):
        w = self.window
        row = self.fixture.row()
        w.current_row = row
        w.auto_complete_after_add = False
        w.log.clear()
        w._on_api_result({'status':'failed','error_code':'POWERSHELL_TIMEOUT','steps':[{'target':'raw-only-target'}]})
        text = w.log.toPlainText()
        self.assertNotIn('raw-only-target',text)
        self.assertNotIn('"status"',text)
        self.assertIn('시간 초과',text)

    def test_long_error_message_is_scrollable_without_losing_last_line(self):
        w = self.window
        self.show()
        self.assertIsInstance(w.status_label,self.m.QTextEdit)
        message = ('\\LSK_S010\\Study Folder\\19087\\Study\\MW: The system cannot find the file specified.\n'*20)+'마지막 오류 줄'
        w.status_label.setText(message)
        self.app.processEvents()
        self.assertEqual(w.status_label.toPlainText(),message)
        bar = w.status_label.verticalScrollBar()
        self.assertGreater(bar.maximum(),0)
        bar.setValue(bar.maximum())
        w.status_label.setText(message)
        self.assertEqual(bar.value(),bar.maximum())
        w.status_label.setText('성공')
        self.app.processEvents()
        self.assertEqual(w.status_label.verticalScrollBar().maximum(),0)
        self.assertLess(w.status_label.height(),50)

    def test_disabled_buttons_have_muted_backgrounds_in_both_themes(self):
        w = self.window
        self.show()
        for mode in ('light','dark'):
            w.apply_theme(mode)
            for button in (w.btn_request,w.btn_run_execute,w.btn_manual,w.btn_stop,w.btn_clear_log):
                button.setEnabled(False)
                self.app.processEvents()
                pixel = button.grab().toImage().pixelColor(2,button.height()//2).name()
                self.assertEqual(pixel,self.m.THEMES[mode]['badge'],(mode,button.text()))
            w.btn_request.setEnabled(True)
            self.app.processEvents()
            pixel = w.btn_request.grab().toImage().pixelColor(6,w.btn_request.height()//2).name()
            self.assertEqual(pixel,self.m.THEMES[mode]['accent'])

    def test_replayed_result_explains_that_it_did_not_rerun(self):
        result = dict(status='partially_succeeded',error_code='POWERSHELL_FAILED',replayed=True)
        self.assertIn('이전 작업 결과',self.window._api_log_summary(result,'프로젝트 1987'))
        self.assertIn('재실행 안 함',self.window._api_log_summary(result,'프로젝트 1987'))

    def test_notification_settings_control_is_disabled_but_saved_value_is_preserved(self):
        dialog = self.m.SettingsDialog(self.window,{'notify_refresh_min':15})
        self.assertFalse(hasattr(dialog,'sb_notify'))
        self.assertEqual(dialog.result()[-1],15)
        dialog.close()
