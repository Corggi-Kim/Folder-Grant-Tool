"""Rendered layout and interaction contracts for the approved desktop mockup."""
import unittest
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
        self.assertIn('1.1.0',self.window.version_badge.text())
        self.assertIn('2026-10-22',self.window.build_label.text())

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
        self.assertIn('미확인',w.execution_mode_label.text())
        w.current_row = self.fixture.row()
        w._on_api_progress(dict(job_id='J',status='running',executor_mode='mock',steps=[]))
        self.assertIn('연결됨',w.api_status_label.text())
        self.assertIn('모의',w.execution_mode_label.text())

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
        w._newdlg.btn_close.click()
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
