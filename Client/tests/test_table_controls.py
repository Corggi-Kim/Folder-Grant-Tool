"""Regression coverage for list controls, preview-only execution and BUS read-only rows."""
import os
import time
import unittest
from unittest.mock import Mock, patch
import test_ui_design as design


class TableControlTests(unittest.TestCase):
    setUp = design.DesignTests.setUp

    def viewer(self):
        v = self.m.NewItemsViewer(self.window)
        self.window._newdlg = v
        v.set_data(['프로젝트코드','프로젝트명'],[['26012','First'],['26013','Second'],['26014','Third']])
        return v

    def test_project_delete_checked_preserves_remaining_request_identity(self):
        v = self.viewer()
        col = v._hidx['프로젝트코드']
        original = dict(v.tbl.item(1,col).data(self.m.Qt.UserRole))
        v.tbl.cellWidget(1,0).findChild(self.m.QCheckBox).setChecked(False)
        v._delete_checked_rows()
        self.assertEqual(v.tbl.rowCount(),1)
        self.assertEqual(v.tbl.item(0,col).text(),'26013')
        self.assertEqual(v.tbl.item(0,col).data(self.m.Qt.UserRole),original)
        self.assertEqual(v.detail_values['프로젝트코드'].text(),'26013')

    def test_project_delete_current_and_all_clear_inspector(self):
        v = self.viewer()
        v._delete_row(1)
        self.assertEqual(v.tbl.rowCount(),2)
        v._delete_all_rows()
        self.assertEqual(v.tbl.rowCount(),0)
        self.assertEqual(v.detail_values,{})
        self.assertEqual(v.detail_result.text(),'요청을 선택하세요')
        v._manual_add('26015','New')
        self.assertEqual(v.tbl.rowCount(),1)
        self.assertEqual(v.detail_values['프로젝트코드'].text(),'26015')

    def test_project_context_menu_matches_permission_actions(self):
        v = self.viewer()
        pos = v.tbl.visualItemRect(v.tbl.item(0,1)).center()
        labels = []
        def choose(menu,point):
            labels.extend(a.text() for a in menu.actions())
            return menu.actions()[0]
        with patch.object(self.m.QMenu,'exec_',choose):
            v._open_table_menu(pos)
        self.assertEqual(labels,['행 삭제','선택 행 삭제','전체 삭제'])
        self.assertEqual(v.tbl.rowCount(),2)
        self.assertEqual(v.tbl.item(0,1).text(),'26013')

    def test_running_blocks_project_and_access_deletions_and_preview(self):
        v = self.viewer()
        self.fixture.row()
        self.window._set_running_ui(True)
        self.assertFalse(v.btn_preview.isEnabled())
        self.assertEqual(v.tbl.contextMenuPolicy(),self.m.Qt.PreventContextMenu)
        for method in (v._delete_row,v._delete_checked_rows,v._delete_all_rows,
                       self.window._delete_row,self.window._delete_checked_rows,self.window._delete_all_rows):
            method(0) if method.__name__=='_delete_row' else method()
        self.assertEqual(v.tbl.rowCount(),3)
        self.assertEqual(self.window.table.rowCount(),1)
        self.window._set_running_ui(False)
        self.assertTrue(v.btn_preview.isEnabled())
        self.assertEqual(v.tbl.contextMenuPolicy(),self.m.Qt.CustomContextMenu)

    def test_project_preview_uses_preview_endpoint_even_after_cached_success(self):
        v = self.viewer()
        v._delete_row(2);v._delete_row(1)
        v._on_project_result(0,dict(status='succeeded',executor_mode='powershell',job_id='previous'))
        api = Mock();api.api_key='test'
        api.request.return_value=dict(valid=True,steps=[])
        self.window._api_client = Mock(return_value=api)
        with patch.object(self.m.CreateWorker,'_bus_click_process') as bus:
            v.btn_preview.click()
            deadline=time.monotonic()+3
            while v.worker_thread is not None and time.monotonic()<deadline:
                self.app.processEvents();time.sleep(.01)
            self.assertIsNone(v.worker_thread)
            bus.assert_not_called()
        self.assertEqual(api.request.call_count,1)
        self.assertEqual(api.request.call_args.args[:2],('POST','/api/v1/project-jobs/preview'))
        self.assertEqual(v.tbl.item(0,v._status_col).text(),'미리보기')
        self.assertTrue(v.btn_preview.isEnabled())

    def test_bus_rows_read_only_manual_and_excel_rows_stay_editable(self):
        file=os.path.join(self.fixture.tmp.name,'requests.xlsx')
        from openpyxl import Workbook
        book=Workbook();book.active.append(['대상자사번','프로젝트코드','폴더Level2','폴더Level3'])
        book.active.append(['E1','26012','Study','DM']);book.save(file)
        w=self.window
        self.assertTrue(w.load_excel(file,source='bus'))
        request_id=w._row_metadata(0)['request_id']
        for col in w.EDITABLE_COLS:
            self.assertFalse(w.table.item(0,col).flags() & self.m.Qt.ItemIsEditable)
        w._set_running_ui(True);w._set_running_ui(False)
        self.assertFalse(w.table.item(0,w.COL_PROJ).flags() & self.m.Qt.ItemIsEditable)
        self.assertEqual(w._row_metadata(0)['request_id'],request_id)
        self.assertTrue(w.load_excel(file,append=True))
        for col in w.EDITABLE_COLS:
            self.assertTrue(w.table.item(1,col).flags() & self.m.Qt.ItemIsEditable)
        old=w._row_metadata(1)['request_id']
        w.table.item(1,w.COL_PROJ).setText('26013')
        self.assertNotEqual(w._row_metadata(1)['request_id'],old)
        self.assertEqual(w._row_metadata(1)['source'],'test-client')
        row=self.fixture.row(source='test-client')
        self.assertTrue(w.table.item(row,w.COL_USER).flags() & self.m.Qt.ItemIsEditable)

    def test_keyboard_editing_is_blocked_only_for_bus_rows(self):
        from PyQt5.QtTest import QTest
        w=self.window
        bus=self.fixture.row()
        manual=self.fixture.row(source='test-client')
        w.show();self.app.processEvents()
        for row,editable in ((bus,False),(manual,True)):
            w.table.setCurrentCell(row,w.COL_PROJ)
            w.table.setFocus()
            QTest.keyClick(w.table,self.m.Qt.Key_F2)
            self.app.processEvents()
            self.assertEqual(w.table.state()==self.m.QAbstractItemView.EditingState,editable)
            if editable:
                QTest.keyClick(w.table.focusWidget(),self.m.Qt.Key_Escape)
                self.app.processEvents()

    def test_both_delegates_hide_focus_without_removing_keyboard_selection(self):
        v=self.viewer();row=self.fixture.row()
        for table,index in ((self.window.table,self.window.table.model().index(row,self.window.COL_PROJ)),
                            (v.tbl,v.tbl.model().index(0,1))):
            option=self.m.QStyleOptionViewItem()
            option.state |= self.m.QStyle.State_HasFocus
            table.itemDelegate().initStyleOption(option,index)
            self.assertFalse(option.state & self.m.QStyle.State_HasFocus)
            self.assertNotEqual(table.focusPolicy(),self.m.Qt.NoFocus)
