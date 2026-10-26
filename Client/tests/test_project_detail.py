import unittest
from unittest.mock import Mock, patch
import test_ui_design as design


class ProjectDetailTests(unittest.TestCase):
    setUp=design.DesignTests.setUp

    def viewer(self):
        w=self.window;w.ensure_bus_session_async=lambda *args:None
        w.open_new_viewer();v=w._newdlg
        v.set_data(['프로젝트코드','프로젝트명','신청자','요청일'],[['26012','First','Tester','2026-10-26'],['26013','Second','Other','2026-10-26']])
        return v

    def job(self, v, states):
        kinds=v.project_step_types[1:-1]
        return dict(job_id='J1',status='running',executor_mode='powershell',steps=[
            dict(type=kind,status=state,target='target-'+kind,description='작업 '+kind,
                 details=dict(template_path='template') if kind=='ensure_project_folder' else {}) for kind,state in zip(kinds,states)])

    def states(self,v):return [icon.property('state') for icon in v.project_step_icons]

    def test_live_steps_show_actual_stage_without_overwriting_authoritative_result(self):
        v=self.viewer();item=v.tbl.item(0,v._hidx['프로젝트코드'])
        meta=item.data(self.m.Qt.UserRole);meta['result']=dict(job_id='old',status='failed',executor_mode='powershell');item.setData(self.m.Qt.UserRole,meta)
        live=self.job(v,['succeeded','running','queued','queued','queued','queued'])
        v._on_project_progress(0,live)
        self.assertEqual(self.states(v),['succeeded','succeeded','running','queued','queued','queued','queued','queued'])
        self.assertEqual(item.data(self.m.Qt.UserRole)['result']['job_id'],'old')
        self.assertEqual(v.detail_values['경로'].text(),'target-ensure_project_folder')
        self.assertEqual(set(v.detail_values),{'프로젝트코드','프로젝트명','경로','신청자'})
        self.assertIn('현재 단계',v.detail_result.text())
        v.tbl.setCurrentCell(1,v._hidx['프로젝트코드']);self.app.processEvents()
        self.assertEqual(self.states(v),['queued']*8)
        v.tbl.setCurrentCell(0,v._hidx['프로젝트코드']);self.app.processEvents()
        self.assertEqual(self.states(v)[2],'running')

    def test_partial_failure_skipped_steps_and_bus_outcome_are_distinct(self):
        v=self.viewer();result=self.job(v,['succeeded','failed','skipped','skipped','skipped','skipped'])
        result.update(status='partially_succeeded',error_code='PATH_NOT_FOUND',error_message='Missing project folder')
        v._on_project_result(0,result)
        self.assertEqual(self.states(v)[1:4],['succeeded','failed','skipped'])
        self.assertEqual(self.states(v)[-1],'queued')
        self.assertIn('Missing project folder',v.detail_result.text())
        result=self.job(v,['succeeded']*6);result.update(status='succeeded',bus_done=True)
        v._on_project_result(0,result)
        self.assertEqual(self.states(v),['succeeded']*8)

    def test_validation_and_mock_never_claim_actual_creation_or_bus_completion(self):
        v=self.viewer()
        v._on_project_result(0,dict(status='validated',validation_status='passed',checks=[]))
        self.assertEqual(self.states(v),['succeeded']+['queued']*7)
        result=self.job(v,['simulated']*6);result.update(status='simulated',executor_mode='mock')
        v._on_project_result(0,result)
        self.assertEqual(self.states(v)[1:-1],['simulated']*6)
        self.assertEqual(self.states(v)[-1],'queued')
        self.assertIn('모의 실행',v.project_step_labels[1].text())

    def test_manual_bus_step_and_interrupted_query_do_not_stay_running(self):
        v=self.viewer();v._manual_add('26014','Manual');v._refresh_project_detail()
        self.assertEqual(self.states(v)[-1],'not_applicable')
        v.tbl.setCurrentCell(0,1);v._on_project_progress(0,dict(status='completion_checking'))
        self.assertEqual(self.states(v)[-1],'running')
        v._project_thread_finished()
        self.assertEqual(self.states(v)[-1],'queued')
        self.assertNotIn('project_action',v.tbl.item(0,1).data(self.m.Qt.UserRole))

    def test_create_worker_forwards_polled_job_steps_and_bus_progress(self):
        w=self.window;m=self.m;events=[]
        worker=m.CreateWorker([dict(proj='26012',name='X',row=0,source='bus',request_id='BUS-X')],dict(id='op',pw='pw'),Mock())
        worker.itemProgress.connect(lambda row,result:events.append((row,result)))
        worker._bus_click_process=Mock(return_value=(True,'완료'))
        def execute(job):
            result=dict(job_id='J1',status='running',executor_mode='powershell',steps=[dict(type='ensure_ad_group',status='running')])
            job.progress.emit(result)
            return dict(result,status='succeeded')
        with patch.object(m.ApiJobWorker,'execute',execute):worker.run()
        self.assertTrue(any(result.get('steps') for row,result in events))
        self.assertTrue(any(result.get('bus_processing') for row,result in events))
        self.assertTrue(all(row==0 for row,result in events))

    def test_long_server_details_scroll_without_changing_list_height(self):
        v=self.viewer();w=self.window;w.resize(1060,900);w.show();self.app.processEvents()
        before=v.tbl.height()
        result=self.job(v,['failed']+['skipped']*5);result.update(status='failed',error_message='error line\n'*80)
        v._on_project_result(0,result)
        for _ in range(5):self.app.processEvents()
        self.assertEqual(v.tbl.height(),before)
        self.assertGreater(v.detail_result_scroll.verticalScrollBar().maximum(),0)
        self.assertEqual(v.detail_result.textFormat(),self.m.Qt.PlainText)
        self.assertEqual(v.detail_scroll.horizontalScrollBar().maximum(),0)
        long_path = r'\\server\share\folder-' + 'long-folder-'*30
        result['steps'][1]['target'] = long_path
        v._on_project_result(0,result)
        for _ in range(5):self.app.processEvents()
        field=v.detail_values['경로']
        self.assertEqual(field.text(),long_path)
        self.assertGreater(field.document().firstBlock().layout().lineCount(),1)

    def test_detail_fields_are_filtered_without_discarding_bus_identity(self):
        v=self.viewer()
        header=['프로젝트코드','프로젝트명','폴더경로','요청자사번','요청자성명','요청자부서/팀','요청일','요청번호','승인자','기타']
        values=['26012','First',r'\\server\share\26012','E1','Tester','DM','2026-10-26','BUS-123','Manager','keep']
        v.set_data(header,[values]);v._refresh_project_detail()
        item=v.tbl.item(0,v._hidx['프로젝트코드']);identity=item.data(self.m.Qt.UserRole)['request_id']
        self.assertEqual(set(v.detail_values),{'프로젝트코드','프로젝트명','경로','요청자사번','요청자성명','요청자부서/팀'})
        self.assertEqual(v.detail_values['경로'].text(),values[2])
        v._toggle_detail_columns(True)
        for name in header:
            self.assertEqual(v.tbl.isColumnHidden(v._hidx[name]),v._project_field_kind(name) is None)
        self.assertEqual(v.tbl.item(0,v._hidx['기타']).text(),'keep')
        self.assertEqual(identity,self.m.make_bus_request_id(dict(proj='26012',project_name='First'),self.m.bus_row_identity(header,values)))
        self.assertNotIn('작업 ID',v.detail_values)
        v._toggle_detail_columns(False)
        self.assertTrue(v.tbl.isColumnHidden(v._hidx['폴더경로']))
        self.assertFalse(v.tbl.isColumnHidden(v._status_col))
