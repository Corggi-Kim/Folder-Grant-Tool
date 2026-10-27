"""Legacy Isolated group consent, preparation, and revalidation integration."""
from unittest.mock import Mock, patch
import unittest
import test_ui_design as design


class IsolatedPreparationTests(unittest.TestCase):
    setUp=design.DesignTests.setUp

    def payload(self):
        return dict(request_id='BUS-X',source='bus',operation='grant',project_status='progress',employee_id='E1',
                    project_code='26012',level2='Isolated',level3='STAT',role='Randomization Statistician',requested_by='op')

    def validation(self,missing=True):
        return dict(request_id='BUS-X',executor_mode='powershell',validation_status='failed' if missing else 'passed',
                    checks=[dict(id='group',label='보안 그룹',target='LSK 26-012 Isolated',status='failed' if missing else 'passed',message='대상 없음' if missing else '존재')],
                    group_preparation=dict(group_name='LSK 26-012 Isolated',root_path=r'\\server\share\26012',isolated_path=r'\\server\share\26012\Isolated') if missing else None)

    def worker(self):
        api=Mock();api.api_key='test';worker=self.m.ApiJobWorker(api,'access',self.payload(),preflight=True)
        worker.poll_interval=.001
        return worker,api

    def test_consent_prepares_separate_job_then_revalidates_original_before_execution(self):
        worker,api=self.worker();worker.groupPreparationRequested.connect(lambda info:worker.respond_to_group_preparation(True))
        calls=[];prepared=False
        def request(method,path,payload=None):
            nonlocal prepared
            calls.append((method,path,payload))
            if path.endswith('/validate'):return self.validation(not prepared)
            if path.endswith('/prepare-group'):
                prepared=True
                return dict(request_id=payload['request_id'],job_id='PREPARATION',status='succeeded',executor_mode='powershell',
                            steps=[dict(type='prepare_isolated_group')],result=dict(job_kind='group_preparation'))
            self.assertTrue(prepared)
            return dict(request_id='BUS-X',job_id='PERMISSION',status='succeeded',executor_mode='powershell')
        api.request.side_effect=request;events=[];worker.progress.connect(events.append)
        result=worker.execute()
        self.assertEqual([path for method,path,data in calls],['/api/v1/access-jobs/validate','/api/v1/access-jobs/prepare-group','/api/v1/access-jobs/validate','/api/v1/access-jobs'])
        self.assertNotEqual(calls[1][2]['request_id'],'BUS-X')
        self.assertIs(calls[1][2]['confirm_group_creation'],True)
        self.assertEqual(calls[-1][2],self.payload())
        self.assertEqual(result['job_id'],'PERMISSION')
        preparation=next(event['preparation_result'] for event in events if event.get('preparation_result'))
        self.assertFalse(self.m.can_complete_bus(preparation,'bus'))
        self.assertTrue(any(event.get('preparation_payload') for event in events))
        self.assertTrue(any(event.get('clear_preparation') for event in events))

    def test_decline_no_handler_and_stop_never_create_or_register(self):
        for action in ['decline','stop','no_handler']:
            worker,api=self.worker();api.request.return_value=self.validation()
            if action=='decline':worker.groupPreparationRequested.connect(lambda info:worker.respond_to_group_preparation(False))
            elif action=='stop':worker.groupPreparationRequested.connect(lambda info:worker.stop())
            result=worker.execute()
            self.assertTrue(result['preflight_blocked'])
            self.assertEqual(api.request.call_count,1)
            self.assertFalse(self.m.can_complete_bus(result,'bus'))

    def test_invalid_offer_scope_never_prompts_or_prepares(self):
        for change in [dict(level2='Study'),dict(operation='revoke'),dict(project_status='closed')]:
            worker,api=self.worker();worker.payload.update(change);api.request.return_value=self.validation()
            asked=[];worker.groupPreparationRequested.connect(asked.append)
            self.assertTrue(worker.execute()['preflight_blocked'])
            self.assertFalse(asked)
            self.assertEqual(api.request.call_count,1)

    def test_failed_preparation_or_failed_revalidation_never_registers_permission(self):
        for failed_prepare in [True,False]:
            worker,api=self.worker();worker.groupPreparationRequested.connect(lambda info:worker.respond_to_group_preparation(True))
            calls=[]
            def request(method,path,payload=None):
                calls.append(path)
                if path.endswith('/validate'):return self.validation()
                self.assertTrue(path.endswith('/prepare-group'))
                return dict(request_id=payload['request_id'],job_id='PREP',status='failed' if failed_prepare else 'succeeded',executor_mode='powershell')
            api.request.side_effect=request
            result=worker.execute()
            self.assertTrue(result['preflight_blocked'])
            self.assertEqual(calls.count('/api/v1/access-jobs/validate'),1 if failed_prepare else 2)
            self.assertNotIn('/api/v1/access-jobs',calls)

    def test_unknown_preparation_reuses_same_approved_identity_on_next_attempt(self):
        worker,api=self.worker();worker.groupPreparationRequested.connect(lambda info:worker.respond_to_group_preparation(True))
        events=[];worker.progress.connect(events.append)
        def lost(method,path,payload=None):
            if path.endswith('/validate'):return self.validation()
            raise self.m.ApiClientError('offline')
        api.request.side_effect=lost
        self.assertTrue(worker.execute()['preflight_blocked'])
        saved=next(event['preparation_payload'] for event in events if event.get('preparation_payload'))
        retry,api=self.worker();retry.group_preparation_payload=saved
        calls=[]
        def recovered(method,path,payload=None):
            calls.append((path,payload))
            if path.endswith('/prepare-group'):return dict(request_id=payload['request_id'],job_id='PREP',status='succeeded',executor_mode='powershell')
            if path.endswith('/validate'):return self.validation(False)
            return dict(request_id='BUS-X',job_id='J1',status='succeeded',executor_mode='powershell')
        api.request.side_effect=recovered
        self.assertEqual(retry.execute()['job_id'],'J1')
        self.assertEqual(calls[0],('/api/v1/access-jobs/prepare-group',saved))

    def test_preparation_progress_never_becomes_permission_authority(self):
        w=self.window;row=self.fixture.row();w.current_row=row
        previous=dict(job_id='old',status='failed',executor_mode='powershell')
        meta=w._row_metadata(row);meta['result']=previous;w._set_row_metadata(row,meta)
        preparation=dict(job_id='prep',status='succeeded',executor_mode='powershell',result=dict(job_kind='group_preparation'))
        saved=dict(self.payload(),request_id='PREP-ID',confirm_group_creation=True)
        w._on_api_progress(dict(status='preparing',preparation_payload=saved))
        w._on_api_progress(dict(status='preparing',preparation_result=preparation))
        self.assertEqual(w._row_metadata(row)['result'],previous)
        self.assertNotEqual(w._row_metadata(row).get('job_id'),'prep')
        self.assertEqual(w._row_metadata(row)['group_preparation_payload'],saved)
        w._on_api_progress(dict(status='preparing',preparation_result=preparation,clear_preparation=True))
        self.assertNotIn('group_preparation_payload',w._row_metadata(row))
        self.assertFalse(self.fixture.completed)

    def test_gui_session_consent_is_explicit_and_not_persisted(self):
        w=self.window;worker,api=self.worker();w.api_worker=worker
        info=dict(self.validation()['group_preparation'],request_id='BUS-X')
        original=self.m.QMessageBox
        def execute(box):
            box.checkBox().setChecked(True)
            return 0
        def accepted(box):return next(button for button in box.buttons() if box.buttonRole(button)==original.AcceptRole)
        try:
            with patch.object(original,'exec_',execute),patch.object(original,'clickedButton',accepted):
                w._on_group_preparation_requested(info)
            self.assertTrue(worker._group_approved)
            self.assertTrue(w._auto_create_isolated_group)
            second,api=self.worker();w.api_worker=second
            with patch.object(original,'exec_',side_effect=AssertionError('second prompt')):
                w._on_group_preparation_requested(info)
            self.assertTrue(second._group_approved)
        finally:w.api_worker=None

    def test_stop_during_preparation_cancels_only_preparation_job(self):
        worker,api=self.worker();worker.groupPreparationRequested.connect(lambda info:worker.respond_to_group_preparation(True))
        calls=[];identity=None
        def response(method,path,payload=None):
            nonlocal identity
            calls.append((method,path))
            if path.endswith('/validate'):return self.validation()
            if path.endswith('/prepare-group'):
                identity=payload['request_id']
                return dict(request_id=identity,job_id='PREP',status='running',executor_mode='powershell')
            if path.endswith('/cancel'):return dict(status='running')
            return dict(request_id=identity,job_id='PREP',status='cancelled',executor_mode='powershell')
        api.request.side_effect=response
        worker.progress.connect(lambda report:worker.stop() if report.get('preparation_progress') else None)
        self.assertTrue(worker.execute()['preflight_blocked'])
        self.assertIn(('POST','/api/v1/access-jobs/PREP/cancel'),calls)
        self.assertNotIn(('POST','/api/v1/access-jobs'),calls)

    def test_pending_preparation_scope_cannot_be_reused_for_different_employee(self):
        worker,api=self.worker()
        worker.group_preparation_payload=dict(worker.payload,request_id='PREP-ID',confirm_group_creation=True,employee_id='OTHER')
        with self.assertRaises(self.m.ApiClientError):worker.execute()
        api.request.assert_not_called()

    def test_cancel_with_checked_session_option_does_not_remember_approval(self):
        w=self.window;worker,api=self.worker();w.api_worker=worker
        box_type=self.m.QMessageBox
        def execute(box):box.checkBox().setChecked(True);return 0
        def cancelled(box):return next(button for button in box.buttons() if box.buttonRole(button)==box_type.RejectRole)
        try:
            with patch.object(box_type,'exec_',execute),patch.object(box_type,'clickedButton',cancelled):
                w._on_group_preparation_requested(dict(self.validation()['group_preparation'],request_id='BUS-X'))
            self.assertFalse(worker._group_approved)
            self.assertFalse(getattr(w,'_auto_create_isolated_group',False))
        finally:w.api_worker=None

    def test_definite_preparation_rejection_clears_pending_identity(self):
        worker,api=self.worker();worker.groupPreparationRequested.connect(lambda info:worker.respond_to_group_preparation(True))
        events=[];worker.progress.connect(events.append)
        def rejected(method,path,payload=None):
            if path.endswith('/validate'):return self.validation()
            raise self.m.ApiClientError('conditions changed; no job registered',409)
        api.request.side_effect=rejected
        self.assertTrue(worker.execute()['preflight_blocked'])
        self.assertTrue(any(event.get('clear_preparation') for event in events))
        self.assertEqual(api.request.call_count,2)
