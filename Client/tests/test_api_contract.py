import unittest
from client_module import load_client


class ContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.client = load_client()

    def test_access_mapping(self):
        values = dict(kind='진행', req='권한해제', user='E123', proj='26-012',
                      lv2='study', lv3='stat', role='Manager')
        p = self.client.make_access_payload(values, 'BUS-1', 'bus', 'operator')
        self.assertEqual(p, dict(request_id='BUS-1', source='bus', operation='revoke',
                                project_status='progress', employee_id='E123',
                                project_code='26012', level2='Study', level3='STAT',
                                role='Manager', requested_by='operator'))
        values.update(kind='종료', req='권한부여')
        p = self.client.make_access_payload(values, 'BUS-1', 'bus', 'operator')
        self.assertEqual(p['operation'], 'grant')
        self.assertEqual(p['project_status'], 'closed')
        self.assertEqual([p[k] for k in ('level2', 'level3', 'role')], [None]*3)

    def test_project_contract_and_invalid_code(self):
        p = self.client.make_project_payload('26-012A1', '프로젝트', 'P1', 'bus', 'operator')
        self.assertEqual(p['project_code'], '26012A1')
        self.assertEqual(p['project_name'], '프로젝트')
        self.assertEqual(set(p), {'request_id','source','project_code','project_name','requested_by'})
        with self.assertRaises(ValueError):
            self.client.make_project_payload('../x', '', 'P1', 'bus', 'operator')

    def test_completion_requires_real_bus_success(self):
        for status in ('simulated', 'preview', 'partially_succeeded', 'failed', 'cancelled', 'queued'):
            self.assertFalse(self.client.can_complete_bus(dict(status=status, executor_mode='powershell'), 'bus'))
        good = dict(status='succeeded', executor_mode='powershell')
        self.assertTrue(self.client.can_complete_bus(good, 'bus'))
        self.assertFalse(self.client.can_complete_bus(good, 'test-client'))
        self.assertFalse(self.client.can_complete_bus(good, 'bus', stopped=True))
        self.assertFalse(self.client.can_complete_bus(dict(status='succeeded', executor_mode='mock'), 'bus'))

    def test_stable_bus_id_distinguishes_operation_and_identity(self):
        values = dict(user='E1', proj='26012', req='권한부여', kind='진행', lv2='Study', lv3='DM')
        first = self.client.make_bus_request_id(values)
        self.assertEqual(first, self.client.make_bus_request_id(dict(values)))
        self.assertNotEqual(first, self.client.make_bus_request_id({**values, 'req':'권한해제'}))
        self.assertNotEqual(self.client.make_bus_request_id(values, 'original-1'),
                            self.client.make_bus_request_id(values, 'original-2'))
        self.assertNotEqual(self.client.new_client_request_id(), self.client.new_client_request_id())


if __name__ == '__main__':
    unittest.main()
