import json
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from client_module import load_client


class WorkerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.m = load_client()

    def setUp(self):
        self.calls = []
        self.scenario = 'success'
        self.polls = 0
        owner = self
        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *args):
                pass

            def do_GET(self):
                self.respond()

            def do_POST(self):
                self.respond()

            def respond(self):
                body = self.rfile.read(int(self.headers.get('Content-Length', 0)))
                owner.calls.append((self.command, self.path, json.loads(body) if body else None,
                                    self.headers.get('X-API-Key')))
                code = 200
                data = dict(job_id='job-1', status='queued', executor_mode='powershell')
                if owner.scenario == 'unauthorized':
                    code, data = 401, {'detail': {'message': 'bad key secret-key'}}
                elif self.path.endswith('/preview'):
                    data = {'valid': True, 'steps': []}
                elif self.path.endswith('/cancel'):
                    owner.scenario = 'cancelled'
                    data['status'] = 'running'
                elif self.command == 'POST' and owner.scenario == 'lost':
                    code, data = 503, {'detail': 'lost response'}
                elif '/by-request/' in self.path:
                    data['status'] = 'running'
                elif self.command == 'GET':
                    owner.polls += 1
                    if owner.scenario == 'poll-error':
                        code, data = 500, {'detail': 'db error'}
                    elif owner.scenario == 'invalid':
                        data = []
                    elif owner.scenario == 'cancelled':
                        data['status'] = 'cancelled'
                    elif owner.scenario == 'slow' and owner.polls == 1:
                        threading.Event().wait(0.15)
                        data['status'] = 'running'
                    else:
                        data['status'] = 'succeeded'
                raw = json.dumps(data).encode()
                self.send_response(code)
                self.send_header('Content-Type','application/json')
                self.send_header('Content-Length', str(len(raw)))
                self.end_headers()
                self.wfile.write(raw)
        self.server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
        self.server.daemon_threads = True
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.api = self.m.FolderGrantApiClient(f'http://127.0.0.1:{self.server.server_port}', 'secret-key')

    def tearDown(self):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join()

    def worker(self, kind='access', dry=False):
        worker = self.m.ApiJobWorker(self.api, kind, dict(request_id='REQ-1', requested_by='operator'), dry)
        worker.poll_interval = 0.01
        self.results, self.errors = [], []
        worker.succeeded.connect(self.results.append)
        worker.failed.connect(self.errors.append)
        return worker

    def test_preview_never_registers(self):
        self.worker(dry=True).run()
        self.assertEqual([x[1] for x in self.calls], ['/api/v1/access-jobs/preview'])
        self.assertEqual(self.results[0]['status'], 'preview')

    def test_lost_registration_recovers_by_request(self):
        self.scenario = 'lost'
        self.worker().run()
        self.assertIn('/api/v1/access-jobs/by-request/REQ-1', [x[1] for x in self.calls])
        self.assertEqual(sum(x[0]=='POST' for x in self.calls), 1)
        self.assertEqual(self.results[0]['status'], 'succeeded')

    def test_project_registration_and_auth(self):
        self.worker(kind='project').run()
        self.assertEqual(self.calls[0][1], '/api/v1/project-jobs')
        self.assertTrue(all(x[3]=='secret-key' for x in self.calls))

    def test_stop_during_network_wait_cancels_then_observes_result(self):
        self.scenario = 'slow'
        worker = self.worker()
        timer = threading.Timer(0.04, worker.stop)
        timer.start()
        worker.run()
        timer.join()
        self.assertTrue(any(x[1].endswith('/cancel') for x in self.calls))
        self.assertEqual(self.results[0]['status'], 'cancelled')
        self.assertTrue(self.results[0]['client_stopped'])

    def test_errors_are_not_success_and_key_is_redacted(self):
        for scenario in ('unauthorized', 'poll-error', 'invalid'):
            with self.subTest(scenario=scenario):
                self.scenario = scenario
                self.worker().run()
                self.assertFalse(self.results)
                self.assertTrue(self.errors)
                self.assertNotIn('secret-key', self.errors[0])

    def test_stop_before_register_does_not_send(self):
        worker = self.worker()
        worker.stop()
        worker.run()
        self.assertFalse(self.calls)
        self.assertEqual(self.results[0]['status'], 'cancelled')
