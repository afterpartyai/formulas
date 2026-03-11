#!/usr/bin/env python
# -*- coding: UTF-8 -*-

import importlib.util
import json
import os
import os.path as osp
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


EXTRAS = os.environ.get('EXTRAS', 'all')
ROOT = Path(__file__).resolve().parents[1]


@unittest.skipIf(EXTRAS not in ('all', 'excel'), 'Not for extra %s.' % EXTRAS)
class TestIntegrations(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp(prefix='formulas-integrations-')

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def run_cmd(self, *args):
        return subprocess.run(
            args,
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            check=False,
        )

    def test_batch_demo_json_command(self):
        out_dir = osp.join(self.tmpdir, 'json')
        result = self.run_cmd(
            sys.executable, '-m', 'formulas.cli', 'calc',
            'test/test_files/excel.xlsx',
            '--batch', 'examples/integrations/batch-demo/scenarios.json',
            '--processes', '2',
            '--output-format', 'json',
            '--output-dir', out_dir,
        )
        self.assertEqual(0, result.returncode, result.stderr)
        summary = json.loads(result.stdout)
        self.assertEqual(['student-alice', 'student-bob'], [item['name'] for item in summary])
        self.assertTrue(osp.isfile(summary[0]['path']))
        with open(summary[0]['path']) as fh:
            self.assertEqual({'result': 10.0}, json.load(fh))

    def test_batch_demo_excel_command(self):
        out_dir = osp.join(self.tmpdir, 'excel')
        result = self.run_cmd(
            sys.executable, '-m', 'formulas.cli', 'calc',
            'test/test_files/excel.xlsx',
            '--batch', 'examples/integrations/batch-demo/scenarios.json',
            '--output-format', 'excel',
            '--output-dir', out_dir,
        )
        self.assertEqual(0, result.returncode, result.stderr)
        summary = json.loads(result.stdout)
        self.assertEqual('ok', summary[0]['status'])
        self.assertTrue(osp.isdir(summary[0]['dir']))

    def test_etl_demo_command(self):
        result = self.run_cmd(sys.executable, 'examples/integrations/etl-demo/transform.py')
        self.assertEqual(0, result.returncode, result.stderr)
        rows = [json.loads(line) for line in result.stdout.splitlines() if line.strip()]
        self.assertEqual([
            {'id': 'row-1', 'result': 10.0},
            {'id': 'row-2', 'result': 11.0},
        ], rows)

    def test_flask_demo_request_examples(self):
        path = ROOT / 'examples' / 'integrations' / 'flask-demo' / 'app.py'
        spec = importlib.util.spec_from_file_location('flask_demo_app', path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        app = module.create_app()
        client = app.test_client()

        health = client.get('/health')
        self.assertEqual(200, health.status_code)
        self.assertEqual({'status': 'ok'}, health.get_json())

        with open(ROOT / 'examples' / 'integrations' / 'flask-demo' / 'sample-request.json') as fh:
            payload = json.load(fh)
        response = client.post('/api/calculate', json=payload)
        self.assertEqual(200, response.status_code)
        self.assertEqual({
            'inputs': {
                "'[excel.xlsx]'!INPUT_A": 3,
                "'[excel.xlsx]DATA'!B3": 1,
            },
            'outputs': {'result': 10.0},
            'warnings': [],
        }, response.get_json())

        invalid = client.post('/api/calculate', json={'inputs': []})
        self.assertEqual(400, invalid.status_code)
        self.assertEqual({'error': '`inputs` must be an object.'}, invalid.get_json())


if __name__ == '__main__':
    unittest.main()
