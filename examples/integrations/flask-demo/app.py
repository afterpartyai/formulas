#!/usr/bin/env python

import json
from pathlib import Path

from flask import Flask, jsonify, render_template, request

from formulas.cli import CliError, _calculate_from_payload, _load_model, _model_info

ROOT = Path(__file__).resolve().parents[3]
DEFAULT_FILES = (str(ROOT / 'test' / 'test_files' / 'excel.xlsx'),)
DEFAULT_INPUTS = json.dumps({
    "'[excel.xlsx]'!INPUT_A": 3,
    "'[excel.xlsx]DATA'!B3": 1,
}, indent=2)
DEFAULT_OUTPUTS = '[]'
DEFAULT_RENDERS = json.dumps([
    "'[excel.xlsx]DATA'!C2=result",
], indent=2)


def _json_text(value, fallback='{}'):
    return json.dumps(value, indent=2) if value is not None else fallback


def _parse_json_text(raw, fallback):
    raw = raw.strip()
    if not raw:
        return fallback
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise CliError('Invalid JSON payload: %s' % exc.msg) from exc


def create_app(files=DEFAULT_FILES):
    app = Flask(__name__)
    model = _load_model(files)
    app.config['FORMULAS_MODEL'] = model
    app.config['FORMULAS_INFO'] = _model_info(model, files)

    @app.get('/')
    def index():
        return render_template(
            'index.html',
            info=app.config['FORMULAS_INFO'],
            inputs_text=DEFAULT_INPUTS,
            outputs_text=DEFAULT_OUTPUTS,
            renders_text=DEFAULT_RENDERS,
            result=None,
            error=None,
        )

    @app.post('/calculate')
    def calculate_page():
        inputs_text = request.form.get('inputs', DEFAULT_INPUTS)
        outputs_text = request.form.get('outputs', DEFAULT_OUTPUTS)
        renders_text = request.form.get('renders', DEFAULT_RENDERS)
        try:
            payload = {
                'inputs': _parse_json_text(inputs_text, {}),
                'outputs': _parse_json_text(outputs_text, []),
                'renders': _parse_json_text(renders_text, []),
            }
            result = _calculate_from_payload(app.config['FORMULAS_MODEL'], payload)
            error = None
        except CliError as exc:
            result = None
            error = exc.format_message()
        return render_template(
            'index.html',
            info=app.config['FORMULAS_INFO'],
            inputs_text=inputs_text,
            outputs_text=outputs_text,
            renders_text=renders_text,
            result=result,
            error=error,
        )

    @app.get('/health')
    def health():
        return jsonify({'status': 'ok'})

    @app.get('/model')
    def model_info():
        return jsonify(app.config['FORMULAS_INFO'])

    @app.post('/api/calculate')
    def calculate_api():
        try:
            result = _calculate_from_payload(
                app.config['FORMULAS_MODEL'], request.get_json(silent=True)
            )
        except CliError as exc:
            return jsonify({'error': exc.format_message()}), 400
        return jsonify(result)

    return app


app = create_app()


if __name__ == '__main__':
    app.run(debug=True)
