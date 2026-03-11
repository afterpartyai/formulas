Reference Integration Demos
===========================

Formulas can be embedded as a calculation engine inside web applications,
automation jobs, and data-processing pipelines.

The repository includes small reference demos under `examples/integrations/`.

.. toctree::
    :maxdepth: 1

    integration-flask
    integration-batch
    integration-etl

Flask + Jinja demo
------------------

Path: `examples/integrations/flask-demo/`

This demo loads a model once at startup, exposes HTTP endpoints, and provides a
small browser interface for entering inputs and viewing calculated outputs.

Batch automation demo
---------------------

Path: `examples/integrations/batch-demo/`

This demo shows how to run one workbook across multiple JSON scenarios and save
repeatable JSON or Excel outputs.

ETL transformer demo
--------------------

Path: `examples/integrations/etl-demo/`

This demo treats a spreadsheet model as a transformation step that reads JSONL
records, calculates selected outputs, and emits structured JSONL results.
