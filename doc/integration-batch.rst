Batch Automation Demo
=====================

This reference integration demo shows how to reuse one spreadsheet model across
multiple scenarios with repeatable JSON or Excel outputs.

Demo location
-------------

The example lives in `examples/integrations/batch-demo/`.

Scenario file
-------------

The demo uses `examples/integrations/batch-demo/scenarios.json`, which defines a
list of run names, input overwrites, and render specs.

JSON output example
-------------------

.. code-block:: console

    $ formulas calc test/test_files/excel.xlsx \
        --batch examples/integrations/batch-demo/scenarios.json \
        --processes 2 \
        --output-format json \
        --output-dir output/json

This produces one JSON result file per scenario, for example:

.. code-block:: json

    [
      {
        "name": "student-alice",
        "path": "output/json/run-0001-student-alice.json",
        "run_id": 1,
        "status": "ok",
        "warnings": []
      },
      {
        "name": "student-bob",
        "path": "output/json/run-0002-student-bob.json",
        "run_id": 2,
        "status": "ok",
        "warnings": []
      }
    ]

Excel output example
--------------------

.. code-block:: console

    $ formulas calc test/test_files/excel.xlsx \
        --batch examples/integrations/batch-demo/scenarios.json \
        --output-format excel \
        --output-dir output/excel

This produces one folder per scenario containing recalculated workbook files.

Integration idea
----------------

Use this pattern when a workbook acts as a reusable template and each batch item
provides different input data, such as report generation or recurring business
processes.
