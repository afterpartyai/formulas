Batch Automation Demo
=====================

This demo shows how to reuse one spreadsheet model for multiple scenarios using
the CLI batch mode.

JSON results
------------

.. code-block:: console

    $ formulas calc test/test_files/excel.xlsx \
        --batch examples/integrations/batch-demo/scenarios.json \
        --processes 2 \
        --output-format json \
        --output-dir output/json

Excel results
-------------

.. code-block:: console

    $ formulas calc test/test_files/excel.xlsx \
        --batch examples/integrations/batch-demo/scenarios.json \
        --output-format excel \
        --output-dir output/excel

Use this pattern when a workbook acts as a reusable template and each batch item
provides different input data.
