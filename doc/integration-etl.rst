ETL Transformer Demo
====================

This reference integration demo shows how to treat a spreadsheet model as a
transformation step in a structured data pipeline.

Demo location
-------------

The example lives in `examples/integrations/etl-demo/`.

Run the demo
------------

.. code-block:: console

    $ python examples/integrations/etl-demo/transform.py

The script loads the workbook once, reads JSON Lines input records, injects
fields into workbook cells, calculates a selected output, and writes JSON Lines
results to stdout.

Input example
-------------

`examples/integrations/etl-demo/input.jsonl`:

.. code-block:: json

    {"id": "row-1", "input_a": 3, "b3": 1}
    {"id": "row-2", "input_a": 4, "b3": 1}

Output example
--------------

.. code-block:: json

    {"id": "row-1", "result": 10.0}
    {"id": "row-2", "result": 11.0}

Custom input file
-----------------

.. code-block:: console

    $ python examples/integrations/etl-demo/transform.py my-input.jsonl

Integration idea
----------------

Use this pattern when spreadsheet logic should remain the transformation source
of truth while the surrounding workflow exchanges structured records.
