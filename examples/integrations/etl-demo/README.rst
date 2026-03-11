ETL Transformer Demo
====================

This demo shows how to treat a spreadsheet model as a transformation step in a
data pipeline.

The script loads the workbook once, reads JSON Lines input records, injects
fields into workbook cells, calculates a selected output, and writes JSON Lines
results to stdout.

Run the demo
------------

.. code-block:: console

    $ python examples/integrations/etl-demo/transform.py

Or provide a custom input file:

.. code-block:: console

    $ python examples/integrations/etl-demo/transform.py my-input.jsonl

Input format
------------

Each JSON Lines record contains the values mapped into workbook inputs:

.. code-block:: json

    {"id": "row-1", "input_a": 3, "b3": 1}

Output format
-------------

Each output record contains the source id and the computed spreadsheet result:

.. code-block:: json

    {"id": "row-1", "result": 10.0}
