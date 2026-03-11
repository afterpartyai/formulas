Calculate Models
================

`formulas calc` calculates workbook and JSON models and writes results as JSON
or Excel artifacts.

Basic usage
-----------

.. code-block:: console

    $ formulas calc test/test_files/excel.xlsx \
        --render "'[excel.xlsx]DATA'!C2=result" \
        --output-format json

Compute scope
-------------

Use `--out` or `--outs` to reduce what the model computes.

.. code-block:: console

    $ formulas calc test/test_files/excel.xlsx \
        --outs examples/cli-calc-outs.json \
        --output-format json

Rendered output
---------------

Use `--render` or `--renders` to reduce what the CLI emits. Render refs are also
added to the computation scope.

.. code-block:: console

    $ formulas calc test/test_files/excel.xlsx \
        --renders examples/cli-calc-renders.json \
        --output-format json

Overwrites
----------

Scalar overwrites use `CELL=VALUE`.

.. code-block:: console

    $ formulas calc test/test_files/excel.xlsx \
        --overwrite "'[excel.xlsx]'!INPUT_A=3" \
        --overwrite "'[excel.xlsx]DATA'!B3=1" \
        --render "'[excel.xlsx]DATA'!C2=result" \
        --output-format json

Strings must use double quotes. Dates must use `YYYY-MM-DD`. Range overwrites are
available only through `--batch`.

Batch execution
---------------

`--batch` accepts a JSON file with a list of scenarios. `--processes` enables
parallel execution.

.. code-block:: console

    $ formulas calc test/test_files/excel.xlsx \
        --batch examples/cli-calc-batch.json \
        --processes 2 \
        --output-format json \
        --output-dir output

Output modes
------------

- `json`: write to stdout, `--output-file`, or batch files in `--output-dir`.
- `excel`: write one run folder per scenario in `--output-dir`.

.. code-block:: console

    $ formulas calc test/test_files/excel.xlsx \
        --batch examples/cli-calc-batch.json \
        --output-format excel \
        --output-dir output
