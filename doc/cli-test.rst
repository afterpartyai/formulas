Test Workbook Reproduction
==========================

`formulas test` checks whether formulas reproduces the values in a reference
workbook.

Reference targets
-----------------

When `--against` is omitted, workbook inputs from `FILES` are used as the
default reference targets.

.. code-block:: console

    $ formulas test test/test_files/excel.xlsx

You can also specify explicit targets:

.. code-block:: console

    $ formulas test model.json --against test/test_files/excel.xlsx

Scoped testing
--------------

Use `--out` or `--outs` to compare only selected cells or ranges.

.. code-block:: console

    $ formulas test test/test_files/excel.xlsx \
        --out "'[excel.xlsx]DATA'!C2"

Tolerance options
-----------------

Comparison tolerances are passed to the underlying comparison engine.

.. code-block:: console

    $ formulas test test/test_files/excel.xlsx \
        --tolerance 0.0 \
        --absolute-tolerance 0.000001

Summary output
--------------

Use `--summary` to print a small metrics table before the comparison report.

.. code-block:: console

    $ formulas test test/test_files/excel.xlsx \
        --out "'[excel.xlsx]DATA'!C2" \
        --summary
