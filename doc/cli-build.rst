Build JSON Models
=================

`formulas build` exports a merged JSON model from workbook and JSON inputs.

Full export
-----------

.. code-block:: console

    $ formulas build test/test_files/excel.xlsx --output-file model.json

Reduced export
--------------

Use `--out` or `--outs` to keep only the requested refs and their dependency
closure.

.. code-block:: console

    $ formulas build test/test_files/excel.xlsx \
        --out "'[excel.xlsx]DATA'!C2" \
        --output-file reduced-model.json

.. code-block:: console

    $ formulas build test/test_files/excel.xlsx tmp/model.json \
        --outs examples/cli-build-outs.json \
        --output-file mixed-model.json

Mixed inputs
------------

Workbook and JSON inputs can be combined in the same export.

Missing refs are logged as warnings. If all requested refs are missing, the
command fails.
