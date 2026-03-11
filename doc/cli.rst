Command Line Interface
======================

`formulas` ships with a command line interface for executing, exporting,
testing, and serving spreadsheet models.

Installation
------------

Install the package with the dependencies needed for workbook support:

.. code-block:: console

    $ pip install formulas[all]

Common concepts
---------------

- Commands accept mixed `.xlsx`, `.ods`, and `.json` model inputs.
- `--out` and `--outs` reduce the computation scope.
- `--render` and `--renders` reduce the emitted JSON payload.
- CLI scalar values use strict parsing: strings need double quotes and dates must
  use `YYYY-MM-DD`.
- Batch execution uses JSON configuration files.

Commands
--------

- :doc:`cli-calc`: calculate workbook and JSON models.
- :doc:`cli-build`: export merged or reduced JSON models.
- :doc:`cli-test`: validate calculated results against reference workbooks.
- :doc:`cli-serve`: expose a loaded model through a Flask API.
