Serve Models Over HTTP
======================

`formulas serve` starts a Flask application and loads the model once at startup.
All API requests reuse that in-memory model.

Starting the server
-------------------

.. code-block:: console

    $ formulas serve test/test_files/excel.xlsx --host 127.0.0.1 --port 5000

Routes
------

- `GET /health`: basic health check.
- `GET /model`: loaded files, workbook ids, and node count.
- `POST /calculate`: calculate outputs for a JSON request payload.

Examples
--------

.. code-block:: console

    $ curl http://127.0.0.1:5000/health

.. code-block:: json

    {
      "status": "ok"
    }

.. code-block:: console

    $ curl http://127.0.0.1:5000/model

.. code-block:: json

    {
      "books": ["EXCEL.XLSX"],
      "files": ["test/test_files/excel.xlsx"],
      "nodes": 16
    }

.. code-block:: console

    $ curl -X POST http://127.0.0.1:5000/calculate \
        -H 'Content-Type: application/json' \
        --data @examples/serve-calculate-request.json

.. code-block:: json

    {
      "inputs": {
        "'[excel.xlsx]'!INPUT_A": 3,
        "'[excel.xlsx]DATA'!B3": 1
      },
      "outputs": {
        "result": 10.0
      },
      "warnings": []
    }

Request body
------------

`POST /calculate` accepts a JSON object with these optional fields:

- `inputs`: object with scalar or range overwrites.
- `outputs`: list of refs to compute.
- `renders`: list of render specs for the JSON response.

The response contains normalized `inputs`, rendered `outputs`, and any warning
messages collected during request processing.
