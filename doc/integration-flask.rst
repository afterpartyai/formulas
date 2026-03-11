Flask + Jinja Demo
==================

This reference integration demo shows how to embed Formulas in a small Flask
application that loads a model once at startup and reuses it for both browser
and API requests.

Demo location
-------------

The example lives in `examples/integrations/flask-demo/`.

Run the demo
------------

From the repository root:

.. code-block:: console

    $ python examples/integrations/flask-demo/app.py

Open `http://127.0.0.1:5000/` in a browser.

Routes
------

- `GET /`: Jinja-based browser interface.
- `GET /health`: health check endpoint.
- `GET /model`: loaded file list, workbook ids, and node count.
- `POST /api/calculate`: JSON calculation endpoint.

Browser interface
-----------------

The browser page provides three JSON text areas:

- `inputs`: object with spreadsheet input overrides.
- `outputs`: optional list of refs to compute.
- `renders`: optional list of render specs.

The page renders calculated outputs, normalized inputs, and warnings on the same
screen so the demo remains small and easy to follow.

Request example
---------------

Sample request body (`examples/integrations/flask-demo/sample-request.json`):

.. code-block:: json

    {
      "inputs": {
        "'[excel.xlsx]'!INPUT_A": 3,
        "'[excel.xlsx]DATA'!B3": 1
      },
      "renders": [
        "'[excel.xlsx]DATA'!C2=result"
      ]
    }

API call example:

.. code-block:: console

    $ curl -X POST http://127.0.0.1:5000/api/calculate \
        -H 'Content-Type: application/json' \
        --data @examples/integrations/flask-demo/sample-request.json

Successful response
-------------------

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

Validation failure example
--------------------------

Invalid request:

.. code-block:: console

    $ curl -X POST http://127.0.0.1:5000/api/calculate \
        -H 'Content-Type: application/json' \
        --data '{"inputs": []}'

Error response:

.. code-block:: json

    {
      "error": "`inputs` must be an object."
    }

How it maps to Formulas
-----------------------

The demo reuses the same model-loading and request-processing logic as the CLI
serve workflow:

- the model is loaded once at startup;
- each request applies input overrides and optional output/render selection;
- responses return normalized inputs, rendered outputs, and warnings.

This keeps the example focused on integration rather than on building a full
spreadsheet interface.
