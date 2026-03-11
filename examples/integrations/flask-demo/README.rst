Flask + Jinja Reference Demo
============================

This demo shows how to embed Formulas in a small Flask application.

The app:

- loads a spreadsheet model once at startup;
- exposes JSON endpoints for health, model metadata, and calculation;
- provides a small Jinja-based browser form for manual calculation requests.

Run the demo
------------

From the repository root:

.. code-block:: console

    $ python examples/integrations/flask-demo/app.py

Open `http://127.0.0.1:5000/` in a browser.

Routes
------

- `GET /health`
- `GET /model`
- `POST /api/calculate`
- `GET /` and `POST /calculate` for the browser interface

Example request
---------------

.. code-block:: console

    $ curl -X POST http://127.0.0.1:5000/api/calculate \
        -H 'Content-Type: application/json' \
        --data @examples/integrations/flask-demo/sample-request.json

This demo is intentionally small. It shows how an existing spreadsheet model can
be reused as a callable service without implementing a spreadsheet editor.
