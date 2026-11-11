# Dashboard de operación (web estática)

- `scripts/actualizar_datos.py`: consulta las APIs del Coordinador y escribe `docs/data/*.json`.
- `.github/workflows/actualizar.yml`: ejecuta el script cada día (y al subir un `budgets.xlsx` nuevo, o manualmente).
- `docs/index.html`: la página web (GitHub Pages). Lee los JSON y genera los Excel en el navegador.
- `budgets.xlsx`: base de datos de meses cerrados y budgets (súbelo tú, no está incluido).

Secrets necesarios en el repo: `API_KEY_PRMTE` y `API_KEY_CMG`.
