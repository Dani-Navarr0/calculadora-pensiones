# Calculadora de Pensiones (Spain)

A static website that shows how much time a given amount of money would cover
of Spain's **contributory pension spending** (Seguridad Social), for any year
since 2008 and projected forward. It also has a "close the gap" reform simulator.
Every figure links to its source.

## Structure

```
index.html                 page
assets/app.js              all calculations, chart, simulator, ES/EN texts
assets/style.css           styles (light & dark)
data/pensions.json         monthly payroll 2008 -> latest, AUTO-GENERATED (do not edit)
data/debt.json             Social Security debt, quarterly 1994 -> latest, AUTO-GENERATED (Banco de España)
data/context.json          other figures (minimum wage, salary, reserve fund...) - hand-checked,
                           each with source URL, date and verbatim quote
update_data.py             rebuilds data/pensions.json from the official INSS Excel files
.github/workflows/update_data.yml   runs update_data.py automatically
```

## Data

**Primary series:** the INSS publishes, every month, an Excel file "Pensiones
contributivas en vigor. Número y cuantía por sexo, régimen y clase". `update_data.py`
downloads every one of those files since January 2008 from the
[official archive](https://www.seg-social.es/wps/portal/wss/internet/EstadisticasPresupuestosEstudios/Estadisticas/EST23/2575)
plus the [latest month](https://www.seg-social.es/wps/portal/wss/internet/EstadisticasPresupuestosEstudios/Estadisticas/EST23/EST24),
reads the *Total sistema* row and stores, per month: number of pensions, ordinary
payroll (€) and the URL of the exact file used.

Older files only give number and average pension per sex; there the payroll is
computed as Σ(number × average). Newer files give the amount directly.

**Validation done (2026-11-18):**
* December pension counts 2008-2025 match exactly the official table
  "Serie histórica de pensiones por pensionista" (INSS book *Evolución mensual de las pensiones*, Oct 2026).
* Every 2026 month matches the book's *S_Total* sheet (count and amount).
* Payroll matches official/press figures within 0.003 %: Dec 2015 8,360.4 M€; Dec 2017 8,881 M€;
  Dec 2020 9,985 M€; Dec 2021 10,309 M€; Dec 2023 12,120.8 M€; Jan 2024 12,651.2 M€;
  Dec 2024 12,974.2 M€; Oct 2026 14,539.96 M€.

The script refuses to write the file if any month fails to parse, a month is
missing, or month-to-month changes fall outside the historical range.

**Debt series:** Banco de España, Boletín Estadístico table 12.16
([CSV](https://www.bde.es/webbe/es/estadisticas/compartido/datos/csv/be1216.csv)): debt of the
Social Security Administrations (Excessive Deficit Procedure definition), total, % of GDP and the
part held by the State, end of each quarter since 1994.

**Debt clock:** latest official quarter + average annual increase of the last 3 complete years
(Dec→Dec), spread evenly per second. The real debt rises in steps (one State loan per year), so
the clock is an estimate of the pace, not an official real-time figure; the page says so.

## Method

* **Annual spending** = Σ 12 ordinary monthly payrolls + 2 extra payments (June,
  November). Each extra payment = that month's payroll × `extra_pay_factor`
  (see `data/context.json`; below 1 because work-accident/occupational-disease
  pensions have their extra payments spread over the 12 months).
* **Daily spending** = annual ÷ 365 (366 in leap years).
* **Current year** = published months + remaining months estimated with the latest payroll.
* **Projection** = last complete year × (1+g)^n, g = compound average growth of the last
  5/10/15 years. A trend extrapolation, **not an official forecast**.
* **Scope:** contributory pensions only. Excludes non-contributory pensions and
  *Clases Pasivas* (civil-servant pensions).
* **Simulator:** static first-order arithmetic on official figures; every formula
  is shown on the page. Not a macroeconomic model.

## Run locally

```bash
pip install -r requirements.txt
python update_data.py            # full rebuild (~230 downloads, a few minutes)
python -m http.server 8000       # then open http://localhost:8000
```
(Opening `index.html` directly from disk does not work: browsers block `fetch` of local files.)

## Publish on GitHub Pages

1. Push this folder to a GitHub repository (branch `main`).
2. *Settings → Pages → Build and deployment → Deploy from a branch → `main` / root*.
3. *Settings → Actions → General → Workflow permissions → Read and write* (lets the
   workflow commit updated data).
4. *Actions → Update pension data → Run workflow* once to test.

The workflow checks for new months every Monday and does a full rebuild on the
2nd of each month. `data/context.json` is **not** updated automatically: review
its `date` fields periodically (e.g. new minimum wage each year).
