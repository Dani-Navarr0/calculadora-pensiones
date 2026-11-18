/* Calculadora de Pensiones — all numbers come from data/pensions.json (official
   INSS monthly files, rebuilt by update_data.py) and data/context.json (hand-checked
   figures, each with source URL + verbatim quote). Nothing is hard-coded here. */
"use strict";

const state = { lang: "es", series: null, ctx: null, annual: [], selYear: null };
const $ = (id) => document.getElementById(id);

/* ------------------------------------------------------------------ i18n */
const T = {
  es: {
    navDebt: "Deuda", clockTitle: "Deuda de la Seguridad Social — estimación en tiempo real",
    clockSession: "deuda nueva desde que abriste esta página",
    debtTitle: "La deuda que financia las pensiones", debtNewTitle: "Deuda nueva cada año",
    debtNewSub: "Aumento de la deuda de la Seguridad Social entre el final de un año y el siguiente (miles de millones de €).",
    debtStockTitle: "Deuda acumulada", debtStockSub: "Saldo al final de cada año (último punto: último trimestre publicado).",
    brand: "Calculadora de Pensiones", navCalc: "Calculadora", navTrend: "Evolución", navReform: "Simulador", navSources: "Fuentes",
    title: "¿Cuánto tiempo de pensiones paga tu dinero?",
    subtitle: "Introduce una cantidad y elige un año: te decimos cuánto tiempo cubriría el gasto en pensiones contributivas de la Seguridad Social.",
    amountLabel: "Cantidad (€)", yearLabel: "Año de referencia",
    howSummary: "¿Cómo se calcula? (fórmula y datos exactos)",
    trendTitle: "Gasto anual en pensiones contributivas",
    trendSub: "Euros corrientes (sin ajustar por inflación). Años completos: dato oficial. Año en curso: estimación. Años futuros: proyección de tendencia.",
    projMethod: "Proyección: crecimiento medio de los últimos", years: "años", projUntil: "hasta",
    tableView: "Ver los datos en tabla",
    reformTitle: "Simulador: ¿cómo cerrarías el agujero?",
    reformDisclaimer: "Estimaciones estáticas de primer orden: aplican cada medida a las cifras oficiales actuales sin efectos de comportamiento, macroeconómicos ni de segunda ronda. Sirven para comparar órdenes de magnitud, no son una previsión. Cada fórmula y su fuente se muestran en la medida.",
    sourcesTitle: "Fuentes y metodología",
    pageTitle: "Calculadora de Pensiones",
    invalid: "Introduce una cantidad mayor que cero.",
    kindActual: "dato oficial", kindEstimate: "estimación", kindProjection: "proyección",
    yearOpt: (y, k) => `${y} (${k})`,
    noteActual: (n) => `Suma de las 12 nóminas oficiales del año + 2 pagas extra (${n} meses de datos).`,
    noteEstimate: (n, m) => `Año en curso: ${n} meses oficiales (último: ${m}); el resto se estima repitiendo la última nómina.`,
    noteProjection: (g, w) => `Proyección: gasto del último año completo × (1 + ${g})^años, siendo ${g} el crecimiento medio anual de los últimos ${w} años.`,
    resultLine: (amt, t) => `Tus <strong>${amt}</strong> pagan las pensiones contributivas durante <strong>${t}</strong>.`,
    breakdown: (perDay, perSec, year) => `En ${year} el sistema gasta ${perDay} al día — ${perSec} cada segundo.`,
    units: { y: ["año", "años"], d: ["día", "días"], h: ["hora", "horas"], m: ["minuto", "minutos"], s: ["segundo", "segundos"] },
    and: "y",
    compareTitle: "Para ponerlo en perspectiva, la misma cantidad equivale a:",
    howSteps: (o) => `<ol>
      <li>Gasto anual estimado para <strong>${o.year}</strong> (${o.kind}): <code>${o.annual}</code>. ${o.note}</li>
      <li>Gasto diario = gasto anual ÷ ${o.days} días = <code>${o.daily}</code>.</li>
      <li>Tiempo cubierto = ${o.amount} ÷ ${o.daily} = <code>${o.daysCovered} días</code>.</li>
      <li>Pagas extra: en junio y noviembre se abona una paga extraordinaria. Se calcula como la nómina ordinaria de ese mes × <code>${o.k}</code> (las pensiones de accidente de trabajo y enfermedad profesional tienen la paga extra prorrateada en la mensual). <span class="src">Fuente del factor: <a href="${o.kUrl}" target="_blank" rel="noopener">${o.kSrc}</a></span></li>
      <li>Nómina mensual usada (último mes disponible, ${o.lastMonth}): <code>${o.lastPayroll}</code> — <a href="${o.lastUrl}" target="_blank" rel="noopener">fichero oficial Excel</a>.</li>
    </ol>`,
    scopeNote: "Alcance: solo pensiones <strong>contributivas</strong> del Sistema de la Seguridad Social (jubilación, incapacidad permanente, viudedad, orfandad y favor de familiares). No incluye pensiones no contributivas ni las de Clases Pasivas (funcionarios del Estado), que se pagan con otros presupuestos.",
    projNote: (g, w, from, to) => `Crecimiento medio anual nominal ${from}–${to}: <strong>${g}</strong> (últimos ${w} años). La proyección supone que ese ritmo se mantiene; no es una previsión oficial: no modela demografía, inflación ni reformas.`,
    chartAria: "Gráfico de línea del gasto anual en pensiones contributivas",
    tblYear: "Año", tblSpend: "Gasto anual", tblGrowth: "Variación", tblKind: "Tipo", tblPensions: "Nº pensiones (dic.)",
    gapIntro: (gap, label) => `Punto de partida: <strong>${gap}</strong> al año — ${label}. Activa medidas y mira cuánto cierran.`,
    closed: (pct, eur) => `Cerrado: <strong>${pct}</strong> (${eur} al año)`,
    remaining: (eur) => `Pendiente: ${eur}`,
    summary: (eur, pct, days) => `Tu paquete aporta <strong>${eur} al año</strong>, el <strong>${pct}</strong> del agujero, equivalente a <strong>${days}</strong> de pensiones.`,
    summaryNone: "Mueve los controles para diseñar tu propia reforma.",
    formula: "Fórmula y fuente", tradeoff: "Quién lo paga",
    sourcesIntro: "Cada cifra de esta web enlaza a su fuente. La serie mensual se reconstruye automáticamente desde los ficheros Excel oficiales del INSS; puedes descargarla aquí:",
    download: "Descargar datos (JSON)",
    srcSeries: (n, a, b) => `Nómina mensual de pensiones contributivas: ${n} meses (${a} → ${b}), un fichero oficial por mes.`,
    validated: "Validación: los recuentos de pensiones de diciembre (2008–2025) coinciden exactamente con la tabla oficial «Serie histórica de pensiones por pensionista» del libro «Evolución mensual de las pensiones» del INSS, y los importes coinciden con las notas de prensa oficiales (p. ej. dic. 2015: 8.360 M€; dic. 2024: 12.974 M€).",
    method: `<h3>Metodología</h3><ul>
      <li><strong>Gasto anual</strong> = suma de las 12 nóminas ordinarias mensuales + 2 pagas extraordinarias (junio y noviembre).</li>
      <li><strong>Gasto diario</strong> = gasto anual ÷ días del año (365 o 366).</li>
      <li><strong>Año en curso</strong> = meses publicados + meses pendientes estimados con la última nómina publicada (sin subida adicional).</li>
      <li><strong>Proyección</strong> = último año completo × (1 + g)<sup>n</sup>, con g = crecimiento medio anual compuesto de los últimos 5, 10 o 15 años. Es una extrapolación de tendencia, no una previsión oficial.</li>
      <li>Todas las cifras están en euros corrientes (nominales).</li></ul>`,
    ctxTitle: "Otras cifras usadas (comparaciones y simulador)",
    updated: (d) => `Datos actualizados: ${d}. Actualización automática mensual vía GitHub Actions.`,
    footer: "Proyecto de divulgación independiente. No es una web oficial. Si encuentras un error, abre una incidencia en el repositorio.",
  },
  en: {
    navDebt: "Debt", clockTitle: "Social Security debt — real-time estimate",
    clockSession: "new debt since you opened this page",
    debtTitle: "The debt that funds pensions", debtNewTitle: "New debt every year",
    debtNewSub: "Increase in Social Security debt from the end of one year to the next (€ billion).",
    debtStockTitle: "Accumulated debt", debtStockSub: "Balance at the end of each year (last point: latest published quarter).",
    brand: "Pension Calculator", navCalc: "Calculator", navTrend: "Trend", navReform: "Simulator", navSources: "Sources",
    title: "How much pension time does your money pay for?",
    subtitle: "Enter an amount and pick a year: we tell you how long it would cover Spain's Social Security contributory pension spending.",
    amountLabel: "Amount (€)", yearLabel: "Reference year",
    howSummary: "How is it calculated? (exact formula and data)",
    trendTitle: "Annual spending on contributory pensions",
    trendSub: "Current euros (not inflation-adjusted). Complete years: official data. Current year: estimate. Future years: trend projection.",
    projMethod: "Projection: average growth of the last", years: "years", projUntil: "until",
    tableView: "Show data as a table",
    reformTitle: "Simulator: how would you close the gap?",
    reformDisclaimer: "Static first-order estimates: each measure is applied to current official figures with no behavioural, macroeconomic or second-round effects. They compare orders of magnitude and are not a forecast. Each measure shows its formula and source.",
    sourcesTitle: "Sources & methodology",
    pageTitle: "Pension Calculator",
    invalid: "Enter an amount greater than zero.",
    kindActual: "official data", kindEstimate: "estimate", kindProjection: "projection",
    yearOpt: (y, k) => `${y} (${k})`,
    noteActual: (n) => `Sum of the year's 12 official monthly payrolls + 2 extra payments (${n} months of data).`,
    noteEstimate: (n, m) => `Current year: ${n} official months (latest: ${m}); the rest is estimated by repeating the latest payroll.`,
    noteProjection: (g, w) => `Projection: last complete year × (1 + ${g})^years, where ${g} is the average annual growth of the last ${w} years.`,
    resultLine: (amt, t) => `Your <strong>${amt}</strong> pays contributory pensions for <strong>${t}</strong>.`,
    breakdown: (perDay, perSec, year) => `In ${year} the system spends ${perDay} per day — ${perSec} every second.`,
    units: { y: ["year", "years"], d: ["day", "days"], h: ["hour", "hours"], m: ["minute", "minutes"], s: ["second", "seconds"] },
    and: "and",
    compareTitle: "To put it in perspective, the same amount equals:",
    howSteps: (o) => `<ol>
      <li>Estimated annual spending for <strong>${o.year}</strong> (${o.kind}): <code>${o.annual}</code>. ${o.note}</li>
      <li>Daily spending = annual spending ÷ ${o.days} days = <code>${o.daily}</code>.</li>
      <li>Time covered = ${o.amount} ÷ ${o.daily} = <code>${o.daysCovered} days</code>.</li>
      <li>Extra payments: an extra payment is made in June and November. It is computed as that month's ordinary payroll × <code>${o.k}</code> (work-accident and occupational-disease pensions have the extra payment spread over the monthly amount). <span class="src">Factor source: <a href="${o.kUrl}" target="_blank" rel="noopener">${o.kSrc}</a></span></li>
      <li>Monthly payroll used (latest available month, ${o.lastMonth}): <code>${o.lastPayroll}</code> — <a href="${o.lastUrl}" target="_blank" rel="noopener">official Excel file</a>.</li>
    </ol>`,
    scopeNote: "Scope: only Social Security <strong>contributory</strong> pensions (retirement, permanent disability, widowhood, orphanhood, family dependants). Excludes non-contributory pensions and civil-servant pensions (Clases Pasivas), which are paid from other budgets.",
    projNote: (g, w, from, to) => `Average nominal annual growth ${from}–${to}: <strong>${g}</strong> (last ${w} years). The projection assumes this pace continues; it is not an official forecast and does not model demographics, inflation or reforms.`,
    chartAria: "Line chart of annual contributory pension spending",
    tblYear: "Year", tblSpend: "Annual spending", tblGrowth: "Change", tblKind: "Type", tblPensions: "No. of pensions (Dec.)",
    gapIntro: (gap, label) => `Starting point: <strong>${gap}</strong> per year — ${label}. Switch on measures and see how much they close.`,
    closed: (pct, eur) => `Closed: <strong>${pct}</strong> (${eur} per year)`,
    remaining: (eur) => `Remaining: ${eur}`,
    summary: (eur, pct, days) => `Your package raises <strong>${eur} per year</strong>, <strong>${pct}</strong> of the gap, equal to <strong>${days}</strong> of pensions.`,
    summaryNone: "Move the sliders to design your own reform.",
    formula: "Formula & source", tradeoff: "Who pays",
    sourcesIntro: "Every figure on this site links to its source. The monthly series is rebuilt automatically from the INSS official Excel files; download it here:",
    download: "Download data (JSON)",
    srcSeries: (n, a, b) => `Monthly contributory pension payroll: ${n} months (${a} → ${b}), one official file per month.`,
    validated: "Validation: December pension counts (2008–2025) match exactly the official table “Serie histórica de pensiones por pensionista” in the INSS book “Evolución mensual de las pensiones”, and amounts match official press releases (e.g. Dec 2015: €8,360M; Dec 2024: €12,974M).",
    method: `<h3>Methodology</h3><ul>
      <li><strong>Annual spending</strong> = sum of the 12 ordinary monthly payrolls + 2 extra payments (June and November).</li>
      <li><strong>Daily spending</strong> = annual spending ÷ days in the year (365 or 366).</li>
      <li><strong>Current year</strong> = published months + pending months estimated with the latest published payroll (no further increase).</li>
      <li><strong>Projection</strong> = last complete year × (1 + g)<sup>n</sup>, g = compound average annual growth of the last 5, 10 or 15 years. Trend extrapolation, not an official forecast.</li>
      <li>All figures are in current (nominal) euros.</li></ul>`,
    ctxTitle: "Other figures used (comparisons and simulator)",
    updated: (d) => `Data updated: ${d}. Automatic monthly update via GitHub Actions.`,
    footer: "Independent public-information project. Not an official website. If you find an error, please open an issue in the repository.",
  },
};
const t = () => T[state.lang];
const loc = () => (state.lang === "es" ? "es-ES" : "en-GB");
const MONTHS = {
  es: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"],
  en: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
};

/* --------------------------------------------------------------- format */
// useGrouping "always": Spanish locale otherwise prints 4-digit numbers without separator (7387)
function fmtNum(v, d = 0) { return v.toLocaleString(loc(), { minimumFractionDigits: d, maximumFractionDigits: d, useGrouping: "always" }); }
function fmtEur(v, d = 0) { return v.toLocaleString(loc(), { style: "currency", currency: "EUR", minimumFractionDigits: d, maximumFractionDigits: d, useGrouping: "always" }); }
function fmtBig(v) {
  const a = Math.abs(v);
  if (state.lang === "es") {
    if (a >= 1e12) return `${fmtNum(v / 1e12, 2)} billones de €`;
    if (a >= 1e9) return `${fmtNum(v / 1e6, 0)} millones de €`;
    if (a >= 1e6) return `${fmtNum(v / 1e6, 2)} millones de €`;
    return fmtEur(v, a < 100 ? 2 : 0);
  }
  if (a >= 1e12) return `€${fmtNum(v / 1e12, 2)} trillion`;
  if (a >= 1e9) return `€${fmtNum(v / 1e9, 2)} billion`;
  if (a >= 1e6) return `€${fmtNum(v / 1e6, 2)} million`;
  return fmtEur(v, a < 100 ? 2 : 0);
}
function fmtPct(v, d = 1) { return (v * 100).toLocaleString(loc(), { minimumFractionDigits: d, maximumFractionDigits: d }) + " %"; }
function monthName(y, m) { return `${MONTHS[state.lang][m - 1]} ${y}`; }
function plural(n, pair) { return Math.abs(n - 1) < 1e-9 ? pair[0] : pair[1]; }

/** Human duration from a number of days: "3 días y 4 horas", "12 minutos y 5 segundos"... */
function fmtDuration(days) {
  const u = t().units;
  const secs = days * 86400;
  if (secs < 60) return `${fmtNum(secs, secs < 10 ? 2 : 1)} ${plural(secs, u.s)}`;
  const parts = [];
  let rest = secs;
  const unitList = [["y", 365 * 86400], ["d", 86400], ["h", 3600], ["m", 60], ["s", 1]];
  for (const [k, size] of unitList) {
    const n = Math.floor(rest / size);
    if (n > 0 || parts.length) { if (n > 0) parts.push(`${fmtNum(n)} ${plural(n, u[k])}`); rest -= n * size; }
    if (parts.length === 2) break;
  }
  return parts.join(` ${t().and} `);
}

/** Accepts "1.000.000", "1,000,000", "1e6", "2,5" etc. Spanish style: "." thousands, "," decimals. */
function parseAmount(raw) {
  if (!raw) return NaN;
  let s = String(raw).trim().replace(/[€\s]/g, "");
  if (/^\d+(\.\d+)?e\d+$/i.test(s)) return Number(s);
  const hasComma = s.includes(","), hasDot = s.includes(".");
  if (hasComma && hasDot) {
    s = s.lastIndexOf(",") > s.lastIndexOf(".") ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
  } else if (hasComma) {
    s = /,\d{3}(,|$)/.test(s) && !/,\d{1,2}$/.test(s) ? s.replace(/,/g, "") : s.replace(",", ".");
  } else if (hasDot) {
    if (/\.\d{3}(\.|$)/.test(s) && (s.match(/\./g).length > 1 || !/\.\d{1,2}$/.test(s))) s = s.replace(/\./g, "");
  }
  return Number(s);
}

/* --------------------------------------------------------------- model */
const ctxVal = (key) => (state.ctx.items[key] ? state.ctx.items[key].value : null);
const daysIn = (y) => ((y % 4 === 0 && y % 100 !== 0) || y % 400 === 0 ? 366 : 365);

function buildAnnual() {
  const k = ctxVal("extra_pay_factor") ?? 1;
  const byYear = {};
  for (const r of state.series.monthly) (byYear[r.year] ||= {})[r.month] = r;
  const years = Object.keys(byYear).map(Number).sort((a, b) => a - b);
  const out = [];
  for (const y of years) {
    const ms = byYear[y];
    const n = Object.keys(ms).length;
    const last = ms[Math.max(...Object.keys(ms).map(Number))];
    const get = (m) => (ms[m] ? ms[m].payroll_eur : last.payroll_eur);
    let total = 0;
    for (let m = 1; m <= 12; m++) total += get(m);
    total += k * (get(6) + get(11));
    out.push({ year: y, spend: total, kind: n === 12 ? "actual" : "estimate", months: n, lastMonth: last.month,
               pensionsDec: ms[12] ? ms[12].pensions : null });
  }
  state.annual = out;
}

function growthRate(window) {
  const actual = state.annual.filter((a) => a.kind === "actual");
  const last = actual[actual.length - 1];
  const first = actual.find((a) => a.year === last.year - window) || actual[0];
  const n = last.year - first.year;
  return { g: Math.pow(last.spend / first.spend, 1 / n) - 1, from: first.year, to: last.year, base: last };
}

function yearInfo(y) {
  const a = state.annual.find((r) => r.year === y);
  if (a) return a;
  const w = Number($("projWindow").value);
  const { g, base } = growthRate(w);
  return { year: y, spend: base.spend * Math.pow(1 + g, y - base.year), kind: "projection", g, w };
}

/* --------------------------------------------------------------- calculator */
function presets() {
  const it = state.ctx.items;
  const list = [];
  const add = (key, val) => { if (it[key] && val) list.push({ key, val }); };
  add("smi_monthly", it.smi_monthly && it.smi_monthly.value);
  add("avg_salary_annual", it.avg_salary_annual && it.avg_salary_annual.value);
  for (const key of ["loan_cap_annual", "reserve_fund_current", "ss_debt"]) add(key, it[key] && it[key].value);
  return list;
}

function renderPresets() {
  const box = $("presetContainer");
  box.innerHTML = "";
  for (const p of presets()) {
    const b = document.createElement("button");
    b.className = "preset-btn";
    b.textContent = state.ctx.items[p.key]["short_" + state.lang];
    b.title = fmtBig(p.val);
    b.onclick = () => { $("moneyInput").value = fmtNum(p.val); calculate(); };
    box.appendChild(b);
  }
}

function comparisons(val) {
  const it = state.ctx.items, out = [];
  const last = state.series.monthly[state.series.monthly.length - 1];
  const avgPension = last.payroll_eur / last.pensions;
  const push = (n, es, en, url, src) => n >= 1 && out.push(`${fmtNum(Math.floor(n))} ${state.lang === "es" ? es : en} <span class="src">(<a href="${url}" target="_blank" rel="noopener">${src}</a>)</span>`);
  push(val / (avgPension * 14),
    `pensiones contributivas medias durante un año (${fmtEur(avgPension, 2)}/mes × 14 pagas, ${monthName(last.year, last.month)})`,
    `average contributory pensions for a whole year (${fmtEur(avgPension, 2)}/month × 14 payments, ${monthName(last.year, last.month)})`,
    last.source, "INSS");
  if (it.avg_salary_annual) push(val / it.avg_salary_annual.value,
    `salarios medios anuales (${fmtEur(it.avg_salary_annual.value, 0)}, ${it.avg_salary_annual.date})`,
    `average annual salaries (${fmtEur(it.avg_salary_annual.value, 0)}, ${it.avg_salary_annual.date})`,
    it.avg_salary_annual.source_url, it.avg_salary_annual.source_name);
  if (it.smi_monthly) push(val / it.smi_monthly.value,
    `mensualidades del salario mínimo (${fmtEur(it.smi_monthly.value, 0)}, ${it.smi_monthly.date})`,
    `monthly minimum-wage payments (${fmtEur(it.smi_monthly.value, 0)}, ${it.smi_monthly.date})`,
    it.smi_monthly.source_url, it.smi_monthly.source_name);
  return out;
}

function calculate() {
  const L = t();
  const val = parseAmount($("moneyInput").value);
  const res = $("result");
  $("inputPreview").textContent = val > 0 ? fmtBig(val) : "";
  if (!(val > 0)) { res.innerHTML = $("moneyInput").value ? `<p class="error">${L.invalid}</p>` : ""; $("howBody").innerHTML = ""; return; }
  const y = state.selYear;
  const info = yearInfo(y);
  const daily = info.spend / daysIn(y);
  const days = val / daily;
  const kindLbl = { actual: L.kindActual, estimate: L.kindEstimate, projection: L.kindProjection }[info.kind];
  const comps = comparisons(val);
  res.innerHTML = `<div class="big">${L.resultLine(fmtBig(val), fmtDuration(days))}</div>
    <div class="breakdown">${L.breakdown(fmtBig(daily), fmtEur(daily / 86400, 0), `${y} (${kindLbl})`)}</div>
    ${comps.length ? `<div class="compare">${L.compareTitle}<ul>${comps.map((c) => `<li>${c}</li>`).join("")}</ul></div>` : ""}`;

  const last = state.series.monthly[state.series.monthly.length - 1];
  const kItem = state.ctx.items.extra_pay_factor;
  $("howBody").innerHTML = L.howSteps({
    year: y, kind: kindLbl, annual: fmtEur(info.spend), note: yearNote(info), days: daysIn(y),
    daily: fmtEur(daily), amount: fmtEur(val), daysCovered: fmtNum(days, days < 1 ? 6 : 2),
    k: kItem ? fmtNum(kItem.value, 4) : "1", kUrl: kItem ? kItem.source_url : "#", kSrc: kItem ? kItem.source_name : "—",
    lastMonth: monthName(last.year, last.month), lastPayroll: fmtEur(last.payroll_eur), lastUrl: last.source,
  }) + `<p class="note">${L.scopeNote}</p>` +
    (state.ctx.items.executed_contributory_spending_2025 ? `<p class="note">${state.ctx.items.executed_contributory_spending_2025["note_" + state.lang]}</p>` : "");
}

function yearNote(info) {
  const L = t();
  if (info.kind === "actual") return L.noteActual(info.months);
  if (info.kind === "estimate") return L.noteEstimate(info.months, monthName(info.year, info.lastMonth));
  return L.noteProjection(fmtPct(info.g, 2), info.w);
}

function renderYearSelect() {
  const L = t(), sel = $("yearSelect");
  const until = Number($("projUntil").value);
  const prev = state.selYear;
  sel.innerHTML = "";
  const all = [...state.annual.map((a) => ({ y: a.year, k: a.kind }))];
  const lastY = state.annual[state.annual.length - 1].year;
  for (let y = lastY + 1; y <= until; y++) all.push({ y, k: "projection" });
  for (const o of all.reverse()) {
    const op = document.createElement("option");
    op.value = o.y;
    op.textContent = L.yearOpt(o.y, { actual: L.kindActual, estimate: L.kindEstimate, projection: L.kindProjection }[o.k]);
    sel.appendChild(op);
  }
  state.selYear = prev && prev <= until ? prev : lastY;
  sel.value = state.selYear;
  $("yearNote").textContent = yearNote(yearInfo(state.selYear));
}

/* --------------------------------------------------------------- chart */
function renderChart() {
  const L = t();
  const w = Number($("projWindow").value), until = Number($("projUntil").value);
  const { g, from, to } = growthRate(w);
  const pts = state.annual.map((a) => ({ ...a }));
  const lastY = pts[pts.length - 1].year;
  for (let y = lastY + 1; y <= until; y++) pts.push(yearInfo(y));
  // draw at the container's real pixel width so text stays legible on phones
  const W = Math.max(300, Math.round($("chart").clientWidth || 860));
  const H = W < 600 ? 240 : 320, m = { l: 58, r: 12, t: 24, b: 32 };
  const x0 = pts[0].year, x1 = pts[pts.length - 1].year;
  const yMax = Math.max(...pts.map((p) => p.spend)) * 1.08;
  const X = (y) => m.l + ((y - x0) / (x1 - x0)) * (W - m.l - m.r);
  const Y = (v) => H - m.b - (v / yMax) * (H - m.t - m.b);
  const step = yMax > 4e11 ? 1e11 : 5e10;
  let grid = "";
  for (let v = 0; v <= yMax; v += step) {
    grid += `<line x1="${m.l}" x2="${W - m.r}" y1="${Y(v)}" y2="${Y(v)}"/>`;
    grid += `<text x="${m.l - 8}" y="${Y(v) + 4}" text-anchor="end">${fmtNum(v / 1e9)}${state.lang === "es" ? " mM€" : " bn€"}</text>`;
  }
  let xt = "";
  const xStep = W < 600 ? (x1 - x0 > 24 ? 8 : 6) : (x1 - x0 > 24 ? 4 : 2);
  for (let y = x0; y <= x1; y++) if (((y - x0) % xStep === 0 && x1 - y >= xStep / 2) || y === x1) xt += `<text x="${X(y)}" y="${H - 10}" text-anchor="middle">${y}</text>`;
  const actual = pts.filter((p) => p.kind === "actual");
  const lastAct = actual[actual.length - 1];
  const path = (arr) => arr.map((p, i) => `${i ? "L" : "M"}${X(p.year).toFixed(1)},${Y(p.spend).toFixed(1)}`).join("");
  const projPts = [lastAct, ...pts.filter((p) => p.kind === "projection")];
  const est = pts.find((p) => p.kind === "estimate");
  const sel = pts.find((p) => p.year === state.selYear);
  const endP = pts[pts.length - 1];
  const svg = `<svg viewBox="0 0 ${W} ${H}" aria-label="${L.chartAria}">
    <g class="grid">${grid.replace(/<text[^>]*>[^<]*<\/text>/g, "")}</g>
    <g class="axis">${grid.replace(/<line[^>]*\/>/g, "")}${xt}</g>
    <path class="line-hist" d="${path(actual)}"/>
    ${projPts.length > 1 ? `<path class="line-proj" d="${path(projPts)}"/>` : ""}
    ${est ? `<circle class="dot est" cx="${X(est.year)}" cy="${Y(est.spend)}" r="4"/>` : ""}
    <circle class="dot" cx="${X(lastAct.year)}" cy="${Y(lastAct.spend)}" r="4"/>
    <text class="lbl" x="${X(lastAct.year) - 6}" y="${Y(lastAct.spend) - 10}" text-anchor="end">${lastAct.year}: ${fmtBig(lastAct.spend)}</text>
    ${endP.kind === "projection" ? `<text class="lbl-2" x="${X(endP.year)}" y="${Y(endP.spend) - 10}" text-anchor="end">${endP.year}: ${fmtBig(endP.spend)}</text>` : ""}
    ${sel ? `<circle class="sel-year" cx="${X(sel.year)}" cy="${Y(sel.spend)}" r="5"/>` : ""}
    <line class="cross" id="cross" y1="${m.t}" y2="${H - m.b}" visibility="hidden"/>
    <rect id="hit" x="${m.l}" y="${m.t}" width="${W - m.l - m.r}" height="${H - m.t - m.b}" fill="transparent"/>
  </svg>`;
  $("chart").innerHTML = svg;
  const off = ["ageing_report_2050", "airef_2050"].map((k) => state.ctx.items[k]).filter(Boolean)
    .map((v) => `<li>${v["label_" + state.lang]} — <a href="${v.source_url}" target="_blank" rel="noopener">${v.source_name}</a></li>`).join("");
  $("projNote").innerHTML = L.projNote(fmtPct(g, 2), w, from, to) +
    (off ? `<br>${state.lang === "es" ? "Proyecciones oficiales de referencia (en % del PIB, no comparables directamente con euros):" : "Official reference projections (% of GDP, not directly comparable with euros):"}<ul>${off}</ul>` : "");

  // hover: crosshair + tooltip on nearest year
  const svgEl = $("chart").querySelector("svg"), hit = svgEl.querySelector("#hit"), cross = svgEl.querySelector("#cross"), tip = $("tooltip");
  const kindLbl = { actual: L.kindActual, estimate: L.kindEstimate, projection: L.kindProjection };
  const move = (ev) => {
    const r = svgEl.getBoundingClientRect();
    const sx = ((ev.clientX - r.left) / r.width) * W;
    const yr = Math.round(x0 + ((sx - m.l) / (W - m.l - m.r)) * (x1 - x0));
    const p = pts.find((q) => q.year === Math.max(x0, Math.min(x1, yr)));
    if (!p) return;
    cross.setAttribute("x1", X(p.year)); cross.setAttribute("x2", X(p.year)); cross.setAttribute("visibility", "visible");
    tip.hidden = false;
    tip.innerHTML = `<strong>${p.year}</strong> · ${kindLbl[p.kind]}<br>${fmtBig(p.spend)}<br><span class="src">${fmtBig(p.spend / daysIn(p.year))} / ${state.lang === "es" ? "día" : "day"}</span>`;
    tip.style.left = Math.min(ev.clientX + 14, window.innerWidth - 270) + "px";
    tip.style.top = ev.clientY + 14 + "px";
  };
  hit.addEventListener("pointermove", move);
  hit.addEventListener("pointerleave", () => { tip.hidden = true; cross.setAttribute("visibility", "hidden"); });
  hit.addEventListener("click", (ev) => {
    const r = svgEl.getBoundingClientRect();
    const yr = Math.round(x0 + ((((ev.clientX - r.left) / r.width) * W - m.l) / (W - m.l - m.r)) * (x1 - x0));
    state.selYear = Math.max(x0, Math.min(x1, yr)); $("yearSelect").value = state.selYear; refresh();
  });

  // table view
  let rows = `<tr><th>${L.tblYear}</th><th>${L.tblSpend}</th><th>${L.tblGrowth}</th><th>${L.tblPensions}</th><th>${L.tblKind}</th></tr>`;
  pts.forEach((p, i) => {
    const gr = i ? p.spend / pts[i - 1].spend - 1 : null;
    rows += `<tr><td>${p.year}</td><td>${fmtEur(p.spend)}</td><td>${gr === null ? "—" : fmtPct(gr)}</td><td>${p.pensionsDec ? fmtNum(p.pensionsDec) : "—"}</td><td>${kindLbl[p.kind]}</td></tr>`;
  });
  $("dataTable").innerHTML = rows;
}

/* --------------------------------------------------------------- debt (Banco de España) */
const YEAR_MS = 365.2425 * 86400 * 1000;

/** Everything the clock and the debt section need, derived from data/debt.json. */
function debtModel() {
  const q = state.debt.quarterly;
  const last = q[q.length - 1];
  const decs = q.filter((r) => r.quarter_end_month === 12);
  const lastDec = decs[decs.length - 1];
  const win = 3;  // growth window: average of the last 3 complete years
  const decBase = decs.find((r) => r.year === lastDec.year - win);
  const ratePerYear = (lastDec.debt_eur - decBase.debt_eur) / win;
  // official figure is the balance on the last day of the quarter
  const baseTime = Date.UTC(last.year, last.quarter_end_month, 0, 23, 59, 59);
  const pre = decs.filter((r) => r.year <= 2016).pop();
  return { q, last, decs, lastDec, decBase, win, ratePerYear, perSec: ratePerYear / (YEAR_MS / 1000), baseTime, pre };
}

function debtNow(dm, now = Date.now()) {
  return dm.last.debt_eur + dm.ratePerYear * ((now - dm.baseTime) / YEAR_MS);
}

function qLabel(r) { return `${MONTHS[state.lang][r.quarter_end_month - 1]} ${r.year}`; }

function renderClockText() {
  const L = t(), dm = state.dm, es = state.lang === "es";
  const cur = state.annual[state.annual.length - 1];
  const spendPerSec = cur.spend / (daysIn(cur.year) * 86400);
  const nw = (x) => `<span style="white-space:nowrap">${x}</span>`;
  $("clockSub").innerHTML = (es
    ? [`+${fmtEur(dm.perSec, 0)} cada segundo`, `+${fmtBig(dm.ratePerYear / 365.2425)} al día`, `+${fmtBig(dm.ratePerYear)} al año`]
    : [`+${fmtEur(dm.perSec, 0)} every second`, `+${fmtBig(dm.ratePerYear / 365.2425)} per day`, `+${fmtBig(dm.ratePerYear)} per year`]).map(nw).join(" · ");
  $("clockSpendLbl").textContent = es ? `gastado en pensiones contributivas en ${cur.year} (estimación, ${fmtEur(spendPerSec, 0)}/s)` : `spent on contributory pensions in ${cur.year} (estimate, ${fmtEur(spendPerSec, 0)}/s)`;
  const src = `<a href="${state.debt.source_page}" target="_blank" rel="noopener">Banco de España, Boletín Estadístico 12.16</a>`;
  $("clockNote").innerHTML = es
    ? `Último dato oficial: <strong>${fmtBig(dm.last.debt_eur)}</strong> (${qLabel(dm.last)}, ${src}). El contador suma desde entonces el ritmo medio de ${dm.decBase.year + 1}–${dm.lastDec.year} (${fmtBig(dm.ratePerYear)}/año). En la realidad la deuda sube a saltos, cuando el Estado concede su préstamo anual; el contador lo reparte de forma continua. Es una estimación, no un dato oficial en tiempo real.`
    : `Latest official figure: <strong>${fmtBig(dm.last.debt_eur)}</strong> (${qLabel(dm.last)}, ${src}). Since then the clock adds the average pace of ${dm.decBase.year + 1}–${dm.lastDec.year} (${fmtBig(dm.ratePerYear)}/year). In reality the debt rises in steps, when the State grants its annual loan; the clock spreads that pace evenly. It is an estimate, not an official real-time figure.`;
  $("clockSr").textContent = es
    ? `Deuda estimada de la Seguridad Social: unos ${fmtBig(debtNow(dm))}, creciendo ${fmtEur(dm.perSec, 0)} por segundo.`
    : `Estimated Social Security debt: about ${fmtBig(debtNow(dm))}, growing ${fmtEur(dm.perSec, 0)} per second.`;
}

function startClock() {
  const dm = state.dm, opened = Date.now();
  const cur = state.annual[state.annual.length - 1];
  const yStart = Date.UTC(cur.year, 0, 1), yLen = daysIn(cur.year) * 86400 * 1000;
  const reduce = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  const tick = () => {
    const now = Date.now();
    $("clockDebt").textContent = fmtEur(debtNow(dm, now), 0);
    $("clockSession").textContent = "+" + fmtEur(dm.ratePerYear * ((now - opened) / YEAR_MS), 0);
    const frac = Math.min(Math.max((now - yStart) / yLen, 0), 1);
    $("clockSpend").textContent = fmtEur(cur.spend * frac, 0);
  };
  tick();
  clearInterval(state.clockTimer);
  state.clockTimer = setInterval(tick, reduce ? 1000 : 100);
}

/** Small reusable SVG chart: kind "bar" or "line"; pts = [{x: label, v: value, tip: html, hl: bool}] */
function miniChart(el, pts, kind, yFmt) {
  const W = Math.max(300, Math.round(el.clientWidth || 860)), H = W < 600 ? 220 : 280;
  const m = { l: 52, r: 12, t: 18, b: 30 };
  const vals = pts.map((p) => p.v);
  const vMax = Math.max(...vals, 0) * 1.1, vMin = Math.min(...vals, 0);
  const iw = W - m.l - m.r, ih = H - m.t - m.b, n = pts.length;
  const Y = (v) => m.t + ih - ((v - vMin) / (vMax - vMin)) * ih;
  const slot = iw / n;
  const X = (i) => m.l + slot * (i + 0.5);
  const raw = (vMax - vMin) / 4, mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const step = [1, 2, 2.5, 5, 10].map((k) => k * mag).find((s) => s >= raw);
  let g = "", ax = "";
  for (let v = Math.ceil(vMin / step) * step; v <= vMax; v += step) {
    g += `<line x1="${m.l}" x2="${W - m.r}" y1="${Y(v)}" y2="${Y(v)}"/>`;
    ax += `<text x="${m.l - 6}" y="${Y(v) + 4}" text-anchor="end">${yFmt(Math.abs(v) < step / 1e6 ? 0 : v)}</text>`;
  }
  const every = Math.ceil(n / (W < 600 ? 5 : 10));
  pts.forEach((p, i) => { if ((n - 1 - i) % every === 0) ax += `<text x="${X(i)}" y="${H - 9}" text-anchor="middle">${p.x}</text>`; });
  let marks = "";
  if (kind === "bar") {
    const bw = Math.max(2, slot - 2);  // 2px gap between bars
    pts.forEach((p, i) => {
      const y0 = Y(0), y1 = Y(p.v), top = Math.min(y0, y1), h = Math.max(Math.abs(y0 - y1), p.v ? 1 : 0);
      const r = Math.min(4, bw / 2, h);
      const x = X(i) - bw / 2;
      // rounded data end (top), square at the baseline
      marks += p.v >= 0
        ? `<path class="bar${p.hl ? " hl" : ""}" data-i="${i}" d="M${x},${y0}V${top + r}Q${x},${top} ${x + r},${top}H${x + bw - r}Q${x + bw},${top} ${x + bw},${top + r}V${y0}Z"/>`
        : `<rect class="bar" data-i="${i}" x="${x}" y="${top}" width="${bw}" height="${h}"/>`;
    });
  } else {
    marks = `<path class="line-hist" d="${pts.map((p, i) => `${i ? "L" : "M"}${X(i).toFixed(1)},${Y(p.v).toFixed(1)}`).join("")}"/>`;
    const lp = pts[n - 1];
    marks += `<circle class="dot" cx="${X(n - 1)}" cy="${Y(lp.v)}" r="4"/><text class="lbl" x="${X(n - 1) - 6}" y="${Y(lp.v) - 10}" text-anchor="end">${lp.x}: ${fmtBig(lp.v)}</text>`;
  }
  el.innerHTML = `<svg viewBox="0 0 ${W} ${H}"><g class="grid">${g}</g><g class="axis">${ax}</g>${marks}
    <line class="cross" y1="${m.t}" y2="${H - m.b}" visibility="hidden"/><rect class="hit" x="${m.l}" y="${m.t}" width="${iw}" height="${ih}" fill="transparent"/></svg>`;
  const svg = el.querySelector("svg"), hit = svg.querySelector(".hit"), cross = svg.querySelector(".cross"), tip = $("tooltip");
  hit.addEventListener("pointermove", (ev) => {
    const r = svg.getBoundingClientRect();
    const i = Math.max(0, Math.min(n - 1, Math.floor((((ev.clientX - r.left) / r.width) * W - m.l) / slot)));
    cross.setAttribute("x1", X(i)); cross.setAttribute("x2", X(i)); cross.setAttribute("visibility", "visible");
    tip.hidden = false; tip.innerHTML = pts[i].tip;
    tip.style.left = Math.min(ev.clientX + 14, window.innerWidth - 270) + "px"; tip.style.top = ev.clientY + 14 + "px";
  });
  hit.addEventListener("pointerleave", () => { tip.hidden = true; cross.setAttribute("visibility", "hidden"); });
}

function renderDebt() {
  const dm = state.dm, es = state.lang === "es", it = state.ctx.items;
  const last = dm.last, mult = last.debt_eur / dm.pre.debt_eur;
  const shareState = last.held_by_state_eur / last.debt_eur;
  const cur = state.annual[state.annual.length - 1];
  const pensionShare = dm.ratePerYear / state.annual.filter((a) => a.kind === "actual").pop().spend;
  const bdeLink = `<a href="${state.debt.source_page}" target="_blank" rel="noopener">Banco de España</a>`;
  $("debtLead").innerHTML = es
    ? `Desde ${dm.pre.year + 1}, el Estado presta cada año dinero a la Seguridad Social para poder pagar las pensiones. Esa deuda ha pasado de <strong>${fmtBig(dm.pre.debt_eur)}</strong> (fin de ${dm.pre.year}) a <strong>${fmtBig(last.debt_eur)}</strong> (${qLabel(last)}): <strong>×${fmtNum(mult, 1)}</strong>.`
    : `Since ${dm.pre.year + 1}, the State has lent money to Social Security every year so that pensions can be paid. That debt went from <strong>${fmtBig(dm.pre.debt_eur)}</strong> (end of ${dm.pre.year}) to <strong>${fmtBig(last.debt_eur)}</strong> (${qLabel(last)}): <strong>×${fmtNum(mult, 1)}</strong>.`;
  const tiles = [
    [fmtBig(last.debt_eur), es ? `deuda oficial, ${qLabel(last)}` : `official debt, ${qLabel(last)}`, bdeLink],
    [`${fmtNum(last.debt_pct_gdp, 1)} %`, es ? "del PIB de España" : "of Spain's GDP", bdeLink],
    [`+${fmtBig(dm.ratePerYear)}`, es ? `deuda nueva por año (media ${dm.decBase.year + 1}–${dm.lastDec.year})` : `new debt per year (average ${dm.decBase.year + 1}–${dm.lastDec.year})`, bdeLink],
    [fmtPct(pensionShare, 1), es ? `de lo que se paga en pensiones contributivas en un año (${fmtBig(dm.ratePerYear)} ÷ gasto de ${state.annual.filter((a) => a.kind === "actual").pop().year})` : `of a year of contributory pension payments (${fmtBig(dm.ratePerYear)} ÷ spending in ${state.annual.filter((a) => a.kind === "actual").pop().year})`, "INSS / Banco de España"],
  ];
  if (it.affiliates) tiles.push([fmtEur(last.debt_eur / it.affiliates.value, 0), es ? `por cada afiliado a la Seguridad Social (${fmtNum(it.affiliates.value)})` : `per Social Security contributor (${fmtNum(it.affiliates.value)})`, `<a href="${it.affiliates.source_url}" target="_blank" rel="noopener">${it.affiliates.source_name}</a>`]);
  $("debtTiles").innerHTML = tiles.map(([v, l, s]) => `<div class="tile"><div class="tile-val">${v}</div><div class="tile-lbl">${l}</div><div class="tile-src">${s}</div></div>`).join("");

  // annual increase (Dec -> Dec) + current year to date
  const bars = [];
  for (let i = 1; i < dm.decs.length; i++) {
    const a = dm.decs[i - 1], b = dm.decs[i], d = b.debt_eur - a.debt_eur;
    bars.push({ x: String(b.year), v: d, hl: false, tip: `<strong>${b.year}</strong><br>${d >= 0 ? "+" : ""}${fmtBig(d)}<br><span class="src">${es ? "deuda a fin de año" : "year-end debt"}: ${fmtBig(b.debt_eur)}</span>` });
  }
  if (last.quarter_end_month !== 12) {
    const d = last.debt_eur - dm.lastDec.debt_eur;
    bars.push({ x: `${last.year}*`, v: d, hl: false, tip: `<strong>${last.year}</strong> (${es ? "hasta" : "to"} ${qLabel(last)})<br>${d >= 0 ? "+" : ""}${fmtBig(d)}<br><span class="src">${es ? "año incompleto" : "incomplete year"}</span>` });
  }
  const bn = (v) => `${fmtNum(v / 1e9, 0)}`;
  miniChart($("debtBars"), bars, "bar", bn);
  const line = dm.decs.map((r) => ({ x: String(r.year), v: r.debt_eur, tip: `<strong>${es ? "Fin de" : "End of"} ${r.year}</strong><br>${fmtBig(r.debt_eur)}${r.debt_pct_gdp ? ` · ${fmtNum(r.debt_pct_gdp, 1)} % ${es ? "del PIB" : "of GDP"}` : ""}` }));
  if (last.quarter_end_month !== 12) line.push({ x: qLabel(last).replace(/^(\w{3})\w*/, "$1"), v: last.debt_eur, tip: `<strong>${qLabel(last)}</strong><br>${fmtBig(last.debt_eur)} · ${fmtNum(last.debt_pct_gdp, 1)} %` });
  miniChart($("debtLine"), line, "line", bn);
  $("debtBars").setAttribute("aria-label", es ? "Gráfico de barras: deuda nueva de la Seguridad Social por año" : "Bar chart: new Social Security debt per year");
  $("debtLine").setAttribute("aria-label", es ? "Gráfico de línea: deuda acumulada de la Seguridad Social" : "Line chart: accumulated Social Security debt");

  const loan = it.loan_2026_approval, cap = it.loan_cap_annual, def = it.ss_deficit_2025, tr = it.state_transfers_2025;
  const a = (x) => x ? ` <span class="src">(<a href="${x.source_url}" target="_blank" rel="noopener">${x.source_name}</a>)</span>` : "";
  const items = es ? [
    `<strong>¿Quién presta?</strong> Prácticamente toda: ${fmtBig(last.held_by_state_eur)} de ${fmtBig(last.debt_eur)} (${fmtPct(shareState, 3)}) están en manos del propio Estado (préstamos a la Tesorería General de la Seguridad Social). El Estado, a su vez, se financia emitiendo deuda pública. <span class="src">(${bdeLink}, serie «en poder del Estado»)</span>`,
    cap ? `<strong>¿Cuánto cada año?</strong> La ley de presupuestos (prorrogada desde 2023) permite prestar hasta ${fmtBig(cap.value)} al año a la Seguridad Social.${a(cap)}` : "",
    loan ? `<strong>2026:</strong> el ${loan.date.split("-").reverse().join("/")} el Consejo de Ministros autorizó un nuevo préstamo; aún no aparece en el último dato trimestral (${qLabel(last)}).${a(loan)}` : "",
    def && tr ? `<strong>Además de los préstamos</strong>, en 2025 el Estado transfirió ${fmtBig(tr.value)} a la Seguridad Social (no son deuda de la Seguridad Social, se pagan con los presupuestos del Estado), y aun así el sistema cerró con un déficit de ${fmtBig(def.value)}.${a(def)}` : "",
    `<strong>Para no confundir:</strong> como acreedor y deudor son administraciones públicas, en la deuda pública total de España (consolidada) esta deuda se compensa y no se suma dos veces. Lo que mide es cuánto ha tenido que adelantar el Estado a la Seguridad Social.`,
  ] : [
    `<strong>Who lends?</strong> Almost all of it: ${fmtBig(last.held_by_state_eur)} of ${fmtBig(last.debt_eur)} (${fmtPct(shareState, 3)}) is held by the State itself (loans to the Social Security Treasury). The State in turn finances itself by issuing public debt. <span class="src">(${bdeLink}, series "held by the State")</span>`,
    cap ? `<strong>How much each year?</strong> The budget law (rolled over since 2023) allows lending up to ${fmtBig(cap.value)} a year to Social Security.${a(cap)}` : "",
    loan ? `<strong>2026:</strong> on ${loan.date} the Council of Ministers authorised a new loan; it is not yet in the latest quarterly figure (${qLabel(last)}).${a(loan)}` : "",
    def && tr ? `<strong>On top of the loans</strong>, in 2025 the State transferred ${fmtBig(tr.value)} to Social Security (not Social Security debt; paid from the State budget), and the system still closed with a ${fmtBig(def.value)} deficit.${a(def)}` : "",
    `<strong>To avoid confusion:</strong> because lender and borrower are both public administrations, this debt is netted out in Spain's total (consolidated) public debt and not counted twice. It measures how much the State has had to advance to Social Security.`,
  ];
  $("debtExplain").innerHTML = `<ul>${items.filter(Boolean).map((x) => `<li>${x}</li>`).join("")}</ul>`;

  let rows = `<tr><th>${es ? "Fecha" : "Date"}</th><th>${es ? "Deuda" : "Debt"}</th><th>${es ? "% PIB" : "% GDP"}</th><th>${es ? "Variación anual" : "Annual change"}</th><th>${es ? "En poder del Estado" : "Held by the State"}</th></tr>`;
  const tbl = [...dm.decs]; if (last.quarter_end_month !== 12) tbl.push(last);
  tbl.forEach((r, i) => {
    const d = i ? r.debt_eur - tbl[i - 1].debt_eur : null;
    rows += `<tr><td>${r.quarter_end_month === 12 ? (es ? "dic. " : "Dec ") + r.year : qLabel(r)}</td><td>${fmtBig(r.debt_eur)}</td><td>${r.debt_pct_gdp ? fmtNum(r.debt_pct_gdp, 1) : "—"}</td><td>${d === null ? "—" : (d >= 0 ? "+" : "") + fmtBig(d)}</td><td>${r.held_by_state_eur ? fmtBig(r.held_by_state_eur) : "—"}</td></tr>`;
  });
  $("debtTable").innerHTML = rows;
}

/* --------------------------------------------------------------- reform simulator */
/* Each lever: value = slider setting; eur() = first-order annual effect in €.
   All inputs come from context.json items (shown with their source). */
function levers() {
  const it = state.ctx.items;
  const lastFull = state.annual.filter((a) => a.kind === "actual").pop();
  const spend = lastFull.spend;
  const L = [];
  const src = (key) => it[key] ? `<a href="${it[key].source_url}" target="_blank" rel="noopener">${it[key].source_name}</a> (${it[key].date})` : "";
  const es = state.lang === "es";

  if (it.contributions_revenue && it.contribution_rate_total) {
    const R = it.contributions_revenue.value, rate = it.contribution_rate_total.value;
    L.push({ id: "rate", min: 0, max: 3, step: 0.1, def: 0,
      title: es ? "Subir las cotizaciones sociales" : "Raise social security contributions",
      setting: (v) => es ? `+${fmtNum(v, 1)} puntos sobre el tipo actual (${fmtNum(rate, 2)} %)` : `+${fmtNum(v, 1)} points on the current rate (${fmtNum(rate, 2)}%)`,
      eur: (v) => (R * v) / rate,
      formula: (v) => `${fmtBig(R)} × ${fmtNum(v, 1)} / ${fmtNum(rate, 2)} — ${src("contributions_revenue")}; ${src("contribution_rate_total")}`,
      tradeoff: es ? "Empresas y trabajadores: encarece el empleo y reduce el salario neto." : "Employers and workers: makes hiring costlier and cuts take-home pay." });
  }
  if (it.new_retirements_year && it.avg_initial_retirement_pension) {
    const N = it.new_retirements_year.value, P = it.avg_initial_retirement_pension.value;
    L.push({ id: "age", min: 0, max: 24, step: 1, def: 0,
      title: es ? "Retrasar la edad efectiva de jubilación" : "Delay the effective retirement age",
      setting: (v) => es ? `${fmtNum(v)} meses más tarde` : `${fmtNum(v)} months later`,
      eur: (v) => N * P * 14 * (v / 12),
      formula: (v) => `${fmtNum(N)} ${es ? "nuevas jubilaciones/año" : "new retirements/year"} × ${fmtEur(P, 2)} × 14 × ${fmtNum(v)}/12 — ${src("new_retirements_year")}. ${es ? "Ahorro aproximado de los primeros años; ignora las cotizaciones adicionales y los complementos por demora." : "Approximate saving in the first years; ignores extra contributions and delayed-retirement bonuses."}`,
      tradeoff: es ? "Futuros jubilados: trabajan más años antes de cobrar." : "Future retirees: they work longer before drawing a pension." });
  }
  L.push({ id: "reval", min: 0, max: 2, step: 0.1, def: 0,
    title: es ? "Revalorizar las pensiones por debajo del IPC" : "Index pensions below inflation",
    setting: (v) => es ? `IPC − ${fmtNum(v, 1)} puntos durante un año` : `CPI − ${fmtNum(v, 1)} points for one year`,
    eur: (v) => spend * v / 100,
    formula: (v) => `${fmtBig(spend)} (${es ? "gasto" : "spending"} ${lastFull.year}, INSS) × ${fmtNum(v, 1)} %. ${es ? "El ahorro se acumula cada año que se mantiene." : "The saving compounds every year it is kept."}`,
    tradeoff: es ? "Pensionistas actuales: pierden poder adquisitivo." : "Current pensioners: they lose purchasing power." });
  if (it.contributions_revenue && it.affiliates) {
    const R = it.contributions_revenue.value, A = it.affiliates.value;
    L.push({ id: "jobs", min: 0, max: 2000, step: 50, def: 0,
      title: es ? "Más empleo cotizando (empleo, inmigración regular, aflorar economía sumergida)" : "More people paying contributions (jobs, regular migration, formalising undeclared work)",
      setting: (v) => es ? `+${fmtNum(v)} mil afiliados` : `+${fmtNum(v)} thousand contributors`,
      eur: (v) => (v * 1000 * R) / A,
      formula: (v) => `${fmtNum(v * 1000)} × ${fmtBig(R)} / ${fmtNum(A)} ${es ? "afiliados" : "contributors"} — ${src("contributions_revenue")}; ${src("affiliates")}. ${es ? "A largo plazo también generan derechos de pensión futuros." : "In the long run they also earn future pension rights."}`,
      tradeoff: es ? "Requiere crecimiento económico o políticas de empleo/migración; no tiene coste directo para nadie." : "Requires economic growth or employment/migration policy; no direct cost to anyone." });
  }
  if (it.gdp_2025) {
    const G = it.gdp_2025.value;
    L.push({ id: "tax", min: 0, max: 1, step: 0.05, def: 0,
      title: es ? "Financiar con impuestos generales (más transferencias del Estado)" : "Fund it from general taxation (bigger State transfers)",
      setting: (v) => es ? `${fmtNum(v, 2)} % del PIB en impuestos adicionales` : `${fmtNum(v, 2)}% of GDP in extra taxes`,
      eur: (v) => G * v / 100,
      formula: (v) => `${fmtBig(G)} × ${fmtNum(v, 2)} % — ${src("gdp_2025")}. ${es ? "No elimina el gasto: lo traslada del sistema de cotizaciones al conjunto de contribuyentes." : "It does not remove spending: it shifts it from contributors to all taxpayers."}`,
      tradeoff: es ? "Todos los contribuyentes (IRPF, IVA, sociedades…)." : "All taxpayers (income tax, VAT, corporate tax…)." });
  }
  return L;
}

function reformTargets() {
  const it = state.ctx.items, es = state.lang === "es", out = [];
  if (it.ss_deficit_2025) out.push({ id: "deficit", value: it.ss_deficit_2025.value,
    label: es ? "Déficit de la Seguridad Social en 2025 (se cubre con préstamos del Estado → deuda)" : "Social Security deficit 2025 (covered with State loans → debt)",
    detail: () => `${fmtBig(it.ss_deficit_2025.value)} — <a href="${it.ss_deficit_2025.source_url}" target="_blank" rel="noopener">${it.ss_deficit_2025.source_name}</a>` });
  if (it.ss_deficit_2025 && it.state_transfers_2025) out.push({ id: "state", value: it.ss_deficit_2025.value + it.state_transfers_2025.value,
    label: es ? "Déficit + transferencias del Estado en 2025 (todo lo que no cubren las cotizaciones)" : "Deficit + State transfers 2025 (everything contributions do not cover)",
    detail: () => `${fmtBig(it.ss_deficit_2025.value)} + ${fmtBig(it.state_transfers_2025.value)} — <a href="${it.state_transfers_2025.source_url}" target="_blank" rel="noopener">${it.state_transfers_2025.source_name}</a>. ${it.state_transfers_2025["note_" + state.lang]}` });
  return out;
}

function renderReform() {
  const L = t(), it = state.ctx.items, es = state.lang === "es";
  const targets = reformTargets();
  if (!targets.length) { $("reform").hidden = true; return; }
  const prevTarget = state.target || "deficit";
  const target = targets.find((x) => x.id === prevTarget) || targets[0];
  state.target = target.id;
  const gap = target.value;
  const radios = targets.map((x) => `<label style="display:block;margin:4px 0"><input type="radio" name="target" value="${x.id}" ${x.id === target.id ? "checked" : ""}> ${x.label}</label>`).join("");
  $("gapIntro").innerHTML = `${L.gapIntro(fmtBig(gap), target.label)}<br><span class="src">${target.detail()}</span><div style="margin-top:10px">${radios}</div>`;
  $("gapIntro").querySelectorAll("input[name=target]").forEach((r) => r.addEventListener("change", (e) => { state.target = e.target.value; renderReform(); }));
  const box = $("levers");
  const prev = {};
  box.querySelectorAll("input[type=range]").forEach((i) => (prev[i.dataset.id] = Number(i.value)));
  box.innerHTML = "";
  const list = levers();
  const update = () => {
    let total = 0;
    for (const lv of list) {
      const v = Number(box.querySelector(`input[data-id=${lv.id}]`).value);
      const eur = lv.eur(v);
      total += eur;
      box.querySelector(`[data-set=${lv.id}]`).textContent = lv.setting(v);
      box.querySelector(`[data-val=${lv.id}]`).textContent = v ? "+" + fmtBig(eur) : "—";
      box.querySelector(`[data-f=${lv.id}]`).innerHTML = lv.formula(v || lv.step);
    }
    const share = Math.min(total / gap, 1);
    $("gapFill").style.width = share * 100 + "%";
    $("gapLegend").innerHTML = `<span>${L.closed(fmtPct(total / gap, 0), fmtBig(total))}</span><span>${L.remaining(fmtBig(Math.max(gap - total, 0)))}</span>`;
    const cur = state.annual[state.annual.length - 1];
    const daily = cur.spend / daysIn(cur.year);
    let extra = "";
    const fund = it.reserve_fund_current, debt = it.ss_debt;
    if (fund) extra += `<li>${es ? "Sin reforma, el Fondo de Reserva" : "Without reform, the Reserve Fund"} (${fmtBig(fund.value)}, ${fund.date}) ${es ? "cubriría este agujero" : "would cover this gap for"} <strong>${fmtNum(fund.value / gap, 1)} ${es ? "años" : "years"}</strong> ${es ? "— o pagaría las pensiones durante" : "— or pay pensions for"} <strong>${fmtDuration(fund.value / daily)}</strong>.</li>`;
    if (debt && total > gap) extra += `<li>${es ? "Con el superávit de tu paquete" : "With your package's surplus"} (${fmtBig(total - gap)}/${es ? "año" : "year"}) ${es ? "se devolvería la deuda de la Seguridad Social" : "the Social Security debt"} (${fmtBig(debt.value)}, ${debt.date}) ${es ? "en" : "would be repaid in"} <strong>${fmtNum(debt.value / (total - gap), 1)} ${es ? "años" : "years"}</strong> (${es ? "sin intereses" : "ignoring interest"}).</li>`;
    else if (debt) extra += `<li>${es ? "Mientras quede agujero, la deuda de la Seguridad Social" : "While a gap remains, the Social Security debt"} (${fmtBig(debt.value)}, ${debt.date}) ${es ? "seguiría creciendo." : "keeps growing."}</li>`;
    $("reformSummary").innerHTML = (total > 0 ? L.summary(fmtBig(total), fmtPct(total / gap, 0), fmtDuration(total / daily)) : L.summaryNone) + (extra ? `<ul>${extra}</ul>` : "");
  };
  for (const lv of list) {
    const d = document.createElement("div");
    d.className = "lever";
    d.innerHTML = `<div class="lever-head"><span class="lever-title">${lv.title}</span><span class="lever-value" data-val="${lv.id}"></span></div>
      <input type="range" min="${lv.min}" max="${lv.max}" step="${lv.step}" value="${prev[lv.id] ?? lv.def}" data-id="${lv.id}" aria-label="${lv.title}">
      <div class="setting" data-set="${lv.id}"></div>
      <div class="tradeoff"><strong>${L.tradeoff}:</strong> ${lv.tradeoff}</div>
      <details><summary>${L.formula}</summary><div data-f="${lv.id}"></div></details>`;
    box.appendChild(d);
  }
  box.querySelectorAll("input[type=range]").forEach((i) => i.addEventListener("input", update));
  update();
}

/* --------------------------------------------------------------- sources */
function renderSources() {
  const L = t(), s = state.series, it = state.ctx.items;
  const first = s.monthly[0], last = s.monthly[s.monthly.length - 1];
  let html = `<p>${L.sourcesIntro} <a href="data/pensions.json" download>${L.download}</a> · <a href="data/context.json" download>context.json</a></p>
    <ul class="sources-list">
      <li>${L.srcSeries(s.monthly.length, monthName(first.year, first.month), monthName(last.year, last.month))}
        <br><a href="${s.archive_url}" target="_blank" rel="noopener">INSS — Histórico de estadísticas de pensiones</a> ·
        <a href="${s.current_url}" target="_blank" rel="noopener">INSS — Pensiones contributivas en vigor</a>
        <br><span class="src">${L.validated}</span></li>
      <li>${state.lang === "es" ? "Deuda de las Administraciones de Seguridad Social (definición PDE), trimestral desde 1994" : "Debt of the Social Security Administrations (EDP definition), quarterly since 1994"}:
        <a href="${state.debt.source_page}" target="_blank" rel="noopener">Banco de España — Boletín Estadístico, tabla 12.16</a> ·
        <a href="data/debt.json" download>debt.json</a>
        <br><span class="src">${state.lang === "es" ? "Reloj de deuda: último dato oficial + ritmo medio de crecimiento de los 3 últimos años completos, repartido de forma continua (estimación)." : "Debt clock: latest official figure + average growth of the last 3 complete years, spread evenly over time (estimate)."}</span></li>
    </ul>${L.method}<h3>${L.ctxTitle}</h3><ul class="sources-list">`;
  for (const [key, v] of Object.entries(it)) {
    const value = v.unit === "event" ? "" : ": " + (v.unit === "EUR" ? fmtBig(v.value) : v.unit === "pct" ? `${fmtNum(v.value, 2)} %` : fmtNum(v.value, v.value % 1 ? 4 : 0));
    html += `<li><strong>${v["label_" + state.lang]}</strong>${value} <span class="muted">(${v.date})</span><br>
      <a href="${v.source_url}" target="_blank" rel="noopener">${v.source_name}</a>
      ${v.quote ? `<blockquote>«${v.quote}»</blockquote>` : ""}${v["note_" + state.lang] ? `<span class="src">${v["note_" + state.lang]}</span>` : ""}</li>`;
  }
  html += "</ul>";
  $("sourcesBody").innerHTML = html;
  const d = new Date(s.generated_at);
  $("footer").innerHTML = `<p>${L.updated(d.toLocaleDateString(loc(), { year: "numeric", month: "long", day: "numeric" }))}</p><p>${L.footer}</p>`;
}

/* --------------------------------------------------------------- wiring */
function applyStatic() {
  const L = t();
  document.documentElement.lang = state.lang;
  document.title = L.pageTitle;
  document.querySelectorAll("[data-i18n]").forEach((el) => { if (L[el.dataset.i18n]) el.textContent = L[el.dataset.i18n]; });
  document.querySelectorAll(".lang-btn").forEach((b) => b.classList.toggle("active", b.dataset.lang === state.lang));
}

function refresh() {
  renderYearSelect();
  renderChart();
  calculate();
}

function setLang(lang) {
  state.lang = lang;
  try { localStorage.setItem("lang", lang); } catch (e) { /* storage unavailable */ }
  const raw = parseAmount($("moneyInput").value);
  applyStatic();
  if (raw > 0) $("moneyInput").value = fmtNum(raw, raw % 1 ? 2 : 0);
  renderPresets(); renderReform(); renderSources(); refresh();
  renderClockText(); renderDebt();
}

async function init() {
  try { const l = localStorage.getItem("lang"); if (l === "en" || l === "es") state.lang = l; } catch (e) { /* ignore */ }
  const q = new URLSearchParams(location.search).get("lang");  // shareable ?lang=en
  if (q === "en" || q === "es") state.lang = q;
  const [series, ctx, debt] = await Promise.all([
    fetch("data/pensions.json").then((r) => r.json()),
    fetch("data/context.json").then((r) => r.json()),
    fetch("data/debt.json").then((r) => r.json()),
  ]);
  state.series = series; state.ctx = ctx; state.debt = debt;
  buildAnnual();
  state.dm = debtModel();
  // expose the latest official debt like any other sourced figure (presets, simulator, sources list)
  const ld = state.dm.last;
  ctx.items.ss_debt = {
    value: ld.debt_eur, unit: "EUR", date: `${ld.year}-${String(ld.quarter_end_month).padStart(2, "0")}`,
    label_es: `Deuda de las Administraciones de Seguridad Social (PDE), ${ld.debt_pct_gdp} % del PIB`,
    label_en: `Debt of the Social Security Administrations (EDP), ${ld.debt_pct_gdp}% of GDP`,
    short_es: "Deuda de la Seguridad Social", short_en: "Social Security debt",
    source_name: "Banco de España — Boletín Estadístico, tabla 12.16 (descarga automática)", source_url: debt.source_page,
  };
  const lastY = state.annual[state.annual.length - 1].year;
  const until = $("projUntil");
  for (let y = lastY + 1; y <= 2050; y++) until.add(new Option(y, y));
  until.value = Math.min(lastY + 9, 2050);
  $("moneyInput").value = fmtNum(1000000);
  document.querySelectorAll(".lang-btn").forEach((b) => b.addEventListener("click", () => setLang(b.dataset.lang)));
  $("moneyInput").addEventListener("input", calculate);
  $("yearSelect").addEventListener("change", (e) => { state.selYear = Number(e.target.value); refresh(); });
  $("projWindow").addEventListener("change", refresh);
  $("projUntil").addEventListener("change", refresh);
  let rz;
  window.addEventListener("resize", () => { clearTimeout(rz); rz = setTimeout(() => { renderChart(); renderDebt(); }, 150); });
  setLang(state.lang);
  startClock();
}
init().catch((e) => {
  document.getElementById("result").innerHTML = `<p class="error">Error loading data: ${e.message}. If you opened the file directly, serve the folder over HTTP (e.g. <code>python -m http.server</code>).</p>`;
});
