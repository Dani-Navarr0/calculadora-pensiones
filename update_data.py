"""
Builds data/pensions.json from PRIMARY official sources only.

Source: Seguridad Social (INSS) - "Pensiones contributivas en vigor.
Número y cuantía por sexo, régimen y clase" monthly Excel files.

  * Historical files: "Histórico de estadísticas de pensiones publicadas"
    https://www.seg-social.es/wps/portal/wss/internet/EstadisticasPresupuestosEstudios/Estadisticas/EST23/2575
  * Latest month:     "Pensiones contributivas en vigor"
    https://www.seg-social.es/wps/portal/wss/internet/EstadisticasPresupuestosEstudios/Estadisticas/EST23/EST24

For every month the script stores: number of pensions, the ordinary monthly
payroll ("nómina ordinaria", euros) and the exact URL of the Excel file the
number was read from, so anyone can re-check every figure.

Run:  pip install openpyxl xlrd
      python update_data.py              # full rebuild (about 230 downloads)
      python update_data.py --latest     # only add missing months (checks this & last year)
      python update_data.py --cache DIR  # keep downloaded Excel files in DIR (debugging)

If anything cannot be parsed or fails the sanity checks, the script exits with
an error and does NOT overwrite data/pensions.json.
"""
import html
import io
import json
import re
import sys
import time
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
import xlrd

BASE = "https://www.seg-social.es"
ARCHIVE = BASE + "/wps/portal/wss/internet/EstadisticasPresupuestosEstudios/Estadisticas/EST23/2575"
CURRENT = BASE + "/wps/portal/wss/internet/EstadisticasPresupuestosEstudios/Estadisticas/EST23/EST24"
OUT = Path(__file__).parent / "data" / "pensions.json"
UA = {"User-Agent": "Mozilla/5.0 (compatible; pension-calculator-data-builder)"}

MONTHS = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7,
          "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}


# ---------------------------------------------------------------- download
CACHE = None  # set with --cache DIR to keep downloaded files (debugging)


def fetch(url, binary=False, tries=4):
    if binary and CACHE:
        f = CACHE / re.sub(r"[^A-Za-z0-9.]+", "_", url.split("/wss/")[-1].split("?")[0])
        if f.exists():
            return f.read_bytes()
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers=UA)
            with urllib.request.urlopen(req, timeout=90) as r:
                data = r.read()
            if binary and CACHE:
                f.write_bytes(data)
            return data if binary else data.decode("utf-8", "ignore")
        except Exception as e:  # network hiccup -> retry
            if i == tries - 1:
                raise
            print(f"   retry {url[:80]}... ({e})", file=sys.stderr)
            time.sleep(3 * (i + 1))


def links(page):
    for m in re.finditer(r'href="([^"]+)"[^>]*>(.*?)</a>', page, flags=re.S):
        url = html.unescape(m.group(1))
        text = html.unescape(re.sub(r"<[^>]+>|\s+", " ", m.group(2))).strip()
        yield text, url if url.startswith("http") else BASE + url


def clean_file_url(url):
    """Keep the stable part of a WCM URL (drop CACHEID etc.)."""
    return url.split("?")[0] + "?MOD=AJPERES"


def regime_files(page_html):
    """Excel links on a 'Pensiones por sexo, regímenes y clase(s)' page."""
    out = []
    for _, u in links(page_html):
        if "/wps/wcm/connect/" in u and re.search(r"\.xlsx?(\?|$)", u, re.I):
            out.append(clean_file_url(u))
    return out


def archive_files(only_years=None):
    idx = fetch(ARCHIVE)
    files = []
    for text, url in links(idx):
        if not re.fullmatch(r"Año \d{4}", text):
            continue
        if only_years and int(text[-4:]) not in only_years:
            continue
        year_page = fetch(url)
        vigor = next((u for t, u in links(year_page)
                      if "contributivas en vigor" in t.lower() and "/2575/" in u and "!ut" not in u), None)
        if not vigor:
            print(f"!! {text}: no 'Pensiones contributivas en vigor' page", file=sys.stderr)
            continue
        vigor_page = fetch(vigor)
        reg = next((u for t, u in links(vigor_page)
                    if "regímenes" in t.lower() and "/2575/" in u and "!ut" not in u), None)
        if not reg:
            print(f"!! {text}: no 'regímenes y clases' page", file=sys.stderr)
            continue
        found = regime_files(fetch(reg))
        print(f"{text}: {len(found)} files")
        files += found
    return files


def current_files():
    page = fetch(CURRENT)
    # On the current page the regime file is the one called REGyyyymm.xlsx
    return [clean_file_url(u) for _, u in links(page)
            if re.search(r"/REG\d{6}\.xlsx?", u, re.I)]


# ---------------------------------------------------------------- parsing
def rows_of(content, url):
    """Return list of sheets, each a list of rows (lists of cell values)."""
    if content[:2] == b"PK":
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        return [[list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets]
    if content[:4] == b"\xd0\xcf\x11\xe0":
        wb = xlrd.open_workbook(file_contents=content)
        return [[sh.row_values(i) for i in range(sh.nrows)] for sh in wb.sheets()]
    raise ValueError(f"Not an Excel file: {url} -> {content[:80]!r}")


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    return None


def reference_date(sheets):
    """Date written inside the workbook, e.g. 'Datos a 1 de diciembre de 2015'."""
    for sheet in sheets:
        for row in sheet[:12]:
            for c in row:
                if isinstance(c, str):
                    m = re.search(r"1\s+de\s+([a-záéíóú]+)\s+(?:de\s+)?(\d{4})", c.lower())
                    if m and m.group(1) in MONTHS:
                        return int(m.group(2)), MONTHS[m.group(1)]
    return None


def filename_date(url):
    """REG202610.xlsx / REG3201512.xlsx -> YYYYMM ; REG3102010.xlsx -> MMYYYY."""
    m = re.search(r"(\d{6})\.xlsx?", url, re.I)
    if not m:
        return None
    d = m.group(1)
    if d[:2] in ("19", "20") and 1 <= int(d[4:]) <= 12:
        return int(d[:4]), int(d[4:])
    if 1 <= int(d[:2]) <= 12 and d[2:4] in ("19", "20"):
        return int(d[2:]), int(d[:2])
    return None


def norm(c):
    return re.sub(r"\s+", " ", c).strip().lower() if isinstance(c, str) else ""


CLASS_WORDS = ("incapacidad", "jubilaci", "viudedad", "orfandad", "favor")
SEX_TOTAL = ("total", "ambos sexos")
SEX_PART = ("hombres", "mujeres", "no consta")


def parse_block(sheet, t):
    """
    Read the 'Total sistema' row at index t using the header rows above it.
    Header structure (top -> bottom, all optional except measure row):
       class row : 'Total' / 'Total pensiones' | 'Incapacidad permanente' | ...
       sex row   : 'Total (*)' / 'Ambos sexos (*)' | 'Hombres' | 'Mujeres' | 'No consta'
       measure   : 'Núm.' / 'Número' | 'P. media' / 'Importe'
    Returns (pensions, euros, method) or None if this row is not the right table.
    """
    # header rows live between the previous table (if any) and this row;
    # 'Total sistema' can be the first OR the last row of a table.
    start = max(0, t - 16)
    for k in range(t - 1, start - 1, -1):
        if any(norm(c).startswith("total sistema") for c in sheet[k][:2]):
            start = k + 1
            break
    rows = sheet[start:t]
    measure_i = next((i for i in range(len(rows) - 1, -1, -1)
                      if any(norm(c).startswith(("núm", "número")) for c in rows[i])), None)
    if measure_i is None:
        return None
    measure = rows[measure_i]
    # optional sex row directly above the measure row
    has_sex = measure_i > 0 and any(norm(c).startswith(SEX_PART) for c in rows[measure_i - 1])
    class_i = measure_i - 2 if has_sex else measure_i - 1
    if class_i < 0 or not any(
            j > 0 and (any(w in norm(c) for w in CLASS_WORDS) or norm(c).startswith("total"))
            for j, c in enumerate(rows[class_i])):
        return None
    class_row = rows[class_i]
    starts = sorted(j for j, c in enumerate(class_row) if norm(c) and j > 0)
    tot = [j for j in starts if re.fullmatch(r"total( pensiones)?( \(\*+\))?", norm(class_row[j]))]
    if len(tot) != 1:
        return None
    c0 = tot[0]
    nxt = [j for j in starts if j > c0]
    c1 = nxt[0] if nxt else max(len(r) for r in rows + [sheet[t]])
    cols = range(c0, c1)
    # optional sex row between class row and measure row
    sex_row = rows[class_i + 1] if class_i + 1 < measure_i else None
    groups = [(c0, c1)]
    if sex_row is not None:
        labels = [(j, norm(sex_row[j])) for j in cols if norm(sex_row[j])]
        tot_lbl = [j for j, l in labels if l.startswith(SEX_TOTAL)]
        if tot_lbl:
            j0 = tot_lbl[0]
            after = [j for j, _ in labels if j > j0]
            groups = [(j0, after[0] if after else c1)]
        elif labels and all(l.startswith(SEX_PART) for _, l in labels):
            js = [j for j, _ in labels] + [c1]
            groups = list(zip(js[:-1], js[1:]))
        else:
            return None
    row = sheet[t]
    total_n = total_eur = 0.0
    method = None
    for g0, g1 in groups:
        mcols = {norm(measure[j]): j for j in range(g0, min(g1, len(measure))) if norm(measure[j])}
        jn = next((j for k, j in mcols.items() if k.startswith(("núm", "número"))), None)
        ji = next((j for k, j in mcols.items() if k.startswith("importe")), None)
        jm = next((j for k, j in mcols.items() if k.startswith(("p.media", "p. media"))), None)
        if jn is None or (ji is None and jm is None):
            return None
        n = num(row[jn]) if jn < len(row) else None
        if n is None:  # empty cell (e.g. no 'No consta' cases) -> 0
            continue
        if ji is not None:
            eur, method = num(row[ji]), "importe"
        else:
            eur, method = n * (num(row[jm]) or 0.0), "num_x_media"
        total_n += n
        total_eur += eur
    if total_n < 1_000_000:  # percentage tables or partial tables, not the system total
        return None
    return int(round(total_n)), total_eur, method


def parse(sheets):
    """Try every 'Total sistema' row of every sheet; all valid readings must agree."""
    found = []
    for sheet in sheets:
        for t, r in enumerate(sheet):
            if any(norm(c).startswith("total sistema") for c in r[:2]):
                res = parse_block(sheet, t)
                if res:
                    found.append(res)
    if not found:
        raise ValueError("no parsable 'Total sistema' table")
    n0, e0, _ = found[0]
    for n, e, _ in found[1:]:
        if n != n0 or abs(e - e0) / e0 > 0.0005:
            raise ValueError(f"inconsistent tables in file: {found}")
    # prefer an exact 'importe' reading when available
    best = next((f for f in found if f[2] == "importe"), found[0])
    return best


# ---------------------------------------------------------------- validation
def validate(monthly):
    """Refuse to publish data that looks wrong. Bounds come from the 2008-2026 history:
    largest monthly payroll jump was Jan 2023 (+8.8 %, revaluation 8.5 %), largest fall
    May 2020 (-0.3 %, COVID-19)."""
    problems = []
    for prev, cur in zip(monthly, monthly[1:]):
        exp = (prev["year"] + (prev["month"] == 12), prev["month"] % 12 + 1)
        if (cur["year"], cur["month"]) != exp:
            problems.append(f"gap: {prev['year']}-{prev['month']} -> {cur['year']}-{cur['month']}")
            continue
        dp = cur["payroll_eur"] / prev["payroll_eur"] - 1
        dn = cur["pensions"] / prev["pensions"] - 1
        hi = 0.12 if cur["month"] == 1 else 0.04
        if not -0.02 <= dp <= hi:
            problems.append(f"{cur['year']}-{cur['month']:02d}: payroll change {dp:+.2%} out of range")
        if abs(dn) > 0.02:
            problems.append(f"{cur['year']}-{cur['month']:02d}: pension-count change {dn:+.2%} out of range")
    for r in monthly:
        if not 7e6 < r["pensions"] < 1.5e7:
            problems.append(f"{r['year']}-{r['month']:02d}: implausible pension count {r['pensions']}")
    return problems


# ---------------------------------------------------------------- Social Security debt (Banco de España)
DEBT_PAGE = "https://www.bde.es/webbe/es/estadisticas/otras-clasificaciones/publicaciones/boletin-estadistico/capitulo-12.html"
DEBT_CSV = "https://www.bde.es/webbe/es/estadisticas/compartido/datos/csv/be1216.csv"
DEBT_OUT = Path(__file__).parent / "data" / "debt.json"
Q_MONTH = {"MAR": 3, "JUN": 6, "SEP": 9, "DIC": 12}
DEBT_COLS = {  # series code in be1216.csv -> field name
    "DTNPDE2010_P00000_PS_ASG": "debt_eur",        # Deuda PDE. Total (miles de euros)
    "DTNPDE2010_P0000P_PS_ASG": "debt_pct_gdp",    # Total en porcentaje del PIBpm
    "DTNPDE2010_P00031_PS_ASG": "held_by_state_eur",  # En poder del Estado
}


def build_debt():
    """Quarterly PDE debt of the Social Security Administrations, Banco de España
    Boletín Estadístico table 12.16. Returns list of errors (empty = written OK)."""
    import csv
    try:
        raw = fetch(DEBT_CSV, binary=True).decode("latin-1")
    except Exception as e:
        return [f"debt: download failed: {e}"]
    rows = list(csv.reader(io.StringIO(raw)))
    head = rows[0]
    idx = {name: head.index(code) for code, name in DEBT_COLS.items() if code in head}
    if len(idx) != len(DEBT_COLS):
        return [f"debt: expected series not found in {DEBT_CSV} (columns changed?)"]
    units = rows[4]
    if "Miles de euros" not in units[idx["debt_eur"]]:
        return [f"debt: unexpected unit '{units[idx['debt_eur']]}'"]
    quarters = []
    for r in rows:
        m = re.fullmatch(r"(MAR|JUN|SEP|DIC) (\d{4})", r[0].strip()) if r else None
        if not m:
            continue
        def val(name, scale):
            v = r[idx[name]].strip()
            return None if v in ("", "_") else round(float(v) * scale, 2)
        rec = {"year": int(m.group(2)), "quarter_end_month": Q_MONTH[m.group(1)],
               "debt_eur": val("debt_eur", 1000), "debt_pct_gdp": val("debt_pct_gdp", 1),
               "held_by_state_eur": val("held_by_state_eur", 1000)}
        if rec["debt_eur"] is not None:
            quarters.append(rec)
    problems = []
    if len(quarters) < 100:
        problems.append(f"debt: only {len(quarters)} quarters parsed")
    for q in quarters:
        if not 1e9 < q["debt_eur"] < 1e12:
            problems.append(f"debt: implausible value {q}")
    if problems:
        return problems
    DEBT_OUT.write_text(json.dumps({
        "description": ("Debt (Excessive Deficit Procedure definition, ESA 2010) of the Spanish Social "
                        "Security Administrations, end of quarter, euros. 'held_by_state_eur' = part "
                        "owed to the central State (State loans)."),
        "publisher": "Banco de España - Boletín Estadístico, tabla 12.16",
        "source_page": DEBT_PAGE,
        "source_csv": DEBT_CSV,
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "quarterly": quarters,
    }, ensure_ascii=False, indent=1), encoding="utf-8")
    last = quarters[-1]
    print(f"Debt: wrote {len(quarters)} quarters; latest {last['year']}-{last['quarter_end_month']:02d} "
          f"= {last['debt_eur']/1e6:,.1f} M€ ({last['debt_pct_gdp']} % GDP)")
    return []


# ---------------------------------------------------------------- main
def main():
    global CACHE
    if "--cache" in sys.argv:
        CACHE = Path(sys.argv[sys.argv.index("--cache") + 1])
        CACHE.mkdir(parents=True, exist_ok=True)
    errors = build_debt()  # independent source: written even if pensions fail
    if "--debt-only" not in sys.argv:
        errors += build_pensions("--latest" in sys.argv)
    if errors:
        print("ERRORS (affected files NOT written):\n  " + "\n  ".join(errors), file=sys.stderr)
        sys.exit(1)


def build_pensions(latest_only):
    existing = {}
    if OUT.exists():
        for r in json.loads(OUT.read_text(encoding="utf-8"))["monthly"]:
            existing[(r["year"], r["month"])] = r

    this_year = datetime.now().year
    urls = (archive_files({this_year, this_year - 1}) if latest_only else archive_files()) + current_files()
    urls = list(dict.fromkeys(urls))  # de-duplicate, keep order
    records = dict(existing) if latest_only else {}
    errors = []
    for url in urls:
        fdate = filename_date(url)
        if latest_only and fdate in existing:
            continue
        sheets = rows_of(fetch(url, binary=True), url)
        ref = reference_date(sheets) or fdate
        if ref is None:
            errors.append(f"cannot determine month of {url}")
            continue
        if fdate and ref != fdate:
            errors.append(f"date mismatch: file name {fdate} vs content {ref} ({url})")
            continue
        try:
            n, imp, layout = parse(sheets)
        except ValueError as e:
            errors.append(f"{ref}: {e} ({url})")
            print(f"!! {ref}: {e}", file=sys.stderr)
            continue
        records[ref] = {"year": ref[0], "month": ref[1], "pensions": n,
                        "payroll_eur": round(imp, 2), "method": layout, "source": url}
        print(f"{ref[0]}-{ref[1]:02d}  pensions={n:>10,}  payroll={imp/1e6:>10,.2f} M€  [{layout}]")

    monthly = [records[k] for k in sorted(records)]
    errors += validate(monthly)
    if errors:  # never publish partial or suspicious data
        return errors
    data = {
        "description": ("Ordinary monthly payroll of CONTRIBUTORY pensions of the Spanish Social "
                        "Security system (pensiones contributivas en vigor a día 1 de cada mes). "
                        "Excludes the two extra payments (June & November), non-contributory "
                        "pensions and civil-servant pensions (Clases Pasivas)."),
        "publisher": "Instituto Nacional de la Seguridad Social (INSS) - Seguridad Social, Gobierno de España",
        "archive_url": ARCHIVE,
        "current_url": CURRENT,
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "monthly": monthly,
    }
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"Wrote {len(monthly)} months to {OUT}")
    return []


if __name__ == "__main__":
    main()
