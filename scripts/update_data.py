#!/usr/bin/env python3
"""
Shiller Explorer — automatic data updater.

Downloads Robert Shiller's ie_data.xls from shillerdata.com, regenerates the
three JSON data files used by the explorer, and refreshes the embedded
fallback data blocks inside Index.html and index.html.

Derivations (reverse-engineered from the existing JSONs and verified to
reproduce decomp10 exactly, 0 mismatches across all 1,621 dates):

  shiller_slim_v3_fixed.json
    d[]: per month —
      d   date (YYYY-MM-01)
      p   nominal S&P composite price          round(P, 2)
      rp  real price (Shiller's column)        round(RealPrice, 1)
      c   CAPE (Shiller's column)              round(CAPE, 2)
      cy  CAPE yield                           round(100/CAPE, 2)   [unrounded CAPE]
      dy  dividend yield                       round(100*D/P, 2)
      ey  earnings yield                       round(100*E/P, 2)
      gs  GS10 long rate                       round(GS10, 2)
      cpi CPI                                  round(CPI, 1)
      e   nominal earnings                     round(E, 2)
      div nominal dividend                     round(D, 2)
      fr  fwd 10y annualised real total return round(((TR[t+120]/TR[t])**0.1 - 1)*100, 2)
    r[]: static NBER recession list (extend manually if NBER declares a new one)
    m : metadata

  shiller_horizon_data.json
    one row per month with CAPE: {cy, d, f: [1..10yr annualised real TR, 2dp]}
    f[k-1] = ((TR[t+12k]/TR[t])**(1/k) - 1) * 100

  shiller_decomp_full.json
    series: [date, round(RealPrice,2), round(TR,2), round(CAPE,2) if CAPE]
            for months where RealPrice and TR both exist
    decomp10 (per month t where CAPE[t], CAPE[t+120], RealPrice, TR exist):
      tr = ann. 10y real total return          from TR
      cr = ann. 10y real price return          from RealPrice
      ir = ((1+tr)/(1+cr) - 1)                 income return
      mc = ann. 10y CAPE multiple change       from CAPE
      eg = ((1+cr)/(1+mc) - 1)                 smoothed real earnings growth
      fr = ((1+ir)*(1+eg) - 1)                 fundamental return
      (all *100, round 2dp)

Safety: the script validates the parsed spreadsheet against frozen history in
the previous JSON files before writing anything. If validation fails it exits
non-zero and writes nothing, so a layout change on shillerdata.com can never
silently corrupt the live data.
"""

import json
import math
import re
import sys
import urllib.request
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
SLIM_PATH = REPO_ROOT / "shiller_slim_v3_fixed.json"
HORIZON_PATH = REPO_ROOT / "shiller_horizon_data.json"
DECOMP_PATH = REPO_ROOT / "shiller_decomp_full.json"
HTML_CAPITAL = REPO_ROOT / "Index.html"
HTML_LOWER = REPO_ROOT / "index.html"

PAGE_URL = "https://shillerdata.com/"
# Documented fallback if scraping ever fails (the ?ver= changes over time):
# https://img1.wsimg.com/blobby/go/e5e77e0b-59d1-44d9-ab25-4763ac982e53/downloads/907c87f4-4176-4a13-9487-abddeadceb1b/ie_data.xls?ver=...

EMBED_CUTOFF = "2011-01-01"  # index.html embeds only data from this date on

USER_AGENT = "Mozilla/5.0 (ShillerExplorer data updater; github.com/BrickwoodAM/Shiller_Explorer)"


# ---------------------------------------------------------------- download --

def find_xls_url(html: str) -> str:
    """Locate the ie_data.xls download link on shillerdata.com.

    The page currently carries two xls download links (ie_data.xls and
    ie_data-0001.xls); we want the one whose filename is exactly ie_data.xls.
    """
    hrefs = re.findall(r'href="([^"]+\.xls[^"]*)"', html)
    hrefs += re.findall(r"href='([^']+\.xls[^']*)'", html)
    exact = [h for h in hrefs if re.search(r"/ie_data\.xls(\?|$)", h)]
    loose = [h for h in hrefs if "ie_data" in h]
    candidates = exact or loose
    if not candidates:
        raise RuntimeError("Could not find an ie_data.xls link on shillerdata.com")
    url = candidates[0]
    if url.startswith("//"):
        url = "https:" + url
    elif url.startswith("/"):
        url = "https://shillerdata.com" + url
    return url


def download_xls(dest: Path) -> None:
    req = urllib.request.Request(PAGE_URL, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=60) as resp:
        html = resp.read().decode("utf-8", errors="replace")
    url = find_xls_url(html)
    print(f"Downloading: {url}")
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=120) as resp:
        data = resp.read()
    if len(data) < 200_000:
        raise RuntimeError(f"Downloaded file suspiciously small ({len(data)} bytes)")
    dest.write_bytes(data)
    print(f"Saved {len(data):,} bytes to {dest}")


# ------------------------------------------------------------------ parse --

def decimal_date_to_iso(val: float) -> str:
    """Shiller dates are decimals: 1871.01 = Jan 1871, 1871.1 = OCT 1871."""
    year = int(val)
    month = int(round((val - year) * 100))
    if month == 1 and abs((val - year) * 100 - 1) > 0.5:
        raise ValueError(f"Unparseable decimal date {val}")
    if not (1 <= month <= 12):
        raise ValueError(f"Month out of range in decimal date {val}")
    return f"{year:04d}-{month:02d}-01"


def norm_label(s: str) -> str:
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def load_sheet(xls_path: Path):
    """Parse the 'Data' sheet. Returns list of row dicts keyed by canonical
    field names: date, p, d, e, cpi, gs10, real_price, tr, cape.
    Column detection is label-based (headers are stacked across several rows),
    then validated against frozen history before use.
    """
    import pandas as pd

    raw = pd.read_excel(xls_path, sheet_name="Data", header=None, engine="xlrd")

    # Locate the header row: the row whose first cell is the label 'Date'
    header_row = None
    for i in range(min(15, len(raw))):
        if norm_label(raw.iat[i, 0]) == "date":
            header_row = i
            break
    if header_row is None:
        raise RuntimeError("Could not locate header row (cell 'Date') in Data sheet")

    # Compound labels: join the header row with up to 4 rows above it
    ncols = raw.shape[1]
    labels = []
    for col in range(ncols):
        parts = []
        for r in range(max(0, header_row - 4), header_row + 1):
            cell = raw.iat[r, col]
            if isinstance(cell, str) and cell.strip():
                parts.append(cell.strip())
        labels.append(norm_label(" ".join(parts)))

    def find_col(predicate, what):
        hits = [i for i, lab in enumerate(labels) if predicate(lab)]
        if len(hits) != 1:
            raise RuntimeError(
                f"Column detection for '{what}' matched {len(hits)} columns "
                f"({[labels[i] for i in hits]}); labels were: {labels}"
            )
        return hits[0]

    def toks(l):
        return re.sub(r"[^a-z0-9/& ]", " ", l).split()

    cols = {
        "date": find_col(lambda l: l == "date", "date"),
        "p": find_col(lambda l: l == "p"
                      or (toks(l)[-1:] == ["p"] and ("comp" in l or "s&p" in l or "500" in l)),
                      "nominal price P"),
        "d": find_col(lambda l: "real" not in toks(l)
                      and (l in ("d", "dividend") or toks(l)[-1:] == ["d"]),
                      "dividend D"),
        "e": find_col(lambda l: "real" not in toks(l) and "scaled" not in toks(l)
                      and (l in ("e", "earnings") or toks(l)[-1:] == ["e"]),
                      "earnings E"),
        "cpi": find_col(lambda l: "cpi" in toks(l) or "consumer price" in l, "CPI"),
        "gs10": find_col(lambda l: "gs10" in toks(l) or "long interest rate" in l, "GS10"),
        "real_price": find_col(lambda l: "real" in toks(l) and "price" in toks(l)
                               and "return" not in toks(l) and "home" not in toks(l)
                               and "cape" not in toks(l), "real price"),
        "tr": find_col(lambda l: "return price" in l
                       and "bond" not in toks(l) and "cape" not in toks(l)
                       and "scaled" not in toks(l), "real total return price"),
        "cape": find_col(lambda l: (l == "cape" or toks(l)[-1:] == ["cape"])
                         and "tr" not in toks(l) and "total" not in toks(l)
                         and "excess" not in toks(l) and "yield" not in toks(l),
                         "CAPE"),
    }
    print("Column mapping:", {k: f"{v} ({labels[v]!r})" for k, v in cols.items()})

    def num(cell):
        if cell is None or isinstance(cell, str):
            return None
        try:
            f = float(cell)
        except (TypeError, ValueError):
            return None
        return None if math.isnan(f) else f

    rows = []
    for i in range(header_row + 1, len(raw)):
        dval = num(raw.iat[i, cols["date"]])
        if dval is None or not (1800 <= dval <= 2200):
            continue
        rows.append({
            "date": decimal_date_to_iso(dval),
            "p": num(raw.iat[i, cols["p"]]),
            "d": num(raw.iat[i, cols["d"]]),
            "e": num(raw.iat[i, cols["e"]]),
            "cpi": num(raw.iat[i, cols["cpi"]]),
            "gs10": num(raw.iat[i, cols["gs10"]]),
            "real_price": num(raw.iat[i, cols["real_price"]]),
            "tr": num(raw.iat[i, cols["tr"]]),
            "cape": num(raw.iat[i, cols["cape"]]),
        })
    rows = [r for r in rows if r["p"] is not None]
    if not rows:
        raise RuntimeError("No data rows parsed from the Data sheet")
    return rows


# --------------------------------------------------------------- validate --

def months_between(d1: str, d2: str) -> int:
    y1, m1 = int(d1[:4]), int(d1[5:7])
    y2, m2 = int(d2[:4]), int(d2[5:7])
    return (y2 - y1) * 12 + (m2 - m1)


def add_months(d: str, n: int) -> str:
    y, m = int(d[:4]), int(d[5:7])
    m += n
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    return f"{y:04d}-{m:02d}-01"


def ann_return(v0, v1, years):
    return ((v1 / v0) ** (1.0 / years) - 1.0) * 100.0


def validate_against_previous(rows, prev_slim, prev_decomp):
    """Refuse to proceed unless frozen history matches the previous data.
    This is the guard against a silent layout change / mis-mapped column."""
    problems = []
    by_date = {r["date"]: r for r in rows}
    prev_by_date = {r["d"]: r for r in prev_slim["d"]}

    # 1. Structure
    if rows[0]["date"] != "1871-01-01":
        problems.append(f"first row is {rows[0]['date']}, expected 1871-01-01")
    for a, b in zip(rows, rows[1:]):
        if months_between(a["date"], b["date"]) != 1:
            problems.append(f"non-consecutive months: {a['date']} -> {b['date']}")
            break
    if len(rows) < len(prev_slim["d"]):
        problems.append(f"row count shrank: {len(rows)} < {len(prev_slim['d'])}")

    # 2. Frozen pass-through values, 1871-1880 (never revised)
    for dt in [f"{y:04d}-{m:02d}-01" for y in range(1871, 1881) for m in (1, 7)]:
        new, old = by_date.get(dt), prev_by_date.get(dt)
        if not new or not old:
            problems.append(f"missing frozen month {dt}")
            continue
        for xls_key, slim_key, nd in (("p", "p", 2), ("d", "div", 2), ("e", "e", 2),
                                      ("cpi", "cpi", 1), ("gs10", "gs", 2)):
            nv, ov = new[xls_key], old[slim_key]
            if nv is None or ov is None or abs(round(nv, nd) - ov) > 10 ** -nd + 1e-9:
                problems.append(f"{dt} {xls_key}: parsed {nv} vs previous {ov}")

    # 3. Frozen CAPE (1881-1890)
    for dt in [f"{y:04d}-01-01" for y in range(1881, 1891)]:
        new, old = by_date.get(dt), prev_by_date.get(dt)
        if new and old and new["cape"] is not None and old["c"] is not None:
            if abs(new["cape"] - old["c"]) > 0.05:
                problems.append(f"{dt} CAPE: parsed {new['cape']} vs previous {old['c']}")

    # 4. Real price column: rp*CPI/P must be ~constant (the CPI rebase base)
    ratios = [r["real_price"] * r["cpi"] / r["p"] for r in rows[:1200]
              if r["real_price"] and r["cpi"] and r["p"]]
    if ratios and (max(ratios) / min(ratios) - 1) > 0.01:
        problems.append("real price column fails rp*CPI/P constancy check "
                        f"(spread {max(ratios)/min(ratios)-1:.4%})")

    # 5. Frozen 10y total-return decompositions (CPI-base independent)
    for dt in [f"{y:04d}-01-01" for y in range(1885, 1990, 10)]:
        dt2 = add_months(dt, 120)
        a, b = by_date.get(dt), by_date.get(dt2)
        old = prev_decomp["decomp10"].get(dt)
        if a and b and old and a["tr"] and b["tr"]:
            tr = ann_return(a["tr"], b["tr"], 10)
            if abs(tr - old["tr"]) > 0.05:
                problems.append(f"{dt} 10y TR: computed {tr:.2f} vs previous {old['tr']}")

    # 6. Latest-month sanity
    last = rows[-1]
    if last["cape"] is not None and not (3 < last["cape"] < 80):
        problems.append(f"latest CAPE {last['cape']} outside sanity range")

    if problems:
        for p in problems:
            print("VALIDATION FAILURE:", p, file=sys.stderr)
        raise SystemExit(1)
    print(f"Validation passed ({len(rows)} rows, "
          f"{rows[0]['date']} to {rows[-1]['date']})")


# ----------------------------------------------------------------- build --

def r2(x):
    return None if x is None else round(x, 2)


def build_outputs(rows, prev_slim):
    by_date = {r["date"]: r for r in rows}

    def fwd(t, months, col):
        a = by_date.get(t)
        b = by_date.get(add_months(t, months))
        if not a or not b or a[col] is None or b[col] is None:
            return None
        return ann_return(a[col], b[col], months / 12.0)

    # slim
    d_list = []
    for r in rows:
        d_list.append({
            "d": r["date"],
            "p": r2(r["p"]),
            "rp": None if r["real_price"] is None else round(r["real_price"], 1),
            "c": r2(r["cape"]),
            "cy": None if r["cape"] is None else round(100.0 / r["cape"], 2),
            "dy": None if (r["d"] is None or not r["p"]) else round(100.0 * r["d"] / r["p"], 2),
            "ey": None if (r["e"] is None or not r["p"]) else round(100.0 * r["e"] / r["p"], 2),
            "gs": r2(r["gs10"]),
            "cpi": None if r["cpi"] is None else round(r["cpi"], 1),
            "e": r2(r["e"]),
            "div": r2(r["d"]),
            "fr": r2(fwd(r["date"], 120, "tr")),
        })
    cape_start = next((x["d"] for x in d_list if x["c"] is not None), None)
    slim = {
        "d": d_list,
        "r": prev_slim["r"],  # static NBER recession list; extend manually
        "m": {
            "source": "Robert J. Shiller, Yale University",
            "url": "http://www.econ.yale.edu/~shiller/data.htm",
            "description": "U.S. Stock Markets 1871-Present and CAPE Ratio",
            "last_updated": d_list[-1]["d"],
            "total_records": len(d_list),
            "cape_start": cape_start,
        },
    }

    # horizon
    horizon = []
    for r in rows:
        if r["cape"] is None:
            continue
        horizon.append({
            "cy": round(100.0 / r["cape"], 2),
            "d": r["date"],
            "f": [r2(fwd(r["date"], 12 * k, "tr")) for k in range(1, 11)],
        })

    # decomp
    series = []
    for r in rows:
        if r["real_price"] is None or r["tr"] is None:
            continue
        row = [r["date"], round(r["real_price"], 2), round(r["tr"], 2)]
        if r["cape"] is not None:
            row.append(round(r["cape"], 2))
        series.append(row)

    decomp10 = {}
    for r in rows:
        t, t2 = r["date"], add_months(r["date"], 120)
        b = by_date.get(t2)
        if (r["cape"] is None or not b or b["cape"] is None
                or r["tr"] is None or b["tr"] is None
                or r["real_price"] is None or b["real_price"] is None):
            continue
        tr = ann_return(r["tr"], b["tr"], 10)
        cr = ann_return(r["real_price"], b["real_price"], 10)
        ir = ((1 + tr / 100) / (1 + cr / 100) - 1) * 100
        mc = ann_return(r["cape"], b["cape"], 10)
        eg = ((1 + cr / 100) / (1 + mc / 100) - 1) * 100
        fr = ((1 + ir / 100) * (1 + eg / 100) - 1) * 100
        decomp10[t] = {"tr": r2(tr), "cr": r2(cr), "ir": r2(ir),
                       "mc": r2(mc), "eg": r2(eg), "fr": r2(fr)}

    decomp = {"series": series, "decomp10": decomp10}
    return slim, horizon, decomp


# ------------------------------------------------------------------ write --

def dumps(obj) -> str:
    return json.dumps(obj, separators=(",", ":"), ensure_ascii=False)


def replace_block(html: str, script_id: str, var_name: str, payload: str) -> str:
    pattern = re.compile(
        r'(<script id="' + re.escape(script_id) + r'">const ' + re.escape(var_name)
        + r' = ).*?(;</script>)', re.S)
    if not pattern.search(html):
        raise RuntimeError(f"Embedded block {script_id}/{var_name} not found in HTML")
    return pattern.sub(lambda m: m.group(1) + payload + m.group(2), html, count=1)


def write_if_changed(path: Path, content: str) -> bool:
    old = path.read_text(encoding="utf-8") if path.exists() else None
    if old == content:
        print(f"  unchanged: {path.name}")
        return False
    path.write_text(content, encoding="utf-8")
    print(f"  updated:   {path.name}")
    return True


def update_files(slim, horizon, decomp):
    changed = False
    changed |= write_if_changed(SLIM_PATH, dumps(slim))
    changed |= write_if_changed(HORIZON_PATH, dumps(horizon))
    changed |= write_if_changed(DECOMP_PATH, dumps(decomp))

    # Index.html embeds the FULL slim + horizon datasets
    html = HTML_CAPITAL.read_text(encoding="utf-8")
    html = replace_block(html, "data-script", "SHILLER_DATA", dumps(slim))
    html = replace_block(html, "horizon-data-script", "HORIZON_DATA", dumps(horizon))
    changed |= write_if_changed(HTML_CAPITAL, html)

    # index.html embeds a 2011+ subset of slim/decomp series, full decomp10
    slim_sub = {
        "d": [r for r in slim["d"] if r["d"] >= EMBED_CUTOFF],
        "r": [r for r in slim["r"] if r[0] >= EMBED_CUTOFF],
    }
    decomp_sub = {
        "series": [r for r in decomp["series"] if r[0] >= EMBED_CUTOFF],
        "decomp10": decomp["decomp10"],
    }
    html = HTML_LOWER.read_text(encoding="utf-8")
    html = replace_block(html, "data-script", "SHILLER_DATA", dumps(slim_sub))
    html = replace_block(html, "decomp-data-script", "DECOMP_DATA", dumps(decomp_sub))
    changed |= write_if_changed(HTML_LOWER, html)
    return changed


# ------------------------------------------------------------------- main --

def main(xls_path: str | None = None):
    prev_slim = json.loads(SLIM_PATH.read_text(encoding="utf-8"))
    prev_decomp = json.loads(DECOMP_PATH.read_text(encoding="utf-8"))

    if xls_path:
        path = Path(xls_path)
        print(f"Using local file: {path}")
    else:
        path = REPO_ROOT / "ie_data_download.xls"
        download_xls(path)

    rows = load_sheet(path)
    validate_against_previous(rows, prev_slim, prev_decomp)
    slim, horizon, decomp = build_outputs(rows, prev_slim)

    print(f"slim: {len(slim['d'])} months, last = {slim['m']['last_updated']}")
    print(f"horizon: {len(horizon)} months")
    print(f"decomp: series {len(decomp['series'])}, decomp10 {len(decomp['decomp10'])}")

    changed = update_files(slim, horizon, decomp)
    if not xls_path:
        path.unlink(missing_ok=True)  # don't leave the raw xls in the repo
    print("DATA_CHANGED" if changed else "NO_CHANGES")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
