"""Download official NYSDOT / NYC DDC / NYC DOT / NYC DEP standards into Standards/.

Data-driven: SOURCES lists index pages (crawled one level for document links)
and direct files. Every file lands in Standards/<agency>/<collection>/ and gets a
row in Standards/registry.json (url, path, bytes, sha256, fetched, edition).

Current editions only by default. Superseded/historical items are skipped
(filters below) but recorded as `skipped` so they can be pulled later.
Never deletes or overwrites a file with different content: a changed remote
file is saved alongside with a date suffix and both stay in the registry.

Usage:
  python scripts/standards_fetch.py            # fetch everything missing
  python scripts/standards_fetch.py --dry-run  # list what would be fetched
  python scripts/standards_fetch.py --only hdm
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
from datetime import date

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STD = os.path.join(ROOT, "Standards")
REGISTRY = os.path.join(STD, "registry.json")
UA = {"User-Agent": "Mozilla/5.0 (standards_fetch; NYSDOT/DDC reference library)"}
DOC_RE = re.compile(r"\.(pdf|zip|xlsx?|docx?)($|\?)", re.I)

# key: (agency, collection, kind, url, link_filter_regex or None)
#   kind = "page"  -> crawl page for document links matching filter
#   kind = "pages" -> crawl page for SUB-pages matching filter, then their documents
#   kind = "file"  -> direct download
SOURCES = {
    # ---- NYSDOT Highway Design Manual (every chapter page + appendices) ----
    "hdm": ("NYSDOT", "hdm", "pages",
            "https://www.dot.ny.gov/divisions/engineering/design/dqab/hdm",
            r"/dqab/hdm/(chapter|Chapter)|hdm-repository/"),
    # ---- NYSDOT e-Standards (current letting edition only) ----
    "e-standards": ("NYSDOT", "specs-and-std-sheets", "page",
                    "https://www.dot.ny.gov/main/business-center/engineering/specifications/busi-e-standards-usc",
                    r"2026_5_(specs|stdsht)_usc"),
    "specs-tracked": ("NYSDOT", "specs-tracked-changes", "page",
                      "https://www.dot.ny.gov/main/business-center/engineering/specifications/updated-standard-specifications-us",
                      r"2026_9_specs_usc_tc|/0130[01]e?\.pdf"),
    # ---- NYSDOT design manuals ----
    "pdm": ("NYSDOT", "pdm", "pages", "https://www.dot.ny.gov/divisions/engineering/design/dqab/pdm",
            r"/dqab/pdm|dqab-repository|pdm-repository|\.pdf"),
    "cpdm": ("NYSDOT", "cpdm", "pages", "https://www.dot.ny.gov/divisions/engineering/design/dqab/cpdm",
             r"/dqab/cpdm|repository|\.pdf"),
    "rmm": ("NYSDOT", "reference-marker", "page", "https://www.dot.ny.gov/divisions/engineering/design/dqab/rmm", None),
    "utility": ("NYSDOT", "utilities", "pages", "https://www.dot.ny.gov/divisions/engineering/design/dqab/util-info",
                r"/util-info|repository|\.pdf"),
    "survey": ("NYSDOT", "land-survey", "pages",
               "https://www.dot.ny.gov/divisions/engineering/design/designmappingautomationbureau/land-survey/standards-procedures",
               r"standards-procedures|repository|\.pdf"),
    "3r": ("NYSDOT", "3r", "file",
           "https://www.dot.ny.gov/divisions/engineering/design/dqab/dqab-repository/standards%20for%20non-freeway3r.pdf", None),
    "design-consultant-manual": ("NYSDOT", "consultant", "file",
           "https://www.dot.ny.gov/divisions/engineering/design/consultant-management/repository/Design_Consultant_Management_Manual.pdf", None),
    "photogrammetry": ("NYSDOT", "survey-specs", "page",
                       "https://www.dot.ny.gov/divisions/engineering/design/documents-and-forms", r"photogrammetry/repository"),
    "cadd-resources": ("NYSDOT", "cadd", "page",
                       "https://www.dot.ny.gov/main/business-center/engineering/cadd-info/resources-settings", None),
    # not linked from the HDM index page (2026-10-23) — direct
    "hdm-16": ("NYSDOT", "hdm", "file",
               "https://www.dot.ny.gov/divisions/engineering/design/dqab/hdm/hdm-repository/chapt_16.pdf", None),
    # ---- NYC DEP stormwater (DDC GR §2.4.H) ----
    "nyc-stormwater-manual": ("NYC-DEP", "stormwater", "file",
        "https://www.nyc.gov/assets/dep/downloads/pdf/water/stormwater/unified-stormwater-rule/uswr_nyc_stormwater_manual.pdf", None),
    "nyc-stormwater-appendices": ("NYC-DEP", "stormwater", "page",
        "https://www.nyc.gov/site/dep/water/unified-stormwater-rule.page", r"unified-stormwater-rule/"),
    # NYC DOT Street Design Manual 4th ed. is web-only: https://www.nycstreetdesign.info/ (not downloadable as PDF)
    # ---- NYC DDC current infrastructure design standards + bulletins ----
    "ddc-bulletins": ("NYC-DDC", "spec-bulletins", "page",
                      "https://www.nyc.gov/site/ddc/resources/specification-bulletins.page", None),
    "ddc-current": ("NYC-DDC", "infrastructure-standards", "page",
                    "https://www.nyc.gov/site/ddc/resources/publications.page",
                    r"Sept12026|SB22-004-NYCDOTSTANDARDDETAILS.*2022-06-06|NYCDEPSTANDARDSEWERANDWATERMAINSPECIFICATIONS|"
                    r"RevisedSep2025|TrunkMainWorkJune1-2026|WATERMAINSTDDWG-20240919|sewer_gen_specs11|sewer_instr_arch|"
                    r"GreenInfrastructureSpecificationsAugust0124|Design_Consultant_Guide_2025|high-performance-infra|bim-guide|"
                    r"active-design-guidelines|ddc-water-matters"),
}

# never fetch (superseded / not standards)
SKIP_RE = re.compile(r"superseded|safety|summit|blueprint|presentation|press-releases|standdown|fall", re.I)


def get(url: str) -> bytes:
    req = urllib.request.Request(url.replace(" ", "%20"), headers=UA)
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read()


def page_links(url: str) -> list[tuple[str, str]]:
    try:
        html = get(url).decode("utf-8", "ignore")
    except Exception as ex:
        print(f"  page fail {url}: {ex}")
        return []
    out = []
    for href, text in re.findall(r"""href\s*=\s*["']([^"'#]+)["'][^>]*>(.*?)</a>""", html, re.S | re.I):
        text = re.sub(r"<[^>]+>|&nbsp;", " ", text)
        text = re.sub(r"\s+", " ", text).strip()
        out.append((urllib.parse.urljoin(url, href.replace("&amp;", "&")), text))
    return out


def plan(only: set[str] | None) -> list[dict]:
    items, seen = [], set()

    def add(key, agency, coll, url, title):
        u = url.split("#")[0]
        if u in seen:
            return
        seen.add(u)
        items.append({"source": key, "agency": agency, "collection": coll, "url": u, "title": title,
                      "skip": bool(SKIP_RE.search(u + " " + title))})

    for key, (agency, coll, kind, url, filt) in SOURCES.items():
        if only and key not in only:
            continue
        fre = re.compile(filt) if filt else None
        if kind == "file":
            add(key, agency, coll, url, key)
            continue
        links = page_links(url)
        pages = [url]
        if kind == "pages":
            dom = urllib.parse.urlparse(url).netloc
            for a, t in links:
                if DOC_RE.search(a):
                    continue
                if urllib.parse.urlparse(a).netloc.endswith(dom.replace("www.", "")) and fre and fre.search(a):
                    pages.append(a)
        for p in dict.fromkeys(pages):
            for a, t in (links if p == url else page_links(p)):
                if not DOC_RE.search(a):
                    continue
                if fre and kind == "page" and not fre.search(a):
                    continue
                if "dot.ny.gov" not in a and "nyc.gov" not in a:
                    continue
                add(key, agency, coll, a, t or os.path.basename(a))
            time.sleep(0.3)
    return items


def load_registry() -> dict:
    if os.path.exists(REGISTRY):
        with open(REGISTRY, encoding="utf-8") as f:
            return json.load(f)
    return {"_comment": "Official standards library. Binaries gitignored; re-fetch with scripts/standards_fetch.py.",
            "documents": []}


def save_registry(reg: dict) -> None:
    reg["documents"].sort(key=lambda d: (d["agency"], d["collection"], d["path"]))
    with open(REGISTRY, "w", encoding="utf-8") as f:
        json.dump(reg, f, indent=1)
        f.write("\n")


def fetch(items: list[dict], dry: bool) -> None:
    reg = load_registry()
    by_url = {d["url"]: d for d in reg["documents"]}
    n_new = n_skip = n_have = n_fail = 0
    for it in items:
        if it["skip"]:
            n_skip += 1
            continue
        if it["url"] in by_url and os.path.exists(os.path.join(ROOT, by_url[it["url"]]["path"])):
            n_have += 1
            continue
        name = urllib.parse.unquote(os.path.basename(urllib.parse.urlparse(it["url"]).path)) or "index.pdf"
        rel = os.path.join("Standards", it["agency"].lower(), it["collection"], name)
        if dry:
            print(f"  would fetch {rel}  <- {it['url']}")
            continue
        try:
            data = get(it["url"])
        except Exception as ex:
            print(f"  FAIL {it['url']}: {ex}")
            n_fail += 1
            continue
        dest = os.path.join(ROOT, rel)
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        sha = hashlib.sha256(data).hexdigest()
        if os.path.exists(dest):
            with open(dest, "rb") as f:
                if hashlib.sha256(f.read()).hexdigest() != sha:
                    b, e = os.path.splitext(rel)
                    rel = f"{b}_{date.today().isoformat()}{e}"
                    dest = os.path.join(ROOT, rel)
        with open(dest, "wb") as f:
            f.write(data)
        row = {k: it[k] for k in ("agency", "collection", "title", "url", "source")}
        row.update(path=rel.replace("\\", "/"), bytes=len(data), sha256=sha, fetched=date.today().isoformat())
        reg["documents"] = [d for d in reg["documents"] if d["url"] != it["url"]] + [row]
        by_url[it["url"]] = row
        n_new += 1
        print(f"  ok {len(data)/1e6:6.1f} MB  {rel}")
        if n_new % 10 == 0:
            save_registry(reg)
        time.sleep(0.5)
    if not dry:
        save_registry(reg)
    print(f"DONE new={n_new} have={n_have} skipped={n_skip} failed={n_fail}")


def register_local() -> None:
    """Register standards already on disk (Standards/ root + Project Documentation PDFs)."""
    reg = load_registry()
    have = {d["path"] for d in reg["documents"]}
    local = [
        ("NYC-DDC", "general-requirements", "Standards/INFRA-GR-APRIL 2024 clean copy 1.pdf",
         "DDC Infrastructure General Requirements for Engineering Design, April 2024"),
        ("NYSDOT", "cadd", "Standards/NYSDOT_Configuration2024_01.zip", "NYSDOT CONNECT Configuration 2024"),
        ("FHWA", "mutcd", "Project Documentation/part6.pdf", "MUTCD Part 6 Temporary Traffic Control"),
        ("NYSDOT", "mutcd-supplement", "Project Documentation/B-2011Supplement-adopted.pdf", "NYS MUTCD Supplement"),
        ("NYSDOT", "specs-and-std-sheets", "Project Documentation/2026_1_stdsht_usc_book_3.pdf",
         "Standard Sheets Book 3 (Jan 2026)"),
    ]
    for agency, coll, rel, title in local:
        p = os.path.join(ROOT, rel)
        if rel in have or not os.path.exists(p):
            continue
        with open(p, "rb") as f:
            data = f.read()
        reg["documents"].append({"agency": agency, "collection": coll, "title": title, "url": "", "source": "local",
                                 "path": rel, "bytes": len(data), "sha256": hashlib.sha256(data).hexdigest(),
                                 "fetched": "local"})
    save_registry(reg)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--only", nargs="*")
    a = ap.parse_args()
    register_local()
    items = plan(set(a.only) if a.only else None)
    print(f"planned {len(items)} documents ({sum(i['skip'] for i in items)} skipped by filter)")
    fetch(items, a.dry_run)
    return 0


if __name__ == "__main__":
    sys.exit(main())
