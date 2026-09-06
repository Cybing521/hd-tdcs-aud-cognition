#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Bibliography validation for the symbolic citation system.

Checks, per REF_DB entry:
1. The key is actually cited somewhere in EN/ZH manuscript text ([@key] tokens).
2. DOI resolves on Crossref and the deposited title matches the stored reference.
3. The canonical numbering (nummap) covers every key with no orphans.

Output: 04_data/processed/ref_validation.json +
        ccfa-review-reports/ref_validation_2026-08-24.md
"""

from __future__ import annotations

import json
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))

from manuscript_content import (  # noqa: E402
    DISCUSSION_EN,
    DISCUSSION_ZH,
    INTRO_EN,
    INTRO_ZH,
    REF_DB,
    assign_reference_numbers,
)
from build_full_manuscript import collect_cite_texts, compute_post_hoc_power  # noqa: E402

DOI_RE = re.compile(r"(?i)(?:doi\s*:\s*|https?://(?:dx\.)?doi\.org/)(10\.\d{4,9}/[^\s<>\"']+)")
KEY_RE = re.compile(r"\[@([^\]]+)\]")

STATS_PATH = ROOT / "04_data" / "processed" / "analysis_results.json"


def normalize_title(title: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", title.lower())


def fetch_crossref(doi: str) -> tuple[str, dict | None]:
    url = f"https://api.crossref.org/works/{urllib.parse.quote(doi.lower())}"
    req = urllib.request.Request(url, headers={"User-Agent": "ref-validator/2.0 (mailto:research@example.com)"})
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            payload = json.loads(resp.read().decode())
            return "ok", payload.get("message", {})
    except urllib.error.HTTPError as exc:
        return f"http_{exc.code}", None
    except Exception as exc:  # noqa: BLE001
        return f"error:{type(exc).__name__}", None


def title_overlap(a: str, b: str) -> float:
    ta = normalize_title(a)
    tb = normalize_title(b)
    if not ta or not tb:
        return 0.0
    shorter, longer = (ta, tb) if len(ta) <= len(tb) else (tb, ta)
    return 1.0 if shorter in longer else len(set(shorter) & set(longer)) / max(len(set(longer)), 1)


def infer_title(ref: str) -> str:
    """Reference format 'Authors. Title. Journal. Year;Vol(Issue):Pages. doi:...'."""
    after_authors = ref.split(". ", 1)
    body = after_authors[1] if len(after_authors) > 1 else ref
    return body.split(". ", 1)[0].strip()


def crossref_year(meta: dict) -> int | None:
    for key in ("published-print", "published-online", "created", "issued"):
        parts = (meta.get(key) or {}).get("date-parts") or []
        if parts and parts[0]:
            return parts[0][0]
    return None


def main() -> int:
    data = json.loads(STATS_PATH.read_text(encoding="utf-8"))

    power = compute_post_hoc_power(data)
    # Canonical numbering exactly as the build computes it.
    en_doc_texts = collect_cite_texts(data, power)
    nummap = assign_reference_numbers(en_doc_texts)
    unassigned = sorted(set(REF_DB) - set(nummap))

    # Citation coverage across ALL languages (EN drives order; ZH must be a subset).
    def keys_in(texts):
        return {
            tok.lstrip("@")
            for t in texts
            for grp in KEY_RE.findall(t)
            for tok in grp.split(",")
            if tok.strip()
        }

    zh_texts = list(INTRO_ZH) + list(DISCUSSION_ZH)
    cited_en = keys_in(en_doc_texts)
    cited_zh = keys_in(zh_texts)

    records = []
    for key in sorted(REF_DB, key=lambda k: nummap.get(k, 999)):
        ref = REF_DB[key]
        doi_match = DOI_RE.search(ref)
        doi = doi_match.group(1).rstrip(".,;") if doi_match else None
        entry = {
            "key": key,
            "ref_num": nummap.get(key),
            "reference": ref,
            "doi": doi,
            "cited_en": key in cited_en,
            "cited_zh": key in cited_zh,
        }
        if doi:
            status, meta = fetch_crossref(doi)
            entry["crossref_status"] = status
            if meta:
                cr_title = " ".join(meta.get("title") or [])
                overlap = title_overlap(infer_title(ref), cr_title)
                entry["crossref_title"] = cr_title
                entry["crossref_year"] = crossref_year(meta)
                entry["title_overlap"] = round(overlap, 3)
                if status == "ok" and overlap >= 0.35:
                    entry["final_status"] = "VERIFIED"
                elif status == "ok":
                    entry["final_status"] = "EXISTS_METADATA_CONFLICT"
                else:
                    entry["final_status"] = "UNRESOLVED"
            else:
                entry["final_status"] = "UNRESOLVED"
            time.sleep(0.25)
        else:
            entry["crossref_status"] = "no_doi"
            entry["final_status"] = "MANUAL_VERIFY"
        records.append(entry)

    uncited = [k for k in sorted(REF_DB) if k not in cited_en]
    zh_orphans = sorted(cited_zh - set(REF_DB))
    conflicts = [r for r in records if r["final_status"] == "EXISTS_METADATA_CONFLICT"]

    out = {
        "checked_at": datetime.now(timezone.utc).isoformat(),
        "system": "symbolic [@key] citations; Vancouver numbering from EN first-appearance order",
        "reference_count": len(REF_DB),
        "assigned_count": len(nummap),
        "unassigned_keys": unassigned,
        "uncited_keys": uncited,
        "zh_orphan_keys": zh_orphans,
        "doi_count": sum(1 for r in records if r["doi"]),
        "verified": sum(1 for r in records if r["final_status"] == "VERIFIED"),
        "manual_verify": sum(1 for r in records if r["final_status"] == "MANUAL_VERIFY"),
        "conflicts": conflicts,
        "records": records,
    }

    out_json = ROOT / "04_data/processed/ref_validation.json"
    out_json.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")

    lines = [
        "# Reference Validation Report (symbolic system)",
        "",
        f"**Checked at:** {out['checked_at']}",
        f"**Numbering:** EN first-appearance order via `assign_reference_numbers`",
        "",
        "## Summary",
        "",
        "| Metric | Value |",
        "|:---|:---|",
        f"| REF_DB entries | {out['reference_count']} |",
        f"| Keys assigned numbers | {out['assigned_count']} |",
        f"| DOI-bearing entries | {out['doi_count']} |",
        f"| Crossref VERIFIED | {out['verified']} |",
        f"| Manual verify (no DOI) | {out['manual_verify']} |",
        f"| Metadata conflicts | {len(conflicts)} |",
        f"| Uncited keys | {len(uncited)} |",
        f"| ZH-only orphan keys | {len(zh_orphans)} |",
        "",
    ]
    if unassigned or uncited or zh_orphans:
        lines.append(f"- Unassigned: {unassigned}")
        lines.append(f"- Uncited (EN): {uncited}")
        lines.append(f"- ZH orphans: {zh_orphans}")
        lines.append("")
    if conflicts:
        lines += ["## Metadata conflicts (action required)", ""]
        for r in conflicts:
            lines.append(f"- **#{r['ref_num']} `{r['key']}`** DOI `{r['doi']}` — overlap {r.get('title_overlap')}")
            lines.append(f"  - Listed: {r['reference'][:120]}...")
            lines.append(f"  - Crossref: {r.get('crossref_title', '')[:120]}")
        lines.append("")

    lines += ["## Record table", "", "| # | Key | Status | DOI | EN | ZH |", "|:---:|:---|:---|:---|:---:|:---:|"]
    for r in records:
        doi = r["doi"] or "—"
        lines.append(
            f"| {r['ref_num']} | {r['key']} | {r['final_status']} | {doi} "
            f"| {'Y' if r['cited_en'] else 'N'} | {'Y' if r['cited_zh'] else 'N'} |"
        )

    lines += ["", "## Manual-verify queue (no DOI)", ""]
    for r in records:
        if r["final_status"] == "MANUAL_VERIFY":
            lines.append(f"- #{r['ref_num']} `{r['key']}` {r['reference'][:100]}")

    report = ROOT / "ccfa-review-reports/ref_validation_2026-08-24.md"
    report.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {out_json}")
    print(f"Wrote {report}")
    print(
        f"VERIFIED={out['verified']} MANUAL={out['manual_verify']} "
        f"CONFLICTS={len(conflicts)} UNCITED={len(uncited)} UNASSIGNED={len(unassigned)}"
    )
    ok = not (conflicts or uncited or unassigned or zh_orphans)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
