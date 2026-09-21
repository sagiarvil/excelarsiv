#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
EXCELARSIV QUALITY FACTORY V1
=============================

Amaç:
- EXCELARSIV_PRODUCTS altındaki tüm Excel ürünlerini otomatik envanterlemek.
- Duplicate / kopya workbook'ları SHA-256 ile tespit etmek.
- OOXML bütünlüğü, formül hataları, harici linkler, sheet mimarisi,
  tab renkleri, gridline, gerçek hyperlink navigasyonu ve iş akışı sırasını denetlemek.
- Kaynak dosyaya dokunmadan güvenli "CALIBRATED" kopya üretmek.
- Her ürün için JSON + HTML kanıt, merkezi CSV manifest üretmek.
- Fail-closed release mantığı uygulamak.

Not:
"500 USD kalite sınıfı" ürünün mühendislik, sunum, denetim ve kullanılabilirlik
hedefidir; gerçek ödeme isteği ancak pazar/satış verisiyle doğrulanabilir.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import html
import json
import re
import sys
import zipfile
from dataclasses import dataclass, asdict, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)

ERROR_TOKENS = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!"}

PALETTE = {
    "START": "FF0F2942",
    "SUMMARY": "FF17324D",
    "INPUT": "FF4472C4",
    "ACCOUNTING": "FF2F75B5",
    "CALC": "FF0F6B78",
    "REPORT": "FF1F4E78",
    "SCENARIO": "FF8A5A1F",
    "DECISION": "FF5B477E",
    "CONTROL": "FF9B2C2C",
    "DOC": "FF7F8C8D",
    "ORACLE": "FF548235",
    "VALUE": "FF997300",
    "OTHER": "FF6B7280",
}

ORDER_RULES = [
    (0,   ["baslangic", "ana_sayfa", "anasayfa", "home", "start"]),
    (10,  ["yonetici_ozeti", "ozet", "dashboard", "panel"]),
    (20,  ["varsayim", "assumption", "girdi", "input"]),
    (30,  ["tarihsel", "historical", "mizan"]),
    (40,  ["thp", "esleme", "mapping"]),
    (50,  ["isletme_sermayesi", "working_capital"]),
    (55,  ["capex", "amortisman"]),
    (60,  ["borc", "finansman", "debt"]),
    (70,  ["gelir_tablosu", "income"]),
    (75,  ["bilanco", "balance"]),
    (80,  ["nakit_akis", "cash_flow"]),
    (90,  ["senaryo", "stres", "scenario", "stress"]),
    (100, ["karar", "decision"]),
    (110, ["kontrol", "check", "release"]),
    (120, ["kilavuz", "guide", "kullanim"]),
    (130, ["meta", "model_meta"]),
    (140, ["oracle", "dogrulama", "validation"]),
    (150, ["deger", "value"]),
]

IGNORE_DIRS = {
    "_SYSTEM", "_QUARANTINE", "_RELEASED", ".git", "node_modules",
    "__pycache__", "QA", "ARCHIVE", "KANIT", "kaynak"
}

@dataclass
class Gate:
    code: str
    severity: str
    passed: bool
    message: str
    evidence: str = ""

@dataclass
class Audit:
    product: str
    path: str
    sha256: str
    size_bytes: int
    modified_at: str
    fmt: str
    zip_ok: bool = False
    sheets: List[str] = field(default_factory=list)
    hidden_sheets: List[str] = field(default_factory=list)
    tab_colors: Dict[str, Optional[str]] = field(default_factory=dict)
    formula_count: int = 0
    error_cells: List[str] = field(default_factory=list)
    external_link_parts: int = 0
    external_formula_count: int = 0
    hyperlink_formula_count: int = 0
    hyperlink_rel_count: int = 0
    chart_count: int = 0
    table_count: int = 0
    conditional_format_count: int = 0
    data_validation_count: int = 0
    gridlines_on_count: int = 0
    freeze_panes_count: int = 0
    merged_ranges: int = 0
    order_mismatches: List[str] = field(default_factory=list)
    gates: List[Gate] = field(default_factory=list)
    status: str = "BLOCKED"
    notes: List[str] = field(default_factory=list)

def norm(s: str) -> str:
    tr = str.maketrans("ıİğĞüÜşŞöÖçÇ", "iIgGuUsSoOcC")
    s = s.translate(tr).lower()
    return re.sub(r"[^a-z0-9]+", "_", s).strip("_")

def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()

def resolve_xl_target(target: str) -> str:
    target = target.replace("\\", "/").lstrip("/")
    return target if target.startswith("xl/") else "xl/" + target

def parse_xml(data: bytes) -> ET.Element:
    return ET.fromstring(data)

def rel_map(data: bytes) -> Dict[str, str]:
    root = parse_xml(data)
    out = {}
    for rel in root.findall(f"{{{NS_PKG_REL}}}Relationship"):
        out[rel.attrib["Id"]] = resolve_xl_target(rel.attrib["Target"])
    return out

def category(sheet_name: str) -> str:
    n = norm(sheet_name)
    if any(x in n for x in ("baslangic","anasayfa","home","start")): return "START"
    if any(x in n for x in ("yonetici_ozeti","ozet","dashboard","panel")): return "SUMMARY"
    if any(x in n for x in ("varsayim","girdi","input","tarihsel","mizan")): return "INPUT"
    if any(x in n for x in ("thp","esleme","muhasebe")): return "ACCOUNTING"
    if any(x in n for x in ("isletme_sermayesi","capex","amortisman","borc","finansman","hesap")): return "CALC"
    if any(x in n for x in ("gelir_tablosu","bilanco","nakit_akis","rapor")): return "REPORT"
    if any(x in n for x in ("senaryo","stres")): return "SCENARIO"
    if any(x in n for x in ("karar","decision")): return "DECISION"
    if any(x in n for x in ("kontrol","release","gate")): return "CONTROL"
    if any(x in n for x in ("kilavuz","guide","meta")): return "DOC"
    if any(x in n for x in ("oracle","dogrulama","validation")): return "ORACLE"
    if any(x in n for x in ("deger","value")): return "VALUE"
    return "OTHER"

def order_rank(sheet_name: str, original_index: int) -> Tuple[int, int]:
    n = norm(sheet_name)
    for rank, tokens in ORDER_RULES:
        if any(tok in n for tok in tokens):
            return rank, original_index
    return 999, original_index

def expected_order(sheets: List[str]) -> List[str]:
    return [x for _, _, x in sorted((order_rank(s, i)[0], i, s) for i, s in enumerate(sheets))]

def audit_workbook(path: Path, product: str) -> Audit:
    audit = Audit(
        product=product,
        path=str(path),
        sha256=sha256(path),
        size_bytes=path.stat().st_size,
        modified_at=datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec="seconds"),
        fmt=path.suffix.lower(),
    )

    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        audit.gates.append(Gate("T001", "P1", False, "OOXML dışı format", path.suffix))
        return audit

    try:
        with zipfile.ZipFile(path, "r") as z:
            bad = z.testzip()
            if bad:
                audit.gates.append(Gate("T001", "P1", False, "ZIP/OOXML bütünlüğü bozuk", bad))
                return audit
            files = {name: z.read(name) for name in z.namelist()}
        audit.zip_ok = True
    except Exception as e:
        audit.gates.append(Gate("T001", "P1", False, "Workbook açılamadı", str(e)))
        return audit

    try:
        wb_root = parse_xml(files["xl/workbook.xml"])
        rels = rel_map(files["xl/_rels/workbook.xml.rels"])
    except Exception as e:
        audit.gates.append(Gate("T002", "P1", False, "Workbook XML okunamadı", str(e)))
        return audit

    sheets_node = wb_root.find(f"{{{NS_MAIN}}}sheets")
    infos = []

    for i, sn in enumerate(list(sheets_node) if sheets_node is not None else []):
        name = sn.attrib.get("name", f"Sheet{i+1}")
        state = sn.attrib.get("state", "visible")
        rid = sn.attrib.get(f"{{{NS_REL}}}id", "")
        target = rels.get(rid, "")
        audit.sheets.append(name)
        if state != "visible":
            audit.hidden_sheets.append(name)
        infos.append((name, state, target))

    audit.external_link_parts = len([
        x for x in files if x.startswith("xl/externalLinks/") and x.endswith(".xml")
    ])

    for name, state, target in infos:
        if not target or target not in files:
            audit.gates.append(Gate("T003", "P1", False, "Sheet XML eksik", f"{name}:{target}"))
            continue

        root = parse_xml(files[target])

        sheet_pr = root.find(f"{{{NS_MAIN}}}sheetPr")
        tab_color = None
        if sheet_pr is not None:
            tc = sheet_pr.find(f"{{{NS_MAIN}}}tabColor")
            if tc is not None:
                tab_color = tc.attrib.get("rgb") or tc.attrib.get("theme") or tc.attrib.get("indexed")
        audit.tab_colors[name] = tab_color

        sheet_views = root.find(f"{{{NS_MAIN}}}sheetViews")
        grid_on = True
        if sheet_views is not None:
            sv = sheet_views.find(f"{{{NS_MAIN}}}sheetView")
            if sv is not None:
                if sv.attrib.get("showGridLines") == "0":
                    grid_on = False
                pane = sv.find(f"{{{NS_MAIN}}}pane")
                if pane is not None and pane.attrib.get("state") in ("frozen","frozenSplit"):
                    audit.freeze_panes_count += 1
        if state == "visible" and grid_on:
            audit.gridlines_on_count += 1

        merge_cells = root.find(f"{{{NS_MAIN}}}mergeCells")
        if merge_cells is not None:
            audit.merged_ranges += int(merge_cells.attrib.get("count", len(list(merge_cells))))

        audit.conditional_format_count += len(root.findall(f"{{{NS_MAIN}}}conditionalFormatting"))
        dvs = root.find(f"{{{NS_MAIN}}}dataValidations")
        if dvs is not None:
            audit.data_validation_count += int(dvs.attrib.get("count", len(list(dvs))))

        hyperlinks = root.find(f"{{{NS_MAIN}}}hyperlinks")
        if hyperlinks is not None:
            audit.hyperlink_rel_count += len(list(hyperlinks))

        for cell in root.iter(f"{{{NS_MAIN}}}c"):
            ref = cell.attrib.get("r", "?")
            formula = cell.find(f"{{{NS_MAIN}}}f")
            if formula is not None:
                audit.formula_count += 1
                txt = formula.text or ""
                if "[" in txt and "]" in txt:
                    audit.external_formula_count += 1
                if txt.strip().upper().startswith("HYPERLINK("):
                    audit.hyperlink_formula_count += 1
            if cell.attrib.get("t") == "e":
                v = cell.find(f"{{{NS_MAIN}}}v")
                err = v.text if v is not None else ""
                audit.error_cells.append(f"{name}!{ref}:{err}")

    audit.chart_count = len([x for x in files if x.startswith("xl/charts/chart") and x.endswith(".xml")])
    audit.table_count = len([x for x in files if x.startswith("xl/tables/table") and x.endswith(".xml")])

    exp = expected_order(audit.sheets)
    if exp != audit.sheets:
        audit.order_mismatches = [
            "Mevcut: " + " > ".join(audit.sheets),
            "Beklenen: " + " > ".join(exp),
        ]

    visible = [s for s in audit.sheets if s not in audit.hidden_sheets]
    real_links = audit.hyperlink_formula_count + audit.hyperlink_rel_count

    audit.gates.extend([
        Gate("T001", "P1", audit.zip_ok, "OOXML paket bütünlüğü", "zip test"),
        Gate("T010", "P1", len(audit.error_cells) == 0,
             "Workbook içinde Excel hata hücresi olmamalı",
             ", ".join(audit.error_cells[:30])),
        Gate("T020", "P1", audit.external_link_parts == 0 and audit.external_formula_count == 0,
             "Kontrolsüz harici link bağımlılığı olmamalı",
             f"parts={audit.external_link_parts}; formulas={audit.external_formula_count}"),
        Gate("U010", "P2", all(audit.tab_colors.get(s) for s in visible),
             "Tüm görünür sheet tabları kurumsal renkte olmalı",
             f"renksiz={sum(1 for s in visible if not audit.tab_colors.get(s))}"),
        Gate("U020", "P2", audit.gridlines_on_count == 0,
             "Kullanıcı sheetlerinde gridline kapalı olmalı",
             f"gridlines_on={audit.gridlines_on_count}"),
        Gate("U030", "P2", not audit.order_mismatches,
             "Sheet sırası kullanıcı iş akışına uygun olmalı",
             " | ".join(audit.order_mismatches)),
        Gate("U040", "P2", len(visible) < 3 or real_links > 0,
             "Çok sayfalı workbook'ta gerçek tıklanabilir navigasyon olmalı",
             f"real_hyperlinks={real_links}"),
        Gate("M010", "P2", audit.formula_count > 0,
             "Ürün yalnız statik tablo olmamalı",
             f"formula_count={audit.formula_count}"),
    ])

    p1_fail = any(g.severity == "P1" and not g.passed for g in audit.gates)
    p2_fail = any(g.severity == "P2" and not g.passed for g in audit.gates)
    audit.status = "BLOCKED" if p1_fail else ("CALIBRATION_REQUIRED" if p2_fail else "READY_FOR_REVIEW")
    return audit

def calibrate_copy(src: Path, dest: Path) -> None:
    """
    Düşük riskli, içerik değiştirmeyen kalibrasyon:
    - native sheet tab renkleri
    - sheet iş akışı sırası
    - gridline kapatma
    - full recalculation
    Kaynak dosya ASLA yerinde değiştirilmez.
    """
    with zipfile.ZipFile(src, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    wb_root = parse_xml(files["xl/workbook.xml"])
    rels_root = parse_xml(files["xl/_rels/workbook.xml.rels"])

    rid_to_target = {}
    for rel in rels_root.findall(f"{{{NS_PKG_REL}}}Relationship"):
        rid_to_target[rel.attrib["Id"]] = resolve_xl_target(rel.attrib["Target"])

    sheets_node = wb_root.find(f"{{{NS_MAIN}}}sheets")
    nodes = list(sheets_node) if sheets_node is not None else []

    ordered_pairs = sorted(enumerate(nodes), key=lambda p: order_rank(p[1].attrib.get("name",""), p[0]))

    if sheets_node is not None:
        for node in list(sheets_node):
            sheets_node.remove(node)
        for _, node in ordered_pairs:
            sheets_node.append(node)

    views = wb_root.find(f"{{{NS_MAIN}}}bookViews")
    if views is not None:
        view = views.find(f"{{{NS_MAIN}}}workbookView")
        if view is not None:
            view.set("activeTab", "0")
            view.set("firstSheet", "0")

    calc = wb_root.find(f"{{{NS_MAIN}}}calcPr")
    if calc is None:
        calc = ET.SubElement(wb_root, f"{{{NS_MAIN}}}calcPr")
    calc.set("calcMode", "auto")
    calc.set("fullCalcOnLoad", "1")
    calc.set("forceFullCalc", "1")

    files["xl/workbook.xml"] = ET.tostring(wb_root, encoding="utf-8", xml_declaration=True)

    for _, sn in ordered_pairs:
        name = sn.attrib.get("name","")
        rid = sn.attrib.get(f"{{{NS_REL}}}id","")
        target = rid_to_target.get(rid)
        if not target or target not in files:
            continue

        root = parse_xml(files[target])

        sheet_pr = root.find(f"{{{NS_MAIN}}}sheetPr")
        if sheet_pr is None:
            sheet_pr = ET.Element(f"{{{NS_MAIN}}}sheetPr")
            root.insert(0, sheet_pr)

        tc = sheet_pr.find(f"{{{NS_MAIN}}}tabColor")
        if tc is None:
            tc = ET.SubElement(sheet_pr, f"{{{NS_MAIN}}}tabColor")
        tc.attrib.clear()
        tc.set("rgb", PALETTE[category(name)])

        sheet_views = root.find(f"{{{NS_MAIN}}}sheetViews")
        if sheet_views is not None:
            sv = sheet_views.find(f"{{{NS_MAIN}}}sheetView")
            if sv is not None:
                sv.set("showGridLines", "0")
                if "zoomScale" not in sv.attrib:
                    sv.set("zoomScale", "90")

        files[target] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    dest.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(dest, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)

def product_dirs(root: Path) -> List[Path]:
    return sorted([
        p for p in root.iterdir()
        if p.is_dir() and p.name not in IGNORE_DIRS and not p.name.startswith(".")
    ])

def excel_candidates(product_dir: Path) -> List[Path]:
    out = []
    for p in product_dir.rglob("*"):
        if not p.is_file() or p.name.startswith("~$"):
            continue
        if p.suffix.lower() not in (".xlsx",".xlsm",".xlsb",".xls"):
            continue
        rel_dirs = {x.lower() for x in p.relative_to(product_dir).parts[:-1]}
        if rel_dirs & {x.lower() for x in IGNORE_DIRS}:
            continue
        out.append(p)
    return sorted(out)

def looks_like_copy(name: str) -> bool:
    n = norm(name)
    return bool(re.search(r"(copy|kopya|_2$|_2_|_3$|_3_)", n))

def choose_canonical(files: List[Path], product_dir: Path) -> Tuple[Optional[Path], List[Path]]:
    if not files:
        return None, []
    def key(p: Path):
        return (
            1 if p.parent == product_dir else 0,
            0 if looks_like_copy(p.stem) else 1,
            p.stat().st_mtime,
            p.stat().st_size
        )
    canonical = sorted(files, key=key, reverse=True)[0]
    return canonical, [p for p in files if p != canonical]

def write_product_report(product_dir: Path, audit: Audit) -> None:
    qa = product_dir / "QA"
    qa.mkdir(exist_ok=True)
    (qa / "audit.json").write_text(
        json.dumps(asdict(audit), ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    trs = []
    for g in audit.gates:
        cls = "pass" if g.passed else "fail"
        trs.append(
            f"<tr><td>{html.escape(g.code)}</td><td>{html.escape(g.severity)}</td>"
            f"<td class='{cls}'>{'PASS' if g.passed else 'FAIL'}</td>"
            f"<td>{html.escape(g.message)}</td><td>{html.escape(g.evidence)}</td></tr>"
        )

    report = f"""<!doctype html>
<html lang="tr"><head><meta charset="utf-8">
<title>ExcelArsiv QA — {html.escape(audit.product)}</title>
<style>
body{{font-family:Arial,sans-serif;background:#f4f6f9;color:#172033;margin:32px}}
.card{{background:white;padding:22px;border-radius:14px;margin-bottom:18px}}
h1{{color:#0F2942}} .status{{font-size:22px;font-weight:700}}
table{{border-collapse:collapse;width:100%}}
th,td{{border-bottom:1px solid #d6dce4;padding:9px;text-align:left;vertical-align:top}}
th{{background:#0F2942;color:white}}
.pass{{color:#1f6b3b;font-weight:700}} .fail{{color:#9b2c2c;font-weight:700}}
</style></head><body>
<div class="card">
<h1>{html.escape(audit.product)}</h1>
<div class="status">Durum: {html.escape(audit.status)}</div>
<p>Dosya: {html.escape(audit.path)}</p>
<p>Sheet: {len(audit.sheets)} | Formül: {audit.formula_count} | Grafik: {audit.chart_count} |
Gerçek hyperlink: {audit.hyperlink_formula_count + audit.hyperlink_rel_count}</p>
<p>SHA-256: {html.escape(audit.sha256)}</p>
</div>
<div class="card"><table>
<thead><tr><th>Kod</th><th>Önem</th><th>Sonuç</th><th>Kontrol</th><th>Kanıt</th></tr></thead>
<tbody>{''.join(trs)}</tbody>
</table></div></body></html>"""
    (qa / "audit.html").write_text(report, encoding="utf-8")

def run(root: Path, apply: bool) -> int:
    system = root / "_SYSTEM"
    manifest_dir = system / "MANIFEST"
    reports_dir = system / "REPORTS"
    calibrated_dir = system / "CALIBRATED"
    quarantine = root / "_QUARANTINE"
    released = root / "_RELEASED"
    for d in (manifest_dir, reports_dir, calibrated_dir, quarantine, released):
        d.mkdir(parents=True, exist_ok=True)

    records = []
    all_files = []

    for pdir in product_dirs(root):
        files = excel_candidates(pdir)
        all_files.extend(files)
        canonical, extras = choose_canonical(files, pdir)

        if canonical is None:
            records.append({
                "product": pdir.name,
                "canonical": "",
                "status": "BLOCKED",
                "sha256": "",
                "extra_count": 0,
                "p1_fail": 1,
                "p2_fail": 0,
                "sheet_count": 0,
                "formula_count": 0,
                "chart_count": 0,
                "real_hyperlinks": 0,
                "reason": "Ana workbook bulunamadı"
            })
            continue

        audit = audit_workbook(canonical, pdir.name)

        if extras:
            audit.gates.append(Gate(
                "R010", "P2", False,
                "Tek kanonik workbook ilkesi",
                " | ".join(str(x) for x in extras)
            ))
            if audit.status == "READY_FOR_REVIEW":
                audit.status = "CALIBRATION_REQUIRED"

        if apply and audit.status != "BLOCKED" and canonical.suffix.lower() in (".xlsx",".xlsm"):
            dest = calibrated_dir / pdir.name / canonical.name
            calibrate_copy(canonical, dest)
            post = audit_workbook(dest, pdir.name)
            if extras:
                post.gates.append(Gate(
                    "R010", "P2", False,
                    "Tek kanonik workbook ilkesi",
                    " | ".join(str(x) for x in extras)
                ))
                if post.status == "READY_FOR_REVIEW":
                    post.status = "CALIBRATION_REQUIRED"
            post.notes.append(f"Kalibre kopya: {dest}")
            audit = post

        write_product_report(pdir, audit)

        records.append({
            "product": pdir.name,
            "canonical": audit.path,
            "status": audit.status,
            "sha256": audit.sha256,
            "extra_count": len(extras),
            "p1_fail": sum(1 for g in audit.gates if g.severity == "P1" and not g.passed),
            "p2_fail": sum(1 for g in audit.gates if g.severity == "P2" and not g.passed),
            "sheet_count": len(audit.sheets),
            "formula_count": audit.formula_count,
            "chart_count": audit.chart_count,
            "real_hyperlinks": audit.hyperlink_formula_count + audit.hyperlink_rel_count,
            "reason": ""
        })

    # Hash duplicate groups
    hash_groups: Dict[str, List[str]] = {}
    for p in all_files:
        try:
            hash_groups.setdefault(sha256(p), []).append(str(p))
        except Exception:
            pass
    duplicate_groups = {h: ps for h, ps in hash_groups.items() if len(ps) > 1}
    (manifest_dir / "duplicate_hashes.json").write_text(
        json.dumps(duplicate_groups, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    fields = [
        "product","canonical","status","sha256","extra_count","p1_fail","p2_fail",
        "sheet_count","formula_count","chart_count","real_hyperlinks","reason"
    ]
    with (manifest_dir / "quality_manifest.csv").open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in records:
            w.writerow(r)

    summary = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "root": str(root),
        "mode": "CALIBRATE_COPY" if apply else "AUDIT_ONLY",
        "products": len(records),
        "statuses": {},
        "duplicate_hash_groups": len(duplicate_groups),
        "quality_target": "500_USD_CLASS_PRODUCT_QUALITY",
        "auto_release": False,
        "human_checker_required": True
    }
    for r in records:
        summary["statuses"][r["status"]] = summary["statuses"].get(r["status"], 0) + 1

    (reports_dir / "quality_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    print("\n=== EXCELARSIV QUALITY FACTORY ===")
    print(f"Kök                : {root}")
    print(f"Ürün sayısı        : {summary['products']}")
    for status, count in sorted(summary["statuses"].items()):
        print(f"{status:<20}: {count}")
    print(f"Duplicate hash grup: {summary['duplicate_hash_groups']}")
    print(f"Manifest           : {manifest_dir / 'quality_manifest.csv'}")
    print(f"Özet               : {reports_dir / 'quality_summary.json'}")
    if apply:
        print(f"Kalibre kopyalar   : {calibrated_dir}")
    print("\nRELEASED otomatik verilmez. READY_FOR_REVIEW sonrası checker onayı zorunludur.")
    return 0

def main() -> int:
    ap = argparse.ArgumentParser(description="ExcelArsiv Quality Factory")
    ap.add_argument("--root", required=True, help="EXCELARSIV_PRODUCTS klasörü")
    ap.add_argument("--apply", action="store_true",
                    help="Kaynağa dokunmadan CALIBRATED kopyalara güvenli UI kalibrasyonu uygula")
    args = ap.parse_args()

    root = Path(args.root).expanduser().resolve()
    if not root.is_dir():
        print(f"HATA: klasör bulunamadı: {root}", file=sys.stderr)
        return 2
    return run(root, args.apply)

if __name__ == "__main__":
    raise SystemExit(main())
