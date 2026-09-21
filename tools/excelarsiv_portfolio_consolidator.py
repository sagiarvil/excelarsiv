#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
EXCELARSIV PORTFOLIO CONSOLIDATOR V1
====================================

Amaç:
- Aynı dosyanın/kopyanın birden fazla sürümünü tekilleştirmek.
- Aynı amaca hizmet etme olasılığı yüksek ürünleri içerik + yapı + isim üzerinden bulmak.
- En güçlü workbook'u "KEEP" olarak seçmek.
- Kaybedenleri önce geri döndürülebilir QUARANTINE alanına taşımak.
- İstenirse ikinci komutla karantinadaki loser dosyaları kalıcı silmek.

Fail-safe:
- --plan varsayılandır, hiçbir şey taşımaz/silmez.
- --commit yalnız HIGH_CONFIDENCE gruplarda loser dosyaları karantinaya taşır.
- --purge ayrı komuttur; kalıcı silme yapar.
- Aynı amaç konusunda düşük/orta güven varsa otomatik silme YOKTUR.
"""

from __future__ import annotations
import argparse, csv, hashlib, json, math, os, re, shutil, sys, time, zipfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT_DEFAULT = "/Users/macair1/projects/excel-arsiv/EXCELARSIV_PRODUCTS"
SYSTEM_DIR = "_SYSTEM"
QUARANTINE_DIR = "_QUARANTINE/CONSOLIDATION"

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

STOPWORDS = {
    "ve","ile","icin","mi","mu","miyim","hesabi","hesap","hesaplayici","sistemi","sistem",
    "motoru","motor","pro","paketi","dosya","dosyasi","robotu","robot","analizi","analiz",
    "simulatoru","simulasyonu","cetveli","kontrol","takip","komuta","merkezi","yonetim",
    "optimizasyonu","optimizasyon","gercek","net","toplu","yeni","2026"
}

# Aynı ailede olup yine de farklı kullanım amacı taşıyabilecek terimler.
# Bu kelimeler farklıysa otomatik birleştirme güveni düşürülür.
DISTINGUISHERS = {
    "satis","zamanlama","vergi","tasarruf","iade","mahsup","tevkifat","bordro","tesvik",
    "hakedis","mutabakat","fiyat","maliyet","nakit","odeme","vade","risk","karlilik",
    "yatirim","fizibilite","performans","uretim","depo","stok","kur","doviz","kredi",
    "covenant","tahsilat","isg","kvkk","sgk","ttk","konkordato","recete","menu"
}

def norm(s: str) -> str:
    tr = str.maketrans("ıİğĞüÜşŞöÖçÇ", "iIgGuUsSoOcC")
    s = s.translate(tr).lower()
    s = re.sub(r"\b(?:v\d+(?:\.\d+)*|final|son|copy|kopya)\b", " ", s)
    s = re.sub(r"\b[23]\b", " ", s)
    return re.sub(r"[^a-z0-9]+", " ", s).strip()

def tokens(s: str):
    return {t for t in norm(s).split() if len(t) > 2 and t not in STOPWORDS}

def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024*1024), b""):
            h.update(chunk)
    return h.hexdigest()

def jaccard(a, b):
    if not a and not b: return 1.0
    if not a or not b: return 0.0
    return len(a & b) / len(a | b)

def cosine_sets(a, b):
    # Binary-set cosine.
    if not a or not b: return 0.0
    return len(a & b) / math.sqrt(len(a) * len(b))

def xml_text_strings(path: Path, max_items=3000):
    """
    Workbook'tan yapısal/etiket fingerprint çıkarır.
    Formül sonuçlarına güvenmez; sheet adları, shared strings ve formül metinleri kullanılır.
    """
    strings = set()
    formulas = set()
    sheets = []
    formula_count = 0
    chart_count = 0
    hyperlinks = 0
    error_cells = 0
    external_links = 0

    if path.suffix.lower() not in (".xlsx",".xlsm"):
        return {
            "strings": strings, "formulas": formulas, "sheets": sheets,
            "formula_count": 0, "chart_count": 0, "hyperlinks": 0,
            "error_cells": 9999, "external_links": 9999
        }

    try:
        with zipfile.ZipFile(path, "r") as z:
            names = z.namelist()
            if z.testzip():
                raise ValueError("zip corruption")
            chart_count = len([n for n in names if n.startswith("xl/charts/chart") and n.endswith(".xml")])
            external_links = len([n for n in names if n.startswith("xl/externalLinks/") and n.endswith(".xml")])

            wb = ET.fromstring(z.read("xl/workbook.xml"))
            for sn in wb.findall(f".//{{{NS_MAIN}}}sheet"):
                sheets.append(sn.attrib.get("name",""))

            if "xl/sharedStrings.xml" in names:
                root = ET.fromstring(z.read("xl/sharedStrings.xml"))
                for si in root.findall(f"{{{NS_MAIN}}}si")[:max_items]:
                    txt = "".join(t.text or "" for t in si.iter(f"{{{NS_MAIN}}}t"))
                    n = norm(txt)
                    if len(n) >= 4:
                        strings.add(n[:180])

            for n in names:
                if not n.startswith("xl/worksheets/sheet") or not n.endswith(".xml"):
                    continue
                root = ET.fromstring(z.read(n))
                h = root.find(f"{{{NS_MAIN}}}hyperlinks")
                if h is not None:
                    hyperlinks += len(list(h))
                for c in root.iter(f"{{{NS_MAIN}}}c"):
                    if c.attrib.get("t") == "e":
                        error_cells += 1
                    f = c.find(f"{{{NS_MAIN}}}f")
                    if f is not None:
                        formula_count += 1
                        txt = norm(f.text or "")
                        if txt:
                            formulas.add(txt[:220])
    except Exception:
        error_cells = 9999
        external_links = 9999

    return {
        "strings": strings, "formulas": formulas, "sheets": [norm(x) for x in sheets],
        "formula_count": formula_count, "chart_count": chart_count, "hyperlinks": hyperlinks,
        "error_cells": error_cells, "external_links": external_links
    }

def load_quality_manifest(root: Path):
    path = root / SYSTEM_DIR / "MANIFEST" / "quality_manifest.csv"
    out = {}
    if not path.exists():
        return out
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            key = row.get("canonical","")
            if key:
                out[str(Path(key).resolve())] = row
    return out

def quality_score(path: Path, fp: dict, qrow: dict | None):
    """
    Skor yalnız KEEP seçmek için kullanılır; release kararı değildir.
    Kritik hata cezaları baskındır.
    """
    score = 0.0
    if qrow:
        status = qrow.get("status","")
        score += {"READY_FOR_REVIEW": 500, "CALIBRATION_REQUIRED": 250, "BLOCKED": 0}.get(status, 0)
        try: score -= int(qrow.get("p1_fail","0") or 0) * 250
        except: pass
        try: score -= int(qrow.get("p2_fail","0") or 0) * 40
        except: pass

    score -= min(fp["error_cells"], 50) * 200
    score -= min(fp["external_links"], 20) * 80

    score += min(fp["formula_count"], 1000) * 0.12
    score += min(fp["chart_count"], 12) * 8
    score += min(fp["hyperlinks"], 30) * 3
    score += min(len(fp["sheets"]), 20) * 2

    # Kopya isimli dosya cezası.
    if re.search(r"(?:\s|\(|_)(?:2|3|copy|kopya)(?:\)|_|\.|$)", path.name, re.I):
        score -= 30

    # Aynı kalite halinde daha yeni dosya hafif avantajlı.
    score += path.stat().st_mtime / 1e10
    return round(score, 3)

def purpose_similarity(a_path: Path, a_fp: dict, b_path: Path, b_fp: dict):
    name_a = tokens(a_path.parent.name + " " + a_path.stem)
    name_b = tokens(b_path.parent.name + " " + b_path.stem)
    name_j = jaccard(name_a, name_b)
    name_c = cosine_sets(name_a, name_b)

    strings_j = jaccard(a_fp["strings"], b_fp["strings"])
    formulas_j = jaccard(a_fp["formulas"], b_fp["formulas"])
    sheets_j = jaccard(set(a_fp["sheets"]), set(b_fp["sheets"]))

    # Domain distinguisher mismatch penalty:
    da = name_a & DISTINGUISHERS
    db = name_b & DISTINGUISHERS
    distinguish_penalty = 0.0
    if da and db and not (da & db):
        distinguish_penalty = 0.12

    # Same-purpose confidence; content is more important than name.
    score = (
        0.28 * name_j +
        0.12 * name_c +
        0.28 * strings_j +
        0.20 * formulas_j +
        0.12 * sheets_j
    ) - distinguish_penalty

    return round(max(0.0, min(1.0, score)), 4), {
        "name_jaccard": round(name_j,4),
        "strings_jaccard": round(strings_j,4),
        "formulas_jaccard": round(formulas_j,4),
        "sheets_jaccard": round(sheets_j,4),
        "distinguish_penalty": distinguish_penalty
    }

def list_workbooks(root: Path):
    out = []
    for product in root.iterdir():
        if not product.is_dir() or product.name.startswith("_") or product.name.startswith("."):
            continue
        for p in product.glob("*.xls*"):
            if p.name.startswith("~$") or p.suffix.lower() not in (".xlsx",".xlsm",".xlsb",".xls"):
                continue
            out.append(p)
    return sorted(out)

class DSU:
    def __init__(self, n):
        self.p = list(range(n))
    def find(self, x):
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]
            x = self.p[x]
        return x
    def union(self, a,b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb: self.p[rb] = ra

def build_plan(root: Path, threshold: float, exact_only: bool):
    files = list_workbooks(root)
    quality = load_quality_manifest(root)
    fps = {p: xml_text_strings(p) for p in files}
    hashes = {p: sha256(p) for p in files}

    dsu = DSU(len(files))
    evidence = {}

    # Exact duplicates always cluster.
    by_hash = defaultdict(list)
    for i,p in enumerate(files):
        by_hash[hashes[p]].append(i)
    for idxs in by_hash.values():
        if len(idxs) > 1:
            first = idxs[0]
            for j in idxs[1:]:
                dsu.union(first,j)
                evidence[(min(first,j),max(first,j))] = ("EXACT_HASH", 1.0, {})

    # Semantic/structural same-purpose.
    if not exact_only:
        for i in range(len(files)):
            for j in range(i+1, len(files)):
                if hashes[files[i]] == hashes[files[j]]:
                    continue
                score, parts = purpose_similarity(files[i], fps[files[i]], files[j], fps[files[j]])
                if score >= threshold:
                    dsu.union(i,j)
                    evidence[(i,j)] = ("HIGH_CONFIDENCE_PURPOSE", score, parts)

    groups = defaultdict(list)
    for i,p in enumerate(files):
        groups[dsu.find(i)].append(p)

    plan = []
    for members in groups.values():
        if len(members) == 1:
            continue
        scored = []
        for p in members:
            qrow = quality.get(str(p.resolve()))
            scored.append((quality_score(p, fps[p], qrow), p))
        scored.sort(key=lambda x: (x[0], x[1].stat().st_mtime), reverse=True)
        keep_score, keep = scored[0]
        losers = scored[1:]

        # Group confidence:
        all_same_hash = len({hashes[p] for p in members}) == 1
        confidence = "EXACT_DUPLICATE" if all_same_hash else "HIGH_CONFIDENCE_PURPOSE"
        plan.append({
            "confidence": confidence,
            "keep": str(keep),
            "keep_score": keep_score,
            "losers": [{"path":str(p),"score":s} for s,p in losers],
            "members": [str(p) for p in members]
        })
    return plan

def write_plan(root: Path, plan):
    report_dir = root / SYSTEM_DIR / "CONSOLIDATION"
    report_dir.mkdir(parents=True, exist_ok=True)
    json_path = report_dir / "consolidation_plan.json"
    csv_path = report_dir / "consolidation_plan.csv"
    json_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["confidence","keep","keep_score","loser","loser_score"])
        for g in plan:
            for loser in g["losers"]:
                w.writerow([g["confidence"],g["keep"],g["keep_score"],loser["path"],loser["score"]])
    return json_path, csv_path

def commit(root: Path, plan):
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    qroot = root / QUARANTINE_DIR / stamp
    qroot.mkdir(parents=True, exist_ok=True)
    moved = []
    for g in plan:
        # Only exact or high-confidence groups produced by threshold logic are eligible.
        for loser in g["losers"]:
            src = Path(loser["path"])
            if not src.exists():
                continue
            rel = src.relative_to(root)
            dst = qroot / rel
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(src), str(dst))
            moved.append({
                "source": str(src),
                "quarantine": str(dst),
                "keep": g["keep"],
                "confidence": g["confidence"]
            })

    manifest = qroot / "moved_manifest.json"
    manifest.write_text(json.dumps(moved, ensure_ascii=False, indent=2), encoding="utf-8")
    return qroot, moved

def purge(root: Path, stamp: str):
    qroot = root / QUARANTINE_DIR / stamp
    if not qroot.exists():
        raise SystemExit(f"Karantina bulunamadı: {qroot}")
    manifest = qroot / "moved_manifest.json"
    if not manifest.exists():
        raise SystemExit("moved_manifest.json yok; güvenlik nedeniyle purge durduruldu.")
    moved = json.loads(manifest.read_text(encoding="utf-8"))
    count = len(moved)
    shutil.rmtree(qroot)
    return count

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default=ROOT_DEFAULT)
    ap.add_argument("--threshold", type=float, default=0.88,
                    help="Aynı amaç otomatik gruplama eşiği. Varsayılan 0.88 (yüksek güven).")
    ap.add_argument("--exact-only", action="store_true",
                    help="Yalnız SHA-256 birebir kopyaları grupla.")
    ap.add_argument("--commit", action="store_true",
                    help="Planlanan loser dosyaları karantinaya taşı.")
    ap.add_argument("--purge", metavar="STAMP",
                    help="Belirli karantina batch'ini kalıcı sil.")
    args = ap.parse_args()

    root = Path(args.root).expanduser().resolve()
    if not root.is_dir():
        raise SystemExit(f"Kök klasör yok: {root}")

    if args.purge:
        n = purge(root, args.purge)
        print(f"KALICI SİLME TAMAMLANDI. Silinen loser dosya: {n}")
        return

    plan = build_plan(root, args.threshold, args.exact_only)
    json_path, csv_path = write_plan(root, plan)

    loser_count = sum(len(g["losers"]) for g in plan)
    exact_groups = sum(1 for g in plan if g["confidence"] == "EXACT_DUPLICATE")
    purpose_groups = sum(1 for g in plan if g["confidence"] == "HIGH_CONFIDENCE_PURPOSE")

    print("=== EXCELARSIV PORTFOLIO CONSOLIDATOR ===")
    print(f"Plan grubu             : {len(plan)}")
    print(f"Exact duplicate grup   : {exact_groups}")
    print(f"Aynı amaç yüksek güven : {purpose_groups}")
    print(f"Elenecek dosya adayı   : {loser_count}")
    print(f"JSON plan              : {json_path}")
    print(f"CSV plan               : {csv_path}")

    if not args.commit:
        print("\nPLAN MODU: hiçbir dosya taşınmadı/silinmedi.")
        print("Onay sonrası aynı komuta --commit ekleyin.")
        return

    qroot, moved = commit(root, plan)
    print(f"\nKARANTİNAYA TAŞINDI: {len(moved)}")
    print(f"Karantina batch: {qroot.name}")
    print(f"Konum: {qroot}")
    print("\nKalıcı silmek için:")
    print(f'python3 "{Path(__file__).resolve()}" --root "{root}" --purge "{qroot.name}"')

if __name__ == "__main__":
    main()
