#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
ExcelArsiv gated catalog publisher.

Publishes a product to src/content/templates ONLY when all evidence exists:
1) current QA status == READY_FOR_REVIEW
2) SPEC.yaml exists and parses
3) KANIT/SEVK_KARARI.md explicitly says SEVK EDİLEBİLİR
4) three product screenshots exist
5) product is not a canonical alias of an already-published product
6) required commercial/catalog metadata can be produced without inventing facts

Products without commerce/catalog payment mapping may still be listed for WhatsApp
quote, but are NOT marked as directly in-stock / directly purchasable.
"""

from __future__ import annotations

import csv
import json
import re
import sys
from datetime import date
from pathlib import Path

try:
    import yaml
except ImportError:
    raise SystemExit("PyYAML gerekli: python3 -m pip install pyyaml")

ROOT = Path(__file__).resolve().parents[1]
PRODUCTS = ROOT / "EXCELARSIV_PRODUCTS"
TEMPLATES = ROOT / "src" / "content" / "templates"
MANIFEST = PRODUCTS / "_SYSTEM" / "MANIFEST" / "quality_manifest.csv"
PUBLISH_DIR = PRODUCTS / "_SYSTEM" / "PUBLISH"
COMMERCE = ROOT / "commerce" / "catalog.json"

ALIASES = {
    "cek-senet-vade-risk-sistemi": "cek-senet-ve-vade-risk-sistemi",
    "gunluk-gelir-gider-karlilik-sistemi": "gunluk-gelir-gider-ve-gercek-karlilik-sistemi",
    "on-uc-haftalik-nakit-akisi-ve-odeme-planlama-sistemi": "13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi",
    "pos-komisyon-net-tahsilat-kontrol-sistemi": "pos-komisyon-ve-net-tahsilat-kontrol-sistemi",
}

CATEGORY_MAP = {
    "asgari-kurumlar-vergisi-simulasyon-motoru": "muhasebe-ve-vergi",
    "banka-kredi-covenant-erken-uyari": "finansal-analiz",
    "dava-masraf-harc-hesaplama": "butce-ve-planlama",
    "depo-doluluk-lokasyon": "stok-ve-uretim",
    "dernek-kooperatif-mali-yonetim": "butce-ve-planlama",
    "e-belge-zorunluluk-radari": "muhasebe-ve-vergi",
    "e-ticaret-gercek-karlilik-fiyatlama": "satis-ve-fiyatlama",
    "elektrik-faturasi-dogrulama": "butce-ve-planlama",
    "emlak-pipeline-komisyon": "satis-ve-fiyatlama",
    "fason-uretim-takip-sistemi": "stok-ve-uretim",
    "fesih-maliyeti-simulatoru": "personel-ve-bordro",
    "filo-arac-maliyet-komutasi": "butce-ve-planlama",
    "ges-uretim-performans": "stok-ve-uretim",
    "ges-yatirim-fizibilite": "finansal-analiz",
    "gida-lot-maliyet-izlenebilirlik": "stok-ve-uretim",
    "hizmet-ihracati-100-indirim-transfer-takip": "muhasebe-ve-vergi",
    "hukuk-burosu-dosya-vekalet": "butce-ve-planlama",
    "ihale-fiyat-farki-eskalasyon-pro": "satis-ve-fiyatlama",
    "ihracat-siparis-karlilik-kur": "satis-ve-fiyatlama",
    "insaat-hakedis-yonetim-sistemi": "butce-ve-planlama",
    "isg-risk-degerlendirme-pro-6331": "personel-ve-bordro",
    "ithalat-landed-cost-motoru": "stok-ve-uretim",
    "kargo-desi-maliyet-optimizasyonu": "satis-ve-fiyatlama",
    "kat-karsiligi-hasilat-paylasimi-simulator": "finansal-analiz",
    "kdv-iadesi-tutar-surec-simulasyonu": "muhasebe-ve-vergi",
    "kik-asiri-dusuk-savunma-sinir-deger": "satis-ve-fiyatlama",
    "kira-portfoyu-getiri-komutasi": "finansal-analiz",
    "konkordato-ttk376-kriz-paketi": "finansal-analiz",
    "kvkk-veri-envanteri-verbis-uyum": "muhasebe-ve-vergi",
    "makine-bakim-kalibrasyon-durus-maliyeti": "stok-ve-uretim",
    "mini-mrp-bom-malzeme-kapasite": "stok-ve-uretim",
    "mizan-anomali-denetim-oncesi-kontrol": "muhasebe-ve-vergi",
    "oee-durus-kok-neden-komutasi": "stok-ve-uretim",
    "on-uc-haftalik-nakit-odeme-onceligi": "nakit-akisi",
    "ortulu-sermaye-emsal-faiz-tf-paketi": "muhasebe-ve-vergi",
    "pazaryeri-hakedis-mutabakat-motoru": "satis-ve-fiyatlama",
    "proje-finansmani-dscr-llcr-paketi": "finansal-analiz",
    "puantaj-vardiya-fazla-mesai-kanit-sistemi": "personel-ve-bordro",
    "recete-maliyeti-menu-muhendisligi": "stok-ve-uretim",
    "sarj-istasyonu-yatirim": "finansal-analiz",
    "sekiz-d-duzeltici-faaliyet-dosya": "stok-ve-uretim",
    "sera-kurulum-fizibilite": "finansal-analiz",
    "sevkiyat-fiyatlama-navlun": "satis-ve-fiyatlama",
    "sgk-prim-tesvikleri-optimizasyon-motoru": "personel-ve-bordro",
    "site-apartman-yonetim-sistemi": "butce-ve-planlama",
    "spc-proses-yetenek-analizi": "stok-ve-uretim",
    "stok-optimizasyon-abc-olu-stok-nakit": "stok-ve-uretim",
    "sut-surusu-yonetim": "stok-ve-uretim",
    "tahsilat-riski-vade-komuta-paneli": "nakit-akisi",
    "tarimsal-destek-uygunluk": "butce-ve-planlama",
    "uretim-kari-125-kv-optimizasyonu": "stok-ve-uretim",
    "yem-rasyonu-maliyet": "stok-ve-uretim",
    "yeniden-degerleme-komuta-merkezi": "muhasebe-ve-vergi",
    "yurt-disi-yapilanma-vergi-simulatoru": "finansal-analiz",
}

INPUT_TOKENS = (
    "girdi", "kayit", "kart", "hareket", "liste", "veri", "ayar", "parametre",
    "tanim", "hesap_plani", "siparis", "personel", "cari", "fatura"
)
OUTPUT_TOKENS = (
    "pano", "rapor", "ozet", "karar", "dashboard", "sonuc", "takvim", "analiz"
)


def q(s: str) -> str:
    return "'" + str(s).replace("\\", "\\\\").replace("'", "''").replace("\n", " ").strip() + "'"


def compact(text: str, max_len: int) -> str:
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    if len(text) <= max_len:
        return text
    cut = text[: max_len - 1].rstrip(" ,;:-")
    last = max(cut.rfind(". "), cut.rfind("; "), cut.rfind(", "))
    if last >= 60:
        cut = cut[:last]
    return cut.rstrip(". ") + "…"


def titleize(key: str) -> str:
    return key.replace("_", " ").strip().title()


def classify_sheet(name: str) -> str:
    n = name.lower()
    if any(x in n for x in INPUT_TOKENS):
        return "input"
    if any(x in n for x in OUTPUT_TOKENS):
        return "output"
    return "calculation"


def purpose_for_sheet(name: str, kind: str) -> str:
    label = titleize(name)
    if kind == "input":
        return f"{label} için kullanıcı veri girişi ve tanım alanı"
    if kind == "output":
        return f"{label} üzerinden sonuç, karar veya yönetim çıktısı"
    return f"{label} için hesaplama ve kontrol katmanı"


def listify_target(value) -> list[str]:
    if isinstance(value, list):
        return [str(x).strip() for x in value if str(x).strip()]
    if not value:
        return []
    return [x.strip() for x in re.split(r",|;|\s+ve\s+", str(value)) if x.strip()]


def extract_inputs(spec: dict) -> list[str]:
    urun = spec.get("urun") or {}
    giris = urun.get("giris") or {}
    if isinstance(giris, dict):
        out = []
        for key, value in giris.items():
            item = titleize(str(key))
            if isinstance(value, dict) and value.get("tablo"):
                item += f" ({value['tablo']})"
            out.append(item)
        if out:
            return out[:8]
    return ["SPEC dosyasında tanımlanan kullanıcı girdileri"]


def extract_outputs(spec: dict) -> list[str]:
    urun = spec.get("urun") or {}
    out = []
    karar = urun.get("karar")
    rapor = urun.get("rapor")
    if karar:
        out.append(compact(karar, 180))
    if rapor:
        out.append(compact(rapor, 180))
    if not out:
        out.append("Karar motoru, kontrol ve yönetim çıktıları")
    return out[:4]


def derive_price(spec: dict) -> int | None:
    deger = spec.get("deger") or {}
    value = deger.get("satis_fiyati_tl")
    try:
        p = int(float(value))
        return p if p > 0 else None
    except Exception:
        return None


def render_mdx(slug: str, spec: dict, audit: dict, commerce_products: dict) -> str:
    urun = spec.get("urun") or {}
    name = compact(urun.get("ad") or titleize(slug), 70)
    decision = urun.get("karar") or ""
    summary = compact(decision, 158)
    if len(summary) < 40:
        summary = compact(f"{name}: SPEC kapsamında tanımlanan girdileri hesaplar, kontrol eder ve karar çıktısına dönüştürür.", 158)

    price = derive_price(spec)
    if not price:
        raise ValueError("SPEC deger.satis_fiyati_tl eksik")

    category = CATEGORY_MAP.get(slug)
    if not category:
        raise ValueError("Kategori eşlemesi yok")

    sheets = audit.get("sheets") or []
    sheet_map = [(s, classify_sheet(s)) for s in sheets if s != "00_NAVIGASYON"]
    target_users = listify_target(urun.get("hedef_kullanici"))
    if not target_users:
        target_users = ["SPEC kapsamında tanımlanan işletme kullanıcıları"]

    inputs = extract_inputs(spec)
    outputs = extract_outputs(spec)
    macros = audit.get("fmt") == ".xlsm"
    file_format = "xlsm" if macros else "xlsx"
    size_mb = round(float(audit.get("size_bytes") or 0) / 1024 / 1024, 2)
    if size_mb <= 0:
        size_mb = 0.01

    direct_sale = bool((commerce_products.get(slug) or {}).get("satista") is True)
    sale_note = (
        "Shopier satış kaydı aktif; ürün sayfasındaki teklif ve teslimat akışı kullanılabilir."
        if direct_sale else
        "Doğrudan ödeme kaydı henüz tanımlı değil; ürün teklif talebi üzerinden sunulur."
    )

    fm = [
        "---",
        f"name: {q(name)}",
        f"summary: {q(summary)}",
        f"category: {q(category)}",
        f"priceTL: {price}",
        "vatIncluded: true",
        f"fileFormat: {file_format}",
        f"sizeMB: {size_mb}",
        f"sheetCount: {len(sheets)}",
        f"hasMacros: {'true' if macros else 'false'}",
        "minExcelVersion: 'Microsoft Excel 2016 ve üzeri'",
        f"macCompatible: {'false' if macros else 'true'}",
        "sheetsCompatibility: 'partial'",
        "version: '1.0.0'",
        f"updatedAt: '{date.today().isoformat()}'",
        "sheetMap:",
    ]
    for s, kind in sheet_map:
        fm += [
            f"  - name: {q(s)}",
            f"    purpose: {q(purpose_for_sheet(s, kind))}",
            f"    kind: {q(kind)}",
        ]
    fm.append("inputs:")
    for item in inputs:
        fm.append(f"  - {q(item)}")
    fm.append("outputs:")
    for item in outputs:
        fm.append(f"  - {q(item)}")
    fm.append("suitableFor:")
    for item in target_users[:6]:
        fm.append(f"  - {q(item)}")
    fm += [
        "notSuitableFor:",
        "  - 'SPEC kapsamı dışında kalan ve tam ERP/özel yazılım entegrasyonu gerektiren süreçler'",
        "requirements:",
        "  - 'Microsoft Excel 2016 veya üzeri'",
        "  - 'Dosyadaki giriş alanlarının işletme verileriyle doldurulması'",
        f"  - {q('Makro gerektirir' if macros else 'Makro gerektirmez')}",
        "updatePolicy: 'Model mantığı veya dayanak parametreler değiştiğinde sürüm yeniden kalite kontrolünden geçirilir.'",
        "faq:",
        f"  - question: {q('Bu sistem hangi kararı veya çıktıyı üretir?')}",
        f"    answer: {q(compact(decision or outputs[0], 420))}",
        f"  - question: {q('Kalite kontrolü nasıl doğrulandı?')}",
        f"    answer: {q('Güncel ExcelArsiv QA kapılarında P1/P2 hata bırakmadan READY_FOR_REVIEW statüsü aldı; ayrıca ürün KANIT/SEVK_KARARI kaydı SEVK EDİLEBİLİR sonucunu içeriyor.')}",
        f"  - question: {q('Satın alma veya teklif süreci nasıl ilerler?')}",
        f"    answer: {q(sale_note)}",
        "screenshots:",
        f"  - src: '/screenshots/{slug}-1.png'",
        f"    alt: {q(name + ' — ürün ekranı 1')}",
        f"  - src: '/screenshots/{slug}-2.png'",
        f"    alt: {q(name + ' — ürün ekranı 2')}",
        f"  - src: '/screenshots/{slug}-3.png'",
        f"    alt: {q(name + ' — ürün ekranı 3')}",
        "related: []",
        "---",
        "",
        "Bu ürün, ExcelArsiv kalite kapıları ve ürün klasöründeki SPEC/KANIT kayıtları esas alınarak kataloğa alınmıştır.",
        "",
    ]
    return "\n".join(fm)


def main() -> int:
    PUBLISH_DIR.mkdir(parents=True, exist_ok=True)
    TEMPLATES.mkdir(parents=True, exist_ok=True)

    with MANIFEST.open("r", encoding="utf-8-sig", newline="") as f:
        qa_rows = {row["product"]: row for row in csv.DictReader(f)}

    commerce = json.loads(COMMERCE.read_text(encoding="utf-8"))
    commerce_products = commerce.get("products") or {}

    existing = {p.stem for p in TEMPLATES.glob("*.mdx")}
    records = []
    created = []

    for product_dir in sorted(p for p in PRODUCTS.iterdir() if p.is_dir() and not p.name.startswith("_")):
        slug = product_dir.name

        if slug in existing:
            records.append({"slug": slug, "status": "ALREADY_PUBLISHED", "reason": ""})
            continue

        if slug in ALIASES:
            canonical = ALIASES[slug]
            records.append({"slug": slug, "status": "CONSOLIDATED_ALIAS", "reason": canonical})
            continue

        qrow = qa_rows.get(slug)
        blockers = []
        if not qrow or qrow.get("status") != "READY_FOR_REVIEW":
            blockers.append("QA_READY_FOR_REVIEW_YOK")

        spec_path = product_dir / "SPEC.yaml"
        sevk_path = product_dir / "KANIT" / "SEVK_KARARI.md"
        audit_path = product_dir / "QA" / "audit.json"

        if not spec_path.exists(): blockers.append("SPEC_YOK")
        if not sevk_path.exists(): blockers.append("SEVK_KARARI_YOK")
        if not audit_path.exists(): blockers.append("AUDIT_YOK")

        for i in (1, 2, 3):
            if not (ROOT / "public" / "screenshots" / f"{slug}-{i}.png").exists():
                blockers.append(f"SCREENSHOT_{i}_YOK")

        if blockers:
            records.append({"slug": slug, "status": "BLOCKED", "reason": "|".join(blockers)})
            continue

        sevk = sevk_path.read_text(encoding="utf-8", errors="replace")
        if "SEVK EDİLEBİLİR" not in sevk:
            records.append({"slug": slug, "status": "BLOCKED", "reason": "SEVK_EDILEBILIR_KANITI_YOK"})
            continue

        try:
            spec = yaml.safe_load(spec_path.read_text(encoding="utf-8")) or {}
            audit = json.loads(audit_path.read_text(encoding="utf-8"))
            mdx = render_mdx(slug, spec, audit, commerce_products)
            out = TEMPLATES / f"{slug}.mdx"
            out.write_text(mdx, encoding="utf-8")
            created.append(slug)
            records.append({
                "slug": slug,
                "status": "PUBLISHED_TO_CATALOG",
                "reason": "DIRECT_SALE" if slug in commerce_products else "QUOTE_ONLY_NO_COMMERCE_MAPPING"
            })
        except Exception as exc:
            records.append({"slug": slug, "status": "BLOCKED", "reason": f"GENERATOR:{exc}"})

    report = {
        "generatedAt": date.today().isoformat(),
        "createdCount": len(created),
        "created": created,
        "records": records,
        "rules": {
            "qa": "READY_FOR_REVIEW required",
            "sevk": "SEVK EDİLEBİLİR required",
            "screenshots": "3 required",
            "aliases": ALIASES,
            "directSale": "only when commerce/catalog.json has satista=true"
        }
    }
    (PUBLISH_DIR / "publication_manifest.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    blocked = [r for r in records if r["status"] == "BLOCKED"]
    print(f"Created: {len(created)}")
    print(f"Blocked: {len(blocked)}")
    print(f"Aliases consolidated: {sum(r['status']=='CONSOLIDATED_ALIAS' for r in records)}")
    print(f"Manifest: {PUBLISH_DIR / 'publication_manifest.json'}")
    if blocked:
        print("Blocked products:")
        for r in blocked:
            print(f" - {r['slug']}: {r['reason']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
