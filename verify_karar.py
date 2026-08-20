#!/usr/bin/env python3
"""
verify_karar.py — MANDATE-EXCELARSIV-KARAR-V1 makine dogrulayicisi

Kullanim:
    python verify_karar.py --check all --base https://excelarsiv.com
    python verify_karar.py --check content --base http://localhost:4321
    python verify_karar.py --check similarity --base http://localhost:4321

Cikis kodu: 0 = YESIL, 1 = KIRMIZI, 2 = calistirma hatasi
Bagimlilik: pip install requests beautifulsoup4
"""

import argparse
import json
import os
import re
import sys
from itertools import combinations

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    print("HATA: 'pip install requests beautifulsoup4' gerekli.", file=sys.stderr)
    sys.exit(2)


SLUGLAR = [
    "hangi-excel-sistemini-almaliyim",
    "kobi-nakit-akisi-excel",
    "kasa-defteri-excel",
    "mali-musavir-cari-takip-excel",
    "pos-komisyon-kontrol-excel",
    "trendyol-pazaryeri-net-kar-excel",
    "kdv-iade-dosyasi-excel",
    "amortisman-yeniden-degerleme-excel",
    "kidem-ihbar-maliyeti-excel",
    "sgk-tesvik-optimizasyon-excel",
    "restoran-kafe-maliyet-excel",
    "insaat-hakedis-excel",
    "ihale-teklif-sinir-deger-excel",
    "stok-devir-nakit-baglanma-excel",
    "sube-karlilik-analizi-excel",
    "ttk-376-sermaye-kaybi-excel",
    "doviz-acik-pozisyon-kur-riski-excel",
    "logo-erp-cari-yaslandirma-excel",
]

CEVAP_MIN, CEVAP_MAKS = 40, 60          # INV-K3.2
KELIME_MIN, KELIME_MAKS = 700, 1200     # INV-K3.3
BENZERLIK_UST = 0.20                    # INV-K3.5
IC_LINK_MIN, IC_LINK_MAKS = 3, 12       # INV-K3.10
SSS_ADET = 5                            # INV-K4.1
SINIR_BASLIK = re.compile(r"almayın eğer|almayin eger|bu sistem size uygun değil", re.I)
YASAK_IDDIA = ["en iyi", "türkiye'nin 1", "binlerce", "lider", "%100 memnuniyet", "milyonlarca"]

sonuclar = []


def kayit(kod, gecti, mesaj):
    sonuclar.append((kod, gecti, mesaj))
    print(f"[{'YESIL' if gecti else 'KIRMIZI'}] {kod}: {mesaj}")
    return gecti


def getir(url):
    r = requests.get(url, timeout=30, headers={"User-Agent": "excelarsiv-karar-verifier/1.0"})
    r.raise_for_status()
    r.encoding = "utf-8"
    return BeautifulSoup(r.text, "html.parser"), r.text


def kelimeler(metin):
    return re.findall(r"\w+", metin.lower(), flags=re.UNICODE)


def shingles(metin, n=5):
    k = kelimeler(metin)
    return {tuple(k[i:i + n]) for i in range(max(0, len(k) - n + 1))}


def jsonld_bloklari(soup):
    bloklar = []
    for s in soup.find_all("script", attrs={"type": "application/ld+json"}):
        try:
            veri = json.loads(s.string or "{}")
        except Exception:
            continue
        bloklar.extend(veri if isinstance(veri, list) else [veri])
    return bloklar


def sema_turleri(bloklar):
    turler = set()
    for b in bloklar:
        t = b.get("@type")
        if isinstance(t, list):
            turler.update(t)
        elif t:
            turler.add(t)
        for g in b.get("@graph", []) or []:
            gt = g.get("@type")
            if isinstance(gt, list):
                turler.update(gt)
            elif gt:
                turler.add(gt)
    return turler


def sayfa_denetle(base, slug, rapor):
    url = f"{base}/karar/{slug}"
    try:
        soup, ham = getir(url)
    except requests.HTTPError as e:
        kayit(f"SAYFA:{slug}", False, f"Erisilemedi: {e}")
        return False, ""

    gecer = True
    govde = soup.find("main") or soup.body
    metin = govde.get_text(" ", strip=True) if govde else ""
    kelime_sayisi = len(kelimeler(metin))

    # INV-K3.1 tek h1 + sorgu ifadesi
    h1ler = soup.find_all("h1")
    gecer &= kayit(f"INV-K3.1:{slug}", len(h1ler) == 1, f"<h1> sayisi {len(h1ler)}")

    # INV-K3.2 cevap blogu
    paragraflar = [p.get_text(" ", strip=True) for p in (govde.find_all("p") if govde else [])]
    cevap = next((p for p in paragraflar if CEVAP_MIN <= len(kelimeler(p)) <= CEVAP_MAKS), None)
    ilk200 = " ".join(kelimeler(metin)[:200])
    cevap_erken = bool(cevap) and " ".join(kelimeler(cevap)[:6]) in ilk200
    gecer &= kayit(f"INV-K3.2:{slug}", bool(cevap) and cevap_erken,
                   f"Cevap blogu {'bulundu' if cevap else 'YOK'}"
                   + (f" ({len(kelimeler(cevap))} kelime)" if cevap else "")
                   + ("" if cevap_erken else " — ilk 200 kelime icinde degil"))

    # INV-K3.3 uzunluk
    gecer &= kayit(f"INV-K3.3:{slug}", KELIME_MIN <= kelime_sayisi <= KELIME_MAKS,
                   f"{kelime_sayisi} kelime (hedef {KELIME_MIN}-{KELIME_MAKS})")

    # INV-K3.6 dogru sinir blogu
    gecer &= kayit(f"INV-K3.6:{slug}", bool(SINIR_BASLIK.search(metin)),
                   "'almayin eger' blogu " + ("var" if SINIR_BASLIK.search(metin) else "YOK"))

    # INV-K3.8 gercek tablo
    tablolar = soup.find_all("table")
    gecer &= kayit(f"INV-K3.8:{slug}", len(tablolar) >= 1, f"<table> sayisi {len(tablolar)}")

    # INV-K3.9 yasak iddia
    dusuk = metin.lower()
    bulunan = [y for y in YASAK_IDDIA if y in dusuk]
    gecer &= kayit(f"INV-K3.9:{slug}", not bulunan, f"Yasak iddia: {bulunan or 'yok'}")

    # INV-K3.10 ic link araligi
    ic = {a.get("href", "") for a in (govde.find_all("a") if govde else [])
          if a.get("href", "").startswith("/")}
    gecer &= kayit(f"INV-K3.10:{slug}", IC_LINK_MIN <= len(ic) <= IC_LINK_MAKS,
                   f"Ic link {len(ic)} (hedef {IC_LINK_MIN}-{IC_LINK_MAKS})")

    # INV-K4.1/2/3 semalar
    turler = sema_turleri(jsonld_bloklari(soup))
    for kod, tur in (("INV-K4.1", "FAQPage"), ("INV-K4.2", "ItemList"), ("INV-K4.3", "BreadcrumbList")):
        gecer &= kayit(f"{kod}:{slug}", tur in turler, f"{tur} {'var' if tur in turler else 'YOK'}")

    # INV-K4.5 canonical + index
    canon = soup.find("link", rel=lambda v: v and "canonical" in v)
    robots = (soup.find("meta", attrs={"name": "robots"}) or {}).get("content", "")
    gecer &= kayit(f"INV-K4.5:{slug}",
                   bool(canon) and slug in (canon.get("href", "") if canon else "")
                   and "noindex" not in robots.lower(),
                   f"canonical={'ok' if canon else 'YOK'} robots='{robots}'")

    # INV-K5.4 title / description
    title = (soup.title.string or "").strip() if soup.title else ""
    desc = (soup.find("meta", attrs={"name": "description"}) or {}).get("content", "")
    gecer &= kayit(f"INV-K5.4:{slug}", len(title) <= 60 and 140 <= len(desc) <= 160,
                   f"title {len(title)} kr, description {len(desc)} kr")

    # INV-K5.5 guncelleme tarihi
    gecer &= kayit(f"INV-K5.5:{slug}", bool(re.search(r"20\d{2}-\d{2}-\d{2}|Son güncelleme", ham)),
                   "Guncelleme tarihi " + ("var" if re.search(r"Son güncelleme", ham) else "YOK"))

    rapor[slug] = {"kelime": kelime_sayisi, "cevap_kelime": len(kelimeler(cevap)) if cevap else 0,
                   "ic_link": len(ic), "semalar": sorted(turler)}
    return gecer, metin


def check_similarity(metinler):
    """INV-K3.5 — sayfalar arasi benzerlik < %20."""
    ihlal = []
    for (a, ma), (b, mb) in combinations(metinler.items(), 2):
        sa, sb = shingles(ma), shingles(mb)
        if not sa or not sb:
            continue
        skor = len(sa & sb) / min(len(sa), len(sb))
        if skor >= BENZERLIK_UST:
            ihlal.append(f"{a} ↔ {b}: {skor:.0%}")
    return kayit("INV-K3.5", not ihlal,
                 f"Benzerlik ihlali: {len(ihlal)}" +
                 ("" if not ihlal else "\n    " + "\n    ".join(ihlal[:15])))


def check_geo(base):
    gecer = True
    try:
        r = requests.get(base + "/llms.txt", timeout=20)
        icerik = r.text if r.status_code == 200 else ""
    except requests.RequestException:
        icerik = ""
    eksik = [s for s in SLUGLAR if s not in icerik]
    gecer &= kayit("INV-K5.1", bool(icerik) and not eksik,
                   f"/llms.txt {'yok' if not icerik else f'var, eksik slug: {len(eksik)}'}")

    try:
        rb = requests.get(base + "/robots.txt", timeout=20).text
    except requests.RequestException:
        rb = ""
    botlar = ["GPTBot", "OAI-SearchBot", "PerplexityBot", "Google-Extended", "ClaudeBot"]
    tanimli = [b for b in botlar if b.lower() in rb.lower()]
    gecer &= kayit("INV-K5.2", len(tanimli) == len(botlar),
                   f"robots.txt'te tanimli AI botu: {tanimli or 'yok'} "
                   f"(karar dosyasi: kanit/GATE-K-robots-karari.md ayrica kontrol edilmeli)")
    return gecer


def check_index(base):
    soup, _ = getir(base + "/karar")
    hrefler = {a.get("href", "").rstrip("/") for a in soup.find_all("a")}
    eksik = [s for s in SLUGLAR if f"/karar/{s}" not in hrefler]
    return kayit("INV-K6.5", not eksik, f"/karar dizininde eksik sayfa: {eksik or 'yok'}")


def check_homepage_links(base):
    soup, _ = getir(base + "/")
    hrefler = {a.get("href", "").rstrip("/") for a in soup.find_all("a")}
    karar_linkleri = [h for h in hrefler if h.startswith("/karar")]
    g1 = kayit("INV-K6.2", len([h for h in karar_linkleri if h != "/karar"]) >= 8,
               f"Ana sayfada karar sayfasi linki: {len([h for h in karar_linkleri if h != '/karar'])} (min 8)")
    bolum = len(soup.find_all("section"))
    g2 = kayit("INV-K6.1", bolum <= 7, f"Ana sayfa <section> sayisi {bolum} (ust sinir 7)")
    return g1 and g2


def main():
    p = argparse.ArgumentParser(description="MANDATE-EXCELARSIV-KARAR-V1 dogrulayici")
    p.add_argument("--check", default="all",
                   choices=["all", "content", "similarity", "geo", "index", "homepage"])
    p.add_argument("--base", default="https://excelarsiv.com")
    a = p.parse_args()
    base = a.base.rstrip("/")
    os.makedirs("kanit", exist_ok=True)

    rapor, metinler = {}, {}
    try:
        if a.check in ("all", "content", "similarity"):
            for slug in SLUGLAR:
                _, metin = sayfa_denetle(base, slug, rapor)
                if metin:
                    metinler[slug] = metin
        if a.check in ("all", "similarity"):
            check_similarity(metinler)
        if a.check in ("all", "geo"):
            check_geo(base)
        if a.check in ("all", "index"):
            check_index(base)
        if a.check in ("all", "homepage"):
            check_homepage_links(base)
    except requests.RequestException as e:
        print(f"HATA (ag): {e}", file=sys.stderr)
        sys.exit(2)

    with open("kanit/GATE-K1-icerik-denetim.json", "w", encoding="utf-8") as f:
        json.dump(rapor, f, ensure_ascii=False, indent=2)

    kirmizi = [k for k, gecti, _ in sonuclar if not gecti]
    print("\n" + "=" * 60)
    print(f"TOPLAM: {len(sonuclar)} kontrol | YESIL {len(sonuclar) - len(kirmizi)} | KIRMIZI {len(kirmizi)}")
    if kirmizi:
        print(f"KIRILAN ({len(kirmizi)}): " + ", ".join(kirmizi[:25]) + (" …" if len(kirmizi) > 25 else ""))
        print("SONUC: GATE KAPALI")
        sys.exit(1)
    print("SONUC: GATE ACIK")
    sys.exit(0)


if __name__ == "__main__":
    main()
