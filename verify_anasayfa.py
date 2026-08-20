#!/usr/bin/env python3
"""
verify_anasayfa.py — MANDATE-EXCELARSIV-ANASAYFA-V1 makine doğrulayıcısı

Kullanim:
    python verify_anasayfa.py --check all --base https://excelarsiv.com
    python verify_anasayfa.py --check prices --base http://localhost:4321
    python verify_anasayfa.py --check shopier --eslesme veri/shopier-eslesme.csv
    python verify_anasayfa.py --check hardcoded --src ./src

Cikis kodu:
    0 = tum secili invariantlar YESIL
    1 = en az bir invariant KIRMIZI
    2 = calistirma hatasi (ag, eksik bagimlilik, eksik dosya)

Bagimlilik: requests, beautifulsoup4
    pip install requests beautifulsoup4
"""

import argparse
import csv
import json
import os
import re
import sys
from collections import defaultdict, Counter

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    print("HATA: 'pip install requests beautifulsoup4' gerekli.", file=sys.stderr)
    sys.exit(2)


# --------------------------------------------------------------------------
# Sabitler — mandate ile birebir
# --------------------------------------------------------------------------

FIYAT_RE = re.compile(r"(?<![\d.,])(\d{1,3}(?:\.\d{3})*|\d+)\s*(?:TL|₺)|₺\s*(\d{1,3}(?:\.\d{3})*|\d+)")
SABIT_FIYAT_RE = re.compile(r"[\"'>]\s*(?:₺\s*)?\d{3,4}(?:\.\d{3})?\s*(?:TL|₺)")
YASAK_HEX = re.compile(r"#(?:[0-9a-fA-F]{6}|[0-9a-fA-F]{3})\b")
PLACEHOLDER_KELIMELER = ["lorem ipsum", "todo", "placeholder", "örnek metin", "ornek metin", "xxx"]
DOGRULANAMAZ_IDDIA = [
    "binlerce işletme", "binlerce isletme", "türkiye'nin 1 numarası",
    "türkiye'nin bir numarası", "en çok tercih edilen", "milyonlarca",
    "%100 memnuniyet", "lider platform",
]
ANASAYFA_ZORUNLU_LINKLER = [
    "/sablonlar", "/demo", "/kurumsal-lisans",
    "/ortaklik-mali-musavir", "/basari-hikayeleri", "/neden-excel-arsiv",
]
ANASAYFA_MIN_FIYAT_TL = 2490      # INV-5.5
MAKS_SECTION = 7                  # INV-5.1
MAKS_EYEBROW = 3                  # INV-D7

sonuclar = []


def kayit(kod, gecti, mesaj):
    sonuclar.append((kod, gecti, mesaj))
    print(f"[{'YESIL' if gecti else 'KIRMIZI'}] {kod}: {mesaj}")
    return gecti


def getir(url):
    r = requests.get(url, timeout=30, headers={"User-Agent": "excelarsiv-mandate-verifier/1.0"})
    r.raise_for_status()
    return BeautifulSoup(r.text, "html.parser"), r.text


def fiyatlari_cikar(metin):
    """Metindeki TL fiyatlarini int listesi olarak dondurur."""
    bulunan = []
    for m in FIYAT_RE.finditer(metin):
        ham = m.group(1) or m.group(2)
        if not ham:
            continue
        try:
            bulunan.append(int(ham.replace(".", "")))
        except ValueError:
            pass
    return bulunan


# --------------------------------------------------------------------------
# INV-G01 / GATE-1 — fiyat tutarliligi
# --------------------------------------------------------------------------

def check_prices(base, urun_veri):
    """Ayni slug icin ana sayfa, katalog ve urun sayfasindaki fiyat esit mi."""
    if not urun_veri:
        return kayit("INV-1.3", False, "Urun veri dosyasi verilmedi (--urunler). Fiyat matrisi kurulamadi.")

    matris = defaultdict(dict)
    hata = 0

    _, anasayfa_html = getir(base + "/")
    _, katalog_html = getir(base + "/sablonlar")

    for u in urun_veri:
        slug, ad, fiyat = u["slug"], u["ad"], int(u["fiyat_tl"])
        matris[slug]["veri"] = fiyat

        # urun detay sayfasi
        try:
            _, detay_html = getir(f"{base}/sablon/{slug}")
            detay_fiyatlar = set(fiyatlari_cikar(detay_html))
            matris[slug]["detay"] = fiyat in detay_fiyatlar
            if fiyat not in detay_fiyatlar:
                hata += 1
        except Exception as e:
            matris[slug]["detay"] = f"HATA: {e}"
            hata += 1

        # ana sayfa ve katalogda urun adi geciyorsa fiyati da dogru gecmeli
        for yuzey, html in (("anasayfa", anasayfa_html), ("katalog", katalog_html)):
            if ad[:35] in html:
                yuzey_fiyatlar = set(fiyatlari_cikar(html))
                matris[slug][yuzey] = fiyat in yuzey_fiyatlar
                if fiyat not in yuzey_fiyatlar:
                    hata += 1

    with open("kanit/GATE-1-fiyat-tutarlilik.json", "w", encoding="utf-8") as f:
        json.dump({"fark_sayisi": hata, "matris": matris}, f, ensure_ascii=False, indent=2)

    return kayit("INV-1.3", hata == 0, f"Fiyat farki sayisi: {hata} (hedef 0)")


def check_hardcoded(src_dizin):
    """INV-1.2 — sablon/bilesen dosyalarinda sabit kodlu fiyat olmamali."""
    if not os.path.isdir(src_dizin):
        return kayit("INV-1.2", False, f"Kaynak dizin bulunamadi: {src_dizin}")

    ihlaller = []
    veri_uzantilari = (".json", ".yaml", ".yml", ".csv", ".ts.data")
    for kok, _, dosyalar in os.walk(src_dizin):
        for d in dosyalar:
            yol = os.path.join(kok, d)
            if d.endswith(veri_uzantilari) or "/veri/" in yol or "/data/" in yol:
                continue
            if not d.endswith((".astro", ".tsx", ".jsx", ".vue", ".html", ".md", ".mdx")):
                continue
            try:
                icerik = open(yol, encoding="utf-8").read()
            except Exception:
                continue
            for m in SABIT_FIYAT_RE.finditer(icerik):
                satir = icerik[: m.start()].count("\n") + 1
                ihlaller.append(f"{yol}:{satir}: {m.group(0).strip()}")

    return kayit("INV-1.2", not ihlaller,
                 f"Sabit kodlu fiyat: {len(ihlaller)} adet" +
                 ("" if not ihlaller else "\n    " + "\n    ".join(ihlaller[:20])))


# --------------------------------------------------------------------------
# INV-G02 / GATE-2 — Shopier eslesmesi
# --------------------------------------------------------------------------

def check_shopier(eslesme_yolu, urun_veri):
    if not os.path.isfile(eslesme_yolu):
        return kayit("INV-2.4", False, f"Eslesme dosyasi yok: {eslesme_yolu}")

    with open(eslesme_yolu, encoding="utf-8") as f:
        satirlar = list(csv.DictReader(f))

    zorunlu = {"slug", "ad", "fiyat_tl", "shopier_id"}
    if not zorunlu.issubset(set(satirlar[0].keys() if satirlar else [])):
        return kayit("INV-2.4", False, f"CSV kolonlari eksik. Zorunlu: {sorted(zorunlu)}")

    idler = [s["shopier_id"].strip() for s in satirlar]
    tekrar = [i for i, adet in Counter(idler).items() if adet > 1]
    g1 = kayit("INV-2.1", not tekrar,
               f"Toplam {len(idler)} kayit, benzersiz {len(set(idler))}. "
               f"Tekrar eden ID: {tekrar if tekrar else 'yok'}")

    g2 = True
    if urun_veri:
        veri_map = {u["slug"]: u for u in urun_veri}
        uyumsuz = []
        for s in satirlar:
            u = veri_map.get(s["slug"])
            if not u:
                uyumsuz.append(f"{s['slug']} (urun verisinde yok)")
            elif str(u["fiyat_tl"]) != s["fiyat_tl"].strip():
                uyumsuz.append(f"{s['slug']} fiyat: veri={u['fiyat_tl']} csv={s['fiyat_tl']}")
        g2 = kayit("INV-2.3", not uyumsuz,
                   f"Fiyat uyumsuzlugu: {len(uyumsuz)}" +
                   ("" if not uyumsuz else "\n    " + "\n    ".join(uyumsuz[:20])))

    return g1 and g2


# --------------------------------------------------------------------------
# INV-G03 — footer esitligi
# --------------------------------------------------------------------------

def check_footer(base, sayfalar):
    imzalar = {}
    for yol in sayfalar:
        soup, _ = getir(base + yol)
        footer = soup.find("footer")
        if footer is None:
            return kayit("INV-3.1", False, f"{yol}: <footer> bulunamadi")
        linkler = sorted({a.get("href", "").rstrip("/") for a in footer.find_all("a")})
        imzalar[yol] = linkler

    referans = imzalar[sayfalar[0]]
    farkli = [y for y, l in imzalar.items() if l != referans]
    g1 = kayit("INV-3.1", not farkli,
               f"Footer farkliligi: {farkli if farkli else 'yok'}")

    zorunlu = ["/kurumsal-lisans", "/ortaklik-mali-musavir", "/basari-hikayeleri", "/neden-excel-arsiv"]
    eksik = {y: [z for z in zorunlu if z not in l] for y, l in imzalar.items()}
    eksik = {y: e for y, e in eksik.items() if e}
    g2 = kayit("INV-3.2", not eksik, f"Footer'da eksik zorunlu link: {eksik if eksik else 'yok'}")
    return g1 and g2


# --------------------------------------------------------------------------
# GATE-5 — ana sayfa mimarisi ve tasarim invariantlari
# --------------------------------------------------------------------------

def check_homepage(base, baseline_ic_link=None):
    soup, html = getir(base + "/")
    gecer = True
    rapor = {}

    # INV-5.1 bolum sayisi
    bolumler = soup.find_all("section")
    rapor["section_sayisi"] = len(bolumler)
    gecer &= kayit("INV-5.1", len(bolumler) <= MAKS_SECTION,
                   f"<section> sayisi {len(bolumler)} (ust sinir {MAKS_SECTION})")

    # INV-G04 tek h1
    h1ler = soup.find_all("h1")
    gecer &= kayit("INV-G04", len(h1ler) == 1, f"<h1> sayisi {len(h1ler)} (olmasi gereken 1)")

    # INV-5.2 h1'de 'sablon' gecmemeli
    h1_metin = h1ler[0].get_text(" ", strip=True).lower() if h1ler else ""
    rapor["h1"] = h1_metin
    gecer &= kayit("INV-5.2", "şablon" not in h1_metin and "sablon" not in h1_metin,
                   f"H1: \"{h1_metin[:80]}\"")

    # INV-5.3 arama kutusu yok
    arama = soup.find_all("input", attrs={"type": ["search", "text"]})
    gecer &= kayit("INV-5.3", len(arama) == 0,
                   f"Ana sayfada arama/metin input sayisi: {len(arama)} (olmasi gereken 0)")

    # INV-5.4 bos src / bos alt yok
    def js_dolduruluyor(img):
        return any(k.startswith("data-") and ("dialog" in k or "src" in k) for k in img.attrs)

    bos_gorsel = [str(i)[:120] for i in soup.find_all("img")
                  if (not (i.get("alt") or "").strip())
                  or (not (i.get("src") or "").strip() and not js_dolduruluyor(i))]
    gecer &= kayit("INV-5.4", not bos_gorsel,
                   f"src/alt bos gorsel: {len(bos_gorsel)}" +
                   ("" if not bos_gorsel else "\n    " + "\n    ".join(bos_gorsel[:5])))

    # INV-5.5 minimum fiyat
    fiyatlar = [f for f in fiyatlari_cikar(soup.get_text(" ", strip=True)) if f >= 100]
    dusukler = sorted({f for f in fiyatlar if f < ANASAYFA_MIN_FIYAT_TL})
    rapor["anasayfa_fiyatlari"] = sorted(set(fiyatlar))
    gecer &= kayit("INV-5.5", not dusukler,
                   f"Ana sayfada {ANASAYFA_MIN_FIYAT_TL} TL altinda fiyat: {dusukler if dusukler else 'yok'}")

    # INV-5.8 zorunlu ic linkler
    hrefler = {a.get("href", "").rstrip("/") for a in soup.find_all("a")}
    eksik = [z for z in ANASAYFA_ZORUNLU_LINKLER if z not in hrefler]
    gecer &= kayit("INV-5.8", not eksik, f"Ana sayfada eksik zorunlu link: {eksik if eksik else 'yok'}")

    # INV-5.7 ic link sayisi gerilemedi
    ic_link = len([h for h in hrefler if h.startswith("/")])
    rapor["ic_link_sayisi"] = ic_link
    if baseline_ic_link is not None:
        gecer &= kayit("INV-5.7", ic_link >= baseline_ic_link,
                       f"Ic link {ic_link} (baseline {baseline_ic_link})")
    else:
        print(f"[ATLA ] INV-5.7: baseline verilmedi, mevcut ic link sayisi {ic_link}")

    # INV-G05 placeholder — yalniz gorunur metin (HTML attribute'lari haric)
    dusuk = soup.get_text(" ", strip=True).lower()
    bulunan_ph = [k for k in PLACEHOLDER_KELIMELER if k in dusuk]
    gecer &= kayit("INV-G05", not bulunan_ph, f"Placeholder ifade: {bulunan_ph if bulunan_ph else 'yok'}")

    # INV-G06 dogrulanamaz iddia
    bulunan_iddia = [k for k in DOGRULANAMAZ_IDDIA if k in dusuk]
    gecer &= kayit("INV-G06", not bulunan_iddia,
                   f"Dogrulanamaz iddia: {bulunan_iddia if bulunan_iddia else 'yok'}")

    os.makedirs("kanit", exist_ok=True)
    with open("kanit/GATE-5-anasayfa-denetim.json", "w", encoding="utf-8") as f:
        json.dump(rapor, f, ensure_ascii=False, indent=2)

    return gecer


def check_design(css_dizin):
    """INV-D1 / INV-D5 — renk ve radius disiplini (CSS kaynak taramasi)."""
    if not os.path.isdir(css_dizin):
        return kayit("INV-D1", False, f"CSS dizini bulunamadi: {css_dizin}")

    izinli_kok = {"14532d", "166534", "15803d", "22c55e", "dcfce7", "f0fdf4",
                  "000000", "ffffff", "fff", "000"}
    yabanci, radius_ihlal = [], []

    for kok, _, dosyalar in os.walk(css_dizin):
        for d in dosyalar:
            if not d.endswith((".css", ".scss", ".astro", ".tsx", ".jsx", ".vue")):
                continue
            yol = os.path.join(kok, d)
            try:
                icerik = open(yol, encoding="utf-8").read()
            except Exception:
                continue
            for m in YASAK_HEX.finditer(icerik):
                hx = m.group(0).lstrip("#").lower()
                if hx in izinli_kok:
                    continue
                r, g, b = (int(hx[i:i + 2], 16) for i in (0, 2, 4)) if len(hx) == 6 else (0, 0, 0)
                notr = max(r, g, b) - min(r, g, b) < 18          # gri skala serbest
                yesil_ailesi = g >= r and g >= b                  # yesil aksan serbest
                if not (notr or yesil_ailesi):
                    yabanci.append(f"{yol}: #{hx}")
            for m in re.finditer(r"border-radius:\s*([^;]+);", icerik):
                deger = m.group(1).strip()
                if deger not in ("0", "0px", "0rem", "var(--radius)"):
                    radius_ihlal.append(f"{yol}: border-radius: {deger}")

    g1 = kayit("INV-D1", not yabanci,
               f"Aksan disi renk: {len(yabanci)}" +
               ("" if not yabanci else "\n    " + "\n    ".join(sorted(set(yabanci))[:20])))
    g2 = kayit("INV-D5", not radius_ihlal,
               f"radius ihlali: {len(radius_ihlal)}" +
               ("" if not radius_ihlal else "\n    " + "\n    ".join(sorted(set(radius_ihlal))[:20])))
    return g1 and g2


# --------------------------------------------------------------------------

def urunleri_yukle(yol):
    if not yol:
        return None
    if not os.path.isfile(yol):
        print(f"UYARI: urun veri dosyasi yok: {yol}", file=sys.stderr)
        return None
    if yol.endswith(".json"):
        return json.load(open(yol, encoding="utf-8"))
    with open(yol, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def main():
    p = argparse.ArgumentParser(description="MANDATE-EXCELARSIV-ANASAYFA-V1 dogrulayici")
    p.add_argument("--check", default="all",
                   choices=["all", "prices", "hardcoded", "shopier", "footer", "homepage", "design"])
    p.add_argument("--base", default="https://excelarsiv.com", help="Site kok URL (preview veya canli)")
    p.add_argument("--urunler", default="veri/urunler.json", help="Urun veri dosyasi (json/csv)")
    p.add_argument("--eslesme", default="veri/shopier-eslesme.csv")
    p.add_argument("--src", default="./src")
    p.add_argument("--css", default="./src")
    p.add_argument("--baseline-ic-link", type=int, default=None)
    a = p.parse_args()

    base = a.base.rstrip("/")
    os.makedirs("kanit", exist_ok=True)
    urunler = urunleri_yukle(a.urunler)
    calisacak = ["prices", "hardcoded", "shopier", "footer", "homepage", "design"] \
        if a.check == "all" else [a.check]

    try:
        if "hardcoded" in calisacak:
            check_hardcoded(a.src)
        if "shopier" in calisacak:
            check_shopier(a.eslesme, urunler)
        if "footer" in calisacak:
            check_footer(base, ["/", "/sablonlar", "/demo"])
        if "homepage" in calisacak:
            check_homepage(base, a.baseline_ic_link)
        if "design" in calisacak:
            check_design(a.css)
        if "prices" in calisacak:
            check_prices(base, urunler)
    except requests.RequestException as e:
        print(f"HATA (ag): {e}", file=sys.stderr)
        sys.exit(2)

    kirmizi = [k for k, gecti, _ in sonuclar if not gecti]
    print("\n" + "=" * 60)
    print(f"TOPLAM: {len(sonuclar)} invariant | YESIL {len(sonuclar) - len(kirmizi)} | KIRMIZI {len(kirmizi)}")
    if kirmizi:
        print("KIRILAN: " + ", ".join(kirmizi))
        print("SONUC: GATE KAPALI")
        sys.exit(1)
    print("SONUC: GATE ACIK")
    sys.exit(0)


if __name__ == "__main__":
    main()
