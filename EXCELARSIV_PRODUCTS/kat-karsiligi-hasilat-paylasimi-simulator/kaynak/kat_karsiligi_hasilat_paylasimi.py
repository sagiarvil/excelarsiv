#!/usr/bin/env python3
"""
Kat Karşılığı / Hasılat Paylaşımı Simülatör — üretim betiği (A3+A1 / manda v6).
Arsa–müteahhit pay paylaşımı, net hasılat, başabaş ve sözleşme→inşaat→teslim→paylaşım zinciri.
"""

from __future__ import annotations

import hashlib
import os
import shutil
import sys
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Font, PatternFill, Protection
from openpyxl.utils import get_column_letter

KOK = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
if KOK not in sys.path:
    sys.path.insert(0, KOK)

from excel_uretim.ortak import (  # noqa: E402
    CATI,
    DIS_REF_YEŞIL,
    GIRIS_SARI,
    GIRIS_YAZI,
    GRİ,
    KOYU_LACIVERT,
    KRITIK,
    NORMAL,
    ORTA_MAVI,
    SIFRE,
    TARİH,
    TL,
    YÜZDE,
    ad_ekle,
    alt_bant,
    bant,
    baski_hazirla,
    baslik_satiri,
    dogrulama,
    genislik,
    giris_hucreleri,
    h,
    sabitle,
    sayfa_hazirla,
    sayfa_koru,
    tablo_ekle,
    tablo_formullerini_hucrelere_yaz,
    yorum_ekle,
)
from ortak.akis_motoru import (  # noqa: E402
    durum_haritasi_dogrula,
    gecersiz_gecis_formul,
    kuyruk_sayim,
    kuyruk_tutar,
    sonraki_kimlik_formul,
    yas_kova_formul,
    yaslandirma_gun,
    yetim_kayit_formul,
)
from ortak.mevzuat_motoru import (  # noqa: E402
    KURAL_BASLIK,
    MOTOR_BASLIK,
    VAKA_BASLIK,
    belirsizlik_beyanlari,
)

URUN_AD = "Kat Karşılığı / Hasılat Paylaşımı Simülatör"
SURUM = "1.0.0"
RENK = "0F2742"
KAPASITE = 1000
DEMO_KART = 6
DEMO_AKIS = 18
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIH = date(2026, 8, 11)

DEMO = {
    "sirket_adi": "Örnek Yapı Ortaklığı A.Ş.",
    "proje_adi": "Cadde Konut Projesi",
    "hesapYili": 2026,
    "arsa": 20_000_000,
    "insaat": 35_000_000,
    "hasilat": 100_000_000,
    "muteahhit_pay": 0.60,
    "arsa_sahibi_pay": 0.40,
    "sure_ay": 24,
    "finansman": 0.12,
    "yorum": "A",
}

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "SOZLESME", "durum_adi": "Sözleşme", "sira": 1,
     "izinli_sonraki_durumlar": "INSAAT|IPTAL", "kategori": "acik"},
    {"durum_id": "INSAAT", "durum_adi": "İnşaat", "sira": 2,
     "izinli_sonraki_durumlar": "TESLIM|IPTAL", "kategori": "acik"},
    {"durum_id": "TESLIM", "durum_adi": "Teslim", "sira": 3,
     "izinli_sonraki_durumlar": "PAYLASIM|IPTAL", "kategori": "acik"},
    {"durum_id": "PAYLASIM", "durum_adi": "Paylaşım", "sira": 4,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "IPTAL", "durum_adi": "İptal", "sira": 5,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "SOZLESME|INSAAT",
    "SOZLESME|IPTAL",
    "INSAAT|TESLIM",
    "INSAAT|IPTAL",
    "TESLIM|PAYLASIM",
    "TESLIM|IPTAL",
]

PROJE_KARTLARI = [
    ("KKH-00001", "Cadde Arsa Ortaklığı", "İstanbul", 4200),
    ("KKH-00002", "Sahil Villa Arsa", "Antalya", 3100),
    ("KKH-00003", "Merkez Ofis Arsa", "Ankara", 2800),
    ("KKH-00004", "Sanayi Lojistik Arsa", "Kocaeli", 5600),
    ("KKH-00005", "Kampüs Konut Arsa", "İzmir", 3900),
    ("KKH-00006", "Park Evleri Arsa", "Bursa", 2500),
]


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: Kat karşılığı / hasılat paylaşımı kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Pay, kâr, başabaş ve karar bozulur | Kaynak: Sözleşme modeli KKH",
    )


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj="", yorum=True):
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    hucre.protection = Protection(locked=False)
    if yorum and baslik:
        _kim(ws, f"{get_column_letter(c)}{r}", baslik, mesaj or "Uygun türde değer girin")
    return hucre


def kural_cek(kural_id: str) -> str:
    """Eşikleri tblKurallar'dan çeker; yıl seçimi kkh_yilTaban ile."""
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-kkh_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(kkh_yuvarMax,IFERROR(N(G12),0)))"


def _guvenli_formul(formul: str, birim: str) -> str:
    if not formul.startswith("="):
        formul = "=" + formul
    inner = formul[1:]
    if inner.upper().startswith("IFERROR("):
        return formul
    yedek = '""' if birim in ("metin", "kod") else "0"
    return f"=IFERROR({inner},{yedek})"


def _demo_akis():
    """Demo akış: SOZLESME/INSAAT/TESLIM/PAYLASIM + yetim + geçersiz geçiş + zincir kırık."""
    satirlar = []
    durumlar = ["SOZLESME", "INSAAT", "TESLIM", "PAYLASIM", "INSAAT", "TESLIM",
                "SOZLESME", "INSAAT", "IPTAL", "INSAAT", "TESLIM", "PAYLASIM",
                "SOZLESME", "INSAAT", "TESLIM", "INSAAT", "SOZLESME", "PAYLASIM"]
    oncekiler = ["", "SOZLESME", "INSAAT", "TESLIM", "SOZLESME", "INSAAT",
                 "", "SOZLESME", "INSAAT", "SOZLESME", "INSAAT", "TESLIM",
                 "", "SOZLESME", "INSAAT", "SOZLESME", "", "TESLIM"]
    for i in range(1, DEMO_AKIS + 1):
        kart = PROJE_KARTLARI[(i - 1) % DEMO_KART][0]
        if i == 14:
            kart = ""
        if i == 16:
            kart = "KKH-99999"
        d = durumlar[i - 1]
        o = oncekiler[i - 1]
        if i == 10:
            o, d = "SOZLESME", "TESLIM"
        if i == 17:
            o, d = "PAYLASIM", "INSAAT"
        soz = 8_000_000 + i * 250_000
        ruh = soz
        ins = soz if i not in (8, 12) else soz + 75_000
        tes = ins if i != 11 else ins - 40_000
        pay = tes if d == "PAYLASIM" else 0
        if d == "IPTAL":
            pay = 0
        tar = RAPOR_TARIH - timedelta(days=(i * 5) % 90)
        satirlar.append({
            "id": f"AKS-{i:05d}",
            "kart": kart,
            "tarih": tar,
            "onceki": o,
            "durum": d,
            "sozlesme": soz,
            "ruhsat": ruh,
            "insaat": ins,
            "teslim": tes,
            "paylasim": pay,
            "tutar": soz if d in ("SOZLESME", "INSAAT", "TESLIM") else (pay if d == "PAYLASIM" else 0),
            "donem": "2026-08",
            "aciklama": f"Kat karşılığı işlem {i}",
        })
    return satirlar


ORNEK_AKIS = _demo_akis()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=40)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=40)
    h(ws, 4, 1,
      "Arsa değeri, inşaat maliyeti ve brüt hasılatı girin; net hasılat, tarafların payı ve kârı, "
      "başabaş ve UYGUN / DİKKAT / UYGUN DEĞİL kararını üretin; sözleşme→inşaat→teslim→paylaşım zincirini izleyin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Net hasılat (KDV ayrıştırmalı) + müteahhit / arsa sahibi pay ve kâr hesabı",
        "Başabaş brüt hasılat ve finansman yükü ile karar eşiği",
        "Çok yıllı kural tablosu (KKH-001..005) — eşikler MOTOR'a sabit yazılmaz",
        "Sözleşme → İnşaat → Teslim → Paylaşım durum makinesi + zincir bayrağı",
        "Altın vakalar, senaryo/tornado ve KANIT_RAPORU (A4 imza)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Müteahhit mali işler — pay, kâr ve başabaşı KANIT_RAPORU ile sunmak",
        "Arsa sahibi / ortak — UYGUN / DİKKAT / UYGUN DEĞİL kararını görmek",
        "SMMM — sözleşme→inşaat→teslim→paylaşım zincirini dosyalamak",
    ], 14):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) HIZLI_BASLANGIC → 2) GIRDI sarı hücreler → 3) PANO karar → 4) KANIT_RAPORU yazdır.",
      yazi="333333", kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 21, 1, f"Sürüm {SURUM} | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 72})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Üç adımda sonuç", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Girdiler",
         "GIRDI sayfasında arsa, inşaat, hasılat, pay oranları ve finansmanı sarı hücrelere girin."),
        ("Adım 2 — Karar",
         "PANO'da müteahhit kârını, başabaşı ve UYGUN / DİKKAT / UYGUN DEĞİL rozetini izleyin."),
        ("Adım 3 — Kanıt",
         "KANIT_RAPORU'nu yazdırıp imzalayın; VAKALAR tutarlılığını ve AKIS zincirini kontrol edin."),
    ]
    for i, (b, m) in enumerate(adimlar, 5):
        h(ws, i, 1, b, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        _kim(ws, f"A{i}", b, "Sırayla izleyin")
    h(ws, 9, 1, "Sık yapılan hatalar", kalin=True, yazi=KRITIK)
    for i, m in enumerate([
        "Müteahhit + arsa sahibi pay toplamını %100 dışına yazmak",
        "Brüt hasılat yerine net hasılat girmek (KDV çift ayrışır)",
        "Finansman oranını yıllık yerine aylık sanmak",
    ], 10):
        h(ws, i, 1, "• " + m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=8)
    h(ws, 14, 1,
      "Bu dosya karar destek aracıdır; kesin hukuki/mali görüş veya sözleşme yerine geçmez.",
      yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 28, "B": 70})


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Proje ve paylaşım girdileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. KDV ve eşikler KURALLAR'dan; karar formülle üretilir.",
      yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Şirket / Ortaklık Adı", DEMO["sirket_adi"], None, "sirket_adi",
         "Şirket veya ortaklık unvanını yazın."),
        (7, "Proje Adı", DEMO["proje_adi"], None, "proje_adi",
         "Proje adını yazın."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili",
         "2025, 2026 veya 2027 seçin."),
        (9, "Arsa Değeri [₺]", DEMO["arsa"], TL, "arsa",
         "Arsa rayiç / katkı değerini TL girin."),
        (10, "İnşaat Maliyeti [₺]", DEMO["insaat"], TL, "insaat",
         "Toplam inşaat maliyetini TL girin."),
        (11, "Satış Hasılatı Tahmini (brüt) [₺]", DEMO["hasilat"], TL, "hasilat",
         "KDV dahil brüt satış hasılatı tahminini TL girin."),
        (12, "Müteahhit Payı [%]", DEMO["muteahhit_pay"], YÜZDE, "muteahhit_pay",
         "Müteahhit pay oranını yüzde girin."),
        (13, "Arsa Sahibi Payı [%]", DEMO["arsa_sahibi_pay"], YÜZDE, "arsa_sahibi_pay",
         "Arsa sahibi pay oranını yüzde girin."),
        (14, "Süre [ay]", DEMO["sure_ay"], CATI, "sure_ay",
         "Proje süresini ay cinsinden girin."),
        (15, "Finansman Maliyeti [%/yıl]", DEMO["finansman"], YÜZDE, "finansman",
         "Yıllık finansman maliyet oranını yüzde girin."),
        (16, "Yorum Modu (A/B)", DEMO["yorum"], None, "yorum",
         "Yorum A (standart) veya B (inşaat ×1,05) seçin."),
    ]
    h(ws, 5, 1, "Alan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 2, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 3, "Birim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    for satir, etiket, deger, sayi, _ad, mesaj in alanlar:
        if sayi == TL:
            birim = "₺"
        elif sayi == YÜZDE:
            birim = "%"
        elif _ad in ("hesapYili", "sure_ay"):
            birim = "yıl" if _ad == "hesapYili" else "ay"
        else:
            birim = "metin"
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi=GIRIS_YAZI)
        ws.cell(satir, 2).protection = Protection(locked=False)
        h(ws, satir, 3, birim, yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 18, 1, "KDV oranı (ayna / KURALLAR)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 2, kural_cek("KKH-002"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Pay toplam hedefi (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2, kural_cek("KKH-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Müteahhit kâr (ayna) [₺]", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=IFERROR(kkh_muteahhitKar,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 21, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B16)=0,kkh_girisBeklenen=0),0,'
      'ROUND(COUNTA(B6:B16)/kkh_girisBeklenen*100,0)),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)

    h(ws, 23, 1, "Giriş tablosu (otomasyon / boş-mod)", kalin=True, yazi=KOYU_LACIVERT)
    sutunlar = ["Alan Anahtarı", "Değer", "Birim", "Zorunlu", "Kaynak", "Kontrol"]
    baslik_satiri(ws, 24, sutunlar)
    formuller = {
        "Kontrol": (
            '=IF(tblGirdi[[#This Row],[Alan Anahtarı]]="","",'
            'IF(tblGirdi[[#This Row],[Değer]]="","EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblGirdi", "A24:F1023", sutunlar, formuller)
    ornek_satir = [
        ("sirket_adi", DEMO["sirket_adi"], "metin", "Evet", "GIRDI"),
        ("proje_adi", DEMO["proje_adi"], "metin", "Evet", "GIRDI"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Evet", "GIRDI"),
        ("arsa", DEMO["arsa"], "TL", "Evet", "GIRDI"),
        ("insaat", DEMO["insaat"], "TL", "Evet", "GIRDI"),
        ("hasilat", DEMO["hasilat"], "TL", "Evet", "GIRDI"),
        ("muteahhit_pay", DEMO["muteahhit_pay"], "oran", "Evet", "GIRDI"),
        ("arsa_sahibi_pay", DEMO["arsa_sahibi_pay"], "oran", "Evet", "GIRDI"),
        ("sure_ay", DEMO["sure_ay"], "ay", "Evet", "GIRDI"),
        ("finansman_orani", DEMO["finansman"], "oran", "Evet", "GIRDI"),
        ("yorum", DEMO["yorum"], "A/B", "Evet", "GIRDI"),
    ]
    tl_alanlar = {"arsa", "insaat", "hasilat"}
    oran_alanlar = {"muteahhit_pay", "arsa_sahibi_pay", "finansman_orani"}
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 25):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in tl_alanlar:
            ws.cell(i, 2).number_format = TL
        elif a in oran_alanlar:
            ws.cell(i, 2).number_format = YÜZDE
        elif a in ("hesapYili", "sure_ay"):
            ws.cell(i, 2).number_format = CATI
        ws.cell(i, 3).value = b
        ws.cell(i, 4).value = z
        ws.cell(i, 5).value = k
        for kolon in range(1, 6):
            ws.cell(i, kolon).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
            ws.cell(i, kolon).protection = Protection(locked=False)
            _kim(ws, f"{get_column_letter(kolon)}{i}", a, "Tablo girişi; PANO ile senkron tutun")

    dogrulama(ws, "list", "ListeYillar", "B8",
              baslik="Hesap Yılı", mesaj="2025, 2026 veya 2027 seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Yalnızca listeden seçin.", bos=False)
    dogrulama(ws, "list", "ListeYorumAB", "B16",
              baslik="Yorum", mesaj="Yorum A veya B seçin.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B seçin.", bos=False)
    dogrulama(ws, "list", "ListeYillar", "B27",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeYorumAB", "B35",
              baslik="Yorum", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.")
    for aralik, baslik in [
        ("B9", "Arsa"), ("B10", "İnşaat"), ("B11", "Hasılat"),
        ("B28", "Arsa"), ("B29", "İnşaat"), ("B30", "Hasılat"),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj="0 ile 1e12 arasında TL.",
                  hata_baslik="Geçersiz tutar", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000000")
    for aralik, baslik in [
        ("B12", "Müteahhit Pay"), ("B13", "Arsa Pay"), ("B15", "Finansman"),
        ("B31", "Müteahhit Pay"), ("B32", "Arsa Pay"), ("B34", "Finansman"),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj="0 ile 1 arasında oran.",
                  hata_baslik="Geçersiz oran", hata_mesaj="0-1 aralığı.",
                  isaret="between", f2="1")
    dogrulama(ws, "whole", "1", "B14",
              baslik="Süre", mesaj="1-120 ay.",
              hata_baslik="Geçersiz", hata_mesaj="1-120.",
              isaret="between", f2="120")

    giris_hucreleri(ws, 6, 16, [2])
    giris_hucreleri(ws, 25, 35, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 42, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1025, "Sarı alanlar giriş; eşik ve karar formülleri kilitlidir.")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Mevzuat / model kural tablosu (yıl yan yana)", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Eşikler MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.",
      yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["KKH-001", "Sözleşme modeli", "pay_toplam_hedef",
         "müteahhit+arsa sahibi pay = hedef", "pay_toplam_hedef",
         1.0, 1.0, 1.0, "01.01.2025", "Pay oranları toplamı hedefi"],
        ["KKH-002", "KDVK", "kdv_orani", "brüt→net ayrıştırma", "kdv_orani",
         0.20, 0.20, 0.20, "01.01.2025", "KDV oranı"],
        ["KKH-003", "İç politika", "dikkat_marj_esik",
         "müteahhit marj < eşik", "dikkat_marj_esik",
         0.05, 0.05, 0.05, "01.01.2025", "DİKKAT marj eşiği"],
        ["KKH-004", "İç politika", "min_kar_esik",
         "müteahhit kâr < eşik", "min_kar_esik_tl",
         0, 0, 0, "01.01.2025", "UYGUN DEĞİL min kâr (TL)"],
        ["KKH-005", "Genel", "yuvarlama", "her hesap", "yuvarlama_ondalik",
         2, 2, 2, "01.01.2025", "Yuvarlama basamağı"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = None
            if k in (6, 7, 8):
                if s[0] in ("KKH-001", "KKH-002", "KKH-003"):
                    sayi = YÜZDE
                else:
                    sayi = CATI if s[0] == "KKH-005" else TL
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-kkh_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "KDV (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("KKH-002"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Dikkat marj eşiği (aktif yıl)", yazi="333333")
    h(ws, 16, 2, kural_cek("KKH-003"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"KKH-001=",TEXT(INDEX(CHOOSE(MAX(1,MIN(3,'
      'IF(OR(hesapYili="",hesapYili=0),1,hesapYili-kkh_yilTaban))),'
      'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
      'MATCH("KKH-001",tblKurallar[kural_id],0)),"0%"),'
      '"KKH-002=",TEXT(B15,"0.0%"),"KKH-003=",TEXT(B16,"0.0%"),'
      '"yıl=",IF(OR(hesapYili="",hesapYili=0),"-",hesapYili)),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    for col in ("F", "G", "H"):
        dogrulama(ws, "decimal", "0", f"{col}7:{col}11",
                  baslik="Kural değeri", mesaj="Yıl kolonuna sayısal parametre girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ve üzeri sayı.",
                  isaret="between", f2="10000000")
    genislik(ws, {get_column_letter(i): w for i, w in enumerate(
        [14, 18, 22, 28, 18, 12, 12, 12, 12, 28, 14], 1)})
    ws.merge_cells("A3:K3")
    ws.merge_cells("A4:K4")
    sabitle(ws, "A7")


def motor(ws):
    sayfa_hazirla(ws, "MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Hesap zinciri — her adımda kural_id", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Eşikler tblKurallar'dan çekilir; sabit eşik yoktur.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")
    baslik_satiri(ws, 6, [*MOTOR_BASLIK, "deger"])

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    # G7.. — kurallar
    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif model yılı", "KKH-005")
    ekle("Pay toplam hedefi", km("KKH-001"), "oran", "Pay toplamı hedefi", "KKH-001")
    ekle("KDV oranı", km("KKH-002"), "oran", "KDV oranı", "KKH-002")
    ekle("Dikkat marj eşiği", km("KKH-003"), "oran", "DİKKAT marj eşiği", "KKH-003")
    ekle("Min kâr eşiği", km("KKH-004"), "TL", "UYGUN DEĞİL min kâr", "KKH-004")
    ekle("Yuvarlama ondalık", km("KKH-005"), "adet", "Yuvarlama basamağı", "KKH-005")
    # G13.. — girdiler
    ekle("Arsa değeri", "=GIRDI!B9", "TL", "Arsa değeri", "KKH-004")
    ekle("İnşaat maliyeti", "=GIRDI!B10", "TL", "İnşaat maliyeti", "KKH-004")
    ekle("Brüt hasılat", "=GIRDI!B11", "TL", "Brüt satış hasılatı", "KKH-002")
    ekle("Müteahhit pay oranı", "=GIRDI!B12", "oran", "Müteahhit payı", "KKH-001")
    ekle("Arsa sahibi pay oranı", "=GIRDI!B13", "oran", "Arsa sahibi payı", "KKH-001")
    ekle("Süre (ay)", "=GIRDI!B14", "adet", "Proje süresi", "KKH-005")
    ekle("Finansman oranı", "=GIRDI!B15", "oran", "Yıllık finansman", "KKH-005")
    ekle("Yorum çarpanı",
         '=IF(GIRDI!B16="B",1+kkh_yorumBCarpan,1)',
         "çarpan", "Yorum A=1 / B=1,05 inşaat", "KKH-005")
    # G21.. — çekirdek
    ekle("Net hasılat",
         f"=ROUND(IF(1+G9=0,0,G15/(1+G9)),{_YV})",
         "TL", "Brüt/(1+KDV)", "KKH-002")
    ekle("Müteahhit pay TL",
         f"=ROUND(G21*G16,{_YV})",
         "TL", "Net × müteahhit pay", "KKH-001")
    ekle("Arsa pay TL",
         f"=ROUND(G21*G17,{_YV})",
         "TL", "Net × arsa pay", "KKH-001")
    ekle("Finansman TL",
         f"=ROUND(G14*G19*(G18/12),{_YV})",
         "TL", "İnşaat×oran×(ay/12)", "KKH-005")
    ekle("İnşaat (yorumlu)",
         f"=ROUND(G14*G20,{_YV})",
         "TL", "Yorum B: inşaat×1,05", "KKH-005")
    ekle("Müteahhit kâr",
         f"=ROUND(G22-G25-G24,{_YV})",
         "TL", "Pay − inşaat(yorum) − finansman", "KKH-004")
    ekle("Arsa sahibi kâr",
         f"=ROUND(G23-G13,{_YV})",
         "TL", "Arsa pay − arsa değeri", "KKH-004")
    ekle("Müteahhit marj",
         "=IF(G22=0,0,G26/G22)",
         "oran", "Müteahhit kâr / pay", "KKH-003")
    ekle("Pay toplamı", "=G16+G17", "oran", "Pay oranları toplamı", "KKH-001")
    ekle("Pay sapması", "=ABS(G29-G8)", "oran", "|toplam−hedef|", "KKH-001")
    ekle("Başabaş net",
         f"=ROUND(IF(G16=0,0,(G14+G24)/G16),{_YV})",
         "TL", "(inşaat+finansman)/pay", "KKH-002")
    ekle("Başabaş brüt",
         f"=ROUND(G31*(1+G9),{_YV})",
         "TL", "Başabaş net×(1+KDV)", "KKH-002")
    ekle("Hasılat / başabaş",
         "=IF(G32=0,0,G15/G32)",
         "oran", "Brüt hasılat / başabaş brüt", "KKH-002")
    ekle("Aylık inşaat harcama",
         f"=ROUND(IF(G18=0,0,G14/G18),{_YV})",
         "TL", "Nakit takvim özeti (aylık)", "KKH-005")
    ekle("Toplam nakit çıkış",
         f"=ROUND(G14+G24,{_YV})",
         "TL", "İnşaat + finansman", "KKH-005")
    # G36.. — karar
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(OR(G26<G11,G30>kkh_payTolerans,G15<G32),"UYGUN_DEGIL",'
         'IF(G28<G10,"DIKKAT","UYGUN")))',
         "kod", "Karar kodu", "KKH-005")
    ekle("Karar metni",
         '=IF(G36="VERI_YOK","VERİ YOK",'
         'IF(G36="UYGUN_DEGIL","UYGUN DEĞİL",'
         'IF(G36="DIKKAT","DİKKAT","UYGUN")))',
         "metin", "Kullanıcıya karar", "KKH-005")
    ekle("Gerekçe",
         '=IF(G36="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G36="UYGUN_DEGIL","Müteahhit kârı eşiğin altında, pay toplamı bozuk veya hasılat başabaşın altında.",'
         'IF(G36="DIKKAT","Müteahhit marjı dikkat eşiğinin altında; pay ve maliyetleri gözden geçirin.",'
         '"Pay dengesi uygun, müteahhit kârı ve hasılat başabaşın üzerinde.")))',
         "metin", "Gerekçe cümlesi", "KKH-005")
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B16)=0,kkh_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B16)/kkh_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "KKH-005")
    ekle("Risk skoru",
         '=IF(G36="VERI_YOK",0,IF(G36="UYGUN",kkh_riskDusuk,'
         'IF(G36="DIKKAT",kkh_riskOrta,kkh_riskYuksek)))',
         "puan", "Paylaşım risk skoru", "KKH-005")
    # G41.. — senaryo / duyarlılık
    ekle("Senaryo iyimser kâr",
         f"=ROUND(G26*kkh_senaryoIyi,{_YV})",
         "TL", "İyimser müteahhit kâr", "KKH-004")
    ekle("Senaryo kötümser kâr",
         f"=ROUND(G26*kkh_senaryoKotu,{_YV})",
         "TL", "Kötümser müteahhit kâr", "KKH-004")
    ekle("M-SEN KDV+1p",
         f"=ROUND(G15/(1+G9+kkh_kdvSikilastirma)*G16-G25-G24,{_YV})",
         "TL", "KDV sıkılaştırmalı kâr", "KKH-002")
    ekle("Tornado: hasılat etkisi",
         f"=ABS(ROUND(G15*kkh_tornadoHasilat/(1+G9)*G16-G25-G24,{_YV})-G26)",
         "TL", "Hasılat ± etki", "KKH-002")
    ekle("Tornado: inşaat etkisi",
         f"=ABS(ROUND(G22-G14*kkh_tornadoMaliyet*G20-G24,{_YV})-G26)",
         "TL", "İnşaat ± etki", "KKH-004")
    ekle("Tornado: pay etkisi",
         f"=ABS(ROUND(G21*G16*kkh_tornadoPay-G25-G24,{_YV})-G26)",
         "TL", "Müteahhit pay ± etki", "KKH-001")
    ekle("Tahmin üst", "=G26*kkh_tahminUst", "TL", "Tahmin üst bant", "KKH-005")
    ekle("Tahmin alt", "=G26*kkh_tahminAlt", "TL", "Tahmin alt bant", "KKH-005")
    ekle("Finansman / inşaat",
         "=IF(G14=0,0,G24/G14)",
         "oran", "Finansman yoğunluğu", "KKH-005")
    ekle("Arsa kâr / arsa",
         "=IF(G13=0,0,G27/G13)",
         "oran", "Arsa sahibi getiri", "KKH-004")
    ekle("Müteahhit kâr (pano)", "=G26", "TL", "Kâr ayna", "KKH-004")
    ekle("Başabaş brüt (pano)", "=G32", "TL", "Başabaş ayna", "KKH-002")
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"KKH-001→pay toplam","KKH-002→KDV",'
         '"KKH-003→dikkat marj","KKH-004→min kâr","KKH-005→yuvarlama")',
         "metin", "Kanıt atıfları", "KKH-001")
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Net=",TEXT(G21,"₺ #,##0"),'
         '"MuteahhitKar=",TEXT(G26,"₺ #,##0"),"ArsaKar=",TEXT(G27,"₺ #,##0"),'
         '"BasabesBrut=",TEXT(G32,"₺ #,##0"),"Karar=",G37)',
         "metin", "Rapor özeti", "KKH-005")
    ekle("Akış yetim adet",
         '=COUNTIF(tblAkis[YetimBayrak],"YETİM")',
         "adet", "Yetim işlem sayısı", "KKH-005")
    ekle("Akış zincir kırık",
         '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")',
         "adet", "Zincir kırık sayısı", "KKH-005")
    ekle("Akış geçersiz geçiş",
         '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")',
         "adet", "Geçersiz geçiş sayısı", "KKH-005")

    assert len(adimlar) >= 40, len(adimlar)

    for i, (ad, formul, birim, ne, kid) in enumerate(adimlar, 1):
        r = 6 + i
        guvenli = _guvenli_formul(formul, birim)
        h(ws, r, 1, i, sayi=CATI, hiza="center")
        h(ws, r, 2, ad, kaydir=True)
        h(ws, r, 3, guvenli, yazi=DIS_REF_YEŞIL, kaydir=True)
        h(ws, r, 4, birim, hiza="center")
        h(ws, r, 5, ne, kaydir=True, yazi="333333")
        h(ws, r, 6, kid, hiza="center", kalin=True, yazi="B08948")
        h(ws, r, 7, guvenli, yazi=DIS_REF_YEŞIL)
        if birim == "TL":
            ws.cell(r, 7).number_format = TL
            ws.cell(r, 3).number_format = TL
        elif birim in ("oran", "çarpan"):
            ws.cell(r, 7).number_format = YÜZDE
            ws.cell(r, 3).number_format = YÜZDE
        elif birim in ("puan", "yıl", "adet"):
            ws.cell(r, 7).number_format = CATI
            ws.cell(r, 3).number_format = CATI

    genislik(ws, {"A": 8, "B": 36, "C": 55, "D": 10, "E": 32, "F": 12, "G": 22})
    sabitle(ws, "A7")
    h(ws, 6 + len(adimlar) + 2, 9,
      "Motor adımları kilitlidir; eşikler yalnızca KURALLAR'dan gelir.", yazi=GRİ, kaydir=True)
    return len(adimlar)


def kartlar(ws):
    sayfa_hazirla(ws, "KARTLAR", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Proje / arsa kartları — her kaydın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=kkh_sonrakiKimlik", kalin=True, yazi=DIS_REF_YEŞIL)
    _kim(ws, "B4", "Sonraki kimlik", "Yeni kartta kullanın")

    sutunlar = ["KartId", "Unvan", "Sehir", "ArsaM2", "SozlesmeNo", "LimitTl",
                "RiskNotu", "Aktif", "ToplamHacim", "OrtYas"]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "ToplamHacim": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'SUMIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],tblAkis[Tutar]))'
        ),
        "OrtYas": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'IFERROR(AVERAGEIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],'
            'tblAkis[YasGun]),0))'
        ),
    }
    for i, (kid, unvan, sehir, m2) in enumerate(PROJE_KARTLARI):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Kart Kimliği", mesaj="KKH-##### formatında")
        _sari(ws, r, 2, unvan, baslik="Unvan", mesaj="Proje/arsa unvanı")
        _sari(ws, r, 3, sehir, baslik="Şehir", mesaj="Şehir")
        _sari(ws, r, 4, m2, sayi=CATI, baslik="Arsa m²", mesaj="Arsa alanı m²")
        _sari(ws, r, 5, f"SZL-KKH-2026-{i + 1:03d}", baslik="Sözleşme No", mesaj="Sözleşme numarası")
        _sari(ws, r, 6, 15_000_000 + i * 2_500_000, sayi=TL, baslik="Limit", mesaj="Sözleşme limiti TL")
        _sari(ws, r, 7, "Normal", baslik="Risk Notu", mesaj="Kısa risk notu")
        _sari(ws, r, 8, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_KART, SON + 1):
        for c in range(1, 9):
            fmt = TL if c == 6 else (CATI if c == 4 else None)
            cell = h(ws, r, c, None, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=fmt)
            cell.protection = Protection(locked=False)
    tablo_ekle(ws, "tblKartlar", f"A{HDR}:J{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Kart Kimliği", mesaj="KKH-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Unvan"), (3, "Şehir"), (5, "Sözleşme"), (7, "Risk")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "decimal", "0", f"D{ILK}:D{SON}",
              baslik="Arsa m²", mesaj="Alan ≥ 0",
              hata_baslik="Alan", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"F{ILK}:F{SON}",
              baslik="Limit", mesaj="Limit ≥ 0",
              hata_baslik="Limit", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeEvetHayir", f"H{ILK}:H{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 11)})
    ws.column_dimensions["B"].width = 24


def akis(ws):
    sayfa_hazirla(ws, "AKIS", RENK, URUN_AD, son_kolon=22)
    alt_bant(ws, 3,
             "İşlem satırları — durum makinesi + yetim + geçersiz geçiş + zincir bayrakları",
             son_kolon=18)
    sutunlar = [
        "IslemId", "KaynakKartId", "Tarih", "OncekiDurum", "Durum",
        "Sozlesme", "Ruhsat", "Insaat", "Teslim", "Paylasim", "Tutar",
        "Donem", "Aciklama",
        "YasGun", "YasKova", "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    yas_f = yaslandirma_gun("tblAkis[[#This Row],[Tarih]]", "raporTarihi").lstrip("=")
    kova_f = yas_kova_formul("tblAkis[[#This Row],[YasGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblAkis[[#This Row],[KaynakKartId]]", "tblKartlar[KartId]").lstrip("=")
    gecis_birlesik = (
        'tblAkis[[#This Row],[OncekiDurum]]&"|"&tblAkis[[#This Row],[Durum]]'
    )
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    # Zincir: sözleşme→ruhsat→inşaat→teslim→paylaşım (ardışık halka sapması)
    zincir_f = (
        'IF(OR(tblAkis[[#This Row],[Sozlesme]]="",tblAkis[[#This Row],[Insaat]]=""),"",'
        'IF(OR('
        'ABS(tblAkis[[#This Row],[Ruhsat]]-tblAkis[[#This Row],[Sozlesme]])>zincirTolerans,'
        'ABS(tblAkis[[#This Row],[Insaat]]-tblAkis[[#This Row],[Ruhsat]])>zincirTolerans,'
        'ABS(tblAkis[[#This Row],[Teslim]]-tblAkis[[#This Row],[Insaat]])>zincirTolerans),'
        '"ZİNCİR KIRIK","OK"))'
    )

    form = {
        "YasGun": f'IF(tblAkis[[#This Row],[IslemId]]="","",{yas_f})',
        "YasKova": f'IF(tblAkis[[#This Row],[IslemId]]="","",{kova_f})',
        "YetimBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{yetim_f})',
        "GecisUyarisi": (
            f'IF(OR(tblAkis[[#This Row],[IslemId]]="",'
            f'tblAkis[[#This Row],[OncekiDurum]]=""),"",{gecis_f})'
        ),
        "ZincirBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{zincir_f})',
        "KayitDolu": 'IF(tblAkis[[#This Row],[IslemId]]="","",1)',
    }

    for i, s in enumerate(ORNEK_AKIS):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="İşlem Kimliği", mesaj="AKS-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Kart", mesaj="KARTLAR'daki KartId")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Olay tarihi")
        _sari(ws, r, 4, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 5, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 6, s["sozlesme"], sayi=TL, baslik="Sözleşme", mesaj="Sözleşme tutarı TL")
        _sari(ws, r, 7, s["ruhsat"], sayi=TL, baslik="Ruhsat", mesaj="Ruhsat halkası TL")
        _sari(ws, r, 8, s["insaat"], sayi=TL, baslik="İnşaat", mesaj="İnşaat halkası TL")
        _sari(ws, r, 9, s["teslim"], sayi=TL, baslik="Teslim", mesaj="Teslim halkası TL")
        _sari(ws, r, 10, s["paylasim"], sayi=TL, baslik="Paylaşım", mesaj="Paylaşım tutarı TL")
        _sari(ws, r, 11, s["tutar"], sayi=TL, baslik="Tutar", mesaj="Kuyruk tutarı TL")
        _sari(ws, r, 12, s["donem"], baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 13, s["aciklama"], baslik="Açıklama", mesaj="Kısa açıklama")

    for r in range(ILK + DEMO_AKIS, SON + 1):
        for c in range(1, 14):
            fmt = TARİH if c == 3 else (TL if c in (6, 7, 8, 9, 10, 11) else None)
            cell = h(ws, r, c, None, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=fmt)
            cell.protection = Protection(locked=False)

    tablo_ekle(ws, "tblAkis", f"A{HDR}:S{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="İşlem Kimliği", mesaj="AKS-##### girin",
              hata_baslik="Kimlik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Kaynak Kart", mesaj="Kart kimliği (boş = yetim riski)",
              hata_baslik="Kart", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"C{ILK}:C{SON}",
              baslik="Tarih", mesaj="Olay tarihi",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "list", "ListeDurumlar", f"D{ILK}:D{SON}",
              baslik="Önceki Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"E{ILK}:E{SON}",
              baslik="Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    for c, ad in [(6, "Sözleşme"), (7, "Ruhsat"), (8, "İnşaat"),
                  (9, "Teslim"), (10, "Paylaşım"), (11, "Tutar")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 20)})
    ws.column_dimensions["M"].width = 20
    ws.column_dimensions["Q"].width = 16


def kuyruklar(ws):
    sayfa_hazirla(ws, "KUYRUKLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Durum bazlı canlı kuyruk — kapalı kayıtlar düşer", son_kolon=10)
    h(ws, 5, 1, "Kuyruk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Tutar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    h(ws, 6, 1, "Sözleşme")
    h(ws, 6, 2, kuyruk_sayim("tblAkis", "SOZLESME"), sayi=CATI)
    h(ws, 6, 3, kuyruk_tutar("tblAkis", "SOZLESME"), sayi=TL)
    h(ws, 7, 1, "İnşaat")
    h(ws, 7, 2, kuyruk_sayim("tblAkis", "INSAAT"), sayi=CATI)
    h(ws, 7, 3, kuyruk_tutar("tblAkis", "INSAAT"), sayi=TL)
    h(ws, 8, 1, "Teslim")
    h(ws, 8, 2, kuyruk_sayim("tblAkis", "TESLIM"), sayi=CATI)
    h(ws, 8, 3, kuyruk_tutar("tblAkis", "TESLIM"), sayi=TL)

    h(ws, 10, 1, "Kapalı (Paylaşım+İptal) — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 1, "Paylaşım")
    h(ws, 11, 2, kuyruk_sayim("tblAkis", "PAYLASIM"), sayi=CATI)
    h(ws, 11, 3, kuyruk_tutar("tblAkis", "PAYLASIM"), sayi=TL)
    h(ws, 12, 1, "İptal")
    h(ws, 12, 2, kuyruk_sayim("tblAkis", "IPTAL"), sayi=CATI)
    h(ws, 12, 3, kuyruk_tutar("tblAkis", "IPTAL"), sayi=TL)

    h(ws, 14, 1, "Kuyruk Özeti", kalin=True)
    h(ws, 14, 2,
      '=_xlfn.TEXTJOIN(" | ",TRUE,"Sözleşme",B6,"İnşaat",B7,"Teslim",B8)',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B14:F14")

    h(ws, 16, 1, "Yaş Kova Dağılımı", kalin=True, yazi=KOYU_LACIVERT)
    for i, kova in enumerate(["0-7", "8-30", "31-60", "60+"], 17):
        h(ws, i, 1, kova)
        h(ws, i, 2, f'=COUNTIF(tblAkis[YasKova],"{kova}")', sayi=CATI)
        h(ws, i, 3, f'=SUMIF(tblAkis[YasKova],"{kova}",tblAkis[Tutar])', sayi=TL)

    h(ws, 22, 1, "Uyarı özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 1, "Yetim")
    h(ws, 23, 2, '=COUNTIF(tblAkis[YetimBayrak],"YETİM")', sayi=CATI)
    h(ws, 24, 1, "Geçersiz geçiş")
    h(ws, 24, 2, '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")', sayi=CATI)
    h(ws, 25, 1, "Zincir kırık")
    h(ws, 25, 2, '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")', sayi=CATI)

    genislik(ws, {"A": 36, "B": 14, "C": 16})


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve duyarlılık", kalin=True, boyut=13, yazi=KOYU_LACIVERT)

    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Çarpan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Müteahhit Kâr", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Kâr (karar)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(kkh_senaryoIyi+N(GIRDI!B9)*kkh_sifirCarpan,0)",
         "=IFERROR(MOTOR!G41,0)", "=IFERROR(MOTOR!G41,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D6<MOTOR!G11,"UYGUN DEĞİL",IF(IFERROR(D6/MAX(MOTOR!G22,1),0)<MOTOR!G10,"DİKKAT","UYGUN")))'),
        ("Baz", "=IFERROR(kkh_bazCarpan+N(GIRDI!B9)*kkh_sifirCarpan,kkh_bazCarpan)",
         "=IFERROR(kkh_muteahhitKar,0)", "=IFERROR(kkh_muteahhitKar,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(kkh_kararMetni,"-"))'),
        ("Kötümser", "=IFERROR(kkh_senaryoKotu+N(GIRDI!B9)*kkh_sifirCarpan,0)",
         "=IFERROR(MOTOR!G42,0)", "=IFERROR(MOTOR!G42,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D8<MOTOR!G11,"UYGUN DEĞİL",IF(IFERROR(D8/MAX(MOTOR!G22,1),0)<MOTOR!G10,"DİKKAT","UYGUN")))'),
        ("M-SEN KDV+1p", "=IFERROR(kkh_bazCarpan+N(GIRDI!B9)*kkh_sifirCarpan,kkh_bazCarpan)",
         "=IFERROR(kkh_muteahhitKar,0)", "=IFERROR(MOTOR!G43,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D9<MOTOR!G11,"UYGUN DEĞİL",IF(IFERROR(D9/MAX(MOTOR!G22,1),0)<MOTOR!G10,"DİKKAT","UYGUN")))'),
    ]
    for i, (ad, carp, mat, tas, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, mat, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, tas, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant genişliği (kötümser−iyimser)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(D6-D8,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser ",TEXT(D6,"₺ #,##0")," · Baz ",TEXT(D7,"₺ #,##0"),'
      '" · Kötümser ",TEXT(D8,"₺ #,##0")," · Bant ",TEXT(B11,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:H12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Tornado hasılat etkisi", yazi="333333")
    h(ws, 16, 2, "=IFERROR(MOTOR!G44,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado inşaat etkisi", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G45,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado pay etkisi", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G46,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 3,
      '=IFERROR(IF(B16>=MAX(B16:B18),kkh_siraBir,IF(B16=MEDIAN(B16:B18),kkh_siraIki,kkh_siraUc)),kkh_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 3,
      '=IFERROR(IF(B17>=MAX(B16:B18),kkh_siraBir,IF(B17=MEDIAN(B16:B18),kkh_siraIki,kkh_siraUc)),kkh_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 3,
      '=IFERROR(IF(B18>=MAX(B16:B18),kkh_siraBir,IF(B18=MEDIAN(B16:B18),kkh_siraIki,kkh_siraUc)),kkh_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(CONCATENATE("1)",INDEX(A16:A18,MATCH(kkh_siraBir,C16:C18,0)),'
      '" 2)",INDEX(A16:A18,MATCH(kkh_siraIki,C16:C18,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B16:B18),"₺ #,##0")," TL — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN kural değişimi kâr", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=D9", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"KDV sıkılaştırma senaryosunda müteahhit kâr ",TEXT(D9,"₺ #,##0")," TL")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı (kâr)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=MOTOR!G48", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 3, "=IFERROR(MOTOR!G26,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 4, "=MOTOR!G47", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Tahmin (AVERAGE)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1, "Seri nokta", yazi=GRİ)
    for i, v in enumerate([1, 2, 3, 4], 28):
        h(ws, i, 1, v, sayi=CATI)
        h(ws, i, 2, f"=D{5+v}", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 32, 1, "Tahmin sonraki", yazi="333333")
    h(ws, 32, 2, "=IFERROR(AVERAGE(B28:B31),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Tahmin yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Tahmin ",TEXT(B32,"₺ #,##0")," · Alt ",TEXT(B25,"₺ #,##0"),'
      '" · Üst ",TEXT(D25,"₺ #,##0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 5, 7, "KarDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([7_260_000, 6_600_000, 5_610_000, 6_100_000], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([3_300_000, 1_750_000, 2_200_000], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)
    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([7_260_000, 6_600_000, 5_610_000, 6_100_000], 28):
        h(ws, i, 4, v, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo Müteahhit Kâr"
    g1.add_data(Reference(ws, min_col=7, min_row=5, max_row=9), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=6, max_row=9))
    g1.height, g1.width = 8, 14
    ws.add_chart(g1, "F5")

    g2 = BarChart()
    g2.type = "bar"
    g2.title = "Duyarlılık Tornado"
    g2.add_data(Reference(ws, min_col=5, min_row=15, max_row=18), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=16, max_row=18))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "F18")

    g3 = LineChart()
    g3.title = "Senaryo Kâr Trendi"
    g3.add_data(Reference(ws, min_col=4, min_row=27, max_row=31), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=28, max_row=31))
    g3.height, g3.width = 8, 12
    ws.add_chart(g3, "F32")

    genislik(ws, {"A": 36, "B": 18, "C": 18, "D": 18, "E": 22, "G": 14})


def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Altın vakalar — kat karşılığı / hasılat paylaşımı", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:H3")
    baslik_satiri(ws, 5, VAKA_BASLIK)

    # V-001: net = 100M/(1+0.20) = 83_333_333.33
    # V-002: muteahhit pay = net*0.60 = 50_000_000
    # V-003: finansman = 35M*0.12*2 = 8_400_000
    # V-004: muteahhit kar = 50M - 35M - 8.4M = 6_600_000
    vakalar_data = [
        ("V-001", "Net hasılat = brüt / (1+KDV)",
         "brut=vaka_hasilat; kdv=KKH-002",
         83333333.33,
         "=IFERROR(ROUND(vaka_hasilat/(1+INDEX(tblKurallar[deger_2026],"
         "MATCH(\"KKH-002\",tblKurallar[kural_id],0))),2),0)",
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=kkh_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "KKH-002 net hasılat"),
        ("V-002", "Müteahhit pay TL = net × pay",
         "net=V-001; pay=vaka_muteahhit_pay",
         50000000,
         "=IFERROR(ROUND(vaka_hasilat/(1+INDEX(tblKurallar[deger_2026],"
         "MATCH(\"KKH-002\",tblKurallar[kural_id],0)))*vaka_muteahhit_pay,2),0)",
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=kkh_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "KKH-001 müteahhit pay"),
        ("V-003", "Finansman = inşaat × oran × (ay/12)",
         "insaat=vaka_insaat; oran=vaka_finansman; ay=vaka_sure",
         8400000,
         "=IFERROR(ROUND(vaka_insaat*vaka_finansman*(vaka_sure/12)*"
         "IF(INDEX(tblKurallar[deger_2026],MATCH(\"KKH-001\",tblKurallar[kural_id],0))>0,1,0),2),0)",
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=kkh_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Finansman yükü"),
        ("V-004", "Müteahhit kâr = pay − inşaat − finansman",
         "yorum A",
         6600000,
         "=IFERROR(ROUND("
         "vaka_hasilat/(1+INDEX(tblKurallar[deger_2026],MATCH(\"KKH-002\",tblKurallar[kural_id],0)))"
         "*vaka_muteahhit_pay-vaka_insaat-vaka_insaat*vaka_finansman*(vaka_sure/12),2),0)",
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=kkh_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "KKH-004 müteahhit kâr"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v,
              yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = TL

    formuller = {
        "fark": "=IF(tblVakalar[[#This Row],[vaka_id]]=\"\",\"\",IFERROR(tblVakalar[[#This Row],[beklenen_sonuc]]-tblVakalar[[#This Row],[hesaplanan]],0))",
        "durum": '=IF(tblVakalar[[#This Row],[vaka_id]]="","",IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=kkh_tolerans,"TUTARLI","KIRIK"),"KIRIK"))',
    }
    tablo_ekle(ws, "tblVakalar", "A5:H9", VAKA_BASLIK, formuller)

    h(ws, 12, 1, "Vaka durum özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")',
      yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 13, 1, "Vaka fark toplamı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2, "=IFERROR(SUM(tblVakalar[fark]),0)", sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 5, 10, "BeklenenDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 11, "HesapDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, (b, he) in enumerate([
        (83333333.33, 83333333.33), (50000000, 50000000),
        (8400000, 8400000), (6600000, 6600000),
    ], 6):
        h(ws, i, 10, b, sayi=TL, yazi=GRİ)
        h(ws, i, 11, he, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Vaka Beklenen vs Hesaplanan"
    g.add_data(Reference(ws, min_col=10, min_row=5, max_col=11, max_row=9), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=6, max_row=9))
    g.height, g.width = 8, 14
    ws.add_chart(g, "A15")

    genislik(ws, {"A": 12, "B": 36, "C": 36, "D": 14, "E": 14, "F": 12, "G": 12, "H": 24})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", "0F2742", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "KANIT RAPORU — Kat Karşılığı / Hasılat Paylaşımı", kalin=True, boyut=14,
      yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:F3")
    h(ws, 4, 1, "Rapor tarihi", yazi="333333")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 5, 1, "Şirket / proje", yazi="333333")
    h(ws, 5, 2, '=CONCATENATE(GIRDI!B6," / ",GIRDI!B7)', yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Hesap yılı", yazi="333333")
    h(ws, 6, 2, "=hesapYili", yazi=DIS_REF_YEŞIL)

    h(ws, 8, 1, "Girdi özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, (et, form, sayi) in enumerate([
        ("Arsa değeri", "=GIRDI!B9", TL),
        ("İnşaat maliyeti", "=GIRDI!B10", TL),
        ("Brüt hasılat", "=GIRDI!B11", TL),
        ("Müteahhit payı", "=GIRDI!B12", YÜZDE),
        ("Arsa sahibi payı", "=GIRDI!B13", YÜZDE),
    ], 9):
        h(ws, i, 1, et, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)

    h(ws, 15, 1, "Sonuç özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, (et, form, sayi) in enumerate([
        ("Net hasılat", "=IFERROR(MOTOR!G21,0)", TL),
        ("Müteahhit kâr", "=IFERROR(kkh_muteahhitKar,0)", TL),
        ("Arsa sahibi kâr", "=IFERROR(MOTOR!G27,0)", TL),
        ("Başabaş brüt", "=IFERROR(kkh_basabesBrut,0)", TL),
        ("Finansman tutarı", "=IFERROR(MOTOR!G24,0)", TL),
    ], 16):
        h(ws, i, 1, et, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)

    h(ws, 22, 1, "Müteahhit kâr (imza satırı)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=IFERROR(kkh_muteahhitKar,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 23, 1, "KARAR", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "=kkh_kararMetni", kalin=True, boyut=14, yazi=DIS_REF_YEŞIL)
    h(ws, 24, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, "=kkh_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B24:F24")
    h(ws, 25, 1, "Madde atıf", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=kkh_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B25:F25")
    h(ws, 26, 1, "Kanıt özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 2, "=kkh_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B26:F26")

    h(ws, 28, 1, "İmza", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 29, 1, "Hazırlayan: ________________", yazi="333333")
    h(ws, 30, 1, "Onaylayan: ________________", yazi="333333")
    h(ws, 32, 1,
      "Bu çıktı karar destektir; kesin hukuki/mali görüş veya sözleşme yerine geçmez. "
      "Hasılat paylaşımı yorumu tartışmalıdır (yorum A/B).",
      yazi=GRİ, boyut=9, kaydir=True)
    ws.merge_cells("A32:F32")
    baski_hazirla(ws, "A1:F34", f"{URUN_AD} · KANIT")
    genislik(ws, {"A": 32, "B": 48})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim paneli — Kat Karşılığı / Hasılat Paylaşımı", kalin=True, boyut=14,
      yazi=KOYU_LACIVERT)

    kpiler = [
        (2, "NetHasilat", "=IFERROR(MOTOR!G21,0)", TL),
        (3, "MuteahhitKar", "=IFERROR(kkh_muteahhitKar,0)", TL),
        (4, "ArsaKar", "=IFERROR(MOTOR!G27,0)", TL),
        (5, "Basabes", "=IFERROR(kkh_basabesBrut,0)", TL),
        (6, "Marj", "=IFERROR(MOTOR!G28,0)", YÜZDE),
        (7, "Guven", "=IFERROR(MOTOR!G39,0)", CATI),
        (8, "Risk", "=IFERROR(MOTOR!G40,0)", CATI),
        (9, "FinansmanTl", "=IFERROR(MOTOR!G24,0)", TL),
        (10, "Yetim", "=IFERROR(MOTOR!G55,0)", CATI),
        (11, "Zincir", "=IFERROR(MOTOR!G56,0)", CATI),
        (12, "M-SEN", "=IFERROR(kkh_kuralDegisimSenaryo,0)", TL),
        (13, "Vaka", "=IFERROR(kkh_vakaDurum,\"-\")", None),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=9, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=kkh_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=kkh_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 10, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Müteahhit kâr / net hasılat",
         "=IFERROR(IF(MOTOR!G21=0,0,MOTOR!G26/MOTOR!G21),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Müteahhit kâr / net ",TEXT(C11,"0.0%"))'),
        ("T2", "Arsa sahibi kâr / arsa değeri",
         "=IFERROR(MOTOR!G50,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Arsa getiri ",TEXT(C12,"0.0%"))'),
        ("O1", "Senaryo kâr sapması",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,STDEV.P(SENARYO!D6:D9)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo sapması ",TEXT(C13,"₺ #,##0")," TL")'),
        ("O2", "Tornado max etki",
         "=IFERROR(kkh_duyarlilikSira,\"-\")",
         "=IFERROR(kkh_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo bant",
         "=IFERROR(kkh_senaryoKarsilastirma,0)",
         "=IFERROR(kkh_yorumSenaryo,\"-\")"),
        ("O6", "Finansman / inşaat",
         "=IFERROR(MOTOR!G49,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Finansman yoğunluğu ",TEXT(C16,"0.0%"))'),
        ("O8", "Kalite skoru",
         "=IFERROR(KONTROLLER!B12,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalite ",TEXT(C17,"0")," puan")'),
        ("M", "M-SEN kâr",
         "=IFERROR(kkh_kuralDegisimSenaryo,0)",
         "=IFERROR(kkh_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin aralık",
         "=IFERROR(kkh_tahminAralik,0)",
         "=IFERROR(kkh_yorumTahmin,\"-\")"),
        ("I2", "Müteahhit kâr P90",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!D6:D9,kkh_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"P90 kâr ",TEXT(C20,"₺ #,##0")," TL")'),
    ]
    for i, (kod, ad, form, yorum) in enumerate(moduller):
        r = 11 + i
        h(ws, r, 1, kod, kalin=True, hiza="center", yazi="B08948")
        h(ws, r, 2, ad, kaydir=True)
        h(ws, r, 3, form, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[r].height = 28
        if kod in ("T1", "T2", "O6"):
            ws.cell(r, 3).number_format = YÜZDE
        elif kod != "O2":
            ws.cell(r, 3).number_format = TL

    h(ws, 23, 1, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Arsa değeri", 20_000_000),
        ("İnşaat maliyeti", 35_000_000),
        ("Brüt hasılat", 100_000_000),
        ("Net hasılat", 83_333_333),
        ("Müteahhit kâr", 6_600_000),
    ], 24):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)

    h(ws, 23, 4, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 5, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Arsa sahibi kâr", 13_333_333),
        ("Finansman tutarı", 8_400_000),
        ("Başabaş brüt", 86_800_000),
        ("İyimser kâr", 7_260_000),
        ("Kötümser kâr", 5_610_000),
    ], 24):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Paylaşım Bileşenleri"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=28), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=28))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G9")

    g2 = BarChart()
    g2.type = "col"
    g2.title = "Kâr / Başabaş"
    g2.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "G23")

    h(ws, 30, 1, "Durum", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 2, "Adet", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, adet) in enumerate([
        ("Sözleşme", 4), ("İnşaat", 6), ("Teslim", 4), ("Paylaşım", 3), ("İptal", 1),
    ], 31):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, adet, sayi=CATI, yazi=GRİ)
    g3 = PieChart()
    g3.title = "Akış Durum Dağılımı"
    g3.add_data(Reference(ws, min_col=2, min_row=30, max_row=35), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=31, max_row=35))
    g3.height, g3.width = 8, 10
    ws.add_chart(g3, "P9")

    h(ws, 30, 4, "KPI", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 5, "Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Müteahhit kâr", 6_600_000),
        ("Arsa kâr", 13_333_333),
        ("Finansman tutarı", 8_400_000),
    ], 31):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "KPI Kâr / Finansman"
    g4.add_data(Reference(ws, min_col=5, min_row=30, max_row=33), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=4, min_row=31, max_row=33))
    g4.height, g4.width = 7, 10
    ws.add_chart(g4, "P23")

    baski_hazirla(ws, "A1:F34", f"{URUN_AD} · PANO")
    genislik(ws, {"A": 28, "B": 16, "C": 18, "D": 55})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Canlı kontrol paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    kontroller_list = [
        ("Giriş tablosu boş mu?",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"BOŞ","DOLU")'),
        ("Negatif arsa?",
         '=IF(GIRDI!B9<0,"NEGATİF","TEMİZ")'),
        ("Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(kkh_yilTaban),N(hesapYili)<=N(kkh_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        ("Pay toplamı bozuk mu?",
         '=IF(IFERROR(MOTOR!G30,0)>kkh_payTolerans,"BOZUK","TEMİZ")'),
        ("Vaka kırık mı?",
         '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")'),
        ("Hasılat başabaş altında mı?",
         '=IF(IFERROR(MOTOR!G15,0)<IFERROR(MOTOR!G32,0),"AŞIYOR","NORMAL")'),
        ("Yorum A/B seçili mi?",
         '=IF(OR(GIRDI!B16="A",GIRDI!B16="B"),"TEMİZ","EKSİK")'),
        ("Motor adım sayısı",
         "=COUNTA(MOTOR!B7:B62)"),
        ("Yetim kayıt var mı?",
         '=IF(IFERROR(MOTOR!G55,0)>0,"YETİM","TEMİZ")'),
        ("Zincir kırık var mı?",
         '=IF(IFERROR(MOTOR!G56,0)>0,"KIRIK","TEMİZ")'),
    ]
    for i, (ad, form) in enumerate(kontroller_list, 5):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Canlı formül | Değiştirmeyin")

    h(ws, 16, 1, "Toplam giriş alanı", yazi="333333")
    h(ws, 16, 2, "=IFERROR(kkh_girisBeklenen+COUNTA(GIRDI!B6:B16)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Dolu giriş", yazi="333333")
    h(ws, 17, 2, "=COUNTA(GIRDI!B6:B16)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Kalite skoru (modül)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=IFERROR(ROUND(IF(OR(B16=0,B16=\"\"),0,B17/B16*100),0),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)

    h(ws, 19, 1, "Müteahhit kâr kontrol", yazi="333333")
    h(ws, 19, 2, "=IFERROR(kkh_muteahhitKar,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Başabaş brüt kontrol", yazi="333333")
    h(ws, 20, 2, "=IFERROR(kkh_basabesBrut,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 21, 1, "Karar kontrol", yazi="333333")
    h(ws, 21, 2, "=IFERROR(kkh_kararMetni,\"-\")", yazi=DIS_REF_YEŞIL)

    genislik(ws, {"A": 40, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek senaryo kütüphanesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:I3")
    h(ws, 4, 1, "Bu sayfa demo değerleri saklar; GIRDI'ye kopyalanabilir.", yazi=GRİ)
    sutunlar = ["Senaryo", "Arsa", "Insaat", "Hasilat", "MuteahhitPay", "ArsaPay",
                "SureAy", "Finansman", "Yil", "NetHasilat"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "NetHasilat": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(tblOrnek[[#This Row],[Hasilat]]/"
            "(1+INDEX(tblKurallar[deger_2026],MATCH(\"KKH-002\",tblKurallar[kural_id],0))),0))"
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:J1004", sutunlar, formuller)
    ornekler = [
        ("Demo UYGUN", 20_000_000, 35_000_000, 100_000_000, 0.60, 0.40, 24, 0.12, 2026),
        ("Dikkat marj", 20_000_000, 42_000_000, 95_000_000, 0.55, 0.45, 30, 0.14, 2026),
        ("Uygun değil zarar", 25_000_000, 55_000_000, 80_000_000, 0.50, 0.50, 36, 0.15, 2026),
        ("2025 referans", 18_000_000, 30_000_000, 90_000_000, 0.60, 0.40, 24, 0.11, 2025),
        ("Yüksek hasılat", 22_000_000, 40_000_000, 140_000_000, 0.65, 0.35, 20, 0.10, 2026),
    ]
    for i, row in enumerate(ornekler, 6):
        for k, v in enumerate(row, 1):
            ws.cell(i, k).value = v
            if k in (2, 3, 4):
                ws.cell(i, k).number_format = TL
            elif k in (5, 6, 8):
                ws.cell(i, k).number_format = YÜZDE
            elif k in (7, 9):
                ws.cell(i, k).number_format = CATI
            ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
            ws.cell(i, k).protection = Protection(locked=False)
            _kim(ws, f"{get_column_letter(k)}{i}", sutunlar[k - 1], "Örnek senaryo değeri")

    for col, ad in [("B", "Arsa"), ("C", "İnşaat"), ("D", "Hasılat")]:
        dogrulama(ws, "decimal", "0", f"{col}6:{col}1004",
                  baslik=ad, mesaj=f"{ad} değerini girin.",
                  hata_baslik="Geçersiz", hata_mesaj="Sınır dışı.",
                  isaret="between", f2="1000000000000")
    for col, ad in [("E", "Müteahhit"), ("F", "Arsa pay"), ("H", "Finansman")]:
        dogrulama(ws, "decimal", "0", f"{col}6:{col}1004",
                  baslik=ad, mesaj="0-1 oran.",
                  hata_baslik="Geçersiz", hata_mesaj="0-1.",
                  isaret="between", f2="1")

    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6, 7, 8, 9])
    h(ws, 5, 12, "NetDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([83_333_333, 79_166_667, 66_666_667, 75_000_000, 116_666_667], 6):
        h(ws, i, 12, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Örnek Senaryo Net Hasılat"
    g.add_data(Reference(ws, min_col=12, min_row=5, max_row=10), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=6, max_row=10))
    g.height, g.width = 8, 12
    ws.add_chart(g, "K5")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 13)})
    sabitle(ws, "A6")


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", "595959", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Doğrulama listeleri", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Yıllar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, y in enumerate([2025, 2026, 2027], 6):
        h(ws, i, 1, y, sayi=CATI, zemin=GIRIS_SARI)
        ws.cell(i, 1).protection = Protection(locked=False)
        _kim(ws, f"A{i}", "Yıl", "Hesap yılı listesi")
    h(ws, 5, 3, "Yorum A/B", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate(["A", "B"], 6):
        h(ws, i, 3, v, zemin=GIRIS_SARI)
        ws.cell(i, 3).protection = Protection(locked=False)
        _kim(ws, f"C{i}", "Yorum", "Hasılat paylaşımı yorumu")
    h(ws, 5, 5, "Evet/Hayır", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate(["EVET", "HAYIR"], 6):
        h(ws, i, 5, v, zemin=GIRIS_SARI)
        ws.cell(i, 5).protection = Protection(locked=False)
        _kim(ws, f"E{i}", "Seçim", "EVET veya HAYIR")
    h(ws, 5, 7, "Durumlar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, d in enumerate(DURUM_HARITASI, 6):
        h(ws, i, 7, d["durum_id"], zemin=GIRIS_SARI)
        ws.cell(i, 7).protection = Protection(locked=False)
        _kim(ws, f"G{i}", "Durum", "Durum haritası kimliği")
    h(ws, 5, 9, "İzinli Geçiş", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, g in enumerate(IZINLI_GECIS, 6):
        h(ws, i, 9, g, zemin=GIRIS_SARI)
        ws.cell(i, 9).protection = Protection(locked=False)
        _kim(ws, f"I{i}", "Geçiş", "İzinli durum geçişi")
    genislik(ws, {"A": 12, "C": 12, "E": 12, "G": 14, "I": 22})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Parametreler (tek kaynak) + durum haritası", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT)
    sutunlar = ["anahtar", "deger", "birim", "aciklama", "kaynak", "yururluk_tarihi",
                "dogrulama_tarihi", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    params = [
        ("hesapYili", DEMO["hesapYili"], "yıl", "Aktif hesap yılı", "Kullanıcı / GIRDI",
         "01.01.2025", "11.08.2026"),
        ("raporTarihi", RAPOR_TARIH, "tarih", "Rapor tarihi (sabit)", "Üretim",
         "01.01.2025", "11.08.2026"),
        ("zincirTolerans", 50000, "TL", "Zincir sapma toleransı", "İş kuralı",
         "01.01.2025", "11.08.2026"),
        ("kkh_tolerans", 0.02, "TL", "Vaka tutarlılık toleransı", "Uygulama notu",
         "01.01.2025", "11.08.2026"),
        ("kkh_payTolerans", 0.001, "oran", "Pay sapma toleransı", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("kkh_senaryoIyi", 1.10, "çarpan", "İyimser kâr çarpanı", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("kkh_senaryoKotu", 0.85, "çarpan", "Kötümser kâr çarpanı", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("kkh_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("kkh_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("kkh_tornadoHasilat", 1.05, "çarpan", "Tornado hasılat şoku", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("kkh_tornadoMaliyet", 1.05, "çarpan", "Tornado inşaat şoku", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("kkh_tornadoPay", 1.05, "çarpan", "Tornado pay şoku", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("kkh_yorumBCarpan", 0.05, "oran", "Yorum B inşaat +%5", "Hasılat yorumu",
         "01.01.2025", "11.08.2026"),
        ("kkh_kdvSikilastirma", 0.01, "oran", "M-SEN KDV +1 puan", "Senaryo motoru",
         "01.01.2025", "11.08.2026"),
        ("kkh_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik",
         "01.01.2025", "11.08.2026"),
        ("kkh_parmakIzi", "KKH-PRO-1.0.0", "metin", "Dosya parmak izi etiketi", "Üretim",
         "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv",
         "01.01.2025", "11.08.2026"),
        ("kkh_girisBeklenen", 11, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("kkh_riskDusuk", 15, "puan", "UYGUN risk skoru", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("kkh_riskOrta", 55, "puan", "DİKKAT risk skoru", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("kkh_riskYuksek", 85, "puan", "UYGUN DEĞİL risk skoru", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("kkh_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("kkh_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("kkh_bazCarpan", 1, "çarpan", "Baz senaryo çarpanı", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("kkh_sifirCarpan", 0, "çarpan", "Nötr çarpan", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("kkh_siraBir", 1, "adet", "Tornado sıra 1", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("kkh_siraIki", 2, "adet", "Tornado sıra 2", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("kkh_siraUc", 3, "adet", "Tornado sıra 3", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("kkh_olcekHedef", 20000, "satir", "Ölçek sözleşmesi", "Manda A3 Ö1",
         "01.01.2025", "11.08.2026"),
        ("vaka_hasilat", 100_000_000, "TL", "Vaka brüt hasılat", "Altın vaka",
         "01.01.2025", "11.08.2026"),
        ("vaka_insaat", 35_000_000, "TL", "Vaka inşaat", "Altın vaka",
         "01.01.2025", "11.08.2026"),
        ("vaka_muteahhit_pay", 0.60, "oran", "Vaka müteahhit pay", "Altın vaka",
         "01.01.2025", "11.08.2026"),
        ("vaka_finansman", 0.12, "oran", "Vaka finansman", "Altın vaka",
         "01.01.2025", "11.08.2026"),
        ("vaka_sure", 24, "ay", "Vaka süre", "Altın vaka",
         "01.01.2025", "11.08.2026"),
    ]
    for i, row in enumerate(params, 6):
        for k, v in enumerate(row, 1):
            sayi = None
            if k == 2:
                if row[2] == "tarih":
                    sayi = TARİH
                elif row[2] in ("oran", "çarpan"):
                    sayi = YÜZDE
                elif row[2] == "TL":
                    sayi = TL
                elif row[2] in ("yıl", "adet", "puan", "ay", "satir"):
                    sayi = CATI
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k == 2 and sayi is None and isinstance(v, (int, float)) and not isinstance(v, bool):
                ws.cell(i, k).number_format = CATI
            if k == 2:
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"B{i}", row[0], row[3])

    for i in range(6, 6 + len(params)):
        h(ws, i, 8,
          f'=IF(A{i}="","",IF(B{i}="","EKSİK","TAM"))',
          yazi=DIS_REF_YEŞIL, hiza="center")

    son_satir = 5 + len(params)

    # Durum haritası (A01/A06)
    h(ws, 6, 10, "Durum Haritası", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    dh_bas = ["durum_id", "durum_adi", "sira", "izinli_sonraki_durumlar", "kategori"]
    baslik_satiri(ws, 7, dh_bas, basla=10)
    for i, d in enumerate(DURUM_HARITASI):
        r = 8 + i
        h(ws, r, 10, d["durum_id"])
        h(ws, r, 11, d["durum_adi"])
        h(ws, r, 12, d["sira"], sayi=CATI)
        h(ws, r, 13, d["izinli_sonraki_durumlar"])
        h(ws, r, 14, d["kategori"])
    tablo_ekle(ws, "tblDurumHaritasi", "J7:N12", dh_bas)

    h(ws, 15, 10, "Sonraki kimlik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 10, "kkh_sonrakiKimlik")
    h(ws, 16, 11,
      sonraki_kimlik_formul("KKH", "tblKartlar[KartId]").replace(
          'MAX(tblKartlar[KartId])', 'COUNTA(tblKartlar[KartId])'),
      yazi=DIS_REF_YEŞIL)

    dogrulama(ws, "list", "ListeYillar", "B6",
              baslik="Hesap Yılı", mesaj="Aktif yılı seçin.",
              hata_baslik="Geçersiz", hata_mesaj="2025-2027.")
    dogrulama(ws, "decimal", "0", "B8",
              baslik="Zincir tolerans", mesaj="TL tolerans.",
              hata_baslik="Geçersiz", hata_mesaj="0 ve üzeri.",
              isaret="greaterThanOrEqual")

    genislik(ws, {"A": 28, "B": 24, "C": 10, "D": 32, "E": 22, "F": 14, "G": 14, "H": 12,
                  "J": 14, "K": 14, "L": 8, "M": 22, "N": 10})
    h(ws, son_satir + 2, 2,
      f"Sürüm {SURUM} | Şifre koruması: {SIFRE} (formül alanları)",
      yazi=GRİ, boyut=9, kaydir=True)


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar kuralları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "VERİ YOK — giriş tablosu boşsa hesap yapılmaz.",
        "UYGUN DEĞİL — müteahhit kârı eşiğin altında, pay toplamı bozuk veya hasılat başabaşın altında.",
        "DİKKAT — müteahhit marjı dikkat eşiğinin altında; pay ve maliyetleri gözden geçirin.",
        "UYGUN — pay dengesi uygun, kâr ve hasılat başabaşın üzerinde.",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 10, 1, "Belirsizlik beyanları (M05)", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari(["hasilat_paylasim_yorumu"]), 11):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 13, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2,
      "Müteahhit kâr = net hasılat × müteahhit pay − inşaat maliyeti − finansman. "
      "Standart maliyet kabulü.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B13:J13")
    h(ws, 14, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 2,
      "Yorum B, inşaat maliyetini ×1,05 sıkılaştırır (daha düşük müteahhit kârı). "
      "Dosya her iki yorumun sonucunu MOTOR'da üretir. Bu nokta tartışmalıdır.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B14:J14")

    h(ws, 16, 1, "Durum makinesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 1,
      "SOZLESME → INSAAT → TESLIM → PAYLASIM (kapalı) / IPTAL (kapalı). "
      "Geçersiz geçiş engellenmez, görünür kılınır. Zincir: Sözleşme→Ruhsat→İnşaat→Teslim→Paylaşım.",
      kaydir=True, yazi="333333")
    ws.merge_cells("A17:J17")

    h(ws, 19, 1, "Kullanım sırası", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 1,
      "GIRDI → KURALLAR → MOTOR → KARTLAR/AKIS → KUYRUKLAR → SENARYO → VAKALAR → PANO → KANIT_RAPORU.",
      kaydir=True, yazi="333333")
    ws.merge_cells("A20:J20")
    h(ws, 22, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 18, "B": 70})


def kosullu_bicimlendirme(wb):
    yesil = Font(color=NORMAL, bold=True)
    kirmizi = Font(color=KRITIK, bold=True)
    amber = Font(color="B7791F", bold=True)
    y_fill = PatternFill("solid", fgColor="E2EFDA")
    k_fill = PatternFill("solid", fgColor="FDE9E9")
    a_fill = PatternFill("solid", fgColor="FFF2CC")

    ws = wb["PANO"]
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("VERİ YOK",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("UYGUN DEĞİL",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['AND(ISNUMBER(SEARCH("UYGUN",B6)),ISERROR(SEARCH("DEĞİL",B6)))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B6))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("C4", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("C4", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("F4", CellIsRule(operator="lessThan", formula=["0.05"], font=amber))
    ws.conditional_formatting.add("G4", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("G4", CellIsRule(operator="between", formula=["50", "79"], font=amber))
    ws.conditional_formatting.add("G4", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))
    ws.conditional_formatting.add("H4", CellIsRule(operator="greaterThanOrEqual", formula=["70"], font=kirmizi))
    ws.conditional_formatting.add("H4", CellIsRule(operator="between", formula=["40", "69"], font=amber))
    ws.conditional_formatting.add("H4", CellIsRule(operator="lessThan", formula=["40"], font=yesil))
    ws.conditional_formatting.add("J4", CellIsRule(operator="greaterThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("K4", CellIsRule(operator="greaterThan", formula=["0"], font=kirmizi))
    for col in range(2, 14):
        harf = get_column_letter(col)
        ws.conditional_formatting.add(f"{harf}4", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))

    ws = wb["GIRDI"]
    for col in ("B",):
        ws.conditional_formatting.add(f"{col}9:{col}11", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(f"{col}9:{col}11", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("F25:F35", FormulaRule(formula=['F25="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F25:F35", FormulaRule(formula=['F25="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F6:F9", CellIsRule(operator="notEqual", formula=["0"], font=amber))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("UYGUN DEĞİL",E6))'], font=kirmizi))
    ws.conditional_formatting.add("E6:E9", FormulaRule(
        formula=['AND(ISNUMBER(SEARCH("UYGUN",E6)),ISERROR(SEARCH("DEĞİL",E6)))'], font=yesil))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",E6))'], font=amber))
    ws.conditional_formatting.add("B16:B18", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add(
        "B5:B14",
        FormulaRule(formula=['OR(B5="KIRIK",B5="HATALI",B5="NEGATİF",B5="BOŞ",B5="EKSİK",B5="AŞIYOR",B5="BOZUK",B5="YETİM")'],
                    font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add(
        "B5:B14",
        FormulaRule(formula=['OR(B5="TEMİZ",B5="DOLU",B5="NORMAL",B5="TUTARLI")'],
                    font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B12", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("B12", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("G7:G62", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("F7:F62", FormulaRule(formula=['F7="KKH-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F62", FormulaRule(formula=['F7="KKH-004"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "C", "D", "J"):
        ws.conditional_formatting.add(f"{harf}6:{harf}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B23", FormulaRule(formula=['ISNUMBER(SEARCH("UYGUN DEĞİL",B23))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B23", FormulaRule(
        formula=['AND(ISNUMBER(SEARCH("UYGUN",B23)),ISERROR(SEARCH("DEĞİL",B23)))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B23", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B23))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B22", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))

    ws = wb["AKIS"]
    ws.conditional_formatting.add("P6:P50", FormulaRule(formula=['P6="YETİM"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("Q6:Q50", FormulaRule(formula=['ISNUMBER(SEARCH("GEÇERSİZ",Q6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("R6:R50", FormulaRule(formula=['R6="ZİNCİR KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("R6:R50", FormulaRule(formula=['R6="OK"'], font=yesil, fill=y_fill))

    ws = wb["KUYRUKLAR"]
    ws.conditional_formatting.add("B6:B8", CellIsRule(operator="greaterThan", formula=["0"], font=amber))
    ws.conditional_formatting.add("B23:B25", CellIsRule(operator="greaterThan", formula=["0"], font=kirmizi))


def ad_tanimlari(wb):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "zincirTolerans", "AYARLAR!$B$8")

    ad_ekle(wb, "kkh_muteahhitKar", "MOTOR!$G$51")
    ad_ekle(wb, "kkh_basabesBrut", "MOTOR!$G$52")
    ad_ekle(wb, "kkh_kararMetni", "MOTOR!$G$37")
    ad_ekle(wb, "kkh_kararGerekce", "MOTOR!$G$38")
    ad_ekle(wb, "kkh_maddeAtifMetni", "MOTOR!$G$53")
    ad_ekle(wb, "kkh_kanitRaporu", "MOTOR!$G$54")
    ad_ekle(wb, "kkh_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "kkh_parmakIzi", "AYARLAR!$B$21")

    ad_ekle(wb, "kkh_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "kkh_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "kkh_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "kkh_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "kkh_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "kkh_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "kkh_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "kkh_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "kkh_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "kkh_vakaFark", "VAKALAR!$B$13")
    ad_ekle(wb, "kkh_sonrakiKimlik", "AYARLAR!$K$16")

    ayar_map = {
        "kkh_tolerans": 9,
        "kkh_payTolerans": 10,
        "kkh_senaryoIyi": 11,
        "kkh_senaryoKotu": 12,
        "kkh_tahminAlt": 13,
        "kkh_tahminUst": 14,
        "kkh_tornadoHasilat": 15,
        "kkh_tornadoMaliyet": 16,
        "kkh_tornadoPay": 17,
        "kkh_yorumBCarpan": 18,
        "kkh_kdvSikilastirma": 19,
        "kkh_yuzdelikOran": 20,
        "kkh_girisBeklenen": 23,
        "kkh_riskDusuk": 24,
        "kkh_riskOrta": 25,
        "kkh_riskYuksek": 26,
        "kkh_yilTaban": 27,
        "kkh_yuvarMax": 28,
        "kkh_bazCarpan": 29,
        "kkh_sifirCarpan": 30,
        "kkh_siraBir": 31,
        "kkh_siraIki": 32,
        "kkh_siraUc": 33,
        "kkh_olcekHedef": 34,
        "vaka_hasilat": 35,
        "vaka_insaat": 36,
        "vaka_muteahhit_pay": 37,
        "vaka_finansman": 38,
        "vaka_sure": 39,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "kkh_modulT1": (11, "C"), "kkh_modulT1Yorum": (11, "D"),
        "kkh_modulT2": (12, "C"), "kkh_modulT2Yorum": (12, "D"),
        "kkh_modulO1": (13, "C"), "kkh_modulO1Yorum": (13, "D"),
        "kkh_modulO6": (16, "C"), "kkh_modulO6Yorum": (16, "D"),
        "kkh_modulO8": (17, "C"), "kkh_modulO8Yorum": (17, "D"),
        "kkh_modulI2": (20, "C"), "kkh_modulI2Yorum": (20, "D"),
    }
    for ad, (satir, kol) in modul_map.items():
        ad_ekle(wb, ad, f"PANO!${kol}${satir}")

    ad_ekle(wb, "ListeYillar", "LISTELER!$A$6:$A$8")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$C$6:$C$7")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$6:$E$7")
    ad_ekle(wb, "ListeDurumlar", "LISTELER!$G$6:$G$10")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$I$6:$I$11")


def main(cikti_yolu=None):
    wb = Workbook()
    wb.remove(wb.active)

    siralar = [
        ("KAPAK", kapak),
        ("HIZLI_BASLANGIC", hizli_baslangic),
        ("PANO", pano),
        ("GIRDI", girdi),
        ("KURALLAR", kurallar),
        ("MOTOR", motor),
        ("KARTLAR", kartlar),
        ("AKIS", akis),
        ("KUYRUKLAR", kuyruklar),
        ("SENARYO", senaryo),
        ("VAKALAR", vakalar),
        ("KANIT_RAPORU", kanit_raporu),
        ("KONTROLLER", kontroller),
        ("ORNEK_VERI", ornek_veri),
        ("LISTELER", listeler),
        ("AYARLAR", ayarlar),
        ("KILAVUZ", kilavuz),
    ]
    for ad, fn in siralar:
        ws = wb.create_sheet(ad)
        fn(ws)

    ad_tanimlari(wb)
    kosullu_bicimlendirme(wb)
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    wb.calculation.fullCalcOnLoad = True

    kok = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    varsayilan = os.path.join(kok, "KatKarsiligiHasilatPaylasimiSimulator.xlsx")
    hedef = cikti_yolu or varsayilan
    os.makedirs(os.path.dirname(os.path.abspath(hedef)) or ".", exist_ok=True)
    wb.save(hedef)

    hsh = hashlib.sha256()
    with open(hedef, "rb") as f:
        for b in iter(lambda: f.read(65536), b""):
            hsh.update(b)
    dig = hsh.hexdigest()

    wb2 = __import__("openpyxl").load_workbook(hedef)
    wb2["AYARLAR"]["B21"] = dig[:16]
    for wsx in wb2.worksheets:
        sayfa_koru(wsx)
    wb2.save(hedef)

    print(f"Dosya: {hedef}")
    print(f"SHA-256: {dig}")
    print(f"Şifre: {SIFRE}")
    return hedef


if __name__ == "__main__":
    yol = sys.argv[1] if len(sys.argv) > 1 else None
    uretilen = main(yol)
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
    cikti = os.path.join(repo, "cikti", "KatKarsiligiHasilatPaylasimiSimulator.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
