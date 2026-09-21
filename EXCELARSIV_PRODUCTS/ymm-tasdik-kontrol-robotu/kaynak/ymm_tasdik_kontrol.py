#!/usr/bin/env python3
"""YMM Tasdik Kontrol Robotu — A4 üretim betiği (manda v6, Hat D D-03)."""

from __future__ import annotations

import hashlib
import os
import shutil
import sys
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

KOK = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
if KOK not in sys.path:
    sys.path.insert(0, KOK)

from excel_uretim.ortak import (  # noqa: E402
    CATI,
    FONT,
    GIRIS_SARI,
    GIRIS_YAZI,
    GRİ,
    GUN,
    KOYU_LACIVERT,
    NORMAL,
    ORTA_MAVI,
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
    zincir_kirik_formul,
)

URUN_AD = "YMM Tasdik Kontrol Robotu"
SURUM = "1.0.0"
RENK = "1A365D"
KAPASITE = 250
DEMO_KART = 8
DEMO_AKIS = 28
DEMO_KONTROL = 24
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 14)

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "TASLAK", "durum_adi": "Taslak", "sira": 1,
     "izinli_sonraki_durumlar": "INCELEME", "kategori": "acik"},
    {"durum_id": "INCELEME", "durum_adi": "İnceleme", "sira": 2,
     "izinli_sonraki_durumlar": "EKSIK|TASDIK|DURDUR", "kategori": "acik"},
    {"durum_id": "EKSIK", "durum_adi": "Eksik belge", "sira": 3,
     "izinli_sonraki_durumlar": "INCELEME|DURDUR", "kategori": "acik"},
    {"durum_id": "TASDIK", "durum_adi": "Tasdik", "sira": 4,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "DURDUR", "durum_adi": "Durdur", "sira": 5,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "TASLAK|INCELEME",
    "INCELEME|EKSIK",
    "INCELEME|TASDIK",
    "INCELEME|DURDUR",
    "EKSIK|INCELEME",
    "EKSIK|DURDUR",
]

MUKELLEFLER = [
    ("MUK-00001", "Demir Tekstil A.Ş.", "1234567890", "İstanbul", "KDV İadesi", 450000),
    ("MUK-00002", "Kaya Gıda Ltd.", "2345678901", "Ankara", "Yıllık Beyan", 320000),
    ("MUK-00003", "Yıldız İnşaat A.Ş.", "3456789012", "İzmir", "ÖTV Tasdik", 780000),
    ("MUK-00004", "Özkan Lojistik", "4567890123", "Bursa", "KDV İadesi", 210000),
    ("MUK-00005", "Arslan Enerji A.Ş.", "5678901234", "Antalya", "Yatırım İndirimi", 950000),
    ("MUK-00006", "Çelik Metal Ltd.", "6789012345", "Kocaeli", "KDV İadesi", 280000),
    ("MUK-00007", "Aydın Turizm A.Ş.", "7890123456", "Adana", "Yıllık Beyan", 190000),
    ("MUK-00008", "Şahin Kimya Ltd.", "8901234567", "Gaziantep", "ÖTV Tasdik", 410000),
]

BELGELER = [
    ("Başvuru dilekçesi", "KRITIK"),
    ("Vergi levhası", "ZORUNLU"),
    ("Faaliyet belgesi", "ZORUNLU"),
    ("Bilanço/gelir tablosu", "ZORUNLU"),
    ("Yevmiye dökümü", "ZORUNLU"),
    ("KDV beyan özeti", "ZORUNLU"),
    ("Banka mutabakatı", "OPSIYONEL"),
    ("Sözleşme ekleri", "OPSIYONEL"),
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: YMM tasdik kontrol ve eksik belge kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_akis():
    """≥25 satır; yetim, geçersiz geçiş ve zincir kırığı kasıtlı."""
    satirlar = []
    durumlar = ["TASLAK", "INCELEME", "EKSIK", "TASDIK", "DURDUR", "INCELEME", "TASLAK", "INCELEME"]
    oncekiler = ["", "TASLAK", "INCELEME", "INCELEME", "INCELEME", "EKSIK", "", "TASLAK"]
    for i in range(1, DEMO_AKIS + 1):
        kart = MUKELLEFLER[(i - 1) % DEMO_KART][0]
        baz = MUKELLEFLER[(i - 1) % DEMO_KART][5]
        if i in (22, 25):
            kart = "" if i == 22 else "MUK-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:
            o, d = "TASLAK", "TASDIK"
        if i == 20:
            o, d = "TASDIK", "INCELEME"
        basvuru = baz
        belge = baz if i not in (12, 19) else baz + 15000
        kontrol = belge
        tasdik = kontrol if d == "TASDIK" else 0
        arsiv = tasdik
        tar = RAPOR_TARIHI - timedelta(days=(i * 3) % 75)
        satirlar.append({
            "id": f"YTK-{i:05d}",
            "kart": kart,
            "tarih": tar,
            "onceki": o,
            "durum": d,
            "basvuru": basvuru,
            "belge": belge,
            "kontrol": kontrol,
            "tasdik": tasdik,
            "arsiv": arsiv,
            "tutar": baz,
            "donem": "2026-08",
            "aciklama": f"Tasdik kontrol kalemi {i}",
        })
    return satirlar


def _demo_kontrol():
    satirlar = []
    n = 0
    for kid, *_rest in MUKELLEFLER:
        for belge_ad, zorun in BELGELER[:3]:
            n += 1
            if n > DEMO_KONTROL:
                break
            durum_b = "YOK" if n in (1, 2, 5, 9, 13) else "VAR"
            zorun_x = "KRITIK" if n == 2 else zorun
            satirlar.append({
                "id": f"KNT-{n:05d}",
                "kart": kid,
                "belge": belge_ad,
                "zorun": zorun_x,
                "durum": durum_b,
                "tarih": RAPOR_TARIHI - timedelta(days=5 + n * 2),
                "not": f"Kontrol satırı {n}",
            })
        if n >= DEMO_KONTROL:
            break
    while len(satirlar) < DEMO_KONTROL:
        n = len(satirlar) + 1
        kid = MUKELLEFLER[(n - 1) % DEMO_KART][0]
        belge_ad, zorun = BELGELER[(n - 1) % len(BELGELER)]
        satirlar.append({
            "id": f"KNT-{n:05d}",
            "kart": "" if n == 22 else kid,
            "belge": belge_ad,
            "zorun": zorun,
            "durum": "YOK" if n % 7 == 0 else "VAR",
            "tarih": RAPOR_TARIHI - timedelta(days=n * 2),
            "not": f"Kontrol satırı {n}",
        })
    return satirlar[:DEMO_KONTROL]


ORNEK = _demo_akis()
ORNEK_KONTROL = _demo_kontrol()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "YMM tasdik dosyasını kontrol listesi, eksik belge kuyruğu ve durum makinesi ile yönetir; "
      "boş veride VERİ YOK, kritik eksikte DURDUR kararı üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Makrosuz durum makinesi: Taslak → İnceleme → Eksik / Tasdik / Durdur",
        "KONTROL_LISTE — belge bazlı zorunluluk ve eksik bayrağı",
        "Eksik belge kuyruğu yaşlandırma (ytk_eksikKuyruk)",
        "Kuyruk + yaşlandırma kovaları (0-7 / 8-30 / 31-60 / 60+)",
        "Karar kapısı: VERİ YOK / TASDIK / EKSİK / DURDUR",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Dosya sorumlusu — kontrol listesini ve belge kuyruğunu günceller",
        "YMM — TASDIK / EKSİK / DURDUR kararını verir",
        "Mükellef / denetçi — kanıt paketini alır",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) KARTLAR'a mükellef kartı girin. 2) KONTROL_LISTE'ye belgeleri işleyin. "
      "3) AKIS'e tasdik olayını yazın. 4) PANO'dan kararı ve RAPOR'dan kanıtı alın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=20)
    sabitle(ws, "A3")
    h(ws, 3, 1, "Karar Destek Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 3, 3, "Rapor Tarihi", yazi=GRİ)
    h(ws, 3, 4, "=raporTarihi", sayi=TARİH)

    h(ws, 4, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2,
      '=IF(COUNTA(tblAkis[IslemId])=0,"VERİ YOK",'
      'IF(OR(ytk_kritikEksik>0,COUNTIF(tblAkis[Tutar],"<0")>0,'
      'IFERROR(MAX(tblAkis[Tutar]),0)>ytk_azamiTutar),"DURDUR",'
      'IF(OR(ytk_eksikAdet>0,ytk_gecikenAdet>ytk_esikGeciken,'
      'ytk_yetimKayit>0,ytk_gecersizGecis>0,ytk_zincirKirik>0),"EKSİK","TASDIK")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: TASDIK | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Açık Kuyruk Adet", "=ytk_acikAdet", CATI),
        (7, "İnceleme Adet", '=COUNTIFS(tblAkis[Durum],"INCELEME",tblAkis[Tutar],">0")', CATI),
        (8, "Eksik Belge Adet", "=ytk_eksikAdet", CATI),
        (9, "Geciken Adet", "=ytk_gecikenAdet", CATI),
        (10, "Açık Tutar", "=ytk_acikTutar", TL),
        (11, "Gecikme Maliyeti", "=ytk_gecikmeMaliyeti", TL),
        (12, "Yetim Kayıt", "=ytk_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=ytk_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=ytk_zincirKirik", CATI),
        (15, "Yaş Kova Özeti", "=ytk_yasKova", None),
        (16, "Tahmin Aralık", "=ytk_tahminAralik", TL),
        (17, "Senaryo Farkı", "=ytk_senaryoKarsilastirma", TL),
        (18, "Tornado Zirve", "=ytk_tornadoZirve", TL),
        (19, "Ödeme Öncelik", "=ytk_odemeOncelik", TL),
        (20, "Pareto Gecikme", "=ytk_paretoGecikme", YÜZDE),
        (21, "Kapalı Düşüm", "=ytk_kapaliDusum", CATI),
        (22, "Eksik Kuyruk", "=ytk_eksikKuyruk", None),
        (23, "Kritik Eksik", "=ytk_kritikEksik", CATI),
        (24, "Kalite Modülü", "=ytk_modulO8", CATI),
        (25, "Yüzdelik 90", "=ytk_modulI2", TL),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=ytk_yorumYas"),
        (8, "=ytk_yorumGecikme"),
        (9, "=ytk_yorumTahmin"),
        (10, "=ytk_yorumSenaryo"),
        (11, "=ytk_yorumTornado"),
        (12, "=ytk_yorumOncelik"),
        (13, "=ytk_yorumPareto"),
        (14, "=ytk_modulO8Yorum"),
        (15, "=ytk_modulI2Yorum"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 27, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 28, 1, "Durum")
    h(ws, 28, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Taslak", 5), ("İnceleme", 9), ("Eksik", 6), ("Tasdik", 5), ("Durdur", 3),
    ], 29):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 28, 4, "Yaş Kova")
    h(ws, 28, 5, "Adet")
    for i, (ad, adet) in enumerate([
        ("0-7", 6), ("8-30", 9), ("31-60", 7), ("60+", 6),
    ], 29):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 28, 7, "Hafta")
    h(ws, 28, 8, "Tutar")
    for i in range(8):
        h(ws, 29 + i, 7, i + 1, sayi=CATI)
        h(ws, 29 + i, 8, 140000 + i * 18000, sayi=TL)

    h(ws, 28, 10, "Senaryo")
    h(ws, 28, 11, "Tutar")
    for i, (ad, t) in enumerate([
        ("Temkinli", 520000), ("Baz", 640000), ("İyimser", 760000),
    ], 29):
        h(ws, i, 10, ad)
        h(ws, i, 11, t, sayi=TL)

    c1 = PieChart()
    c1.title = "Durum Dağılımı"
    c1.add_data(Reference(ws, min_col=2, min_row=28, max_row=33), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=29, max_row=33))
    c1.width, c1.height = 10, 7
    ws.add_chart(c1, "A39")

    c2 = BarChart()
    c2.title = "Yaşlandırma Kovaları"
    c2.add_data(Reference(ws, min_col=5, min_row=28, max_row=32), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=29, max_row=32))
    c2.width, c2.height = 10, 7
    ws.add_chart(c2, "F39")

    c3 = LineChart()
    c3.title = "Haftalık Tasdik Tutarı"
    c3.add_data(Reference(ws, min_col=8, min_row=28, max_row=36), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=7, min_row=29, max_row=36))
    c3.width, c3.height = 12, 7
    ws.add_chart(c3, "A54")

    c4 = BarChart()
    c4.title = "Senaryo Karşılaştırma"
    c4.add_data(Reference(ws, min_col=11, min_row=28, max_row=31), titles_from_data=True)
    c4.set_categories(Reference(ws, min_col=10, min_row=29, max_row=31))
    c4.width, c4.height = 10, 7
    ws.add_chart(c4, "F54")

    h(ws, 28, 13, "Uyarı")
    h(ws, 28, 14, "Adet")
    for i, (ad, adet) in enumerate([
        ("Yetim", 2), ("Geçersiz", 2), ("Zincir", 2), ("Eksik", 4),
    ], 29):
        h(ws, i, 13, ad)
        h(ws, i, 14, adet, sayi=CATI)

    h(ws, 28, 16, "Ay")
    h(ws, 28, 17, "Tasdik")
    for i in range(6):
        h(ws, 29 + i, 16, i + 1, sayi=CATI)
        h(ws, 29 + i, 17, 160000 + i * 22000, sayi=TL)

    c5 = BarChart()
    c5.title = "Uyarı Türleri"
    c5.add_data(Reference(ws, min_col=14, min_row=28, max_row=32), titles_from_data=True)
    c5.set_categories(Reference(ws, min_col=13, min_row=29, max_row=32))
    c5.width, c5.height = 10, 7
    ws.add_chart(c5, "A69")

    c6 = LineChart()
    c6.title = "Aylık Tasdik"
    c6.add_data(Reference(ws, min_col=17, min_row=28, max_row=34), titles_from_data=True)
    c6.set_categories(Reference(ws, min_col=16, min_row=29, max_row=34))
    c6.width, c6.height = 10, 7
    ws.add_chart(c6, "F69")

    c7 = PieChart()
    c7.title = "Kova Payı"
    c7.add_data(Reference(ws, min_col=5, min_row=28, max_row=32), titles_from_data=True)
    c7.set_categories(Reference(ws, min_col=4, min_row=29, max_row=32))
    c7.width, c7.height = 9, 7
    ws.add_chart(c7, "A84")

    c8 = BarChart()
    c8.title = "Durum Adet"
    c8.add_data(Reference(ws, min_col=2, min_row=28, max_row=33), titles_from_data=True)
    c8.set_categories(Reference(ws, min_col=1, min_row=29, max_row=33))
    c8.width, c8.height = 10, 7
    ws.add_chart(c8, "F84")

    baski_hazirla(ws, "A1:N26", f"{URUN_AD} | Pano | {SURUM}")
    genislik(ws, {"A": 22, "B": 16, "C": 14, "D": 14, "E": 12})


def kartlar(ws):
    sayfa_hazirla(ws, "KARTLAR", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Mükellef / tasdik dosya kartları — her kaydın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=ytk_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: Kimlik disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: MUK-00009 | Risk: Elle çakışma")

    sutunlar = ["KartId", "Unvan", "VKN", "Sehir", "TasdikTuru", "DosyaTutar",
                "RiskNotu", "Aktif", "EksikBelge", "ToplamIslem", "OrtGecikme"]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "EksikBelge": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'COUNTIFS(tblKontrol[KaynakKartId],tblKartlar[[#This Row],[KartId]],'
            'tblKontrol[BelgeDurum],"YOK"))'
        ),
        "ToplamIslem": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'SUMIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],tblAkis[Tutar]))'
        ),
        "OrtGecikme": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'IFERROR(AVERAGEIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],'
            'tblAkis[YasGun]),0))'
        ),
    }
    for i, (kid, unvan, vkn, sehir, tur, tutar) in enumerate(MUKELLEFLER):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Kart Kimliği", mesaj="MUK-##### formatında")
        _sari(ws, r, 2, unvan, baslik="Unvan", mesaj="Mükellef unvanı")
        _sari(ws, r, 3, vkn, baslik="VKN", mesaj="10 haneli VKN")
        _sari(ws, r, 4, sehir, baslik="Şehir", mesaj="Şehir")
        _sari(ws, r, 5, tur, baslik="Tasdik Türü", mesaj="Tasdik konusu")
        _sari(ws, r, 6, tutar, sayi=TL, baslik="Dosya Tutarı", mesaj="Dosya tutarı TL")
        _sari(ws, r, 7, "Normal", baslik="Risk Notu", mesaj="Kısa risk notu")
        _sari(ws, r, 8, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_KART, SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None, sayi=TL if c == 6 else None,
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblKartlar", f"A{HDR}:K{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Kart Kimliği", mesaj="MUK-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Unvan"), (3, "VKN"), (4, "Şehir"), (5, "Tasdik Türü"), (7, "Risk")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "decimal", "0", f"F{ILK}:F{SON}",
              baslik="Dosya Tutarı", mesaj="Tutar ≥ 0",
              hata_baslik="Tutar", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeEvetHayir", f"H{ILK}:H{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 12)})
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["E"].width = 18


def akis(ws):
    sayfa_hazirla(ws, "AKIS", RENK, URUN_AD, son_kolon=22)
    alt_bant(ws, 3,
             "Tasdik olay satırları — durum makinesi + yetim + geçersiz geçiş + zincir bayrakları",
             son_kolon=18)
    sutunlar = [
        "IslemId", "KaynakKartId", "Tarih", "OncekiDurum", "Durum",
        "BasvuruTutar", "BelgeTutar", "KontrolTutar", "TasdikTutar", "ArsivTutar", "Tutar",
        "Donem", "Aciklama",
        "YasGun", "YasKova", "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    yas_f = yaslandirma_gun(
        "tblAkis[[#This Row],[Tarih]]", "raporTarihi").lstrip("=")
    kova_f = yas_kova_formul("tblAkis[[#This Row],[YasGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblAkis[[#This Row],[KaynakKartId]]", "tblKartlar[KartId]").lstrip("=")
    gecis_birlesik = (
        'tblAkis[[#This Row],[OncekiDurum]]&"|"&tblAkis[[#This Row],[Durum]]'
    )
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    zincir_f = zincir_kirik_formul(
        "tblAkis[[#This Row],[BasvuruTutar]]",
        "tblAkis[[#This Row],[BelgeTutar]]",
        "zincirTolerans",
    ).lstrip("=")

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

    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="İşlem Kimliği", mesaj="YTK-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Kart", mesaj="KARTLAR'daki KartId")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Olay tarihi")
        _sari(ws, r, 4, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 5, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 6, s["basvuru"], sayi=TL, baslik="Başvuru", mesaj="Başvuru halkası TL")
        _sari(ws, r, 7, s["belge"], sayi=TL, baslik="Belge", mesaj="Belge halkası TL")
        _sari(ws, r, 8, s["kontrol"], sayi=TL, baslik="Kontrol", mesaj="Kontrol halkası TL")
        _sari(ws, r, 9, s["tasdik"], sayi=TL, baslik="Tasdik", mesaj="Tasdik halkası TL")
        _sari(ws, r, 10, s["arsiv"], sayi=TL, baslik="Arşiv", mesaj="Arşiv halkası TL")
        _sari(ws, r, 11, s["tutar"], sayi=TL, baslik="Tutar", mesaj="Kuyruk tutarı TL")
        _sari(ws, r, 12, s["donem"], baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 13, s["aciklama"], baslik="Açıklama", mesaj="Kısa açıklama")

    for r in range(ILK + DEMO_AKIS, SON + 1):
        for c in range(1, 14):
            fmt = TARİH if c == 3 else (TL if c in (6, 7, 8, 9, 10, 11) else None)
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblAkis", f"A{HDR}:S{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="İşlem Kimliği", mesaj="YTK-##### girin",
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
    for c, ad in [(6, "Başvuru"), (7, "Belge"), (8, "Kontrol"),
                  (9, "Tasdik"), (10, "Arşiv"), (11, "Tutar")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "textLength", "0", f"L{ILK}:L{SON}",
              baslik="Dönem", mesaj="YYYY-AA",
              hata_baslik="Dönem", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="20")
    dogrulama(ws, "textLength", "0", f"M{ILK}:M{SON}",
              baslik="Açıklama", mesaj="Kısa açıklama",
              hata_baslik="Açıklama", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="120")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 20)})
    ws.column_dimensions["M"].width = 18
    ws.column_dimensions["Q"].width = 16


def kuyruklar(ws):
    sayfa_hazirla(ws, "KUYRUKLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Durum bazlı canlı kuyruk — kapalı kayıtlar düşer", son_kolon=10)
    h(ws, 5, 1, "Kuyruk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Tutar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    h(ws, 6, 1, "İnceleme")
    h(ws, 6, 2, kuyruk_sayim("tblAkis", "INCELEME"), sayi=CATI)
    h(ws, 6, 3, kuyruk_tutar("tblAkis", "INCELEME"), sayi=TL)

    h(ws, 7, 1, "Eksik belge")
    h(ws, 7, 2, kuyruk_sayim("tblAkis", "EKSIK"), sayi=CATI)
    h(ws, 7, 3, kuyruk_tutar("tblAkis", "EKSIK"), sayi=TL)

    h(ws, 8, 1, "Taslak")
    h(ws, 8, 2, kuyruk_sayim("tblAkis", "TASLAK"), sayi=CATI)
    h(ws, 8, 3, kuyruk_tutar("tblAkis", "TASLAK"), sayi=TL)

    h(ws, 9, 1, "Geciken (açık, yaş>eşik)")
    h(ws, 9, 2,
      '=COUNTIFS(tblAkis[YasGun],">"&ytk_esikGecikmeGun,tblAkis[Durum],"<>TASDIK",'
      'tblAkis[Durum],"<>DURDUR",tblAkis[Tutar],">0")',
      sayi=CATI)
    h(ws, 9, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[YasGun],">"&ytk_esikGecikmeGun,'
      'tblAkis[Durum],"<>TASDIK",tblAkis[Durum],"<>DURDUR")',
      sayi=TL)

    h(ws, 11, 1, "Kapalı (Tasdik+Durdur) — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 1, "Tasdik")
    h(ws, 12, 2, kuyruk_sayim("tblAkis", "TASDIK"), sayi=CATI)
    h(ws, 12, 3, kuyruk_tutar("tblAkis", "TASDIK"), sayi=TL)
    h(ws, 13, 1, "Durdur")
    h(ws, 13, 2, kuyruk_sayim("tblAkis", "DURDUR"), sayi=CATI)
    h(ws, 13, 3, kuyruk_tutar("tblAkis", "DURDUR"), sayi=TL)

    h(ws, 15, 1, "Eksik Belge Kuyruğu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 2, "=ytk_eksikKuyruk", kaydir=True)
    ws.merge_cells("B15:F15")

    h(ws, 17, 1, "Yaş Kova Dağılımı", kalin=True, yazi=KOYU_LACIVERT)
    for i, kova in enumerate(["0-7", "8-30", "31-60", "60+"], 18):
        h(ws, i, 1, kova)
        h(ws, i, 2, f'=COUNTIF(tblAkis[YasKova],"{kova}")', sayi=CATI)
        h(ws, i, 3, f'=SUMIF(tblAkis[YasKova],"{kova}",tblAkis[Tutar])', sayi=TL)

    genislik(ws, {"A": 36, "B": 14, "C": 16})


def kontrol_liste(ws):
    sayfa_hazirla(ws, "KONTROL_LISTE", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3,
             "YMM tasdik kontrol listesi — zorunluluk, eksik bayrağı, yaşlandırma",
             son_kolon=14)
    h(ws, 4, 1, "Eksik Adet", yazi=GRİ)
    h(ws, 4, 2, '=COUNTIF(tblKontrol[BelgeDurum],"YOK")', sayi=CATI, kalin=True)
    h(ws, 4, 3, "Kritik Eksik", yazi=GRİ)
    h(ws, 4, 4,
      '=COUNTIFS(tblKontrol[BelgeDurum],"YOK",tblKontrol[Zorunluluk],"KRITIK")',
      sayi=CATI, kalin=True)
    h(ws, 4, 5, "Eksik Kuyruk Özeti", yazi=GRİ)
    h(ws, 4, 6, "=ytk_eksikKuyruk", kaydir=True)

    sutunlar = [
        "KontrolId", "KaynakKartId", "BelgeAdi", "Zorunluluk", "BelgeDurum",
        "Tarih", "Not", "YasGun", "YasKova", "YetimBayrak", "EksikBayrak", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    yas_f = yaslandirma_gun(
        "tblKontrol[[#This Row],[Tarih]]", "raporTarihi").lstrip("=")
    kova_f = yas_kova_formul("tblKontrol[[#This Row],[YasGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblKontrol[[#This Row],[KaynakKartId]]", "tblKartlar[KartId]").lstrip("=")
    form = {
        "YasGun": f'IF(tblKontrol[[#This Row],[KontrolId]]="","",{yas_f})',
        "YasKova": f'IF(tblKontrol[[#This Row],[KontrolId]]="","",{kova_f})',
        "YetimBayrak": f'IF(tblKontrol[[#This Row],[KontrolId]]="","",{yetim_f})',
        "EksikBayrak": (
            'IF(tblKontrol[[#This Row],[KontrolId]]="","",'
            'IF(tblKontrol[[#This Row],[BelgeDurum]]="YOK",'
            'IF(tblKontrol[[#This Row],[Zorunluluk]]="KRITIK","KRİTİK EKSİK","EKSİK"),"TAM"))'
        ),
        "KayitDolu": 'IF(tblKontrol[[#This Row],[KontrolId]]="","",1)',
    }

    for i, s in enumerate(ORNEK_KONTROL):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="Kontrol Kimliği", mesaj="KNT-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Kart", mesaj="KARTLAR'daki KartId")
        _sari(ws, r, 3, s["belge"], baslik="Belge Adı", mesaj="Kontrol belgesi")
        _sari(ws, r, 4, s["zorun"], baslik="Zorunluluk", mesaj="KRITIK/ZORUNLU/OPSIYONEL")
        _sari(ws, r, 5, s["durum"], baslik="Belge Durum", mesaj="VAR veya YOK")
        _sari(ws, r, 6, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Kontrol tarihi")
        _sari(ws, r, 7, s["not"], baslik="Not", mesaj="Kısa not")

    for r in range(ILK + DEMO_KONTROL, SON + 1):
        for c in range(1, 8):
            fmt = TARİH if c == 6 else None
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblKontrol", f"A{HDR}:L{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Kontrol Kimliği", mesaj="KNT-##### girin",
              hata_baslik="Kimlik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Kaynak Kart", mesaj="Kart kimliği",
              hata_baslik="Kart", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "textLength", "0", f"C{ILK}:C{SON}",
              baslik="Belge Adı", mesaj="Belge adı",
              hata_baslik="Belge", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "list", "ListeZorunluluk", f"D{ILK}:D{SON}",
              baslik="Zorunluluk", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeBelgeDurum", f"E{ILK}:E{SON}",
              baslik="Belge Durum", mesaj="VAR veya YOK",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "date", "1", f"F{ILK}:F{SON}",
              baslik="Tarih", mesaj="Kontrol tarihi",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "textLength", "0", f"G{ILK}:G{SON}",
              baslik="Not", mesaj="Kısa not",
              hata_baslik="Not", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="120")

    h(ws, 4, 8, "Yaş Kova (eksik)", kalin=True, yazi=KOYU_LACIVERT)
    for i, kova in enumerate(["0-7", "8-30", "31-60", "60+"], 5):
        h(ws, i, 8, kova)
        h(ws, i, 9,
          f'=COUNTIFS(tblKontrol[YasKova],"{kova}",tblKontrol[BelgeDurum],"YOK")',
          sayi=CATI)

    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 13)})
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["G"].width = 18
    ws.column_dimensions["K"].width = 14


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Hesap motoru — ≥40 adım; sabitler AYARLAR'dan", son_kolon=10)
    h(ws, 5, 1, "Adım", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Açıklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    adimlar = [
        ("H01", "Toplam işlem", '=COUNTA(tblAkis[IslemId])'),
        ("H02", "Dolu kayıt", '=SUM(tblAkis[KayitDolu])'),
        ("H03", "Toplam tutar", "=SUM(tblAkis[Tutar])"),
        ("H04", "Başvuru toplam", "=SUM(tblAkis[BasvuruTutar])"),
        ("H05", "Belge toplam", "=SUM(tblAkis[BelgeTutar])"),
        ("H06", "Kontrol toplam", "=SUM(tblAkis[KontrolTutar])"),
        ("H07", "Tasdik toplam", "=SUM(tblAkis[TasdikTutar])"),
        ("H08", "Arşiv toplam", "=SUM(tblAkis[ArsivTutar])"),
        ("H09", "Taslak adet", '=COUNTIF(tblAkis[Durum],"TASLAK")'),
        ("H10", "İnceleme adet", '=COUNTIF(tblAkis[Durum],"INCELEME")'),
        ("H11", "Eksik adet (akis)", '=COUNTIF(tblAkis[Durum],"EKSIK")'),
        ("H12", "Tasdik adet", '=COUNTIF(tblAkis[Durum],"TASDIK")'),
        ("H13", "Durdur adet", '=COUNTIF(tblAkis[Durum],"DURDUR")'),
        ("H14", "Açık adet", "=C14+C15+C16"),
        ("H15", "Kapalı düşüm", "=C17+C18"),
        ("H16", "Açık tutar",
         '=SUMIF(tblAkis[Durum],"TASLAK",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"INCELEME",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"EKSIK",tblAkis[Tutar])'),
        ("H17", "Yetim adet", '=COUNTIF(tblAkis[YetimBayrak],"YETİM")'),
        ("H18", "Geçersiz geçiş", '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")'),
        ("H19", "Zincir kırık", '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")'),
        ("H20", "Zincir tamam", '=COUNTIF(tblAkis[ZincirBayrak],"OK")'),
        ("H21", "Yaş 0-7", '=COUNTIF(tblAkis[YasKova],"0-7")'),
        ("H22", "Yaş 8-30", '=COUNTIF(tblAkis[YasKova],"8-30")'),
        ("H23", "Yaş 31-60", '=COUNTIF(tblAkis[YasKova],"31-60")'),
        ("H24", "Yaş 60+", '=COUNTIF(tblAkis[YasKova],"60+")'),
        ("H25", "Ort yaş gün", "=IFERROR(AVERAGE(tblAkis[YasGun]),0)"),
        ("H26", "Geciken adet",
         '=COUNTIFS(tblAkis[YasGun],">"&ytk_esikGecikmeGun,tblAkis[Durum],"<>TASDIK",'
         'tblAkis[Durum],"<>DURDUR",tblAkis[Tutar],">0")'),
        ("H27", "Geciken tutar",
         '=SUMIFS(tblAkis[Tutar],tblAkis[YasGun],">"&ytk_esikGecikmeGun,'
         'tblAkis[Durum],"<>TASDIK",tblAkis[Durum],"<>DURDUR")'),
        ("H28", "Gecikme maliyeti", "=C32*ytk_gecikmeOranGun"),
        ("H29", "Temkinli senaryo", "=C21*ytk_senaryoTemkinli"),
        ("H30", "Baz senaryo", "=C21*ytk_senaryoBaz"),
        ("H31", "İyimser senaryo", "=C21*ytk_senaryoIyimser"),
        ("H32", "Senaryo farkı", "=C36-C34"),
        ("H33", "Tornado (oran×tutar)", "=C21*ytk_tornadoOran"),
        ("H34", "Ödeme öncelik havuzu",
         '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"EKSIK",tblAkis[YasGun],">"&ytk_esikGecikmeGun)'),
        ("H35", "Pareto gecikme payı", "=IFERROR(C32/MAX(C21,1),0)"),
        ("H36", "Tahmin yüzdelik",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[Tutar],ytk_yuzdelikOran),0)"),
        ("H37", "Tahmin eğilim",
         "=IFERROR(FORECAST(COUNTA(tblAkis[IslemId])+1,tblAkis[Tutar],"
         "tblAkis[KayitDolu]),C41)"),
        ("H38", "Tahmin aralık", "=IFERROR((C41+C42)/2,0)"),
        ("H39", "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        ("H40", "Dosya tutar toplam", "=SUM(tblKartlar[DosyaTutar])"),
        ("H41", "Limit kullanım oranı", "=IFERROR(C8/MAX(C45,1),0)"),
        ("H42", "Kuyruk özeti metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"İnceleme",C15,"Eksik",C16,"Geciken",C31)'),
        ("H43", "Yaş kova metin",
         '=_xlfn.TEXTJOIN("/",TRUE,C26,C27,C28,C29)'),
        ("H44", "Karar skoru ham",
         "=IF(C6=0,0,MAX(0,100-C22*10-C23*10-C24*10-C31*2))"),
        ("H45", "Veri hazırlık", "=IF(C6=0,0,C7/MAX(C6,1))"),
        ("H46", "Kontrol kayıt adet", '=COUNTA(tblKontrol[KontrolId])'),
        ("H47", "Eksik belge adet", '=COUNTIF(tblKontrol[BelgeDurum],"YOK")'),
        ("H48", "Kritik eksik adet",
         '=COUNTIFS(tblKontrol[BelgeDurum],"YOK",tblKontrol[Zorunluluk],"KRITIK")'),
        ("H49", "Tam belge adet", '=COUNTIF(tblKontrol[BelgeDurum],"VAR")'),
        ("H50", "Kontrol yetim", '=COUNTIF(tblKontrol[YetimBayrak],"YETİM")'),
        ("H51", "Kritik eksik bayrak", '=COUNTIF(tblKontrol[EksikBayrak],"KRİTİK EKSİK")'),
        ("H52", "Kalite modülü", "=C49"),
        ("H53", "Yüzdelik 90",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[YasGun],ytk_yuzdelikOran),0)"),
        ("H54", "Eksik kuyruk özet",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Eksik",C52,"Kritik",C53,"Tam",C54,"Yetim",C55)'),
        ("H55", "Kanıt özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Açık",C19,"Yetim",C22,"Zincir",C24,'
         '"Geciken",C31,"Eksik",C52,"Kritik",C53)'),
    ]
    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        fmt = TL if any(x in acik.lower() for x in (
            "tutar", "maliyet", "senaryo", "tornado", "öncelik", "tahmin", "limit", "fark",
            "toplam", "havuz", "dosya", "yüzdelik",
        )) and "oran" not in acik.lower() and "adet" not in acik.lower() else (
            YÜZDE if "oran" in acik.lower() or "payı" in acik.lower() or "hazırlık" in acik.lower()
            else (None if "metin" in acik.lower() or "özet" in acik.lower() else CATI)
        )
        if "yaş gün" in acik.lower():
            fmt = GUN
        h(ws, r, 3, form, sayi=fmt, kalin=True)

    h(ws, 64, 1, "Senaryo", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 64, 2, "Çarpan", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 64, 3, "Sonuç", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 65, 1, "Temkinli")
    h(ws, 65, 2, "=ytk_senaryoTemkinli*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 65, 3, "=C34", sayi=TL)
    h(ws, 66, 1, "Baz")
    h(ws, 66, 2, "=ytk_senaryoBaz*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 66, 3, "=C35", sayi=TL)
    h(ws, 67, 1, "İyimser")
    h(ws, 67, 2, "=ytk_senaryoIyimser*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 67, 3, "=C36", sayi=TL)

    genislik(ws, {"A": 10, "B": 28, "C": 50})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "YMM Tasdik Kontrol Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Açık Kuyruk", "=ytk_acikAdet", CATI),
        (10, "Açık Tutar", "=ytk_acikTutar", TL),
        (11, "Gecikme Maliyeti", "=ytk_gecikmeMaliyeti", TL),
        (12, "Yetim", "=ytk_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=ytk_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=ytk_zincirKirik", CATI),
        (15, "Tahmin Aralık", "=ytk_tahminAralik", TL),
        (16, "Eksik Kuyruk", "=ytk_eksikKuyruk", None),
        (17, "Kritik Eksik", "=ytk_kritikEksik", CATI),
        (18, "Kanıt Özeti", "=ytk_kanitRaporu", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 20, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 1,
      "Durum haritası AYARLAR'dadır. Geçersiz geçiş engellenmez, görünür kılınır. "
      "Eksik belge kuyruğu KONTROL_LISTE'den türetilir. "
      "Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A21:H21")
    h(ws, 23, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 1, "1. Yetim kayıtları KARTLAR'a bağlayın veya satırı düzeltin.")
    h(ws, 25, 1, "2. Geçersiz geçişleri durum haritasına uygun ilerletin.")
    h(ws, 26, 1, "3. Kritik eksik belgeleri kapatın; eksik kuyruğu yaş kovalarına göre önceliklendirin.")
    h(ws, 28, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    baski_hazirla(ws, "A1:H28", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 62, "B": 40, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "İşlem sayısı", '=COUNTA(tblAkis[IslemId])'),
        (6, "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        (7, "Taslak", '=COUNTIF(tblAkis[Durum],"TASLAK")'),
        (8, "İnceleme", '=COUNTIF(tblAkis[Durum],"INCELEME")'),
        (9, "Eksik", '=COUNTIF(tblAkis[Durum],"EKSIK")'),
        (10, "Tasdik", '=COUNTIF(tblAkis[Durum],"TASDIK")'),
        (11, "Durdur", '=COUNTIF(tblAkis[Durum],"DURDUR")'),
        (12, "Yetim", "=ytk_yetimKayit"),
        (13, "Geçersiz", "=ytk_gecersizGecis"),
        (14, "Zincir kırık", "=ytk_zincirKirik"),
        (15, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (16, "Kalite skoru", "=HESAP!C49"),
        (17, "Açık tutar", "=ytk_acikTutar"),
        (18, "Gecikme maliyeti", "=ytk_gecikmeMaliyeti"),
        (19, "Kapalı düşüm", "=ytk_kapaliDusum"),
        (20, "Eksik belge", "=ytk_eksikAdet"),
        (21, "Kritik eksik", "=ytk_kritikEksik"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, kalin=True,
          sayi=TL if r in (17, 18) else (CATI if r not in (15,) else None))
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO_KART} kart, {DEMO_KONTROL} kontrol ve {DEMO_AKIS} akış satırı vardır. "
      "Yetim (22, 25), geçersiz geçiş (18, 20) ve zincir kırık (12, 19) kasıtlıdır. "
      "Eksik belge kuyruğu KONTROL_LISTE'den türetilir. "
      "KARTLAR, AKIS ve KONTROL_LISTE'yi temizleyip kendi verinizi girebilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Durumlar: Taslak/İnceleme/Eksik/Tasdik/Durdur — kapalılar kuyruktan düşer")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 5000, sayi=CATI)
    genislik(ws, {"A": 70, "B": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Durum Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "DurumId", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, d in enumerate(DURUM_HARITASI, 7):
        _sari(ws, i, 1, d["durum_id"], baslik="Durum", mesaj="Durum kimliği")

    h(ws, 3, 3, "İzinli Geçişler", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 3, "Gecis", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, g in enumerate(IZINLI_GECIS, 7):
        _sari(ws, i, 3, g, baslik="Geçiş", mesaj="eski|yeni")

    h(ws, 3, 5, "Evet / Hayır", kalin=True)
    h(ws, 6, 5, "Secim", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    _sari(ws, 7, 5, "EVET", baslik="Seçim", mesaj="EVET/HAYIR")
    _sari(ws, 8, 5, "HAYIR", baslik="Seçim", mesaj="EVET/HAYIR")

    h(ws, 3, 7, "Zorunluluk", kalin=True)
    h(ws, 6, 7, "Tip", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(["KRITIK", "ZORUNLU", "OPSIYONEL"], 7):
        _sari(ws, i, 7, b, baslik="Zorunluluk", mesaj="Zorunluluk tipi")

    h(ws, 3, 9, "Belge Durum", kalin=True)
    h(ws, 6, 9, "Durum", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    _sari(ws, 7, 9, "VAR", baslik="Durum", mesaj="VAR/YOK")
    _sari(ws, 8, 9, "YOK", baslik="Durum", mesaj="VAR/YOK")

    h(ws, 3, 11, "Kart Kimlik", kalin=True)
    h(ws, 6, 11, "KartId", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, k in enumerate(MUKELLEFLER, 7):
        _sari(ws, i, 11, k[0], baslik="Kart", mesaj="MUK-#####")
    genislik(ws, {"A": 14, "C": 22, "E": 12, "G": 14, "I": 12, "K": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + durum haritası (A01/A06)", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("zincirTolerans", 1000, "TL", "İş kuralı — zincir sapma", "01.01.2026",
         "ZİNCİR KIRIK eşiği"),
        ("ytk_esikGecikmeGun", 30, "gün", "İç politika", "01.01.2026", "Gecikme eşiği"),
        ("ytk_esikGeciken", 5, "adet", "İç politika", "01.01.2026", "Karar eşiği geciken"),
        ("ytk_gecikmeOranGun", 0.0005, "oran", "Finans varsayımı", "01.01.2026",
         "Günlük gecikme maliyeti oranı"),
        ("ytk_senaryoTemkinli", 0.85, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("ytk_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("ytk_senaryoIyimser", 1.15, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("ytk_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("ytk_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("ytk_olcekHedef", 5000, "satir", "Manda A4 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("ytk_azamiTutar", 50000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul"),
        ("ytk_eksikEsik", 1, "adet", "Eksik belge eşiği", "01.01.2026",
         "EKSİK karar eşiği"),
        ("ytk_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi",
         "01.01.2026", "Durum makinesi kaynağı"),
        ("ytk_kuyrukOzeti_metin", "KUYRUKLAR canlı özet", "metin", "A05 kuyruk", "01.01.2026",
         "Kuyruk görünümü"),
        ("ytk_kapaliDusum_metin", "kapali kategori kuyruktan düşer", "metin", "A06", "01.01.2026",
         "Kapalı düşüm kuralı"),
        ("ytk_kanitRaporu_metin", "RAPOR kanıt çıktısı", "metin", "Dönemsel rapor", "01.01.2026",
         "Kanıt raporu"),
        ("ytk_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("ytk_ymmUnvan", "Örnek YMM", "metin", "Kullanıcı girişi", "01.01.2026",
         "YMM unvanı"),
    ]
    for i, (ana, deg, bir, kay, yur, acik) in enumerate(params):
        r = 6 + i
        h(ws, r, 1, ana)
        if isinstance(deg, date):
            _sari(ws, r, 2, deg, sayi=TARİH, baslik=ana, mesaj=acik)
        elif isinstance(deg, float) and deg <= 1:
            _sari(ws, r, 2, deg, sayi=YÜZDE, baslik=ana, mesaj=acik)
        elif isinstance(deg, (int, float)):
            _sari(ws, r, 2, deg, sayi=TL if bir == "TL" else CATI, baslik=ana, mesaj=acik)
        else:
            _sari(ws, r, 2, deg, baslik=ana, mesaj=acik)
        h(ws, r, 3, bir)
        h(ws, r, 4, kay)
        h(ws, r, 5, yur)
        h(ws, r, 6, acik, kaydir=True)

    h(ws, 28, 8, "Durum Haritası", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    dh_bas = ["durum_id", "durum_adi", "sira", "izinli_sonraki_durumlar", "kategori"]
    baslik_satiri(ws, 29, dh_bas, basla=8)
    for i, d in enumerate(DURUM_HARITASI):
        r = 30 + i
        h(ws, r, 8, d["durum_id"])
        h(ws, r, 9, d["durum_adi"])
        h(ws, r, 10, d["sira"], sayi=CATI)
        h(ws, r, 11, d["izinli_sonraki_durumlar"])
        h(ws, r, 12, d["kategori"])
    tablo_ekle(ws, "tblDurumHaritasi", "H29:L34", dh_bas)

    h(ws, 38, 8, "Motor Çıktıları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 39, 8, "anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 39, 9, "deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    motor = [
        (40, "ytk_acikAdet", "=HESAP!C19"),
        (41, "ytk_acikTutar", "=HESAP!C21"),
        (42, "ytk_gecikenAdet", "=HESAP!C31"),
        (43, "ytk_gecikmeMaliyeti", "=HESAP!C33"),
        (44, "ytk_yetimKayit", "=HESAP!C22"),
        (45, "ytk_gecersizGecis", "=HESAP!C23"),
        (46, "ytk_zincirKirik", "=HESAP!C24"),
        (47, "ytk_yasKova", "=HESAP!C48"),
        (48, "ytk_tahminAralik", "=HESAP!C43"),
        (49, "ytk_senaryoKarsilastirma", "=HESAP!C37"),
        (50, "ytk_tornadoZirve", "=HESAP!C38"),
        (51, "ytk_odemeOncelik", "=HESAP!C39"),
        (52, "ytk_paretoGecikme", "=HESAP!C40"),
        (53, "ytk_kapaliDusum", "=HESAP!C20"),
        (54, "ytk_kuyrukOzeti", "=HESAP!C47"),
        (55, "ytk_kanitRaporu", "=HESAP!C60"),
        (56, "ytk_sonrakiKimlik",
         sonraki_kimlik_formul("MUK", "tblKartlar[KartId]").replace(
             'MAX(tblKartlar[KartId])',
             'COUNTA(tblKartlar[KartId])')),
        (57, "ytk_eksikKuyruk", "=HESAP!C59"),
        (58, "ytk_modulO8", "=HESAP!C57"),
        (59, "ytk_modulI2", "=HESAP!C58"),
        (60, "ytk_eksikAdet", "=HESAP!C52"),
        (61, "ytk_kritikEksik", "=HESAP!C53"),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 63, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (64, "ytk_yorumYas",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yaş kovaları",ytk_yasKova)'),
        (65, "ytk_yorumGecikme",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Gecikme maliyeti",TEXT(ytk_gecikmeMaliyeti,"0.00"),"TL")'),
        (66, "ytk_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(ytk_tahminAralik,"0.00"),"TL")'),
        (67, "ytk_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(ytk_senaryoKarsilastirma,"0.00"),"TL")'),
        (68, "ytk_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(ytk_tornadoZirve,"0.00"),"TL")'),
        (69, "ytk_yorumOncelik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Ödeme öncelik havuzu",TEXT(ytk_odemeOncelik,"0.00"),"TL")'),
        (70, "ytk_yorumPareto",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Pareto gecikme payı",TEXT(ytk_paretoGecikme,"0.0%"))'),
        (71, "ytk_modulO8Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite",TEXT(ytk_modulO8,"0"),"yetim/geçiş/zincir düşer")'),
        (72, "ytk_modulI2Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yüzdelik 90 tutar",TEXT(ytk_modulI2,"0.00"),"TL")'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 22, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. KARTLAR sayfasına mükellef kartı ekleyin; VKN ve tasdik türünü yazın.",
        "2. KONTROL_LISTE sayfasına zorunlu/kritik belgeleri girin; VAR/YOK işaretleyin.",
        "3. AKIS sayfasına her tasdik olayını yazın; durum kolonunu listeden seçin.",
        "4. Önceki durum boş bırakılırsa geçiş uyarısı çalışmaz; bilinçli geçişte doldurun.",
        "5. Geçersiz geçiş engellenmez — kırmızı uyarı üretir (makrosuz kural).",
        "6. Kapalı durumlar (Tasdik/Durdur) açık kuyruklardan düşer.",
        "7. Kritik eksik DURDUR, eksik belge EKSİK, uygun dosya TASDIK, boş veri VERİ YOK üretir.",
        "8. RAPOR sayfasından kanıt paketini yazdırın.",
        "9. raporTarihi AYARLAR'dadır; TODAY kullanılmaz.",
        "10. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı hücreler açıktır.",
        f"11. Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv",
    ]
    for i, m in enumerate(maddeler, 5):
        h(ws, i, 1, m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 25):
        dogrulama(ws, "textLength", "0", f"D{r}",
                  baslik="Kaynak", mesaj="Kaynak metni girin",
                  hata_baslik="Kaynak", hata_mesaj="Boş bırakmayın",
                  isaret="greaterThanOrEqual", f2="120")
    ws = wb["LISTELER"]
    for r in range(7, 12):
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Durum", mesaj="Durum kimliği",
                  hata_baslik="Durum", hata_mesaj="Geçersiz",
                  isaret="greaterThan", f2="30")
    for r in range(7, 13):
        dogrulama(ws, "textLength", "1", f"C{r}",
                  baslik="Geçiş", mesaj="eski|yeni",
                  hata_baslik="Geçiş", hata_mesaj="Geçersiz",
                  isaret="greaterThan", f2="40")
    ws = wb["RAPOR"]
    for r, ad in [(5, "Sürüm"), (24, "Aksiyon1"), (25, "Aksiyon2"), (26, "Aksiyon3")]:
        dogrulama(ws, "textLength", "0", f"B{r}" if r == 5 else f"A{r}",
                  baslik=ad, mesaj="Metin alanı",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="200")
    ws = wb["KILAVUZ"]
    for r in range(5, 16):
        dogrulama(ws, "textLength", "0", f"A{r}",
                  baslik="Kılavuz", mesaj="Bilgi satırı",
                  hata_baslik="Metin", hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="200")


def kosullu_bicim(wb):
    yesil = PatternFill("solid", fgColor="C6EFCE")
    kirmizi = PatternFill("solid", fgColor="FFC7CE")
    sari = PatternFill("solid", fgColor="FFEB9C")
    yf = Font(color="006100", name=FONT)
    kf = Font(color="9C0006", name=FONT)

    ws = wb["PANO"]
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"TASDIK"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"EKSİK"'], fill=sari))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(6, 26):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["AKIS"]
    for col, val, fill, font in [
        ("P", '"YETİM"', kirmizi, kf),
        ("Q", '"GEÇERSİZ GEÇİŞ"', kirmizi, kf),
        ("R", '"ZİNCİR KIRIK"', kirmizi, kf),
        ("R", '"OK"', yesil, yf),
    ]:
        ws.conditional_formatting.add(
            f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=[val], fill=fill, font=font))
    for durum, fill in [("EKSIK", sari), ("TASDIK", yesil), ("DURDUR", kirmizi)]:
        ws.conditional_formatting.add(
            f"E{ILK}:E{SON}",
            CellIsRule(operator="equal", formula=[f'"{durum}"'], fill=fill))

    ws = wb["KUYRUKLAR"]
    for r in range(6, 10):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=sari))
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
    for r in range(18, 22):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["HESAP"]
    for r in range(6, 61):
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add(
        "B15", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B15", CellIsRule(operator="equal", formula=['"VERİ VAR"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B16", CellIsRule(operator="greaterThanOrEqual", formula=["90"], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B16", CellIsRule(operator="lessThan", formula=["70"], fill=kirmizi, font=kf))
    for r in range(5, 15):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["RAPOR"]
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"TASDIK"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"EKSİK"'], fill=sari))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(9, 19):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["KARTLAR"]
    ws.conditional_formatting.add(
        f"H{ILK}:H{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"F{ILK}:F{SON}",
        CellIsRule(operator="greaterThan", formula=["500000"], fill=yesil))

    ws = wb["KONTROL_LISTE"]
    ws.conditional_formatting.add(
        f"K{ILK}:K{SON}",
        CellIsRule(operator="equal", formula=['"KRİTİK EKSİK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"K{ILK}:K{SON}",
        CellIsRule(operator="equal", formula=['"EKSİK"'], fill=sari))
    ws.conditional_formatting.add(
        f"K{ILK}:K{SON}",
        CellIsRule(operator="equal", formula=['"TAM"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        f"E{ILK}:E{SON}",
        CellIsRule(operator="equal", formula=['"YOK"'], fill=sari))
    ws.conditional_formatting.add(
        f"J{ILK}:J{SON}",
        CellIsRule(operator="equal", formula=['"YETİM"'], fill=kirmizi, font=kf))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "zincirTolerans", "ytk_esikGecikmeGun", "ytk_esikGeciken",
        "ytk_gecikmeOranGun", "ytk_senaryoTemkinli", "ytk_senaryoBaz", "ytk_senaryoIyimser",
        "ytk_tornadoOran", "ytk_yuzdelikOran", "ytk_olcekHedef", "ytk_azamiTutar",
        "ytk_eksikEsik", "ytk_durumHaritasi", "ytk_kuyrukOzeti_metin",
        "ytk_kapaliDusum_metin", "ytk_kanitRaporu_metin", "ytk_dosyaSurumu", "ytk_ymmUnvan",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "ytk_acikAdet": 40, "ytk_acikTutar": 41, "ytk_gecikenAdet": 42,
        "ytk_gecikmeMaliyeti": 43, "ytk_yetimKayit": 44, "ytk_gecersizGecis": 45,
        "ytk_zincirKirik": 46, "ytk_yasKova": 47, "ytk_tahminAralik": 48,
        "ytk_senaryoKarsilastirma": 49, "ytk_tornadoZirve": 50, "ytk_odemeOncelik": 51,
        "ytk_paretoGecikme": 52, "ytk_kapaliDusum": 53, "ytk_kuyrukOzeti": 54,
        "ytk_kanitRaporu": 55, "ytk_sonrakiKimlik": 56, "ytk_eksikKuyruk": 57,
        "ytk_modulO8": 58, "ytk_modulI2": 59, "ytk_eksikAdet": 60, "ytk_kritikEksik": 61,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "ytk_yorumYas": 64, "ytk_yorumGecikme": 65, "ytk_yorumTahmin": 66,
        "ytk_yorumSenaryo": 67, "ytk_yorumTornado": 68, "ytk_yorumOncelik": 69,
        "ytk_yorumPareto": 70, "ytk_modulO8Yorum": 71, "ytk_modulI2Yorum": 72,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ytk_kararKapi", "PANO!$B$4")
    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$C$7:$C$12")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")
    ad_ekle(wb, "ListeZorunluluk", "LISTELER!$G$7:$G$9")
    ad_ekle(wb, "ListeBelgeDurum", "LISTELER!$I$7:$I$8")
    ad_ekle(wb, "ListeKartlar", "LISTELER!$K$7:$K$14")


def _csv_yaz(urun_dir):
    yol = os.path.join(urun_dir, "ornek_veri.csv")
    with open(yol, "w", encoding="utf-8") as f:
        f.write("IslemId,KaynakKartId,Tarih,OncekiDurum,Durum,"
                "BasvuruTutar,BelgeTutar,KontrolTutar,TasdikTutar,ArsivTutar,"
                "Tutar,Donem,Aciklama\n")
        for s in ORNEK:
            f.write(
                f"{s['id']},{s['kart']},{s['tarih']:%d.%m.%Y},"
                f"{s['onceki']},{s['durum']},{s['basvuru']},{s['belge']},"
                f"{s['kontrol']},{s['tasdik']},{s['arsiv']},{s['tutar']},"
                f"{s['donem']},{s['aciklama']}\n")
    return yol


def main(cikti_yolu=None):
    wb = Workbook()
    siralar = [
        (kapak, "KAPAK"),
        (pano, "PANO"),
        (kartlar, "KARTLAR"),
        (akis, "AKIS"),
        (kuyruklar, "KUYRUKLAR"),
        (kontrol_liste, "KONTROL_LISTE"),
        (hesap, "HESAP"),
        (rapor, "RAPOR"),
        (kontroller, "KONTROLLER"),
        (ornek_veri, "ORNEK_VERI"),
        (listeler, "LISTELER"),
        (ayarlar, "AYARLAR"),
        (kilavuz, "KILAVUZ"),
    ]
    ilk = wb.active
    for i, (fn, _ad) in enumerate(siralar):
        ws = ilk if i == 0 else wb.create_sheet()
        fn(ws)

    adlari_bagla(wb)
    _cok_dogrulama(wb)
    kosullu_bicim(wb)
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=ILK, satir_sonu=SON)

    for ws in wb.worksheets:
        sayfa_koru(ws)

    wb.calculation.fullCalcOnLoad = True
    urun_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    dosya_ad = "YmmTasdikKontrolRobotu.xlsx"
    hedef = cikti_yolu or os.path.join(urun_dir, dosya_ad)
    os.makedirs(os.path.dirname(hedef) or ".", exist_ok=True)
    wb.save(hedef)

    cikti = os.path.join(KOK, "cikti", dosya_ad)
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(hedef) != os.path.abspath(cikti):
        shutil.copy2(hedef, cikti)

    csv_yol = _csv_yaz(urun_dir)
    hsh = hashlib.sha256()
    with open(hedef, "rb") as fh:
        for blok in iter(lambda: fh.read(65536), b""):
            hsh.update(blok)
    print(f"Dosya oluşturuldu: {hedef}")
    print(f"Kopya: {cikti}")
    print(f"CSV: {csv_yol}")
    print(f"SHA-256: {hsh.hexdigest()}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
