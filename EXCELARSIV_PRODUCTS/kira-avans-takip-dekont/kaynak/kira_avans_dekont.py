#!/usr/bin/env python3
"""Kira ve Avans Takip + Dekont — A3 üretim betiği (manda v6, Hat C C-10)."""

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

URUN_AD = "Kira ve Avans Takip + Dekont"
SURUM = "1.0.0"
RENK = "0F3D4C"
KAPASITE = 250
DEMO_KART = 8
DEMO_AKIS = 28
DEMO_AVANS = 8
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 14)

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "BEKLIYOR", "durum_adi": "Bekliyor", "sira": 1,
     "izinli_sonraki_durumlar": "TAHSIL|IPTAL", "kategori": "acik"},
    {"durum_id": "TAHSIL", "durum_adi": "Tahsil", "sira": 2,
     "izinli_sonraki_durumlar": "MAHSUP|KAPALI", "kategori": "acik"},
    {"durum_id": "MAHSUP", "durum_adi": "Mahsup", "sira": 3,
     "izinli_sonraki_durumlar": "KAPALI", "kategori": "acik"},
    {"durum_id": "KAPALI", "durum_adi": "Kapalı", "sira": 4,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "IPTAL", "durum_adi": "İptal", "sira": 5,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "BEKLIYOR|TAHSIL",
    "BEKLIYOR|IPTAL",
    "TAHSIL|MAHSUP",
    "TAHSIL|KAPALI",
    "MAHSUP|KAPALI",
]

KIRACILAR = [
    ("KIR-00001", "Ayşe Demir", "Kadıköy Daire 3", "İstanbul", 18000),
    ("KIR-00002", "Mehmet Kaya", "Çankaya Ofis 12", "Ankara", 22000),
    ("KIR-00003", "Elif Yıldız", "Bornova Dükkan", "İzmir", 16000),
    ("KIR-00004", "Can Özkan", "Nilüfer Konut", "Bursa", 14000),
    ("KIR-00005", "Zeynep Arslan", "Muratpaşa Villa", "Antalya", 28000),
    ("KIR-00006", "Hakan Çelik", "İzmit Loft", "Kocaeli", 15000),
    ("KIR-00007", "Selin Aydın", "Seyhan Daire", "Adana", 12000),
    ("KIR-00008", "Burak Şahin", "Şehitkamil Dükkan", "Gaziantep", 11000),
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Kira/avans mahsup ve dekont kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_akis():
    """≥25 satır; yetim, geçersiz geçiş ve zincir kırığı kasıtlı."""
    satirlar = []
    durumlar = ["BEKLIYOR", "TAHSIL", "MAHSUP", "KAPALI", "IPTAL", "TAHSIL", "BEKLIYOR", "TAHSIL"]
    oncekiler = ["", "BEKLIYOR", "TAHSIL", "MAHSUP", "BEKLIYOR", "BEKLIYOR", "", "TAHSIL"]
    for i in range(1, DEMO_AKIS + 1):
        kart = KIRACILAR[(i - 1) % DEMO_KART][0]
        kira_ay = KIRACILAR[(i - 1) % DEMO_KART][4]
        if i in (22, 25):
            kart = "" if i == 22 else "KIR-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:
            o, d = "BEKLIYOR", "KAPALI"
        if i == 20:
            o, d = "KAPALI", "TAHSIL"
        sozlesme = kira_ay
        avans_zincir = kira_ay * 2
        kira = sozlesme if i not in (12, 19) else sozlesme + 4500
        mahsup = min(kira_ay, avans_zincir // 2) if d == "MAHSUP" else 0
        dekont = kira - mahsup if d in ("TAHSIL", "MAHSUP", "KAPALI") else 0
        tutar = kira
        tar = RAPOR_TARIHI - timedelta(days=(i * 3) % 75)
        satirlar.append({
            "id": f"KAD-{i:05d}",
            "kart": kart,
            "tarih": tar,
            "onceki": o,
            "durum": d,
            "sozlesme": sozlesme,
            "avans": avans_zincir,
            "kira": kira,
            "mahsup": mahsup,
            "dekont": dekont,
            "tutar": tutar,
            "donem": "2026-08",
            "aciklama": f"Kira tahsil kalemi {i}",
        })
    return satirlar


def _demo_avans():
    satirlar = []
    for i, (kid, _unvan, _tas, _sehir, kira) in enumerate(KIRACILAR, 1):
        tar = RAPOR_TARIHI - timedelta(days=30 + i * 4)
        satirlar.append({
            "id": f"AVN-{i:05d}",
            "kart": kid,
            "tarih": tar,
            "avans": kira * 2,
            "donem": "2026-07",
            "aciklama": f"Sözleşme avansı {i}",
        })
    return satirlar


ORNEK = _demo_akis()
ORNEK_AVANS = _demo_avans()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Kira tahsilatını avans mahsup motoru ve yazdırılabilir dekont ile yönetir; "
      "durum makinesi, yaşlandırma ve boş veride VERİ YOK kararı üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Makrosuz durum makinesi: Bekliyor → Tahsil → Mahsup → Kapalı / İptal",
        "Avans mahsup motoru — kalan bakiye ve aşım bayrağı",
        "Yazdırılabilir tahsilat dekontu (PDF baskı alanı)",
        "Kuyruk + yaşlandırma kovaları (0-7 / 8-30 / 31-60 / 60+)",
        "Karar kapısı: VERİ YOK / UYGUN / DİKKAT / KRİTİK",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Ev sahibi — kira ve avans tahsilatını işler",
        "Malik — bakiye ve mahsup kararını görür",
        "Kiracı / SMMM — yazdırılabilir dekontu alır",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) KARTLAR'a kiracı kartı girin. 2) AVANS'a depozito/avans yazın. "
      "3) AKIS'e kira olayını işleyin. 4) DEKONT'tan baskı alın, PANO'dan kararı görün.",
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
      'IF(OR(kad_gecikenAdet>kad_esikGeciken*2,COUNTIF(tblAkis[Tutar],"<0")>0,'
      'IFERROR(MAX(tblAkis[Tutar]),0)>kad_azamiTutar),"KRİTİK",'
      'IF(OR(kad_gecikenAdet>kad_esikGeciken,kad_yetimKayit>0,'
      'kad_gecersizGecis>0,kad_zincirKirik>0),"DİKKAT","UYGUN")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Açık Kuyruk Adet", "=kad_acikAdet", CATI),
        (7, "Bekliyor Adet", '=COUNTIFS(tblAkis[Durum],"BEKLIYOR",tblAkis[Tutar],">0")', CATI),
        (8, "Tahsil Adet", '=COUNTIFS(tblAkis[Durum],"TAHSIL",tblAkis[Tutar],">0")', CATI),
        (9, "Geciken Adet", "=kad_gecikenAdet", CATI),
        (10, "Açık Tutar", "=kad_acikTutar", TL),
        (11, "Gecikme Maliyeti", "=kad_gecikmeMaliyeti", TL),
        (12, "Yetim Kayıt", "=kad_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=kad_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=kad_zincirKirik", CATI),
        (15, "Yaş Kova Özeti", "=kad_yasKova", None),
        (16, "Tahmin Aralık", "=kad_tahminAralik", TL),
        (17, "Senaryo Farkı", "=kad_senaryoKarsilastirma", TL),
        (18, "Tornado Zirve", "=kad_tornadoZirve", TL),
        (19, "Ödeme Öncelik", "=kad_odemeOncelik", TL),
        (20, "Pareto Gecikme", "=kad_paretoGecikme", YÜZDE),
        (21, "Kapalı Düşüm", "=kad_kapaliDusum", CATI),
        (22, "Avans Mahsup", "=kad_avansMahsup", None),
        (23, "Dekont Özeti", "=kad_dekont", None),
        (24, "Kalite Modülü", "=kad_modulO8", CATI),
        (25, "Yüzdelik 90", "=kad_modulI2", TL),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=kad_yorumYas"),
        (8, "=kad_yorumGecikme"),
        (9, "=kad_yorumTahmin"),
        (10, "=kad_yorumSenaryo"),
        (11, "=kad_yorumTornado"),
        (12, "=kad_yorumOncelik"),
        (13, "=kad_yorumPareto"),
        (14, "=kad_modulO8Yorum"),
        (15, "=kad_modulI2Yorum"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 27, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 28, 1, "Durum")
    h(ws, 28, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Bekliyor", 6), ("Tahsil", 10), ("Mahsup", 4), ("Kapalı", 5), ("İptal", 3),
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
    h(ws, 28, 8, "Kira")
    for i in range(8):
        h(ws, 29 + i, 7, i + 1, sayi=CATI)
        h(ws, 29 + i, 8, 140000 + i * 18000, sayi=TL)

    h(ws, 28, 10, "Senaryo")
    h(ws, 28, 11, "Tutar")
    for i, (ad, t) in enumerate([
        ("Temkinli", 320000), ("Baz", 380000), ("İyimser", 440000),
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
    c3.title = "Haftalık Kira"
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
        ("Yetim", 2), ("Geçersiz", 2), ("Zincir", 2), ("Geciken", 8),
    ], 29):
        h(ws, i, 13, ad)
        h(ws, i, 14, adet, sayi=CATI)

    h(ws, 28, 16, "Ay")
    h(ws, 28, 17, "Tahsil")
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
    c6.title = "Aylık Tahsil"
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
    alt_bant(ws, 3, "Kiracı / taşınmaz kartları — her kaydın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=kad_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: Kimlik disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: KIR-00009 | Risk: Elle çakışma")

    sutunlar = ["KartId", "Unvan", "Tasinmaz", "Sehir", "SozlesmeNo", "AylikKira",
                "RiskNotu", "Aktif", "AvansBakiye", "ToplamKira", "OrtGecikme"]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "AvansBakiye": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'SUMIF(tblAvans[KaynakKartId],tblKartlar[[#This Row],[KartId]],tblAvans[KalanBakiye]))'
        ),
        "ToplamKira": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'SUMIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],tblAkis[Tutar]))'
        ),
        "OrtGecikme": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'IFERROR(AVERAGEIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],'
            'tblAkis[YasGun]),0))'
        ),
    }
    for i, (kid, unvan, tasinmaz, sehir, kira) in enumerate(KIRACILAR):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Kart Kimliği", mesaj="KIR-##### formatında")
        _sari(ws, r, 2, unvan, baslik="Unvan", mesaj="Kiracı adı soyadı")
        _sari(ws, r, 3, tasinmaz, baslik="Taşınmaz", mesaj="Daire / dükkan tanımı")
        _sari(ws, r, 4, sehir, baslik="Şehir", mesaj="Şehir")
        _sari(ws, r, 5, f"SZL-2026-{i + 1:03d}", baslik="Sözleşme No", mesaj="Sözleşme numarası")
        _sari(ws, r, 6, kira, sayi=TL, baslik="Aylık Kira", mesaj="Aylık kira TL")
        _sari(ws, r, 7, "Normal", baslik="Risk Notu", mesaj="Kısa risk notu")
        _sari(ws, r, 8, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_KART, SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None, sayi=TL if c == 6 else None,
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblKartlar", f"A{HDR}:K{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Kart Kimliği", mesaj="KIR-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Unvan"), (3, "Taşınmaz"), (4, "Şehir"), (5, "Sözleşme"), (7, "Risk")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "decimal", "0", f"F{ILK}:F{SON}",
              baslik="Aylık Kira", mesaj="Kira ≥ 0",
              hata_baslik="Kira", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeEvetHayir", f"H{ILK}:H{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 12)})
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 22


def akis(ws):
    sayfa_hazirla(ws, "AKIS", RENK, URUN_AD, son_kolon=22)
    alt_bant(ws, 3,
             "Kira olay satırları — durum makinesi + yetim + geçersiz geçiş + zincir bayrakları",
             son_kolon=18)
    sutunlar = [
        "IslemId", "KaynakKartId", "Tarih", "OncekiDurum", "Durum",
        "SozlesmeTutar", "AvansTutar", "KiraTutar", "MahsupTutar", "DekontTutar", "Tutar",
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
        "tblAkis[[#This Row],[SozlesmeTutar]]",
        "tblAkis[[#This Row],[KiraTutar]]",
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
        _sari(ws, r, 1, s["id"], baslik="İşlem Kimliği", mesaj="KAD-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Kart", mesaj="KARTLAR'daki KartId")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Olay tarihi")
        _sari(ws, r, 4, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 5, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 6, s["sozlesme"], sayi=TL, baslik="Sözleşme", mesaj="Sözleşme kira tutarı TL")
        _sari(ws, r, 7, s["avans"], sayi=TL, baslik="Avans Zincir", mesaj="Avans halkası TL")
        _sari(ws, r, 8, s["kira"], sayi=TL, baslik="Kira", mesaj="Dönem kira tutarı TL")
        _sari(ws, r, 9, s["mahsup"], sayi=TL, baslik="Mahsup", mesaj="Avans mahsup tutarı TL")
        _sari(ws, r, 10, s["dekont"], sayi=TL, baslik="Dekont", mesaj="Dekonta yazılacak net TL")
        _sari(ws, r, 11, s["tutar"], sayi=TL, baslik="Tutar", mesaj="Kuyruk tutarı TL")
        _sari(ws, r, 12, s["donem"], baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 13, s["aciklama"], baslik="Açıklama", mesaj="Kısa açıklama")

    for r in range(ILK + DEMO_AKIS, SON + 1):
        for c in range(1, 14):
            fmt = TARİH if c == 3 else (TL if c in (6, 7, 8, 9, 10, 11) else None)
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblAkis", f"A{HDR}:S{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="İşlem Kimliği", mesaj="KAD-##### girin",
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
    for c, ad in [(6, "Sözleşme"), (7, "Avans"), (8, "Kira"),
                  (9, "Mahsup"), (10, "Dekont"), (11, "Tutar")]:
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

    h(ws, 6, 1, "Bekliyor")
    h(ws, 6, 2, kuyruk_sayim("tblAkis", "BEKLIYOR"), sayi=CATI)
    h(ws, 6, 3, kuyruk_tutar("tblAkis", "BEKLIYOR"), sayi=TL)

    h(ws, 7, 1, "Tahsil")
    h(ws, 7, 2, kuyruk_sayim("tblAkis", "TAHSIL"), sayi=CATI)
    h(ws, 7, 3, kuyruk_tutar("tblAkis", "TAHSIL"), sayi=TL)

    h(ws, 8, 1, "Mahsup")
    h(ws, 8, 2, kuyruk_sayim("tblAkis", "MAHSUP"), sayi=CATI)
    h(ws, 8, 3, kuyruk_tutar("tblAkis", "MAHSUP"), sayi=TL)

    h(ws, 9, 1, "Geciken (açık, yaş>eşik)")
    h(ws, 9, 2,
      '=COUNTIFS(tblAkis[YasGun],">"&kad_esikGecikmeGun,tblAkis[Durum],"<>KAPALI",'
      'tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")',
      sayi=CATI)
    h(ws, 9, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[YasGun],">"&kad_esikGecikmeGun,'
      'tblAkis[Durum],"<>KAPALI",tblAkis[Durum],"<>IPTAL")',
      sayi=TL)

    h(ws, 11, 1, "Kapalı (Kapalı+İptal) — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 1, "Kapalı")
    h(ws, 12, 2, kuyruk_sayim("tblAkis", "KAPALI"), sayi=CATI)
    h(ws, 12, 3, kuyruk_tutar("tblAkis", "KAPALI"), sayi=TL)
    h(ws, 13, 1, "İptal")
    h(ws, 13, 2, kuyruk_sayim("tblAkis", "IPTAL"), sayi=CATI)
    h(ws, 13, 3, kuyruk_tutar("tblAkis", "IPTAL"), sayi=TL)

    h(ws, 15, 1, "Kuyruk Özeti", kalin=True)
    h(ws, 15, 2, "=kad_kuyrukOzeti", kaydir=True)
    ws.merge_cells("B15:F15")

    h(ws, 17, 1, "Yaş Kova Dağılımı", kalin=True, yazi=KOYU_LACIVERT)
    for i, kova in enumerate(["0-7", "8-30", "31-60", "60+"], 18):
        h(ws, i, 1, kova)
        h(ws, i, 2, f'=COUNTIF(tblAkis[YasKova],"{kova}")', sayi=CATI)
        h(ws, i, 3, f'=SUMIF(tblAkis[YasKova],"{kova}",tblAkis[Tutar])', sayi=TL)

    genislik(ws, {"A": 36, "B": 14, "C": 16})


def avans(ws):
    sayfa_hazirla(ws, "AVANS", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3,
             "Avans mahsup motoru — anapara, mahsup edilen, kalan bakiye, aşım bayrağı",
             son_kolon=14)
    h(ws, 4, 1, "Avans Anapara", yazi=GRİ)
    h(ws, 4, 2, "=SUM(tblAvans[AvansTutar])", sayi=TL, kalin=True)
    h(ws, 4, 3, "Mahsup Edilen", yazi=GRİ)
    h(ws, 4, 4, "=SUM(tblAvans[MahsupEdilen])", sayi=TL, kalin=True)
    h(ws, 4, 5, "Kalan Bakiye", yazi=GRİ)
    h(ws, 4, 6, "=SUM(tblAvans[KalanBakiye])", sayi=TL, kalin=True)
    h(ws, 4, 7, "Aşım Adet", yazi=GRİ)
    h(ws, 4, 8, '=COUNTIF(tblAvans[AsimBayrak],"AŞIM")', sayi=CATI, kalin=True)

    sutunlar = [
        "AvansId", "KaynakKartId", "Tarih", "AvansTutar", "Donem", "Aciklama",
        "MahsupEdilen", "KalanBakiye", "YasGun", "YasKova", "YetimBayrak", "AsimBayrak",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    yas_f = yaslandirma_gun(
        "tblAvans[[#This Row],[Tarih]]", "raporTarihi").lstrip("=")
    kova_f = yas_kova_formul("tblAvans[[#This Row],[YasGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblAvans[[#This Row],[KaynakKartId]]", "tblKartlar[KartId]").lstrip("=")
    form = {
        "MahsupEdilen": (
            'IF(tblAvans[[#This Row],[AvansId]]="","",'
            'SUMIFS(tblAkis[MahsupTutar],tblAkis[KaynakKartId],'
            'tblAvans[[#This Row],[KaynakKartId]],tblAkis[Durum],"MAHSUP"))'
        ),
        "KalanBakiye": (
            'IF(tblAvans[[#This Row],[AvansId]]="","",'
            'MAX(0,tblAvans[[#This Row],[AvansTutar]]-tblAvans[[#This Row],[MahsupEdilen]]))'
        ),
        "YasGun": f'IF(tblAvans[[#This Row],[AvansId]]="","",{yas_f})',
        "YasKova": f'IF(tblAvans[[#This Row],[AvansId]]="","",{kova_f})',
        "YetimBayrak": f'IF(tblAvans[[#This Row],[AvansId]]="","",{yetim_f})',
        "AsimBayrak": (
            'IF(tblAvans[[#This Row],[AvansId]]="","",'
            'IF(tblAvans[[#This Row],[MahsupEdilen]]>'
            'tblAvans[[#This Row],[AvansTutar]]+kad_mahsupTolerans,"AŞIM","UYGUN"))'
        ),
    }
    for i, s in enumerate(ORNEK_AVANS):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="Avans Kimliği", mesaj="AVN-#####")
        _sari(ws, r, 2, s["kart"], baslik="Kaynak Kart", mesaj="KARTLAR'daki KartId")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Avans tarihi")
        _sari(ws, r, 4, s["avans"], sayi=TL, baslik="Avans Tutarı", mesaj="Anapara TL")
        _sari(ws, r, 5, s["donem"], baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 6, s["aciklama"], baslik="Açıklama", mesaj="Kısa açıklama")
    for r in range(ILK + DEMO_AVANS, SON + 1):
        for c in range(1, 7):
            fmt = TARİH if c == 3 else (TL if c == 4 else None)
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblAvans", f"A{HDR}:L{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Avans Kimliği", mesaj="AVN-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Kaynak Kart", mesaj="Kart kimliği",
              hata_baslik="Kart", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"C{ILK}:C{SON}",
              baslik="Tarih", mesaj="Avans tarihi",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "decimal", "0", f"D{ILK}:D{SON}",
              baslik="Avans Tutarı", mesaj="Tutar ≥ 0",
              hata_baslik="Tutar", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "textLength", "0", f"E{ILK}:E{SON}",
              baslik="Dönem", mesaj="YYYY-AA",
              hata_baslik="Dönem", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="20")
    dogrulama(ws, "textLength", "0", f"F{ILK}:F{SON}",
              baslik="Açıklama", mesaj="Kısa açıklama",
              hata_baslik="Açıklama", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="120")

    h(ws, 4, 15, "Kart")
    h(ws, 4, 16, "Anapara")
    h(ws, 4, 17, "Kalan")
    for i, s in enumerate(ORNEK_AVANS):
        h(ws, 5 + i, 15, s["kart"])
        h(ws, 5 + i, 16, s["avans"], sayi=TL)
        h(ws, 5 + i, 17, max(0, s["avans"] - 8000), sayi=TL)

    c = BarChart()
    c.title = "Avans Anapara ve Kalan"
    c.add_data(Reference(ws, min_col=16, min_row=4, max_row=12), titles_from_data=True)
    c.add_data(Reference(ws, min_col=17, min_row=4, max_row=12), titles_from_data=True)
    c.set_categories(Reference(ws, min_col=15, min_row=5, max_row=12))
    c.width, c.height = 12, 7
    ws.add_chart(c, "A20")

    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 18)})
    ws.column_dimensions["F"].width = 22


def dekont(ws):
    sayfa_hazirla(ws, "DEKONT", RENK, URUN_AD, son_kolon=10)
    bant(ws, 3, "Tahsilat Dekontu — yazdırılabilir kanıt belgesi", boyut=16, son_kolon=8)
    h(ws, 4, 1, "Malik")
    h(ws, 4, 2, "=kad_malikAdi", kalin=True)
    h(ws, 5, 1, "Kart Seçimi")
    _sari(ws, 5, 2, "KIR-00001", baslik="Kart Seçimi",
          mesaj="Dekonta yazılacak kiracı kartını seçin")
    h(ws, 6, 1, "Dönem")
    _sari(ws, 6, 2, "2026-08", baslik="Dönem", mesaj="YYYY-AA")
    h(ws, 7, 1, "Rapor Tarihi")
    h(ws, 7, 2, "=raporTarihi", sayi=TARİH)

    h(ws, 9, 1, "Kiracı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 9, 2,
      '=IFERROR(INDEX(tblKartlar[Unvan],MATCH(B5,tblKartlar[KartId],0)),"")')
    h(ws, 10, 1, "Taşınmaz")
    h(ws, 10, 2,
      '=IFERROR(INDEX(tblKartlar[Tasinmaz],MATCH(B5,tblKartlar[KartId],0)),"")')
    h(ws, 11, 1, "Aylık Kira")
    h(ws, 11, 2,
      '=IFERROR(INDEX(tblKartlar[AylikKira],MATCH(B5,tblKartlar[KartId],0)),0)',
      sayi=TL)
    h(ws, 12, 1, "Dönem Kira")
    h(ws, 12, 2,
      '=IF(B5="","",SUMIFS(tblAkis[KiraTutar],tblAkis[KaynakKartId],B5,tblAkis[Donem],B6))',
      sayi=TL)
    h(ws, 13, 1, "Avans Mahsup")
    h(ws, 13, 2,
      '=IF(B5="","",SUMIFS(tblAkis[MahsupTutar],tblAkis[KaynakKartId],B5,tblAkis[Donem],B6))',
      sayi=TL)
    h(ws, 14, 1, "Net Tahsil")
    h(ws, 14, 2, "=IFERROR(B12-B13,0)", sayi=TL, kalin=True, boyut=14)
    h(ws, 15, 1, "Avans Kalan")
    h(ws, 15, 2,
      '=IF(B5="","",SUMIF(tblAvans[KaynakKartId],B5,tblAvans[KalanBakiye]))',
      sayi=TL)
    h(ws, 16, 1, "Açıklama")
    _sari(ws, 16, 2, "Dönem kira tahsilatı — avans mahsup düşülmüştür",
          baslik="Açıklama", mesaj="Dekont açıklaması")
    h(ws, 17, 1, "Teslim Alan")
    _sari(ws, 17, 2, "Ayşe Demir", baslik="Teslim Alan", mesaj="Ad soyad")

    h(ws, 19, 1, "Dekont Metni", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 1,
      '=_xlfn.TEXTJOIN(" | ",TRUE,"Kart",B5,"Dönem",B6,"Net",TEXT(B14,"0"),'
      '"Mahsup",TEXT(B13,"0"),"Kalan",TEXT(B15,"0"))',
      kaydir=True)
    ws.merge_cells("A20:H20")

    h(ws, 22, 1, "Beyan")
    h(ws, 23, 1,
      "Bu dekont karar destek belgesidir; resmi makbuz yerine geçmez. "
      "Tutarlar AKIS ve AVANS tablolarından türetilir.",
      kaydir=True)
    ws.merge_cells("A23:H23")
    h(ws, 25, 1, "Malik imza")
    h(ws, 25, 4, "Kiracı imza")
    h(ws, 27, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı", yazi=GRİ)

    dogrulama(ws, "list", "ListeKartlar", "B5",
              baslik="Kart Seçimi", mesaj="Kiracı kartını listeden seçin",
              hata_baslik="Kart", hata_mesaj="Listeden seçin")
    dogrulama(ws, "textLength", "0", "B6",
              baslik="Dönem", mesaj="YYYY-AA",
              hata_baslik="Dönem", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="20")
    dogrulama(ws, "textLength", "0", "B16",
              baslik="Açıklama", mesaj="Dekont açıklaması",
              hata_baslik="Açıklama", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="200")
    dogrulama(ws, "textLength", "0", "B17",
              baslik="Teslim Alan", mesaj="Ad soyad",
              hata_baslik="Ad", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="80")

    baski_hazirla(ws, "A1:H28", f"{URUN_AD} | Dekont | {SURUM}")
    genislik(ws, {"A": 22, "B": 48, "C": 14, "D": 18, "E": 14})


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
        ("H04", "Sözleşme toplam", "=SUM(tblAkis[SozlesmeTutar])"),
        ("H05", "Avans zincir toplam", "=SUM(tblAkis[AvansTutar])"),
        ("H06", "Kira toplam", "=SUM(tblAkis[KiraTutar])"),
        ("H07", "Mahsup akış toplam", "=SUM(tblAkis[MahsupTutar])"),
        ("H08", "Dekont toplam", "=SUM(tblAkis[DekontTutar])"),
        ("H09", "Bekliyor adet", '=COUNTIF(tblAkis[Durum],"BEKLIYOR")'),
        ("H10", "Tahsil adet", '=COUNTIF(tblAkis[Durum],"TAHSIL")'),
        ("H11", "Mahsup adet", '=COUNTIF(tblAkis[Durum],"MAHSUP")'),
        ("H12", "Kapalı adet", '=COUNTIF(tblAkis[Durum],"KAPALI")'),
        ("H13", "İptal adet", '=COUNTIF(tblAkis[Durum],"IPTAL")'),
        ("H14", "Açık adet", "=C14+C15+C16"),
        ("H15", "Kapalı düşüm", "=C17+C18"),
        ("H16", "Açık tutar",
         '=SUMIF(tblAkis[Durum],"BEKLIYOR",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"TAHSIL",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"MAHSUP",tblAkis[Tutar])'),
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
         '=COUNTIFS(tblAkis[YasGun],">"&kad_esikGecikmeGun,tblAkis[Durum],"<>KAPALI",'
         'tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")'),
        ("H27", "Geciken tutar",
         '=SUMIFS(tblAkis[Tutar],tblAkis[YasGun],">"&kad_esikGecikmeGun,'
         'tblAkis[Durum],"<>KAPALI",tblAkis[Durum],"<>IPTAL")'),
        ("H28", "Gecikme maliyeti", "=C32*kad_gecikmeOranGun"),
        ("H29", "Temkinli senaryo", "=C21*kad_senaryoTemkinli"),
        ("H30", "Baz senaryo", "=C21*kad_senaryoBaz"),
        ("H31", "İyimser senaryo", "=C21*kad_senaryoIyimser"),
        ("H32", "Senaryo farkı", "=C36-C34"),
        ("H33", "Tornado (oran×tutar)", "=C21*kad_tornadoOran"),
        ("H34", "Ödeme öncelik havuzu",
         '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"TAHSIL",tblAkis[YasGun],">"&kad_esikGecikmeGun)'),
        ("H35", "Pareto gecikme payı", "=IFERROR(C32/MAX(C21,1),0)"),
        ("H36", "Tahmin yüzdelik",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[Tutar],kad_yuzdelikOran),0)"),
        ("H37", "Tahmin eğilim",
         "=IFERROR(FORECAST(COUNTA(tblAkis[IslemId])+1,tblAkis[Tutar],"
         "tblAkis[KayitDolu]),C41)"),
        ("H38", "Tahmin aralık", "=IFERROR((C41+C42)/2,0)"),
        ("H39", "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        ("H40", "Kira limit toplam", "=SUM(tblKartlar[AylikKira])"),
        ("H41", "Limit kullanım oranı", "=IFERROR(C8/MAX(C45,1),0)"),
        ("H42", "Kuyruk özeti metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Bekliyor",C15,"Tahsil",C16,"Geciken",C31)'),
        ("H43", "Yaş kova metin",
         '=_xlfn.TEXTJOIN("/",TRUE,C26,C27,C28,C29)'),
        ("H44", "Karar skoru ham",
         "=IF(C6=0,0,MAX(0,100-C22*10-C23*10-C24*10-C31*2))"),
        ("H45", "Veri hazırlık", "=IF(C6=0,0,C7/MAX(C6,1))"),
        ("H46", "Avans kayıt adet", '=COUNTA(tblAvans[AvansId])'),
        ("H47", "Avans anapara", "=SUM(tblAvans[AvansTutar])"),
        ("H48", "Mahsup edilen", "=SUM(tblAvans[MahsupEdilen])"),
        ("H49", "Avans kalan", "=SUM(tblAvans[KalanBakiye])"),
        ("H50", "Dekont net toplam", "=SUM(tblAkis[DekontTutar])"),
        ("H51", "Mahsup aşım adet", '=COUNTIF(tblAvans[AsimBayrak],"AŞIM")'),
        ("H52", "Kalite modülü", "=C49"),
        ("H53", "Yüzdelik 90",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[YasGun],kad_yuzdelikOran),0)"),
        ("H54", "Avans mahsup özet",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Anapara",TEXT(C52,"0"),"Mahsup",TEXT(C53,"0"),'
         '"Kalan",TEXT(C54,"0"),"Aşım",C56)'),
        ("H55", "Kanıt özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Açık",C19,"Yetim",C22,"Zincir",C24,'
         '"Geciken",C31,"Mahsup kalan",TEXT(C54,"0"))'),
    ]
    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        fmt = TL if any(x in acik.lower() for x in (
            "tutar", "maliyet", "senaryo", "tornado", "öncelik", "tahmin", "limit", "fark",
            "toplam", "havuz", "anapara", "mahsup", "kalan", "kira", "dekont", "yüzdelik",
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
    h(ws, 65, 2, "=kad_senaryoTemkinli*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 65, 3, "=C34", sayi=TL)
    h(ws, 66, 1, "Baz")
    h(ws, 66, 2, "=kad_senaryoBaz*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 66, 3, "=C35", sayi=TL)
    h(ws, 67, 1, "İyimser")
    h(ws, 67, 2, "=kad_senaryoIyimser*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 67, 3, "=C36", sayi=TL)

    genislik(ws, {"A": 10, "B": 28, "C": 50})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Dönemsel Kira Avans ve Dekont Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Açık Kuyruk", "=kad_acikAdet", CATI),
        (10, "Açık Tutar", "=kad_acikTutar", TL),
        (11, "Gecikme Maliyeti", "=kad_gecikmeMaliyeti", TL),
        (12, "Yetim", "=kad_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=kad_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=kad_zincirKirik", CATI),
        (15, "Tahmin Aralık", "=kad_tahminAralik", TL),
        (16, "Avans Mahsup", "=kad_avansMahsup", None),
        (17, "Dekont Özeti", "=kad_dekont", None),
        (18, "Kanıt Özeti", "=kad_kanitRaporu", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 20, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 1,
      "Durum haritası AYARLAR'dadır. Geçersiz geçiş engellenmez, görünür kılınır. "
      "Avans mahsup kart bazında SUMIFS ile türetilir. "
      "Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A21:H21")
    h(ws, 23, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 1, "1. Yetim kayıtları KARTLAR'a bağlayın veya satırı düzeltin.")
    h(ws, 25, 1, "2. Geçersiz geçişleri durum haritasına uygun ilerletin.")
    h(ws, 26, 1, "3. Avans aşım bayraklarını mahsup tutarı ile kapatın; DEKONT'tan baskı alın.")
    h(ws, 28, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    baski_hazirla(ws, "A1:H28", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 62, "B": 40, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "İşlem sayısı", '=COUNTA(tblAkis[IslemId])'),
        (6, "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        (7, "Bekliyor", '=COUNTIF(tblAkis[Durum],"BEKLIYOR")'),
        (8, "Tahsil", '=COUNTIF(tblAkis[Durum],"TAHSIL")'),
        (9, "Mahsup", '=COUNTIF(tblAkis[Durum],"MAHSUP")'),
        (10, "Kapalı", '=COUNTIF(tblAkis[Durum],"KAPALI")'),
        (11, "İptal", '=COUNTIF(tblAkis[Durum],"IPTAL")'),
        (12, "Yetim", "=kad_yetimKayit"),
        (13, "Geçersiz", "=kad_gecersizGecis"),
        (14, "Zincir kırık", "=kad_zincirKirik"),
        (15, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (16, "Kalite skoru", "=HESAP!C49"),
        (17, "Açık tutar", "=kad_acikTutar"),
        (18, "Gecikme maliyeti", "=kad_gecikmeMaliyeti"),
        (19, "Kapalı düşüm", "=kad_kapaliDusum"),
        (20, "Avans kalan", "=SUM(tblAvans[KalanBakiye])"),
        (21, "Avans adet", '=COUNTA(tblAvans[AvansId])'),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, kalin=True,
          sayi=TL if r in (17, 18, 20) else (CATI if r not in (15,) else None))
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO_KART} kart, {DEMO_AVANS} avans ve {DEMO_AKIS} akış satırı vardır. "
      "Yetim (22, 25), geçersiz geçiş (18, 20) ve zincir kırık (12, 19) kasıtlıdır. "
      "Avans mahsup kart bazında toplanır; DEKONT yazdırılabilir. "
      "KARTLAR, AKIS ve AVANS'ı temizleyip kendi verinizi girebilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Durumlar: Bekliyor/Tahsil/Mahsup/Kapalı/İptal — kapalılar kuyruktan düşer")
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

    h(ws, 3, 7, "Taşınmaz Tip", kalin=True)
    h(ws, 6, 7, "Tip", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(["Daire", "Dükkan", "Ofis", "Villa", "Loft", "Depo"], 7):
        _sari(ws, i, 7, b, baslik="Tip", mesaj="Taşınmaz tipi")

    h(ws, 3, 9, "Kart Kimlik", kalin=True)
    h(ws, 6, 9, "KartId", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, k in enumerate(KIRACILAR, 7):
        _sari(ws, i, 9, k[0], baslik="Kart", mesaj="KIR-#####")
    genislik(ws, {"A": 14, "C": 20, "E": 12, "G": 14, "I": 14})


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
        ("kad_esikGecikmeGun", 30, "gün", "İç politika", "01.01.2026", "Gecikme eşiği"),
        ("kad_esikGeciken", 5, "adet", "İç politika", "01.01.2026", "Karar eşiği geciken"),
        ("kad_gecikmeOranGun", 0.0005, "oran", "Finans varsayımı", "01.01.2026",
         "Günlük gecikme maliyeti oranı"),
        ("kad_senaryoTemkinli", 0.85, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("kad_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("kad_senaryoIyimser", 1.15, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("kad_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("kad_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("kad_olcekHedef", 5000, "satir", "Manda A3 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("kad_azamiTutar", 50000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul"),
        ("kad_mahsupTolerans", 100, "TL", "Avans mahsup eşiği", "01.01.2026",
         "Aşım = mahsup > anapara + tolerans"),
        ("kad_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi",
         "01.01.2026", "Durum makinesi kaynağı"),
        ("kad_kuyrukOzeti_metin", "KUYRUKLAR canlı özet", "metin", "A05 kuyruk", "01.01.2026",
         "Kuyruk görünümü"),
        ("kad_kapaliDusum_metin", "kapali kategori kuyruktan düşer", "metin", "A06", "01.01.2026",
         "Kapalı düşüm kuralı"),
        ("kad_kanitRaporu_metin", "RAPOR kanıt çıktısı", "metin", "Dönemsel rapor", "01.01.2026",
         "Kanıt raporu"),
        ("kad_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("kad_malikAdi", "Örnek Malik", "metin", "Kullanıcı girişi", "01.01.2026",
         "Malik unvanı"),
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
        (40, "kad_acikAdet", "=HESAP!C19"),
        (41, "kad_acikTutar", "=HESAP!C21"),
        (42, "kad_gecikenAdet", "=HESAP!C31"),
        (43, "kad_gecikmeMaliyeti", "=HESAP!C33"),
        (44, "kad_yetimKayit", "=HESAP!C22"),
        (45, "kad_gecersizGecis", "=HESAP!C23"),
        (46, "kad_zincirKirik", "=HESAP!C24"),
        (47, "kad_yasKova", "=HESAP!C48"),
        (48, "kad_tahminAralik", "=HESAP!C43"),
        (49, "kad_senaryoKarsilastirma", "=HESAP!C37"),
        (50, "kad_tornadoZirve", "=HESAP!C38"),
        (51, "kad_odemeOncelik", "=HESAP!C39"),
        (52, "kad_paretoGecikme", "=HESAP!C40"),
        (53, "kad_kapaliDusum", "=HESAP!C20"),
        (54, "kad_kuyrukOzeti", "=HESAP!C47"),
        (55, "kad_kanitRaporu", "=HESAP!C60"),
        (56, "kad_sonrakiKimlik",
         sonraki_kimlik_formul("KIR", "tblKartlar[KartId]").replace(
             'MAX(tblKartlar[KartId])',
             'COUNTA(tblKartlar[KartId])')),
        (57, "kad_avansMahsup", "=HESAP!C59"),
        (58, "kad_modulO8", "=HESAP!C57"),
        (59, "kad_modulI2", "=HESAP!C58"),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 62, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (63, "kad_yorumYas",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yaş kovaları",kad_yasKova)'),
        (64, "kad_yorumGecikme",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Gecikme maliyeti",TEXT(kad_gecikmeMaliyeti,"0.00"),"TL")'),
        (65, "kad_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(kad_tahminAralik,"0.00"),"TL")'),
        (66, "kad_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(kad_senaryoKarsilastirma,"0.00"),"TL")'),
        (67, "kad_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(kad_tornadoZirve,"0.00"),"TL")'),
        (68, "kad_yorumOncelik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Ödeme öncelik havuzu",TEXT(kad_odemeOncelik,"0.00"),"TL")'),
        (69, "kad_yorumPareto",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Pareto gecikme payı",TEXT(kad_paretoGecikme,"0.0%"))'),
        (70, "kad_modulO8Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite",TEXT(kad_modulO8,"0"),"yetim/geçiş/zincir düşer")'),
        (71, "kad_modulI2Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yüzdelik 90 tutar",TEXT(kad_modulI2,"0.00"),"TL")'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 22, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. KARTLAR sayfasına kiracı kartı ekleyin; aylık kira ve taşınmazı yazın.",
        "2. AVANS sayfasına depozito/avans anaparasını girin; mahsup otomatik toplanır.",
        "3. AKIS sayfasına her kira olayını yazın; durum kolonunu listeden seçin.",
        "4. Önceki durum boş bırakılırsa geçiş uyarısı çalışmaz; bilinçli geçişte doldurun.",
        "5. Geçersiz geçiş engellenmez — kırmızı uyarı üretir (makrosuz kural).",
        "6. Kapalı durumlar (Kapalı/İptal) açık kuyruklardan düşer.",
        "7. Gecikme eşiği aşımı KRİTİK, yetim/geçiş DİKKAT, boş veri VERİ YOK üretir.",
        "8. DEKONT sayfasında kart ve dönem seçin; PDF olarak yazdırın.",
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
    for r in range(7, 12):
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
        "B4", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"DİKKAT"'], fill=sari))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"KRİTİK"'], fill=kirmizi, font=kf))
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
    for durum, fill in [("BEKLIYOR", sari), ("TAHSIL", yesil), ("IPTAL", kirmizi)]:
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
        "B7", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"DİKKAT"'], fill=sari))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"KRİTİK"'], fill=kirmizi, font=kf))
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
        CellIsRule(operator="greaterThan", formula=["20000"], fill=yesil))

    ws = wb["AVANS"]
    ws.conditional_formatting.add(
        f"L{ILK}:L{SON}",
        CellIsRule(operator="equal", formula=['"AŞIM"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"L{ILK}:L{SON}",
        CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        f"H{ILK}:H{SON}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=sari))
    ws.conditional_formatting.add(
        f"K{ILK}:K{SON}",
        CellIsRule(operator="equal", formula=['"YETİM"'], fill=kirmizi, font=kf))

    ws = wb["DEKONT"]
    ws.conditional_formatting.add(
        "B14", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B14", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B13", CellIsRule(operator="greaterThan", formula=["0"], fill=sari))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "zincirTolerans", "kad_esikGecikmeGun", "kad_esikGeciken",
        "kad_gecikmeOranGun", "kad_senaryoTemkinli", "kad_senaryoBaz", "kad_senaryoIyimser",
        "kad_tornadoOran", "kad_yuzdelikOran", "kad_olcekHedef", "kad_azamiTutar",
        "kad_mahsupTolerans", "kad_durumHaritasi", "kad_kuyrukOzeti_metin",
        "kad_kapaliDusum_metin", "kad_kanitRaporu_metin", "kad_dosyaSurumu", "kad_malikAdi",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "kad_acikAdet": 40, "kad_acikTutar": 41, "kad_gecikenAdet": 42,
        "kad_gecikmeMaliyeti": 43, "kad_yetimKayit": 44, "kad_gecersizGecis": 45,
        "kad_zincirKirik": 46, "kad_yasKova": 47, "kad_tahminAralik": 48,
        "kad_senaryoKarsilastirma": 49, "kad_tornadoZirve": 50, "kad_odemeOncelik": 51,
        "kad_paretoGecikme": 52, "kad_kapaliDusum": 53, "kad_kuyrukOzeti": 54,
        "kad_kanitRaporu": 55, "kad_sonrakiKimlik": 56, "kad_avansMahsup": 57,
        "kad_modulO8": 58, "kad_modulI2": 59,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "kad_yorumYas": 63, "kad_yorumGecikme": 64, "kad_yorumTahmin": 65,
        "kad_yorumSenaryo": 66, "kad_yorumTornado": 67, "kad_yorumOncelik": 68,
        "kad_yorumPareto": 69, "kad_modulO8Yorum": 70, "kad_modulI2Yorum": 71,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "kad_kararKapi", "PANO!$B$4")
    ad_ekle(wb, "kad_dekont", "DEKONT!$A$20")
    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$C$7:$C$11")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")
    ad_ekle(wb, "ListeKartlar", "LISTELER!$I$7:$I$14")


def _csv_yaz(urun_dir):
    yol = os.path.join(urun_dir, "ornek_veri.csv")
    with open(yol, "w", encoding="utf-8") as f:
        f.write("IslemId,KaynakKartId,Tarih,OncekiDurum,Durum,"
                "SozlesmeTutar,AvansTutar,KiraTutar,MahsupTutar,DekontTutar,"
                "Tutar,Donem,Aciklama\n")
        for s in ORNEK:
            f.write(
                f"{s['id']},{s['kart']},{s['tarih']:%d.%m.%Y},"
                f"{s['onceki']},{s['durum']},{s['sozlesme']},{s['avans']},"
                f"{s['kira']},{s['mahsup']},{s['dekont']},{s['tutar']},"
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
        (avans, "AVANS"),
        (dekont, "DEKONT"),
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
    dosya_ad = "KiraAvansTakipDekont.xlsx"
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
