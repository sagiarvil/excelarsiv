#!/usr/bin/env python3
"""İnşaat Hakediş ve Şantiye Maliyet Takip — A3 üretim betiği (manda v6, Hat C C-05)."""

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

URUN_AD = "İnşaat Hakediş ve Şantiye Maliyet Takip"
SURUM = "1.0.0"
RENK = "1F4E79"
KAPASITE = 250
DEMO_KART = 8
DEMO_AKIS = 28
DEMO_SANTIYE = 4
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 14)

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "DRAFT", "durum_adi": "Taslak", "sira": 1,
     "izinli_sonraki_durumlar": "BEKLIYOR", "kategori": "acik"},
    {"durum_id": "BEKLIYOR", "durum_adi": "Bekliyor", "sira": 2,
     "izinli_sonraki_durumlar": "ONAY|IPTAL", "kategori": "acik"},
    {"durum_id": "ONAY", "durum_adi": "Onaylandı", "sira": 3,
     "izinli_sonraki_durumlar": "ODEME|IPTAL", "kategori": "acik"},
    {"durum_id": "ODEME", "durum_adi": "Ödendi", "sira": 4,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "IPTAL", "durum_adi": "İptal", "sira": 5,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "DRAFT|BEKLIYOR",
    "BEKLIYOR|ONAY",
    "BEKLIYOR|IPTAL",
    "ONAY|ODEME",
    "ONAY|IPTAL",
]

SANTIYELER = [
    ("SAN-00001", "Kadıköy Konut", "İstanbul", 2_000_000, 450_000, 520_000, 180_000, 90_000),
    ("SAN-00002", "Çankaya İş Merkezi", "Ankara", 12_000_000, 800_000, 900_000, 250_000, 120_000),
    ("SAN-00003", "Bornova AVM", "İzmir", 9_000_000, 600_000, 700_000, 200_000, 100_000),
    ("SAN-00004", "Nilüfer Villa", "Bursa", 5_000_000, 300_000, 400_000, 150_000, 80_000),
]

TASERONLAR = [
    ("KAR-00001", "Atlas Kalıp Ltd.", "Kalıp", "İstanbul", "SAN-00001"),
    ("KAR-00002", "Demir Yapı A.Ş.", "Demir", "Ankara", "SAN-00001"),
    ("KAR-00003", "Nova Beton", "Beton", "İzmir", "SAN-00002"),
    ("KAR-00004", "Proje Elektrik", "Elektrik", "Bursa", "SAN-00002"),
    ("KAR-00005", "Su Tesisat Pro", "Tesisat", "Antalya", "SAN-00003"),
    ("KAR-00006", "Cephe Sistemleri", "Cephe", "Kocaeli", "SAN-00003"),
    ("KAR-00007", "Zemin Stabil", "Zemin", "Adana", "SAN-00004"),
    ("KAR-00008", "İskele Güven", "İskele", "Gaziantep", "SAN-00004"),
]

KART_SANTIYE = {k[0]: k[4] for k in TASERONLAR}


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Hakediş kuyruğu ve şantiye maliyet kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_akis():
    """≥25 satır; yetim, geçersiz geçiş ve zincir kırığı kasıtlı."""
    satirlar = []
    durumlar = ["DRAFT", "BEKLIYOR", "ONAY", "ODEME", "IPTAL", "BEKLIYOR", "ONAY", "BEKLIYOR"]
    oncekiler = ["", "DRAFT", "BEKLIYOR", "ONAY", "BEKLIYOR", "DRAFT", "BEKLIYOR", "ONAY"]
    for i in range(1, DEMO_AKIS + 1):
        kart = TASERONLAR[(i - 1) % DEMO_KART][0]
        if i in (22, 25):
            kart = "" if i == 22 else "KAR-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:
            o, d = "DRAFT", "ONAY"
        if i == 20:
            o, d = "ONAY", "BEKLIYOR"
        teklif = 100000 + i * 12500
        siparis = teklif
        hakedis = teklif if i not in (12, 19) else teklif + 45000
        fatura = hakedis if i != 15 else hakedis - 8000
        odeme = fatura if d == "ODEME" else 0
        tar = RAPOR_TARIHI - timedelta(days=(i * 3) % 75)
        sid = KART_SANTIYE.get(kart, SANTIYELER[(i - 1) % DEMO_SANTIYE][0])
        if not kart:
            sid = SANTIYELER[0][0]
        satirlar.append({
            "id": f"HAK-{i:05d}",
            "kart": kart,
            "santiye": sid,
            "tarih": tar,
            "onceki": o,
            "durum": d,
            "teklif": teklif,
            "siparis": siparis,
            "hakedis": hakedis,
            "fatura": fatura,
            "odeme": odeme,
            "tutar": hakedis,
            "donem": "2026-08",
            "aciklama": f"Hakediş kalemi {i}",
        })
    return satirlar


ORNEK = _demo_akis()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Hakediş kuyruğu ile şantiye maliyet kartını tek dosyada birleştirir; "
      "durum makinesi, yaşlandırma ve bütçe sapması ile karar üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Makrosuz durum makinesi: Taslak → Bekliyor → Onay → Ödeme / İptal",
        "Hakediş kuyruk + yaşlandırma kovaları (0-7 / 8-30 / 31-60 / 60+)",
        "Şantiye maliyet kartı: bütçe, işçilik, malzeme, makine, genel gider, hakediş",
        "Geçersiz geçiş, yetim kayıt ve zincir kırık uyarıları",
        "Karar kapısı: VERİ YOK / UYGUN / DİKKAT / KRİTİK",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Şantiye muhasebesi — kuyruk ve maliyet kartını günceller",
        "Proje müdürü — sapma ve gecikme kararını görür",
        "SMMM — dönemsel kanıt raporunu dosyalar",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) KARTLAR'a taşeron kartı girin. 2) AKIS'e hakediş yazın. "
      "3) SANTIYE_MALIYET kartını doldurun. 4) KUYRUKLAR ve PANO'dan kararı görün.",
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
      'IF(OR(ihs_maliyetAsim>0,COUNTIF(tblAkis[Tutar],"<0")>0,'
      'IFERROR(MAX(tblAkis[Tutar]),0)>ihs_azamiTutar),"KRİTİK",'
      'IF(OR(ihs_gecikenAdet>ihs_esikGeciken,ihs_yetimKayit>0,'
      'ihs_gecersizGecis>0,ihs_zincirKirik>0),"DİKKAT","UYGUN")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Açık Kuyruk Adet", "=ihs_acikAdet", CATI),
        (7, "Bekliyor Adet", '=COUNTIFS(tblAkis[Durum],"BEKLIYOR",tblAkis[Tutar],">0")', CATI),
        (8, "Onaylandı Adet", '=COUNTIFS(tblAkis[Durum],"ONAY",tblAkis[Tutar],">0")', CATI),
        (9, "Geciken Adet", "=ihs_gecikenAdet", CATI),
        (10, "Açık Tutar", "=ihs_acikTutar", TL),
        (11, "Gecikme Maliyeti", "=ihs_gecikmeMaliyeti", TL),
        (12, "Yetim Kayıt", "=ihs_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=ihs_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=ihs_zincirKirik", CATI),
        (15, "Yaş Kova Özeti", "=ihs_yasKova", None),
        (16, "Tahmin Aralık", "=ihs_tahminAralik", TL),
        (17, "Senaryo Farkı", "=ihs_senaryoKarsilastirma", TL),
        (18, "Tornado Zirve", "=ihs_tornadoZirve", TL),
        (19, "Ödeme Öncelik", "=ihs_odemeOncelik", TL),
        (20, "Pareto Gecikme", "=ihs_paretoGecikme", YÜZDE),
        (21, "Kapalı Düşüm", "=ihs_kapaliDusum", CATI),
        (22, "Şantiye Aşım", "=ihs_maliyetAsim", CATI),
        (23, "Şantiye Maliyet", "=ihs_santiyeMaliyet", None),
        (24, "Kalite Modülü", "=ihs_modulO8", CATI),
        (25, "Yüzdelik 90", "=ihs_modulI2", TL),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=ihs_yorumYas"),
        (8, "=ihs_yorumGecikme"),
        (9, "=ihs_yorumTahmin"),
        (10, "=ihs_yorumSenaryo"),
        (11, "=ihs_yorumTornado"),
        (12, "=ihs_yorumOncelik"),
        (13, "=ihs_yorumPareto"),
        (14, "=ihs_modulO8Yorum"),
        (15, "=ihs_modulI2Yorum"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 27, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 28, 1, "Durum")
    h(ws, 28, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Taslak", 4), ("Bekliyor", 8), ("Onay", 7), ("Ödeme", 5), ("İptal", 4),
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
    h(ws, 28, 8, "Hakediş")
    for i in range(8):
        h(ws, 29 + i, 7, i + 1, sayi=CATI)
        h(ws, 29 + i, 8, 180000 + i * 22000, sayi=TL)

    h(ws, 28, 10, "Senaryo")
    h(ws, 28, 11, "Tutar")
    for i, (ad, t) in enumerate([
        ("Temkinli", 4200000), ("Baz", 5100000), ("İyimser", 5900000),
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
    c3.title = "Haftalık Hakediş"
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
    h(ws, 28, 17, "Ödeme")
    for i in range(6):
        h(ws, 29 + i, 16, i + 1, sayi=CATI)
        h(ws, 29 + i, 17, 900000 + i * 110000, sayi=TL)

    c5 = BarChart()
    c5.title = "Uyarı Türleri"
    c5.add_data(Reference(ws, min_col=14, min_row=28, max_row=32), titles_from_data=True)
    c5.set_categories(Reference(ws, min_col=13, min_row=29, max_row=32))
    c5.width, c5.height = 10, 7
    ws.add_chart(c5, "A69")

    c6 = LineChart()
    c6.title = "Aylık Ödeme"
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
    alt_bant(ws, 3, "Taşeron / proje kartları — her kaydın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=ihs_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: KAR-00009 | Risk: Elle çakışma")

    sutunlar = ["KartId", "Unvan", "Brans", "Sehir", "SozlesmeNo", "LimitTl",
                "RiskNotu", "Aktif", "SantiyeId", "ToplamHacim", "OrtGecikme"]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "ToplamHacim": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'SUMIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],tblAkis[Tutar]))'
        ),
        "OrtGecikme": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'IFERROR(AVERAGEIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],'
            'tblAkis[YasGun]),0))'
        ),
    }
    for i, (kid, unvan, brans, sehir, sid) in enumerate(TASERONLAR):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Kart Kimliği", mesaj="KAR-##### formatında")
        _sari(ws, r, 2, unvan, baslik="Unvan", mesaj="Taşeron/firma unvanı")
        _sari(ws, r, 3, brans, baslik="Branş", mesaj="İş kalemi branşı")
        _sari(ws, r, 4, sehir, baslik="Şehir", mesaj="Şehir")
        _sari(ws, r, 5, f"SZL-2026-{i + 1:03d}", baslik="Sözleşme No", mesaj="Sözleşme numarası")
        _sari(ws, r, 6, 2500000 + i * 150000, sayi=TL, baslik="Limit", mesaj="Sözleşme limiti TL")
        _sari(ws, r, 7, "Normal", baslik="Risk Notu", mesaj="Kısa risk notu")
        _sari(ws, r, 8, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
        _sari(ws, r, 9, sid, baslik="Şantiye Kimliği", mesaj="SANTIYE_MALIYET kartındaki kimlik")
    for r in range(ILK + DEMO_KART, SON + 1):
        for c in range(1, 10):
            _sari(ws, r, c, None, sayi=TL if c == 6 else None,
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblKartlar", f"A{HDR}:K{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Kart Kimliği", mesaj="KAR-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Unvan"), (3, "Branş"), (4, "Şehir"), (5, "Sözleşme"), (7, "Risk")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "decimal", "0", f"F{ILK}:F{SON}",
              baslik="Limit", mesaj="Limit ≥ 0",
              hata_baslik="Limit", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeEvetHayir", f"H{ILK}:H{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "textLength", "0", f"I{ILK}:I{SON}",
              baslik="Şantiye", mesaj="SAN-##### veya boş",
              hata_baslik="Şantiye", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 12)})
    ws.column_dimensions["B"].width = 22


def akis(ws):
    sayfa_hazirla(ws, "AKIS", RENK, URUN_AD, son_kolon=22)
    alt_bant(ws, 3,
             "İşlem satırları — durum makinesi + yetim + geçersiz geçiş + zincir bayrakları",
             son_kolon=18)
    sutunlar = [
        "IslemId", "KaynakKartId", "SantiyeId", "Tarih", "OncekiDurum", "Durum",
        "Teklif", "Siparis", "Hakedis", "Fatura", "Odeme", "Tutar",
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
        "tblAkis[[#This Row],[Teklif]]",
        "tblAkis[[#This Row],[Hakedis]]",
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
        _sari(ws, r, 1, s["id"], baslik="İşlem Kimliği", mesaj="HAK-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Kart", mesaj="KARTLAR'daki KartId")
        _sari(ws, r, 3, s["santiye"], baslik="Şantiye", mesaj="SANTIYE_MALIYET kimliği")
        _sari(ws, r, 4, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Olay tarihi")
        _sari(ws, r, 5, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 6, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 7, s["teklif"], sayi=TL, baslik="Teklif", mesaj="Teklif tutarı TL")
        _sari(ws, r, 8, s["siparis"], sayi=TL, baslik="Sipariş", mesaj="Sipariş tutarı TL")
        _sari(ws, r, 9, s["hakedis"], sayi=TL, baslik="Hakediş", mesaj="Hakediş tutarı TL")
        _sari(ws, r, 10, s["fatura"], sayi=TL, baslik="Fatura", mesaj="Fatura tutarı TL")
        _sari(ws, r, 11, s["odeme"], sayi=TL, baslik="Ödeme", mesaj="Ödeme tutarı TL")
        _sari(ws, r, 12, s["tutar"], sayi=TL, baslik="Tutar", mesaj="Kuyruk tutarı TL")
        _sari(ws, r, 13, s["donem"], baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 14, s["aciklama"], baslik="Açıklama", mesaj="Kısa açıklama")

    for r in range(ILK + DEMO_AKIS, SON + 1):
        for c in range(1, 15):
            fmt = TARİH if c == 4 else (TL if c in (7, 8, 9, 10, 11, 12) else None)
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblAkis", f"A{HDR}:T{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="İşlem Kimliği", mesaj="HAK-##### girin",
              hata_baslik="Kimlik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Kaynak Kart", mesaj="Kart kimliği (boş = yetim riski)",
              hata_baslik="Kart", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "textLength", "0", f"C{ILK}:C{SON}",
              baslik="Şantiye", mesaj="SAN-#####",
              hata_baslik="Şantiye", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"D{ILK}:D{SON}",
              baslik="Tarih", mesaj="Olay tarihi",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "list", "ListeDurumlar", f"E{ILK}:E{SON}",
              baslik="Önceki Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"F{ILK}:F{SON}",
              baslik="Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    for c, ad in [(7, "Teklif"), (8, "Sipariş"), (9, "Hakediş"),
                  (10, "Fatura"), (11, "Ödeme"), (12, "Tutar")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "textLength", "0", f"M{ILK}:M{SON}",
              baslik="Dönem", mesaj="YYYY-AA",
              hata_baslik="Dönem", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="20")
    dogrulama(ws, "textLength", "0", f"N{ILK}:N{SON}",
              baslik="Açıklama", mesaj="Kısa açıklama",
              hata_baslik="Açıklama", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="120")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 21)})
    ws.column_dimensions["N"].width = 18
    ws.column_dimensions["R"].width = 16


def kuyruklar(ws):
    sayfa_hazirla(ws, "KUYRUKLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Durum bazlı canlı kuyruk — kapalı kayıtlar düşer", son_kolon=10)
    h(ws, 5, 1, "Kuyruk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Tutar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    h(ws, 6, 1, "Bekliyor")
    h(ws, 6, 2, kuyruk_sayim("tblAkis", "BEKLIYOR"), sayi=CATI)
    h(ws, 6, 3, kuyruk_tutar("tblAkis", "BEKLIYOR"), sayi=TL)

    h(ws, 7, 1, "Onaylandı")
    h(ws, 7, 2, kuyruk_sayim("tblAkis", "ONAY"), sayi=CATI)
    h(ws, 7, 3, kuyruk_tutar("tblAkis", "ONAY"), sayi=TL)

    h(ws, 8, 1, "Taslak")
    h(ws, 8, 2, kuyruk_sayim("tblAkis", "DRAFT"), sayi=CATI)
    h(ws, 8, 3, kuyruk_tutar("tblAkis", "DRAFT"), sayi=TL)

    h(ws, 9, 1, "Geciken (açık, yaş>eşik)")
    h(ws, 9, 2,
      '=COUNTIFS(tblAkis[YasGun],">"&ihs_esikGecikmeGun,tblAkis[Durum],"<>ODEME",'
      'tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")',
      sayi=CATI)
    h(ws, 9, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[YasGun],">"&ihs_esikGecikmeGun,'
      'tblAkis[Durum],"<>ODEME",tblAkis[Durum],"<>IPTAL")',
      sayi=TL)

    h(ws, 11, 1, "Kapalı (Ödeme+İptal) — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 1, "Ödendi")
    h(ws, 12, 2, kuyruk_sayim("tblAkis", "ODEME"), sayi=CATI)
    h(ws, 12, 3, kuyruk_tutar("tblAkis", "ODEME"), sayi=TL)
    h(ws, 13, 1, "İptal")
    h(ws, 13, 2, kuyruk_sayim("tblAkis", "IPTAL"), sayi=CATI)
    h(ws, 13, 3, kuyruk_tutar("tblAkis", "IPTAL"), sayi=TL)

    h(ws, 15, 1, "Kuyruk Özeti", kalin=True)
    h(ws, 15, 2, "=ihs_kuyrukOzeti", kaydir=True)
    ws.merge_cells("B15:F15")

    h(ws, 17, 1, "Yaş Kova Dağılımı", kalin=True, yazi=KOYU_LACIVERT)
    for i, kova in enumerate(["0-7", "8-30", "31-60", "60+"], 18):
        h(ws, i, 1, kova)
        h(ws, i, 2, f'=COUNTIF(tblAkis[YasKova],"{kova}")', sayi=CATI)
        h(ws, i, 3, f'=SUMIF(tblAkis[YasKova],"{kova}",tblAkis[Tutar])', sayi=TL)

    genislik(ws, {"A": 36, "B": 14, "C": 16})


def santiye_maliyet(ws):
    sayfa_hazirla(ws, "SANTIYE_MALIYET", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3,
             "Şantiye maliyet kartı — bütçe, işçilik, malzeme, makine, genel gider + bağlı hakediş",
             son_kolon=14)
    h(ws, 4, 1, "Toplam Bütçe", yazi=GRİ)
    h(ws, 4, 2, "=SUM(tblSantiye[ButceTl])", sayi=TL, kalin=True)
    h(ws, 4, 3, "Toplam Maliyet", yazi=GRİ)
    h(ws, 4, 4, f"=SUM($J${ILK}:$J${SON})", sayi=TL, kalin=True)
    h(ws, 4, 5, "Sapma", yazi=GRİ)
    h(ws, 4, 6, f"=SUM($K${ILK}:$K${SON})", sayi=TL, kalin=True)
    h(ws, 4, 7, "Aşım Adet", yazi=GRİ)
    h(ws, 4, 8, f'=COUNTIF($M${ILK}:$M${SON},"AŞIM")', sayi=CATI, kalin=True)

    tablo_sut = ["SantiyeId", "SantiyeAdi", "Sehir", "ButceTl", "IscilikTl"]
    baslik_satiri(ws, HDR, tablo_sut + [
        "MalzemeTl", "MakineTl", "GenelGiderTl",
        "HakedisToplam", "ToplamMaliyet", "SapmaTl", "SapmaOran", "AsimBayrak",
    ])
    for i, (sid, ad, sehir, butce, isc, mal, mak, gen) in enumerate(SANTIYELER):
        r = ILK + i
        _sari(ws, r, 1, sid, baslik="Şantiye Kimliği", mesaj="SAN-#####")
        _sari(ws, r, 2, ad, baslik="Şantiye Adı", mesaj="Proje / şantiye adı")
        _sari(ws, r, 3, sehir, baslik="Şehir", mesaj="Şehir")
        _sari(ws, r, 4, butce, sayi=TL, baslik="Bütçe", mesaj="Şantiye bütçesi TL")
        _sari(ws, r, 5, isc, sayi=TL, baslik="İşçilik", mesaj="İşçilik maliyeti TL")
        _sari(ws, r, 6, mal, sayi=TL, baslik="Malzeme", mesaj="Malzeme maliyeti TL")
        _sari(ws, r, 7, mak, sayi=TL, baslik="Makine", mesaj="Makine maliyeti TL")
        _sari(ws, r, 8, gen, sayi=TL, baslik="Genel Gider", mesaj="Genel gider TL")
    for r in range(ILK + DEMO_SANTIYE, SON + 1):
        for c, ad in enumerate(
                ["SantiyeId", "SantiyeAdi", "Sehir", "ButceTl", "IscilikTl",
                 "MalzemeTl", "MakineTl", "GenelGiderTl"], 1):
            _sari(ws, r, c, None, sayi=TL if c >= 4 else None,
                  baslik=ad, mesaj="Manuel giriş")
    for r in range(ILK, SON + 1):
        h(ws, r, 9,
          f'=IF(A{r}="","",SUMIF(tblAkis[SantiyeId],A{r},tblAkis[Tutar]))',
          sayi=TL)
        h(ws, r, 10, f'=IF(A{r}="","",E{r}+F{r}+G{r}+H{r}+I{r})', sayi=TL)
        h(ws, r, 11, f'=IF(A{r}="","",J{r}-D{r})', sayi=TL)
        h(ws, r, 12, f'=IF(A{r}="","",IFERROR(K{r}/MAX(D{r},1),0))', sayi=YÜZDE)
        h(ws, r, 13,
          f'=IF(A{r}="","",IF(J{r}>D{r}*ihs_asimCarpan,"AŞIM","UYGUN"))')
    tablo_ekle(ws, "tblSantiye", f"A{HDR}:E{SON}", tablo_sut)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Şantiye Kimliği", mesaj="SAN-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Şantiye Adı", mesaj="Proje adı",
              hata_baslik="Ad", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "textLength", "0", f"C{ILK}:C{SON}",
              baslik="Şehir", mesaj="Şehir",
              hata_baslik="Şehir", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    for c, ad in [(4, "Bütçe"), (5, "İşçilik"), (6, "Malzeme"),
                  (7, "Makine"), (8, "Genel Gider")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")

    h(ws, 4, 15, "Şantiye", kalin=True)
    h(ws, 4, 16, "Bütçe", kalin=True)
    h(ws, 4, 17, "Maliyet", kalin=True)
    for i, (sid, ad, _sehir, butce, isc, mal, mak, gen) in enumerate(SANTIYELER):
        h(ws, 5 + i, 15, ad)
        h(ws, 5 + i, 16, butce, sayi=TL)
        h(ws, 5 + i, 17, isc + mal + mak + gen + 1_800_000, sayi=TL)

    c = BarChart()
    c.title = "Bütçe ve Maliyet"
    c.add_data(Reference(ws, min_col=16, min_row=4, max_row=8), titles_from_data=True)
    c.add_data(Reference(ws, min_col=17, min_row=4, max_row=8), titles_from_data=True)
    c.set_categories(Reference(ws, min_col=15, min_row=5, max_row=8))
    c.width, c.height = 12, 7
    ws.add_chart(c, "A20")

    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 18)})
    ws.column_dimensions["B"].width = 22


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
        ("H04", "Teklif toplam", "=SUM(tblAkis[Teklif])"),
        ("H05", "Sipariş toplam", "=SUM(tblAkis[Siparis])"),
        ("H06", "Hakediş toplam", "=SUM(tblAkis[Hakedis])"),
        ("H07", "Fatura toplam", "=SUM(tblAkis[Fatura])"),
        ("H08", "Ödeme toplam", "=SUM(tblAkis[Odeme])"),
        ("H09", "Taslak adet", '=COUNTIF(tblAkis[Durum],"DRAFT")'),
        ("H10", "Bekliyor adet", '=COUNTIF(tblAkis[Durum],"BEKLIYOR")'),
        ("H11", "Onay adet", '=COUNTIF(tblAkis[Durum],"ONAY")'),
        ("H12", "Ödeme adet", '=COUNTIF(tblAkis[Durum],"ODEME")'),
        ("H13", "İptal adet", '=COUNTIF(tblAkis[Durum],"IPTAL")'),
        ("H14", "Açık adet", "=C14+C15+C16"),
        ("H15", "Kapalı adet", "=C17+C18"),
        ("H16", "Açık tutar",
         '=SUMIF(tblAkis[Durum],"DRAFT",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"BEKLIYOR",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"ONAY",tblAkis[Tutar])'),
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
         '=COUNTIFS(tblAkis[YasGun],">"&ihs_esikGecikmeGun,tblAkis[Durum],"<>ODEME",'
         'tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")'),
        ("H27", "Geciken tutar",
         '=SUMIFS(tblAkis[Tutar],tblAkis[YasGun],">"&ihs_esikGecikmeGun,'
         'tblAkis[Durum],"<>ODEME",tblAkis[Durum],"<>IPTAL")'),
        ("H28", "Gecikme maliyeti", "=C32*ihs_gecikmeOranGun"),
        ("H29", "Temkinli senaryo", "=C21*ihs_senaryoTemkinli"),
        ("H30", "Baz senaryo", "=C21*ihs_senaryoBaz"),
        ("H31", "İyimser senaryo", "=C21*ihs_senaryoIyimser"),
        ("H32", "Senaryo farkı", "=C36-C34"),
        ("H33", "Tornado (oran×tutar)", "=C21*ihs_tornadoOran"),
        ("H34", "Ödeme öncelik havuzu",
         '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"ONAY",tblAkis[YasGun],">"&ihs_esikGecikmeGun)'),
        ("H35", "Pareto gecikme payı", "=IFERROR(C32/MAX(C21,1),0)"),
        ("H36", "Tahmin yüzdelik",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[Tutar],ihs_yuzdelikOran),0)"),
        ("H37", "Tahmin eğilim",
         "=IFERROR(FORECAST(COUNTA(tblAkis[IslemId])+1,tblAkis[Tutar],"
         "tblAkis[KayitDolu]),C41)"),
        ("H38", "Tahmin aralık", "=IFERROR((C41+C42)/2,0)"),
        ("H39", "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        ("H40", "Limit toplam", "=SUM(tblKartlar[LimitTl])"),
        ("H41", "Limit kullanım oranı", "=IFERROR(C8/MAX(C45,1),0)"),
        ("H42", "Kuyruk özeti metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Bekliyor",C15,"Onay",C16,"Geciken",C31)'),
        ("H43", "Yaş kova metin",
         '=_xlfn.TEXTJOIN("/",TRUE,C26,C27,C28,C29)'),
        ("H44", "Karar skoru ham",
         "=IF(C6=0,0,MAX(0,100-C22*10-C23*10-C24*10-C31*2))"),
        ("H45", "Veri hazırlık", "=IF(C6=0,0,C7/MAX(C6,1))"),
        ("H46", "Şantiye adet", '=COUNTA(tblSantiye[SantiyeId])'),
        ("H47", "Şantiye bütçe", "=SUM(tblSantiye[ButceTl])"),
        ("H48", "Şantiye hakediş", "=SUM(SANTIYE_MALIYET!$I$6:$I$255)"),
        ("H49", "Şantiye toplam maliyet", "=SUM(SANTIYE_MALIYET!$J$6:$J$255)"),
        ("H50", "Şantiye sapma", "=SUM(SANTIYE_MALIYET!$K$6:$K$255)"),
        ("H51", "Şantiye aşım adet", '=COUNTIF(SANTIYE_MALIYET!$M$6:$M$255,"AŞIM")'),
        ("H52", "Kalite modülü", "=C49"),
        ("H53", "Yüzdelik 90",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[YasGun],ihs_yuzdelikOran),0)"),
        ("H54", "Şantiye özet",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Sapma",TEXT(C55,"0"),"Aşım",C56)'),
        ("H55", "Kanıt özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Açık",C19,"Yetim",C22,"Zincir",C24,'
         '"Geciken",C31,"Aşım",C56)'),
    ]
    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        fmt = TL if any(x in acik.lower() for x in (
            "tutar", "maliyet", "senaryo", "tornado", "öncelik", "tahmin", "limit", "fark",
            "toplam", "havuz", "bütçe", "hakediş", "sapma", "yüzdelik",
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
    h(ws, 65, 2, "=ihs_senaryoTemkinli*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 65, 3, "=C34", sayi=TL)
    h(ws, 66, 1, "Baz")
    h(ws, 66, 2, "=ihs_senaryoBaz*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 66, 3, "=C35", sayi=TL)
    h(ws, 67, 1, "İyimser")
    h(ws, 67, 2, "=ihs_senaryoIyimser*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 67, 3, "=C36", sayi=TL)

    genislik(ws, {"A": 10, "B": 28, "C": 50})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Dönemsel Hakediş ve Şantiye Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Açık Kuyruk", "=ihs_acikAdet", CATI),
        (10, "Açık Tutar", "=ihs_acikTutar", TL),
        (11, "Gecikme Maliyeti", "=ihs_gecikmeMaliyeti", TL),
        (12, "Yetim", "=ihs_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=ihs_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=ihs_zincirKirik", CATI),
        (15, "Tahmin Aralık", "=ihs_tahminAralik", TL),
        (16, "Şantiye Aşım", "=ihs_maliyetAsim", CATI),
        (17, "Şantiye Maliyet", "=ihs_santiyeMaliyet", None),
        (18, "Kanıt Özeti", "=ihs_kanitRaporu", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 20, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 1,
      "Durum haritası AYARLAR'dadır. Geçersiz geçiş engellenmez, görünür kılınır. "
      "Şantiye aşımı bütçe × ihs_asimCarpan eşiğine bağlıdır. "
      "Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A21:H21")
    h(ws, 23, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 1, "1. Yetim kayıtları KARTLAR'a bağlayın veya satırı düzeltin.")
    h(ws, 25, 1, "2. Geçersiz geçişleri durum haritasına uygun ilerletin.")
    h(ws, 26, 1, "3. Aşım bayraklı şantiyelerde bütçe veya maliyet kalemini revize edin.")
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
        (8, "Onay", '=COUNTIF(tblAkis[Durum],"ONAY")'),
        (9, "Ödeme", '=COUNTIF(tblAkis[Durum],"ODEME")'),
        (10, "İptal", '=COUNTIF(tblAkis[Durum],"IPTAL")'),
        (11, "Açık toplam", '=B7+B8+COUNTIF(tblAkis[Durum],"DRAFT")'),
        (12, "Yetim", "=ihs_yetimKayit"),
        (13, "Geçersiz", "=ihs_gecersizGecis"),
        (14, "Zincir kırık", "=ihs_zincirKirik"),
        (15, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (16, "Kalite skoru", "=HESAP!C49"),
        (17, "Açık tutar", "=ihs_acikTutar"),
        (18, "Gecikme maliyeti", "=ihs_gecikmeMaliyeti"),
        (19, "Kapalı düşüm", "=ihs_kapaliDusum"),
        (20, "Şantiye aşım", "=ihs_maliyetAsim"),
        (21, "Şantiye adet", '=COUNTA(tblSantiye[SantiyeId])'),
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
      f"Dosyada {DEMO_KART} kart, {DEMO_SANTIYE} şantiye ve {DEMO_AKIS} akış satırı vardır. "
      "Yetim (22, 25), geçersiz geçiş (18, 20) ve zincir kırık (12, 19) kasıtlıdır. "
      "Kadıköy Konut bütçesi aşım üretir; karar KRİTİK olabilir. "
      "KARTLAR, AKIS ve SANTIYE_MALIYET'i temizleyip kendi verinizi girebilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Durumlar: Taslak/Bekliyor/Onay/Ödeme/İptal — kapalılar kuyruktan düşer")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 10000, sayi=CATI)
    genislik(ws, {"A": 70, "B": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=10)
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

    h(ws, 3, 7, "Branş", kalin=True)
    h(ws, 6, 7, "Brans", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(["Kalıp", "Demir", "Beton", "Elektrik", "Tesisat", "Cephe"], 7):
        _sari(ws, i, 7, b, baslik="Branş", mesaj="Branş listesi")

    h(ws, 3, 9, "Şantiye Kimlik", kalin=True)
    h(ws, 6, 9, "SantiyeId", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, s in enumerate(SANTIYELER, 7):
        _sari(ws, i, 9, s[0], baslik="Şantiye", mesaj="SAN-#####")
    genislik(ws, {"A": 14, "C": 20, "E": 12, "G": 14, "I": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + durum haritası (A01/A06)", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("zincirTolerans", 10000, "TL", "İş kuralı — zincir sapma", "01.01.2026",
         "ZİNCİR KIRIK eşiği"),
        ("ihs_esikGecikmeGun", 30, "gün", "İç politika", "01.01.2026", "Gecikme eşiği"),
        ("ihs_esikGeciken", 5, "adet", "İç politika", "01.01.2026", "Karar eşiği geciken"),
        ("ihs_gecikmeOranGun", 0.0005, "oran", "Finans varsayımı", "01.01.2026",
         "Günlük gecikme maliyeti oranı"),
        ("ihs_senaryoTemkinli", 0.85, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("ihs_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("ihs_senaryoIyimser", 1.15, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("ihs_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("ihs_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("ihs_olcekHedef", 10000, "satir", "Manda A3 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("ihs_azamiTutar", 50000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul"),
        ("ihs_asimCarpan", 1.0, "oran", "Şantiye bütçe eşiği", "01.01.2026",
         "Aşım = maliyet > bütçe × çarpan"),
        ("ihs_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi",
         "01.01.2026", "Durum makinesi kaynağı"),
        ("ihs_kuyrukOzeti_metin", "KUYRUKLAR canlı özet", "metin", "A05 kuyruk", "01.01.2026",
         "Kuyruk görünümü"),
        ("ihs_kapaliDusum_metin", "kapali kategori kuyruktan düşer", "metin", "A06", "01.01.2026",
         "Kapalı düşüm kuralı"),
        ("ihs_kanitRaporu_metin", "RAPOR kanıt çıktısı", "metin", "Dönemsel rapor", "01.01.2026",
         "Kanıt raporu"),
        ("ihs_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("ihs_firmaUnvan", "Örnek İnşaat A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma"),
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
        (40, "ihs_acikAdet", "=HESAP!C19"),
        (41, "ihs_acikTutar", "=HESAP!C21"),
        (42, "ihs_gecikenAdet", "=HESAP!C31"),
        (43, "ihs_gecikmeMaliyeti", "=HESAP!C33"),
        (44, "ihs_yetimKayit", "=HESAP!C22"),
        (45, "ihs_gecersizGecis", "=HESAP!C23"),
        (46, "ihs_zincirKirik", "=HESAP!C24"),
        (47, "ihs_yasKova", "=HESAP!C48"),
        (48, "ihs_tahminAralik", "=HESAP!C43"),
        (49, "ihs_senaryoKarsilastirma", "=HESAP!C37"),
        (50, "ihs_tornadoZirve", "=HESAP!C38"),
        (51, "ihs_odemeOncelik", "=HESAP!C39"),
        (52, "ihs_paretoGecikme", "=HESAP!C40"),
        (53, "ihs_kapaliDusum", "=HESAP!C20"),
        (54, "ihs_kuyrukOzeti", "=HESAP!C47"),
        (55, "ihs_kanitRaporu", "=HESAP!C60"),
        (56, "ihs_sonrakiKimlik",
         sonraki_kimlik_formul("KAR", "tblKartlar[KartId]").replace(
             'MAX(tblKartlar[KartId])',
             'COUNTA(tblKartlar[KartId])')),
        (57, "ihs_santiyeMaliyet", "=HESAP!C59"),
        (58, "ihs_maliyetAsim", "=HESAP!C56"),
        (59, "ihs_modulO8", "=HESAP!C57"),
        (60, "ihs_modulI2", "=HESAP!C58"),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 62, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (63, "ihs_yorumYas",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yaş kovaları",ihs_yasKova)'),
        (64, "ihs_yorumGecikme",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Gecikme maliyeti",TEXT(ihs_gecikmeMaliyeti,"0.00"),"TL")'),
        (65, "ihs_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(ihs_tahminAralik,"0.00"),"TL")'),
        (66, "ihs_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(ihs_senaryoKarsilastirma,"0.00"),"TL")'),
        (67, "ihs_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(ihs_tornadoZirve,"0.00"),"TL")'),
        (68, "ihs_yorumOncelik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Ödeme öncelik havuzu",TEXT(ihs_odemeOncelik,"0.00"),"TL")'),
        (69, "ihs_yorumPareto",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Pareto gecikme payı",TEXT(ihs_paretoGecikme,"0.0%"))'),
        (70, "ihs_modulO8Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite",TEXT(ihs_modulO8,"0"),"yetim/geçiş/zincir düşer")'),
        (71, "ihs_modulI2Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yüzdelik 90 tutar",TEXT(ihs_modulI2,"0.00"),"TL")'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 22, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. KARTLAR sayfasına taşeron kartı ekleyin; şantiye kimliğini bağlayın.",
        "2. AKIS sayfasına her hakediş olayını yazın; durum kolonunu listeden seçin.",
        "3. SANTIYE_MALIYET kartına bütçe ve maliyet kalemlerini girin; hakediş otomatik toplanır.",
        "4. Önceki durum boş bırakılırsa geçiş uyarısı çalışmaz; bilinçli geçişte doldurun.",
        "5. Geçersiz geçiş engellenmez — kırmızı uyarı üretir (makrosuz kural).",
        "6. Kapalı durumlar (Ödendi/İptal) açık kuyruklardan düşer.",
        "7. Bütçe aşımı KRİTİK, gecikme eşiği DİKKAT, boş veri VERİ YOK üretir.",
        "8. raporTarihi AYARLAR'dadır; TODAY kullanılmaz.",
        "9. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı hücreler açıktır.",
        "10. RAPOR sayfasını PDF olarak yazdırabilirsiniz.",
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
        ("Q", '"YETİM"', kirmizi, kf),
        ("R", '"GEÇERSİZ GEÇİŞ"', kirmizi, kf),
        ("S", '"ZİNCİR KIRIK"', kirmizi, kf),
        ("S", '"OK"', yesil, yf),
    ]:
        ws.conditional_formatting.add(
            f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=[val], fill=fill, font=font))
    for durum, fill in [("BEKLIYOR", sari), ("ONAY", yesil), ("IPTAL", kirmizi)]:
        ws.conditional_formatting.add(
            f"F{ILK}:F{SON}",
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
        CellIsRule(operator="greaterThan", formula=["5000000"], fill=yesil))

    ws = wb["SANTIYE_MALIYET"]
    ws.conditional_formatting.add(
        f"M{ILK}:M{SON}",
        CellIsRule(operator="equal", formula=['"AŞIM"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"M{ILK}:M{SON}",
        CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        f"K{ILK}:K{SON}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"K{ILK}:K{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=yesil, font=yf))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "zincirTolerans", "ihs_esikGecikmeGun", "ihs_esikGeciken",
        "ihs_gecikmeOranGun", "ihs_senaryoTemkinli", "ihs_senaryoBaz", "ihs_senaryoIyimser",
        "ihs_tornadoOran", "ihs_yuzdelikOran", "ihs_olcekHedef", "ihs_azamiTutar",
        "ihs_asimCarpan", "ihs_durumHaritasi", "ihs_kuyrukOzeti_metin",
        "ihs_kapaliDusum_metin", "ihs_kanitRaporu_metin", "ihs_dosyaSurumu", "ihs_firmaUnvan",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "ihs_acikAdet": 40, "ihs_acikTutar": 41, "ihs_gecikenAdet": 42,
        "ihs_gecikmeMaliyeti": 43, "ihs_yetimKayit": 44, "ihs_gecersizGecis": 45,
        "ihs_zincirKirik": 46, "ihs_yasKova": 47, "ihs_tahminAralik": 48,
        "ihs_senaryoKarsilastirma": 49, "ihs_tornadoZirve": 50, "ihs_odemeOncelik": 51,
        "ihs_paretoGecikme": 52, "ihs_kapaliDusum": 53, "ihs_kuyrukOzeti": 54,
        "ihs_kanitRaporu": 55, "ihs_sonrakiKimlik": 56, "ihs_santiyeMaliyet": 57,
        "ihs_maliyetAsim": 58, "ihs_modulO8": 59, "ihs_modulI2": 60,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "ihs_yorumYas": 63, "ihs_yorumGecikme": 64, "ihs_yorumTahmin": 65,
        "ihs_yorumSenaryo": 66, "ihs_yorumTornado": 67, "ihs_yorumOncelik": 68,
        "ihs_yorumPareto": 69, "ihs_modulO8Yorum": 70, "ihs_modulI2Yorum": 71,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ihs_kararKapi", "PANO!$B$4")
    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$C$7:$C$11")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")
    ad_ekle(wb, "ListeSantiye", "LISTELER!$I$7:$I$10")


def _csv_yaz(urun_dir):
    yol = os.path.join(urun_dir, "ornek_veri.csv")
    with open(yol, "w", encoding="utf-8") as f:
        f.write("IslemId,KaynakKartId,SantiyeId,Tarih,OncekiDurum,Durum,"
                "Teklif,Siparis,Hakedis,Fatura,Odeme,Tutar,Donem,Aciklama\n")
        for s in ORNEK:
            f.write(
                f"{s['id']},{s['kart']},{s['santiye']},{s['tarih']:%d.%m.%Y},"
                f"{s['onceki']},{s['durum']},{s['teklif']},{s['siparis']},"
                f"{s['hakedis']},{s['fatura']},{s['odeme']},{s['tutar']},"
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
        (santiye_maliyet, "SANTIYE_MALIYET"),
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
    dosya_ad = "InsaatHakedisSantiyeMaliyet.xlsx"
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
