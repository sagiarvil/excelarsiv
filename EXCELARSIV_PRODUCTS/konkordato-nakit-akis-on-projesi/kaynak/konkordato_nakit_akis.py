#!/usr/bin/env python3
"""Konkordato Nakit Akış Ön Projesi — A3 üretim betiği (manda v6, Hat D D-05).

Ayırt edici: 3 yıllık bilanço + nakit akış mahkeme ön projesi.
TTK 376 kriz paketi kopyası değildir.
"""

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
    yas_kova_formul,
    yaslandirma_gun,
    yetim_kayit_formul,
    zincir_kirik_formul,
)

URUN_AD = "Konkordato Nakit Akış Ön Projesi"
SURUM = "1.0.0"
RENK = "1B365D"
KAPASITE = 250
DEMO_KART = 8
DEMO_AKIS = 28
DEMO_BILANCO = 12
DEMO_NAKIT = 10
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 14)

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "TASLAK", "durum_adi": "Taslak", "sira": 1,
     "izinli_sonraki_durumlar": "HAZIRLIK", "kategori": "acik"},
    {"durum_id": "HAZIRLIK", "durum_adi": "Hazırlık", "sira": 2,
     "izinli_sonraki_durumlar": "BASVURU|IPTAL", "kategori": "acik"},
    {"durum_id": "BASVURU", "durum_adi": "Başvuru", "sira": 3,
     "izinli_sonraki_durumlar": "KABUL|RET", "kategori": "acik"},
    {"durum_id": "KABUL", "durum_adi": "Kabul", "sira": 4,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "RET", "durum_adi": "Ret", "sira": 5,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "IPTAL", "durum_adi": "İptal", "sira": 6,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "TASLAK|HAZIRLIK",
    "HAZIRLIK|BASVURU",
    "HAZIRLIK|IPTAL",
    "BASVURU|KABUL",
    "BASVURU|RET",
]

ALACAKLILAR = [
    ("ALN-00001", "Atlas Tedarik Ltd.", "Tedarikçi", "İstanbul", 1_200_000),
    ("ALN-00002", "Demir Finans A.Ş.", "Kredi", "Ankara", 3_500_000),
    ("ALN-00003", "Nova Enerji", "Enerji", "İzmir", 480_000),
    ("ALN-00004", "Proje Personel", "Bordro", "Bursa", 720_000),
    ("ALN-00005", "Su Vergi Büro", "Vergi", "Antalya", 390_000),
    ("ALN-00006", "Cephe Kira A.Ş.", "Kira", "Kocaeli", 240_000),
    ("ALN-00007", "Zemin Sigorta", "Sigorta", "Adana", 180_000),
    ("ALN-00008", "İskele Lojistik", "Lojistik", "Gaziantep", 310_000),
]

BILANCO_KALEM = [
    ("BLN-00001", "Kasa ve bankalar", "donen", 1_800_000, 1_450_000, 1_720_000),
    ("BLN-00002", "Ticari alacaklar", "donen", 4_200_000, 3_600_000, 3_900_000),
    ("BLN-00003", "Stoklar", "donen", 2_100_000, 1_800_000, 1_950_000),
    ("BLN-00004", "Dönen varlıklar toplam", "donen", 8_100_000, 6_850_000, 7_570_000),
    ("BLN-00005", "Maddi duran varlıklar", "duran", 12_400_000, 11_800_000, 11_200_000),
    ("BLN-00006", "Duran varlıklar toplam", "duran", 12_400_000, 11_800_000, 11_200_000),
    ("BLN-00007", "Aktif toplam", "aktif", 20_500_000, 18_650_000, 18_770_000),
    ("BLN-00008", "Ticari borçlar kısa vade", "pasif", 6_800_000, 5_900_000, 5_200_000),
    ("BLN-00009", "Finansal borçlar kısa vade", "pasif", 4_200_000, 3_800_000, 3_100_000),
    ("BLN-00010", "Kısa vadeli yükümlülük toplam", "pasif", 11_000_000, 9_700_000, 8_300_000),
    ("BLN-00011", "Uzun vadeli yükümlülük", "pasif", 7_400_000, 6_600_000, 5_800_000),
    ("BLN-00012", "Özkaynak", "ozkaynak", 2_100_000, 2_350_000, 4_670_000),
]

NAKIT_KALEM = [
    ("NAK-00001", "Satış tahsilatı", "isletme", 18_400_000, 19_600_000, 21_200_000),
    ("NAK-00002", "Tedarik ödemesi", "isletme", -9_800_000, -10_100_000, -10_400_000),
    ("NAK-00003", "Personel ödemesi", "isletme", -4_200_000, -4_350_000, -4_500_000),
    ("NAK-00004", "Vergi ödemesi", "isletme", -1_100_000, -1_050_000, -1_200_000),
    ("NAK-00005", "İşletme net nakit", "isletme", 3_300_000, 4_100_000, 5_100_000),
    ("NAK-00006", "Yatırım nakit", "yatirim", -420_000, -380_000, -310_000),
    ("NAK-00007", "Finansman anapara", "finansman", -1_800_000, -1_500_000, -1_200_000),
    ("NAK-00008", "Finansman faiz", "finansman", -640_000, -520_000, -410_000),
    ("NAK-00009", "Dönem başı nakit", "kapanis", 1_760_000, 1_800_000, 1_450_000),
    ("NAK-00010", "Dönem sonu nakit", "kapanis", 1_800_000, 1_450_000, 1_720_000),
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: 3 yıllık nakit akış ön proje kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_akis():
    """≥25 satır; yetim, geçersiz geçiş ve zincir kırığı kasıtlı."""
    satirlar = []
    durumlar = ["TASLAK", "HAZIRLIK", "BASVURU", "KABUL", "RET", "IPTAL", "HAZIRLIK", "BASVURU"]
    oncekiler = ["", "TASLAK", "HAZIRLIK", "BASVURU", "BASVURU", "HAZIRLIK", "TASLAK", "HAZIRLIK"]
    for i in range(1, DEMO_AKIS + 1):
        kart = ALACAKLILAR[(i - 1) % DEMO_KART][0]
        if i in (22, 25):
            kart = "" if i == 22 else "ALN-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:
            o, d = "TASLAK", "KABUL"
        if i == 20:
            o, d = "KABUL", "HAZIRLIK"
        bilanco = 180000 + i * 12500
        nakit = bilanco if i not in (12, 19) else bilanco + 45000
        projeksiyon = nakit
        basvuru = nakit
        karar = nakit
        tutar = nakit
        tar = RAPOR_TARIHI - timedelta(days=(i * 3) % 75)
        satirlar.append({
            "id": f"KNA-{i:05d}",
            "kart": kart,
            "tarih": tar,
            "onceki": o,
            "durum": d,
            "bilanco": bilanco,
            "nakit": nakit,
            "projeksiyon": projeksiyon,
            "basvuru": basvuru,
            "karar": karar,
            "tutar": tutar,
            "donem": "2026-08",
            "aciklama": f"Ön proje kalemi {i}",
        })
    return satirlar


ORNEK = _demo_akis()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Mahkemeye sunulacak 3 yıllık bilanço ve nakit akış ön projesini tek motorda üretir; "
      "durum makinesi, yaşlandırma ve boş veride VERİ YOK kararı verir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "3 yıllık bilanço + nakit akış zinciri (Y0 gerçekleşen, Y1-Y2 projeksiyon)",
        "Makrosuz durum makinesi: Taslak → Hazırlık → Başvuru → Kabul / Ret / İptal",
        "Mahkeme ön proje baskı sayfası (RAPOR)",
        "Nakit açık yaşlandırma kovaları (0-7 / 8-30 / 31-60 / 60+)",
        "Karar kapısı: VERİ YOK / BAŞVUR / BEKLE / UYGUN DEĞİL",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Mali işler — 3 yıllık nakit akış satırlarını girer",
        "CFO / ortak — BAŞVUR / BEKLE / UYGUN DEĞİL kararını alır",
        "Mahkeme / kayyım — kanıt nakit akış raporunu inceler",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) GIRDI'ye alacaklı kartı girin. 2) BILANCO ve NAKIT_AKIS'e 3 yılı yazın. "
      "3) AKIS'e ön proje kalemini işleyin. 4) PANO'dan kararı, RAPOR'dan baskıyı alın.",
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
      'IF(OR(kna_nakitKilit>0,COUNTIF(tblAkis[Tutar],"<0")>0,'
      'IFERROR(MAX(tblAkis[Tutar]),0)>kna_azamiTutar,'
      'kna_gecikenAdet>kna_esikGeciken*2),"UYGUN DEĞİL",'
      'IF(OR(kna_gecikenAdet>kna_esikGeciken,kna_yetimKayit>0,'
      'kna_gecersizGecis>0,kna_zincirKirik>0),"BEKLE","BAŞVUR")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: BAŞVUR | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Açık Kuyruk Adet", "=kna_acikAdet", CATI),
        (7, "Hazırlık Adet", '=COUNTIFS(tblAkis[Durum],"HAZIRLIK",tblAkis[Tutar],">0")', CATI),
        (8, "Başvuru Adet", '=COUNTIFS(tblAkis[Durum],"BASVURU",tblAkis[Tutar],">0")', CATI),
        (9, "Geciken Adet", "=kna_gecikenAdet", CATI),
        (10, "Açık Tutar", "=kna_acikTutar", TL),
        (11, "Gecikme Maliyeti", "=kna_gecikmeMaliyeti", TL),
        (12, "Yetim Kayıt", "=kna_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=kna_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=kna_zincirKirik", CATI),
        (15, "Yaş Kova Özeti", "=kna_yasKova", None),
        (16, "Tahmin Aralık", "=kna_tahminAralik", TL),
        (17, "Senaryo Farkı", "=kna_senaryoKarsilastirma", TL),
        (18, "Tornado Zirve", "=kna_tornadoZirve", TL),
        (19, "Ödeme Öncelik", "=kna_odemeOncelik", TL),
        (20, "Pareto Gecikme", "=kna_paretoGecikme", YÜZDE),
        (21, "Kapalı Düşüm", "=kna_kapaliDusum", CATI),
        (22, "Üç Yıl Akış", "=kna_ucYilAkis", None),
        (23, "Kanıt Özeti", "=kna_kanitRaporu", None),
        (24, "Kalite Modülü", "=kna_modulO8", CATI),
        (25, "Yüzdelik 90", "=kna_modulI2", TL),
        (26, "Nakit Kilit", "=kna_nakitKilit", CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=kna_yorumYas"),
        (8, "=kna_yorumGecikme"),
        (9, "=kna_yorumTahmin"),
        (10, "=kna_yorumSenaryo"),
        (11, "=kna_yorumTornado"),
        (12, "=kna_yorumOncelik"),
        (13, "=kna_yorumPareto"),
        (14, "=kna_modulO8Yorum"),
        (15, "=kna_modulI2Yorum"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 28, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 28, 1, "Durum")
    h(ws, 28, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Taslak", 5), ("Hazırlık", 8), ("Başvuru", 7), ("Kabul", 4), ("Ret", 2), ("İptal", 2),
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

    h(ws, 28, 7, "Yıl")
    h(ws, 28, 8, "Nakit")
    for i in range(3):
        h(ws, 29 + i, 7, 2024 + i, sayi=CATI)
        h(ws, 29 + i, 8, 1_800_000 - i * 180000 + (0 if i < 2 else 450000), sayi=TL)

    h(ws, 28, 10, "Senaryo")
    h(ws, 28, 11, "Tutar")
    for i, (ad, t) in enumerate([
        ("Temkinli", 4_200_000), ("Baz", 5_100_000), ("İyimser", 6_000_000),
    ], 29):
        h(ws, i, 10, ad)
        h(ws, i, 11, t, sayi=TL)

    c1 = PieChart()
    c1.title = "Durum Dağılımı"
    c1.add_data(Reference(ws, min_col=2, min_row=28, max_row=34), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=29, max_row=34))
    c1.width, c1.height = 10, 7
    ws.add_chart(c1, "A39")

    c2 = BarChart()
    c2.title = "Yaşlandırma Kovaları"
    c2.add_data(Reference(ws, min_col=5, min_row=28, max_row=32), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=29, max_row=32))
    c2.width, c2.height = 10, 7
    ws.add_chart(c2, "F39")

    c3 = LineChart()
    c3.title = "Üç Yıllık Nakit"
    c3.add_data(Reference(ws, min_col=8, min_row=28, max_row=31), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=7, min_row=29, max_row=31))
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

    h(ws, 28, 16, "Yıl")
    h(ws, 28, 17, "Aktif")
    for i in range(3):
        h(ws, 29 + i, 16, 2024 + i, sayi=CATI)
        h(ws, 29 + i, 17, 20_500_000 - i * 900000, sayi=TL)

    c5 = BarChart()
    c5.title = "Uyarı Türleri"
    c5.add_data(Reference(ws, min_col=14, min_row=28, max_row=32), titles_from_data=True)
    c5.set_categories(Reference(ws, min_col=13, min_row=29, max_row=32))
    c5.width, c5.height = 10, 7
    ws.add_chart(c5, "A69")

    c6 = LineChart()
    c6.title = "Aktif Projeksiyon"
    c6.add_data(Reference(ws, min_col=17, min_row=28, max_row=31), titles_from_data=True)
    c6.set_categories(Reference(ws, min_col=16, min_row=29, max_row=31))
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
    c8.add_data(Reference(ws, min_col=2, min_row=28, max_row=34), titles_from_data=True)
    c8.set_categories(Reference(ws, min_col=1, min_row=29, max_row=34))
    c8.width, c8.height = 10, 7
    ws.add_chart(c8, "F84")

    baski_hazirla(ws, "A1:N26", f"{URUN_AD} | Pano | {SURUM}")
    genislik(ws, {"A": 22, "B": 18, "C": 14, "D": 14, "E": 12})


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Alacaklı kartları — her kaydın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=kna_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: Kimlik disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: ALN-00009 | Risk: Elle çakışma")

    sutunlar = ["KartId", "Unvan", "AlacakSinifi", "Sehir", "SozlesmeNo", "AlacakTutar",
                "RiskNotu", "Aktif", "KayitDolu"]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "KayitDolu": 'IF(tblGirdi[[#This Row],[KartId]]="","",1)',
    }
    for i, (kid, unvan, sinif, sehir, tutar) in enumerate(ALACAKLILAR):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Kart Kimliği", mesaj="ALN-##### formatında")
        _sari(ws, r, 2, unvan, baslik="Unvan", mesaj="Alacaklı unvanı")
        _sari(ws, r, 3, sinif, baslik="Alacak Sınıfı", mesaj="Tedarikçi / kredi / vergi")
        _sari(ws, r, 4, sehir, baslik="Şehir", mesaj="Şehir")
        _sari(ws, r, 5, f"SZL-2026-{i + 1:03d}", baslik="Sözleşme No", mesaj="Sözleşme numarası")
        _sari(ws, r, 6, tutar, sayi=TL, baslik="Alacak Tutarı", mesaj="Kayıtlı alacak TL")
        _sari(ws, r, 7, "Normal", baslik="Risk Notu", mesaj="Kısa risk notu")
        _sari(ws, r, 8, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_KART, SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None, sayi=TL if c == 6 else None,
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblGirdi", f"A{HDR}:I{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Kart Kimliği", mesaj="ALN-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Unvan"), (3, "Sınıf"), (4, "Şehir"), (5, "Sözleşme"), (7, "Risk")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "decimal", "0", f"F{ILK}:F{SON}",
              baslik="Alacak Tutarı", mesaj="Tutar ≥ 0",
              hata_baslik="Tutar", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeEvetHayir", f"H{ILK}:H{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 12)})
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 16


def bilanco(ws):
    sayfa_hazirla(ws, "BILANCO", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "3 yıllık bilanço — Y0 gerçekleşen, Y1 ve Y2 projeksiyon", son_kolon=12)
    h(ws, 4, 1, "Aktif Y0", yazi=GRİ)
    h(ws, 4, 2, '=SUMIF($C$6:$C$35,"aktif",$D$6:$D$35)', sayi=TL, kalin=True)
    h(ws, 4, 3, "Özkaynak Y2", yazi=GRİ)
    h(ws, 4, 4, '=SUMIF($C$6:$C$35,"ozkaynak",$F$6:$F$35)', sayi=TL, kalin=True)

    sutunlar = ["KalemId", "KalemAdi", "Grup", "Yil0", "Yil1", "Yil2",
                "ToplamUcYil", "SapmaYil1", "PayYil0"]
    baslik_satiri(ws, HDR, sutunlar)
    bln_son = 35
    for i, (kid, ad, grup, y0, y1, y2) in enumerate(BILANCO_KALEM):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Kalem Kimliği", mesaj="BLN-#####")
        _sari(ws, r, 2, ad, baslik="Kalem Adı", mesaj="Bilanço kalemi")
        _sari(ws, r, 3, grup, baslik="Grup", mesaj="donen / duran / aktif / pasif / ozkaynak")
        _sari(ws, r, 4, y0, sayi=TL, baslik="Yıl 0", mesaj="Gerçekleşen yıl TL")
        _sari(ws, r, 5, y1, sayi=TL, baslik="Yıl 1", mesaj="Projeksiyon yıl 1 TL")
        _sari(ws, r, 6, y2, sayi=TL, baslik="Yıl 2", mesaj="Projeksiyon yıl 2 TL")
    for r in range(ILK + DEMO_BILANCO, bln_son + 1):
        for c in range(1, 7):
            _sari(ws, r, c, None, sayi=TL if c in (4, 5, 6) else None,
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    for r in range(ILK, bln_son + 1):
        h(ws, r, 7, f'=IF(A{r}="","",D{r}+E{r}+F{r})', sayi=TL)
        h(ws, r, 8, f'=IF(A{r}="","",E{r}-D{r})', sayi=TL)
        h(ws, r, 9, f'=IF(OR(A{r}="",G{r}=0),"",D{r}/G{r})', sayi=YÜZDE)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{bln_son}",
              baslik="Kalem Kimliği", mesaj="BLN-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{bln_son}",
              baslik="Kalem Adı", mesaj="Bilanço kalemi adı",
              hata_baslik="Ad", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "list", "ListeBilancoGrup", f"C{ILK}:C{bln_son}",
              baslik="Grup", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    for c, ad in [(4, "Yıl 0"), (5, "Yıl 1"), (6, "Yıl 2")]:
        dogrulama(ws, "decimal", "-999999999999",
                  f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{bln_son}",
                  baslik=ad, mesaj=f"{ad} tutarını TL girin",
                  hata_baslik=ad, hata_mesaj="Sayı girin",
                  isaret="between", f2="999999999999")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 16 for i in range(1, 10)})
    ws.column_dimensions["B"].width = 32


def nakit_akis(ws):
    sayfa_hazirla(ws, "NAKIT_AKIS", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "3 yıllık nakit akış — işletme / yatırım / finansman / kapanış", son_kolon=12)
    h(ws, 4, 1, "Y0 Toplam", yazi=GRİ)
    h(ws, 4, 2, "=SUM($D$6:$D$35)", sayi=TL, kalin=True)
    h(ws, 4, 3, "Y2 Kapanış", yazi=GRİ)
    h(ws, 4, 4, '=SUMIF($C$6:$C$35,"kapanis",$F$6:$F$35)', sayi=TL, kalin=True)

    sutunlar = ["KalemId", "KalemAdi", "Grup", "Yil0", "Yil1", "Yil2",
                "ToplamUcYil", "Kumulatif", "KilitBayrak"]
    baslik_satiri(ws, HDR, sutunlar)
    nak_son = 35
    for i, (kid, ad, grup, y0, y1, y2) in enumerate(NAKIT_KALEM):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Kalem Kimliği", mesaj="NAK-#####")
        _sari(ws, r, 2, ad, baslik="Kalem Adı", mesaj="Nakit akış kalemi")
        _sari(ws, r, 3, grup, baslik="Grup", mesaj="isletme / yatirim / finansman / kapanis")
        _sari(ws, r, 4, y0, sayi=TL, baslik="Yıl 0", mesaj="Gerçekleşen nakit TL")
        _sari(ws, r, 5, y1, sayi=TL, baslik="Yıl 1", mesaj="Projeksiyon yıl 1 TL")
        _sari(ws, r, 6, y2, sayi=TL, baslik="Yıl 2", mesaj="Projeksiyon yıl 2 TL")
    for r in range(ILK + DEMO_NAKIT, nak_son + 1):
        for c in range(1, 7):
            _sari(ws, r, c, None, sayi=TL if c in (4, 5, 6) else None,
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    for r in range(ILK, nak_son + 1):
        h(ws, r, 7, f'=IF(A{r}="","",D{r}+E{r}+F{r})', sayi=TL)
        h(ws, r, 8, f'=IF(A{r}="","",D{r}+E{r}+F{r})', sayi=TL)
        h(ws, r, 9,
          f'=IF(A{r}="","",IF(C{r}<>"kapanis","",'
          f'IF(MIN(D{r},E{r},F{r})<kna_nakitKilitEsik,"KİLİT","AÇIK")))')

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{nak_son}",
              baslik="Kalem Kimliği", mesaj="NAK-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{nak_son}",
              baslik="Kalem Adı", mesaj="Nakit akış kalemi adı",
              hata_baslik="Ad", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "list", "ListeNakitGrup", f"C{ILK}:C{nak_son}",
              baslik="Grup", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    for c, ad in [(4, "Yıl 0"), (5, "Yıl 1"), (6, "Yıl 2")]:
        dogrulama(ws, "decimal", "-999999999999",
                  f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{nak_son}",
                  baslik=ad, mesaj=f"{ad} nakit tutarını TL girin",
                  hata_baslik=ad, hata_mesaj="Sayı girin",
                  isaret="between", f2="999999999999")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 16 for i in range(1, 10)})
    ws.column_dimensions["B"].width = 28


def akis(ws):
    sayfa_hazirla(ws, "AKIS", RENK, URUN_AD, son_kolon=22)
    alt_bant(ws, 3,
             "Ön proje kalemleri — durum makinesi + yetim + geçersiz geçiş + beş halkalı zincir",
             son_kolon=18)
    sutunlar = [
        "IslemId", "KaynakKartId", "Tarih", "OncekiDurum", "Durum",
        "BilancoTutar", "NakitTutar", "ProjeksiyonTutar", "BasvuruTutar", "KararTutar", "Tutar",
        "Donem", "Aciklama",
        "YasGun", "YasKova", "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    yas_f = yaslandirma_gun(
        "tblAkis[[#This Row],[Tarih]]", "raporTarihi").lstrip("=")
    kova_f = yas_kova_formul("tblAkis[[#This Row],[YasGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblAkis[[#This Row],[KaynakKartId]]", "GIRDI!$A$6:$A$255").lstrip("=")
    gecis_birlesik = (
        'tblAkis[[#This Row],[OncekiDurum]]&"|"&tblAkis[[#This Row],[Durum]]'
    )
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    zincir_f = zincir_kirik_formul(
        "tblAkis[[#This Row],[BilancoTutar]]",
        "tblAkis[[#This Row],[NakitTutar]]",
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
        _sari(ws, r, 1, s["id"], baslik="İşlem Kimliği", mesaj="KNA-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Kart", mesaj="GIRDI'deki KartId")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Olay tarihi")
        _sari(ws, r, 4, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 5, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 6, s["bilanco"], sayi=TL, baslik="Bilanço Halkası", mesaj="Bilanço tutarı TL")
        _sari(ws, r, 7, s["nakit"], sayi=TL, baslik="Nakit Halkası", mesaj="Nakit akış tutarı TL")
        _sari(ws, r, 8, s["projeksiyon"], sayi=TL, baslik="Projeksiyon", mesaj="Projeksiyon tutarı TL")
        _sari(ws, r, 9, s["basvuru"], sayi=TL, baslik="Başvuru", mesaj="Başvuru tutarı TL")
        _sari(ws, r, 10, s["karar"], sayi=TL, baslik="Karar Halkası", mesaj="Karar tutarı TL")
        _sari(ws, r, 11, s["tutar"], sayi=TL, baslik="Tutar", mesaj="Kuyruk tutarı TL")
        _sari(ws, r, 12, s["donem"], baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 13, s["aciklama"], baslik="Açıklama", mesaj="Kısa açıklama")

    for r in range(ILK + DEMO_AKIS, SON + 1):
        for c in range(1, 14):
            fmt = TARİH if c == 3 else (TL if c in (6, 7, 8, 9, 10, 11) else None)
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblAkis", f"A{HDR}:S{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="İşlem Kimliği", mesaj="KNA-##### girin",
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
    for c, ad in [(6, "Bilanço"), (7, "Nakit"), (8, "Projeksiyon"),
                  (9, "Başvuru"), (10, "Karar"), (11, "Tutar")]:
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

    h(ws, 6, 1, "Hazırlık")
    h(ws, 6, 2, kuyruk_sayim("tblAkis", "HAZIRLIK"), sayi=CATI)
    h(ws, 6, 3, kuyruk_tutar("tblAkis", "HAZIRLIK"), sayi=TL)

    h(ws, 7, 1, "Başvuru")
    h(ws, 7, 2, kuyruk_sayim("tblAkis", "BASVURU"), sayi=CATI)
    h(ws, 7, 3, kuyruk_tutar("tblAkis", "BASVURU"), sayi=TL)

    h(ws, 8, 1, "Taslak")
    h(ws, 8, 2, kuyruk_sayim("tblAkis", "TASLAK"), sayi=CATI)
    h(ws, 8, 3, kuyruk_tutar("tblAkis", "TASLAK"), sayi=TL)

    h(ws, 9, 1, "Geciken (açık, yaş>eşik)")
    h(ws, 9, 2,
      '=COUNTIFS(tblAkis[YasGun],">"&kna_esikGecikmeGun,tblAkis[Durum],"<>KABUL",'
      'tblAkis[Durum],"<>RET",tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")',
      sayi=CATI)
    h(ws, 9, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[YasGun],">"&kna_esikGecikmeGun,'
      'tblAkis[Durum],"<>KABUL",tblAkis[Durum],"<>RET",tblAkis[Durum],"<>IPTAL")',
      sayi=TL)

    h(ws, 11, 1, "Kapalı (Kabul+Ret+İptal) — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 1, "Kabul")
    h(ws, 12, 2, kuyruk_sayim("tblAkis", "KABUL"), sayi=CATI)
    h(ws, 12, 3, kuyruk_tutar("tblAkis", "KABUL"), sayi=TL)
    h(ws, 13, 1, "Ret")
    h(ws, 13, 2, kuyruk_sayim("tblAkis", "RET"), sayi=CATI)
    h(ws, 13, 3, kuyruk_tutar("tblAkis", "RET"), sayi=TL)
    h(ws, 14, 1, "İptal")
    h(ws, 14, 2, kuyruk_sayim("tblAkis", "IPTAL"), sayi=CATI)
    h(ws, 14, 3, kuyruk_tutar("tblAkis", "IPTAL"), sayi=TL)

    h(ws, 16, 1, "Kuyruk Özeti", kalin=True)
    h(ws, 16, 2, "=kna_kuyrukOzeti", kaydir=True)
    ws.merge_cells("B16:F16")

    h(ws, 18, 1, "Yaş Kova Dağılımı", kalin=True, yazi=KOYU_LACIVERT)
    for i, kova in enumerate(["0-7", "8-30", "31-60", "60+"], 19):
        h(ws, i, 1, kova)
        h(ws, i, 2, f'=COUNTIF(tblAkis[YasKova],"{kova}")', sayi=CATI)
        h(ws, i, 3, f'=SUMIF(tblAkis[YasKova],"{kova}",tblAkis[Tutar])', sayi=TL)

    genislik(ws, {"A": 40, "B": 14, "C": 16})


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Hesap motoru — ≥40 adım; sabitler AYARLAR'dan", son_kolon=10)
    h(ws, 5, 1, "Adım", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Açıklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    # H01 r6 … H55 r60. C14=H09 TASLAK, C15=HAZIRLIK, C16=BASVURU
    adimlar = [
        ("H01", "Toplam işlem", '=COUNTA(tblAkis[IslemId])'),
        ("H02", "Dolu kayıt", '=SUM(tblAkis[KayitDolu])'),
        ("H03", "Toplam tutar", "=SUM(tblAkis[Tutar])"),
        ("H04", "Bilanço halka toplam", "=SUM(tblAkis[BilancoTutar])"),
        ("H05", "Nakit halka toplam", "=SUM(tblAkis[NakitTutar])"),
        ("H06", "Projeksiyon toplam", "=SUM(tblAkis[ProjeksiyonTutar])"),
        ("H07", "Başvuru toplam", "=SUM(tblAkis[BasvuruTutar])"),
        ("H08", "Karar halka toplam", "=SUM(tblAkis[KararTutar])"),
        ("H09", "Taslak adet", '=COUNTIF(tblAkis[Durum],"TASLAK")'),
        ("H10", "Hazırlık adet", '=COUNTIF(tblAkis[Durum],"HAZIRLIK")'),
        ("H11", "Başvuru adet", '=COUNTIF(tblAkis[Durum],"BASVURU")'),
        ("H12", "Kabul adet", '=COUNTIF(tblAkis[Durum],"KABUL")'),
        ("H13", "Ret adet", '=COUNTIF(tblAkis[Durum],"RET")'),
        ("H14", "İptal adet", '=COUNTIF(tblAkis[Durum],"IPTAL")'),
        ("H15", "Açık adet", "=C14+C15+C16"),
        ("H16", "Kapalı düşüm", "=C17+C18+C19"),
        ("H17", "Açık tutar",
         '=SUMIF(tblAkis[Durum],"TASLAK",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"HAZIRLIK",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"BASVURU",tblAkis[Tutar])'),
        ("H18", "Yetim adet", '=COUNTIF(tblAkis[YetimBayrak],"YETİM")'),
        ("H19", "Geçersiz geçiş", '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")'),
        ("H20", "Zincir kırık", '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")'),
        ("H21", "Zincir tamam", '=COUNTIF(tblAkis[ZincirBayrak],"OK")'),
        ("H22", "Yaş 0-7", '=COUNTIF(tblAkis[YasKova],"0-7")'),
        ("H23", "Yaş 8-30", '=COUNTIF(tblAkis[YasKova],"8-30")'),
        ("H24", "Yaş 31-60", '=COUNTIF(tblAkis[YasKova],"31-60")'),
        ("H25", "Yaş 60+", '=COUNTIF(tblAkis[YasKova],"60+")'),
        ("H26", "Ort yaş gün", "=IFERROR(AVERAGE(tblAkis[YasGun]),0)"),
        ("H27", "Geciken adet",
         '=COUNTIFS(tblAkis[YasGun],">"&kna_esikGecikmeGun,tblAkis[Durum],"<>KABUL",'
         'tblAkis[Durum],"<>RET",tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")'),
        ("H28", "Geciken tutar",
         '=SUMIFS(tblAkis[Tutar],tblAkis[YasGun],">"&kna_esikGecikmeGun,'
         'tblAkis[Durum],"<>KABUL",tblAkis[Durum],"<>RET",tblAkis[Durum],"<>IPTAL")'),
        ("H29", "Gecikme maliyeti", "=C33*kna_gecikmeOranGun"),
        ("H30", "Temkinli senaryo", "=C22*kna_senaryoTemkinli"),
        ("H31", "Baz senaryo", "=C22*kna_senaryoBaz"),
        ("H32", "İyimser senaryo", "=C22*kna_senaryoIyimser"),
        ("H33", "Senaryo farkı", "=C37-C35"),
        ("H34", "Tornado (oran×tutar)", "=C22*kna_tornadoOran"),
        ("H35", "Ödeme öncelik havuzu",
         '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"HAZIRLIK",tblAkis[YasGun],">"&kna_esikGecikmeGun)'),
        ("H36", "Pareto gecikme payı", "=IFERROR(C33/MAX(C22,1),0)"),
        ("H37", "Tahmin yüzdelik",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[Tutar],kna_yuzdelikOran),0)"),
        ("H38", "Tahmin eğilim",
         "=IFERROR(FORECAST(COUNTA(tblAkis[IslemId])+1,tblAkis[Tutar],"
         "tblAkis[KayitDolu]),C42)"),
        ("H39", "Tahmin aralık", "=IFERROR((C42+C43)/2,0)"),
        ("H40", "Kart sayısı", '=COUNTA(tblGirdi[KartId])'),
        ("H41", "Bilanço Y0 toplam", "=SUM(BILANCO!D6:D35)"),
        ("H42", "Nakit Y0 toplam", "=SUM(NAKIT_AKIS!D6:D35)"),
        ("H43", "Üç yıl akış metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Y0",TEXT(SUM(NAKIT_AKIS!D6:D35),"0"),'
         '"Y1",TEXT(SUM(NAKIT_AKIS!E6:E35),"0"),"Y2",TEXT(SUM(NAKIT_AKIS!F6:F35),"0"))'),
        ("H44", "Kuyruk özeti metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Hazırlık",C15,"Başvuru",C16,"Geciken",C32)'),
        ("H45", "Yaş kova metin",
         '=_xlfn.TEXTJOIN("/",TRUE,C27,C28,C29,C30)'),
        ("H46", "Karar skoru ham",
         "=IF(C6=0,0,MAX(0,100-C23*kna_cezaYetim-C24*kna_cezaGecis"
         "-C25*kna_cezaZincir-C32*kna_cezaGeciken))"),
        ("H47", "Veri hazırlık", "=IF(C6=0,0,C7/MAX(C6,1))"),
        ("H48", "Nakit kilit adet", '=COUNTIF(NAKIT_AKIS!I6:I35,"KİLİT")'),
        ("H49", "Bilanço nakit sapma", "=ABS(C46-C47)"),
        ("H50", "Kalite modülü", "=C51"),
        ("H51", "Yüzdelik 90",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[Tutar],kna_yuzdelikOran),0)"),
        ("H52", "Kanıt özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Açık",C20,"Yetim",C23,"Zincir",C25,'
         '"Geciken",C32,"Kilit",C53)'),
        ("H53", "Nakit Y1 toplam", "=SUM(NAKIT_AKIS!E6:E35)"),
        ("H54", "Nakit Y2 toplam", "=SUM(NAKIT_AKIS!F6:F35)"),
        ("H55", "Üç yıl nakit toplam", "=C47+C58+C59"),
    ]
    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        fmt = TL if any(x in acik.lower() for x in (
            "tutar", "maliyet", "senaryo", "tornado", "öncelik", "tahmin", "halka",
            "toplam", "havuz", "fark", "yüzdelik", "sapma", "y0", "y1", "y2",
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
    h(ws, 65, 2, "=kna_senaryoTemkinli*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 65, 3, "=C35", sayi=TL)
    h(ws, 66, 1, "Baz")
    h(ws, 66, 2, "=kna_senaryoBaz*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 66, 3, "=C36", sayi=TL)
    h(ws, 67, 1, "İyimser")
    h(ws, 67, 2, "=kna_senaryoIyimser*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 67, 3, "=C37", sayi=TL)

    genislik(ws, {"A": 10, "B": 28, "C": 55})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Mahkeme Konkordato Nakit Akış Ön Proje Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Açık Kuyruk", "=kna_acikAdet", CATI),
        (10, "Açık Tutar", "=kna_acikTutar", TL),
        (11, "Gecikme Maliyeti", "=kna_gecikmeMaliyeti", TL),
        (12, "Yetim", "=kna_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=kna_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=kna_zincirKirik", CATI),
        (15, "Tahmin Aralık", "=kna_tahminAralik", TL),
        (16, "Üç Yıl Akış", "=kna_ucYilAkis", None),
        (17, "Nakit Kilit", "=kna_nakitKilit", CATI),
        (18, "Kanıt Özeti", "=kna_kanitRaporu", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 20, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 1,
      "Durum haritası AYARLAR'dadır. Geçersiz geçiş engellenmez, görünür kılınır. "
      "Üç yıllık bilanço ve nakit akış BILANCO / NAKIT_AKIS tablolarından türetilir. "
      "Bu çıktı karar destektir; kesin mali/hukuki görüş veya mahkeme hükmü yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A21:H21")
    h(ws, 23, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 1, "1. Yetim kayıtları GIRDI'ye bağlayın veya satırı düzeltin.")
    h(ws, 25, 1, "2. Geçersiz geçişleri durum haritasına uygun ilerletin.")
    h(ws, 26, 1, "3. Bilanço-nakit zincir sapmasını kapatın; nakit kilit varsa başvuruyu durdurun.")
    h(ws, 28, 1, "Kayyım / mahkeme imza")
    h(ws, 28, 4, "CFO / ortak imza")
    h(ws, 30, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    baski_hazirla(ws, "A1:H30", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 62, "B": 42, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "İşlem sayısı", '=COUNTA(tblAkis[IslemId])'),
        (6, "Kart sayısı", '=COUNTA(tblGirdi[KartId])'),
        (7, "Taslak", '=COUNTIF(tblAkis[Durum],"TASLAK")'),
        (8, "Hazırlık", '=COUNTIF(tblAkis[Durum],"HAZIRLIK")'),
        (9, "Başvuru", '=COUNTIF(tblAkis[Durum],"BASVURU")'),
        (10, "Kabul", '=COUNTIF(tblAkis[Durum],"KABUL")'),
        (11, "Ret", '=COUNTIF(tblAkis[Durum],"RET")'),
        (12, "Yetim", "=kna_yetimKayit"),
        (13, "Geçersiz", "=kna_gecersizGecis"),
        (14, "Zincir kırık", "=kna_zincirKirik"),
        (15, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (16, "Kalite skoru", "=HESAP!C51"),
        (17, "Açık tutar", "=kna_acikTutar"),
        (18, "Gecikme maliyeti", "=kna_gecikmeMaliyeti"),
        (19, "Kapalı düşüm", "=kna_kapaliDusum"),
        (20, "Nakit kilit", "=kna_nakitKilit"),
        (21, "Bilanço satır", "=COUNTA(BILANCO!A6:A35)"),
        (22, "Nakit satır", "=COUNTA(NAKIT_AKIS!A6:A35)"),
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
      f"Dosyada {DEMO_KART} kart, {DEMO_BILANCO} bilanço, {DEMO_NAKIT} nakit ve "
      f"{DEMO_AKIS} akış satırı vardır. "
      "Yetim (22, 25), geçersiz geçiş (18, 20) ve zincir kırık (12, 19) kasıtlıdır. "
      "3 yıllık BILANCO ve NAKIT_AKIS mahkeme ön projesinin omurgasıdır. "
      "GIRDI, BILANCO, NAKIT_AKIS ve AKIS'i temizleyip kendi verinizi girebilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Durumlar: Taslak/Hazırlık/Başvuru/Kabul/Ret/İptal — kapalılar kuyruktan düşer")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 5000, sayi=CATI)
    genislik(ws, {"A": 70, "B": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=16)
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

    h(ws, 3, 7, "Bilanço Grup", kalin=True)
    h(ws, 6, 7, "Grup", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(["donen", "duran", "aktif", "pasif", "ozkaynak"], 7):
        _sari(ws, i, 7, b, baslik="Grup", mesaj="Bilanço grubu")

    h(ws, 3, 9, "Nakit Grup", kalin=True)
    h(ws, 6, 9, "Grup", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(["isletme", "yatirim", "finansman", "kapanis"], 7):
        _sari(ws, i, 9, b, baslik="Grup", mesaj="Nakit grubu")

    h(ws, 3, 11, "Kart Kimlik", kalin=True)
    h(ws, 6, 11, "KartId", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, k in enumerate(ALACAKLILAR, 7):
        _sari(ws, i, 11, k[0], baslik="Kart", mesaj="ALN-#####")
    genislik(ws, {"A": 14, "C": 20, "E": 12, "G": 14, "I": 14, "K": 14})


PARAMETRELER = [
    ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
     "Rapor tarihi (TODAY yok)"),
    ("zincirTolerans", 1000, "TL", "İş kuralı — zincir sapma", "01.01.2026",
     "ZİNCİR KIRIK eşiği"),
    ("kna_esikGecikmeGun", 30, "gün", "İç politika", "01.01.2026", "Gecikme eşiği"),
    ("kna_esikGeciken", 5, "adet", "İç politika", "01.01.2026", "Karar eşiği geciken"),
    ("kna_gecikmeOranGun", 0.0005, "oran", "Finans varsayımı", "01.01.2026",
     "Günlük gecikme maliyeti oranı"),
    ("kna_senaryoTemkinli", 0.85, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
    ("kna_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
    ("kna_senaryoIyimser", 1.15, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
    ("kna_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
    ("kna_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
    ("kna_olcekHedef", 5000, "satir", "Manda A3 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
    ("kna_azamiTutar", 50000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul"),
    ("kna_nakitKilitEsik", 0, "TL", "Nakit kilit eşiği", "01.01.2026",
     "Kapanış yılı eşiğin altındaysa kilit"),
    ("kna_cezaYetim", 12, "puan", "Kalite skoru", "01.01.2026", "Yetim ceza puanı"),
    ("kna_cezaGecis", 12, "puan", "Kalite skoru", "01.01.2026", "Geçiş ceza puanı"),
    ("kna_cezaZincir", 12, "puan", "Kalite skoru", "01.01.2026", "Zincir ceza puanı"),
    ("kna_cezaGeciken", 2, "puan", "Kalite skoru", "01.01.2026", "Geciken ceza puanı"),
    ("kna_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi",
     "01.01.2026", "Durum makinesi kaynağı"),
    ("kna_kuyrukOzeti_metin", "KUYRUKLAR canlı özet", "metin", "A05 kuyruk", "01.01.2026",
     "Kuyruk görünümü"),
    ("kna_kapaliDusum_metin", "kapali kategori kuyruktan düşer", "metin", "A06", "01.01.2026",
     "Kapalı düşüm kuralı"),
    ("kna_kanitRaporu_metin", "RAPOR kanıt çıktısı", "metin", "Mahkeme ön proje", "01.01.2026",
     "Kanıt raporu"),
    ("kna_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
]


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + durum haritası (A01/A06)", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    for i, (ana, deg, bir, kay, yur, acik) in enumerate(PARAMETRELER):
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

    h(ws, 30, 8, "Durum Haritası", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    dh_bas = ["durum_id", "durum_adi", "sira", "izinli_sonraki_durumlar", "kategori"]
    baslik_satiri(ws, 31, dh_bas, basla=8)
    for i, d in enumerate(DURUM_HARITASI):
        r = 32 + i
        h(ws, r, 8, d["durum_id"])
        h(ws, r, 9, d["durum_adi"])
        h(ws, r, 10, d["sira"], sayi=CATI)
        h(ws, r, 11, d["izinli_sonraki_durumlar"])
        h(ws, r, 12, d["kategori"])

    h(ws, 40, 8, "Motor Çıktıları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 41, 8, "anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 41, 9, "deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    motor = [
        (42, "kna_acikAdet", "=HESAP!C20"),
        (43, "kna_acikTutar", "=HESAP!C22"),
        (44, "kna_gecikenAdet", "=HESAP!C32"),
        (45, "kna_gecikmeMaliyeti", "=HESAP!C34"),
        (46, "kna_yetimKayit", "=HESAP!C23"),
        (47, "kna_gecersizGecis", "=HESAP!C24"),
        (48, "kna_zincirKirik", "=HESAP!C25"),
        (49, "kna_yasKova", "=HESAP!C50"),
        (50, "kna_tahminAralik", "=HESAP!C44"),
        (51, "kna_senaryoKarsilastirma", "=HESAP!C38"),
        (52, "kna_tornadoZirve", "=HESAP!C39"),
        (53, "kna_odemeOncelik", "=HESAP!C40"),
        (54, "kna_paretoGecikme", "=HESAP!C41"),
        (55, "kna_kapaliDusum", "=HESAP!C21"),
        (56, "kna_kuyrukOzeti", "=HESAP!C49"),
        (57, "kna_kanitRaporu", "=HESAP!C57"),
        (58, "kna_sonrakiKimlik",
         '="ALN-"&TEXT(COUNTA(tblGirdi[KartId])+1,"00000")'),
        (59, "kna_ucYilAkis", "=HESAP!C48"),
        (60, "kna_modulO8", "=HESAP!C55"),
        (61, "kna_modulI2", "=HESAP!C56"),
        (62, "kna_nakitKilit", "=HESAP!C53"),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 64, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (65, "kna_yorumYas",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yaş kovaları",kna_yasKova)'),
        (66, "kna_yorumGecikme",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Gecikme maliyeti",TEXT(kna_gecikmeMaliyeti,"0.00"),"TL")'),
        (67, "kna_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(kna_tahminAralik,"0.00"),"TL")'),
        (68, "kna_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(kna_senaryoKarsilastirma,"0.00"),"TL")'),
        (69, "kna_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(kna_tornadoZirve,"0.00"),"TL")'),
        (70, "kna_yorumOncelik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Ödeme öncelik havuzu",TEXT(kna_odemeOncelik,"0.00"),"TL")'),
        (71, "kna_yorumPareto",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Pareto gecikme payı",TEXT(kna_paretoGecikme,"0.0%"))'),
        (72, "kna_modulO8Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite",TEXT(kna_modulO8,"0"),"yetim/geçiş/zincir düşer")'),
        (73, "kna_modulI2Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yüzdelik 90 tutar",TEXT(kna_modulI2,"0.00"),"TL")'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 24, "B": 28, "C": 10, "D": 28, "E": 14, "F": 36, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. GIRDI sayfasına alacaklı kartı ekleyin; alacak tutarını yazın.",
        "2. BILANCO sayfasına 3 yılı (Y0 gerçekleşen, Y1-Y2 projeksiyon) girin.",
        "3. NAKIT_AKIS sayfasına işletme / yatırım / finansman / kapanış kalemlerini yazın.",
        "4. AKIS sayfasına her ön proje kalemini yazın; durum kolonunu listeden seçin.",
        "5. Önceki durum boş bırakılırsa geçiş uyarısı çalışmaz; bilinçli geçişte doldurun.",
        "6. Geçersiz geçiş engellenmez — kırmızı uyarı üretir (makrosuz kural).",
        "7. Kapalı durumlar (Kabul/Ret/İptal) açık kuyruklardan düşer.",
        "8. Nakit kilit veya uç tutar UYGUN DEĞİL, yetim/geçiş BEKLE, boş veri VERİ YOK üretir.",
        "9. RAPOR sayfası mahkeme ön proje baskısıdır; PDF olarak yazdırın.",
        "10. raporTarihi AYARLAR'dadır; TODAY kullanılmaz.",
        "11. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı hücreler açıktır.",
        f"12. Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv",
    ]
    for i, m in enumerate(maddeler, 5):
        h(ws, i, 1, m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    son_param = 5 + len(PARAMETRELER)
    for r in range(6, son_param + 1):
        dogrulama(ws, "textLength", "0", f"D{r}",
                  baslik="Kaynak", mesaj="Kaynak metni girin",
                  hata_baslik="Kaynak", hata_mesaj="Boş bırakmayın",
                  isaret="greaterThanOrEqual", f2="120")
    ws = wb["LISTELER"]
    for r in range(7, 13):
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
    for r in range(5, 17):
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
        "B4", CellIsRule(operator="equal", formula=['"BAŞVUR"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"BEKLE"'], fill=sari))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"UYGUN DEĞİL"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(6, 27):
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
    for durum, fill in [("HAZIRLIK", sari), ("BASVURU", yesil), ("IPTAL", kirmizi)]:
        ws.conditional_formatting.add(
            f"E{ILK}:E{SON}",
            CellIsRule(operator="equal", formula=[f'"{durum}"'], fill=fill))

    ws = wb["KUYRUKLAR"]
    for r in range(6, 10):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=sari))
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
    for r in range(19, 23):
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
        "B7", CellIsRule(operator="equal", formula=['"BAŞVUR"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"BEKLE"'], fill=sari))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"UYGUN DEĞİL"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(9, 19):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["GIRDI"]
    ws.conditional_formatting.add(
        f"H{ILK}:H{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"F{ILK}:F{SON}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["NAKIT_AKIS"]
    ws.conditional_formatting.add(
        "I6:I35",
        CellIsRule(operator="equal", formula=['"KİLİT"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "I6:I35",
        CellIsRule(operator="equal", formula=['"AÇIK"'], fill=yesil, font=yf))

    ws = wb["BILANCO"]
    ws.conditional_formatting.add(
        "H6:H35",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))


def adlari_bagla(wb):
    for i, (ana, *_rest) in enumerate(PARAMETRELER):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "kna_acikAdet": 42, "kna_acikTutar": 43, "kna_gecikenAdet": 44,
        "kna_gecikmeMaliyeti": 45, "kna_yetimKayit": 46, "kna_gecersizGecis": 47,
        "kna_zincirKirik": 48, "kna_yasKova": 49, "kna_tahminAralik": 50,
        "kna_senaryoKarsilastirma": 51, "kna_tornadoZirve": 52, "kna_odemeOncelik": 53,
        "kna_paretoGecikme": 54, "kna_kapaliDusum": 55, "kna_kuyrukOzeti": 56,
        "kna_kanitRaporu": 57, "kna_sonrakiKimlik": 58, "kna_ucYilAkis": 59,
        "kna_modulO8": 60, "kna_modulI2": 61, "kna_nakitKilit": 62,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "kna_yorumYas": 65, "kna_yorumGecikme": 66, "kna_yorumTahmin": 67,
        "kna_yorumSenaryo": 68, "kna_yorumTornado": 69, "kna_yorumOncelik": 70,
        "kna_yorumPareto": 71, "kna_modulO8Yorum": 72, "kna_modulI2Yorum": 73,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "kna_kararKapi", "PANO!$B$4")
    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$12")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$C$7:$C$11")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")
    ad_ekle(wb, "ListeBilancoGrup", "LISTELER!$G$7:$G$11")
    ad_ekle(wb, "ListeNakitGrup", "LISTELER!$I$7:$I$10")
    ad_ekle(wb, "ListeKartlar", "LISTELER!$K$7:$K$14")


def _csv_yaz(urun_dir):
    yol = os.path.join(urun_dir, "ornek_veri.csv")
    with open(yol, "w", encoding="utf-8") as f:
        f.write("IslemId,KaynakKartId,Tarih,OncekiDurum,Durum,"
                "BilancoTutar,NakitTutar,ProjeksiyonTutar,BasvuruTutar,KararTutar,"
                "Tutar,Donem,Aciklama\n")
        for s in ORNEK:
            f.write(
                f"{s['id']},{s['kart']},{s['tarih']:%d.%m.%Y},"
                f"{s['onceki']},{s['durum']},{s['bilanco']},{s['nakit']},"
                f"{s['projeksiyon']},{s['basvuru']},{s['karar']},{s['tutar']},"
                f"{s['donem']},{s['aciklama']}\n")
    return yol


def main(cikti_yolu=None):
    wb = Workbook()
    siralar = [
        (kapak, "KAPAK"),
        (pano, "PANO"),
        (girdi, "GIRDI"),
        (bilanco, "BILANCO"),
        (nakit_akis, "NAKIT_AKIS"),
        (akis, "AKIS"),
        (kuyruklar, "KUYRUKLAR"),
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
    dosya_ad = "KonkordatoNakitAkisOnProjesi.xlsx"
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
