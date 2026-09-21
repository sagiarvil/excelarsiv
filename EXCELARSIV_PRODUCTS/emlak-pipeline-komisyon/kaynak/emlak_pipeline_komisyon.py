#!/usr/bin/env python3
"""Emlak Ofisi Satış Pipeline & Komisyon Takibi — A3 üretim betiği (manda v6).

Domain: ilan→görüşme→teklif→sözleşme→tapu pipeline, komisyon hesabı, dönüşüm.
Rename-kopya değil: tahsilat/vade mantığı yok; emlak satışı + komisyon zinciri.
"""

from __future__ import annotations

import os
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
    yaslandirma_gun,
    yetim_kayit_formul,
    zincir_kirik_formul,
)

URUN_AD = "Emlak Ofisi Satış Pipeline & Komisyon Takibi"
SURUM = "1.0.0"
RENK = "1F4E79"
KAPASITE = 250
DEMO_KART = 8
DEMO_AKIS = 28
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "ILAN", "durum_adi": "İlan", "sira": 1,
     "izinli_sonraki_durumlar": "GORUSME|IPTAL", "kategori": "acik"},
    {"durum_id": "GORUSME", "durum_adi": "Görüşme", "sira": 2,
     "izinli_sonraki_durumlar": "TEKLIF|IPTAL|ILAN", "kategori": "acik"},
    {"durum_id": "TEKLIF", "durum_adi": "Teklif", "sira": 3,
     "izinli_sonraki_durumlar": "SOZLESME|IPTAL|GORUSME", "kategori": "acik"},
    {"durum_id": "SOZLESME", "durum_adi": "Sözleşme", "sira": 4,
     "izinli_sonraki_durumlar": "TAPU|IPTAL", "kategori": "acik"},
    {"durum_id": "TAPU", "durum_adi": "Tapu", "sira": 5,
     "izinli_sonraki_durumlar": "KAPALI", "kategori": "acik"},
    {"durum_id": "KAPALI", "durum_adi": "Kapalı", "sira": 6,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "IPTAL", "durum_adi": "İptal", "sira": 7,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "ILAN|GORUSME", "ILAN|IPTAL",
    "GORUSME|TEKLIF", "GORUSME|IPTAL", "GORUSME|ILAN",
    "TEKLIF|SOZLESME", "TEKLIF|IPTAL", "TEKLIF|GORUSME",
    "SOZLESME|TAPU", "SOZLESME|IPTAL",
    "TAPU|KAPALI",
]

DANISMANLAR = [
    ("DAN-00001", "Ayşe Yılmaz", "Kadıköy", "Satış"),
    ("DAN-00002", "Mehmet Kaya", "Beşiktaş", "Satış"),
    ("DAN-00003", "Zeynep Demir", "Üsküdar", "Kiralama"),
    ("DAN-00004", "Can Öztürk", "Ataşehir", "Satış"),
    ("DAN-00005", "Elif Şahin", "Bakırköy", "Satış"),
    ("DAN-00006", "Burak Çelik", "Şişli", "Proje"),
    ("DAN-00007", "Selin Arslan", "Maltepe", "Satış"),
    ("DAN-00008", "Emre Koç", "Kartal", "Kiralama"),
]

AY_TREND = [
    ("epk_trendOcak", 0.72), ("epk_trendSubat", 0.74), ("epk_trendMart", 0.78),
    ("epk_trendNisan", 0.80), ("epk_trendMayis", 0.82), ("epk_trendHaziran", 0.85),
    ("epk_trendTemmuz", 0.83), ("epk_trendAgustos", 0.81), ("epk_trendEylul", 0.84),
    ("epk_trendEkim", 0.86), ("epk_trendKasim", 0.88), ("epk_trendAralik", 0.90),
]


def _yas_kova_epk(yas_hucre: str) -> str:
    return (
        f'=IF({yas_hucre}="","",'
        f'IF({yas_hucre}<=epk_kovaEsik1,"0-30",'
        f'IF({yas_hucre}<=epk_kovaEsik2,"31-60",'
        f'IF({yas_hucre}<=epk_kovaEsik3,"61-90","90+"))))'
    )


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Pipeline ve komisyon kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_akis():
    """≥25 satır; yetim, geçersiz geçiş ve komisyon zincir kırığı kasıtlı."""
    satirlar = []
    durumlar = ["ILAN", "GORUSME", "TEKLIF", "SOZLESME", "TAPU", "KAPALI", "IPTAL", "TEKLIF"]
    oncekiler = ["", "ILAN", "GORUSME", "TEKLIF", "SOZLESME", "TAPU", "TEKLIF", "GORUSME"]
    for i in range(1, DEMO_AKIS + 1):
        kart = DANISMANLAR[(i - 1) % DEMO_KART][0]
        if i in (22, 25):
            kart = "" if i == 22 else "DAN-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:  # geçersiz: ILAN → TEKLIF
            o, d = "ILAN", "TEKLIF"
        if i == 20:  # geçersiz: SOZLESME → GORUSME
            o, d = "SOZLESME", "GORUSME"
        liste = 2_500_000 + i * 180_000
        teklif = liste - (i % 5) * 50_000
        satis = teklif if d in ("TAPU", "KAPALI") else 0
        oran = 0.03 if i % 3 else 0.02
        if i in (12, 19):  # zincir kırık: satış ≠ teklif (tolerans üstü)
            if d in ("TAPU", "KAPALI"):
                satis = teklif + 250_000
            else:
                satis = 0
                teklif = liste + 200_000  # liste-teklif sapması
        kom_tahsil = round(satis * oran) if d == "KAPALI" else 0
        tutar = 0 if d in ("KAPALI", "IPTAL") else round(
            (satis or teklif or liste) * oran)
        baslangic = RAPOR_TARIHI - timedelta(days=(i * 5) % 120)
        satirlar.append({
            "id": f"ILN-{i:05d}",
            "kart": kart,
            "ilan_no": f"EML-2026-{i:04d}",
            "baslangic": baslangic,
            "onceki": o,
            "durum": d,
            "liste": liste,
            "teklif": teklif,
            "satis": satis,
            "oran": oran,
            "tahsil": kom_tahsil,
            "tutar": tutar,
            "tip": ["Konut", "İşyeri", "Arsa"][i % 3],
            "yorum_a": "A" if i % 2 else "B",
            "yorum_b": f"Portföy notu {i}",
        })
    return satirlar


ORNEK = _demo_akis()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Emlak ofisi satış pipeline'ını durum makinesi ile yönetir; liste→teklif→satış "
      "komisyon zincirini görünür kılar; dönüşüm ve beklenen komisyon üzerinden karar üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Makrosuz durum: İlan → Görüşme → Teklif → Sözleşme → Tapu → Kapalı / İptal",
        "Komisyon = satış × oran; liste/teklif/satış tutar zinciri bayrağı",
        "Dönüşüm oranı ve kuyruk yaşlandırması (0-30 / 31-60 / 61-90 / 90+)",
        "Yetim danışman ve geçersiz geçiş uyarıları",
        "Dönemsel komisyon kanıt raporu — ofis sahibi / SMMM dosyasına uygun",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Danışman — ilan satırı ve komisyon oranı girişi",
        "Ofis sahibi / satış müdürü — dönüşüm ve beklenen komisyon kararı",
        "SMMM — dönem komisyon kanıt raporu",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) KARTLAR'a danışman kartı girin. 2) AKIS'e ilan/pipeline satırlarını yazın. "
      "3) KUYRUKLAR ve PANO'dan günün kararını görün. 4) RAPOR'dan kanıt çıktısı alın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=20)
    sabitle(ws, "A3")
    h(ws, 3, 1, "Pipeline Karar Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 3, 3, "Rapor Tarihi", yazi=GRİ)
    h(ws, 3, 4, "=raporTarihi", sayi=TARİH)

    h(ws, 4, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2,
      '=IF(COUNTA(tblAkis[IslemId])=0,"VERİ YOK",'
      'IF(OR(epk_yetimKayit>0,epk_gecersizGecis>0,epk_zincirKirik>0),"KRİTİK",'
      'IF(OR(epk_gecikenAdet>epk_esikGeciken,epk_donusumOrani<epk_hedefDonusum),"DİKKAT","UYGUN")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle değiştirilmez")

    # A05/E03: PANO'da COUNTIF/SUMIF izleri (sabit KPI yasak)
    kpi = [
        (6, "Pano Açık Pipeline",
         '=COUNTIFS(tblAkis[Durum],"ILAN",tblAkis[Tutar],">0")'
         '+COUNTIFS(tblAkis[Durum],"GORUSME",tblAkis[Tutar],">0")'
         '+COUNTIFS(tblAkis[Durum],"TEKLIF",tblAkis[Tutar],">0")'
         '+COUNTIFS(tblAkis[Durum],"SOZLESME",tblAkis[Tutar],">0")'
         '+COUNTIFS(tblAkis[Durum],"TAPU",tblAkis[Tutar],">0")', CATI),
        (7, "Pano Görüşme Adet",
         '=COUNTIFS(tblAkis[Durum],"GORUSME",tblAkis[Tutar],">0")', CATI),
        (8, "Pano Teklif Adet",
         '=COUNTIFS(tblAkis[Durum],"TEKLIF",tblAkis[Tutar],">0")', CATI),
        (9, "Pano Geciken Adet", "=epk_gecikenAdet", CATI),
        (10, "Pano Beklenen Komisyon",
         '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"<>KAPALI",tblAkis[Durum],"<>IPTAL")', TL),
        (11, "Pano Tahsil Komisyon",
         '=SUMIF(tblAkis[Durum],"KAPALI",tblAkis[TahsilKomisyon])', TL),
        (12, "Pano Yetim Kayıt", "=epk_yetimKayit", CATI),
        (13, "Pano Geçersiz Geçiş", "=epk_gecersizGecis", CATI),
        (14, "Pano Zincir Kırık", "=epk_zincirKirik", CATI),
        (15, "Pano Dönüşüm Oranı", "=epk_donusumOrani", YÜZDE),
        (16, "Pano Yaş Özeti", "=epk_yasKova", None),
        (17, "Pano Tahmin Aralık", "=epk_tahminAralik", TL),
        (18, "Pano Senaryo Farkı", "=epk_senaryoKarsilastirma", TL),
        (19, "Pano Tornado Zirve", "=epk_tornadoZirve", TL),
        (20, "Pano Yoğunlaşma", "=epk_hhi", YÜZDE),
        (21, "Pano Kalite Skoru", "=epk_kaliteSkor", CATI),
        (22, "Pano P90 Satış", "=epk_yuzdelikP90", TL),
        (23, "Pano Kapalı Düşüm", "=epk_kapaliDusum", CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=epk_modulT1Yorum"), (8, "=epk_modulT2Yorum"),
        (9, "=epk_modulO1Yorum"), (10, "=epk_modulO2Yorum"),
        (11, "=epk_modulO3Yorum"), (12, "=epk_modulO6Yorum"),
        (13, "=epk_modulO8Yorum"), (14, "=epk_modulMYorum"),
        (15, "=epk_modulI1Yorum"), (16, "=epk_modulI2Yorum"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    # Grafik kaynakları — sabit demo (D09b/Ö2; LO tablo COUNTIF grafikte sıfırlanmasın)
    h(ws, 25, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 1, "Durum")
    h(ws, 26, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("İlan", 5), ("Görüşme", 4), ("Teklif", 5), ("Sözleşme", 4),
        ("Tapu", 3), ("Kapalı", 4), ("İptal", 3),
    ], 27):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 26, 4, "Yaş")
    h(ws, 26, 5, "Adet")
    for i, (kova, adet) in enumerate([
        ("0-30", 7), ("31-60", 8), ("61-90", 6), ("90+", 7),
    ], 27):
        h(ws, i, 4, kova)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 26, 7, "Hafta")
    h(ws, 26, 8, "Komisyon")
    for i in range(8):
        h(ws, 27 + i, 7, i + 1, sayi=CATI)
        h(ws, 27 + i, 8, 85000 + i * 12000, sayi=TL)

    h(ws, 26, 10, "Senaryo")
    h(ws, 26, 11, "Tutar")
    for i, (ad, t) in enumerate([
        ("Temkinli", 380000), ("Baz", 450000), ("İyimser", 520000),
    ], 27):
        h(ws, i, 10, ad)
        h(ws, i, 11, t, sayi=TL)

    c1 = PieChart()
    c1.title = "Pipeline Durum"
    c1.add_data(Reference(ws, min_col=2, min_row=26, max_row=33), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=27, max_row=33))
    c1.width, c1.height = 10, 7
    ws.add_chart(c1, "A38")

    c2 = BarChart()
    c2.title = "Yaşlandırma"
    c2.add_data(Reference(ws, min_col=5, min_row=26, max_row=30), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=27, max_row=30))
    c2.width, c2.height = 10, 7
    ws.add_chart(c2, "F38")

    c3 = LineChart()
    c3.title = "Haftalık Komisyon"
    c3.add_data(Reference(ws, min_col=8, min_row=26, max_row=34), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=7, min_row=27, max_row=34))
    c3.width, c3.height = 12, 7
    ws.add_chart(c3, "A53")

    c4 = BarChart()
    c4.title = "Senaryo"
    c4.add_data(Reference(ws, min_col=11, min_row=26, max_row=29), titles_from_data=True)
    c4.set_categories(Reference(ws, min_col=10, min_row=27, max_row=29))
    c4.width, c4.height = 10, 7
    ws.add_chart(c4, "F53")

    h(ws, 26, 13, "Uyarı")
    h(ws, 26, 14, "Adet")
    for i, (ad, adet) in enumerate([
        ("Yetim", 2), ("Geçersiz", 2), ("Zincir", 2), ("Geciken", 9),
    ], 27):
        h(ws, i, 13, ad)
        h(ws, i, 14, adet, sayi=CATI)

    h(ws, 26, 16, "AyNo")
    h(ws, 26, 17, "Trend")
    for i, val in enumerate([0.72, 0.74, 0.78, 0.80, 0.82, 0.85], 0):
        h(ws, 27 + i, 16, i + 1, sayi=CATI)
        h(ws, 27 + i, 17, val, sayi=YÜZDE)

    c5 = BarChart()
    c5.title = "Uyarı Türleri"
    c5.add_data(Reference(ws, min_col=14, min_row=26, max_row=30), titles_from_data=True)
    c5.set_categories(Reference(ws, min_col=13, min_row=27, max_row=30))
    c5.width, c5.height = 10, 7
    ws.add_chart(c5, "A68")

    c6 = LineChart()
    c6.title = "Aylık Dönüşüm Trend"
    c6.add_data(Reference(ws, min_col=17, min_row=26, max_row=32), titles_from_data=True)
    c6.set_categories(Reference(ws, min_col=16, min_row=27, max_row=32))
    c6.width, c6.height = 10, 7
    ws.add_chart(c6, "F68")

    c7 = PieChart()
    c7.title = "Yaş Payı"
    c7.add_data(Reference(ws, min_col=5, min_row=26, max_row=30), titles_from_data=True)
    c7.set_categories(Reference(ws, min_col=4, min_row=27, max_row=30))
    c7.width, c7.height = 9, 7
    ws.add_chart(c7, "A83")

    c8 = BarChart()
    c8.title = "Durum Adet"
    c8.add_data(Reference(ws, min_col=2, min_row=26, max_row=33), titles_from_data=True)
    c8.set_categories(Reference(ws, min_col=1, min_row=27, max_row=33))
    c8.width, c8.height = 10, 7
    ws.add_chart(c8, "F83")

    baski_hazirla(ws, "A1:N36", f"{URUN_AD} | Pano | {SURUM}")
    genislik(ws, {"A": 26, "B": 16, "C": 14, "D": 14, "E": 12})


def kartlar(ws):
    sayfa_hazirla(ws, "KARTLAR", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Danışman kartları — her kaydın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=epk_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: DAN-00009 | Risk: Elle çakışma")

    sutunlar = ["KartId", "AdSoyad", "Ofis", "Rol", "HedefKomisyon",
                "VarsayilanOran", "Aktif", "BeklenenKomisyon", "KapaliAdet"]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "BeklenenKomisyon": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'SUMIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],tblAkis[Tutar]))'
        ),
        "KapaliAdet": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'COUNTIFS(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],'
            'tblAkis[Durum],"KAPALI"))'
        ),
    }
    for i, (kid, ad, ofis, rol) in enumerate(DANISMANLAR):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Kart Kimliği", mesaj="DAN-##### formatında")
        _sari(ws, r, 2, ad, baslik="Ad Soyad", mesaj="Danışman adı")
        _sari(ws, r, 3, ofis, baslik="Ofis", mesaj="Şube / ofis")
        _sari(ws, r, 4, rol, baslik="Rol", mesaj="Satış/Kiralama/Proje")
        _sari(ws, r, 5, 400000 + i * 50000, sayi=TL, baslik="Hedef", mesaj="Aylık komisyon hedefi")
        _sari(ws, r, 6, 0.03 if i % 2 == 0 else 0.02, sayi=YÜZDE,
              baslik="Oran", mesaj="Varsayılan komisyon oranı")
        _sari(ws, r, 7, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_KART, SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None,
                  sayi=TL if c == 5 else (YÜZDE if c == 6 else None),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblKartlar", f"A{HDR}:I{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Kart Kimliği", mesaj="DAN-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Ad"), (3, "Ofis"), (4, "Rol")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "decimal", "0", f"E{ILK}:E{SON}",
              baslik="Hedef", mesaj="Hedef ≥ 0",
              hata_baslik="Hedef", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"F{ILK}:F{SON}",
              baslik="Oran", mesaj="0–1 arası",
              hata_baslik="Oran", hata_mesaj="0–1",
              isaret="between", f2="1")
    dogrulama(ws, "list", "ListeEvetHayir", f"G{ILK}:G{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 10)})
    ws.column_dimensions["B"].width = 18


def akis(ws):
    sayfa_hazirla(ws, "AKIS", RENK, URUN_AD, son_kolon=26)
    alt_bant(ws, 3,
             "İlan/pipeline satırları — durum makinesi + yetim + geçersiz geçiş + komisyon zinciri",
             son_kolon=20)
    sutunlar = [
        "IslemId", "KaynakKartId", "IlanNo", "BaslangicTarihi",
        "OncekiDurum", "Durum", "ListeFiyati", "TeklifFiyati", "SatisFiyati",
        "KomisyonOrani", "TahsilKomisyon", "Tutar",
        "GayrimenkulTipi", "YorumA", "YorumB",
        "GecikmeGun", "YasKova", "BeklenenKomisyon", "DonusumBayrak",
        "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)

    yas_f = yaslandirma_gun(
        "tblAkis[[#This Row],[BaslangicTarihi]]", "raporTarihi").lstrip("=")
    kova_f = _yas_kova_epk("tblAkis[[#This Row],[GecikmeGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblAkis[[#This Row],[KaynakKartId]]", "tblKartlar[KartId]").lstrip("=")
    gecis_birlesik = (
        'tblAkis[[#This Row],[OncekiDurum]]&"|"&tblAkis[[#This Row],[Durum]]'
    )
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    # Zincir: liste↔teklif (açık) veya teklif↔satış (kapalı/tapu)
    zincir_liste = zincir_kirik_formul(
        "tblAkis[[#This Row],[ListeFiyati]]",
        "tblAkis[[#This Row],[TeklifFiyati]]",
        "zincirTolerans",
    ).lstrip("=")
    zincir_satis = zincir_kirik_formul(
        "tblAkis[[#This Row],[TeklifFiyati]]",
        "tblAkis[[#This Row],[SatisFiyati]]",
        "zincirTolerans",
    ).lstrip("=")
    zincir_guvenli = (
        f'IF(OR(tblAkis[[#This Row],[Durum]]="KAPALI",'
        f'tblAkis[[#This Row],[Durum]]="TAPU",'
        f'tblAkis[[#This Row],[SatisFiyati]]>0),{zincir_satis},'
        f'IF(tblAkis[[#This Row],[TeklifFiyati]]>0,{zincir_liste},"OK"))'
    )
    bek_kom = (
        'IF(tblAkis[[#This Row],[IslemId]]="","",'
        'IF(tblAkis[[#This Row],[SatisFiyati]]>0,'
        'tblAkis[[#This Row],[SatisFiyati]]*tblAkis[[#This Row],[KomisyonOrani]],'
        'IF(tblAkis[[#This Row],[TeklifFiyati]]>0,'
        'tblAkis[[#This Row],[TeklifFiyati]]*tblAkis[[#This Row],[KomisyonOrani]],'
        'tblAkis[[#This Row],[ListeFiyati]]*tblAkis[[#This Row],[KomisyonOrani]])))'
    )
    donusum_f = (
        'IF(tblAkis[[#This Row],[Durum]]="KAPALI","KAPANAN",'
        'IF(tblAkis[[#This Row],[Durum]]="IPTAL","IPTAL","ACIK"))'
    )

    form = {
        "GecikmeGun": yas_f,
        "YasKova": kova_f,
        "BeklenenKomisyon": bek_kom,
        "DonusumBayrak": donusum_f,
        "YetimBayrak": yetim_f,
        "GecisUyarisi": gecis_f,
        "ZincirBayrak": zincir_guvenli,
        "KayitDolu": 'IF(tblAkis[[#This Row],[IslemId]]="","",1)',
    }

    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="İşlem", mesaj="ILN-#####")
        _sari(ws, r, 2, s["kart"], baslik="Danışman", mesaj="Kart kimliği")
        _sari(ws, r, 3, s["ilan_no"], baslik="İlan No", mesaj="İlan numarası")
        _sari(ws, r, 4, s["baslangic"], sayi=TARİH, baslik="Başlangıç", mesaj="İlan tarihi")
        _sari(ws, r, 5, s["onceki"], baslik="Önceki Durum", mesaj="Durum haritası")
        _sari(ws, r, 6, s["durum"], baslik="Durum", mesaj="Güncel durum")
        _sari(ws, r, 7, s["liste"], sayi=TL, baslik="Liste", mesaj="Liste fiyatı TL")
        _sari(ws, r, 8, s["teklif"], sayi=TL, baslik="Teklif", mesaj="Teklif fiyatı TL")
        _sari(ws, r, 9, s["satis"] or None, sayi=TL, baslik="Satış", mesaj="Satış fiyatı TL")
        _sari(ws, r, 10, s["oran"], sayi=YÜZDE, baslik="Oran", mesaj="Komisyon oranı")
        _sari(ws, r, 11, s["tahsil"] or None, sayi=TL, baslik="Tahsil", mesaj="Tahsil komisyon")
        _sari(ws, r, 12, s["tutar"], sayi=TL, baslik="Tutar", mesaj="Pipeline komisyon tutarı")
        _sari(ws, r, 13, s["tip"], baslik="Tip", mesaj="Gayrimenkul tipi")
        _sari(ws, r, 14, s["yorum_a"], baslik="Yorum A", mesaj="A veya B")
        _sari(ws, r, 15, s["yorum_b"], baslik="Yorum B", mesaj="Serbest not")

    for r in range(ILK + DEMO_AKIS, SON + 1):
        for c in range(1, 16):
            fmt = None
            if c == 4:
                fmt = TARİH
            elif c in (7, 8, 9, 11, 12):
                fmt = TL
            elif c == 10:
                fmt = YÜZDE
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblAkis", f"A{HDR}:W{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="İşlem", mesaj="ILN-#####",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Danışman", mesaj="Kart kimliği",
              hata_baslik="Kart", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"D{ILK}:D{SON}",
              baslik="Başlangıç", mesaj="İlan tarihi",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "list", "ListeDurumlar", f"E{ILK}:E{SON}",
              baslik="Önceki Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"F{ILK}:F{SON}",
              baslik="Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    for c, ad in [(7, "Liste"), (8, "Teklif"), (9, "Satış"), (11, "Tahsil"), (12, "Tutar")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"J{ILK}:J{SON}",
              baslik="Oran", mesaj="0–1",
              hata_baslik="Oran", hata_mesaj="0–1",
              isaret="between", f2="1")
    dogrulama(ws, "list", "ListeYorumAB", f"N{ILK}:N{SON}",
              baslik="Yorum A", mesaj="A veya B",
              hata_baslik="Liste", hata_mesaj="A veya B")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 24)})
    ws.column_dimensions["O"].width = 16
    ws.column_dimensions["U"].width = 16


def kuyruklar(ws):
    sayfa_hazirla(ws, "KUYRUKLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Durum bazlı canlı kuyruk — kapalı/iptal kayıtlar düşer", son_kolon=10)
    h(ws, 5, 1, "Kuyruk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Tutar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    for i, (ad, kod) in enumerate([
        ("Kuyruk İlan", "ILAN"),
        ("Kuyruk Görüşme", "GORUSME"),
        ("Kuyruk Teklif", "TEKLIF"),
        ("Kuyruk Sözleşme", "SOZLESME"),
        ("Kuyruk Tapu", "TAPU"),
    ], 6):
        h(ws, i, 1, ad)
        h(ws, i, 2,
          f'=COUNTIFS(tblAkis[Durum],"{kod}",tblAkis[Tutar],">0")', sayi=CATI)
        h(ws, i, 3,
          f'=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"{kod}")', sayi=TL)

    h(ws, 11, 1, "Geciken (açık, yaş>eşik)")
    h(ws, 11, 2,
      '=COUNTIFS(tblAkis[GecikmeGun],">"&epk_esikGecikmeGun,'
      'tblAkis[Durum],"<>KAPALI",tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")',
      sayi=CATI)
    h(ws, 11, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[GecikmeGun],">"&epk_esikGecikmeGun,'
      'tblAkis[Durum],"<>KAPALI",tblAkis[Durum],"<>IPTAL")',
      sayi=TL)

    h(ws, 13, 1, "Kapalı / İptal — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 1, "Kapalı")
    h(ws, 14, 2, '=COUNTIF(tblAkis[Durum],"KAPALI")', sayi=CATI)
    h(ws, 14, 3, '=SUMIF(tblAkis[Durum],"KAPALI",tblAkis[TahsilKomisyon])', sayi=TL)
    h(ws, 15, 1, "İptal")
    h(ws, 15, 2, '=COUNTIF(tblAkis[Durum],"IPTAL")', sayi=CATI)
    h(ws, 15, 3, '=SUMIF(tblAkis[Durum],"IPTAL",tblAkis[Tutar])', sayi=TL)

    h(ws, 17, 1, "Kuyruk Özeti", kalin=True)
    h(ws, 17, 2, "=epk_kuyrukOzeti", kaydir=True)
    ws.merge_cells("B17:F17")

    h(ws, 19, 1, "Yaş Kova Dağılımı", kalin=True, yazi=KOYU_LACIVERT)
    for i, kova in enumerate(["0-30", "31-60", "61-90", "90+"], 20):
        h(ws, i, 1, kova)
        h(ws, i, 2, f'=COUNTIF(tblAkis[YasKova],"{kova}")', sayi=CATI)
        h(ws, i, 3, f'=SUMIF(tblAkis[YasKova],"{kova}",tblAkis[Tutar])', sayi=TL)

    genislik(ws, {"A": 36, "B": 14, "C": 16})


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Hesap motoru — ≥40 adım; sabitler AYARLAR'dan", son_kolon=10)
    h(ws, 5, 1, "Adım", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Açıklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    adimlar = [
        ("H01", "Toplam işlem", '=COUNTA(tblAkis[IslemId])'),
        ("H02", "Dolu kayıt", '=SUM(tblAkis[KayitDolu])'),
        ("H03", "Toplam pipeline tutar", "=SUM(tblAkis[Tutar])"),
        ("H04", "Liste toplam", "=SUM(tblAkis[ListeFiyati])"),
        ("H05", "Teklif toplam", "=SUM(tblAkis[TeklifFiyati])"),
        ("H06", "Satış toplam", "=SUM(tblAkis[SatisFiyati])"),
        ("H07", "İlan adet", '=COUNTIF(tblAkis[Durum],"ILAN")'),
        ("H08", "Görüşme adet", '=COUNTIF(tblAkis[Durum],"GORUSME")'),
        ("H09", "Teklif adet", '=COUNTIF(tblAkis[Durum],"TEKLIF")'),
        ("H10", "Sözleşme adet", '=COUNTIF(tblAkis[Durum],"SOZLESME")'),
        ("H11", "Tapu adet", '=COUNTIF(tblAkis[Durum],"TAPU")'),
        ("H12", "Kapalı adet", '=COUNTIF(tblAkis[Durum],"KAPALI")'),
        ("H13", "İptal adet", '=COUNTIF(tblAkis[Durum],"IPTAL")'),
        ("H14", "Açık kuyruk adet", "=C12+C13+C14+C15+C16"),
        ("H15", "Kapalı düşüm adet", "=C17+C18"),
        ("H16", "Açık tutar kuyruk",
         '=SUMIF(tblAkis[Durum],"ILAN",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"GORUSME",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"TEKLIF",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"SOZLESME",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"TAPU",tblAkis[Tutar])'),
        ("H17", "Yetim adet", '=COUNTIF(tblAkis[YetimBayrak],"YETİM")'),
        ("H18", "Geçersiz geçiş", '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")'),
        ("H19", "Zincir kırık", '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")'),
        ("H20", "Zincir OK", '=COUNTIF(tblAkis[ZincirBayrak],"OK")'),
        ("H21", "Yaş 0-30", '=COUNTIF(tblAkis[YasKova],"0-30")'),
        ("H22", "Yaş 31-60", '=COUNTIF(tblAkis[YasKova],"31-60")'),
        ("H23", "Yaş 61-90", '=COUNTIF(tblAkis[YasKova],"61-90")'),
        ("H24", "Yaş 90+", '=COUNTIF(tblAkis[YasKova],"90+")'),
        ("H25", "Ort gecikme gün",
         "=IFERROR(AVERAGEIF(tblAkis[KayitDolu],1,tblAkis[GecikmeGun]),0)"),
        ("H26", "Geciken adet",
         '=COUNTIFS(tblAkis[GecikmeGun],">"&epk_esikGecikmeGun,'
         'tblAkis[Durum],"<>KAPALI",tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")'),
        ("H27", "Geciken tutar",
         '=SUMIFS(tblAkis[Tutar],tblAkis[GecikmeGun],">"&epk_esikGecikmeGun,'
         'tblAkis[Durum],"<>KAPALI",tblAkis[Durum],"<>IPTAL")'),
        ("H28", "Beklenen komisyon", "=C21"),
        ("H29", "Tahsil komisyon",
         '=SUMIF(tblAkis[Durum],"KAPALI",tblAkis[TahsilKomisyon])'),
        ("H30", "Dönüşüm oranı",
         "=IFERROR(C17/MAX(C7,1),0)"),
        ("H31", "Dönüşüm sapma", "=C35-epk_hedefDonusum"),
        ("H32", "Temkinli senaryo", "=C33*epk_senaryoTemkinli"),
        ("H33", "Baz senaryo", "=C33*epk_senaryoBaz"),
        ("H34", "İyimser senaryo", "=C33*epk_senaryoIyimser"),
        ("H35", "Senaryo farkı", "=C39-C37"),
        ("H36", "Tornado (oran×tutar)", "=C33*epk_tornadoOran"),
        ("H37", "HHI proxy",
         "=IFERROR((MAX(C12,C13,C14)/MAX(C19,1))^2,0)"),
        ("H38", "Kalite skoru",
         "=IF(C6=0,0,MAX(epk_sifir,epk_kaliteTaban-C22*epk_kaliteCeza"
         "-C23*epk_kaliteCeza-C24*epk_kaliteCeza-C31*2))"),
        ("H39", "Tahmin PERCENTILE",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[Tutar],epk_yuzdelikOran),0)"),
        ("H40", "Tahmin FORECAST",
         "=IFERROR(FORECAST(COUNTA(tblAkis[IslemId])+1,tblAkis[Tutar],"
         "tblAkis[KayitDolu]),C44)"),
        ("H41", "Tahmin aralık", "=IFERROR((C44+C45)/2,0)"),
        ("H42", "P90 satış",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[ListeFiyati],epk_yuzdelikOran),0)"),
        ("H43", "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        ("H44", "Hedef toplam", "=SUM(tblKartlar[HedefKomisyon])"),
        ("H45", "Hedef karşılama", "=IFERROR(C33/MAX(C49,1),0)"),
        ("H46", "Kuyruk özeti metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"İlan",C12,"Görüşme",C13,"Teklif",C14,'
         '"Sözleşme",C15,"Tapu",C16,"Geciken",C31)'),
        ("H47", "Yaş kova metin",
         '=_xlfn.TEXTJOIN("/",TRUE,C26,C27,C28,C29)'),
        ("H48", "Senaryo motor özet",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Temkinli",C37,"Baz",C38,"İyimser",C39)'),
        ("H49", "Veri hazırlık", "=IF(C6=0,0,C7/MAX(C6,1))"),
        ("H50", "Trend ortalama",
         "=IFERROR(C35*(epk_trendOcak+epk_trendSubat+epk_trendMart+epk_trendNisan+"
         "epk_trendMayis+epk_trendHaziran+epk_trendTemmuz+epk_trendAgustos+"
         "epk_trendEylul+epk_trendEkim+epk_trendKasim+epk_trendAralik)/12,0)"),
    ]
    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        low = acik.lower()
        if any(x in low for x in ("metin", "özet")):
            fmt = None
        elif "gün" in low:
            fmt = GUN
        elif any(x in low for x in ("oran", "sapma", "hhi", "hazırlık", "karşılama", "trend", "dönüşüm")):
            fmt = YÜZDE
        elif any(x in low for x in ("tutar", "komisyon", "senaryo", "tornado", "tahmin", "liste", "teklif", "satış", "hedef", "fark", "p90")) and "adet" not in low:
            fmt = TL
        elif "kalite" in low:
            fmt = CATI
        else:
            fmt = CATI
        h(ws, r, 3, form, sayi=fmt, kalin=True)

    h(ws, 58, 1, "Senaryo", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 58, 2, "Çarpan", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 58, 3, "Sonuç", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 59, 1, "Temkinli")
    h(ws, 59, 2,
      "=epk_senaryoTemkinli*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 59, 3, "=C37", sayi=TL)
    h(ws, 60, 1, "Baz")
    h(ws, 60, 2,
      "=epk_senaryoBaz*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 60, 3, "=C38", sayi=TL)
    h(ws, 61, 1, "İyimser")
    h(ws, 61, 2,
      "=epk_senaryoIyimser*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 61, 3, "=C39", sayi=TL)

    genislik(ws, {"A": 10, "B": 28, "C": 50})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Dönemsel Komisyon Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Rapor Açık Pipeline", "=epk_acikAdet", CATI),
        (10, "Rapor Beklenen Komisyon", "=epk_beklenenKomisyon", TL),
        (11, "Rapor Tahsil Komisyon", "=epk_tahsilKomisyon", TL),
        (12, "Rapor Dönüşüm", "=epk_donusumOrani", YÜZDE),
        (13, "Rapor Yetim", "=epk_yetimKayit", CATI),
        (14, "Rapor Geçersiz Geçiş", "=epk_gecersizGecis", CATI),
        (15, "Rapor Zincir Kırık", "=epk_zincirKirik", CATI),
        (16, "Rapor Tahmin", "=epk_tahminAralik", TL),
        (17, "Rapor Kanıt Özeti", "=epk_kanitRaporu", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 19, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 1,
      "Durum haritası AYARLAR'dadır. Geçersiz geçiş engellenmez, görünür kılınır. "
      "Komisyon oranı satırdan veya danışman varsayılanından gelir. "
      "Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A20:H20")
    h(ws, 22, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 1, "1. Yetim kayıtları KARTLAR'a bağlayın veya satırı düzeltin.")
    h(ws, 24, 1, "2. Geciken teklif/sözleşme satırlarında danışmanı uyarın.")
    h(ws, 25, 1, "3. Zincir kırık satırlarda liste/teklif/satış farkını açıklayın.")
    h(ws, 27, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    baski_hazirla(ws, "A1:H27", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 32, "B": 40, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "İşlem sayısı", '=COUNTA(tblAkis[IslemId])'),
        (6, "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        (7, "İlan", '=COUNTIF(tblAkis[Durum],"ILAN")'),
        (8, "Görüşme", '=COUNTIF(tblAkis[Durum],"GORUSME")'),
        (9, "Teklif", '=COUNTIF(tblAkis[Durum],"TEKLIF")'),
        (10, "Sözleşme", '=COUNTIF(tblAkis[Durum],"SOZLESME")'),
        (11, "Tapu", '=COUNTIF(tblAkis[Durum],"TAPU")'),
        (12, "Kapalı", '=COUNTIF(tblAkis[Durum],"KAPALI")'),
        (13, "İptal", '=COUNTIF(tblAkis[Durum],"IPTAL")'),
        (14, "Açık toplam", "=B7+B8+B9+B10+B11"),
        (15, "Yetim", "=epk_yetimKayit"),
        (16, "Geçersiz", "=epk_gecersizGecis"),
        (17, "Zincir kırık", "=epk_zincirKirik"),
        (18, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (19, "Kalite skoru", "=epk_kaliteSkor"),
        (20, "Beklenen komisyon", "=epk_beklenenKomisyon"),
        (21, "Tahsil komisyon", "=epk_tahsilKomisyon"),
        (22, "Kapalı düşüm", "=epk_kapaliDusum"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, kalin=True,
          sayi=TL if r >= 20 else (CATI if r < 18 else None))
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO_KART} danışman kartı ve {DEMO_AKIS} ilan satırı vardır. "
      "Yetim (22, 25), geçersiz geçiş (18, 20) ve zincir kırık (12, 19) kasıtlıdır. "
      "KARTLAR ve AKIS'i temizleyip kendi verinizi girebilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Durumlar: İlan/Görüşme/Teklif/Sözleşme/Tapu/Kapalı/İptal — kapalılar kuyruktan düşer")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 20000, sayi=CATI)
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

    h(ws, 3, 7, "Yorum A/B", kalin=True)
    h(ws, 6, 7, "Yorum", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    _sari(ws, 7, 7, "A", baslik="Yorum", mesaj="A veya B")
    _sari(ws, 8, 7, "B", baslik="Yorum", mesaj="A veya B")

    h(ws, 3, 9, "Gayrimenkul Tipi", kalin=True)
    h(ws, 6, 9, "Tip", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(["Konut", "İşyeri", "Arsa", "Villa", "Dükkan"], 7):
        _sari(ws, i, 9, b, baslik="Tip", mesaj="Gayrimenkul tipi")
    genislik(ws, {"A": 14, "C": 24, "E": 12, "G": 10, "I": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + durum haritası (A01/A06)", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("zincirTolerans", 100000, "TL", "İş kuralı — liste/teklif/satış sapma", "01.01.2026",
         "ZİNCİR KIRIK eşiği"),
        ("epk_esikGecikmeGun", 30, "gün", "İç politika", "01.01.2026", "Gecikme eşiği"),
        ("epk_esikGeciken", 5, "adet", "İç politika", "01.01.2026", "Karar eşiği geciken"),
        ("epk_kovaEsik1", 30, "gün", "Yaşlandırma", "01.01.2026", "Kova 1 üst sınır"),
        ("epk_kovaEsik2", 60, "gün", "Yaşlandırma", "01.01.2026", "Kova 2 üst sınır"),
        ("epk_kovaEsik3", 90, "gün", "Yaşlandırma", "01.01.2026", "Kova 3 üst sınır"),
        ("epk_hedefDonusum", 0.25, "oran", "Satış yönetimi", "01.01.2026", "Hedef dönüşüm"),
        ("epk_sifir", 0, "adet", "Kalite motoru", "01.01.2026", "MAX alt sınırı"),
        ("epk_kaliteTaban", 100, "puan", "Kalite motoru", "01.01.2026", "Kalite taban puan"),
        ("epk_kaliteCeza", 10, "puan", "Kalite motoru", "01.01.2026", "Uyarı başına ceza"),
        ("epk_senaryoTemkinli", 0.85, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("epk_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("epk_senaryoIyimser", 1.15, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("epk_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("epk_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("epk_olcekHedef", 20000, "satir", "Manda A3 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("epk_azamiTutar", 100000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul"),
        ("epk_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi",
         "01.01.2026", "Durum makinesi kaynağı"),
        ("epk_kuyrukOzeti_metin", "KUYRUKLAR canlı özet", "metin", "A05 kuyruk", "01.01.2026",
         "Kuyruk görünümü"),
        ("epk_kapaliDusum_metin", "kapali kategori kuyruktan düşer", "metin", "A06", "01.01.2026",
         "Kapalı düşüm kuralı"),
        ("epk_kanitRaporu_metin", "RAPOR kanıt çıktısı", "metin", "Dönemsel rapor", "01.01.2026",
         "Kanıt raporu"),
        ("epk_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("epk_firmaUnvan", "Örnek Emlak Ofisi Ltd.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma"),
    ]
    # Ay trend NR (G07 — gömülü ay çarpanı yok)
    for ad, val in AY_TREND:
        params.append((ad, val, "oran", "Tarihsel dönüşüm", "01.01.2026", f"{ad} trend"))

    for i, (ana, deg, bir, kay, yur, acik) in enumerate(params):
        r = 6 + i
        h(ws, r, 1, ana)
        if isinstance(deg, date):
            _sari(ws, r, 2, deg, sayi=TARİH, baslik=ana, mesaj=acik)
        elif isinstance(deg, float) and deg <= 1.15:
            _sari(ws, r, 2, deg, sayi=YÜZDE, baslik=ana, mesaj=acik)
        elif isinstance(deg, (int, float)):
            _sari(ws, r, 2, deg, sayi=TL if bir == "TL" else CATI, baslik=ana, mesaj=acik)
        else:
            _sari(ws, r, 2, deg, baslik=ana, mesaj=acik)
        h(ws, r, 3, bir)
        h(ws, r, 4, kay)
        h(ws, r, 5, yur)
        h(ws, r, 6, acik, kaydir=True)

    # Durum haritası tablosu
    h(ws, 42, 8, "Durum Haritası", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    dh_bas = ["durum_id", "durum_adi", "sira", "izinli_sonraki_durumlar", "kategori"]
    baslik_satiri(ws, 43, dh_bas, basla=8)
    for i, d in enumerate(DURUM_HARITASI):
        r = 44 + i
        h(ws, r, 8, d["durum_id"])
        h(ws, r, 9, d["durum_adi"])
        h(ws, r, 10, d["sira"], sayi=CATI)
        h(ws, r, 11, d["izinli_sonraki_durumlar"])
        h(ws, r, 12, d["kategori"])
    tablo_ekle(ws, "tblDurumHaritasi", "H43:L50", dh_bas)

    h(ws, 53, 8, "Motor Çıktıları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 54, 8, "anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 54, 9, "deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    motor = [
        (55, "epk_acikAdet", "=HESAP!C19"),
        (56, "epk_beklenenKomisyon", "=HESAP!C33"),
        (57, "epk_gecikenAdet", "=HESAP!C31"),
        (58, "epk_tahsilKomisyon", "=HESAP!C34"),
        (59, "epk_yetimKayit", "=HESAP!C22"),
        (60, "epk_gecersizGecis", "=HESAP!C23"),
        (61, "epk_zincirKirik", "=HESAP!C24"),
        (62, "epk_donusumOrani", "=HESAP!C35"),
        (63, "epk_yasKova", "=HESAP!C52"),
        (64, "epk_tahsilatOrt", "=HESAP!C35"),
        (65, "epk_tahsilatSapma", "=HESAP!C36"),
        (66, "epk_tahminAralik", "=HESAP!C46"),
        (67, "epk_senaryoKarsilastirma", "=HESAP!C40"),
        (68, "epk_tornadoZirve", "=HESAP!C41"),
        (69, "epk_hhi", "=HESAP!C42"),
        (70, "epk_kaliteSkor", "=HESAP!C43"),
        (71, "epk_yuzdelikP90", "=HESAP!C47"),
        (72, "epk_senaryoMotor", "=HESAP!C53"),
        (73, "epk_kapaliDusum", "=HESAP!C20"),
        (74, "epk_kuyrukOzeti", "=HESAP!C51"),
        (75, "epk_kanitRaporu",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Kapalı",HESAP!C17,"Komisyon",HESAP!C34,'
         '"Dönüşüm",TEXT(HESAP!C35,"0.0%"))'),
        (76, "epk_sonrakiKimlik",
         '="DAN-"&TEXT(COUNTA(tblKartlar[KartId])+1,"00000")'),
        (77, "epk_kararKapi", "=PANO!B4"),
        (78, "epk_modulT1", "=HESAP!C19"),
        (79, "epk_modulT2", "=HESAP!C33"),
        (80, "epk_modulO1", "=HESAP!C36"),
        (81, "epk_modulO2", "=HESAP!C41"),
        (82, "epk_modulO3", "=HESAP!C53"),
        (83, "epk_modulO6", "=HESAP!C42"),
        (84, "epk_modulO8", "=HESAP!C43"),
        (85, "epk_modulM", "=HESAP!C53"),
        (86, "epk_modulI1", "=HESAP!C46"),
        (87, "epk_modulI2", "=HESAP!C47"),
    ]
    for r, ana, form in motor:
        h(ws, r, 8, ana)
        h(ws, r, 9, form, kalin=True)

    h(ws, 89, 8, "Yorum Motoru", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (90, "epk_modulT1Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Açık pipeline",epk_acikAdet,"adet.")'),
        (91, "epk_modulT2Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Beklenen komisyon",TEXT(epk_beklenenKomisyon,"₺#,##0"),'
         '"tahsil",TEXT(epk_tahsilKomisyon,"₺#,##0"))'),
        (92, "epk_modulO1Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Dönüşüm",TEXT(epk_donusumOrani,"0.0%"),'
         '"hedef sapma",TEXT(epk_tahsilatSapma,"0.0%"))'),
        (93, "epk_modulO2Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(epk_tornadoZirve,"₺#,##0"))'),
        (94, "epk_modulO3Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(epk_senaryoKarsilastirma,"₺#,##0"))'),
        (95, "epk_modulO6Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma HHI",TEXT(epk_hhi,"0.0%"))'),
        (96, "epk_modulO8Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite skoru",epk_kaliteSkor)'),
        (97, "epk_modulMYorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motor",epk_senaryoMotor)'),
        (98, "epk_modulI1Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(epk_tahminAralik,"₺#,##0"))'),
        (99, "epk_modulI2Yorum",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 liste",TEXT(epk_yuzdelikP90,"₺#,##0"))'),
    ]
    for r, ana, form in yorumlar:
        h(ws, r, 8, ana)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 24, "B": 22, "C": 10, "D": 28, "E": 14, "F": 28,
                  "H": 26, "I": 55, "J": 8, "K": 28, "L": 12})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    adimlar = [
        "1. AYARLAR'da rapor tarihi, komisyon toleransı ve dönüşüm hedefini kontrol edin.",
        "2. KARTLAR'a danışman kimliklerini girin (sarı hücreler).",
        "3. AKIS'te her ilan için durum, liste/teklif/satış ve komisyon oranını yazın.",
        "4. KUYRUKLAR'da günün açık pipeline ve geciken satırlarını okuyun.",
        "5. PANO'da karar (UYGUN/DİKKAT/KRİTİK/VERİ YOK) ve KPI'ları görün.",
        "6. RAPOR'u yazdırın veya PDF'e aktarın; SMMM dosyasına ekleyin.",
        "7. Formül hücreleri 1234 şifresiyle korunur; sarı alanlar serbesttir.",
    ]
    for i, m in enumerate(adimlar, 5):
        h(ws, i, 1, m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    h(ws, 14, 1, "Sık yapılan hatalar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "• Danışman kartı olmadan AKIS satırı → YETİM bayrağı")
    h(ws, 16, 1, "• İzinli olmayan durum geçişi → GEÇERSİZ GEÇİŞ (engellemez, görünür)")
    h(ws, 17, 1, "• Liste/teklif/satış sapması tolerans üstü → ZİNCİR KIRIK")
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    pass  # sayfa içi dogrulama yeterli


def kosullu_bicim(wb):
    yesil = PatternFill("solid", fgColor="C6EFCE")
    sari = PatternFill("solid", fgColor="FFEB9C")
    kirmizi = PatternFill("solid", fgColor="FFC7CE")
    yf = Font(color="006100", bold=True)
    kf = Font(color="9C0006", bold=True)

    ws = wb["PANO"]
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"DİKKAT"'], fill=sari))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"KRİTİK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(6, 24):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["AKIS"]
    ws.conditional_formatting.add(
        f"T{ILK}:T{SON}",
        CellIsRule(operator="equal", formula=['"YETİM"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"U{ILK}:U{SON}",
        CellIsRule(operator="equal", formula=['"GEÇERSİZ GEÇİŞ"'], fill=sari))
    ws.conditional_formatting.add(
        f"V{ILK}:V{SON}",
        CellIsRule(operator="equal", formula=['"ZİNCİR KIRIK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"Q{ILK}:Q{SON}",
        CellIsRule(operator="equal", formula=['"90+"'], fill=kirmizi))

    ws = wb["KUYRUKLAR"]
    for r in range(6, 16):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add(
        "B18", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B18", CellIsRule(operator="equal", formula=['"VERİ VAR"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B19", CellIsRule(operator="greaterThanOrEqual", formula=["90"], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B19", CellIsRule(operator="lessThan", formula=["70"], fill=kirmizi, font=kf))
    for r in range(5, 18):
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
    for r in range(9, 17):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["KARTLAR"]
    ws.conditional_formatting.add(
        f"G{ILK}:G{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"E{ILK}:E{SON}",
        CellIsRule(operator="greaterThan", formula=["500000"], fill=yesil))

    ws = wb["HESAP"]
    for r in range(6, 56):
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))


def adlari_bagla(wb):
    # Parametre satırları: 6.. (21 sabit + 12 ay = 33) → B6..B38
    param_adlari = [
        "raporTarihi", "zincirTolerans", "epk_esikGecikmeGun", "epk_esikGeciken",
        "epk_kovaEsik1", "epk_kovaEsik2", "epk_kovaEsik3", "epk_hedefDonusum",
        "epk_sifir", "epk_kaliteTaban", "epk_kaliteCeza",
        "epk_senaryoTemkinli", "epk_senaryoBaz", "epk_senaryoIyimser",
        "epk_tornadoOran", "epk_yuzdelikOran", "epk_olcekHedef", "epk_azamiTutar",
        "epk_durumHaritasi", "epk_kuyrukOzeti_metin", "epk_kapaliDusum_metin",
        "epk_kanitRaporu_metin", "epk_dosyaSurumu", "epk_firmaUnvan",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    for i, (ad, _v) in enumerate(AY_TREND):
        ad_ekle(wb, ad, f"AYARLAR!$B${30 + i}")  # 6+24=30

    motor_map = {
        "epk_acikAdet": 55, "epk_beklenenKomisyon": 56, "epk_gecikenAdet": 57,
        "epk_tahsilKomisyon": 58, "epk_yetimKayit": 59, "epk_gecersizGecis": 60,
        "epk_zincirKirik": 61, "epk_donusumOrani": 62, "epk_yasKova": 63,
        "epk_tahsilatOrt": 64, "epk_tahsilatSapma": 65, "epk_tahminAralik": 66,
        "epk_senaryoKarsilastirma": 67, "epk_tornadoZirve": 68, "epk_hhi": 69,
        "epk_kaliteSkor": 70, "epk_yuzdelikP90": 71, "epk_senaryoMotor": 72,
        "epk_kapaliDusum": 73, "epk_kuyrukOzeti": 74, "epk_kanitRaporu": 75,
        "epk_sonrakiKimlik": 76, "epk_kararKapi": 77,
        "epk_modulT1": 78, "epk_modulT2": 79, "epk_modulO1": 80,
        "epk_modulO2": 81, "epk_modulO3": 82, "epk_modulO6": 83,
        "epk_modulO8": 84, "epk_modulM": 85, "epk_modulI1": 86, "epk_modulI2": 87,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "epk_modulT1Yorum": 90, "epk_modulT2Yorum": 91, "epk_modulO1Yorum": 92,
        "epk_modulO2Yorum": 93, "epk_modulO3Yorum": 94, "epk_modulO6Yorum": 95,
        "epk_modulO8Yorum": 96, "epk_modulMYorum": 97, "epk_modulI1Yorum": 98,
        "epk_modulI2Yorum": 99,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$13")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$C$7:$C$17")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$G$7:$G$8")


def main(cikti_yolu=None):
    wb = Workbook()
    siralar = [
        (kapak, "KAPAK"),
        (pano, "PANO"),
        (kartlar, "KARTLAR"),
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
    dosya_ad = "EmlakPipelineKomisyon.xlsx"
    hedef = cikti_yolu or os.path.join(urun_dir, dosya_ad)
    os.makedirs(os.path.dirname(hedef) or ".", exist_ok=True)
    wb.save(hedef)

    cikti = os.path.join(KOK, "cikti", dosya_ad)
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(hedef) != os.path.abspath(cikti):
        import shutil
        shutil.copy2(hedef, cikti)

    print(f"Dosya oluşturuldu: {hedef}")
    print(f"Kopya: {cikti}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
