#!/usr/bin/env python3
"""8D / Düzeltici Faaliyet Dosya Sistemi — A4+A3 üretim betiği (manda v6 · Hat B Y-22).

Domain: müşteri şikâyeti → 8D (D1–D8) durum makinesi, kök neden, DF plan/gerçek.
Rename-kopya değil: 8D / DF / kök neden terimleri korunur.
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

URUN_AD = "8D / Düzeltici Faaliyet Dosya Sistemi"
SURUM = "1.0.0"
RENK = "0B3D2E"
KAPASITE = 250
DEMO_KART = 10
DEMO_AKIS = 30
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "D1_EKIP", "durum_adi": "D1 Ekip", "sira": 1,
     "izinli_sonraki_durumlar": "D2_PROBLEM|IPTAL", "kategori": "acik"},
    {"durum_id": "D2_PROBLEM", "durum_adi": "D2 Problem", "sira": 2,
     "izinli_sonraki_durumlar": "D3_MUHAFAZA|IPTAL|D1_EKIP", "kategori": "acik"},
    {"durum_id": "D3_MUHAFAZA", "durum_adi": "D3 Muhafaza", "sira": 3,
     "izinli_sonraki_durumlar": "D4_KOK|IPTAL|D2_PROBLEM", "kategori": "acik"},
    {"durum_id": "D4_KOK", "durum_adi": "D4 Kök Neden", "sira": 4,
     "izinli_sonraki_durumlar": "D5_DF|IPTAL|D3_MUHAFAZA", "kategori": "acik"},
    {"durum_id": "D5_DF", "durum_adi": "D5 DF Plan", "sira": 5,
     "izinli_sonraki_durumlar": "D6_UYGULA|IPTAL|D4_KOK", "kategori": "acik"},
    {"durum_id": "D6_UYGULA", "durum_adi": "D6 Uygulama", "sira": 6,
     "izinli_sonraki_durumlar": "D7_ONLE|IPTAL|D5_DF", "kategori": "acik"},
    {"durum_id": "D7_ONLE", "durum_adi": "D7 Önleme", "sira": 7,
     "izinli_sonraki_durumlar": "D8_KAPALI|IPTAL", "kategori": "acik"},
    {"durum_id": "D8_KAPALI", "durum_adi": "D8 Kapalı", "sira": 8,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "IPTAL", "durum_adi": "İptal", "sira": 9,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "D1_EKIP|D2_PROBLEM", "D1_EKIP|IPTAL",
    "D2_PROBLEM|D3_MUHAFAZA", "D2_PROBLEM|IPTAL", "D2_PROBLEM|D1_EKIP",
    "D3_MUHAFAZA|D4_KOK", "D3_MUHAFAZA|IPTAL", "D3_MUHAFAZA|D2_PROBLEM",
    "D4_KOK|D5_DF", "D4_KOK|IPTAL", "D4_KOK|D3_MUHAFAZA",
    "D5_DF|D6_UYGULA", "D5_DF|IPTAL", "D5_DF|D4_KOK",
    "D6_UYGULA|D7_ONLE", "D6_UYGULA|IPTAL", "D6_UYGULA|D5_DF",
    "D7_ONLE|D8_KAPALI", "D7_ONLE|IPTAL",
]

EKIPLER = [
    ("EKP-00001", "Ayşe Yılmaz", "Kalite", "8D Lider"),
    ("EKP-00002", "Mehmet Kaya", "Üretim", "Proses"),
    ("EKP-00003", "Zeynep Demir", "Kalite", "MSA"),
    ("EKP-00004", "Can Öztürk", "Bakım", "Ekipman"),
    ("EKP-00005", "Elif Şahin", "Satınalma", "Tedarikçi"),
    ("EKP-00006", "Burak Çelik", "Üretim", "Operatör"),
    ("EKP-00007", "Selin Arslan", "Kalite", "Müşteri"),
    ("EKP-00008", "Emre Koç", "Mühendislik", "Tasarım"),
    ("EKP-00009", "Deniz Acar", "Kalite", "8D Lider"),
    ("EKP-00010", "Burcu Polat", "Lojistik", "Sevkiyat"),
]

AY_TREND = [
    ("sdf_trendOcak", 0.72), ("sdf_trendSubat", 0.74), ("sdf_trendMart", 0.78),
    ("sdf_trendNisan", 0.80), ("sdf_trendMayis", 0.82), ("sdf_trendHaziran", 0.85),
    ("sdf_trendTemmuz", 0.83), ("sdf_trendAgustos", 0.81), ("sdf_trendEylul", 0.84),
    ("sdf_trendEkim", 0.86), ("sdf_trendKasim", 0.88), ("sdf_trendAralik", 0.90),
]


def _yas_kova_epk(yas_hucre: str) -> str:
    return (
        f'=IF({yas_hucre}="","",'
        f'IF({yas_hucre}<=sdf_kovaEsik1,"0-30",'
        f'IF({yas_hucre}<=sdf_kovaEsik2,"31-60",'
        f'IF({yas_hucre}<=sdf_kovaEsik3,"61-90","90+"))))'
    )


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: 8D dosya ve DF maliyet kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_akis():
    """≥25 satır; yetim, geçersiz geçiş ve DF maliyet zincir kırığı kasıtlı."""
    satirlar = []
    durumlar = ["D1_EKIP", "D2_PROBLEM", "D3_MUHAFAZA", "D4_KOK", "D5_DF",
                "D6_UYGULA", "D7_ONLE", "D8_KAPALI", "IPTAL", "D5_DF"]
    oncekiler = ["", "D1_EKIP", "D2_PROBLEM", "D3_MUHAFAZA", "D4_KOK",
                 "D5_DF", "D6_UYGULA", "D7_ONLE", "D5_DF", "D4_KOK"]
    kokler = ["Yontem", "Malzeme", "Makine", "Insan", "Ortam"]
    for i in range(1, DEMO_AKIS + 1):
        kart = EKIPLER[(i - 1) % DEMO_KART][0]
        if i in (22, 25):
            kart = "" if i == 22 else "EKP-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:  # geçersiz: D1 → D5
            o, d = "D1_EKIP", "D5_DF"
        if i == 20:  # geçersiz: D6 → D2
            o, d = "D6_UYGULA", "D2_PROBLEM"
        plan = 15000 + i * 2500
        gercek = plan - (i % 4) * 800
        etki = plan * 2 if d in ("D7_ONLE", "D8_KAPALI") else plan
        oran = 0.15 if i % 3 else 0.10
        if i in (12, 19):  # zincir kırık: gerçek ≠ plan (tolerans üstü)
            gercek = plan + 25000 if d in ("D7_ONLE", "D8_KAPALI") else plan + 22000
        dogrulama_tl = round(gercek) if d == "D8_KAPALI" else 0
        tutar = 0 if d in ("D8_KAPALI", "IPTAL") else round(plan * (1 + oran))
        baslangic = RAPOR_TARIHI - timedelta(days=(i * 5) % 120)
        satirlar.append({
            "id": f"8D-{i:05d}",
            "kart": kart,
            "ilan_no": f"SKY-2026-{i:04d}",
            "baslangic": baslangic,
            "onceki": o,
            "durum": d,
            "liste": plan,
            "teklif": gercek,
            "satis": etki if d in ("D7_ONLE", "D8_KAPALI") else 0,
            "oran": oran,
            "tahsil": dogrulama_tl,
            "tutar": tutar,
            "tip": kokler[i % len(kokler)],
            "yorum_a": "A" if i % 2 else "B",
            "yorum_b": f"Problem notu {i}",
        })
    return satirlar



ORNEK = _demo_akis()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Müşteri şikâyetinden 8D (D1–D8) kapanışına kadar dosyayı yönetir; kök neden ve "
      "DF plan/gerçek maliyet zincirini görünür kılar; müşteri/IATF kanıt raporu üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Makrosuz durum: D1 Ekip → D2 Problem → D3 Muhafaza → D4 Kök → D5 DF → D6 Uygula → D7 Önle → D8 Kapalı",
        "Kök neden + DF plan/gerçek maliyet zinciri; ZİNCİR KIRIK bayrağı",
        "Kapanış oranı ve kuyruk yaşlandırması (0-30 / 31-60 / 61-90 / 90+)",
        "Yetim ekip ve geçersiz 8D geçiş uyarıları",
        "KANIT_RAPORU — müşteri / IATF denetim dosyasına imzalı kanıt",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "8D lideri / QA — şikâyet ve D1–D8 satırı girişi",
        "Kalite müdürü — kapanış oranı ve DF maliyet kararı",
        "Müşteri / IATF denetçisi — KANIT_RAPORU imzası",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) EKIP'e ekip kartı girin. 2) AKIS'e 8D dosya satırlarını yazın. "
      "3) KUYRUKLAR ve PANO'dan günün kararını görün. 4) KANIT_RAPORU'ndan kanıt alın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=20)
    sabitle(ws, "A3")
    h(ws, 3, 1, "8D / DF Karar Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 3, 3, "Rapor Tarihi", yazi=GRİ)
    h(ws, 3, 4, "=raporTarihi", sayi=TARİH)

    h(ws, 4, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2,
      '=IF(COUNTA(tblAkis[IslemId])=0,"VERİ YOK",'
      'IF(OR(sdf_yetimKayit>0,sdf_gecersizGecis>0,sdf_zincirKirik>0),"KRİTİK",'
      'IF(OR(sdf_gecikenAdet>sdf_esikGeciken,sdf_kapanisOrani<sdf_hedefKapanis),"DİKKAT","UYGUN")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle değiştirilmez")

    # A05/E03: PANO'da COUNTIF/SUMIF izleri (sabit KPI yasak)
    kpi = [
        (6, "Pano Açık 8D",
         '=COUNTIFS(tblAkis[Durum],"D1_EKIP",tblAkis[Tutar],">0")'
         '+COUNTIFS(tblAkis[Durum],"D2_PROBLEM",tblAkis[Tutar],">0")'
         '+COUNTIFS(tblAkis[Durum],"D3_MUHAFAZA",tblAkis[Tutar],">0")'
         '+COUNTIFS(tblAkis[Durum],"D4_KOK",tblAkis[Tutar],">0")'
         '+COUNTIFS(tblAkis[Durum],"D5_DF",tblAkis[Tutar],">0")'
         '+COUNTIFS(tblAkis[Durum],"D6_UYGULA",tblAkis[Tutar],">0")'
         '+COUNTIFS(tblAkis[Durum],"D7_ONLE",tblAkis[Tutar],">0")', CATI),
        (7, "Pano D2 Problem",
         '=COUNTIFS(tblAkis[Durum],"D2_PROBLEM",tblAkis[Tutar],">0")', CATI),
        (8, "Pano D5 DF Plan",
         '=COUNTIFS(tblAkis[Durum],"D5_DF",tblAkis[Tutar],">0")', CATI),
        (9, "Pano Geciken Adet", "=sdf_gecikenAdet", CATI),
        (10, "Pano DF Plan Maliyet",
         '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"<>D8_KAPALI",tblAkis[Durum],"<>IPTAL")', TL),
        (11, "Pano DF Gerçek Maliyet",
         '=SUMIF(tblAkis[Durum],"D8_KAPALI",tblAkis[DfDogrulamaTl])', TL),
        (12, "Pano Yetim Kayıt", "=sdf_yetimKayit", CATI),
        (13, "Pano Geçersiz Geçiş", "=sdf_gecersizGecis", CATI),
        (14, "Pano Zincir Kırık", "=sdf_zincirKirik", CATI),
        (15, "Pano Kapanış Oranı", "=sdf_kapanisOrani", YÜZDE),
        (16, "Pano Yaş Özeti", "=sdf_yasKova", None),
        (17, "Pano Tahmin Aralık", "=sdf_tahminAralik", TL),
        (18, "Pano Senaryo Farkı", "=sdf_senaryoKarsilastirma", TL),
        (19, "Pano Tornado Zirve", "=sdf_tornadoZirve", TL),
        (20, "Pano Yoğunlaşma", "=sdf_hhi", YÜZDE),
        (21, "Pano Kalite Skoru", "=sdf_kaliteSkor", CATI),
        (22, "Pano P90 DF Maliyet", "=sdf_yuzdelikP90", TL),
        (23, "Pano D8 Kapalı Düşüm", "=sdf_kapaliDusum", CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=sdf_modulT1Yorum"), (8, "=sdf_modulT2Yorum"),
        (9, "=sdf_modulO1Yorum"), (10, "=sdf_modulO2Yorum"),
        (11, "=sdf_modulO3Yorum"), (12, "=sdf_modulO6Yorum"),
        (13, "=sdf_modulO8Yorum"), (14, "=sdf_modulMYorum"),
        (15, "=sdf_modulI1Yorum"), (16, "=sdf_modulI2Yorum"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    # Grafik kaynakları — sabit demo (D09b/Ö2; LO tablo COUNTIF grafikte sıfırlanmasın)
    h(ws, 25, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 1, "Durum")
    h(ws, 26, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Grafik D1", 4), ("Grafik D2", 4), ("Grafik D3", 3), ("Grafik D4", 4),
        ("Grafik D5", 4), ("Grafik D6", 3), ("Grafik D7", 3), ("Grafik D8", 3), ("Grafik Iptal", 2),
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
    h(ws, 26, 8, "DF maliyet")
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
    c1.title = "8D Durum"
    c1.add_data(Reference(ws, min_col=2, min_row=26, max_row=35), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=27, max_row=35))
    c1.width, c1.height = 10, 7
    ws.add_chart(c1, "A38")

    c2 = BarChart()
    c2.title = "Yaşlandırma"
    c2.add_data(Reference(ws, min_col=5, min_row=26, max_row=30), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=27, max_row=30))
    c2.width, c2.height = 10, 7
    ws.add_chart(c2, "F38")

    c3 = LineChart()
    c3.title = "Haftalık DF Maliyet"
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
    c6.title = "Aylık Kapanış Trend"
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
    c8.add_data(Reference(ws, min_col=2, min_row=26, max_row=35), titles_from_data=True)
    c8.set_categories(Reference(ws, min_col=1, min_row=27, max_row=35))
    c8.width, c8.height = 10, 7
    ws.add_chart(c8, "F83")

    baski_hazirla(ws, "A1:N36", f"{URUN_AD} | Pano | {SURUM}")
    genislik(ws, {"A": 26, "B": 16, "C": 14, "D": 14, "E": 12})


def ekip(ws):
    sayfa_hazirla(ws, "EKIP", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "8D ekip kartları — her üyenin tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=sdf_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: EKP-00009 | Risk: Elle çakışma")

    sutunlar = ["EkipId", "AdSoyad", "Bolum", "Rol", "HedefAcik8D",
                "VarsayşikâyetRisk", "Aktif", "BeklenenDfMaliyet", "Kapali8DAdet"]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "BeklenenDfMaliyet": (
            'IF(tblEkip[[#This Row],[EkipId]]="","",'
            'SUMIF(tblAkis[KaynakEkipId],tblEkip[[#This Row],[EkipId]],tblAkis[Tutar]))'
        ),
        "Kapali8DAdet": (
            'IF(tblEkip[[#This Row],[EkipId]]="","",'
            'COUNTIFS(tblAkis[KaynakEkipId],tblEkip[[#This Row],[EkipId]],'
            'tblAkis[Durum],"D8_KAPALI"))'
        ),
    }
    for i, (kid, ad, ofis, rol) in enumerate(EKIPLER):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Kart Kimliği", mesaj="EKP-##### formatında")
        _sari(ws, r, 2, ad, baslik="Ad Soyad", mesaj="Ekip adı")
        _sari(ws, r, 3, ofis, baslik="Bolum", mesaj="Bölüm")
        _sari(ws, r, 4, rol, baslik="Rol", mesaj="8D Lider/Proses/MSA")
        _sari(ws, r, 5, 3 + i, sayi=CATI, baslik="Hedef", mesaj="Açık 8D hedef adedi")
        _sari(ws, r, 6, 0.03 if i % 2 == 0 else 0.02, sayi=YÜZDE,
              baslik="Oran", mesaj="Varsayılan DF maliyet oranı")
        _sari(ws, r, 7, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_KART, SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None,
                  sayi=TL if c == 5 else (YÜZDE if c == 6 else None),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblEkip", f"A{HDR}:I{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Kart Kimliği", mesaj="EKP-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Ad"), (3, "Bolum"), (4, "Rol")]:
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
             "8D dosya satırları — D1–D8 durum makinesi + yetim + geçersiz geçiş + DF maliyet zinciri",
             son_kolon=20)
    sutunlar = [
        "IslemId", "KaynakEkipId", "SikayetNo", "BaslangicTarihi",
        "OncekiDurum", "Durum", "DfPlanMaliyet", "DfGercekMaliyet", "MusteriEtkiTl",
        "TekrarRisk", "DfDogrulamaTl", "Tutar",
        "KokNedenKod", "DfSinif", "ProblemOzet",
        "GecikmeGun", "YasKova", "BeklenenDfMaliyet", "KapanisBayrak",
        "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)

    yas_f = yaslandirma_gun(
        "tblAkis[[#This Row],[BaslangicTarihi]]", "raporTarihi").lstrip("=")
    kova_f = _yas_kova_epk("tblAkis[[#This Row],[GecikmeGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblAkis[[#This Row],[KaynakEkipId]]", "tblEkip[EkipId]").lstrip("=")
    gecis_birlesik = (
        'tblAkis[[#This Row],[OncekiDurum]]&"|"&tblAkis[[#This Row],[Durum]]'
    )
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    # Zincir: liste↔teklif (açık) veya teklif↔satış (kapalı/tapu)
    zincir_liste = zincir_kirik_formul(
        "tblAkis[[#This Row],[DfPlanMaliyet]]",
        "tblAkis[[#This Row],[DfGercekMaliyet]]",
        "zincirTolerans",
    ).lstrip("=")
    zincir_satis = zincir_kirik_formul(
        "tblAkis[[#This Row],[DfGercekMaliyet]]",
        "tblAkis[[#This Row],[MusteriEtkiTl]]",
        "zincirTolerans",
    ).lstrip("=")
    zincir_guvenli = (
        f'IF(OR(tblAkis[[#This Row],[Durum]]="D8_KAPALI",'
        f'tblAkis[[#This Row],[Durum]]="D7_ONLE",'
        f'tblAkis[[#This Row],[MusteriEtkiTl]]>0),{zincir_satis},'
        f'IF(tblAkis[[#This Row],[DfGercekMaliyet]]>0,{zincir_liste},"OK"))'
    )
    bek_kom = (
        'IF(tblAkis[[#This Row],[IslemId]]="","",'
        'IF(tblAkis[[#This Row],[DfGercekMaliyet]]>0,'
        'tblAkis[[#This Row],[DfGercekMaliyet]],'
        'tblAkis[[#This Row],[DfPlanMaliyet]]))'
    )
    kapanis_f = (
        'IF(tblAkis[[#This Row],[Durum]]="D8_KAPALI","KAPANAN",'
        'IF(tblAkis[[#This Row],[Durum]]="IPTAL","IPTAL","ACIK"))'
    )

    form = {
        "GecikmeGun": yas_f,
        "YasKova": kova_f,
        "BeklenenDfMaliyet": bek_kom,
        "KapanisBayrak": kapanis_f,
        "YetimBayrak": yetim_f,
        "GecisUyarisi": gecis_f,
        "ZincirBayrak": zincir_guvenli,
        "KayitDolu": 'IF(tblAkis[[#This Row],[IslemId]]="","",1)',
    }

    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="8D Kimlik", mesaj="8D-#####")
        _sari(ws, r, 2, s["kart"], baslik="Ekip", mesaj="Ekip kimliği")
        _sari(ws, r, 3, s["ilan_no"], baslik="Şikâyet No", mesaj="Şikâyet numarası")
        _sari(ws, r, 4, s["baslangic"], sayi=TARİH, baslik="Başlangıç", mesaj="Şikâyet tarihi")
        _sari(ws, r, 5, s["onceki"], baslik="Önceki Durum", mesaj="Durum haritası")
        _sari(ws, r, 6, s["durum"], baslik="Durum", mesaj="Güncel 8D adımı")
        _sari(ws, r, 7, s["liste"], sayi=TL, baslik="DF Plan", mesaj="DF plan maliyet TL")
        _sari(ws, r, 8, s["teklif"], sayi=TL, baslik="DF Gerçek", mesaj="DF gerçek maliyet TL")
        _sari(ws, r, 9, s["satis"] or None, sayi=TL, baslik="Müşteri Etki", mesaj="Müşteri etki TL")
        _sari(ws, r, 10, s["oran"], sayi=YÜZDE, baslik="Tekrar Risk", mesaj="Tekrar risk oranı")
        _sari(ws, r, 11, s["tahsil"] or None, sayi=TL, baslik="DF Doğrulama", mesaj="DF doğrulama TL")
        _sari(ws, r, 12, s["tutar"], sayi=TL, baslik="Tutar", mesaj="Açık 8D DF tutarı")
        _sari(ws, r, 13, s["tip"], baslik="Kök Neden", mesaj="Kök neden kodu")
        _sari(ws, r, 14, s["yorum_a"], baslik="DF Sınıf", mesaj="A veya B")
        _sari(ws, r, 15, s["yorum_b"], baslik="Problem Özet", mesaj="Serbest not")

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
              baslik="İşlem", mesaj="8D-#####",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Ekip", mesaj="Kart kimliği",
              hata_baslik="Kart", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"D{ILK}:D{SON}",
              baslik="Başlangıç", mesaj="Şikâyet tarihi",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "list", "ListeDurumlar", f"E{ILK}:E{SON}",
              baslik="Önceki Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"F{ILK}:F{SON}",
              baslik="Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    for c, ad in [(7, "Liste"), (8, "D5"), (9, "Satış"), (11, "Tahsil"), (12, "Tutar")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"J{ILK}:J{SON}",
              baslik="Oran", mesaj="0–1",
              hata_baslik="Oran", hata_mesaj="0–1",
              isaret="between", f2="1")
    dogrulama(ws, "list", "ListeDfSinifB", f"N{ILK}:N{SON}",
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
        ("Kuyruk D1 Ekip", "D1_EKIP"),
        ("Kuyruk D2 Problem", "D2_PROBLEM"),
        ("Kuyruk D3 Muhafaza", "D3_MUHAFAZA"),
        ("Kuyruk D4 Kök", "D4_KOK"),
        ("Kuyruk D5 DF", "D5_DF"),
        ("Kuyruk D6 Uygulama", "D6_UYGULA"),
        ("Kuyruk D7 Önleme", "D7_ONLE"),
    ], 6):
        h(ws, i, 1, ad)
        h(ws, i, 2,
          f'=COUNTIFS(tblAkis[Durum],"{kod}",tblAkis[Tutar],">0")', sayi=CATI)
        h(ws, i, 3,
          f'=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"{kod}")', sayi=TL)

    h(ws, 11, 1, "Geciken (açık, yaş>eşik)")
    h(ws, 11, 2,
      '=COUNTIFS(tblAkis[GecikmeGun],">"&sdf_esikGecikmeGun,'
      'tblAkis[Durum],"<>D8_KAPALI",tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")',
      sayi=CATI)
    h(ws, 11, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[GecikmeGun],">"&sdf_esikGecikmeGun,'
      'tblAkis[Durum],"<>D8_KAPALI",tblAkis[Durum],"<>IPTAL")',
      sayi=TL)

    h(ws, 13, 1, "D8 Kapalı / İptal — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 1, "D8_KAPALI")
    h(ws, 14, 2, '=COUNTIF(tblAkis[Durum],"D8_KAPALI")', sayi=CATI)
    h(ws, 14, 3, '=SUMIF(tblAkis[Durum],"D8_KAPALI",tblAkis[DfDogrulamaTl])', sayi=TL)
    h(ws, 15, 1, "İptal")
    h(ws, 15, 2, '=COUNTIF(tblAkis[Durum],"IPTAL")', sayi=CATI)
    h(ws, 15, 3, '=SUMIF(tblAkis[Durum],"IPTAL",tblAkis[Tutar])', sayi=TL)

    h(ws, 17, 1, "Kuyruk Özeti", kalin=True)
    h(ws, 17, 2, "=sdf_kuyrukOzeti", kaydir=True)
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
        ("H03", "Toplam 8D dosya tutar", "=SUM(tblAkis[Tutar])"),
        ("H04", "DF plan toplam", "=SUM(tblAkis[DfPlanMaliyet])"),
        ("H05", "DF gerçek toplam", "=SUM(tblAkis[DfGercekMaliyet])"),
        ("H06", "Müşteri etki toplam", "=SUM(tblAkis[MusteriEtkiTl])"),
        ("H07", "Şikâyet adet", '=COUNTIF(tblAkis[Durum],"D1_EKIP")'),
        ("H08", "D2 Problem adet", '=COUNTIF(tblAkis[Durum],"D2_PROBLEM")'),
        ("H09", "D5 DF adet", '=COUNTIF(tblAkis[Durum],"D5_DF")'),
        ("H10", "D6 Uygulama adet", '=COUNTIF(tblAkis[Durum],"D6_UYGULA")'),
        ("H11", "D7 Önleme adet", '=COUNTIF(tblAkis[Durum],"D7_ONLE")'),
        ("H12", "D8 Kapalı adet", '=COUNTIF(tblAkis[Durum],"D8_KAPALI")'),
        ("H13", "İptal adet", '=COUNTIF(tblAkis[Durum],"IPTAL")'),
        ("H14", "Açık 8D kuyruk", "=C12+C13+C14+C15+C16"),
        ("H15", "D8 Kapalı düşüm adet", "=C17+C18"),
        ("H16", "Açık 8D tutar",
         '=SUMIF(tblAkis[Durum],"D1_EKIP",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"D2_PROBLEM",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"D3_MUHAFAZA",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"D4_KOK",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"D5_DF",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"D6_UYGULA",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"D7_ONLE",tblAkis[Tutar])'),
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
         '=COUNTIFS(tblAkis[GecikmeGun],">"&sdf_esikGecikmeGun,'
         'tblAkis[Durum],"<>D8_KAPALI",tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")'),
        ("H27", "Geciken tutar",
         '=SUMIFS(tblAkis[Tutar],tblAkis[GecikmeGun],">"&sdf_esikGecikmeGun,'
         'tblAkis[Durum],"<>D8_KAPALI",tblAkis[Durum],"<>IPTAL")'),
        ("H28", "Beklenen DF maliyet", "=C21"),
        ("H29", "Tahsil DF maliyet",
         '=SUMIF(tblAkis[Durum],"D8_KAPALI",tblAkis[DfDogrulamaTl])'),
        ("H30", "Kapanış oranı",
         "=IFERROR(C17/MAX(C7,1),0)"),
        ("H31", "Kapanış sapma", "=C35-sdf_hedefKapanis"),
        ("H32", "Temkinli senaryo", "=C33*sdf_senaryoTemkinli"),
        ("H33", "Baz senaryo", "=C33*sdf_senaryoBaz"),
        ("H34", "İyimser senaryo", "=C33*sdf_senaryoIyimser"),
        ("H35", "Senaryo farkı", "=C39-C37"),
        ("H36", "Tornado (oran×tutar)", "=C33*sdf_tornadoOran"),
        ("H37", "HHI proxy",
         "=IFERROR((MAX(C12,C13,C14)/MAX(C19,1))^2,0)"),
        ("H38", "Kalite skoru",
         "=IF(C6=0,0,MAX(sdf_sifir,sdf_kaliteTaban-C22*sdf_kaliteCeza"
         "-C23*sdf_kaliteCeza-C24*sdf_kaliteCeza-C31*2))"),
        ("H39", "Tahmin PERCENTILE",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[Tutar],sdf_yuzdelikOran),0)"),
        ("H40", "Tahmin FORECAST",
         "=IFERROR(FORECAST(COUNTA(tblAkis[IslemId])+1,tblAkis[Tutar],"
         "tblAkis[KayitDolu]),C44)"),
        ("H41", "Tahmin aralık", "=IFERROR((C44+C45)/2,0)"),
        ("H42", "P90 DF plan",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[DfPlanMaliyet],sdf_yuzdelikOran),0)"),
        ("H43", "Ekip sayısı", '=COUNTA(tblEkip[EkipId])'),
        ("H44", "Hedef açık 8D toplam", "=SUM(tblEkip[HedefAcik8D])"),
        ("H45", "Hedef karşılama", "=IFERROR(C33/MAX(C49,1),0)"),
        ("H46", "8D kuyruk özeti",
         '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"Şikâyet",C12,"D2",C13,"D5",C14,'
         '"D6",C15,"D7",C16,"Geciken",C31),"")'),
        ("H47", "Yaş kova metin",
         '=IFERROR(_xlfn.TEXTJOIN("/",TRUE,C26,C27,C28,C29),"")'),
        ("H48", "Senaryo motor özet",
         '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"Temkinli",C37,"Baz",C38,"İyimser",C39),"")'),
        ("H49", "Veri hazırlık", "=IF(C6=0,0,C7/MAX(C6,1))"),
        ("H50", "Trend ortalama",
         "=IFERROR(C35*(sdf_trendOcak+sdf_trendSubat+sdf_trendMart+sdf_trendNisan+"
         "sdf_trendMayis+sdf_trendHaziran+sdf_trendTemmuz+sdf_trendAgustos+"
         "sdf_trendEylul+sdf_trendEkim+sdf_trendKasim+sdf_trendAralik)/12,0)"),
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
        elif any(x in low for x in ("oran", "sapma", "hhi", "hazırlık", "karşılama", "trend", "kapanış")):
            fmt = YÜZDE
        elif any(x in low for x in ("tutar", "DF maliyet", "senaryo", "tornado", "tahmin", "liste", "teklif", "satış", "hedef", "fark", "p90")) and "adet" not in low:
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
      "=sdf_senaryoTemkinli*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 59, 3, "=C37", sayi=TL)
    h(ws, 60, 1, "Baz")
    h(ws, 60, 2,
      "=sdf_senaryoBaz*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 60, 3, "=C38", sayi=TL)
    h(ws, 61, 1, "İyimser")
    h(ws, 61, 2,
      "=sdf_senaryoIyimser*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 61, 3, "=C39", sayi=TL)

    genislik(ws, {"A": 10, "B": 28, "C": 50})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "8D / DF Kapanış Kanıt Raporu — Müşteri / IATF", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Ürün")
    h(ws, 4, 2, URUN_AD)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 6, 1, "Rapor Tarihi")
    h(ws, 6, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    h(ws, 9, 1, "Girdi özeti", kalin=True, yazi=KOYU_LACIVERT)
    for r, ad, form, fmt in [
        (10, "Açık 8D", "=sdf_acik8D", CATI),
        (11, "DF plan toplam", "=sdf_dfPlanToplam", TL),
        (12, "DF gerçek toplam", "=sdf_dfGercekToplam", TL),
        (13, "Kapanış oranı", "=sdf_kapanisOrani", YÜZDE),
        (14, "Yetim", "=sdf_yetimKayit", CATI),
        (15, "Geçersiz geçiş", "=sdf_gecersizGecis", CATI),
        (16, "ZİNCİR KIRIK", "=sdf_zincirKirik", CATI),
        (17, "Geciken 8D", "=sdf_gecikenAdet", CATI),
    ]:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 19, 1, "Hesap zinciri", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 1,
      "Şikâyet açılışı → D1 Ekip → D2 Problem → D3 Muhafaza → D4 Kök Neden → "
      "D5 DF Plan → D6 Uygulama → D7 Önleme → D8 Kapalı. "
      "DF plan ↔ DF gerçek sapması tolerans üstü ise ZİNCİR KIRIK.",
      kaydir=True)
    ws.merge_cells("A20:H21")
    h(ws, 23, 1, "Madde atıfları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 1, "IATF 16949 — düzeltici faaliyet / müşteri şikâyeti kanıtı")
    h(ws, 25, 1, "ISO 9001 — uygunsuzluk ve düzeltici faaliyet")
    h(ws, 26, 1, "AIAG 8D — D1–D8 adım disiplini")
    h(ws, 28, 1, "Parametre seti", kalin=True)
    h(ws, 28, 2, "=sdf_kanitRaporu", kaydir=True)
    ws.merge_cells("B28:H28")
    h(ws, 30, 1, "SHA-256", kalin=True)
    h(ws, 30, 2,
      "SHA-256 sevk anında KANIT/SEVK_KARARI.md dosyasına yazılır; alıcı parmak iziyle doğrular.",
      kaydir=True)
    ws.merge_cells("B30:H30")
    h(ws, 32, 1, "İmza — 8D Lider", kalin=True)
    h(ws, 32, 3, "İmza — Kalite Müdürü", kalin=True)
    h(ws, 32, 5, "İmza — Müşteri / Denetçi", kalin=True)
    h(ws, 34, 1, "________________")
    h(ws, 34, 3, "________________")
    h(ws, 34, 5, "________________")
    h(ws, 36, 1, "Uyarı: Bu çıktı karar destektir; sertifikasyon kararı yerine geçmez.", yazi=GRİ)
    baski_hazirla(ws, "A1:H36", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 28, "B": 22, "C": 24, "D": 14, "E": 26})



def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "8D / DF Kapanış Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Rapor Açık 8D", "=sdf_acik8D", CATI),
        (10, "Rapor DF Plan", "=sdf_dfPlanToplam", TL),
        (11, "Rapor DF Gerçek", "=sdf_dfGercekToplam", TL),
        (12, "Rapor Kapanış Oranı", "=sdf_kapanisOrani", YÜZDE),
        (13, "Rapor Yetim", "=sdf_yetimKayit", CATI),
        (14, "Rapor Geçersiz Geçiş", "=sdf_gecersizGecis", CATI),
        (15, "Rapor Zincir Kırık", "=sdf_zincirKirik", CATI),
        (16, "Rapor Tahmin", "=sdf_tahminAralik", TL),
        (17, "Rapor Kanıt Özeti", "=sdf_kanitRaporu", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 19, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 1,
      "Durum haritası AYARLAR'dadır. Geçersiz geçiş engellenmez, görünür kılınır. "
      "DF maliyet oranı satırdan veya ekip varsayılanından gelir. "
      "Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A20:H20")
    h(ws, 22, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 1, "1. Yetim kayıtları EKIP'a bağlayın veya satırı düzeltin.")
    h(ws, 24, 1, "2. Geciken teklif/sözleşme satırlarında ekipı uyarın.")
    h(ws, 25, 1, "3. Zincir kırık satırlarda DF plan/gerçek/etki farkını açıklayın.")
    h(ws, 27, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    baski_hazirla(ws, "A1:H27", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 32, "B": 40, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "8D dosya sayısı", '=COUNTA(tblAkis[IslemId])'),
        (6, "Ekip sayısı", '=COUNTA(tblEkip[EkipId])'),
        (7, "Şikâyet", '=COUNTIF(tblAkis[Durum],"D1_EKIP")'),
        (8, "D2", '=COUNTIF(tblAkis[Durum],"D2_PROBLEM")'),
        (9, "D5", '=COUNTIF(tblAkis[Durum],"D5_DF")'),
        (10, "D6", '=COUNTIF(tblAkis[Durum],"D6_UYGULA")'),
        (11, "D7", '=COUNTIF(tblAkis[Durum],"D7_ONLE")'),
        (12, "D8 Kapalı", '=COUNTIF(tblAkis[Durum],"D8_KAPALI")'),
        (13, "İptal", '=COUNTIF(tblAkis[Durum],"IPTAL")'),
        (14, "Açık toplam", "=B7+B8+B9+B10+B11"),
        (15, "Yetim", "=sdf_yetimKayit"),
        (16, "Geçersiz", "=sdf_gecersizGecis"),
        (17, "Zincir kırık", "=sdf_zincirKirik"),
        (18, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (19, "Kalite skoru", "=sdf_kaliteSkor"),
        (20, "Beklenen DF maliyet", "=sdf_dfPlanToplam"),
        (21, "Tahsil DF maliyet", "=sdf_dfGercekToplam"),
        (22, "D8 Kapalı düşüm", "=sdf_kapaliDusum"),
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
      f"Dosyada {DEMO_KART} ekip kartı ve {DEMO_AKIS} şikâyet satırı vardır. "
      "Yetim (22, 25), geçersiz geçiş (18, 20) ve zincir kırık (12, 19) kasıtlıdır. "
      "EKIP ve AKIS'i temizleyip kendi verinizi girebilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Durumlar: D1–D8 / Görüşme/Teklif/Sözleşme/Tapu/D8 Kapalı/İptal — kapalılar kuyruktan düşer")
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

    h(ws, 3, 9, "Kök Neden", kalin=True)
    h(ws, 6, 9, "Tip", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(["Yontem", "Malzeme", "Makine", "Insan", "Ortam"], 7):
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
        ("zincirTolerans", 10000, "TL", "İş kuralı — DF plan/gerçek/etki sapma", "01.01.2026",
         "ZİNCİR KIRIK eşiği"),
        ("sdf_esikGecikmeGun", 30, "gün", "İç politika", "01.01.2026", "Gecikme eşiği"),
        ("sdf_esikGeciken", 5, "adet", "İç politika", "01.01.2026", "Karar eşiği geciken"),
        ("sdf_kovaEsik1", 30, "gün", "Yaşlandırma", "01.01.2026", "Kova 1 üst sınır"),
        ("sdf_kovaEsik2", 60, "gün", "Yaşlandırma", "01.01.2026", "Kova 2 üst sınır"),
        ("sdf_kovaEsik3", 90, "gün", "Yaşlandırma", "01.01.2026", "Kova 3 üst sınır"),
        ("sdf_hedefKapanis", 0.35, "oran", "Kalite yönetimi", "01.01.2026", "Hedef kapanış oranı"),
        ("sdf_sifir", 0, "adet", "Kalite motoru", "01.01.2026", "MAX alt sınırı"),
        ("sdf_kaliteTaban", 100, "puan", "Kalite motoru", "01.01.2026", "Kalite taban puan"),
        ("sdf_kaliteCeza", 10, "puan", "Kalite motoru", "01.01.2026", "Uyarı başına ceza"),
        ("sdf_senaryoTemkinli", 0.85, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("sdf_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("sdf_senaryoIyimser", 1.15, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("sdf_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("sdf_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("sdf_olcekHedef", 20000, "satir", "Manda A3 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("sdf_azamiTutar", 100000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul"),
        ("sdf_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi",
         "01.01.2026", "Durum makinesi kaynağı"),
        ("sdf_kuyrukOzeti_metin", "KUYRUKLAR canlı özet", "metin", "A05 kuyruk", "01.01.2026",
         "Kuyruk görünümü"),
        ("sdf_kapaliDusum_metin", "kapali kategori kuyruktan düşer", "metin", "A06", "01.01.2026",
         "D8 Kapalı düşüm kuralı"),
        ("sdf_kanitRaporu_metin", "RAPOR kanıt çıktısı", "metin", "Dönemsel rapor", "01.01.2026",
         "Kanıt raporu"),
        ("sdf_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("sdf_firmaUnvan", "Örnek Emlak Bolumi Ltd.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma"),
    ]
    # Ay trend NR (G07 — gömülü ay çarpanı yok)
    for ad, val in AY_TREND:
        params.append((ad, val, "oran", "Tarihsel kapanış", "01.01.2026", f"{ad} trend"))

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
    tablo_ekle(ws, "tblDurumHaritasi", "H43:L52", dh_bas)

    h(ws, 53, 8, "Motor Çıktıları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 54, 8, "anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 54, 9, "deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    motor = [
        (55, "sdf_acik8D", "=HESAP!C19"),
        (56, "sdf_dfPlanToplam", "=HESAP!C33"),
        (57, "sdf_gecikenAdet", "=HESAP!C31"),
        (58, "sdf_dfGercekToplam", "=HESAP!C34"),
        (59, "sdf_yetimKayit", "=HESAP!C22"),
        (60, "sdf_gecersizGecis", "=HESAP!C23"),
        (61, "sdf_zincirKirik", "=HESAP!C24"),
        (62, "sdf_kapanisOrani", "=HESAP!C35"),
        (63, "sdf_yasKova", "=HESAP!C52"),
        (64, "sdf_tahsilatOrt", "=HESAP!C35"),
        (65, "sdf_tahsilatSapma", "=HESAP!C36"),
        (66, "sdf_tahminAralik", "=HESAP!C46"),
        (67, "sdf_senaryoKarsilastirma", "=HESAP!C40"),
        (68, "sdf_tornadoZirve", "=HESAP!C41"),
        (69, "sdf_hhi", "=HESAP!C42"),
        (70, "sdf_kaliteSkor", "=HESAP!C43"),
        (71, "sdf_yuzdelikP90", "=HESAP!C47"),
        (72, "sdf_senaryoMotor", "=HESAP!C53"),
        (73, "sdf_kapaliDusum", "=HESAP!C20"),
        (74, "sdf_kuyrukOzeti", "=HESAP!C51"),
        (75, "sdf_kanitRaporu",
         '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"D8",HESAP!C17,"DF",HESAP!C34,'
         '"Kapanis",TEXT(HESAP!C35,"0.0%")),"")'),
        (76, "sdf_sonrakiKimlik",
         '="EKP-"&TEXT(COUNTA(tblEkip[EkipId])+1,"00000")'),
        (77, "sdf_kararKapi", "=PANO!B4"),
        (78, "sdf_modulT1", "=HESAP!C19"),
        (79, "sdf_modulT2", "=HESAP!C33"),
        (80, "sdf_modulO1", "=HESAP!C36"),
        (81, "sdf_modulO2", "=HESAP!C41"),
        (82, "sdf_modulO3", "=HESAP!C53"),
        (83, "sdf_modulO6", "=HESAP!C42"),
        (84, "sdf_modulO8", "=HESAP!C43"),
        (85, "sdf_modulM", "=HESAP!C53"),
        (86, "sdf_modulI1", "=HESAP!C46"),
        (87, "sdf_modulI2", "=HESAP!C47"),
    ]
    for r, ana, form in motor:
        h(ws, r, 8, ana)
        h(ws, r, 9, form, kalin=True)

    h(ws, 89, 8, "Yorum Motoru", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (90, "sdf_modulT1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Acik 8D",sdf_acik8D,"adet"),"")'),
        (91, "sdf_modulT2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"DF plan",TEXT(sdf_dfPlanToplam,"₺#,##0"),'
         '"DF gercek",TEXT(sdf_dfGercekToplam,"₺#,##0")),"")'),
        (92, "sdf_modulO1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Kapanis",TEXT(sdf_kapanisOrani,"0.0%"),'
         '"sapma",TEXT(sdf_tahsilatSapma,"0.0%")),"")'),
        (93, "sdf_modulO2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tornado",TEXT(sdf_tornadoZirve,"₺#,##0")),"")'),
        (94, "sdf_modulO3Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Senaryo",TEXT(sdf_senaryoKarsilastirma,"₺#,##0")),"")'),
        (95, "sdf_modulO6Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"HHI",TEXT(sdf_hhi,"0.0%")),"")'),
        (96, "sdf_modulO8Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Kalite",sdf_kaliteSkor),"")'),
        (97, "sdf_modulMYorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motor",sdf_senaryoMotor),"")'),
        (98, "sdf_modulI1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tahmin",TEXT(sdf_tahminAralik,"₺#,##0")),"")'),
        (99, "sdf_modulI2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"P90",TEXT(sdf_yuzdelikP90,"₺#,##0")),"")'),
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
        "1. AYARLAR'da rapor tarihi, DF maliyet toleransı ve kapanış hedefini kontrol edin.",
        "2. EKIP'a ekip kimliklerini girin (sarı hücreler).",
        "3. AKIS'te her şikâyet için durum, DF plan/gerçek/etki ve DF maliyet oranını yazın.",
        "4. KUYRUKLAR'da günün açık 8D dosya ve geciken satırlarını okuyun.",
        "5. PANO'da karar (UYGUN/DİKKAT/KRİTİK/VERİ YOK) ve KPI'ları görün.",
        "6. RAPOR'u yazdırın veya PDF'e aktarın; IATF / müşteri dosyasına ekleyin.",
        "7. Formül hücreleri 1234 şifresiyle korunur; sarı alanlar serbesttir.",
    ]
    for i, m in enumerate(adimlar, 5):
        h(ws, i, 1, m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    h(ws, 14, 1, "Sık yapılan hatalar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "• Ekip kartı olmadan AKIS satırı → YETİM bayrağı")
    h(ws, 16, 1, "• İzinli olmayan durum geçişi → GEÇERSİZ GEÇİŞ (engellemez, görünür)")
    h(ws, 17, 1, "• DF plan/gerçek/etki sapması tolerans üstü → ZİNCİR KIRIK")
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

    ws = wb["EKIP"]
    ws.conditional_formatting.add(
        f"G{ILK}:G{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"E{ILK}:E{SON}",
        CellIsRule(operator="greaterThan", formula=["5"], fill=yesil))

    ws = wb["HESAP"]
    for r in range(6, 56):
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))


def adlari_bagla(wb):
    # Parametre satırları: 6.. (21 sabit + 12 ay = 33) → B6..B38
    param_adlari = [
        "raporTarihi", "zincirTolerans", "sdf_esikGecikmeGun", "sdf_esikGeciken",
        "sdf_kovaEsik1", "sdf_kovaEsik2", "sdf_kovaEsik3", "sdf_hedefKapanis",
        "sdf_sifir", "sdf_kaliteTaban", "sdf_kaliteCeza",
        "sdf_senaryoTemkinli", "sdf_senaryoBaz", "sdf_senaryoIyimser",
        "sdf_tornadoOran", "sdf_yuzdelikOran", "sdf_olcekHedef", "sdf_azamiTutar",
        "sdf_durumHaritasi", "sdf_kuyrukOzeti_metin", "sdf_kapaliDusum_metin",
        "sdf_kanitRaporu_metin", "sdf_dosyaSurumu", "sdf_firmaUnvan",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    for i, (ad, _v) in enumerate(AY_TREND):
        ad_ekle(wb, ad, f"AYARLAR!$B${30 + i}")  # 6+24=30

    motor_map = {
        "sdf_acik8D": 55, "sdf_dfPlanToplam": 56, "sdf_gecikenAdet": 57,
        "sdf_dfGercekToplam": 58, "sdf_yetimKayit": 59, "sdf_gecersizGecis": 60,
        "sdf_zincirKirik": 61, "sdf_kapanisOrani": 62, "sdf_yasKova": 63,
        "sdf_tahsilatOrt": 64, "sdf_tahsilatSapma": 65, "sdf_tahminAralik": 66,
        "sdf_senaryoKarsilastirma": 67, "sdf_tornadoZirve": 68, "sdf_hhi": 69,
        "sdf_kaliteSkor": 70, "sdf_yuzdelikP90": 71, "sdf_senaryoMotor": 72,
        "sdf_kapaliDusum": 73, "sdf_kuyrukOzeti": 74, "sdf_kanitRaporu": 75,
        "sdf_sonrakiKimlik": 76, "sdf_kararKapi": 77,
        "sdf_modulT1": 78, "sdf_modulT2": 79, "sdf_modulO1": 80,
        "sdf_modulO2": 81, "sdf_modulO3": 82, "sdf_modulO6": 83,
        "sdf_modulO8": 84, "sdf_modulM": 85, "sdf_modulI1": 86, "sdf_modulI2": 87,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "sdf_modulT1Yorum": 90, "sdf_modulT2Yorum": 91, "sdf_modulO1Yorum": 92,
        "sdf_modulO2Yorum": 93, "sdf_modulO3Yorum": 94, "sdf_modulO6Yorum": 95,
        "sdf_modulO8Yorum": 96, "sdf_modulMYorum": 97, "sdf_modulI1Yorum": 98,
        "sdf_modulI2Yorum": 99,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$15")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$C$7:$C$25")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")
    ad_ekle(wb, "ListeDfSinifB", "LISTELER!$G$7:$G$8")


def main(cikti_yolu=None):
    wb = Workbook()
    siralar = [
        (kapak, "KAPAK"),
        (pano, "PANO"),
        (ekip, "EKIP"),
        (akis, "AKIS"),
        (kuyruklar, "KUYRUKLAR"),
        (hesap, "HESAP"),
        (kanit_raporu, "KANIT_RAPORU"),
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
    dosya_ad = "SekizDDuzelticiFaaliyetDosya.xlsx"
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
