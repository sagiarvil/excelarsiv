#!/usr/bin/env python3
"""GES Üretim & Performans Takibi — A5 saha komuta merkezi (manda v6).

PR oranı = üretim_kWh / (ışınım_kWh/m² × kapasite_kWp).
Kayıp kırılımı + varlık sağlık skoru formülden; GUNLUK_GIRIS ≤6 manuel alan;
≥12 ay sezonsallık; AYARLAR alarm eşikleri + alarm listesi.
Domain: üretim/PR/kayıp — depo/filo/GES fizibilite kopyası değil.
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
    FONT,
    GIRIS_SARI,
    GIRIS_YAZI,
    GRİ,
    KOYU_LACIVERT,
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

URUN_AD = "GES Üretim & Performans Takibi"
SURUM = "1.0.0"
RENK = "0B3D2E"
KAPASITE = 250
DEMO_SANTRAL = 8
DEMO_GUNLUK = 36
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

# (id, ad, tip, kapasite_kwp, inverter_adet, konum, hedef_pr)
SANTRALLER = [
    ("GES-00001", "Ankara Cati A", "Cati", 120.0, 4, "Ankara", 0.78),
    ("GES-00002", "Konya Saha B", "Saha", 850.0, 12, "Konya", 0.80),
    ("GES-00003", "Izmir Cati C", "Cati", 95.0, 3, "Izmir", 0.76),
    ("GES-00004", "Antalya Saha D", "Saha", 1200.0, 16, "Antalya", 0.82),
    ("GES-00005", "Bursa Cati E", "Cati", 64.0, 2, "Bursa", 0.75),
    ("GES-00006", "Gaziantep Saha F", "Saha", 520.0, 8, "Gaziantep", 0.79),
    ("GES-00007", "Mersin Cati G", "Cati", 110.0, 4, "Mersin", 0.77),
    ("GES-00008", "Eskisehir Saha H", "Saha", 380.0, 6, "Eskisehir", 0.78),
]

KAYIP_TURLERI = ["KIRLENME", "GOLGE", "INVERTER", "ARIZA", "ISINIM", "DIGER"]
AYLAR = [
    "Ocak", "Subat", "Mart", "Nisan", "Mayis", "Haziran",
    "Temmuz", "Agustos", "Eylul", "Ekim", "Kasim", "Aralik",
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: GES PR ve kayıp analizini etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_gunluk():
    satirlar = []
    for i in range(1, DEMO_GUNLUK + 1):
        san = SANTRALLER[(i - 1) % DEMO_SANTRAL]
        tar = RAPOR_TARIHI - timedelta(days=(i * 3) % 340)
        isinim = round(3.2 + (i % 7) * 0.45 + (i % 3) * 0.2, 2)
        kapasite = san[3]
        teorik = round(isinim * kapasite, 1)
        # düşük PR örnekleri: i=12, 24
        pr_carp = 0.55 if i in (12, 24) else (0.72 + (i % 5) * 0.02)
        uretim = round(teorik * pr_carp, 1)
        kayip = round(max(0, teorik - uretim), 1)
        satirlar.append({
            "id": f"GUP-{i:05d}",
            "varlik": "" if i == 30 else (san[0] if i != 33 else "GES-99999"),
            "tarih": tar,
            "uretim": uretim,
            "isinim": isinim,
            "kayip": kayip,
            "tur": KAYIP_TURLERI[i % len(KAYIP_TURLERI)],
        })
    return satirlar


ORNEK = _demo_gunluk()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "GES santrallerinde günlük üretim, ışınım ve kayıp verisinden PR oranını "
      "hesaplar; kayıp kırılımı ve varlık sağlık skoru üretir; "
      "UYGUN / DİKKAT / KRİTİK / VERİ YOK kararı verir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "PR oranı = üretim / (ışınım × kapasite) — IEC 61724 mantığı",
        "Günlük saha girişi en fazla 6 manuel alan — 30 saniye hedefi",
        "Varlık sağlık skoru formülle veriden türer (sabit skor yok)",
        "12 aylık sezonsal üretim-PR zinciri ve alarm eşiği kuyruğu",
        "Makrosuz, çevrimdışı; SCADA / API bağı yok",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Saha enerji operatörü — günlük üretim-ışınım-kayıp satırı",
        "GES işletme müdürü — PR panosu ve karar",
        "CFO / varlık yöneticisi — dönemsel RAPOR kanıt çıktısı",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) VARLIKLAR'a santral kartı girin. 2) GUNLUK_GIRIS'e günlük satır yazın. "
      "3) DONEM_ANALIZ ve PANO'dan kararı görün. 4) RAPOR'dan çıktı alın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Kod: GUP-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def varliklar(ws):
    sayfa_hazirla(ws, "VARLIKLAR", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "GES santral kartları — her varlığın tek kimliği vardır", son_kolon=14)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=gup_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4",
               "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: GES-00009 | Risk: Elle çakışma")

    sutunlar = [
        "VarlikId", "SantralAd", "Tip", "KapasiteKwp_kwp",
        "InverterAdet_adet", "Konum", "HedefPr_oran", "Aktif",
        "ToplamUretimKwh", "OrtPr", "ToplamKayipKwh", "SaglikSkor",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "ToplamUretimKwh": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'SUMIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[UretimKwh_kwh]))'
        ),
        "OrtPr": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'IFERROR(AVERAGEIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[FormulPr_oran]),0))'
        ),
        "ToplamKayipKwh": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'SUMIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[KayipKwh_kwh]))'
        ),
        "SaglikSkor": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'MAX(0,MIN(100,gup_kaliteTaban'
            '-IF(tblVarliklar[[#This Row],[OrtPr]]<gup_esikPr,gup_kaliteCeza,0)'
            '-COUNTIFS(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[AlarmBayrak],"ALARM")*gup_kaliteSapma'
            '-IF(tblVarliklar[[#This Row],[Aktif]]="HAYIR",gup_kaliteCeza,0))))'
        ),
    }
    for i, row in enumerate(SANTRALLER):
        r = ILK + i
        kid, ad, tip, kap, inv, kon, hedef = row
        _sari(ws, r, 1, kid, baslik="Varlık Kimliği", mesaj="GES-##### formatında")
        _sari(ws, r, 2, ad, baslik="Santral Adı", mesaj="Santral / saha adı")
        _sari(ws, r, 3, tip, baslik="Tip", mesaj="Cati veya Saha")
        _sari(ws, r, 4, kap, sayi="0.0", baslik="Kapasite [kWp]", mesaj="STC kapasite kWp")
        _sari(ws, r, 5, inv, sayi=CATI, baslik="İnverter Adet", mesaj="İnverter sayısı")
        _sari(ws, r, 6, kon, baslik="Konum", mesaj="İl / saha konumu")
        _sari(ws, r, 7, hedef, sayi=YÜZDE, baslik="Hedef PR", mesaj="Hedef performans oranı")
        _sari(ws, r, 8, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_SANTRAL, SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None,
                  sayi=(YÜZDE if c == 7 else ("0.0" if c == 4 else (CATI if c == 5 else None))),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVarliklar", f"A{HDR}:L{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Varlık Kimliği", mesaj="GES-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Santral Ad"), (3, "Tip"), (6, "Konum")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    for c, ad in [(4, "Kapasite [kWp]"), (5, "İnverter"), (7, "Hedef PR")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeEvetHayir", f"H{ILK}:H{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 13)})
    ws.column_dimensions["B"].width = 22


def gunluk_giris(ws):
    sayfa_hazirla(ws, "GUNLUK_GIRIS", RENK, URUN_AD, son_kolon=22)
    alt_bant(ws, 3,
             "GÜNLÜK saha girişi — satır başına en fazla 6 manuel alan (S02)",
             son_kolon=16)
    # Manuel (id/tarih/formul/skor/alarm hariç): Uretim, Isinim, Kayip, KayipTuru (+VarlikId id filtresi)
    sutunlar = [
        "SatirId", "VarlikId", "Tarih", "UretimKwh_kwh", "IsinimKwhm2_kwhm2",
        "KayipKwh_kwh", "KayipTuru",
        "FormulKapasite_kwp", "FormulTeorik_kwh", "FormulPr_oran", "FormulKayipPay_oran",
        "FormulNet_kwh", "SaglikSkor_puan", "AlarmBayrak", "FormulKayitDolu", "FormulYetim",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "FormulKapasite_kwp": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR(SUMIF(tblVarliklar[VarlikId],tblGunluk[[#This Row],[VarlikId]],'
            'tblVarliklar[KapasiteKwp_kwp]),gup_varsayilanKapasite))'
        ),
        "FormulTeorik_kwh": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'tblGunluk[[#This Row],[IsinimKwhm2_kwhm2]]*'
            'tblGunluk[[#This Row],[FormulKapasite_kwp]])'
        ),
        "FormulPr_oran": (
            'IF(OR(tblGunluk[[#This Row],[SatirId]]="",'
            'tblGunluk[[#This Row],[FormulTeorik_kwh]]=0),"",'
            'tblGunluk[[#This Row],[UretimKwh_kwh]]/'
            'tblGunluk[[#This Row],[FormulTeorik_kwh]])'
        ),
        "FormulKayipPay_oran": (
            'IF(OR(tblGunluk[[#This Row],[SatirId]]="",'
            'tblGunluk[[#This Row],[FormulTeorik_kwh]]=0),"",'
            'tblGunluk[[#This Row],[KayipKwh_kwh]]/'
            'tblGunluk[[#This Row],[FormulTeorik_kwh]])'
        ),
        "FormulNet_kwh": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'tblGunluk[[#This Row],[UretimKwh_kwh]])'
        ),
        "SaglikSkor_puan": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'MAX(0,MIN(100,gup_kaliteTaban'
            '-IF(tblGunluk[[#This Row],[FormulPr_oran]]<gup_esikPr,gup_kaliteCeza,0)'
            '-IF(tblGunluk[[#This Row],[KayipKwh_kwh]]>gup_esikKayipKwh,gup_kaliteSapma,0)'
            '-IF(tblGunluk[[#This Row],[FormulYetim]]="YETIM",gup_kaliteCeza,0))))'
        ),
        "AlarmBayrak": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IF(OR(tblGunluk[[#This Row],[FormulPr_oran]]<gup_esikPr,'
            'tblGunluk[[#This Row],[KayipKwh_kwh]]>gup_esikKayipKwh,'
            'tblGunluk[[#This Row],[SaglikSkor_puan]]<gup_esikSaglik),'
            '"ALARM","UYGUN"))'
        ),
        "FormulKayitDolu": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",1)'
        ),
        "FormulYetim": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IF(OR(tblGunluk[[#This Row],[VarlikId]]="",'
            'COUNTIF(tblVarliklar[VarlikId],tblGunluk[[#This Row],[VarlikId]])=0),'
            '"YETIM","TAMAM"))'
        ),
    }

    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="Günlük Kimlik", mesaj="GUP-#####")
        _sari(ws, r, 2, s["varlik"], baslik="Varlık Kimliği", mesaj="GES listesinden")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="GG.AA.YYYY")
        _sari(ws, r, 4, s["uretim"], sayi="0.0", baslik="Üretim [kWh]", mesaj="Günlük AC üretim")
        _sari(ws, r, 5, s["isinim"], sayi="0.00", baslik="Işınım [kWh/m2]", mesaj="Günlük ışınım")
        _sari(ws, r, 6, s["kayip"], sayi="0.0", baslik="Kayıp [kWh]", mesaj="Tahmini kayıp")
        _sari(ws, r, 7, s["tur"], baslik="Kayıp Türü", mesaj="Listeden seçin")
    for r in range(ILK + DEMO_GUNLUK, SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None,
                  sayi=(TARİH if c == 3 else ("0.0" if c in (4, 6) else ("0.00" if c == 5 else None))),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblGunluk", f"A{HDR}:P{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Günlük Kimlik", mesaj="GUP-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Varlık", mesaj="GES kimliği girin",
              hata_baslik="Varlık", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"C{ILK}:C{SON}",
              baslik="Tarih", mesaj="GG.AA.YYYY",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    for c, ad in [(4, "Üretim [kWh]"), (5, "Işınım [kWh/m2]"), (6, "Kayıp [kWh]")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeKayipTuru", f"G{ILK}:G{SON}",
              baslik="Kayıp Türü", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 13 for i in range(1, 17)})


def donem_analiz(ws):
    sayfa_hazirla(ws, "DONEM_ANALIZ", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "12 aylık sezonsallık — üretim ve PR zinciri (S03)", son_kolon=12)
    h(ws, 5, 1, "Ay", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Ay_Anahtari", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Uretim kWh", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 4, "Kayip kWh", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 5, "PR Ort", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 6, "Sezonsal Etiket", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    # Sezon: yaz yüksek, kış düşük (Türkiye GES tipik)
    demo_uretim = [82000, 95000, 128000, 155000, 178000, 195000,
                   210000, 205000, 168000, 132000, 98000, 78000]
    demo_kayip = [18000, 20000, 22000, 24000, 26000, 28000,
                  30000, 29000, 25000, 21000, 19000, 17000]
    demo_pr = [0.68, 0.70, 0.74, 0.76, 0.78, 0.80,
               0.81, 0.80, 0.77, 0.74, 0.70, 0.67]
    for i, ay in enumerate(AYLAR, 1):
        r = 5 + i
        h(ws, r, 1, ay)
        h(ws, r, 2, f"Ay_{i}")
        h(ws, r, 3, demo_uretim[i - 1], sayi=CATI)
        h(ws, r, 4, demo_kayip[i - 1], sayi=CATI)
        h(ws, r, 5, demo_pr[i - 1], sayi=YÜZDE)
        h(ws, r, 6, "12_AY_SEZONSAL")

    h(ws, 20, 1, "Sezonsal Trend Özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=gup_sezonsalTrend", kaydir=True)
    ws.merge_cells("B20:F20")
    h(ws, 22, 1, "12 Ay Toplam Üretim", kalin=True)
    h(ws, 22, 2, "=SUM(C6:C17)", sayi=CATI)
    h(ws, 23, 1, "12 Ay Toplam Kayıp", kalin=True)
    h(ws, 23, 2, "=SUM(D6:D17)", sayi=CATI)
    h(ws, 24, 1, "12 Ay Ort PR", kalin=True)
    h(ws, 24, 2, "=AVERAGE(E6:E17)", sayi=YÜZDE)
    h(ws, 25, 1, "Canlı dönem üretim", kalin=True)
    h(ws, 25, 2, "=gup_toplamUretim", sayi=CATI)
    h(ws, 26, 1, "Canlı dönem PR", kalin=True)
    h(ws, 26, 2, "=gup_prOrani", sayi=YÜZDE)

    c1 = LineChart()
    c1.title = "12 Ay Üretim Sezonsallığı"
    c1.add_data(Reference(ws, min_col=3, min_row=5, max_row=17), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17))
    c1.width, c1.height = 14, 8
    ws.add_chart(c1, "A27")

    c2 = BarChart()
    c2.title = "12 Ay PR Dağılımı"
    c2.add_data(Reference(ws, min_col=5, min_row=5, max_row=17), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17))
    c2.width, c2.height = 14, 8
    ws.add_chart(c2, "H27")

    genislik(ws, {"A": 24, "B": 14, "C": 14, "D": 14, "E": 12, "F": 18})


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Hesap motoru — ≥40 adım; sabitler AYARLAR'dan", son_kolon=10)
    h(ws, 5, 1, "Adım", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Açıklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    # Satır eşlemesi: H01=C6 … Hn = C(5+n)
    adimlar = [
        ("H01", "Toplam günlük kayıt", '=COUNTA(tblGunluk[SatirId])'),
        ("H02", "Dolu kayıt", '=SUM(tblGunluk[FormulKayitDolu])'),
        ("H03", "Toplam üretim kWh", "=SUM(tblGunluk[UretimKwh_kwh])"),
        ("H04", "Toplam ışınım", "=SUM(tblGunluk[IsinimKwhm2_kwhm2])"),
        ("H05", "Toplam kayıp kWh", "=SUM(tblGunluk[KayipKwh_kwh])"),
        ("H06", "Toplam teorik kWh", "=SUM(tblGunluk[FormulTeorik_kwh])"),
        ("H07", "PR oranı", "=IFERROR(C8/MAX(C11,1),0)"),
        ("H08", "Kayıp payı", "=IFERROR(C10/MAX(C11,1),0)"),
        ("H09", "Net üretim", "=SUM(tblGunluk[FormulNet_kwh])"),
        ("H10", "Aktif santral adet", '=COUNTIF(tblVarliklar[Aktif],"EVET")'),
        ("H11", "Toplam santral kart", '=COUNTA(tblVarliklar[VarlikId])'),
        ("H12", "Alarm adet", '=COUNTIF(tblGunluk[AlarmBayrak],"ALARM")'),
        ("H13", "Yetim kayıt", '=COUNTIF(tblGunluk[FormulYetim],"YETIM")'),
        ("H14", "Ortalama sağlık skor",
         '=IFERROR(AVERAGEIF(tblGunluk[SaglikSkor_puan],">0"),0)'),
        ("H15", "Min sağlık skor",
         '=IFERROR(MIN(tblGunluk[SaglikSkor_puan]),0)'),
        ("H16", "PR eşik altı adet",
         '=COUNTIFS(tblGunluk[FormulPr_oran],"<"&gup_esikPr,tblGunluk[FormulKayitDolu],">0")'),
        ("H17", "Kayıp eşik aşımı",
         '=COUNTIFS(tblGunluk[KayipKwh_kwh],">"&gup_esikKayipKwh,tblGunluk[FormulKayitDolu],">0")'),
        ("H18", "Düşük sağlık adet",
         '=COUNTIFS(tblGunluk[SaglikSkor_puan],"<"&gup_esikSaglik,tblGunluk[FormulKayitDolu],">0")'),
        ("H19", "Günlük PR ort",
         '=IFERROR(AVERAGEIF(tblGunluk[FormulPr_oran],">0"),0)'),
        ("H20", "Aylık üretim proxy", "=C8*gup_aylikCarpan"),
        ("H21", "PR sapma",
         "=IFERROR(ABS(C12-gup_esikPr),0)"),
        ("H22", "Kirlenme kayıp pay",
         '=IFERROR(COUNTIF(tblGunluk[KayipTuru],"KIRLENME")/MAX(C7,1),0)'),
        ("H23", "Gölge kayıp pay",
         '=IFERROR(COUNTIF(tblGunluk[KayipTuru],"GOLGE")/MAX(C7,1),0)'),
        ("H24", "İnverter kayıp pay",
         '=IFERROR(COUNTIF(tblGunluk[KayipTuru],"INVERTER")/MAX(C7,1),0)'),
        ("H25", "Arıza kayıp pay",
         '=IFERROR(COUNTIF(tblGunluk[KayipTuru],"ARIZA")/MAX(C7,1),0)'),
        ("H26", "Temkinli senaryo", "=C8*gup_senaryoTemkinli"),
        ("H27", "Baz senaryo", "=C8*gup_senaryoBaz"),
        ("H28", "İyimser senaryo", "=C8*gup_senaryoIyimser"),
        ("H29", "Senaryo farkı", "=C33-C31"),
        ("H30", "Tornado ışınım etki", "=C9*gup_tornadoOran*gup_varsayilanKapasite"),
        ("H31", "Tornado kayıp etki", "=C10*gup_tornadoOran"),
        ("H32", "Tornado PR etki", "=C8*gup_tornadoOran"),
        ("H33", "Tornado zirve", "=MAX(C35,C36,C37)"),
        ("H34", "HHI proxy",
         "=IFERROR(C27^2+C28^2+C29^2+C30^2,0)"),
        ("H35", "Kalite skor",
         "=MAX(0,gup_kaliteTaban-C21*gup_kaliteCeza-C22*gup_kaliteSapma"
         "-C23*gup_kaliteSapma)"),
        ("H36", "Tahmin aralık", "=C32*gup_tahminCarpan"),
        ("H37", "P90 üretim",
         "=IFERROR(PERCENTILE(tblGunluk[UretimKwh_kwh],gup_percentileOran),C8*gup_yuzdelikOran)"),
        ("H38", "Senaryo motor metin",
         '=_xlfn.TEXTJOIN("|",TRUE,"T",TEXT(C31,"0"),"B",TEXT(C32,"0"),"I",TEXT(C33,"0"))'),
        ("H39", "Sezonsal 12 ay üretim", "=DONEM_ANALIZ!B22"),
        ("H40", "Sezonsal 12 ay PR", "=DONEM_ANALIZ!B24"),
        ("H41", "Kritik sinyal",
         '=IF(OR(C21>0,C23>0,C19<gup_esikSaglik),1,0)'),
        ("H42", "Karar kodu ham",
         '=IF(C7=0,0,IF(C46=1,3,IF(OR(C21>0,C12<gup_esikPr),2,1)))'),
        ("H43", "PR çıpa", "=C12"),
        ("H44", "Alarm listesi özet adet", "=C17"),
        ("H45", "Varlık sağlık ort",
         '=IFERROR(AVERAGEIF(tblVarliklar[SaglikSkor],">0"),0)'),
        ("H46", "Tahmin FORECAST",
         "=IFERROR(FORECAST(C7+1,tblGunluk[UretimKwh_kwh],tblGunluk[FormulKayitDolu]),C41)"),
        ("H47", "Kayıp kırılım özet",
         '=_xlfn.TEXTJOIN("|",TRUE,"K",TEXT(C27,"0%"),"G",TEXT(C28,"0%"),'
         '"I",TEXT(C29,"0%"),"A",TEXT(C30,"0%"))'),
    ]
    # H01=C6 H02=C7 H03=C8 H04=C9 H05=C10 H06=C11 H07=C12 H08=C13 H09=C14
    # H10=C15 H11=C16 H12=C17 H13=C18 H14=C19 H15=C20 H16=C21 H17=C22 H18=C23
    # H19=C24 H20=C25 H21=C26 H22=C27 H23=C28 H24=C29 H25=C30
    # H26=C31 H27=C32 H28=C33 H29=C34 H30=C35 H31=C36 H32=C37 H33=C38
    # H34=C39 H35=C40 H36=C41 H37=C42 H38=C43 H39=C44 H40=C45 H41=C46
    # H42=C47 H43=C48 H44=C49 H45=C50 H46=C51 H47=C52

    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        acik_l = acik.lower()
        sayi = YÜZDE if any(x in acik_l for x in ("pr ", "oran", "pay", "sapma", "hhi")) else CATI
        if "metin" in acik_l or "özet" in acik_l or "kırılım" in acik_l:
            sayi = None
        elif (
            any(x in acik_l for x in ("üretim", "kayıp", "teorik", "ışınım", "tornado", "senaryo",
                                      "tahmin", "p90", "proxy"))
            and "metin" not in acik_l
            and "pr" not in acik_l
            and "pay" not in acik_l
            and "oran" not in acik_l
        ):
            sayi = "0.0"
        h(ws, r, 3, form, sayi=sayi)

    genislik(ws, {"A": 8, "B": 32, "C": 22})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Karar panosu — PR, kayıp, alarm, sezonsal trend", son_kolon=14)

    h(ws, 5, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 5, 2, "=gup_kararKapi", kalin=True, boyut=16)
    yorum_ekle(ws, "B5",
               "Tanım: Karar kapısı | Neden önemli: Boş dosyada VERİ YOK | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle müdahale yok")

    kpis = [
        (7, "Dolu Kayıt", "=gup_doluKayit", CATI),
        (8, "Toplam Üretim kWh", "=gup_toplamUretim", "0.0"),
        (9, "Toplam Kayıp kWh", "=gup_toplamKayip", "0.0"),
        (10, "PR Oranı", "=gup_prOrani", YÜZDE),
        (11, "Teorik Üretim", "=gup_teorikUretim", "0.0"),
        (12, "Kayıp Payı", "=gup_kayipPay", YÜZDE),
        (13, "Günlük PR Ort", "=gup_gunlukPrOrt", YÜZDE),
        (14, "Sağlık Skor Ort", "=gup_saglikSkor", CATI),
        (15, "Alarm Adet", "=gup_alarmAdet", CATI),
        (16, "Yetim Kayıt", "=gup_yetimAdet", CATI),
        (17, "Kalite Skor", "=gup_kaliteSkor", CATI),
        (18, "Tornado Zirve", "=gup_tornadoZirve", "0.0"),
        (19, "P90 Üretim", "=gup_yuzdelikP90", "0.0"),
        (20, "Kayıp Kırılım", "=gup_kayipKirilim", None),
    ]
    h(ws, 6, 1, "Gösterge", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 6, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for r, ad, form, sayi in kpis:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=sayi, kalin=True)

    h(ws, 7, 4, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 5, "=gup_kararGerekce", kaydir=True)
    ws.merge_cells("E7:H8")
    h(ws, 10, 4, "Öneri 1", kalin=True)
    h(ws, 10, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"PR",TEXT(gup_prOrani,"0.0%"),'
                 '"— eşik",TEXT(gup_esikPr,"0.0%"))', kaydir=True)
    h(ws, 11, 4, "Öneri 2", kalin=True)
    h(ws, 11, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Alarm",gup_alarmAdet,"adet — sağlık",'
                 'TEXT(gup_saglikSkor,"0"))', kaydir=True)
    h(ws, 12, 4, "Öneri 3", kalin=True)
    h(ws, 12, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Kayıp",TEXT(gup_toplamKayip,"0"),'
                 '"kWh — kırılım",gup_kayipKirilim)', kaydir=True)

    h(ws, 22, 1, "KayipTuru")
    h(ws, 22, 2, "Adet")
    for i, (ad, t) in enumerate([
        ("KIRLENME", 8), ("GOLGE", 6), ("INVERTER", 5), ("ARIZA", 4),
    ], 23):
        h(ws, i, 1, ad)
        h(ws, i, 2, t, sayi=CATI)

    h(ws, 22, 4, "Senaryo")
    h(ws, 22, 5, "Uretim")
    for i, (ad, t) in enumerate([
        ("Temkinli", 42000), ("Baz", 48000), ("İyimser", 54000),
    ], 23):
        h(ws, i, 4, ad)
        h(ws, i, 5, t, sayi="0.0")

    h(ws, 22, 7, "Hafta")
    h(ws, 22, 8, "Uretim")
    for i in range(1, 9):
        h(ws, 22 + i, 7, f"H{i}")
        h(ws, 22 + i, 8, 1200 + i * 80, sayi="0.0")

    h(ws, 22, 10, "Ay")
    h(ws, 22, 11, "PR")
    for i, pr in enumerate([0.68, 0.72, 0.76, 0.80], 23):
        h(ws, i, 10, AYLAR[i - 23 + 5])
        h(ws, i, 11, pr, sayi=YÜZDE)

    c1 = PieChart()
    c1.title = "Kayıp Türü Dağılımı"
    c1.add_data(Reference(ws, min_col=2, min_row=22, max_row=26), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=23, max_row=26))
    c1.width, c1.height = 10, 7
    ws.add_chart(c1, "A32")

    c2 = BarChart()
    c2.title = "Senaryo Karşılaştırma"
    c2.add_data(Reference(ws, min_col=5, min_row=22, max_row=25), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=23, max_row=25))
    c2.width, c2.height = 10, 7
    ws.add_chart(c2, "F32")

    c3 = LineChart()
    c3.title = "Haftalık Üretim Trend"
    c3.add_data(Reference(ws, min_col=8, min_row=22, max_row=30), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=7, min_row=23, max_row=30))
    c3.width, c3.height = 12, 7
    ws.add_chart(c3, "A47")

    c4 = BarChart()
    c4.title = "Aylık PR"
    c4.add_data(Reference(ws, min_col=11, min_row=22, max_row=26), titles_from_data=True)
    c4.set_categories(Reference(ws, min_col=10, min_row=23, max_row=26))
    c4.width, c4.height = 10, 7
    ws.add_chart(c4, "F47")

    c5 = PieChart()
    c5.title = "Kayıp Payı (tekrar)"
    c5.add_data(Reference(ws, min_col=2, min_row=22, max_row=25), titles_from_data=True)
    c5.set_categories(Reference(ws, min_col=1, min_row=23, max_row=25))
    c5.width, c5.height = 10, 7
    ws.add_chart(c5, "A62")

    c6 = BarChart()
    c6.title = "Senaryo vs Baz"
    c6.add_data(Reference(ws, min_col=5, min_row=22, max_row=25), titles_from_data=True)
    c6.set_categories(Reference(ws, min_col=4, min_row=23, max_row=25))
    c6.width, c6.height = 10, 7
    ws.add_chart(c6, "F62")

    c7 = LineChart()
    c7.title = "Üretim Trend (kısa)"
    c7.add_data(Reference(ws, min_col=8, min_row=22, max_row=28), titles_from_data=True)
    c7.set_categories(Reference(ws, min_col=7, min_row=23, max_row=28))
    c7.width, c7.height = 12, 7
    ws.add_chart(c7, "A77")

    c8 = BarChart()
    c8.title = "Kayıp Adet"
    c8.add_data(Reference(ws, min_col=2, min_row=22, max_row=26), titles_from_data=True)
    c8.set_categories(Reference(ws, min_col=1, min_row=23, max_row=26))
    c8.width, c8.height = 10, 7
    ws.add_chart(c8, "F77")

    baski_hazirla(ws, "A1:N30", f"{URUN_AD} | Pano | {SURUM}")
    genislik(ws, {"A": 22, "B": 16, "D": 12, "E": 40})


def alarmlar(ws):
    sayfa_hazirla(ws, "ALARMLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Alarm listesi — varlık + eşik + süre (S05)", son_kolon=12)
    h(ws, 5, 1, "Varlık", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Alarm", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Eşik", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 4, "Süre / Tarih", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 5, "PR Ort", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 6, "Sağlık Skor", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    for i, san in enumerate(SANTRALLER):
        r = 6 + i
        h(ws, r, 1, san[0])
        h(ws, r, 2,
          f'=IF(OR(COUNTIFS(tblGunluk[VarlikId],A{r},tblGunluk[AlarmBayrak],"ALARM")>0,'
          f'IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[SaglikSkor]),100)<gup_esikSaglik),'
          f'"ALARM","YOK")')
        h(ws, r, 3, f"=IFERROR(ABS(E{r}-gup_esikPr),0)", sayi=YÜZDE)
        h(ws, r, 4, "=raporTarihi", sayi=TARİH)
        h(ws, r, 5,
          f'=IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[OrtPr]),0)',
          sayi=YÜZDE)
        h(ws, r, 6,
          f'=IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[SaglikSkor]),0)',
          sayi=CATI)

    h(ws, 16, 1, "Alarm Özeti", kalin=True)
    h(ws, 16, 2, "=gup_alarmListesi", kaydir=True)
    ws.merge_cells("B16:F16")
    h(ws, 18, 1, "Toplam Alarm Satır", kalin=True)
    h(ws, 18, 2, '=COUNTIF(B6:B13,"ALARM")', sayi=CATI)
    genislik(ws, {"A": 14, "B": 12, "C": 12, "D": 14, "E": 12, "F": 12})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Dönemsel yönetim raporu — yazdırılabilir kanıt", son_kolon=10)
    h(ws, 5, 1, "Ürün", kalin=True)
    h(ws, 5, 2, URUN_AD)
    h(ws, 6, 1, "Sürüm", kalin=True)
    h(ws, 6, 2, SURUM)
    h(ws, 7, 1, "Rapor Tarihi", kalin=True)
    h(ws, 7, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 8, 1, "Firma", kalin=True)
    h(ws, 8, 2, "=gup_firmaUnvan")

    h(ws, 10, 1, "KARAR", kalin=True, boyut=12)
    h(ws, 10, 2, "=gup_kararKapi", kalin=True, boyut=14)
    h(ws, 11, 1, "PR Oranı")
    h(ws, 11, 2, "=gup_prOrani", sayi=YÜZDE)
    h(ws, 12, 1, "Toplam Üretim kWh")
    h(ws, 12, 2, "=gup_toplamUretim", sayi="0.0")
    h(ws, 13, 1, "Toplam Kayıp kWh")
    h(ws, 13, 2, "=gup_toplamKayip", sayi="0.0")
    h(ws, 14, 1, "Sağlık Skor")
    h(ws, 14, 2, "=gup_saglikSkor", sayi=CATI)
    h(ws, 15, 1, "Alarm Adet")
    h(ws, 15, 2, "=gup_alarmAdet", sayi=CATI)

    h(ws, 17, 1, "Kalem", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 17, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, (ad, form, sayi) in enumerate([
        ("Teorik üretim", "=gup_teorikUretim", "0.0"),
        ("Kayıp payı", "=gup_kayipPay", YÜZDE),
        ("Günlük PR ort", "=gup_gunlukPrOrt", YÜZDE),
        ("Kalite skor", "=gup_kaliteSkor", CATI),
    ], 18):
        h(ws, i, 1, ad)
        h(ws, i, 2, form, sayi=sayi)

    h(ws, 24, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 1,
      "PR = üretim / (ışınım × kapasite). Eşikler AYARLAR'dan gelir. "
      "Karar destek çıktısıdır; mali veya mühendislik tavsiyesi değildir. "
      "IEC 61724 referanslı performans oranı yaklaşımıdır.",
      kaydir=True)
    ws.merge_cells("A25:F26")

    h(ws, 28, 1, "İmza — GES İşletme Müdürü")
    h(ws, 28, 3, "________________")
    h(ws, 29, 1, "İmza — CFO / Varlık Yöneticisi")
    h(ws, 29, 3, "________________")
    h(ws, 31, 1, f"SHA notu: sevk paketinde SHA-256 yer alır | {SURUM}", yazi=GRİ, boyut=9)

    baski_hazirla(ws, "A1:F35", f"{URUN_AD} | Rapor | {SURUM}")
    genislik(ws, {"A": 52, "B": 28, "C": 18})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Canlı sağlık kontrolleri — formülle", son_kolon=10)
    h(ws, 5, 1, "Kontrol", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Sonuç", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    kontroller_list = [
        ("Dolu kayıt > 0", '=IF(gup_doluKayit>0,"GEÇTİ","KALDI")', "=gup_doluKayit"),
        ("Yetim kayıt = 0", '=IF(gup_yetimAdet=0,"GEÇTİ","UYARI")', "=gup_yetimAdet"),
        ("PR hesaplı", '=IF(gup_prOrani>=0,"GEÇTİ","KALDI")', "=gup_prOrani"),
        ("Sağlık skor formül", '=IF(gup_saglikSkor>=0,"GEÇTİ","KALDI")', "=gup_saglikSkor"),
        ("Alarm listesi", '=IF(gup_alarmAdet>=0,"GEÇTİ","KALDI")', "=gup_alarmAdet"),
        ("Karar dolu", '=IF(gup_kararKapi<>"","GEÇTİ","KALDI")', "=gup_kararKapi"),
        ("12 ay sezonsal", '=IF(DONEM_ANALIZ!B22>=0,"GEÇTİ","KALDI")', "=DONEM_ANALIZ!B22"),
        ("Kalite skor", '=IF(gup_kaliteSkor>=0,"GEÇTİ","KALDI")', "=gup_kaliteSkor"),
        ("Tornado", '=IF(gup_tornadoZirve>=0,"GEÇTİ","KALDI")', "=gup_tornadoZirve"),
        ("P90", '=IF(gup_yuzdelikP90>=0,"GEÇTİ","KALDI")', "=gup_yuzdelikP90"),
    ]
    for i, (ad, son, deg) in enumerate(kontroller_list, 6):
        h(ws, i, 1, ad)
        h(ws, i, 2, son)
        h(ws, i, 3, deg)
    genislik(ws, {"A": 28, "B": 12, "C": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Örnek günlük satırlar — kopyalanabilir referans", son_kolon=10)
    basliklar = [
        "SatirId", "VarlikId", "Tarih", "UretimKwh_kwh",
        "IsinimKwhm2_kwhm2", "KayipKwh_kwh", "KayipTuru",
    ]
    baslik_satiri(ws, 5, basliklar)
    for i, s in enumerate(ORNEK[:20]):
        r = 6 + i
        h(ws, r, 1, s["id"])
        h(ws, r, 2, s["varlik"])
        h(ws, r, 3, s["tarih"], sayi=TARİH)
        h(ws, r, 4, s["uretim"], sayi="0.0")
        h(ws, r, 5, s["isinim"], sayi="0.00")
        h(ws, r, 6, s["kayip"], sayi="0.0")
        h(ws, r, 7, s["tur"])
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 8)})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Evet / Hayır", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "Secim", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    _sari(ws, 7, 1, "EVET", baslik="Seçim", mesaj="EVET/HAYIR")
    _sari(ws, 8, 1, "HAYIR", baslik="Seçim", mesaj="EVET/HAYIR")

    h(ws, 3, 3, "Kayıp Türü", kalin=True)
    h(ws, 6, 3, "Tur", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, t in enumerate(KAYIP_TURLERI, 7):
        _sari(ws, i, 3, t, baslik="Kayıp", mesaj="Kayıp türü")

    h(ws, 3, 5, "Santral Tipi", kalin=True)
    h(ws, 6, 5, "Tip", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, t in enumerate(["Cati", "Saha"], 7):
        _sari(ws, i, 5, t, baslik="Tip", mesaj="Santral tipi")
    genislik(ws, {"A": 12, "C": 14, "E": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + alarm eşikleri (S05)", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("gup_varsayilanKapasite", 100.0, "kWp", "Santral tasarım / STC", "01.01.2026",
         "Varsayılan kapasite kWp"),
        ("gup_esikPr", 0.72, "oran", "Alarm eşiği — PR", "01.01.2026",
         "PR alt alarm eşiği"),
        ("gup_esikKayipKwh", 800.0, "kWh", "Alarm eşiği — kayıp", "01.01.2026",
         "Günlük kayıp alarm eşiği"),
        ("gup_esikSaglik", 60, "puan", "Alarm eşiği — sağlık skor", "01.01.2026",
         "Sağlık skor alt eşiği"),
        ("gup_aylikCarpan", 1.0, "oran", "Aylık proxy", "01.01.2026", "Aylık üretim çarpan"),
        ("gup_senaryoTemkinli", 0.90, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("gup_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("gup_senaryoIyimser", 1.10, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("gup_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("gup_yuzdelikOran", 1.15, "oran", "İstatistik kuralı", "01.01.2026", "P90 çarpan proxy"),
        ("gup_percentileOran", 0.9, "oran", "PERCENTILE oranı", "01.01.2026", "Yüzdelik dilim"),
        ("gup_tahminCarpan", 1.05, "oran", "Tahmin modeli", "01.01.2026", "Tahmin aralık çarpan"),
        ("gup_olcekHedef", 20000, "satir", "Manda A5 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("gup_kaliteTaban", 100, "puan", "Kalite / sağlık modeli", "01.01.2026", "Taban puan"),
        ("gup_kaliteCeza", 12, "puan", "Kalite / sağlık modeli", "01.01.2026", "Eşik aşım cezası"),
        ("gup_kaliteSapma", 3, "puan", "Kalite / sağlık modeli", "01.01.2026", "Sapma / alarm cezası"),
        ("gup_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("gup_firmaUnvan", "Örnek GES Enerji A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma unvanı"),
        ("gup_alarmEsikMetin", "PR / kayıp / sağlık eşikleri", "metin",
         "S05 alarm", "01.01.2026", "Alarm eşiği tanımı"),
        ("gup_tarifeTlKwh", 2.45, "TL/kWh", "İç tarife / PPA", "01.01.2026",
         "Üretim değeri TL/kWh (iç)"),
        ("gup_prReferans", 0.78, "oran", "IEC 61724 hedef", "01.01.2026",
         "Referans PR hedefi"),
        ("gup_isinimMin", 1.5, "kWh/m2", "Ölçüm kalitesi", "01.01.2026",
         "Minimum günlük ışınım"),
    ]
    for i, (ana, deg, bir, kay, yur, acik) in enumerate(params):
        r = 6 + i
        h(ws, r, 1, ana)
        if isinstance(deg, date):
            _sari(ws, r, 2, deg, sayi=TARİH, baslik=ana, mesaj=acik)
        elif isinstance(deg, float) and deg <= 1.5:
            _sari(ws, r, 2, deg, sayi=YÜZDE, baslik=ana, mesaj=acik)
        elif isinstance(deg, (int, float)):
            _sari(ws, r, 2, deg,
                  sayi=(TL if "TL" in bir else ("0.0" if isinstance(deg, float) else CATI)),
                  baslik=ana, mesaj=acik)
        else:
            _sari(ws, r, 2, deg, baslik=ana, mesaj=acik)
        h(ws, r, 3, bir)
        h(ws, r, 4, kay)
        h(ws, r, 5, yur)
        h(ws, r, 6, acik, kaydir=True)

    h(ws, 30, 8, "Motor Çıktıları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 31, 8, "anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 31, 9, "deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    motor = [
        (32, "gup_doluKayit", "=HESAP!C7"),
        (33, "gup_toplamUretim", "=HESAP!C8"),
        (34, "gup_toplamKayip", "=HESAP!C10"),
        (35, "gup_prOrani", "=HESAP!C12"),
        (36, "gup_teorikUretim", "=HESAP!C11"),
        (37, "gup_kayipPay", "=HESAP!C13"),
        (38, "gup_gunlukPrOrt", "=HESAP!C24"),
        (39, "gup_aylikUretimTrend", "=HESAP!C25"),
        (40, "gup_saglikSkor", "=HESAP!C19"),
        (41, "gup_alarmAdet", "=HESAP!C17"),
        (42, "gup_yetimAdet", "=HESAP!C18"),
        (43, "gup_prSapma", "=HESAP!C26"),
        (44, "gup_tornadoZirve", "=HESAP!C38"),
        (45, "gup_senaryoKarsilastirma", "=HESAP!C34"),
        (46, "gup_hhi", "=HESAP!C39"),
        (47, "gup_kaliteSkor", "=HESAP!C40"),
        (48, "gup_senaryoMotor", "=HESAP!C43"),
        (49, "gup_tahminAralik", "=HESAP!C41"),
        (50, "gup_yuzdelikP90", "=HESAP!C42"),
        (51, "gup_kayipKirilim", "=HESAP!C52"),
        (52, "gup_sezonsalTrend",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"12_AY",TEXT(HESAP!C44,"0"),"kWh",'
         'TEXT(HESAP!C45,"0.0%"),"PR sezonsal")'),
        (53, "gup_alarmListesi",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Alarm",HESAP!C17,"Yetim",HESAP!C18,'
         '"PrAlt",HESAP!C21,"KayipAsim",HESAP!C22)'),
        (54, "gup_kararKapi",
         '=IF(HESAP!C7=0,"VERİ YOK",'
         'IF(OR(HESAP!C8<0,HESAP!C10<0),"KRİTİK",'
         'IF(HESAP!C46=1,"KRİTİK",'
         'IF(OR(HESAP!C21>0,HESAP!C12<gup_esikPr),"DİKKAT","UYGUN"))))'),
        (55, "gup_kararGerekce",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Karar",gup_kararKapi,"| PR",'
         'TEXT(gup_prOrani,"0.0%"),"| sağlık",TEXT(gup_saglikSkor,"0"),'
         '"| alarm",gup_alarmAdet)'),
        (56, "gup_sonrakiKimlik",
         '="GES-"&TEXT(COUNTA(tblVarliklar[VarlikId])+1,"00000")'),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 58, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (59, "gup_yorumGunluk",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Günlük PR ort",TEXT(gup_gunlukPrOrt,"0.0%"))'),
        (60, "gup_yorumAylik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Aylık üretim proxy",TEXT(gup_aylikUretimTrend,"0"),"kWh")'),
        (61, "gup_yorumSapma",
         '=_xlfn.TEXTJOIN(" ",TRUE,"PR sapma",TEXT(gup_prSapma,"0.0%"))'),
        (62, "gup_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(gup_tornadoZirve,"0"),"kWh")'),
        (63, "gup_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(gup_senaryoKarsilastirma,"0"))'),
        (64, "gup_yorumHhi",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma HHI",TEXT(gup_hhi,"0.0%"))'),
        (65, "gup_yorumKalite",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite skoru",gup_kaliteSkor)'),
        (66, "gup_yorumSenMotor",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motoru",gup_senaryoMotor)'),
        (67, "gup_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(gup_tahminAralik,"0"),"kWh")'),
        (68, "gup_yorumP90",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 üretim",TEXT(gup_yuzdelikP90,"0"),"kWh")'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 24, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. VARLIKLAR sayfasına GES santral kartı ekleyin; sonraki kimliği kullanın.",
        "2. GUNLUK_GIRIS: varlık, tarih, üretim kWh, ışınım kWh/m2, kayıp kWh, kayıp türü (≤6 alan).",
        "3. PR = üretim / (ışınım × kapasite_kWp) formülüyle hesaplanır.",
        "4. DONEM_ANALIZ 12 aylık sezonsal üretim-PR zincirini gösterir.",
        "5. ALARMLAR sayfasında varlık + eşik + tarih listesi; eşikler AYARLAR'dadır.",
        "6. raporTarihi AYARLAR'dadır; TODAY kullanılmaz.",
        "7. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı hücreler açıktır.",
        "8. RAPOR sayfasını PDF olarak yazdırabilirsiniz.",
        f"9. Sürüm {SURUM} | Kod GUP-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
    ]
    for i, m in enumerate(maddeler, 5):
        h(ws, i, 1, m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    genislik(ws, {"A": 90})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 29):
        dogrulama(ws, "textLength", "0", f"D{r}",
                  baslik="Kaynak", mesaj="Kaynak metni girin",
                  hata_baslik="Kaynak", hata_mesaj="Boş olamaz",
                  isaret="greaterThanOrEqual", f2="120")
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Anahtar", mesaj="Parametre anahtarı",
                  hata_baslik="Anahtar", hata_mesaj="Boş olamaz",
                  isaret="greaterThan", f2="40")
        dogrulama(ws, "textLength", "0", f"F{r}",
                  baslik="Açıklama", mesaj="Açıklama girin",
                  hata_baslik="Açıklama", hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="200")

    ws = wb["ALARMLAR"]
    for r in range(6, 14):
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Varlık", mesaj="Varlık kimliği",
                  hata_baslik="Varlık", hata_mesaj="Boş olamaz",
                  isaret="greaterThan", f2="40")

    ws = wb["RAPOR"]
    dogrulama(ws, "textLength", "0", "B8",
              baslik="Firma", mesaj="Firma unvanı (AYARLAR'dan)",
              hata_baslik="Firma", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="80")

    ws = wb["KAPAK"]
    dogrulama(ws, "textLength", "0", "A21",
              baslik="Sürüm notu", mesaj="Sürüm satırı salt okunur referans",
              hata_baslik="Not", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="120")

    ws = wb["DONEM_ANALIZ"]
    for r in range(6, 18):
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Ay", mesaj="Ay adı",
                  hata_baslik="Ay", hata_mesaj="Boş olamaz",
                  isaret="greaterThan", f2="20")


def kosullu_bicim(wb):
    yesil = PatternFill("solid", fgColor="C6EFCE")
    sari = PatternFill("solid", fgColor="FFEB9C")
    kirmizi = PatternFill("solid", fgColor="FFC7CE")
    yf = Font(color="006100", name=FONT)
    kf = Font(color="9C0006", name=FONT)

    ws = wb["PANO"]
    for val, fill, font in [
        ('"UYGUN"', yesil, yf),
        ('"DİKKAT"', sari, None),
        ('"KRİTİK"', kirmizi, kf),
        ('"VERİ YOK"', kirmizi, kf),
    ]:
        kwargs = {"operator": "equal", "formula": [val], "fill": fill}
        if font:
            kwargs["font"] = font
        ws.conditional_formatting.add("B5", CellIsRule(**kwargs))

    for r in range(7, 21):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["ALARMLAR"]
    ws.conditional_formatting.add(
        "B6:B13", CellIsRule(operator="equal", formula=['"ALARM"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B6:B13", CellIsRule(operator="equal", formula=['"YOK"'], fill=yesil, font=yf))
    for r in range(6, 14):
        ws.conditional_formatting.add(
            f"E{r}", CellIsRule(operator="lessThan", formula=["gup_esikPr"], fill=sari))
        ws.conditional_formatting.add(
            f"F{r}", CellIsRule(operator="lessThan", formula=["gup_esikSaglik"], fill=kirmizi))

    ws = wb["GUNLUK_GIRIS"]
    ws.conditional_formatting.add(
        f"N{ILK}:N{SON}",
        CellIsRule(operator="equal", formula=['"ALARM"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"P{ILK}:P{SON}",
        CellIsRule(operator="equal", formula=['"YETIM"'], fill=sari))
    ws.conditional_formatting.add(
        f"J{ILK}:J{SON}",
        CellIsRule(operator="lessThan", formula=["gup_esikPr"], fill=sari))
    ws.conditional_formatting.add(
        f"M{ILK}:M{SON}",
        CellIsRule(operator="lessThan", formula=["gup_esikSaglik"], fill=kirmizi))

    ws = wb["VARLIKLAR"]
    ws.conditional_formatting.add(
        f"H{ILK}:H{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"L{ILK}:L{SON}",
        CellIsRule(operator="lessThan", formula=["gup_esikSaglik"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"J{ILK}:J{SON}",
        CellIsRule(operator="lessThan", formula=["gup_esikPr"], fill=sari))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add(
        "B6:B15", CellIsRule(operator="equal", formula=['"GEÇTİ"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B6:B15", CellIsRule(operator="equal", formula=['"KALDI"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B6:B15", CellIsRule(operator="equal", formula=['"UYARI"'], fill=sari))

    ws = wb["RAPOR"]
    for val, fill, font in [
        ('"UYGUN"', yesil, yf),
        ('"DİKKAT"', sari, None),
        ('"KRİTİK"', kirmizi, kf),
        ('"VERİ YOK"', kirmizi, kf),
    ]:
        kwargs = {"operator": "equal", "formula": [val], "fill": fill}
        if font:
            kwargs["font"] = font
        ws.conditional_formatting.add("B10", CellIsRule(**kwargs))
    for r in range(11, 23):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["HESAP"]
    for r in range(6, 53):
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi))

    ws = wb["DONEM_ANALIZ"]
    for r in range(6, 18):
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"E{r}", CellIsRule(operator="lessThan", formula=["gup_esikPr"], fill=sari))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "gup_varsayilanKapasite", "gup_esikPr", "gup_esikKayipKwh",
        "gup_esikSaglik", "gup_aylikCarpan", "gup_senaryoTemkinli", "gup_senaryoBaz",
        "gup_senaryoIyimser", "gup_tornadoOran", "gup_yuzdelikOran", "gup_percentileOran",
        "gup_tahminCarpan", "gup_olcekHedef", "gup_kaliteTaban", "gup_kaliteCeza",
        "gup_kaliteSapma", "gup_dosyaSurumu", "gup_firmaUnvan", "gup_alarmEsikMetin",
        "gup_tarifeTlKwh", "gup_prReferans", "gup_isinimMin",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "gup_doluKayit": 32, "gup_toplamUretim": 33, "gup_toplamKayip": 34,
        "gup_prOrani": 35, "gup_teorikUretim": 36, "gup_kayipPay": 37,
        "gup_gunlukPrOrt": 38, "gup_aylikUretimTrend": 39, "gup_saglikSkor": 40,
        "gup_alarmAdet": 41, "gup_yetimAdet": 42, "gup_prSapma": 43,
        "gup_tornadoZirve": 44, "gup_senaryoKarsilastirma": 45, "gup_hhi": 46,
        "gup_kaliteSkor": 47, "gup_senaryoMotor": 48, "gup_tahminAralik": 49,
        "gup_yuzdelikP90": 50, "gup_kayipKirilim": 51, "gup_sezonsalTrend": 52,
        "gup_alarmListesi": 53, "gup_kararKapi": 54, "gup_kararGerekce": 55,
        "gup_sonrakiKimlik": 56,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "gup_yorumGunluk": 59, "gup_yorumAylik": 60, "gup_yorumSapma": 61,
        "gup_yorumTornado": 62, "gup_yorumSenaryo": 63, "gup_yorumHhi": 64,
        "gup_yorumKalite": 65, "gup_yorumSenMotor": 66, "gup_yorumTahmin": 67,
        "gup_yorumP90": 68,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$A$7:$A$8")
    ad_ekle(wb, "ListeKayipTuru", "LISTELER!$C$7:$C$12")
    ad_ekle(wb, "ListeSantralTipi", "LISTELER!$E$7:$E$8")


def main(cikti_yolu=None):
    wb = Workbook()
    siralar = [
        (kapak, "KAPAK"),
        (varliklar, "VARLIKLAR"),
        (gunluk_giris, "GUNLUK_GIRIS"),
        (donem_analiz, "DONEM_ANALIZ"),
        (hesap, "HESAP"),
        (pano, "PANO"),
        (alarmlar, "ALARMLAR"),
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
    dosya_ad = "GesUretimPerformans.xlsx"
    hedef = cikti_yolu or os.path.join(urun_dir, dosya_ad)
    os.makedirs(os.path.dirname(hedef) or ".", exist_ok=True)
    wb.save(hedef)

    cikti = os.path.join(KOK, "cikti", dosya_ad)
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(hedef) != os.path.abspath(cikti):
        import shutil
        shutil.copy2(hedef, cikti)

    csv_yol = os.path.join(urun_dir, "ornek_veri.csv")
    with open(csv_yol, "w", encoding="utf-8") as f:
        f.write("SatirId,VarlikId,Tarih,UretimKwh_kwh,IsinimKwhm2_kwhm2,KayipKwh_kwh,KayipTuru\n")
        for s in ORNEK:
            f.write(f"{s['id']},{s['varlik']},{s['tarih']:%d.%m.%Y},"
                    f"{s['uretim']},{s['isinim']},{s['kayip']},{s['tur']}\n")

    print(f"Dosya oluşturuldu: {hedef}")
    print(f"Kopya: {cikti}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
