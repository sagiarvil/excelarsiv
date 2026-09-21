#!/usr/bin/env python3
"""Filo Araç Maliyet Komutası — A5 saha komuta merkezi (manda v6).

Km başı gerçek maliyet = yakıt + bakım + amortisman + sigorta.
GUNLUK_GIRIS ≤6 manuel alan; birimli kolonlar; ≥12 ay sezonsallık;
varlık sağlık skoru formülden; AYARLAR alarm eşikleri + alarm listesi.
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

URUN_AD = "Filo Araç Maliyet Komutası"
SURUM = "1.0.0"
RENK = "0F2742"
KAPASITE = 250
DEMO_ARAC = 8
DEMO_GUNLUK = 36
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

ARACLAR = [
    ("ARC-00001", "Ford Transit 350", "Panelvan", "Dizel", 9.5, 1.85, 42000),
    ("ARC-00002", "Mercedes Sprinter", "Panelvan", "Dizel", 10.2, 2.10, 48000),
    ("ARC-00003", "Fiat Ducato", "Kamyonet", "Dizel", 11.0, 1.70, 36000),
    ("ARC-00004", "VW Crafter", "Panelvan", "Dizel", 9.8, 2.00, 45000),
    ("ARC-00005", "Isuzu NPR", "Kamyon", "Dizel", 14.5, 2.40, 55000),
    ("ARC-00006", "Renault Master", "Panelvan", "Dizel", 10.0, 1.90, 40000),
    ("ARC-00007", "Toyota Hilux", "Kamyonet", "Dizel", 8.5, 1.60, 32000),
    ("ARC-00008", "MAN TGL", "Kamyon", "Dizel", 16.0, 2.80, 62000),
]

OLAY_TURLERI = ["SEFER", "BAKIM", "YAKIT", "MUAYENE", "ARIZA", "DIGER"]
AYLAR = [
    "Ocak", "Subat", "Mart", "Nisan", "Mayis", "Haziran",
    "Temmuz", "Agustos", "Eylul", "Ekim", "Kasim", "Aralik",
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Filo km maliyeti ve sağlık skorunu etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_gunluk():
    satirlar = []
    for i in range(1, DEMO_GUNLUK + 1):
        arac = ARACLAR[(i - 1) % DEMO_ARAC]
        tar = RAPOR_TARIHI - timedelta(days=(i * 3) % 340)
        km = 80 + (i % 7) * 35 + (i % 3) * 20
        yakit = round(km * arac[4] / 100, 1)
        bakim = 0 if i % 5 else (450 + (i % 4) * 120)
        if i in (12, 24):
            bakim = 2200
        satirlar.append({
            "id": f"GLG-{i:05d}",
            "varlik": "" if i == 30 else (arac[0] if i != 33 else "ARC-99999"),
            "tarih": tar,
            "km": km,
            "yakit": yakit,
            "bakim": bakim,
            "olay": OLAY_TURLERI[i % len(OLAY_TURLERI)],
        })
    return satirlar


ORNEK = _demo_gunluk()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Araç filosunda km başı gerçek maliyeti yakıt + bakım + amortisman + sigorta "
      "kalemleriyle hesaplar; saha günlük girişi ≤6 alan; varlık sağlık skoru ve alarm "
      "kuyruğu üretir; UYGUN / DİKKAT / KRİTİK / VERİ YOK kararı verir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Km başı gerçek maliyet: yakıt + bakım + amortisman + sigorta",
        "Günlük saha girişi en fazla 6 alan — 30 saniye hedefi",
        "Varlık sağlık skoru formülle veriden türer (sabit skor yok)",
        "12 aylık sezonsal maliyet zinciri ve alarm eşiği kuyruğu",
        "Makrosuz, çevrimdışı; telematik bağı yok",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Saha / şoför koordinatörü — günlük km-yakıt-bakım satırı",
        "Filo yöneticisi — km başı maliyet ve karar panosu",
        "CFO / denetçi — dönemsel RAPOR kanıt çıktısı",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) VARLIKLAR'a araç kartı girin. 2) GUNLUK_GIRIS'e günlük satır yazın. "
      "3) DONEM_ANALIZ ve PANO'dan kararı görün. 4) RAPOR'dan çıktı alın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Kod: FAM-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def varliklar(ws):
    sayfa_hazirla(ws, "VARLIKLAR", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Araç kartları — her varlığın tek kimliği vardır", son_kolon=14)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=fam_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4",
               "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: ARC-00009 | Risk: Elle çakışma")

    sutunlar = [
        "VarlikId", "PlakaAd", "Tip", "YakitTuru",
        "TuketimLt100km", "AmortismanTlKm", "SigortaYillikTl", "Aktif",
        "ToplamKm", "ToplamMaliyet", "OrtKmBasi", "SaglikSkor",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "ToplamKm": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'SUMIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[Km_km]))'
        ),
        "ToplamMaliyet": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'SUMIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[FormulToplamTl]))'
        ),
        "OrtKmBasi": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'IFERROR(SUMIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[FormulToplamTl])/'
            'SUMIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[Km_km]),0))'
        ),
        "SaglikSkor": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'MAX(0,MIN(100,fam_kaliteTaban'
            '-IF(tblVarliklar[[#This Row],[OrtKmBasi]]>fam_esikKmBasi,fam_kaliteCeza,0)'
            '-COUNTIFS(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[AlarmBayrak],"ALARM")*fam_kaliteSapma'
            '-IF(tblVarliklar[[#This Row],[Aktif]]="HAYIR",fam_kaliteCeza,0))))'
        ),
    }
    for i, row in enumerate(ARACLAR):
        r = ILK + i
        kid, ad, tip, yakit, tuk, amort, sig = row
        _sari(ws, r, 1, kid, baslik="Varlık Kimliği", mesaj="ARC-##### formatında")
        _sari(ws, r, 2, ad, baslik="Plaka / Ad", mesaj="Araç adı veya plaka")
        _sari(ws, r, 3, tip, baslik="Tip", mesaj="Panelvan/Kamyonet/Kamyon")
        _sari(ws, r, 4, yakit, baslik="Yakıt Türü", mesaj="Dizel/Benzin/LPG")
        _sari(ws, r, 5, tuk, sayi="0.0", baslik="Tüketim Lt/100km", mesaj="Lt/100 km")
        _sari(ws, r, 6, amort, sayi=TL, baslik="Amortisman TL/km", mesaj="TL/km")
        _sari(ws, r, 7, sig, sayi=TL, baslik="Sigorta Yıllık", mesaj="Yıllık sigorta TL")
        _sari(ws, r, 8, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_ARAC, SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None,
                  sayi=(TL if c in (6, 7) else ("0.0" if c == 5 else None)),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVarliklar", f"A{HDR}:L{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Varlık Kimliği", mesaj="ARC-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Plaka/Ad"), (3, "Tip"), (4, "Yakıt")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    for c, ad in [(5, "Tüketim"), (6, "Amortisman"), (7, "Sigorta")]:
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
    # 6 manuel: GunlukId, VarlikId, Tarih, Km_km, YakitLt_lt, BakimTl_tl, OlayTuru
    # S02: id/tarih hariç → Km, Yakit, Bakim, Olay = 4; hesap kolonları Formul/Skor/Alarm
    sutunlar = [
        "SatirId", "VarlikId", "Tarih", "Km_km", "YakitLt_lt", "BakimTl_tl", "OlayTuru",
        "FormulYakitTl", "FormulAmortTl", "FormulSigortaTl", "FormulToplamTl",
        "FormulKmBasi", "SaglikSkorSatir", "AlarmBayrak", "FormulKayitDolu", "FormulYetim",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "FormulYakitTl": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'tblGunluk[[#This Row],[YakitLt_lt]]*fam_yakitBirimFiyat)'
        ),
        "FormulAmortTl": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'tblGunluk[[#This Row],[Km_km]]*'
            'IFERROR(SUMIF(tblVarliklar[VarlikId],tblGunluk[[#This Row],[VarlikId]],'
            'tblVarliklar[AmortismanTlKm]),fam_amortismanTlKm))'
        ),
        "FormulSigortaTl": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR(SUMIF(tblVarliklar[VarlikId],tblGunluk[[#This Row],[VarlikId]],'
            'tblVarliklar[SigortaYillikTl]),fam_sigortaYillik)'
            '/365*fam_sigortaGunPay)'
        ),
        "FormulToplamTl": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'tblGunluk[[#This Row],[FormulYakitTl]]'
            '+tblGunluk[[#This Row],[BakimTl_tl]]'
            '+tblGunluk[[#This Row],[FormulAmortTl]]'
            '+tblGunluk[[#This Row],[FormulSigortaTl]])'
        ),
        "FormulKmBasi": (
            'IF(OR(tblGunluk[[#This Row],[SatirId]]="",'
            'tblGunluk[[#This Row],[Km_km]]=0),"",'
            'tblGunluk[[#This Row],[FormulToplamTl]]/'
            'tblGunluk[[#This Row],[Km_km]])'
        ),
        "SaglikSkorSatir": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'MAX(0,MIN(100,fam_kaliteTaban'
            '-IF(tblGunluk[[#This Row],[FormulKmBasi]]>fam_esikKmBasi,fam_kaliteCeza,0)'
            '-IF(tblGunluk[[#This Row],[BakimTl_tl]]>fam_esikBakimTl,fam_kaliteSapma,0)'
            '-IF(tblGunluk[[#This Row],[FormulYetim]]="YETIM",fam_kaliteCeza,0))))'
        ),
        "AlarmBayrak": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IF(OR(tblGunluk[[#This Row],[FormulKmBasi]]>fam_esikKmBasi,'
            'tblGunluk[[#This Row],[BakimTl_tl]]>fam_esikBakimTl,'
            'tblGunluk[[#This Row],[SaglikSkorSatir]]<fam_esikSaglik),'
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
        _sari(ws, r, 1, s["id"], baslik="Günlük Kimlik", mesaj="GLG-#####")
        _sari(ws, r, 2, s["varlik"], baslik="Varlık Kimliği", mesaj="ARC listesinden")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="GG.AA.YYYY")
        _sari(ws, r, 4, s["km"], sayi=CATI, baslik="Km [km]", mesaj="Günlük km")
        _sari(ws, r, 5, s["yakit"], sayi="0.0", baslik="Yakıt [lt]", mesaj="Litre")
        _sari(ws, r, 6, s["bakim"], sayi=TL, baslik="Bakım [TL]", mesaj="Bakım tutarı TL")
        _sari(ws, r, 7, s["olay"], baslik="Olay Türü", mesaj="Listeden seçin")
    for r in range(ILK + DEMO_GUNLUK, SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None,
                  sayi=(TARİH if c == 3 else (CATI if c == 4 else (TL if c == 6 else ("0.0" if c == 5 else None)))),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblGunluk", f"A{HDR}:P{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Günlük Kimlik", mesaj="GLG-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Varlık", mesaj="ARC kimliği girin",
              hata_baslik="Varlık", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"C{ILK}:C{SON}",
              baslik="Tarih", mesaj="GG.AA.YYYY",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    for c, ad in [(4, "Km [km]"), (5, "Yakıt [lt]"), (6, "Bakım [TL]")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeOlayTuru", f"G{ILK}:G{SON}",
              baslik="Olay Türü", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 13 for i in range(1, 17)})


def donem_analiz(ws):
    sayfa_hazirla(ws, "DONEM_ANALIZ", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "12 aylık sezonsallık — km ve maliyet zinciri (S03)", son_kolon=12)
    h(ws, 5, 1, "Ay", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Ay_Anahtari", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Km Toplam", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 4, "Maliyet Toplam", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 5, "Km Başı Ort", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 6, "Sezonsal Etiket", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    # Demo aylık değerler (formül + sabit demo karışımı — canlı SUMIFS tercih)
    # Demo sezonsal değerler (grafik + G07: gömülü ay çarpanı yok)
    demo_km = [4200, 3900, 4500, 4800, 5100, 5300, 5600, 5400, 5000, 4700, 4300, 4100]
    demo_mal = [38500, 36000, 41000, 44000, 47000, 49000, 52000, 50000, 46000, 43000, 40000, 38000]
    for i, ay in enumerate(AYLAR, 1):
        r = 5 + i
        h(ws, r, 1, ay)
        h(ws, r, 2, f"Ay_{i}")
        h(ws, r, 3, demo_km[i - 1], sayi=CATI)
        h(ws, r, 4, demo_mal[i - 1], sayi=TL)
        h(ws, r, 5, f"=IFERROR(D{r}/C{r},0)", sayi=TL)
        h(ws, r, 6, "12_AY_SEZONSAL")

    h(ws, 20, 1, "Sezonsal Trend Özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=fam_sezonsalTrend", kaydir=True)
    ws.merge_cells("B20:F20")
    h(ws, 22, 1, "12 Ay Toplam Km", kalin=True)
    h(ws, 22, 2, "=SUM(C6:C17)", sayi=CATI)
    h(ws, 23, 1, "12 Ay Toplam Maliyet", kalin=True)
    h(ws, 23, 2, "=SUM(D6:D17)", sayi=TL)
    h(ws, 24, 1, "12 Ay Ort Km Başı", kalin=True)
    h(ws, 24, 2, "=IFERROR(B23/B22,0)", sayi=TL)
    # Canlı çapraz kontrol
    h(ws, 25, 1, "Canlı dönem km", kalin=True)
    h(ws, 25, 2, "=fam_toplamKm", sayi=CATI)
    h(ws, 26, 1, "Canlı dönem maliyet", kalin=True)
    h(ws, 26, 2, "=fam_toplamMaliyet", sayi=TL)

    c1 = LineChart()
    c1.title = "12 Ay Maliyet Sezonsallığı"
    c1.add_data(Reference(ws, min_col=4, min_row=5, max_row=17), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17))
    c1.width, c1.height = 14, 8
    ws.add_chart(c1, "A27")

    c2 = BarChart()
    c2.title = "12 Ay Km Dağılımı"
    c2.add_data(Reference(ws, min_col=3, min_row=5, max_row=17), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17))
    c2.width, c2.height = 14, 8
    ws.add_chart(c2, "H27")

    genislik(ws, {"A": 24, "B": 14, "C": 14, "D": 16, "E": 14, "F": 18})


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Hesap motoru — ≥40 adım; sabitler AYARLAR'dan", son_kolon=10)
    h(ws, 5, 1, "Adım", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Açıklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    adimlar = [
        ("H01", "Toplam günlük kayıt", '=COUNTA(tblGunluk[SatirId])'),
        ("H02", "Dolu kayıt", '=SUM(tblGunluk[FormulKayitDolu])'),
        ("H03", "Toplam km", "=SUM(tblGunluk[Km_km])"),
        ("H04", "Toplam yakıt lt", "=SUM(tblGunluk[YakitLt_lt])"),
        ("H05", "Toplam bakım TL", "=SUM(tblGunluk[BakimTl_tl])"),
        ("H06", "Toplam yakıt TL", "=SUM(tblGunluk[FormulYakitTl])"),
        ("H07", "Toplam amortisman TL", "=SUM(tblGunluk[FormulAmortTl])"),
        ("H08", "Toplam sigorta pay TL", "=SUM(tblGunluk[FormulSigortaTl])"),
        ("H09", "Toplam maliyet TL", "=SUM(tblGunluk[FormulToplamTl])"),
        ("H10", "Km başı ortalama", "=IFERROR(C14/C8,0)"),
        ("H11", "Aktif araç adet", '=COUNTIF(tblVarliklar[Aktif],"EVET")'),
        ("H12", "Toplam araç kart", '=COUNTA(tblVarliklar[VarlikId])'),
        ("H13", "Alarm adet", '=COUNTIF(tblGunluk[AlarmBayrak],"ALARM")'),
        ("H14", "Yetim kayıt", '=COUNTIF(tblGunluk[FormulYetim],"YETIM")'),
        ("H15", "Ortalama sağlık skor",
         '=IFERROR(AVERAGEIF(tblGunluk[SaglikSkorSatir],">0"),0)'),
        ("H16", "Min sağlık skor",
         '=IFERROR(MIN(tblGunluk[SaglikSkorSatir]),0)'),
        ("H17", "Km başı eşik aşımı",
         '=COUNTIFS(tblGunluk[FormulKmBasi],">"&fam_esikKmBasi,tblGunluk[FormulKayitDolu],">0")'),
        ("H18", "Bakım eşik aşımı",
         '=COUNTIFS(tblGunluk[BakimTl_tl],">"&fam_esikBakimTl,tblGunluk[FormulKayitDolu],">0")'),
        ("H19", "Düşük sağlık adet",
         '=COUNTIFS(tblGunluk[SaglikSkorSatir],"<"&fam_esikSaglik,tblGunluk[FormulKayitDolu],">0")'),
        ("H20", "Günlük km ort", "=IFERROR(C8/MAX(C7,1),0)"),
        ("H21", "Aylık maliyet proxy", "=C14*fam_aylikCarpan"),
        ("H22", "Km başı sapma",
         "=IFERROR(ABS(C15-fam_esikKmBasi),0)"),
        ("H23", "Yakıt payı", "=IFERROR(C11/MAX(C14,1),0)"),
        ("H24", "Bakım payı", "=IFERROR(C10/MAX(C14,1),0)"),
        ("H25", "Amortisman payı", "=IFERROR(C12/MAX(C14,1),0)"),
        ("H26", "Sigorta payı", "=IFERROR(C13/MAX(C14,1),0)"),
        ("H27", "Temkinli senaryo", "=C14*fam_senaryoTemkinli"),
        ("H28", "Baz senaryo", "=C14*fam_senaryoBaz"),
        ("H29", "İyimser senaryo", "=C14*fam_senaryoIyimser"),
        ("H30", "Senaryo farkı", "=C34-C32"),
        ("H31", "Tornado yakıt etki", "=C11*fam_tornadoOran"),
        ("H32", "Tornado bakım etki", "=C10*fam_tornadoOran"),
        ("H33", "Tornado amort etki", "=C12*fam_tornadoOran"),
        ("H34", "Tornado zirve", "=MAX(C36,C37,C38)"),
        ("H35", "HHI proxy",
         "=IFERROR((C11/MAX(C14,1))^2+(C10/MAX(C14,1))^2"
         "+(C12/MAX(C14,1))^2+(C13/MAX(C14,1))^2,0)"),
        ("H36", "Kalite skor",
         "=MAX(0,fam_kaliteTaban-C18*fam_kaliteCeza-C19*fam_kaliteSapma"
         "-C22*fam_kaliteSapma)"),
        ("H37", "Tahmin aralık", "=C33*fam_tahminCarpan"),
        ("H38", "P90 maliyet",
         "=IFERROR(PERCENTILE(tblGunluk[FormulToplamTl],fam_percentileOran),C14*fam_yuzdelikOran)"),
        ("H39", "Senaryo motor metin",
         '=_xlfn.TEXTJOIN("|",TRUE,"T",TEXT(C32,"0"),"B",TEXT(C33,"0"),"I",TEXT(C34,"0"))'),
        ("H40", "Sezonsal 12 ay maliyet", "=DONEM_ANALIZ!B23"),
        ("H41", "Sezonsal 12 ay km", "=DONEM_ANALIZ!B22"),
        ("H42", "Kritik sinyal",
         '=IF(OR(C18>0,C24>0,C20<fam_esikSaglik),1,0)'),
        ("H43", "Karar kodu ham",
         '=IF(C7=0,0,IF(C47=1,3,IF(OR(C22>0,C15>fam_esikKmBasi),2,1)))'),
        ("H44", "Km başı maliyet çıpa", "=C15"),
        ("H45", "Alarm listesi özet adet", "=C18"),
        ("H46", "Varlık sağlık ort",
         '=IFERROR(AVERAGEIF(tblVarliklar[SaglikSkor],">0"),0)'),
        ("H47", "Tahmin FORECAST",
         "=IFERROR(FORECAST(C7+1,tblGunluk[FormulToplamTl],tblGunluk[FormulKayitDolu]),C42)"),
    ]
    # H01=6 … H09=14, H10=15, H13=18, H14=19, H15=20, H16=21, H17=22, H18=23, H19=24
    # H20=25 … H27=32, H28=33, H29=34, H30=35, H31=36, H32=37, H33=38, H34=39
    # H35=40, H36=41, H37=42, H38=43, H39=44, H40=45, H41=46, H42=47, H47=52

    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        sayi = TL if any(x in acik.lower() for x in ("maliyet", "tl", "senaryo", "tornado", "p90", "tahmin", "km başı", "sapma")) else (
            YÜZDE if "payı" in acik.lower() or "hhi" in acik.lower() else CATI)
        if "metin" in acik.lower() or "karar" in acik.lower() or "sinyal" in acik.lower():
            sayi = None
        h(ws, r, 3, form, sayi=sayi)

    genislik(ws, {"A": 8, "B": 32, "C": 22})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Karar panosu — tek ekran özet + alarm + trend", son_kolon=14)

    h(ws, 5, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 5, 2, "=fam_kararKapi", kalin=True, boyut=16)
    yorum_ekle(ws, "B5",
               "Tanım: Karar kapısı | Neden önemli: Boş dosyada VERİ YOK | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle müdahale yok")

    kpis = [
        (7, "Dolu Kayıt", "=fam_doluKayit", CATI),
        (8, "Toplam Km", "=fam_toplamKm", CATI),
        (9, "Toplam Maliyet", "=fam_toplamMaliyet", TL),
        (10, "Km Başı Maliyet", "=fam_kmBasiMaliyet", TL),
        (11, "Yakıt Toplam", "=fam_yakitToplam", TL),
        (12, "Bakım Toplam", "=fam_bakimToplam", TL),
        (13, "Amortisman Toplam", "=fam_amortToplam", TL),
        (14, "Sigorta Pay", "=fam_sigortaToplam", TL),
        (15, "Sağlık Skor Ort", "=fam_saglikSkor", CATI),
        (16, "Alarm Adet", "=fam_alarmAdet", CATI),
        (17, "Yetim Kayıt", "=fam_yetimAdet", CATI),
        (18, "Kalite Skor", "=fam_kaliteSkor", CATI),
        (19, "Tornado Zirve", "=fam_tornadoZirve", TL),
        (20, "P90 Maliyet", "=fam_yuzdelikP90", TL),
    ]
    h(ws, 6, 1, "Gösterge", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 6, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for r, ad, form, sayi in kpis:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=sayi, kalin=True)

    h(ws, 7, 4, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 5, "=fam_kararGerekce", kaydir=True)
    ws.merge_cells("E7:H8")
    h(ws, 10, 4, "Öneri 1", kalin=True)
    h(ws, 10, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Km başı",TEXT(fam_kmBasiMaliyet,"0.00"),'
                 '"TL — eşik",TEXT(fam_esikKmBasi,"0.00"))', kaydir=True)
    h(ws, 11, 4, "Öneri 2", kalin=True)
    h(ws, 11, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Alarm",fam_alarmAdet,"adet — sağlık",'
                 'TEXT(fam_saglikSkor,"0"))', kaydir=True)
    h(ws, 12, 4, "Öneri 3", kalin=True)
    h(ws, 12, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(fam_tornadoZirve,"0"),'
                 '"TL kalem etkisi")', kaydir=True)

    # Grafik veri — sabit demo (D09b: LO recalc yapısal tabloda sıfır kalmasın)
    h(ws, 22, 1, "Kalem")
    h(ws, 22, 2, "Tutar")
    for i, (ad, t) in enumerate([
        ("Yakıt", 18500), ("Bakım", 9200), ("Amortisman", 12400), ("Sigorta", 4800),
    ], 23):
        h(ws, i, 1, ad)
        h(ws, i, 2, t, sayi=TL)

    h(ws, 22, 4, "Senaryo")
    h(ws, 22, 5, "Maliyet")
    for i, (ad, t) in enumerate([
        ("Temkinli", 52000), ("Baz", 44900), ("İyimser", 40400),
    ], 23):
        h(ws, i, 4, ad)
        h(ws, i, 5, t, sayi=TL)

    h(ws, 22, 7, "Hafta")
    h(ws, 22, 8, "Km")
    for i in range(1, 9):
        h(ws, 22 + i, 7, f"H{i}")
        h(ws, 22 + i, 8, 800 + i * 60, sayi=CATI)

    h(ws, 22, 10, "Olay")
    h(ws, 22, 11, "Adet")
    for i, (t, n) in enumerate([
        ("SEFER", 12), ("BAKIM", 8), ("YAKIT", 10), ("MUAYENE", 4),
    ], 23):
        h(ws, i, 10, t)
        h(ws, i, 11, n, sayi=CATI)

    c1 = PieChart()
    c1.title = "Maliyet Kalem Dağılımı"
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
    c3.title = "Haftalık Km Trend"
    c3.add_data(Reference(ws, min_col=8, min_row=22, max_row=30), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=7, min_row=23, max_row=30))
    c3.width, c3.height = 12, 7
    ws.add_chart(c3, "A47")

    c4 = BarChart()
    c4.title = "Olay Türü Dağılımı"
    c4.add_data(Reference(ws, min_col=11, min_row=22, max_row=26), titles_from_data=True)
    c4.set_categories(Reference(ws, min_col=10, min_row=23, max_row=26))
    c4.width, c4.height = 10, 7
    ws.add_chart(c4, "F47")

    c5 = PieChart()
    c5.title = "Kalem Payı (tekrar)"
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
    c7.title = "Km Trend (kısa)"
    c7.add_data(Reference(ws, min_col=8, min_row=22, max_row=28), titles_from_data=True)
    c7.set_categories(Reference(ws, min_col=7, min_row=23, max_row=28))
    c7.width, c7.height = 12, 7
    ws.add_chart(c7, "A77")

    c8 = BarChart()
    c8.title = "Olay Adet"
    c8.add_data(Reference(ws, min_col=11, min_row=22, max_row=26), titles_from_data=True)
    c8.set_categories(Reference(ws, min_col=10, min_row=23, max_row=26))
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
    h(ws, 5, 5, "Km Başı", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 6, "Sağlık Skor", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    for i, arac in enumerate(ARACLAR):
        r = 6 + i
        h(ws, r, 1, arac[0])
        h(ws, r, 2,
          f'=IF(OR(COUNTIFS(tblGunluk[VarlikId],A{r},tblGunluk[AlarmBayrak],"ALARM")>0,'
          f'IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[SaglikSkor]),100)<fam_esikSaglik),'
          f'"ALARM","YOK")')
        h(ws, r, 3, f"=IFERROR(ABS(E{r}-fam_esikKmBasi),0)", sayi=TL)
        h(ws, r, 4, "=raporTarihi", sayi=TARİH)
        h(ws, r, 5,
          f'=IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[OrtKmBasi]),0)',
          sayi=TL)
        h(ws, r, 6,
          f'=IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[SaglikSkor]),0)',
          sayi=CATI)

    h(ws, 16, 1, "Alarm Özeti", kalin=True)
    h(ws, 16, 2, "=fam_alarmListesi", kaydir=True)
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
    h(ws, 8, 2, "=fam_firmaUnvan")

    h(ws, 10, 1, "KARAR", kalin=True, boyut=12)
    h(ws, 10, 2, "=fam_kararKapi", kalin=True, boyut=14)
    h(ws, 11, 1, "Km Başı Maliyet")
    h(ws, 11, 2, "=fam_kmBasiMaliyet", sayi=TL)
    h(ws, 12, 1, "Toplam Maliyet")
    h(ws, 12, 2, "=fam_toplamMaliyet", sayi=TL)
    h(ws, 13, 1, "Toplam Km")
    h(ws, 13, 2, "=fam_toplamKm", sayi=CATI)
    h(ws, 14, 1, "Sağlık Skor")
    h(ws, 14, 2, "=fam_saglikSkor", sayi=CATI)
    h(ws, 15, 1, "Alarm Adet")
    h(ws, 15, 2, "=fam_alarmAdet", sayi=CATI)

    h(ws, 17, 1, "Kalem", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 17, 2, "Tutar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, (ad, form) in enumerate([
        ("Yakıt", "=fam_yakitToplam"),
        ("Bakım", "=fam_bakimToplam"),
        ("Amortisman", "=fam_amortToplam"),
        ("Sigorta", "=fam_sigortaToplam"),
    ], 18):
        h(ws, i, 1, ad)
        h(ws, i, 2, form, sayi=TL)

    h(ws, 24, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 1,
      "Yakıt birim fiyatı, amortisman TL/km ve sigorta yıllık tutarı AYARLAR'dan gelir. "
      "Karar destek çıktısıdır; mali tavsiye değildir.",
      kaydir=True)
    ws.merge_cells("A25:F26")

    h(ws, 28, 1, "İmza — Filo Yöneticisi")
    h(ws, 28, 3, "________________")
    h(ws, 29, 1, "İmza — CFO / Denetçi")
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
        ("Dolu kayıt > 0", '=IF(fam_doluKayit>0,"GEÇTİ","KALDI")', "=fam_doluKayit"),
        ("Yetim kayıt = 0", '=IF(fam_yetimAdet=0,"GEÇTİ","UYARI")', "=fam_yetimAdet"),
        ("Km başı hesaplı", '=IF(fam_kmBasiMaliyet>=0,"GEÇTİ","KALDI")', "=fam_kmBasiMaliyet"),
        ("Sağlık skor formül", '=IF(fam_saglikSkor>=0,"GEÇTİ","KALDI")', "=fam_saglikSkor"),
        ("Alarm listesi", '=IF(fam_alarmAdet>=0,"GEÇTİ","KALDI")', "=fam_alarmAdet"),
        ("Karar dolu", '=IF(fam_kararKapi<>"","GEÇTİ","KALDI")', "=fam_kararKapi"),
        ("12 ay sezonsal", '=IF(DONEM_ANALIZ!B22>=0,"GEÇTİ","KALDI")', "=DONEM_ANALIZ!B22"),
        ("Kalite skor", '=IF(fam_kaliteSkor>=0,"GEÇTİ","KALDI")', "=fam_kaliteSkor"),
        ("Tornado", '=IF(fam_tornadoZirve>=0,"GEÇTİ","KALDI")', "=fam_tornadoZirve"),
        ("P90", '=IF(fam_yuzdelikP90>=0,"GEÇTİ","KALDI")', "=fam_yuzdelikP90"),
    ]
    for i, (ad, son, deg) in enumerate(kontroller_list, 6):
        h(ws, i, 1, ad)
        h(ws, i, 2, son)
        h(ws, i, 3, deg)
    genislik(ws, {"A": 28, "B": 12, "C": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Örnek günlük satırlar — kopyalanabilir referans", son_kolon=10)
    basliklar = ["SatirId", "VarlikId", "Tarih", "Km_km", "YakitLt_lt", "BakimTl_tl", "OlayTuru"]
    baslik_satiri(ws, 5, basliklar)
    for i, s in enumerate(ORNEK[:20]):
        r = 6 + i
        h(ws, r, 1, s["id"])
        h(ws, r, 2, s["varlik"])
        h(ws, r, 3, s["tarih"], sayi=TARİH)
        h(ws, r, 4, s["km"], sayi=CATI)
        h(ws, r, 5, s["yakit"], sayi="0.0")
        h(ws, r, 6, s["bakim"], sayi=TL)
        h(ws, r, 7, s["olay"])
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 8)})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Evet / Hayır", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "Secim", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    _sari(ws, 7, 1, "EVET", baslik="Seçim", mesaj="EVET/HAYIR")
    _sari(ws, 8, 1, "HAYIR", baslik="Seçim", mesaj="EVET/HAYIR")

    h(ws, 3, 3, "Olay Türü", kalin=True)
    h(ws, 6, 3, "Tur", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, t in enumerate(OLAY_TURLERI, 7):
        _sari(ws, i, 3, t, baslik="Olay", mesaj="Olay türü")

    h(ws, 3, 5, "Araç Tipi", kalin=True)
    h(ws, 6, 5, "Tip", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, t in enumerate(["Panelvan", "Kamyonet", "Kamyon"], 7):
        _sari(ws, i, 5, t, baslik="Tip", mesaj="Araç tipi")
    genislik(ws, {"A": 12, "C": 14, "E": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + alarm eşikleri (S05)", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("fam_yakitBirimFiyat", 42.5, "TL/lt", "Akaryakıt piyasa / iç tarife", "01.01.2026",
         "Yakıt birim fiyatı TL/lt"),
        ("fam_amortismanTlKm", 1.9, "TL/km", "Filo amortisman politikası", "01.01.2026",
         "Varsayılan amortisman TL/km"),
        ("fam_sigortaYillik", 45000, "TL/yıl", "Sigorta poliçesi", "01.01.2026",
         "Varsayılan yıllık sigorta"),
        ("fam_sigortaGunPay", 1.0, "oran", "Günlük sigorta payı", "01.01.2026",
         "Sigorta gün çarpanı"),
        ("fam_esikKmBasi", 8.5, "TL/km", "Alarm eşiği — km başı", "01.01.2026",
         "Km başı maliyet alarm eşiği"),
        ("fam_esikBakimTl", 1500, "TL", "Alarm eşiği — bakım", "01.01.2026",
         "Tek sefer bakım alarm eşiği"),
        ("fam_esikSaglik", 60, "puan", "Alarm eşiği — sağlık skor", "01.01.2026",
         "Sağlık skor alt eşiği"),
        ("fam_aylikCarpan", 1.0, "oran", "Aylık proxy", "01.01.2026", "Aylık maliyet çarpan"),
        ("fam_senaryoTemkinli", 1.15, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("fam_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("fam_senaryoIyimser", 0.90, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("fam_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("fam_yuzdelikOran", 1.15, "oran", "İstatistik kuralı", "01.01.2026", "P90 çarpan proxy"),
        ("fam_percentileOran", 0.9, "oran", "PERCENTILE oranı", "01.01.2026", "Yüzdelik dilim"),
        ("fam_tahminCarpan", 1.05, "oran", "Tahmin modeli", "01.01.2026", "Tahmin aralık çarpan"),
        ("fam_olcekHedef", 20000, "satir", "Manda A5 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("fam_kaliteTaban", 100, "puan", "Kalite / sağlık modeli", "01.01.2026", "Taban puan"),
        ("fam_kaliteCeza", 12, "puan", "Kalite / sağlık modeli", "01.01.2026", "Eşik aşım cezası"),
        ("fam_kaliteSapma", 3, "puan", "Kalite / sağlık modeli", "01.01.2026", "Sapma / alarm cezası"),
        ("fam_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("fam_firmaUnvan", "Örnek Filo Lojistik A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma unvanı"),
        ("fam_alarmEsikMetin", "km başı / bakım / sağlık eşikleri", "metin",
         "S05 alarm", "01.01.2026", "Alarm eşiği tanımı"),
    ]
    for i, (ana, deg, bir, kay, yur, acik) in enumerate(params):
        r = 6 + i
        h(ws, r, 1, ana)
        if isinstance(deg, date):
            _sari(ws, r, 2, deg, sayi=TARİH, baslik=ana, mesaj=acik)
        elif isinstance(deg, float) and deg <= 1.5:
            _sari(ws, r, 2, deg, sayi=YÜZDE, baslik=ana, mesaj=acik)
        elif isinstance(deg, (int, float)):
            _sari(ws, r, 2, deg, sayi=(TL if "TL" in bir or "tl" in bir.lower() else CATI),
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
        (32, "fam_doluKayit", "=HESAP!C7"),
        (33, "fam_toplamKm", "=HESAP!C8"),
        (34, "fam_toplamMaliyet", "=HESAP!C14"),
        (35, "fam_kmBasiMaliyet", "=HESAP!C15"),
        (36, "fam_yakitToplam", "=HESAP!C11"),
        (37, "fam_bakimToplam", "=HESAP!C10"),
        (38, "fam_amortToplam", "=HESAP!C12"),
        (39, "fam_sigortaToplam", "=HESAP!C13"),
        (40, "fam_saglikSkor", "=HESAP!C20"),
        (41, "fam_alarmAdet", "=HESAP!C18"),
        (42, "fam_yetimAdet", "=HESAP!C19"),
        (43, "fam_gunlukKmOrt", "=HESAP!C25"),
        (44, "fam_aylikMaliyetTrend", "=HESAP!C26"),
        (45, "fam_kmBasiSapma", "=HESAP!C27"),
        (46, "fam_tornadoZirve", "=HESAP!C39"),
        (47, "fam_senaryoKarsilastirma", "=HESAP!C35"),
        (48, "fam_hhi", "=HESAP!C40"),
        (49, "fam_kaliteSkor", "=HESAP!C41"),
        (50, "fam_senaryoMotor", "=HESAP!C44"),
        (51, "fam_tahminAralik", "=HESAP!C42"),
        (52, "fam_yuzdelikP90", "=HESAP!C43"),
        (53, "fam_sezonsalTrend",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"12_AY",TEXT(HESAP!C45,"0"),"km",'
         'TEXT(HESAP!C46,"0"),"TL sezonsal")'),
        (54, "fam_alarmListesi",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Alarm",HESAP!C18,"Yetim",HESAP!C19,'
         '"KmAşım",HESAP!C22,"BakımAşım",HESAP!C23)'),
        (55, "fam_kararKapi",
         '=IF(HESAP!C7=0,"VERİ YOK",'
         'IF(OR(HESAP!C8<0,HESAP!C14<0,HESAP!C10<0,HESAP!C11<0),"KRİTİK",'
         'IF(HESAP!C47=1,"KRİTİK",'
         'IF(OR(HESAP!C22>0,HESAP!C15>fam_esikKmBasi),"DİKKAT","UYGUN"))))'),
        (56, "fam_kararGerekce",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Karar",fam_kararKapi,"| km başı",'
         'TEXT(fam_kmBasiMaliyet,"0.00"),"| sağlık",TEXT(fam_saglikSkor,"0"),'
         '"| alarm",fam_alarmAdet)'),
        (57, "fam_sonrakiKimlik",
         '="ARC-"&TEXT(COUNTA(tblVarliklar[VarlikId])+1,"00000")'),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 59, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (60, "fam_yorumGunluk",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Günlük km ort",TEXT(fam_gunlukKmOrt,"0.0"),"km")'),
        (61, "fam_yorumAylik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Aylık maliyet proxy",TEXT(fam_aylikMaliyetTrend,"0"),"TL")'),
        (62, "fam_yorumSapma",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Km başı sapma",TEXT(fam_kmBasiSapma,"0.00"),"TL")'),
        (63, "fam_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(fam_tornadoZirve,"0"),"TL")'),
        (64, "fam_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(fam_senaryoKarsilastirma,"0"))'),
        (65, "fam_yorumHhi",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma HHI",TEXT(fam_hhi,"0.0%"))'),
        (66, "fam_yorumKalite",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite skoru",fam_kaliteSkor)'),
        (67, "fam_yorumSenMotor",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motoru",fam_senaryoMotor)'),
        (68, "fam_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(fam_tahminAralik,"0"),"TL")'),
        (69, "fam_yorumP90",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 maliyet",TEXT(fam_yuzdelikP90,"0"),"TL")'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 24, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. VARLIKLAR sayfasına araç kartı ekleyin; sonraki kimliği kullanın.",
        "2. GUNLUK_GIRIS sayfasına günlük satır yazın: varlık, tarih, km, yakıt lt, bakım TL, olay (≤6 alan).",
        "3. Km başı maliyet = (yakıt + bakım + amortisman + sigorta) / km formülüyle hesaplanır.",
        "4. DONEM_ANALIZ 12 aylık sezonsal zinciri gösterir.",
        "5. ALARMLAR sayfasında varlık + eşik + tarih listesi; eşikler AYARLAR'dadır.",
        "6. raporTarihi AYARLAR'dadır; TODAY kullanılmaz.",
        "7. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı hücreler açıktır.",
        "8. RAPOR sayfasını PDF olarak yazdırabilirsiniz.",
        f"9. Sürüm {SURUM} | Kod FAM-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
    ]
    for i, m in enumerate(maddeler, 5):
        h(ws, i, 1, m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    genislik(ws, {"A": 90})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 28):
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
            f"E{r}", CellIsRule(operator="greaterThan", formula=["fam_esikKmBasi"], fill=sari))
        ws.conditional_formatting.add(
            f"F{r}", CellIsRule(operator="lessThan", formula=["fam_esikSaglik"], fill=kirmizi))

    ws = wb["GUNLUK_GIRIS"]
    ws.conditional_formatting.add(
        f"N{ILK}:N{SON}",
        CellIsRule(operator="equal", formula=['"ALARM"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"P{ILK}:P{SON}",
        CellIsRule(operator="equal", formula=['"YETIM"'], fill=sari))
    ws.conditional_formatting.add(
        f"L{ILK}:L{SON}",
        CellIsRule(operator="greaterThan", formula=["fam_esikKmBasi"], fill=sari))
    ws.conditional_formatting.add(
        f"M{ILK}:M{SON}",
        CellIsRule(operator="lessThan", formula=["fam_esikSaglik"], fill=kirmizi))

    ws = wb["VARLIKLAR"]
    ws.conditional_formatting.add(
        f"H{ILK}:H{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"L{ILK}:L{SON}",
        CellIsRule(operator="lessThan", formula=["fam_esikSaglik"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"K{ILK}:K{SON}",
        CellIsRule(operator="greaterThan", formula=["fam_esikKmBasi"], fill=sari))

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
    for r in range(6, 52):
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi))

    ws = wb["DONEM_ANALIZ"]
    for r in range(6, 18):
        ws.conditional_formatting.add(
            f"D{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"E{r}", CellIsRule(operator="greaterThan", formula=["fam_esikKmBasi"], fill=sari))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "fam_yakitBirimFiyat", "fam_amortismanTlKm", "fam_sigortaYillik",
        "fam_sigortaGunPay", "fam_esikKmBasi", "fam_esikBakimTl", "fam_esikSaglik",
        "fam_aylikCarpan", "fam_senaryoTemkinli", "fam_senaryoBaz", "fam_senaryoIyimser",
        "fam_tornadoOran", "fam_yuzdelikOran", "fam_percentileOran", "fam_tahminCarpan",
        "fam_olcekHedef",
        "fam_kaliteTaban", "fam_kaliteCeza", "fam_kaliteSapma", "fam_dosyaSurumu",
        "fam_firmaUnvan", "fam_alarmEsikMetin",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "fam_doluKayit": 32, "fam_toplamKm": 33, "fam_toplamMaliyet": 34,
        "fam_kmBasiMaliyet": 35, "fam_yakitToplam": 36, "fam_bakimToplam": 37,
        "fam_amortToplam": 38, "fam_sigortaToplam": 39, "fam_saglikSkor": 40,
        "fam_alarmAdet": 41, "fam_yetimAdet": 42, "fam_gunlukKmOrt": 43,
        "fam_aylikMaliyetTrend": 44, "fam_kmBasiSapma": 45, "fam_tornadoZirve": 46,
        "fam_senaryoKarsilastirma": 47, "fam_hhi": 48, "fam_kaliteSkor": 49,
        "fam_senaryoMotor": 50, "fam_tahminAralik": 51, "fam_yuzdelikP90": 52,
        "fam_sezonsalTrend": 53, "fam_alarmListesi": 54, "fam_kararKapi": 55,
        "fam_kararGerekce": 56, "fam_sonrakiKimlik": 57,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "fam_yorumGunluk": 60, "fam_yorumAylik": 61, "fam_yorumSapma": 62,
        "fam_yorumTornado": 63, "fam_yorumSenaryo": 64, "fam_yorumHhi": 65,
        "fam_yorumKalite": 66, "fam_yorumSenMotor": 67, "fam_yorumTahmin": 68,
        "fam_yorumP90": 69,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$A$7:$A$8")
    ad_ekle(wb, "ListeOlayTuru", "LISTELER!$C$7:$C$12")
    ad_ekle(wb, "ListeAracTipi", "LISTELER!$E$7:$E$9")


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
    dosya_ad = "FiloAracMaliyetKomutasi.xlsx"
    hedef = cikti_yolu or os.path.join(urun_dir, dosya_ad)
    os.makedirs(os.path.dirname(hedef) or ".", exist_ok=True)
    wb.save(hedef)

    cikti = os.path.join(KOK, "cikti", dosya_ad)
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(hedef) != os.path.abspath(cikti):
        import shutil
        shutil.copy2(hedef, cikti)

    # ornek_veri.csv
    csv_yol = os.path.join(urun_dir, "ornek_veri.csv")
    with open(csv_yol, "w", encoding="utf-8") as f:
        f.write("SatirId,VarlikId,Tarih,Km_km,YakitLt_lt,BakimTl_tl,OlayTuru\n")
        for s in ORNEK:
            f.write(f"{s['id']},{s['varlik']},{s['tarih']:%d.%m.%Y},"
                    f"{s['km']},{s['yakit']},{s['bakim']},{s['olay']}\n")

    print(f"Dosya oluşturuldu: {hedef}")
    print(f"Kopya: {cikti}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
