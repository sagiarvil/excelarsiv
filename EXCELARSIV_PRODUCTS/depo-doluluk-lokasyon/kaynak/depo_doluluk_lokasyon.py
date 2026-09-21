#!/usr/bin/env python3
"""Depo Doluluk & Lokasyon Optimizasyonu — A5 saha komuta merkezi (manda v6).

Doluluk % + koridor/lokasyon verim + alarm kuyruğu.
GUNLUK_GIRIS ≤6 manuel alan; birimli kolonlar; ≥12 ay sezonsallık;
lokasyon sağlık skoru formülden; AYARLAR alarm eşikleri + alarm listesi.
Domain: depo — filo/km maliyet kopyası değildir.
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

URUN_AD = "Depo Doluluk & Lokasyon Optimizasyonu"
SURUM = "1.0.0"
RENK = "0F2742"
KAPASITE = 250
DEMO_LOK = 8
DEMO_GUNLUK = 36
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

# VarlikId, KoridorAd, Bolge, Tip, KapasiteM3_m3, KapasitePalet_adet, HedefDolulukPct
LOKASYONLAR = [
    ("LOC-00001", "K-A01", "Kuzey", "Raf", 140, 90, 0.85),
    ("LOC-00002", "K-A02", "Kuzey", "Raf", 120, 80, 0.85),
    ("LOC-00003", "K-B01", "Guney", "Zemin", 200, 120, 0.80),
    ("LOC-00004", "K-B02", "Guney", "Raf", 110, 70, 0.85),
    ("LOC-00005", "K-C01", "Dogu", "Mezanin", 90, 55, 0.80),
    ("LOC-00006", "K-C02", "Dogu", "Raf", 130, 85, 0.85),
    ("LOC-00007", "K-D01", "Bati", "Zemin", 180, 110, 0.75),
    ("LOC-00008", "K-D02", "Bati", "Raf", 100, 65, 0.85),
]

OLAY_TURLERI = ["SAYIM", "MAL_KABUL", "TOPLAMA", "SEVK", "STOK_DUZELTME", "DIGER"]
AYLAR = [
    "Ocak", "Subat", "Mart", "Nisan", "Mayis", "Haziran",
    "Temmuz", "Agustos", "Eylul", "Ekim", "Kasim", "Aralik",
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Depo doluluk ve lokasyon verimini etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_gunluk():
    satirlar = []
    for i in range(1, DEMO_GUNLUK + 1):
        lok = LOKASYONLAR[(i - 1) % DEMO_LOK]
        tar = RAPOR_TARIHI - timedelta(days=(i * 3) % 340)
        doluluk = round(0.55 + (i % 7) * 0.05 + (i % 3) * 0.03, 2)
        if doluluk > 0.98:
            doluluk = 0.98
        palet = int(lok[5] * doluluk) + (i % 4)
        pick = 20 + (i % 9) * 8 + (i % 5) * 3
        satirlar.append({
            "id": f"GLG-{i:05d}",
            "varlik": "" if i == 30 else (lok[0] if i != 33 else "LOC-99999"),
            "tarih": tar,
            "doluluk": doluluk,
            "palet": palet,
            "pick": pick,
            "olay": OLAY_TURLERI[i % len(OLAY_TURLERI)],
        })
    return satirlar


ORNEK = _demo_gunluk()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Depo lokasyonlarında doluluk yüzdesini ve koridor verimini izler; "
      "günlük saha girişi ≤6 alan; lokasyon sağlık skoru ve alarm kuyruğu üretir; "
      "UYGUN / DİKKAT / KRİTİK / VERİ YOK kararı verir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Lokasyon bazlı doluluk % ve koridor verim skoru",
        "Günlük saha girişi en fazla 6 alan — 30 saniye hedefi",
        "Lokasyon sağlık skoru formülle veriden türer (sabit skor yok)",
        "12 aylık sezonsal doluluk zinciri ve alarm eşiği kuyruğu",
        "Makrosuz, çevrimdışı; WMS/API bağı yok",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Saha / depo operatörü — günlük doluluk-palet-pick satırı",
        "Depo müdürü — doluluk % ve karar panosu",
        "CFO / denetçi — dönemsel RAPOR kanıt çıktısı",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) VARLIKLAR'a lokasyon kartı girin. 2) GUNLUK_GIRIS'e günlük satır yazın. "
      "3) DONEM_ANALIZ ve PANO'dan kararı görün. 4) RAPOR'dan çıktı alın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Kod: DDL-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def varliklar(ws):
    sayfa_hazirla(ws, "VARLIKLAR", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Lokasyon kartları — her varlığın tek kimliği vardır", son_kolon=14)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=ddl_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4",
               "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: LOC-00009 | Risk: Elle çakışma")

    sutunlar = [
        "VarlikId", "KoridorAd", "Bolge", "Tip",
        "KapasiteM3_m3", "KapasitePalet_adet", "HedefDolulukPct_pct", "Aktif",
        "OrtDoluluk", "OrtPick", "OrtVerim", "SaglikSkor",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "OrtDoluluk": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'IFERROR(AVERAGEIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[DolulukPct_pct]),0))'
        ),
        "OrtPick": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'IFERROR(AVERAGEIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[PickAdet_adet]),0))'
        ),
        "OrtVerim": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'IFERROR(AVERAGEIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[FormulVerim_pct]),0))'
        ),
        "SaglikSkor": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'MAX(0,MIN(100,ddl_kaliteTaban'
            '-IF(tblVarliklar[[#This Row],[OrtDoluluk]]>ddl_esikDoluluk,ddl_kaliteCeza,0)'
            '-IF(tblVarliklar[[#This Row],[OrtVerim]]<ddl_esikVerim,ddl_kaliteCeza,0)'
            '-COUNTIFS(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[AlarmBayrak],"ALARM")*ddl_kaliteSapma'
            '-IF(tblVarliklar[[#This Row],[Aktif]]="HAYIR",ddl_kaliteCeza,0))))'
        ),
    }
    for i, row in enumerate(LOKASYONLAR):
        r = ILK + i
        kid, kor, bolge, tip, m3, pal, hedef = row
        _sari(ws, r, 1, kid, baslik="Varlık Kimliği", mesaj="LOC-##### formatında")
        _sari(ws, r, 2, kor, baslik="Koridor", mesaj="Koridor kodu")
        _sari(ws, r, 3, bolge, baslik="Bölge", mesaj="Depo bölgesi")
        _sari(ws, r, 4, tip, baslik="Tip", mesaj="Raf/Zemin/Mezanin")
        _sari(ws, r, 5, m3, sayi=CATI, baslik="Kapasite [m3]", mesaj="Metreküp")
        _sari(ws, r, 6, pal, sayi=CATI, baslik="Kapasite [palet]", mesaj="Palet adet")
        _sari(ws, r, 7, hedef, sayi=YÜZDE, baslik="Hedef Doluluk [%]", mesaj="0-1 arası")
        _sari(ws, r, 8, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_LOK, SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None,
                  sayi=(YÜZDE if c == 7 else (CATI if c in (5, 6) else None)),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVarliklar", f"A{HDR}:L{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Varlık Kimliği", mesaj="LOC-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Koridor"), (3, "Bölge"), (4, "Tip")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    for c, ad in [(5, "Kapasite m3"), (6, "Kapasite palet")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"G{ILK}:G{SON}",
              baslik="Hedef Doluluk", mesaj="0 ile 1 arası",
              hata_baslik="Doluluk", hata_mesaj="0-1 aralığı",
              isaret="between", f2="1")
    dogrulama(ws, "list", "ListeEvetHayir", f"H{ILK}:H{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 13)})
    ws.column_dimensions["B"].width = 12


def gunluk_giris(ws):
    sayfa_hazirla(ws, "GUNLUK_GIRIS", RENK, URUN_AD, son_kolon=22)
    alt_bant(ws, 3,
             "GÜNLÜK saha girişi — satır başına en fazla 6 manuel alan (S02)",
             son_kolon=16)
    # Manuel: SatirId, VarlikId, Tarih, DolulukPct_pct, PaletAdet_adet, PickAdet_adet, OlayTuru
    # S02 filtresi id/tarih/formul/skor/alarm düşer → Doluluk, Palet, Pick, Olay ≤6
    sutunlar = [
        "SatirId", "VarlikId", "Tarih", "DolulukPct_pct", "PaletAdet_adet", "PickAdet_adet",
        "OlayTuru",
        "FormulKapasite_pct", "FormulVerim_pct", "FormulBosPalet_adet",
        "SaglikSkor_puan", "AlarmBayrak", "FormulKayitDolu_adet", "FormulYetim",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "FormulKapasite_pct": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR(tblGunluk[[#This Row],[PaletAdet_adet]]/'
            'IFERROR(SUMIF(tblVarliklar[VarlikId],tblGunluk[[#This Row],[VarlikId]],'
            'tblVarliklar[KapasitePalet_adet]),ddl_varsayilanKapasitePalet),0))'
        ),
        "FormulVerim_pct": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR(tblGunluk[[#This Row],[PickAdet_adet]]/'
            'MAX(tblGunluk[[#This Row],[PaletAdet_adet]],1)*ddl_verimCarpan,0))'
        ),
        "FormulBosPalet_adet": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR(IFERROR(SUMIF(tblVarliklar[VarlikId],tblGunluk[[#This Row],[VarlikId]],'
            'tblVarliklar[KapasitePalet_adet]),ddl_varsayilanKapasitePalet)'
            '-tblGunluk[[#This Row],[PaletAdet_adet]],0))'
        ),
        "SaglikSkor_puan": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'MAX(0,MIN(100,ddl_kaliteTaban'
            '-IF(tblGunluk[[#This Row],[DolulukPct_pct]]>ddl_esikDoluluk,ddl_kaliteCeza,0)'
            '-IF(tblGunluk[[#This Row],[FormulVerim_pct]]<ddl_esikVerim,ddl_kaliteSapma,0)'
            '-IF(tblGunluk[[#This Row],[FormulYetim]]="YETIM",ddl_kaliteCeza,0))))'
        ),
        "AlarmBayrak": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IF(OR(tblGunluk[[#This Row],[DolulukPct_pct]]>ddl_esikDoluluk,'
            'tblGunluk[[#This Row],[FormulVerim_pct]]<ddl_esikVerim,'
            'tblGunluk[[#This Row],[SaglikSkor_puan]]<ddl_esikSaglik),'
            '"ALARM","UYGUN"))'
        ),
        "FormulKayitDolu_adet": (
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
        _sari(ws, r, 2, s["varlik"], baslik="Varlık Kimliği", mesaj="LOC listesinden")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="GG.AA.YYYY")
        _sari(ws, r, 4, s["doluluk"], sayi=YÜZDE, baslik="Doluluk [%]", mesaj="0-1 arası")
        _sari(ws, r, 5, s["palet"], sayi=CATI, baslik="Palet [adet]", mesaj="Palet adedi")
        _sari(ws, r, 6, s["pick"], sayi=CATI, baslik="Pick [adet]", mesaj="Toplama adedi")
        _sari(ws, r, 7, s["olay"], baslik="Olay Türü", mesaj="Listeden seçin")
    for r in range(ILK + DEMO_GUNLUK, SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None,
                  sayi=(TARİH if c == 3 else (YÜZDE if c == 4 else (CATI if c in (5, 6) else None))),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblGunluk", f"A{HDR}:N{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Günlük Kimlik", mesaj="GLG-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Varlık", mesaj="LOC kimliği girin",
              hata_baslik="Varlık", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"C{ILK}:C{SON}",
              baslik="Tarih", mesaj="GG.AA.YYYY",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "decimal", "0", f"D{ILK}:D{SON}",
              baslik="Doluluk [%]", mesaj="0 ile 1 arası",
              hata_baslik="Doluluk", hata_mesaj="0-1 aralığı",
              isaret="between", f2="1")
    for c, ad in [(5, "Palet [adet]"), (6, "Pick [adet]")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeOlayTuru", f"G{ILK}:G{SON}",
              baslik="Olay Türü", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 13 for i in range(1, 15)})


def donem_analiz(ws):
    sayfa_hazirla(ws, "DONEM_ANALIZ", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "12 aylık sezonsallık — doluluk ve pick zinciri (S03)", son_kolon=12)
    h(ws, 5, 1, "Ay", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Ay_Anahtari", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Ort Doluluk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 4, "Toplam Pick", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 5, "Ort Verim", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 6, "Sezonsal Etiket", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    demo_dol = [0.62, 0.65, 0.70, 0.74, 0.78, 0.82, 0.88, 0.86, 0.80, 0.75, 0.68, 0.64]
    demo_pick = [4200, 4400, 4800, 5100, 5400, 5700, 6200, 6000, 5500, 5000, 4600, 4300]
    for i, ay in enumerate(AYLAR, 1):
        r = 5 + i
        h(ws, r, 1, ay)
        h(ws, r, 2, f"Ay_{i}")
        h(ws, r, 3, demo_dol[i - 1], sayi=YÜZDE)
        h(ws, r, 4, demo_pick[i - 1], sayi=CATI)
        h(ws, r, 5, f"=IFERROR(D{r}/MAX(C{r}*1000,1)*ddl_verimCarpan,0)", sayi=YÜZDE)
        h(ws, r, 6, "12_AY_SEZONSAL")

    h(ws, 20, 1, "Sezonsal Trend Özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=ddl_sezonsalTrend", kaydir=True)
    ws.merge_cells("B20:F20")
    h(ws, 22, 1, "12 Ay Ort Doluluk", kalin=True)
    h(ws, 22, 2, "=AVERAGE(C6:C17)", sayi=YÜZDE)
    h(ws, 23, 1, "12 Ay Toplam Pick", kalin=True)
    h(ws, 23, 2, "=SUM(D6:D17)", sayi=CATI)
    h(ws, 24, 1, "12 Ay Ort Verim", kalin=True)
    h(ws, 24, 2, "=AVERAGE(E6:E17)", sayi=YÜZDE)
    h(ws, 25, 1, "Canlı dönem doluluk", kalin=True)
    h(ws, 25, 2, "=ddl_ortDoluluk", sayi=YÜZDE)
    h(ws, 26, 1, "Canlı dönem pick", kalin=True)
    h(ws, 26, 2, "=ddl_toplamPick", sayi=CATI)

    c1 = LineChart()
    c1.title = "12 Ay Doluluk Sezonsallığı"
    c1.add_data(Reference(ws, min_col=3, min_row=5, max_row=17), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17))
    c1.width, c1.height = 14, 8
    ws.add_chart(c1, "A27")

    c2 = BarChart()
    c2.title = "12 Ay Pick Dağılımı"
    c2.add_data(Reference(ws, min_col=4, min_row=5, max_row=17), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17))
    c2.width, c2.height = 14, 8
    ws.add_chart(c2, "H27")

    genislik(ws, {"A": 24, "B": 14, "C": 14, "D": 14, "E": 14, "F": 18})


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Hesap motoru — ≥40 adım; sabitler AYARLAR'dan", son_kolon=10)
    h(ws, 5, 1, "Adım", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Açıklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    # Satır: H01=C6 … indeks = 5+n
    adimlar = [
        ("H01", "Toplam günlük kayıt", '=COUNTA(tblGunluk[SatirId])'),
        ("H02", "Dolu kayıt", '=SUM(tblGunluk[FormulKayitDolu_adet])'),
        ("H03", "Ortalama doluluk", '=IFERROR(AVERAGEIF(tblGunluk[DolulukPct_pct],">0"),0)'),
        ("H04", "Toplam palet", "=SUM(tblGunluk[PaletAdet_adet])"),
        ("H05", "Toplam pick", "=SUM(tblGunluk[PickAdet_adet])"),
        ("H06", "Ortalama verim", '=IFERROR(AVERAGEIF(tblGunluk[FormulVerim_pct],">0"),0)'),
        ("H07", "Ortalama kapasite kullanım",
         '=IFERROR(AVERAGEIF(tblGunluk[FormulKapasite_pct],">0"),0)'),
        ("H08", "Toplam boş palet", "=SUM(tblGunluk[FormulBosPalet_adet])"),
        ("H09", "Aktif lokasyon adet", '=COUNTIF(tblVarliklar[Aktif],"EVET")'),
        ("H10", "Toplam lokasyon kart", '=COUNTA(tblVarliklar[VarlikId])'),
        ("H11", "Alarm adet", '=COUNTIF(tblGunluk[AlarmBayrak],"ALARM")'),
        ("H12", "Yetim kayıt", '=COUNTIF(tblGunluk[FormulYetim],"YETIM")'),
        ("H13", "Ortalama sağlık skor",
         '=IFERROR(AVERAGEIF(tblGunluk[SaglikSkor_puan],">0"),0)'),
        ("H14", "Min sağlık skor",
         '=IFERROR(MIN(tblGunluk[SaglikSkor_puan]),0)'),
        ("H15", "Doluluk eşik aşımı",
         '=COUNTIFS(tblGunluk[DolulukPct_pct],">"&ddl_esikDoluluk,tblGunluk[FormulKayitDolu_adet],">0")'),
        ("H16", "Verim eşik altı",
         '=COUNTIFS(tblGunluk[FormulVerim_pct],"<"&ddl_esikVerim,tblGunluk[FormulKayitDolu_adet],">0")'),
        ("H17", "Düşük sağlık adet",
         '=COUNTIFS(tblGunluk[SaglikSkor_puan],"<"&ddl_esikSaglik,tblGunluk[FormulKayitDolu_adet],">0")'),
        ("H18", "Günlük doluluk ort", "=C8"),
        ("H19", "Aylık doluluk proxy", "=C8*ddl_aylikCarpan"),
        ("H20", "Doluluk sapma", "=IFERROR(ABS(C8-ddl_esikDoluluk),0)"),
        ("H21", "Palet payı proxy", "=IFERROR(C9/MAX(C9+C13,1),0)"),
        ("H22", "Pick payı proxy", "=IFERROR(C10/MAX(C10+C9,1),0)"),
        ("H23", "Kapasite payı", "=IFERROR(C12/MAX(C12+C13,1),0)"),
        ("H24", "Boş palet payı", "=IFERROR(C13/MAX(C9+C13,1),0)"),
        ("H25", "Temkinli senaryo doluluk", "=C8*ddl_senaryoTemkinli"),
        ("H26", "Baz senaryo doluluk", "=C8*ddl_senaryoBaz"),
        ("H27", "İyimser senaryo doluluk", "=C8*ddl_senaryoIyimser"),
        ("H28", "Senaryo farkı", "=C32-C30"),
        ("H29", "Tornado doluluk etki", "=C8*ddl_tornadoOran"),
        ("H30", "Tornado pick etki", "=C10*ddl_tornadoOran"),
        ("H31", "Tornado verim etki", "=C11*ddl_tornadoOran"),
        ("H32", "Tornado zirve", "=MAX(C34,C35,C36)"),
        ("H33", "HHI proxy",
         "=IFERROR((C26)^2+(C27)^2+(C28)^2+(C29)^2,0)"),
        ("H34", "Kalite skor",
         "=MAX(0,ddl_kaliteTaban-C20*ddl_kaliteCeza-C21*ddl_kaliteSapma"
         "-C22*ddl_kaliteSapma)"),
        ("H35", "Tahmin aralık", "=C31*ddl_tahminCarpan"),
        ("H36", "P90 doluluk",
         "=IFERROR(PERCENTILE(tblGunluk[DolulukPct_pct],ddl_percentileOran),C8*ddl_yuzdelikOran)"),
        ("H37", "Senaryo motor metin",
         '=_xlfn.TEXTJOIN("|",TRUE,"T",TEXT(C30,"0.0%"),"B",TEXT(C31,"0.0%"),'
         '"I",TEXT(C32,"0.0%"))'),
        ("H38", "Sezonsal 12 ay pick", "=DONEM_ANALIZ!B23"),
        ("H39", "Sezonsal 12 ay doluluk", "=DONEM_ANALIZ!B22"),
        ("H40", "Kritik sinyal",
         '=IF(OR(C20>0,C22>0,C18<ddl_esikSaglik),1,0)'),
        ("H41", "Karar kodu ham",
         '=IF(C7=0,0,IF(C45=1,3,IF(OR(C20>0,C8>ddl_esikDoluluk),2,1)))'),
        ("H42", "Doluluk çıpa", "=C8"),
        ("H43", "Alarm listesi özet adet", "=C16"),
        ("H44", "Varlık sağlık ort",
         '=IFERROR(AVERAGEIF(tblVarliklar[SaglikSkor],">0"),0)'),
        ("H45", "Tahmin FORECAST",
         "=IFERROR(FORECAST(C7+1,tblGunluk[DolulukPct_pct],tblGunluk[FormulKayitDolu_adet]),C40)"),
        ("H46", "Pick yoğunluk", "=IFERROR(C10/MAX(C7,1),0)"),
        ("H47", "Koridor verim ort", "=C11"),
    ]
    # H01=C6 H02=C7 H03=C8 H04=C9 H05=C10 H06=C11 H07=C12 H08=C13
    # H09=C14 H10=C15 H11=C16 H12=C17 H13=C18 H14=C19 H15=C20 H16=C21 H17=C22
    # H18=C23 H19=C24 H20=C25 H21=C26 H22=C27 H23=C28 H24=C29
    # H25=C30 H26=C31 H27=C32 H28=C33 H29=C34 H30=C35 H31=C36 H32=C37
    # H33=C38 H34=C39 H35=C40 H36=C41 H37=C42 H38=C43 H39=C44 H40=C45
    # H41=C46 H42=C47 H43=C48 H44=C49 H45=C50 H46=C51 H47=C52

    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        sayi = YÜZDE if any(x in acik.lower() for x in (
            "doluluk", "verim", "payı", "hhi", "senaryo", "kapasite kullanım", "p90"
        )) else CATI
        if "metin" in acik.lower() or "karar" in acik.lower() or "sinyal" in acik.lower():
            sayi = None
        if (
            "tornado pick" in acik.lower()
            or "toplam pick" in acik.lower()
            or "pick yoğun" in acik.lower()
        ):
            sayi = CATI
        if "tahmin" in acik.lower() and "aralık" in acik.lower():
            sayi = YÜZDE
        h(ws, r, 3, form, sayi=sayi)

    genislik(ws, {"A": 8, "B": 36, "C": 22})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Karar panosu — tek ekran özet + alarm + trend", son_kolon=14)

    h(ws, 5, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 5, 2, "=ddl_kararKapi", kalin=True, boyut=16)
    yorum_ekle(ws, "B5",
               "Tanım: Karar kapısı | Neden önemli: Boş dosyada VERİ YOK | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle müdahale yok")

    kpis = [
        (7, "Dolu Kayıt", "=ddl_doluKayit", CATI),
        (8, "Ort Doluluk", "=ddl_ortDoluluk", YÜZDE),
        (9, "Toplam Palet", "=ddl_toplamPalet", CATI),
        (10, "Toplam Pick", "=ddl_toplamPick", CATI),
        (11, "Ort Verim", "=ddl_ortVerim", YÜZDE),
        (12, "Kapasite Kullanım", "=ddl_kapasiteKullanim", YÜZDE),
        (13, "Boş Palet", "=ddl_bosPalet", CATI),
        (14, "Aktif Lokasyon", "=ddl_aktifLokasyon", CATI),
        (15, "Sağlık Skor Ort", "=ddl_saglikSkor", CATI),
        (16, "Alarm Adet", "=ddl_alarmAdet", CATI),
        (17, "Yetim Kayıt", "=ddl_yetimAdet", CATI),
        (18, "Kalite Skor", "=ddl_kaliteSkor", CATI),
        (19, "Tornado Zirve", "=ddl_tornadoZirve", CATI),
        (20, "P90 Doluluk", "=ddl_yuzdelikP90", YÜZDE),
    ]
    h(ws, 6, 1, "Gösterge", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 6, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for r, ad, form, sayi in kpis:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=sayi, kalin=True)

    h(ws, 7, 4, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 5, "=ddl_kararGerekce", kaydir=True)
    ws.merge_cells("E7:H8")
    h(ws, 10, 4, "Öneri 1", kalin=True)
    h(ws, 10, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Ort doluluk",TEXT(ddl_ortDoluluk,"0.0%"),'
                 '"— eşik",TEXT(ddl_esikDoluluk,"0.0%"))', kaydir=True)
    h(ws, 11, 4, "Öneri 2", kalin=True)
    h(ws, 11, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Alarm",ddl_alarmAdet,"adet — sağlık",'
                 'TEXT(ddl_saglikSkor,"0"))', kaydir=True)
    h(ws, 12, 4, "Öneri 3", kalin=True)
    h(ws, 12, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Verim",TEXT(ddl_ortVerim,"0.0%"),'
                 '"— koridor densitesi izle")', kaydir=True)

    h(ws, 22, 1, "Bolge")
    h(ws, 22, 2, "Doluluk")
    for i, (ad, t) in enumerate([
        ("Kuzey", 0.78), ("Guney", 0.72), ("Dogu", 0.81), ("Bati", 0.68),
    ], 23):
        h(ws, i, 1, ad)
        h(ws, i, 2, t, sayi=YÜZDE)

    h(ws, 22, 4, "Senaryo")
    h(ws, 22, 5, "Doluluk")
    for i, (ad, t) in enumerate([
        ("Temkinli", 0.92), ("Baz", 0.78), ("İyimser", 0.68),
    ], 23):
        h(ws, i, 4, ad)
        h(ws, i, 5, t, sayi=YÜZDE)

    h(ws, 22, 7, "Hafta")
    h(ws, 22, 8, "Pick")
    for i in range(1, 9):
        h(ws, 22 + i, 7, f"H{i}")
        h(ws, 22 + i, 8, 900 + i * 70, sayi=CATI)

    h(ws, 22, 10, "Olay")
    h(ws, 22, 11, "Adet")
    for i, (t, n) in enumerate([
        ("SAYIM", 10), ("MAL_KABUL", 8), ("TOPLAMA", 12), ("SEVK", 6),
    ], 23):
        h(ws, i, 10, t)
        h(ws, i, 11, n, sayi=CATI)

    c1 = PieChart()
    c1.title = "Bölge Doluluk Dağılımı"
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
    c3.title = "Haftalık Pick Trend"
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
    c5.title = "Bölge Payı (tekrar)"
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
    c7.title = "Pick Trend (kısa)"
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
    h(ws, 5, 5, "Ort Doluluk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 6, "Sağlık Skor", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    for i, lok in enumerate(LOKASYONLAR):
        r = 6 + i
        h(ws, r, 1, lok[0])
        h(ws, r, 2,
          f'=IF(OR(COUNTIFS(tblGunluk[VarlikId],A{r},tblGunluk[AlarmBayrak],"ALARM")>0,'
          f'IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[SaglikSkor]),100)<ddl_esikSaglik),'
          f'"ALARM","YOK")')
        h(ws, r, 3, f"=IFERROR(ABS(E{r}-ddl_esikDoluluk),0)", sayi=YÜZDE)
        h(ws, r, 4, "=raporTarihi", sayi=TARİH)
        h(ws, r, 5,
          f'=IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[OrtDoluluk]),0)',
          sayi=YÜZDE)
        h(ws, r, 6,
          f'=IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[SaglikSkor]),0)',
          sayi=CATI)

    h(ws, 16, 1, "Alarm Özeti", kalin=True)
    h(ws, 16, 2, "=ddl_alarmListesi", kaydir=True)
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
    h(ws, 8, 2, "=ddl_firmaUnvan")

    h(ws, 10, 1, "KARAR", kalin=True, boyut=12)
    h(ws, 10, 2, "=ddl_kararKapi", kalin=True, boyut=14)
    h(ws, 11, 1, "Ort Doluluk")
    h(ws, 11, 2, "=ddl_ortDoluluk", sayi=YÜZDE)
    h(ws, 12, 1, "Ort Verim")
    h(ws, 12, 2, "=ddl_ortVerim", sayi=YÜZDE)
    h(ws, 13, 1, "Toplam Pick")
    h(ws, 13, 2, "=ddl_toplamPick", sayi=CATI)
    h(ws, 14, 1, "Sağlık Skor")
    h(ws, 14, 2, "=ddl_saglikSkor", sayi=CATI)
    h(ws, 15, 1, "Alarm Adet")
    h(ws, 15, 2, "=ddl_alarmAdet", sayi=CATI)

    h(ws, 17, 1, "Kalem", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 17, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, (ad, form, sayi) in enumerate([
        ("Toplam Palet", "=ddl_toplamPalet", CATI),
        ("Boş Palet", "=ddl_bosPalet", CATI),
        ("Kapasite Kullanım", "=ddl_kapasiteKullanim", YÜZDE),
        ("Aktif Lokasyon", "=ddl_aktifLokasyon", CATI),
    ], 18):
        h(ws, i, 1, ad)
        h(ws, i, 2, form, sayi=sayi)

    h(ws, 24, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 1,
      "Doluluk ve verim eşikleri AYARLAR'dan gelir. "
      "Karar destek çıktısıdır; operasyonel tavsiye değildir.",
      kaydir=True)
    ws.merge_cells("A25:F26")

    h(ws, 28, 1, "İmza — Depo Müdürü")
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
        ("Dolu kayıt > 0", '=IF(ddl_doluKayit>0,"GEÇTİ","KALDI")', "=ddl_doluKayit"),
        ("Yetim kayıt = 0", '=IF(ddl_yetimAdet=0,"GEÇTİ","UYARI")', "=ddl_yetimAdet"),
        ("Doluluk hesaplı", '=IF(ddl_ortDoluluk>=0,"GEÇTİ","KALDI")', "=ddl_ortDoluluk"),
        ("Sağlık skor formül", '=IF(ddl_saglikSkor>=0,"GEÇTİ","KALDI")', "=ddl_saglikSkor"),
        ("Alarm listesi", '=IF(ddl_alarmAdet>=0,"GEÇTİ","KALDI")', "=ddl_alarmAdet"),
        ("Karar dolu", '=IF(ddl_kararKapi<>"","GEÇTİ","KALDI")', "=ddl_kararKapi"),
        ("12 ay sezonsal", '=IF(DONEM_ANALIZ!B23>=0,"GEÇTİ","KALDI")', "=DONEM_ANALIZ!B23"),
        ("Kalite skor", '=IF(ddl_kaliteSkor>=0,"GEÇTİ","KALDI")', "=ddl_kaliteSkor"),
        ("Tornado", '=IF(ddl_tornadoZirve>=0,"GEÇTİ","KALDI")', "=ddl_tornadoZirve"),
        ("P90", '=IF(ddl_yuzdelikP90>=0,"GEÇTİ","KALDI")', "=ddl_yuzdelikP90"),
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
        "SatirId", "VarlikId", "Tarih", "DolulukPct_pct", "PaletAdet_adet",
        "PickAdet_adet", "OlayTuru",
    ]
    baslik_satiri(ws, 5, basliklar)
    for i, s in enumerate(ORNEK[:20]):
        r = 6 + i
        h(ws, r, 1, s["id"])
        h(ws, r, 2, s["varlik"])
        h(ws, r, 3, s["tarih"], sayi=TARİH)
        h(ws, r, 4, s["doluluk"], sayi=YÜZDE)
        h(ws, r, 5, s["palet"], sayi=CATI)
        h(ws, r, 6, s["pick"], sayi=CATI)
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

    h(ws, 3, 5, "Lokasyon Tipi", kalin=True)
    h(ws, 6, 5, "Tip", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, t in enumerate(["Raf", "Zemin", "Mezanin"], 7):
        _sari(ws, i, 5, t, baslik="Tip", mesaj="Lokasyon tipi")
    genislik(ws, {"A": 12, "C": 16, "E": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + alarm eşikleri (S05)", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("ddl_esikDoluluk", 0.85, "oran", "Alarm eşiği — doluluk", "01.01.2026",
         "Doluluk üst alarm eşiği"),
        ("ddl_esikVerim", 0.70, "oran", "Alarm eşiği — verim", "01.01.2026",
         "Koridor verim alt eşiği"),
        ("ddl_esikSaglik", 60, "puan", "Alarm eşiği — sağlık skor", "01.01.2026",
         "Sağlık skor alt eşiği"),
        ("ddl_varsayilanKapasitePalet", 80, "adet", "Varsayılan kapasite", "01.01.2026",
         "Lokasyon yoksa palet kapasitesi"),
        ("ddl_verimCarpan", 1.0, "oran", "Verim modeli", "01.01.2026",
         "Pick/palet verim çarpanı"),
        ("ddl_aylikCarpan", 1.0, "oran", "Aylık proxy", "01.01.2026", "Aylık doluluk çarpan"),
        ("ddl_senaryoTemkinli", 1.15, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("ddl_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("ddl_senaryoIyimser", 0.90, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("ddl_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("ddl_yuzdelikOran", 1.05, "oran", "İstatistik kuralı", "01.01.2026", "P90 çarpan proxy"),
        ("ddl_percentileOran", 0.9, "oran", "PERCENTILE oranı", "01.01.2026", "Yüzdelik dilim"),
        ("ddl_tahminCarpan", 1.05, "oran", "Tahmin modeli", "01.01.2026", "Tahmin aralık çarpan"),
        ("ddl_olcekHedef", 10000, "satir", "Manda A5 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("ddl_kaliteTaban", 100, "puan", "Kalite / sağlık modeli", "01.01.2026", "Taban puan"),
        ("ddl_kaliteCeza", 12, "puan", "Kalite / sağlık modeli", "01.01.2026", "Eşik aşım cezası"),
        ("ddl_kaliteSapma", 3, "puan", "Kalite / sağlık modeli", "01.01.2026", "Sapma / alarm cezası"),
        ("ddl_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("ddl_firmaUnvan", "Örnek Depo Lojistik A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma unvanı"),
        ("ddl_alarmEsikMetin", "doluluk / verim / sağlık eşikleri", "metin",
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
            _sari(ws, r, 2, deg, sayi=CATI, baslik=ana, mesaj=acik)
        else:
            _sari(ws, r, 2, deg, baslik=ana, mesaj=acik)
        h(ws, r, 3, bir)
        h(ws, r, 4, kay)
        h(ws, r, 5, yur)
        h(ws, r, 6, acik, kaydir=True)

    h(ws, 30, 8, "Motor Çıktıları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 31, 8, "anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 31, 9, "deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    # HESAP satır: H01=C6 … H02=C7 H03=C8 H04=C9 H05=C10 H06=C11 H07=C12 H08=C13
    # H09=C14 H11=C16 H12=C17 H13=C18 H15=C20 H18=C23 H19=C24 H20=C25
    # H28=C33 H32=C37 H33=C38 H34=C39 H35=C40 H36=C41 H37=C42 H38=C43 H39=C44 H40=C45
    motor = [
        (32, "ddl_doluKayit", "=HESAP!C7"),
        (33, "ddl_ortDoluluk", "=HESAP!C8"),
        (34, "ddl_toplamPalet", "=HESAP!C9"),
        (35, "ddl_toplamPick", "=HESAP!C10"),
        (36, "ddl_ortVerim", "=HESAP!C11"),
        (37, "ddl_kapasiteKullanim", "=HESAP!C12"),
        (38, "ddl_bosPalet", "=HESAP!C13"),
        (39, "ddl_aktifLokasyon", "=HESAP!C14"),
        (40, "ddl_saglikSkor", "=HESAP!C18"),
        (41, "ddl_alarmAdet", "=HESAP!C16"),
        (42, "ddl_yetimAdet", "=HESAP!C17"),
        (43, "ddl_gunlukDolulukOrt", "=HESAP!C23"),
        (44, "ddl_aylikDolulukTrend", "=HESAP!C24"),
        (45, "ddl_dolulukSapma", "=HESAP!C25"),
        (46, "ddl_tornadoZirve", "=HESAP!C37"),
        (47, "ddl_senaryoKarsilastirma", "=HESAP!C33"),
        (48, "ddl_hhi", "=HESAP!C38"),
        (49, "ddl_kaliteSkor", "=HESAP!C39"),
        (50, "ddl_senaryoMotor", "=HESAP!C42"),
        (51, "ddl_tahminAralik", "=HESAP!C40"),
        (52, "ddl_yuzdelikP90", "=HESAP!C41"),
        (53, "ddl_sezonsalTrend",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"12_AY",TEXT(HESAP!C44,"0.0%"),"doluluk",'
         'TEXT(HESAP!C43,"0"),"pick sezonsal")'),
        (54, "ddl_alarmListesi",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Alarm",HESAP!C16,"Yetim",HESAP!C17,'
         '"DolulukAşım",HESAP!C20,"VerimAlt",HESAP!C21)'),
        (55, "ddl_kararKapi",
         '=IF(HESAP!C7=0,"VERİ YOK",'
         'IF(OR(HESAP!C8<0,HESAP!C9<0,HESAP!C10<0),"KRİTİK",'
         'IF(HESAP!C45=1,"KRİTİK",'
         'IF(OR(HESAP!C20>0,HESAP!C8>ddl_esikDoluluk),"DİKKAT","UYGUN"))))'),
        (56, "ddl_kararGerekce",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Karar",ddl_kararKapi,"| doluluk",'
         'TEXT(ddl_ortDoluluk,"0.0%"),"| sağlık",TEXT(ddl_saglikSkor,"0"),'
         '"| alarm",ddl_alarmAdet)'),
        (57, "ddl_sonrakiKimlik",
         '="LOC-"&TEXT(COUNTA(tblVarliklar[VarlikId])+1,"00000")'),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 59, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (60, "ddl_yorumGunluk",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Günlük doluluk ort",TEXT(ddl_gunlukDolulukOrt,"0.0%"))'),
        (61, "ddl_yorumAylik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Aylık doluluk proxy",TEXT(ddl_aylikDolulukTrend,"0.0%"))'),
        (62, "ddl_yorumSapma",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Doluluk sapma",TEXT(ddl_dolulukSapma,"0.0%"))'),
        (63, "ddl_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(ddl_tornadoZirve,"0"))'),
        (64, "ddl_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(ddl_senaryoKarsilastirma,"0.0%"))'),
        (65, "ddl_yorumHhi",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma HHI",TEXT(ddl_hhi,"0.0%"))'),
        (66, "ddl_yorumKalite",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite skoru",ddl_kaliteSkor)'),
        (67, "ddl_yorumSenMotor",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motoru",ddl_senaryoMotor)'),
        (68, "ddl_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(ddl_tahminAralik,"0.0%"))'),
        (69, "ddl_yorumP90",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 doluluk",TEXT(ddl_yuzdelikP90,"0.0%"))'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 28, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. VARLIKLAR sayfasına lokasyon kartı ekleyin; sonraki kimliği kullanın.",
        "2. GUNLUK_GIRIS sayfasına günlük satır yazın: varlık, tarih, doluluk %, palet, pick, olay (≤6 alan).",
        "3. Doluluk ve koridor verim skorları formülle hesaplanır; eşikler AYARLAR'dadır.",
        "4. DONEM_ANALIZ 12 aylık sezonsal doluluk zincirini gösterir.",
        "5. ALARMLAR sayfasında varlık + eşik + tarih listesi.",
        "6. raporTarihi AYARLAR'dadır; TODAY kullanılmaz.",
        "7. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı hücreler açıktır.",
        "8. RAPOR sayfasını PDF olarak yazdırabilirsiniz.",
        f"9. Sürüm {SURUM} | Kod DDL-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
    ]
    for i, m in enumerate(maddeler, 5):
        h(ws, i, 1, m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    genislik(ws, {"A": 90})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 27):
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
            f"E{r}", CellIsRule(operator="greaterThan", formula=["ddl_esikDoluluk"], fill=sari))
        ws.conditional_formatting.add(
            f"F{r}", CellIsRule(operator="lessThan", formula=["ddl_esikSaglik"], fill=kirmizi))

    ws = wb["GUNLUK_GIRIS"]
    ws.conditional_formatting.add(
        f"L{ILK}:L{SON}",
        CellIsRule(operator="equal", formula=['"ALARM"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"N{ILK}:N{SON}",
        CellIsRule(operator="equal", formula=['"YETIM"'], fill=sari))
    ws.conditional_formatting.add(
        f"D{ILK}:D{SON}",
        CellIsRule(operator="greaterThan", formula=["ddl_esikDoluluk"], fill=sari))
    ws.conditional_formatting.add(
        f"K{ILK}:K{SON}",
        CellIsRule(operator="lessThan", formula=["ddl_esikSaglik"], fill=kirmizi))

    ws = wb["VARLIKLAR"]
    ws.conditional_formatting.add(
        f"H{ILK}:H{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"L{ILK}:L{SON}",
        CellIsRule(operator="lessThan", formula=["ddl_esikSaglik"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"I{ILK}:I{SON}",
        CellIsRule(operator="greaterThan", formula=["ddl_esikDoluluk"], fill=sari))

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
            f"C{r}", CellIsRule(operator="greaterThan", formula=["ddl_esikDoluluk"], fill=sari))
        ws.conditional_formatting.add(
            f"D{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "ddl_esikDoluluk", "ddl_esikVerim", "ddl_esikSaglik",
        "ddl_varsayilanKapasitePalet", "ddl_verimCarpan", "ddl_aylikCarpan",
        "ddl_senaryoTemkinli", "ddl_senaryoBaz", "ddl_senaryoIyimser",
        "ddl_tornadoOran", "ddl_yuzdelikOran", "ddl_percentileOran", "ddl_tahminCarpan",
        "ddl_olcekHedef", "ddl_kaliteTaban", "ddl_kaliteCeza", "ddl_kaliteSapma",
        "ddl_dosyaSurumu", "ddl_firmaUnvan", "ddl_alarmEsikMetin",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "ddl_doluKayit": 32, "ddl_ortDoluluk": 33, "ddl_toplamPalet": 34,
        "ddl_toplamPick": 35, "ddl_ortVerim": 36, "ddl_kapasiteKullanim": 37,
        "ddl_bosPalet": 38, "ddl_aktifLokasyon": 39, "ddl_saglikSkor": 40,
        "ddl_alarmAdet": 41, "ddl_yetimAdet": 42, "ddl_gunlukDolulukOrt": 43,
        "ddl_aylikDolulukTrend": 44, "ddl_dolulukSapma": 45, "ddl_tornadoZirve": 46,
        "ddl_senaryoKarsilastirma": 47, "ddl_hhi": 48, "ddl_kaliteSkor": 49,
        "ddl_senaryoMotor": 50, "ddl_tahminAralik": 51, "ddl_yuzdelikP90": 52,
        "ddl_sezonsalTrend": 53, "ddl_alarmListesi": 54, "ddl_kararKapi": 55,
        "ddl_kararGerekce": 56, "ddl_sonrakiKimlik": 57,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "ddl_yorumGunluk": 60, "ddl_yorumAylik": 61, "ddl_yorumSapma": 62,
        "ddl_yorumTornado": 63, "ddl_yorumSenaryo": 64, "ddl_yorumHhi": 65,
        "ddl_yorumKalite": 66, "ddl_yorumSenMotor": 67, "ddl_yorumTahmin": 68,
        "ddl_yorumP90": 69,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$A$7:$A$8")
    ad_ekle(wb, "ListeOlayTuru", "LISTELER!$C$7:$C$12")
    ad_ekle(wb, "ListeLokasyonTipi", "LISTELER!$E$7:$E$9")


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
    dosya_ad = "DepoDolulukLokasyon.xlsx"
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
        f.write("SatirId,VarlikId,Tarih,DolulukPct_pct,PaletAdet_adet,PickAdet_adet,OlayTuru\n")
        for s in ORNEK:
            f.write(f"{s['id']},{s['varlik']},{s['tarih']:%d.%m.%Y},"
                    f"{s['doluluk']},{s['palet']},{s['pick']},{s['olay']}\n")

    print(f"Dosya oluşturuldu: {hedef}")
    print(f"Kopya: {cikti}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
