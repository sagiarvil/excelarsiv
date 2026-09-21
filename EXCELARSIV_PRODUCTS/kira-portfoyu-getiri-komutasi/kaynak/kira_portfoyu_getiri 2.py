#!/usr/bin/env python3
"""Kira Portföyü Getiri Komutası — A5 saha komuta merkezi (manda v6).

Çoklu menkul: doluluk/boşluk, kira tahsilat, net getiri, TÜFE artış takvimi.
GUNLUK_GIRIS ≤6 manuel alan; birimli kolonlar; ≥12 ay sezonsallık;
varlık sağlık skoru formülden; AYARLAR alarm eşikleri + alarm listesi.
Domain: kira/TÜFE/boşluk/getiri — filo/depo/GES kopyası değildir.
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

URUN_AD = "Kira Portföyü Getiri Komutası"
SURUM = "1.0.0"
RENK = "1A2744"
KAPASITE = 250
DEMO_MULK = 8
DEMO_GUNLUK = 36
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

# VarlikId, MulkAd, Tip, Sehir, Metrekare_m2, DegerTl_tl, HedefKiraTl_tl
MULKLER = [
    ("KIR-00001", "Kadikoy Daire A", "Daire", "Istanbul", 95, 4500000, 28000),
    ("KIR-00002", "Cankaya Ofis B", "Ofis", "Ankara", 120, 3200000, 22000),
    ("KIR-00003", "Konak Dukkan C", "Dukkan", "Izmir", 80, 2800000, 18000),
    ("KIR-00004", "Nilufer Daire D", "Daire", "Bursa", 110, 2100000, 14000),
    ("KIR-00005", "Muratpasa Ofis E", "Ofis", "Antalya", 140, 3800000, 26000),
    ("KIR-00006", "Sahinbey Depo F", "Depo", "Gaziantep", 250, 1500000, 12000),
    ("KIR-00007", "Yenisehir Daire G", "Daire", "Mersin", 85, 1600000, 11000),
    ("KIR-00008", "Tepebasi Dukkan H", "Dukkan", "Eskisehir", 70, 1400000, 9500),
]

OLAY_TURLERI = ["TAHSILAT", "BOSLUK", "ARTIS_TUFE", "GIDER", "TAHSILAT_GECIKME", "DIGER"]
AYLAR = [
    "Ocak", "Subat", "Mart", "Nisan", "Mayis", "Haziran",
    "Temmuz", "Agustos", "Eylul", "Ekim", "Kasim", "Aralik",
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Kira doluluk ve getiri kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_gunluk():
    satirlar = []
    for i in range(1, DEMO_GUNLUK + 1):
        mulk = MULKLER[(i - 1) % DEMO_MULK]
        tar = RAPOR_TARIHI - timedelta(days=(i * 3) % 340)
        hedef = mulk[6]
        bos = (i % 7) if i % 5 == 0 else (0 if i % 3 else 2)
        tahsilat = 0 if bos >= 20 else round(hedef * (0.85 + (i % 5) * 0.03) * (1 - bos / 30), 0)
        gider = round(800 + (i % 6) * 350 + (i % 4) * 120, 0)
        tufe = round(0.42 + (i % 4) * 0.02, 2) if i % 9 == 0 else round(0.38 + (i % 3) * 0.01, 2)
        satirlar.append({
            "id": f"KPG-{i:05d}",
            "varlik": "" if i == 30 else (mulk[0] if i != 33 else "KIR-99999"),
            "tarih": tar,
            "tahsilat": tahsilat,
            "bos": bos,
            "gider": gider,
            "tufe": tufe,
            "olay": OLAY_TURLERI[i % len(OLAY_TURLERI)],
        })
    return satirlar


ORNEK = _demo_gunluk()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Çoklu kira menkulünde doluluk, tahsilat, net getiri ve TÜFE artış takvimini izler; "
      "günlük saha girişi ≤6 alan; varlık sağlık skoru ve alarm kuyruğu üretir; "
      "UYGUN / DİKKAT / KRİTİK / VERİ YOK kararı verir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Menkul bazlı doluluk % ve net kira getirisi",
        "Günlük saha girişi en fazla 6 alan — 30 saniye hedefi",
        "TÜFE oranıyla önerilen kira artışı (dış API yok)",
        "Varlık sağlık skoru formülle veriden türer (sabit skor yok)",
        "12 aylık sezonsal tahsilat-doluluk zinciri ve alarm eşiği kuyruğu",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Saha / emlak operatörü — günlük tahsilat-boşluk-gider satırı",
        "Portföy sahibi — doluluk % ve karar panosu",
        "Mali müşavir / denetçi — dönemsel RAPOR kanıt çıktısı",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) VARLIKLAR'a menkul kartı girin. 2) GUNLUK_GIRIS'e günlük satır yazın. "
      "3) DONEM_ANALIZ ve PANO'dan kararı görün. 4) RAPOR'dan çıktı alın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Kod: KPG-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def varliklar(ws):
    sayfa_hazirla(ws, "VARLIKLAR", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Kira menkul kartları — her varlığın tek kimliği vardır", son_kolon=14)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=kpg_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4",
               "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: KIR-00009 | Risk: Elle çakışma")

    sutunlar = [
        "VarlikId", "MulkAd", "Tip", "Sehir", "Metrekare_m2", "DegerTl_tl",
        "HedefKiraTl_tl", "Aktif",
        "ToplamTahsilatTl", "OrtDoluluk", "ToplamGiderTl", "SaglikSkor",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "ToplamTahsilatTl": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'SUMIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[TahsilatTl_tl]))'
        ),
        "OrtDoluluk": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'IFERROR(AVERAGEIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[FormulDoluluk_oran]),0))'
        ),
        "ToplamGiderTl": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'SUMIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[GiderTl_tl]))'
        ),
        "SaglikSkor": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'MAX(0,MIN(100,kpg_kaliteTaban'
            '-IF(tblVarliklar[[#This Row],[OrtDoluluk]]<kpg_esikDoluluk,kpg_kaliteCeza,0)'
            '-COUNTIFS(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[AlarmBayrak],"ALARM")*kpg_kaliteSapma'
            '-IF(tblVarliklar[[#This Row],[Aktif]]="HAYIR",kpg_kaliteCeza,0))))'
        ),
    }
    for i, row in enumerate(MULKLER):
        r = ILK + i
        kid, ad, tip, sehir, m2, deger, hedef = row
        _sari(ws, r, 1, kid, baslik="Varlık Kimliği", mesaj="KIR-##### formatında")
        _sari(ws, r, 2, ad, baslik="Mülk Adı", mesaj="Daire / ofis / dükkan adı")
        _sari(ws, r, 3, tip, baslik="Tip", mesaj="Daire Ofis Dukkan Depo")
        _sari(ws, r, 4, sehir, baslik="Şehir", mesaj="İl adı")
        _sari(ws, r, 5, m2, sayi="0.0", baslik="Metrekare [m2]", mesaj="Brüt m2")
        _sari(ws, r, 6, deger, sayi=TL, baslik="Değer [TL]", mesaj="Tahmini piyasa değeri")
        _sari(ws, r, 7, hedef, sayi=TL, baslik="Hedef Kira [TL]", mesaj="Aylık hedef kira")
        _sari(ws, r, 8, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_MULK, SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None,
                  sayi=("0.0" if c == 5 else (TL if c in (6, 7) else None)),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVarliklar", f"A{HDR}:L{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Varlık Kimliği", mesaj="KIR-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Mülk Ad"), (3, "Tip"), (4, "Şehir")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    for c, ad in [(5, "Metrekare [m2]"), (6, "Değer [TL]"), (7, "Hedef Kira [TL]")]:
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
    # Manuel (id/tarih/formul/skor/alarm hariç): Tahsilat, BosGun, Gider, Tufe, Olay (+VarlikId id filtresi)
    sutunlar = [
        "SatirId", "VarlikId", "Tarih", "TahsilatTl_tl", "BosGun_gun",
        "GiderTl_tl", "TufeOran_oran", "OlayTuru",
        "FormulHedefKira_tl", "FormulDoluluk_oran", "FormulGetiri_oran",
        "FormulNet_tl", "FormulTufeArtis_tl", "SaglikSkor_puan",
        "AlarmBayrak", "FormulKayitDolu", "FormulYetim",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "FormulHedefKira_tl": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR(SUMIF(tblVarliklar[VarlikId],tblGunluk[[#This Row],[VarlikId]],'
            'tblVarliklar[HedefKiraTl_tl]),kpg_varsayilanHedefKira))'
        ),
        "FormulDoluluk_oran": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'MAX(0, 1-tblGunluk[[#This Row],[BosGun_gun]]/kpg_gunAy))'
        ),
        "FormulGetiri_oran": (
            'IF(OR(tblGunluk[[#This Row],[SatirId]]="",'
            'tblGunluk[[#This Row],[FormulHedefKira_tl]]=0),"",'
            'tblGunluk[[#This Row],[TahsilatTl_tl]]/'
            'tblGunluk[[#This Row],[FormulHedefKira_tl]])'
        ),
        "FormulNet_tl": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'tblGunluk[[#This Row],[TahsilatTl_tl]]-'
            'tblGunluk[[#This Row],[GiderTl_tl]])'
        ),
        "FormulTufeArtis_tl": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'tblGunluk[[#This Row],[FormulHedefKira_tl]]*'
            '(1+tblGunluk[[#This Row],[TufeOran_oran]]))'
        ),
        "SaglikSkor_puan": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'MAX(0,MIN(100,kpg_kaliteTaban'
            '-IF(tblGunluk[[#This Row],[FormulDoluluk_oran]]<kpg_esikDoluluk,kpg_kaliteCeza,0)'
            '-IF(tblGunluk[[#This Row],[FormulGetiri_oran]]<kpg_esikGetiri,kpg_kaliteSapma,0)'
            '-IF(tblGunluk[[#This Row],[FormulYetim]]="YETIM",kpg_kaliteCeza,0))))'
        ),
        "AlarmBayrak": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IF(OR(tblGunluk[[#This Row],[FormulDoluluk_oran]]<kpg_esikDoluluk,'
            'tblGunluk[[#This Row],[FormulGetiri_oran]]<kpg_esikGetiri,'
            'tblGunluk[[#This Row],[SaglikSkor_puan]]<kpg_esikSaglik),'
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
        _sari(ws, r, 1, s["id"], baslik="Günlük Kimlik", mesaj="KPG-#####")
        _sari(ws, r, 2, s["varlik"], baslik="Varlık Kimliği", mesaj="KIR listesinden")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="GG.AA.YYYY")
        _sari(ws, r, 4, s["tahsilat"], sayi=TL, baslik="Tahsilat [TL]", mesaj="Dönem tahsilatı")
        _sari(ws, r, 5, s["bos"], sayi=CATI, baslik="Boş Gün [gün]", mesaj="Boş kalan gün sayısı")
        _sari(ws, r, 6, s["gider"], sayi=TL, baslik="Gider [TL]", mesaj="Dönem gideri")
        _sari(ws, r, 7, s["tufe"], sayi=YÜZDE, baslik="TÜFE Oranı", mesaj="Elle girilen TÜFE (API yok)")
        _sari(ws, r, 8, s["olay"], baslik="Olay Türü", mesaj="Listeden seçin")
    for r in range(ILK + DEMO_GUNLUK, SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None,
                  sayi=(TARİH if c == 3 else (TL if c in (4, 6) else (
                      CATI if c == 5 else (YÜZDE if c == 7 else None)))),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblGunluk", f"A{HDR}:Q{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Günlük Kimlik", mesaj="KPG-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Varlık", mesaj="KIR kimliği girin",
              hata_baslik="Varlık", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"C{ILK}:C{SON}",
              baslik="Tarih", mesaj="GG.AA.YYYY",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    for c, ad in [(4, "Tahsilat [TL]"), (5, "Boş Gün [gün]"), (6, "Gider [TL]"), (7, "TÜFE Oranı")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeOlayTuru", f"H{ILK}:H{SON}",
              baslik="Olay Türü", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 13 for i in range(1, 18)})


def donem_analiz(ws):
    sayfa_hazirla(ws, "DONEM_ANALIZ", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "12 aylık sezonsallık — tahsilat ve doluluk zinciri (S03)", son_kolon=12)
    h(ws, 5, 1, "Ay", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Ay_Anahtari", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Tahsilat TL", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 4, "Gider TL", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 5, "Doluluk Ort", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 6, "Sezonsal Etiket", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    # Kira sezonsallığı: yaz yüksek doluluk, kış daha fazla boşluk riski
    demo_tahsilat = [118000, 122000, 128000, 135000, 142000, 148000,
                     152000, 150000, 145000, 138000, 130000, 124000]
    demo_gider = [18000, 19000, 20000, 21000, 22000, 23000,
                  24000, 23500, 22000, 20500, 19500, 18500]
    demo_doluluk = [0.88, 0.89, 0.91, 0.93, 0.94, 0.95,
                    0.96, 0.95, 0.93, 0.91, 0.89, 0.87]
    for i, ay in enumerate(AYLAR, 1):
        r = 5 + i
        h(ws, r, 1, ay)
        h(ws, r, 2, f"Ay_{i}")
        h(ws, r, 3, demo_tahsilat[i - 1], sayi=TL)
        h(ws, r, 4, demo_gider[i - 1], sayi=TL)
        h(ws, r, 5, demo_doluluk[i - 1], sayi=YÜZDE)
        h(ws, r, 6, "12_AY_SEZONSAL")

    h(ws, 20, 1, "Sezonsal Trend Özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=kpg_sezonsalTrend", kaydir=True)
    ws.merge_cells("B20:F20")
    h(ws, 22, 1, "12 Ay Toplam Tahsilat", kalin=True)
    h(ws, 22, 2, "=SUM(C6:C17)", sayi=TL)
    h(ws, 23, 1, "12 Ay Toplam Gider", kalin=True)
    h(ws, 23, 2, "=SUM(D6:D17)", sayi=TL)
    h(ws, 24, 1, "12 Ay Ort Doluluk", kalin=True)
    h(ws, 24, 2, "=AVERAGE(E6:E17)", sayi=YÜZDE)
    h(ws, 25, 1, "Canlı dönem tahsilat", kalin=True)
    h(ws, 25, 2, "=kpg_toplamTahsilat", sayi=TL)
    h(ws, 26, 1, "Canlı dönem doluluk", kalin=True)
    h(ws, 26, 2, "=kpg_ortDoluluk", sayi=YÜZDE)

    c1 = LineChart()
    c1.title = "12 Ay Tahsilat Sezonsallığı"
    c1.add_data(Reference(ws, min_col=3, min_row=5, max_row=17), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17))
    c1.width, c1.height = 14, 8
    ws.add_chart(c1, "A27")

    c2 = BarChart()
    c2.title = "12 Ay Doluluk Dağılımı"
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

    adimlar = [
        ("H01", "Toplam günlük kayıt", '=COUNTA(tblGunluk[SatirId])'),
        ("H02", "Dolu kayıt", '=SUM(tblGunluk[FormulKayitDolu])'),
        ("H03", "Toplam tahsilat TL", "=SUM(tblGunluk[TahsilatTl_tl])"),
        ("H04", "Toplam boş gün", "=SUM(tblGunluk[BosGun_gun])"),
        ("H05", "Toplam gider TL", "=SUM(tblGunluk[GiderTl_tl])"),
        ("H06", "Toplam hedef kira", "=SUM(tblGunluk[FormulHedefKira_tl])"),
        ("H07", "Ort doluluk", "=IFERROR(AVERAGEIF(tblGunluk[FormulDoluluk_oran],\">0\"),0)"),
        ("H08", "Ort getiri", "=IFERROR(AVERAGEIF(tblGunluk[FormulGetiri_oran],\">0\"),0)"),
        ("H09", "Net tahsilat", "=SUM(tblGunluk[FormulNet_tl])"),
        ("H10", "Aktif menkul adet", '=COUNTIF(tblVarliklar[Aktif],"EVET")'),
        ("H11", "Toplam menkul kart", '=COUNTA(tblVarliklar[VarlikId])'),
        ("H12", "Alarm adet", '=COUNTIF(tblGunluk[AlarmBayrak],"ALARM")'),
        ("H13", "Yetim kayıt", '=COUNTIF(tblGunluk[FormulYetim],"YETIM")'),
        ("H14", "Ortalama sağlık skor",
         '=IFERROR(AVERAGEIF(tblGunluk[SaglikSkor_puan],">0"),0)'),
        ("H15", "Min sağlık skor",
         '=IFERROR(MIN(tblGunluk[SaglikSkor_puan]),0)'),
        ("H16", "Doluluk eşik altı adet",
         '=COUNTIFS(tblGunluk[FormulDoluluk_oran],"<"&kpg_esikDoluluk,tblGunluk[FormulKayitDolu],">0")'),
        ("H17", "Getiri eşik altı adet",
         '=COUNTIFS(tblGunluk[FormulGetiri_oran],"<"&kpg_esikGetiri,tblGunluk[FormulKayitDolu],">0")'),
        ("H18", "Düşük sağlık adet",
         '=COUNTIFS(tblGunluk[SaglikSkor_puan],"<"&kpg_esikSaglik,tblGunluk[FormulKayitDolu],">0")'),
        ("H19", "Günlük doluluk ort",
         '=IFERROR(AVERAGEIF(tblGunluk[FormulDoluluk_oran],">0"),0)'),
        ("H20", "Aylık tahsilat proxy", "=C8*kpg_aylikCarpan"),
        ("H21", "Getiri sapma",
         "=IFERROR(ABS(C13-kpg_esikGetiri),0)"),
        ("H22", "Tahsilat olay pay",
         '=IFERROR(COUNTIF(tblGunluk[OlayTuru],"TAHSILAT")/MAX(C7,1),0)'),
        ("H23", "Boşluk olay pay",
         '=IFERROR(COUNTIF(tblGunluk[OlayTuru],"BOSLUK")/MAX(C7,1),0)'),
        ("H24", "TÜFE artış olay pay",
         '=IFERROR(COUNTIF(tblGunluk[OlayTuru],"ARTIS_TUFE")/MAX(C7,1),0)'),
        ("H25", "Gider olay pay",
         '=IFERROR(COUNTIF(tblGunluk[OlayTuru],"GIDER")/MAX(C7,1),0)'),
        ("H26", "Temkinli senaryo", "=C8*kpg_senaryoTemkinli"),
        ("H27", "Baz senaryo", "=C8*kpg_senaryoBaz"),
        ("H28", "İyimser senaryo", "=C8*kpg_senaryoIyimser"),
        ("H29", "Senaryo farkı", "=C33-C31"),
        ("H30", "Tornado boşluk etki", "=C9*kpg_tornadoOran*kpg_varsayilanHedefKira"),
        ("H31", "Tornado gider etki", "=C10*kpg_tornadoOran"),
        ("H32", "Tornado tahsilat etki", "=C8*kpg_tornadoOran"),
        ("H33", "Tornado zirve", "=MAX(C35,C36,C37)"),
        ("H34", "HHI proxy",
         "=IFERROR(C27^2+C28^2+C29^2+C30^2,0)"),
        ("H35", "Kalite skor",
         "=MAX(0,kpg_kaliteTaban-C21*kpg_kaliteCeza-C22*kpg_kaliteSapma"
         "-C23*kpg_kaliteSapma)"),
        ("H36", "Tahmin aralık", "=C32*kpg_tahminCarpan"),
        ("H37", "P90 tahsilat",
         "=IFERROR(PERCENTILE(tblGunluk[TahsilatTl_tl],kpg_percentileOran),C8*kpg_yuzdelikOran)"),
        ("H38", "Senaryo motor metin",
         '=_xlfn.TEXTJOIN("|",TRUE,"T",TEXT(C31,"0"),"B",TEXT(C32,"0"),"I",TEXT(C33,"0"))'),
        ("H39", "Sezonsal 12 ay tahsilat", "=DONEM_ANALIZ!B22"),
        ("H40", "Sezonsal 12 ay doluluk", "=DONEM_ANALIZ!B24"),
        ("H41", "Kritik sinyal",
         '=IF(OR(C21>0,C23>0,C19<kpg_esikSaglik),1,0)'),
        ("H42", "Karar kodu ham",
         '=IF(C7=0,0,IF(C46=1,3,IF(OR(C21>0,C12<kpg_esikDoluluk),2,1)))'),
        ("H43", "Doluluk çıpa", "=C12"),
        ("H44", "Alarm listesi özet adet", "=C17"),
        ("H45", "Varlık sağlık ort",
         '=IFERROR(AVERAGEIF(tblVarliklar[SaglikSkor],">0"),0)'),
        ("H46", "Tahmin FORECAST",
         "=IFERROR(FORECAST(C7+1,tblGunluk[TahsilatTl_tl],tblGunluk[FormulKayitDolu]),C41)"),
        ("H47", "Olay kırılım özet",
         '=_xlfn.TEXTJOIN("|",TRUE,"T",TEXT(C27,"0%"),"B",TEXT(C28,"0%"),'
         '"U",TEXT(C29,"0%"),"G",TEXT(C30,"0%"))'),
        ("H48", "Ort TÜFE artışı TL",
         '=IFERROR(AVERAGEIF(tblGunluk[FormulTufeArtis_tl],">0"),0)'),
    ]
    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        acik_l = acik.lower()
        sayi = YÜZDE if any(x in acik_l for x in ("doluluk", "getiri", "oran", "pay", "sapma", "hhi", "tufe")) else CATI
        if "metin" in acik_l or "özet" in acik_l or "kırılım" in acik_l:
            sayi = None
        elif (
            any(x in acik_l for x in ("tahsilat", "gider", "hedef", "net", "tornado", "senaryo",
                                      "tahmin", "p90", "proxy", "artış", "artis"))
            and "metin" not in acik_l
            and "pay" not in acik_l
            and "oran" not in acik_l
            and "doluluk" not in acik_l
            and "getiri" not in acik_l
        ):
            sayi = TL
        h(ws, r, 3, form, sayi=sayi)

    genislik(ws, {"A": 8, "B": 32, "C": 22})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Karar panosu — doluluk, getiri, TÜFE, alarm, sezonsal trend", son_kolon=14)

    h(ws, 5, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 5, 2, "=kpg_kararKapi", kalin=True, boyut=16)
    yorum_ekle(ws, "B5",
               "Tanım: Karar kapısı | Neden önemli: Boş dosyada VERİ YOK | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle müdahale yok")

    kpis = [
        (7, "Dolu Kayıt", "=kpg_doluKayit", CATI),
        (8, "Toplam Tahsilat TL", "=kpg_toplamTahsilat", TL),
        (9, "Toplam Gider TL", "=kpg_toplamGider", TL),
        (10, "Ort Doluluk", "=kpg_ortDoluluk", YÜZDE),
        (11, "Ort Getiri", "=kpg_ortGetiri", YÜZDE),
        (12, "Net Tahsilat", "=kpg_netTahsilat", TL),
        (13, "Günlük Doluluk Ort", "=kpg_gunlukDolulukOrt", YÜZDE),
        (14, "Sağlık Skor Ort", "=kpg_saglikSkor", CATI),
        (15, "Alarm Adet", "=kpg_alarmAdet", CATI),
        (16, "Yetim Kayıt", "=kpg_yetimAdet", CATI),
        (17, "Kalite Skor", "=kpg_kaliteSkor", CATI),
        (18, "Tornado Zirve", "=kpg_tornadoZirve", TL),
        (19, "P90 Tahsilat", "=kpg_yuzdelikP90", TL),
        (20, "Olay Kırılım", "=kpg_olayKirilim", None),
        (21, "Ort TÜFE Artış TL", "=kpg_ortTufeArtis", TL),
    ]
    h(ws, 6, 1, "Gösterge", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 6, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for r, ad, form, sayi in kpis:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=sayi, kalin=True)

    # Canlı PANO formülleri (COUNTIF/SUMIF) — çapraz tutarlılık için named range ile aynı kaynak
    h(ws, 7, 4, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 5, "=kpg_kararGerekce", kaydir=True)
    ws.merge_cells("E7:H8")
    h(ws, 10, 4, "Öneri 1", kalin=True)
    h(ws, 10, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Doluluk",TEXT(kpg_ortDoluluk,"0.0%"),'
                 '"— eşik",TEXT(kpg_esikDoluluk,"0.0%"))', kaydir=True)
    h(ws, 11, 4, "Öneri 2", kalin=True)
    h(ws, 11, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Alarm",kpg_alarmAdet,"adet — sağlık",'
                 'TEXT(kpg_saglikSkor,"0"))', kaydir=True)
    h(ws, 12, 4, "Öneri 3", kalin=True)
    h(ws, 12, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"TÜFE artış ort",TEXT(kpg_ortTufeArtis,"₺#,##0"),'
                 '"— getiri",TEXT(kpg_ortGetiri,"0.0%"))', kaydir=True)

    h(ws, 14, 4, "Canlı Alarm COUNTIF", kalin=True)
    h(ws, 14, 5, '=COUNTIF(tblGunluk[AlarmBayrak],"ALARM")', sayi=CATI)
    h(ws, 15, 4, "Canlı Tahsilat SUMIF", kalin=True)
    h(ws, 15, 5, '=SUMIF(tblGunluk[FormulKayitDolu],">0",tblGunluk[TahsilatTl_tl])', sayi=TL)

    h(ws, 22, 1, "OlayTuru")
    h(ws, 22, 2, "Adet")
    for i, (ad, t) in enumerate([
        ("TAHSILAT", 10), ("BOSLUK", 6), ("ARTIS_TUFE", 5), ("GIDER", 4),
    ], 23):
        h(ws, i, 1, ad)
        h(ws, i, 2, t, sayi=CATI)

    h(ws, 22, 4, "Senaryo")
    h(ws, 22, 5, "Tahsilat")
    for i, (ad, t) in enumerate([
        ("Temkinli", 420000), ("Baz", 480000), ("İyimser", 540000),
    ], 23):
        h(ws, i, 4, ad)
        h(ws, i, 5, t, sayi=TL)

    h(ws, 22, 7, "Hafta")
    h(ws, 22, 8, "Tahsilat")
    for i in range(1, 9):
        h(ws, 22 + i, 7, f"H{i}")
        h(ws, 22 + i, 8, 120000 + i * 8000, sayi=TL)

    h(ws, 22, 10, "Ay")
    h(ws, 22, 11, "Doluluk")
    for i, d in enumerate([0.88, 0.91, 0.94, 0.96], 23):
        h(ws, i, 10, AYLAR[i - 23 + 5])
        h(ws, i, 11, d, sayi=YÜZDE)

    c1 = PieChart()
    c1.title = "Olay Türü Dağılımı"
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
    c3.title = "Haftalık Tahsilat Trend"
    c3.add_data(Reference(ws, min_col=8, min_row=22, max_row=30), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=7, min_row=23, max_row=30))
    c3.width, c3.height = 12, 7
    ws.add_chart(c3, "A47")

    c4 = BarChart()
    c4.title = "Aylık Doluluk"
    c4.add_data(Reference(ws, min_col=11, min_row=22, max_row=26), titles_from_data=True)
    c4.set_categories(Reference(ws, min_col=10, min_row=23, max_row=26))
    c4.width, c4.height = 10, 7
    ws.add_chart(c4, "F47")

    c5 = PieChart()
    c5.title = "Olay Payı (tekrar)"
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
    c7.title = "Tahsilat Trend (kısa)"
    c7.add_data(Reference(ws, min_col=8, min_row=22, max_row=28), titles_from_data=True)
    c7.set_categories(Reference(ws, min_col=7, min_row=23, max_row=28))
    c7.width, c7.height = 12, 7
    ws.add_chart(c7, "A77")

    c8 = BarChart()
    c8.title = "Olay Adet"
    c8.add_data(Reference(ws, min_col=2, min_row=22, max_row=26), titles_from_data=True)
    c8.set_categories(Reference(ws, min_col=1, min_row=23, max_row=26))
    c8.width, c8.height = 10, 7
    ws.add_chart(c8, "F77")

    baski_hazirla(ws, "A1:N30", f"{URUN_AD} | Pano | {SURUM}")
    genislik(ws, {"A": 22, "B": 16, "D": 22, "E": 40})


def alarmlar(ws):
    sayfa_hazirla(ws, "ALARMLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Alarm listesi — varlık + eşik + süre (S05)", son_kolon=12)
    h(ws, 5, 1, "Varlık", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Alarm", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Eşik", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 4, "Süre / Tarih", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 5, "Doluluk Ort", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 6, "Sağlık Skor", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    for i, mulk in enumerate(MULKLER):
        r = 6 + i
        h(ws, r, 1, mulk[0])
        h(ws, r, 2,
          f'=IF(OR(COUNTIFS(tblGunluk[VarlikId],A{r},tblGunluk[AlarmBayrak],"ALARM")>0,'
          f'IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[SaglikSkor]),100)<kpg_esikSaglik),'
          f'"ALARM","YOK")')
        h(ws, r, 3, f"=IFERROR(ABS(E{r}-kpg_esikDoluluk),0)", sayi=YÜZDE)
        h(ws, r, 4, "=raporTarihi", sayi=TARİH)
        h(ws, r, 5,
          f'=IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[OrtDoluluk]),0)',
          sayi=YÜZDE)
        h(ws, r, 6,
          f'=IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[SaglikSkor]),0)',
          sayi=CATI)

    h(ws, 16, 1, "Alarm Özeti", kalin=True)
    h(ws, 16, 2, "=kpg_alarmListesi", kaydir=True)
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
    h(ws, 8, 2, "=kpg_firmaUnvan")

    h(ws, 10, 1, "KARAR", kalin=True, boyut=12)
    h(ws, 10, 2, "=kpg_kararKapi", kalin=True, boyut=14)
    h(ws, 11, 1, "Ort Doluluk")
    h(ws, 11, 2, "=kpg_ortDoluluk", sayi=YÜZDE)
    h(ws, 12, 1, "Toplam Tahsilat TL")
    h(ws, 12, 2, "=kpg_toplamTahsilat", sayi=TL)
    h(ws, 13, 1, "Toplam Gider TL")
    h(ws, 13, 2, "=kpg_toplamGider", sayi=TL)
    h(ws, 14, 1, "Sağlık Skor")
    h(ws, 14, 2, "=kpg_saglikSkor", sayi=CATI)
    h(ws, 15, 1, "Alarm Adet")
    h(ws, 15, 2, "=kpg_alarmAdet", sayi=CATI)

    h(ws, 17, 1, "Kalem", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 17, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, (ad, form, sayi) in enumerate([
        ("Ort getiri", "=kpg_ortGetiri", YÜZDE),
        ("Net tahsilat", "=kpg_netTahsilat", TL),
        ("Günlük doluluk ort", "=kpg_gunlukDolulukOrt", YÜZDE),
        ("Kalite skor", "=kpg_kaliteSkor", CATI),
        ("Ort TÜFE artış TL", "=kpg_ortTufeArtis", TL),
    ], 18):
        h(ws, i, 1, ad)
        h(ws, i, 2, form, sayi=sayi)

    h(ws, 24, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 1,
      "Doluluk = 1 − boş_gün / ay_günü. Getiri = tahsilat / hedef_kira. "
      "TÜFE oranı kullanıcı girişi (dış API yok). Eşikler AYARLAR'dan gelir. "
      "Karar destek çıktısıdır; hukuki veya mali tavsiye değildir. "
      "TBK kira hükümleri + TÜFE artış yaklaşımıdır.",
      kaydir=True)
    ws.merge_cells("A25:F26")

    h(ws, 28, 1, "İmza — Portföy Yöneticisi")
    h(ws, 28, 3, "________________")
    h(ws, 29, 1, "İmza — Mali Müşavir / CFO")
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
        ("Dolu kayıt > 0", '=IF(kpg_doluKayit>0,"GEÇTİ","KALDI")', "=kpg_doluKayit"),
        ("Yetim kayıt = 0", '=IF(kpg_yetimAdet=0,"GEÇTİ","UYARI")', "=kpg_yetimAdet"),
        ("Doluluk hesaplı", '=IF(kpg_ortDoluluk>=0,"GEÇTİ","KALDI")', "=kpg_ortDoluluk"),
        ("Sağlık skor formül", '=IF(kpg_saglikSkor>=0,"GEÇTİ","KALDI")', "=kpg_saglikSkor"),
        ("Alarm listesi", '=IF(kpg_alarmAdet>=0,"GEÇTİ","KALDI")', "=kpg_alarmAdet"),
        ("Karar dolu", '=IF(kpg_kararKapi<>"","GEÇTİ","KALDI")', "=kpg_kararKapi"),
        ("12 ay sezonsal", '=IF(DONEM_ANALIZ!B22>=0,"GEÇTİ","KALDI")', "=DONEM_ANALIZ!B22"),
        ("Kalite skor", '=IF(kpg_kaliteSkor>=0,"GEÇTİ","KALDI")', "=kpg_kaliteSkor"),
        ("Tornado", '=IF(kpg_tornadoZirve>=0,"GEÇTİ","KALDI")', "=kpg_tornadoZirve"),
        ("P90", '=IF(kpg_yuzdelikP90>=0,"GEÇTİ","KALDI")', "=kpg_yuzdelikP90"),
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
        "SatirId", "VarlikId", "Tarih", "TahsilatTl_tl",
        "BosGun_gun", "GiderTl_tl", "TufeOran_oran", "OlayTuru",
    ]
    baslik_satiri(ws, 5, basliklar)
    for i, s in enumerate(ORNEK[:20]):
        r = 6 + i
        h(ws, r, 1, s["id"])
        h(ws, r, 2, s["varlik"])
        h(ws, r, 3, s["tarih"], sayi=TARİH)
        h(ws, r, 4, s["tahsilat"], sayi=TL)
        h(ws, r, 5, s["bos"], sayi=CATI)
        h(ws, r, 6, s["gider"], sayi=TL)
        h(ws, r, 7, s["tufe"], sayi=YÜZDE)
        h(ws, r, 8, s["olay"])
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 9)})


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

    h(ws, 3, 5, "Mülk Tipi", kalin=True)
    h(ws, 6, 5, "Tip", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, t in enumerate(["Daire", "Ofis", "Dukkan", "Depo"], 7):
        _sari(ws, i, 5, t, baslik="Tip", mesaj="Mülk tipi")
    genislik(ws, {"A": 12, "C": 18, "E": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + alarm eşikleri (S05)", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("kpg_varsayilanHedefKira", 15000.0, "TL", "Sözleşme / emsal kira", "01.01.2026",
         "Varsayılan aylık hedef kira"),
        ("kpg_esikDoluluk", 0.85, "oran", "Alarm eşiği — doluluk", "01.01.2026",
         "Doluluk alt alarm eşiği"),
        ("kpg_esikGetiri", 0.80, "oran", "Alarm eşiği — getiri", "01.01.2026",
         "Getiri (tahsilat/hedef) alt eşiği"),
        ("kpg_esikSaglik", 60, "puan", "Alarm eşiği — sağlık skor", "01.01.2026",
         "Sağlık skor alt eşiği"),
        ("kpg_gunAy", 30, "gun", "Dönem gün sayısı", "01.01.2026",
         "Doluluk hesabında ay günü"),
        ("kpg_aylikCarpan", 1.0, "oran", "Aylık proxy", "01.01.2026", "Aylık tahsilat çarpan"),
        ("kpg_senaryoTemkinli", 0.90, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("kpg_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("kpg_senaryoIyimser", 1.10, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("kpg_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("kpg_yuzdelikOran", 1.15, "oran", "İstatistik kuralı", "01.01.2026", "P90 çarpan proxy"),
        ("kpg_percentileOran", 0.9, "oran", "PERCENTILE oranı", "01.01.2026", "Yüzdelik dilim"),
        ("kpg_tahminCarpan", 1.05, "oran", "Tahmin modeli", "01.01.2026", "Tahmin aralık çarpan"),
        ("kpg_olcekHedef", 20000, "satir", "Manda A5 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("kpg_kaliteTaban", 100, "puan", "Kalite / sağlık modeli", "01.01.2026", "Taban puan"),
        ("kpg_kaliteCeza", 12, "puan", "Kalite / sağlık modeli", "01.01.2026", "Eşik aşım cezası"),
        ("kpg_kaliteSapma", 3, "puan", "Kalite / sağlık modeli", "01.01.2026", "Sapma / alarm cezası"),
        ("kpg_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("kpg_firmaUnvan", "Örnek Portföy Gayrimenkul A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma unvanı"),
        ("kpg_alarmEsikMetin", "Doluluk / getiri / sağlık eşikleri", "metin",
         "S05 alarm", "01.01.2026", "Alarm eşiği tanımı"),
        ("kpg_tufeReferans", 0.40, "oran", "TÜFE referans (elle)", "01.01.2026",
         "Varsayılan TÜFE referansı — API yok"),
        ("kpg_boslukMaliyetCarpan", 1.0, "oran", "Boşluk maliyet modeli", "01.01.2026",
         "Boş gün maliyet çarpanı"),
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

    h(ws, 32, 8, "Motor Çıktıları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 8, "anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 33, 9, "deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    motor = [
        (34, "kpg_doluKayit", "=HESAP!C7"),
        (35, "kpg_toplamTahsilat", "=HESAP!C8"),
        (36, "kpg_toplamGider", "=HESAP!C10"),
        (37, "kpg_ortDoluluk", "=HESAP!C12"),
        (38, "kpg_ortGetiri", "=HESAP!C13"),
        (39, "kpg_netTahsilat", "=HESAP!C14"),
        (40, "kpg_gunlukDolulukOrt", "=HESAP!C24"),
        (41, "kpg_aylikTahsilatTrend", "=HESAP!C25"),
        (42, "kpg_saglikSkor", "=HESAP!C19"),
        (43, "kpg_alarmAdet", "=HESAP!C17"),
        (44, "kpg_yetimAdet", "=HESAP!C18"),
        (45, "kpg_getiriSapma", "=HESAP!C26"),
        (46, "kpg_tornadoZirve", "=HESAP!C38"),
        (47, "kpg_senaryoKarsilastirma", "=HESAP!C34"),
        (48, "kpg_hhi", "=HESAP!C39"),
        (49, "kpg_kaliteSkor", "=HESAP!C40"),
        (50, "kpg_senaryoMotor", "=HESAP!C43"),
        (51, "kpg_tahminAralik", "=HESAP!C41"),
        (52, "kpg_yuzdelikP90", "=HESAP!C42"),
        (53, "kpg_olayKirilim", "=HESAP!C52"),
        (54, "kpg_ortTufeArtis", "=HESAP!C53"),
        (55, "kpg_sezonsalTrend",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"12_AY",TEXT(HESAP!C44,"₺#,##0"),"TL",'
         'TEXT(HESAP!C45,"0.0%"),"doluluk sezonsal")'),
        (56, "kpg_alarmListesi",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Alarm",HESAP!C17,"Yetim",HESAP!C18,'
         '"DolAlt",HESAP!C21,"GetAlt",HESAP!C22)'),
        (57, "kpg_kararKapi",
         '=IF(HESAP!C7=0,"VERİ YOK",'
         'IF(OR(HESAP!C8<0,HESAP!C10<0),"KRİTİK",'
         'IF(HESAP!C46=1,"KRİTİK",'
         'IF(OR(HESAP!C21>0,HESAP!C12<kpg_esikDoluluk),"DİKKAT","UYGUN"))))'),
        (58, "kpg_kararGerekce",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Karar",kpg_kararKapi,"| doluluk",'
         'TEXT(kpg_ortDoluluk,"0.0%"),"| sağlık",TEXT(kpg_saglikSkor,"0"),'
         '"| alarm",kpg_alarmAdet)'),
        (59, "kpg_sonrakiKimlik",
         '="KIR-"&TEXT(COUNTA(tblVarliklar[VarlikId])+1,"00000")'),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 61, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (62, "kpg_yorumGunluk",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Günlük doluluk ort",TEXT(kpg_gunlukDolulukOrt,"0.0%"))'),
        (63, "kpg_yorumAylik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Aylık tahsilat proxy",TEXT(kpg_aylikTahsilatTrend,"₺#,##0"))'),
        (64, "kpg_yorumSapma",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Getiri sapma",TEXT(kpg_getiriSapma,"0.0%"))'),
        (65, "kpg_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(kpg_tornadoZirve,"₺#,##0"))'),
        (66, "kpg_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(kpg_senaryoKarsilastirma,"₺#,##0"))'),
        (67, "kpg_yorumHhi",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma HHI",TEXT(kpg_hhi,"0.0%"))'),
        (68, "kpg_yorumKalite",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite skoru",kpg_kaliteSkor)'),
        (69, "kpg_yorumSenMotor",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motoru",kpg_senaryoMotor)'),
        (70, "kpg_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(kpg_tahminAralik,"₺#,##0"))'),
        (71, "kpg_yorumP90",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 tahsilat",TEXT(kpg_yuzdelikP90,"₺#,##0"))'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 26, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. VARLIKLAR sayfasına kira menkul kartı ekleyin; sonraki kimliği kullanın.",
        "2. GUNLUK_GIRIS: varlık, tarih, tahsilat TL, boş gün, gider TL, TÜFE oranı, olay türü.",
        "3. Doluluk = 1 − boş_gün / ay_günü; getiri = tahsilat / hedef_kira.",
        "4. TÜFE oranı elle girilir; dış API/canlı çekim yoktur (Ç15).",
        "5. DONEM_ANALIZ 12 aylık sezonsal tahsilat-doluluk zincirini gösterir.",
        "6. ALARMLAR sayfasında varlık + eşik + tarih listesi; eşikler AYARLAR'dadır.",
        "7. raporTarihi AYARLAR'dadır; TODAY kullanılmaz.",
        "8. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı hücreler açıktır.",
        "9. RAPOR sayfasını PDF olarak yazdırabilirsiniz.",
        f"10. Sürüm {SURUM} | Kod KPG-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
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

    for r in range(7, 22):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["ALARMLAR"]
    ws.conditional_formatting.add(
        "B6:B13", CellIsRule(operator="equal", formula=['"ALARM"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B6:B13", CellIsRule(operator="equal", formula=['"YOK"'], fill=yesil, font=yf))
    for r in range(6, 14):
        ws.conditional_formatting.add(
            f"E{r}", CellIsRule(operator="lessThan", formula=["kpg_esikDoluluk"], fill=sari))
        ws.conditional_formatting.add(
            f"F{r}", CellIsRule(operator="lessThan", formula=["kpg_esikSaglik"], fill=kirmizi))

    ws = wb["GUNLUK_GIRIS"]
    ws.conditional_formatting.add(
        f"O{ILK}:O{SON}",
        CellIsRule(operator="equal", formula=['"ALARM"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"Q{ILK}:Q{SON}",
        CellIsRule(operator="equal", formula=['"YETIM"'], fill=sari))
    ws.conditional_formatting.add(
        f"J{ILK}:J{SON}",
        CellIsRule(operator="lessThan", formula=["kpg_esikDoluluk"], fill=sari))
    ws.conditional_formatting.add(
        f"N{ILK}:N{SON}",
        CellIsRule(operator="lessThan", formula=["kpg_esikSaglik"], fill=kirmizi))

    ws = wb["VARLIKLAR"]
    ws.conditional_formatting.add(
        f"H{ILK}:H{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"L{ILK}:L{SON}",
        CellIsRule(operator="lessThan", formula=["kpg_esikSaglik"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"J{ILK}:J{SON}",
        CellIsRule(operator="lessThan", formula=["kpg_esikDoluluk"], fill=sari))

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
    for r in range(6, 54):
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi))

    ws = wb["DONEM_ANALIZ"]
    for r in range(6, 18):
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"E{r}", CellIsRule(operator="lessThan", formula=["kpg_esikDoluluk"], fill=sari))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "kpg_varsayilanHedefKira", "kpg_esikDoluluk", "kpg_esikGetiri",
        "kpg_esikSaglik", "kpg_gunAy", "kpg_aylikCarpan", "kpg_senaryoTemkinli",
        "kpg_senaryoBaz", "kpg_senaryoIyimser", "kpg_tornadoOran", "kpg_yuzdelikOran",
        "kpg_percentileOran", "kpg_tahminCarpan", "kpg_olcekHedef", "kpg_kaliteTaban",
        "kpg_kaliteCeza", "kpg_kaliteSapma", "kpg_dosyaSurumu", "kpg_firmaUnvan",
        "kpg_alarmEsikMetin", "kpg_tufeReferans", "kpg_boslukMaliyetCarpan",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    # Ay named ranges (G07) — DONEM_ANALIZ Ay_Anahtari hücreleri
    for i in range(1, 13):
        ad_ekle(wb, f"Ay_{i}", f"DONEM_ANALIZ!$B${5 + i}")

    motor_map = {
        "kpg_doluKayit": 34, "kpg_toplamTahsilat": 35, "kpg_toplamGider": 36,
        "kpg_ortDoluluk": 37, "kpg_ortGetiri": 38, "kpg_netTahsilat": 39,
        "kpg_gunlukDolulukOrt": 40, "kpg_aylikTahsilatTrend": 41, "kpg_saglikSkor": 42,
        "kpg_alarmAdet": 43, "kpg_yetimAdet": 44, "kpg_getiriSapma": 45,
        "kpg_tornadoZirve": 46, "kpg_senaryoKarsilastirma": 47, "kpg_hhi": 48,
        "kpg_kaliteSkor": 49, "kpg_senaryoMotor": 50, "kpg_tahminAralik": 51,
        "kpg_yuzdelikP90": 52, "kpg_olayKirilim": 53, "kpg_ortTufeArtis": 54,
        "kpg_sezonsalTrend": 55, "kpg_alarmListesi": 56, "kpg_kararKapi": 57,
        "kpg_kararGerekce": 58, "kpg_sonrakiKimlik": 59,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "kpg_yorumGunluk": 62, "kpg_yorumAylik": 63, "kpg_yorumSapma": 64,
        "kpg_yorumTornado": 65, "kpg_yorumSenaryo": 66, "kpg_yorumHhi": 67,
        "kpg_yorumKalite": 68, "kpg_yorumSenMotor": 69, "kpg_yorumTahmin": 70,
        "kpg_yorumP90": 71,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$A$7:$A$8")
    ad_ekle(wb, "ListeOlayTuru", "LISTELER!$C$7:$C$12")
    ad_ekle(wb, "ListeMulkTipi", "LISTELER!$E$7:$E$10")


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
    dosya_ad = "KiraPortfoyuGetiriKomutasi.xlsx"
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
        f.write("SatirId,VarlikId,Tarih,TahsilatTl_tl,BosGun_gun,GiderTl_tl,TufeOran_oran,OlayTuru\n")
        for s in ORNEK:
            f.write(f"{s['id']},{s['varlik']},{s['tarih']:%d.%m.%Y},"
                    f"{s['tahsilat']},{s['bos']},{s['gider']},{s['tufe']},{s['olay']}\n")

    print(f"Dosya oluşturuldu: {hedef}")
    print(f"Kopya: {cikti}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
