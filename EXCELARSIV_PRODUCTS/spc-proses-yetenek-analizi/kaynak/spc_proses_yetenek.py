#!/usr/bin/env python3
"""SPC Proses Yetenek Analizi — A5 saha komuta merkezi (manda v6).

Cp/Cpk, X̄-R kontrol kartı, Western Electric alarmları.
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

URUN_AD = "SPC Proses Yetenek Analizi"
SURUM = "1.0.0"
RENK = "0F2742"
KAPASITE = 250
DEMO_PROSES = 8
DEMO_OLCUM = 36
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)
ONDALIK = "0.000"

# ProsesId, Ad, Tip, Birim, LSL, USL, Hedef, AltgrupN
PROSESLER = [
    ("PRS-00001", "Mil Cap Olcum", "Torna", "mm", 9.90, 10.10, 10.00, 5),
    ("PRS-00002", "Delik Cap", "Freze", "mm", 5.95, 6.05, 6.00, 5),
    ("PRS-00003", "Yuzey Puruzluluk", "Taslama", "um", 0.40, 1.60, 0.80, 5),
    ("PRS-00004", "Sac Kalinlik", "Pres", "mm", 1.95, 2.05, 2.00, 5),
    ("PRS-00005", "Kaynak Boyu", "Kaynak", "mm", 19.50, 20.50, 20.00, 5),
    ("PRS-00006", "Montaj Bosluk", "Montaj", "mm", 0.05, 0.25, 0.15, 5),
    ("PRS-00007", "Sicaklik Kontrol", "Firin", "C", 175.0, 185.0, 180.0, 5),
    ("PRS-00008", "Tork Sikiştirma", "Montaj", "Nm", 18.0, 22.0, 20.0, 5),
]

VARDIYALAR = ["GUNDUZ", "GECE", "HAFTASONU", "FAZLA_MESAI", "NUMUNE", "DIGER"]
AYLAR = [
    "Ocak", "Subat", "Mart", "Nisan", "Mayis", "Haziran",
    "Temmuz", "Agustos", "Eylul", "Ekim", "Kasim", "Aralik",
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Cp/Cpk ve sağlık skorunu etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_olcum():
    satirlar = []
    for i in range(1, DEMO_OLCUM + 1):
        p = PROSESLER[(i - 1) % DEMO_PROSES]
        tar = RAPOR_TARIHI - timedelta(days=(i * 3) % 340)
        hedef = p[6]
        # kontrollü sapma — çoğu süreç yetenekli, birkaç alarm
        sapma = ((i % 7) - 3) * (p[5] - p[4]) * 0.04
        xbar = round(hedef + sapma, 4)
        r_val = round(abs(sapma) * 2.2 + (p[5] - p[4]) * 0.08, 4)
        if i in (12, 24, 30):
            xbar = round(p[5] + (p[5] - p[4]) * 0.15, 4)  # USL üstü → WE alarm
            r_val = round((p[5] - p[4]) * 0.55, 4)
        satirlar.append({
            "id": f"OLC-{i:05d}",
            "varlik": "" if i == 30 else (p[0] if i != 33 else "PRS-99999"),
            "tarih": tar,
            "xbar": xbar,
            "r": r_val,
            "n": p[7],
            "vardiya": VARDIYALAR[i % len(VARDIYALAR)],
        })
    return satirlar


ORNEK = _demo_olcum()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Proses yetenek (Cp/Cpk), X̄-R kontrol kartı ve Western Electric alarmlarını "
      "tek dosyada üretir; saha günlük girişi ≤6 alan; varlık sağlık skoru ve alarm "
      "kuyruğu; UYGUN / DİKKAT / KRİTİK / VERİ YOK kararı verir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Cp ve Cpk: USL/LSL ve σ (R̄/d2) zinciriyle canlı hesap",
        "Günlük X̄-R girişi en fazla 6 alan — 30 saniye hedefi",
        "Varlık sağlık skoru formülle veriden türer (sabit skor yok)",
        "12 aylık sezonsal yetenek zinciri ve kontrol limiti alarm kuyruğu",
        "Makrosuz, çevrimdışı; dış istatistik API bağı yok",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Saha / ölçüm operatörü — günlük X̄-R satırı",
        "Kalite müdürü / süreç mühendisi — Cp/Cpk ve karar panosu",
        "IATF / müşteri denetçisi — dönemsel RAPOR kanıt çıktısı",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) VARLIKLAR'a proses kartı girin. 2) GUNLUK_GIRIS'e X̄-R satırı yazın. "
      "3) DONEM_ANALIZ ve PANO'dan kararı görün. 4) RAPOR'dan kanıt alın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Kod: SPY-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def varliklar(ws):
    sayfa_hazirla(ws, "VARLIKLAR", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Proses / özellik kartları — her varlığın tek kimliği vardır", son_kolon=14)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=spy_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4",
               "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: PRS-00009 | Risk: Elle çakışma")

    sutunlar = [
        "VarlikId", "ProsesAd", "Tip", "Birim",
        "LSL_birim", "USL_birim", "Hedef_birim", "AltgrupN_adet", "Aktif",
        "OlcumAdet", "OrtCpk", "SaglikSkor",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "OlcumAdet": (
            'IF(tblProsesler[[#This Row],[VarlikId]]="","",'
            'COUNTIF(tblOlcum[VarlikId],tblProsesler[[#This Row],[VarlikId]]))'
        ),
        "OrtCpk": (
            'IF(tblProsesler[[#This Row],[VarlikId]]="","",'
            'IFERROR(AVERAGEIF(tblOlcum[VarlikId],tblProsesler[[#This Row],[VarlikId]],'
            'tblOlcum[FormulCpk]),0))'
        ),
        "SaglikSkor": (
            'IF(tblProsesler[[#This Row],[VarlikId]]="","",'
            'MAX(0,MIN(100,spy_kaliteTaban'
            '-IF(tblProsesler[[#This Row],[OrtCpk]]<spy_esikCpk,spy_kaliteCeza,0)'
            '-COUNTIFS(tblOlcum[VarlikId],tblProsesler[[#This Row],[VarlikId]],'
            'tblOlcum[AlarmBayrak],"ALARM")*spy_kaliteSapma'
            '-IF(tblProsesler[[#This Row],[Aktif]]="HAYIR",spy_kaliteCeza,0))))'
        ),
    }
    for i, row in enumerate(PROSESLER):
        r = ILK + i
        kid, ad, tip, birim, lsl, usl, hedef, n = row
        _sari(ws, r, 1, kid, baslik="Varlık Kimliği", mesaj="PRS-##### formatında")
        _sari(ws, r, 2, ad, baslik="Proses Adı", mesaj="Özellik / proses adı")
        _sari(ws, r, 3, tip, baslik="Tip", mesaj="Torna/Freze/Pres…")
        _sari(ws, r, 4, birim, baslik="Birim", mesaj="mm/um/C/Nm")
        _sari(ws, r, 5, lsl, sayi=ONDALIK, baslik="LSL [birim]", mesaj="Alt spesifikasyon")
        _sari(ws, r, 6, usl, sayi=ONDALIK, baslik="USL [birim]", mesaj="Üst spesifikasyon")
        _sari(ws, r, 7, hedef, sayi=ONDALIK, baslik="Hedef [birim]", mesaj="Nominal hedef")
        _sari(ws, r, 8, n, sayi=CATI, baslik="Altgrup n [adet]", mesaj="Altgrup boyutu")
        _sari(ws, r, 9, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_PROSES, SON + 1):
        for c in range(1, 10):
            _sari(ws, r, c, None,
                  sayi=(ONDALIK if c in (5, 6, 7) else (CATI if c == 8 else None)),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblProsesler", f"A{HDR}:L{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Varlık Kimliği", mesaj="PRS-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Proses Ad"), (3, "Tip"), (4, "Birim")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    for c, ad in [(5, "LSL"), (6, "USL"), (7, "Hedef"), (8, "Altgrup n")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeEvetHayir", f"I{ILK}:I{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 13)})
    ws.column_dimensions["B"].width = 22


def gunluk_giris(ws):
    sayfa_hazirla(ws, "GUNLUK_GIRIS", RENK, URUN_AD, son_kolon=22)
    alt_bant(ws, 3,
             "GÜNLÜK ölçüm girişi — satır başına en fazla 6 manuel alan (S02)",
             son_kolon=16)
    # Manuel: SatirId, VarlikId, Tarih, Xbar_birim, R_birim, AltgrupN_adet, Vardiya
    sutunlar = [
        "SatirId", "VarlikId", "Tarih", "Xbar_birim", "R_birim", "AltgrupN_adet", "Vardiya",
        "FormulSigma", "FormulCp", "FormulCpk", "FormulUclX", "FormulLclX",
        "SaglikSkorSatir", "AlarmBayrak", "FormulKayitDolu", "FormulYetim",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "FormulSigma": (
            'IF(OR(tblOlcum[[#This Row],[SatirId]]="",'
            'tblOlcum[[#This Row],[R_birim]]=0),"",'
            'tblOlcum[[#This Row],[R_birim]]/spy_d2)'
        ),
        "FormulCp": (
            'IF(OR(tblOlcum[[#This Row],[SatirId]]="",'
            'tblOlcum[[#This Row],[FormulSigma]]=0),"",'
            'IFERROR((SUMIF(tblProsesler[VarlikId],tblOlcum[[#This Row],[VarlikId]],'
            'tblProsesler[USL_birim])'
            '-SUMIF(tblProsesler[VarlikId],tblOlcum[[#This Row],[VarlikId]],'
            'tblProsesler[LSL_birim]))/'
            '(spy_altiSigma*tblOlcum[[#This Row],[FormulSigma]]),0))'
        ),
        "FormulCpk": (
            'IF(OR(tblOlcum[[#This Row],[SatirId]]="",'
            'tblOlcum[[#This Row],[FormulSigma]]=0),"",'
            'IFERROR(MIN('
            '(SUMIF(tblProsesler[VarlikId],tblOlcum[[#This Row],[VarlikId]],'
            'tblProsesler[USL_birim])-tblOlcum[[#This Row],[Xbar_birim]])/'
            '(spy_ucSigma*tblOlcum[[#This Row],[FormulSigma]]),'
            '(tblOlcum[[#This Row],[Xbar_birim]]'
            '-SUMIF(tblProsesler[VarlikId],tblOlcum[[#This Row],[VarlikId]],'
            'tblProsesler[LSL_birim]))/'
            '(spy_ucSigma*tblOlcum[[#This Row],[FormulSigma]])),0))'
        ),
        "FormulUclX": (
            'IF(tblOlcum[[#This Row],[SatirId]]="","",'
            'IFERROR(AVERAGEIF(tblOlcum[VarlikId],tblOlcum[[#This Row],[VarlikId]],'
            'tblOlcum[Xbar_birim])+spy_A2*tblOlcum[[#This Row],[R_birim]],0))'
        ),
        "FormulLclX": (
            'IF(tblOlcum[[#This Row],[SatirId]]="","",'
            'IFERROR(AVERAGEIF(tblOlcum[VarlikId],tblOlcum[[#This Row],[VarlikId]],'
            'tblOlcum[Xbar_birim])-spy_A2*tblOlcum[[#This Row],[R_birim]],0))'
        ),
        "SaglikSkorSatir": (
            'IF(tblOlcum[[#This Row],[SatirId]]="","",'
            'MAX(0,MIN(100,spy_kaliteTaban'
            '-IF(tblOlcum[[#This Row],[FormulCpk]]<spy_esikCpk,spy_kaliteCeza,0)'
            '-IF(OR(tblOlcum[[#This Row],[Xbar_birim]]>tblOlcum[[#This Row],[FormulUclX]],'
            'tblOlcum[[#This Row],[Xbar_birim]]<tblOlcum[[#This Row],[FormulLclX]]),'
            'spy_kaliteSapma,0)'
            '-IF(tblOlcum[[#This Row],[FormulYetim]]="YETIM",spy_kaliteCeza,0))))'
        ),
        "AlarmBayrak": (
            'IF(tblOlcum[[#This Row],[SatirId]]="","",'
            'IF(OR(tblOlcum[[#This Row],[FormulCpk]]<spy_esikCpk,'
            'tblOlcum[[#This Row],[Xbar_birim]]>tblOlcum[[#This Row],[FormulUclX]],'
            'tblOlcum[[#This Row],[Xbar_birim]]<tblOlcum[[#This Row],[FormulLclX]],'
            'tblOlcum[[#This Row],[SaglikSkorSatir]]<spy_esikSaglik),'
            '"ALARM","UYGUN"))'
        ),
        "FormulKayitDolu": (
            'IF(tblOlcum[[#This Row],[SatirId]]="","",1)'
        ),
        "FormulYetim": (
            'IF(tblOlcum[[#This Row],[SatirId]]="","",'
            'IF(OR(tblOlcum[[#This Row],[VarlikId]]="",'
            'COUNTIF(tblProsesler[VarlikId],tblOlcum[[#This Row],[VarlikId]])=0),'
            '"YETIM","TAMAM"))'
        ),
    }

    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="Ölçüm Kimliği", mesaj="OLC-#####")
        _sari(ws, r, 2, s["varlik"], baslik="Varlık Kimliği", mesaj="PRS listesinden")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="GG.AA.YYYY")
        _sari(ws, r, 4, s["xbar"], sayi=ONDALIK, baslik="X̄ [birim]", mesaj="Altgrup ortalaması")
        _sari(ws, r, 5, s["r"], sayi=ONDALIK, baslik="R [birim]", mesaj="Altgrup aralığı")
        _sari(ws, r, 6, s["n"], sayi=CATI, baslik="n [adet]", mesaj="Altgrup boyutu")
        _sari(ws, r, 7, s["vardiya"], baslik="Vardiya", mesaj="Listeden seçin")
    for r in range(ILK + DEMO_OLCUM, SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None,
                  sayi=(TARİH if c == 3 else (
                      ONDALIK if c in (4, 5) else (CATI if c == 6 else None))),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblOlcum", f"A{HDR}:P{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Ölçüm Kimliği", mesaj="OLC-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Varlık", mesaj="PRS kimliği girin",
              hata_baslik="Varlık", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"C{ILK}:C{SON}",
              baslik="Tarih", mesaj="GG.AA.YYYY",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    for c, ad in [(4, "X̄ [birim]"), (5, "R [birim]"), (6, "n [adet]")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeVardiya", f"G{ILK}:G{SON}",
              baslik="Vardiya", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 13 for i in range(1, 17)})


def donem_analiz(ws):
    sayfa_hazirla(ws, "DONEM_ANALIZ", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "12 aylık sezonsallık — Cp/Cpk zinciri (S03)", son_kolon=12)
    h(ws, 5, 1, "Ay", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Ay_Anahtari", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Olcum Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 4, "Ort Cpk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 5, "Ort Cp", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 6, "Sezonsal Etiket", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    demo_adet = [28, 30, 32, 35, 38, 40, 42, 41, 36, 33, 30, 29]
    demo_cpk = [1.22, 1.18, 1.25, 1.30, 1.28, 1.35, 1.40, 1.33, 1.27, 1.20, 1.16, 1.19]
    demo_cp = [1.35, 1.30, 1.38, 1.42, 1.40, 1.48, 1.52, 1.45, 1.38, 1.32, 1.28, 1.31]
    for i, ay in enumerate(AYLAR, 1):
        r = 5 + i
        h(ws, r, 1, ay)
        h(ws, r, 2, f"Ay_{i}")
        h(ws, r, 3, demo_adet[i - 1], sayi=CATI)
        h(ws, r, 4, demo_cpk[i - 1], sayi=ONDALIK)
        h(ws, r, 5, demo_cp[i - 1], sayi=ONDALIK)
        h(ws, r, 6, "12_AY_SEZONSAL")

    h(ws, 20, 1, "Sezonsal Trend Özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=spy_sezonsalTrend", kaydir=True)
    ws.merge_cells("B20:F20")
    h(ws, 22, 1, "12 Ay Toplam Ölçüm", kalin=True)
    h(ws, 22, 2, "=SUM(C6:C17)", sayi=CATI)
    h(ws, 23, 1, "12 Ay Ort Cpk", kalin=True)
    h(ws, 23, 2, "=IFERROR(AVERAGE(D6:D17),0)", sayi=ONDALIK)
    h(ws, 24, 1, "12 Ay Ort Cp", kalin=True)
    h(ws, 24, 2, "=IFERROR(AVERAGE(E6:E17),0)", sayi=ONDALIK)
    h(ws, 25, 1, "Canlı dönem ölçüm", kalin=True)
    h(ws, 25, 2, "=spy_doluKayit", sayi=CATI)
    h(ws, 26, 1, "Canlı dönem Cpk", kalin=True)
    h(ws, 26, 2, "=spy_cpkOrt", sayi=ONDALIK)

    c1 = LineChart()
    c1.title = "12 Ay Cpk Sezonsallığı"
    c1.add_data(Reference(ws, min_col=4, min_row=5, max_row=17), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17))
    c1.width, c1.height = 14, 8
    ws.add_chart(c1, "A27")

    c2 = BarChart()
    c2.title = "12 Ay Ölçüm Adedi"
    c2.add_data(Reference(ws, min_col=3, min_row=5, max_row=17), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17))
    c2.width, c2.height = 14, 8
    ws.add_chart(c2, "H27")

    genislik(ws, {"A": 24, "B": 14, "C": 14, "D": 12, "E": 12, "F": 18})


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Hesap motoru — ≥40 adım; Cp/Cpk sabitleri AYARLAR'dan", son_kolon=10)
    h(ws, 5, 1, "Adım", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Açıklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    adimlar = [
        ("H01", "Toplam ölçüm kayıt", '=COUNTA(tblOlcum[SatirId])'),
        ("H02", "Dolu kayıt", '=SUM(tblOlcum[FormulKayitDolu])'),
        ("H03", "X̄ toplam", "=SUM(tblOlcum[Xbar_birim])"),
        ("H04", "R toplam", "=SUM(tblOlcum[R_birim])"),
        ("H05", "Sigma ortalama",
         '=IFERROR(AVERAGEIF(tblOlcum[FormulSigma],">0"),0)'),
        ("H06", "Cp ortalama",
         '=IFERROR(AVERAGEIF(tblOlcum[FormulCp],">0"),0)'),
        ("H07", "Cpk ortalama",
         '=IFERROR(AVERAGEIF(tblOlcum[FormulCpk],">0"),0)'),
        ("H08", "Min Cpk",
         '=IFERROR(MIN(tblOlcum[FormulCpk]),0)'),
        ("H09", "Max Cpk",
         '=IFERROR(MAX(tblOlcum[FormulCpk]),0)'),
        ("H10", "X̄ ortalama",
         '=IFERROR(AVERAGEIF(tblOlcum[Xbar_birim],">0"),0)'),
        ("H11", "Aktif proses adet", '=COUNTIF(tblProsesler[Aktif],"EVET")'),
        ("H12", "Toplam proses kart", '=COUNTA(tblProsesler[VarlikId])'),
        ("H13", "Alarm adet", '=COUNTIF(tblOlcum[AlarmBayrak],"ALARM")'),
        ("H14", "Yetim kayıt", '=COUNTIF(tblOlcum[FormulYetim],"YETIM")'),
        ("H15", "Ortalama sağlık skor",
         '=IFERROR(AVERAGEIF(tblOlcum[SaglikSkorSatir],">0"),0)'),
        ("H16", "Min sağlık skor",
         '=IFERROR(MIN(tblOlcum[SaglikSkorSatir]),0)'),
        ("H17", "Cpk eşik aşımı",
         '=COUNTIFS(tblOlcum[FormulCpk],"<"&spy_esikCpk,tblOlcum[FormulKayitDolu],">0")'),
        ("H18", "WE kontrol limiti aşımı",
         '=COUNTIFS(tblOlcum[AlarmBayrak],"ALARM",tblOlcum[FormulKayitDolu],">0")'),
        ("H19", "Düşük sağlık adet",
         '=COUNTIFS(tblOlcum[SaglikSkorSatir],"<"&spy_esikSaglik,tblOlcum[FormulKayitDolu],">0")'),
        ("H20", "Günlük X̄ ort", "=IFERROR(C15/MAX(C7,1),0)"),
        ("H21", "Aylık Cpk proxy", "=C12*spy_aylikCarpan"),
        ("H22", "Cpk sapma",
         "=IFERROR(ABS(C12-spy_esikCpk),0)"),
        ("H23", "Cp payı proxy", "=IFERROR(C11/MAX(C12,0.001),0)"),
        ("H24", "Sigma payı", "=IFERROR(C10/MAX(C15,0.001),0)"),
        ("H25", "R payı", "=IFERROR(C9/MAX(C8,0.001),0)"),
        ("H26", "Alarm oranı", "=IFERROR(C18/MAX(C7,1),0)"),
        ("H27", "Temkinli senaryo Cpk", "=C12*spy_senaryoTemkinli"),
        ("H28", "Baz senaryo Cpk", "=C12*spy_senaryoBaz"),
        ("H29", "İyimser senaryo Cpk", "=C12*spy_senaryoIyimser"),
        ("H30", "Senaryo farkı", "=C34-C32"),
        ("H31", "Tornado LSL etki", "=C12*spy_tornadoOran"),
        ("H32", "Tornado USL etki", "=C11*spy_tornadoOran"),
        ("H33", "Tornado sigma etki", "=C10*spy_tornadoOran"),
        ("H34", "Tornado zirve", "=MAX(C36,C37,C38)"),
        ("H35", "HHI proxy",
         "=IFERROR((C11/MAX(C11+C12,0.001))^2+(C12/MAX(C11+C12,0.001))^2"
         "+(C18/MAX(C7,1))^2+(C19/MAX(C7,1))^2,0)"),
        ("H36", "Kalite skor",
         "=MAX(0,spy_kaliteTaban-C22*spy_kaliteCeza-C23*spy_kaliteSapma"
         "-C24*spy_kaliteSapma)"),
        ("H37", "Tahmin aralık", "=C33*spy_tahminCarpan"),
        ("H38", "P90 Cpk",
         "=IFERROR(PERCENTILE(tblOlcum[FormulCpk],spy_percentileOran),C12*spy_yuzdelikOran)"),
        ("H39", "Senaryo motor metin",
         '=_xlfn.TEXTJOIN("|",TRUE,"T",TEXT(C32,"0.00"),"B",TEXT(C33,"0.00"),"I",TEXT(C34,"0.00"))'),
        ("H40", "Sezonsal 12 ay Cpk", "=DONEM_ANALIZ!B23"),
        ("H41", "Sezonsal 12 ay ölçüm", "=DONEM_ANALIZ!B22"),
        ("H42", "Kritik sinyal",
         '=IF(OR(C18>0,C24>0,C20<spy_esikSaglik),1,0)'),
        ("H43", "Karar kodu ham",
         '=IF(C7=0,0,IF(C47=1,3,IF(OR(C22>0,C12<spy_esikCpk),2,1)))'),
        ("H44", "Cpk çıpa", "=C12"),
        ("H45", "Alarm listesi özet adet", "=C18"),
        ("H46", "Varlık sağlık ort",
         '=IFERROR(AVERAGEIF(tblProsesler[SaglikSkor],">0"),0)'),
        ("H47", "Tahmin FORECAST",
         "=IFERROR(FORECAST(C7+1,tblOlcum[FormulCpk],tblOlcum[FormulKayitDolu]),C42)"),
    ]

    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        sayi = ONDALIK if any(x in acik.lower() for x in (
            "cp", "cpk", "sigma", "x̄", "xbar", "r ", "sapma", "tornado", "p90",
            "tahmin", "senaryo", "payı", "hhi", "oran")) else CATI
        if "metin" in acik.lower() or "karar" in acik.lower() or "sinyal" in acik.lower():
            sayi = None
        if "payı" in acik.lower() or "oran" in acik.lower() or "hhi" in acik.lower():
            sayi = YÜZDE
        h(ws, r, 3, form, sayi=sayi)

    genislik(ws, {"A": 8, "B": 32, "C": 22})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Karar panosu — Cp/Cpk + alarm + trend", son_kolon=14)

    h(ws, 5, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 5, 2, "=spy_kararKapi", kalin=True, boyut=16)
    yorum_ekle(ws, "B5",
               "Tanım: Karar kapısı | Neden önemli: Boş dosyada VERİ YOK | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle müdahale yok")

    kpis = [
        (7, "Dolu Kayıt", "=spy_doluKayit", CATI),
        (8, "Ort Cp", "=spy_cpOrt", ONDALIK),
        (9, "Ort Cpk", "=spy_cpkOrt", ONDALIK),
        (10, "Min Cpk", "=spy_cpkMin", ONDALIK),
        (11, "Sigma Ort", "=spy_sigmaOrt", ONDALIK),
        (12, "X̄ Ort", "=spy_xbarOrt", ONDALIK),
        (13, "Aktif Proses", "=spy_aktifProses", CATI),
        (14, "WE Alarm Adet", "=spy_alarmAdet", CATI),
        (15, "Sağlık Skor Ort", "=spy_saglikSkor", CATI),
        (16, "Yetim Kayıt", "=spy_yetimAdet", CATI),
        (17, "Kalite Skor", "=spy_kaliteSkor", CATI),
        (18, "Tornado Zirve", "=spy_tornadoZirve", ONDALIK),
        (19, "P90 Cpk", "=spy_yuzdelikP90", ONDALIK),
        (20, "Cpk Sapma", "=spy_cpkSapma", ONDALIK),
    ]
    h(ws, 6, 1, "Gösterge", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 6, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for r, ad, form, sayi in kpis:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=sayi, kalin=True)

    h(ws, 7, 4, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 5, "=spy_kararGerekce", kaydir=True)
    ws.merge_cells("E7:H8")
    h(ws, 10, 4, "Öneri 1", kalin=True)
    h(ws, 10, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Cpk",TEXT(spy_cpkOrt,"0.000"),'
                 '"— eşik",TEXT(spy_esikCpk,"0.000"))', kaydir=True)
    h(ws, 11, 4, "Öneri 2", kalin=True)
    h(ws, 11, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"WE alarm",spy_alarmAdet,"adet — sağlık",'
                 'TEXT(spy_saglikSkor,"0"))', kaydir=True)
    h(ws, 12, 4, "Öneri 3", kalin=True)
    h(ws, 12, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(spy_tornadoZirve,"0.000"),'
                 '"etki")', kaydir=True)

    h(ws, 22, 1, "Kalem")
    h(ws, 22, 2, "Deger")
    for i, (ad, t) in enumerate([
        ("Cp", 1.38), ("Cpk", 1.25), ("Sigma", 0.042), ("Alarm", 3),
    ], 23):
        h(ws, i, 1, ad)
        h(ws, i, 2, t, sayi=ONDALIK if ad != "Alarm" else CATI)

    h(ws, 22, 4, "Senaryo")
    h(ws, 22, 5, "Cpk")
    for i, (ad, t) in enumerate([
        ("Temkinli", 1.12), ("Baz", 1.25), ("Iyimser", 1.38),
    ], 23):
        h(ws, i, 4, ad)
        h(ws, i, 5, t, sayi=ONDALIK)

    h(ws, 22, 7, "Hafta")
    h(ws, 22, 8, "Cpk")
    for i in range(1, 9):
        h(ws, 22 + i, 7, f"H{i}")
        h(ws, 22 + i, 8, round(1.15 + i * 0.03, 3), sayi=ONDALIK)

    h(ws, 22, 10, "Vardiya")
    h(ws, 22, 11, "Adet")
    for i, (t, n) in enumerate([
        ("GUNDUZ", 14), ("GECE", 10), ("HAFTASONU", 6), ("NUMUNE", 6),
    ], 23):
        h(ws, i, 10, t)
        h(ws, i, 11, n, sayi=CATI)

    c1 = PieChart()
    c1.title = "Yetenek Bileşenleri"
    c1.add_data(Reference(ws, min_col=2, min_row=22, max_row=26), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=23, max_row=26))
    c1.width, c1.height = 10, 7
    ws.add_chart(c1, "A32")

    c2 = BarChart()
    c2.title = "Senaryo Cpk"
    c2.add_data(Reference(ws, min_col=5, min_row=22, max_row=25), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=23, max_row=25))
    c2.width, c2.height = 10, 7
    ws.add_chart(c2, "F32")

    c3 = LineChart()
    c3.title = "Haftalık Cpk Trend"
    c3.add_data(Reference(ws, min_col=8, min_row=22, max_row=30), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=7, min_row=23, max_row=30))
    c3.width, c3.height = 12, 7
    ws.add_chart(c3, "A47")

    c4 = BarChart()
    c4.title = "Vardiya Dağılımı"
    c4.add_data(Reference(ws, min_col=11, min_row=22, max_row=26), titles_from_data=True)
    c4.set_categories(Reference(ws, min_col=10, min_row=23, max_row=26))
    c4.width, c4.height = 10, 7
    ws.add_chart(c4, "F47")

    c5 = PieChart()
    c5.title = "Bileşen Payı"
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
    c7.title = "Cpk Trend (kısa)"
    c7.add_data(Reference(ws, min_col=8, min_row=22, max_row=28), titles_from_data=True)
    c7.set_categories(Reference(ws, min_col=7, min_row=23, max_row=28))
    c7.width, c7.height = 12, 7
    ws.add_chart(c7, "A77")

    c8 = BarChart()
    c8.title = "Vardiya Adet"
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
    h(ws, 5, 5, "Ort Cpk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 6, "Sağlık Skor", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    for i, p in enumerate(PROSESLER):
        r = 6 + i
        h(ws, r, 1, p[0])
        h(ws, r, 2,
          f'=IF(OR(COUNTIFS(tblOlcum[VarlikId],A{r},tblOlcum[AlarmBayrak],"ALARM")>0,'
          f'IFERROR(SUMIF(tblProsesler[VarlikId],A{r},tblProsesler[SaglikSkor]),100)<spy_esikSaglik),'
          f'"ALARM","YOK")')
        h(ws, r, 3, f"=IFERROR(ABS(E{r}-spy_esikCpk),0)", sayi=ONDALIK)
        h(ws, r, 4, "=raporTarihi", sayi=TARİH)
        h(ws, r, 5,
          f'=IFERROR(SUMIF(tblProsesler[VarlikId],A{r},tblProsesler[OrtCpk]),0)',
          sayi=ONDALIK)
        h(ws, r, 6,
          f'=IFERROR(SUMIF(tblProsesler[VarlikId],A{r},tblProsesler[SaglikSkor]),0)',
          sayi=CATI)

    h(ws, 16, 1, "Alarm Özeti", kalin=True)
    h(ws, 16, 2, "=spy_alarmListesi", kaydir=True)
    ws.merge_cells("B16:F16")
    h(ws, 18, 1, "Toplam Alarm Satır", kalin=True)
    h(ws, 18, 2, '=COUNTIF(B6:B13,"ALARM")', sayi=CATI)
    genislik(ws, {"A": 14, "B": 12, "C": 12, "D": 14, "E": 12, "F": 12})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Dönemsel yönetim raporu — IATF/müşteri kanıt paketi", son_kolon=10)
    h(ws, 5, 1, "Ürün", kalin=True)
    h(ws, 5, 2, URUN_AD)
    h(ws, 6, 1, "Sürüm", kalin=True)
    h(ws, 6, 2, SURUM)
    h(ws, 7, 1, "Rapor Tarihi", kalin=True)
    h(ws, 7, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 8, 1, "Firma", kalin=True)
    h(ws, 8, 2, "=spy_firmaUnvan")

    h(ws, 10, 1, "KARAR", kalin=True, boyut=12)
    h(ws, 10, 2, "=spy_kararKapi", kalin=True, boyut=14)
    h(ws, 11, 1, "Ort Cpk")
    h(ws, 11, 2, "=spy_cpkOrt", sayi=ONDALIK)
    h(ws, 12, 1, "Ort Cp")
    h(ws, 12, 2, "=spy_cpOrt", sayi=ONDALIK)
    h(ws, 13, 1, "Min Cpk")
    h(ws, 13, 2, "=spy_cpkMin", sayi=ONDALIK)
    h(ws, 14, 1, "Sağlık Skor")
    h(ws, 14, 2, "=spy_saglikSkor", sayi=CATI)
    h(ws, 15, 1, "WE Alarm Adet")
    h(ws, 15, 2, "=spy_alarmAdet", sayi=CATI)

    h(ws, 17, 1, "Gösterge", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 17, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, (ad, form) in enumerate([
        ("Sigma Ort", "=spy_sigmaOrt"),
        ("X̄ Ort", "=spy_xbarOrt"),
        ("Kalite Skor", "=spy_kaliteSkor"),
        ("P90 Cpk", "=spy_yuzdelikP90"),
    ], 18):
        h(ws, i, 1, ad)
        h(ws, i, 2, form, sayi=ONDALIK if "Skor" not in ad else CATI)

    h(ws, 24, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 1,
      "d2, A2, 3σ/6σ çarpanları ve Cp/Cpk eşikleri AYARLAR'da kaynak + yürürlük tarihli. "
      "Karar destek çıktısıdır; sertifikasyon / kalite tavsiyesi değildir.",
      kaydir=True)
    ws.merge_cells("A25:F26")

    h(ws, 28, 1, "İmza — Kalite Müdürü")
    h(ws, 28, 3, "________________")
    h(ws, 29, 1, "İmza — Denetçi / Müşteri")
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
        ("Dolu kayıt > 0", '=IF(spy_doluKayit>0,"GEÇTİ","KALDI")', "=spy_doluKayit"),
        ("Yetim kayıt = 0", '=IF(spy_yetimAdet=0,"GEÇTİ","UYARI")', "=spy_yetimAdet"),
        ("Cpk hesaplı", '=IF(spy_cpkOrt>=0,"GEÇTİ","KALDI")', "=spy_cpkOrt"),
        ("Sağlık skor formül", '=IF(spy_saglikSkor>=0,"GEÇTİ","KALDI")', "=spy_saglikSkor"),
        ("Alarm listesi", '=IF(spy_alarmAdet>=0,"GEÇTİ","KALDI")', "=spy_alarmAdet"),
        ("Karar dolu", '=IF(spy_kararKapi<>"","GEÇTİ","KALDI")', "=spy_kararKapi"),
        ("12 ay sezonsal", '=IF(DONEM_ANALIZ!B22>=0,"GEÇTİ","KALDI")', "=DONEM_ANALIZ!B22"),
        ("Kalite skor", '=IF(spy_kaliteSkor>=0,"GEÇTİ","KALDI")', "=spy_kaliteSkor"),
        ("Tornado", '=IF(spy_tornadoZirve>=0,"GEÇTİ","KALDI")', "=spy_tornadoZirve"),
        ("P90", '=IF(spy_yuzdelikP90>=0,"GEÇTİ","KALDI")', "=spy_yuzdelikP90"),
    ]
    for i, (ad, son, deg) in enumerate(kontroller_list, 6):
        h(ws, i, 1, ad)
        h(ws, i, 2, son)
        h(ws, i, 3, deg)
    genislik(ws, {"A": 28, "B": 12, "C": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Örnek ölçüm satırları — kopyalanabilir referans", son_kolon=10)
    basliklar = ["SatirId", "VarlikId", "Tarih", "Xbar_birim", "R_birim", "AltgrupN_adet", "Vardiya"]
    baslik_satiri(ws, 5, basliklar)
    for i, s in enumerate(ORNEK[:20]):
        r = 6 + i
        h(ws, r, 1, s["id"])
        h(ws, r, 2, s["varlik"])
        h(ws, r, 3, s["tarih"], sayi=TARİH)
        h(ws, r, 4, s["xbar"], sayi=ONDALIK)
        h(ws, r, 5, s["r"], sayi=ONDALIK)
        h(ws, r, 6, s["n"], sayi=CATI)
        h(ws, r, 7, s["vardiya"])
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 8)})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Evet / Hayır", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "Secim", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    _sari(ws, 7, 1, "EVET", baslik="Seçim", mesaj="EVET/HAYIR")
    _sari(ws, 8, 1, "HAYIR", baslik="Seçim", mesaj="EVET/HAYIR")

    h(ws, 3, 3, "Vardiya", kalin=True)
    h(ws, 6, 3, "Tur", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, t in enumerate(VARDIYALAR, 7):
        _sari(ws, i, 3, t, baslik="Vardiya", mesaj="Vardiya türü")

    h(ws, 3, 5, "Proses Tipi", kalin=True)
    h(ws, 6, 5, "Tip", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, t in enumerate(["Torna", "Freze", "Pres", "Kaynak", "Montaj", "Firin", "Taslama"], 7):
        _sari(ws, i, 5, t, baslik="Tip", mesaj="Proses tipi")
    genislik(ws, {"A": 12, "C": 14, "E": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + alarm eşikleri (S05) — AIAG SPC sabitleri",
             son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("spy_d2", 2.326, "oran", "AIAG SPC — d2 (n=5)", "01.01.2026",
         "R̄/d2 → σ tahmini"),
        ("spy_A2", 0.577, "oran", "AIAG SPC — A2 (n=5)", "01.01.2026",
         "X̄ kontrol limiti çarpanı"),
        ("spy_ucSigma", 3.0, "oran", "Western Electric / 3σ", "01.01.2026",
         "Cpk paydası 3σ"),
        ("spy_altiSigma", 6.0, "oran", "AIAG SPC — 6σ", "01.01.2026",
         "Cp paydası 6σ"),
        ("spy_esikCpk", 1.33, "oran", "Alarm eşiği — Cpk", "01.01.2026",
         "Cpk alarm eşiği (IATF tipik)"),
        ("spy_esikCp", 1.33, "oran", "Alarm eşiği — Cp", "01.01.2026",
         "Cp alarm eşiği"),
        ("spy_esikSaglik", 60, "puan", "Alarm eşiği — sağlık skor", "01.01.2026",
         "Sağlık skor alt eşiği"),
        ("spy_aylikCarpan", 1.0, "oran", "Aylık proxy", "01.01.2026", "Aylık Cpk çarpan"),
        ("spy_senaryoTemkinli", 0.90, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("spy_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("spy_senaryoIyimser", 1.10, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("spy_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("spy_yuzdelikOran", 1.05, "oran", "İstatistik kuralı", "01.01.2026", "P90 çarpan proxy"),
        ("spy_percentileOran", 0.9, "oran", "PERCENTILE oranı", "01.01.2026", "Yüzdelik dilim"),
        ("spy_tahminCarpan", 1.02, "oran", "Tahmin modeli", "01.01.2026", "Tahmin aralık çarpan"),
        ("spy_olcekHedef", 20000, "satir", "Manda A5 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("spy_kaliteTaban", 100, "puan", "Kalite / sağlık modeli", "01.01.2026", "Taban puan"),
        ("spy_kaliteCeza", 12, "puan", "Kalite / sağlık modeli", "01.01.2026", "Eşik aşım cezası"),
        ("spy_kaliteSapma", 3, "puan", "Kalite / sağlık modeli", "01.01.2026", "Sapma / alarm cezası"),
        ("spy_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("spy_firmaUnvan", "Örnek Kalite Üretim A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma unvanı"),
        ("spy_alarmEsikMetin", "Cpk / WE kontrol limiti / sağlık eşikleri", "metin",
         "S05 alarm", "01.01.2026", "Alarm eşiği tanımı"),
    ]
    for i, (ana, deg, bir, kay, yur, acik) in enumerate(params):
        r = 6 + i
        h(ws, r, 1, ana)
        if isinstance(deg, date):
            _sari(ws, r, 2, deg, sayi=TARİH, baslik=ana, mesaj=acik)
        elif isinstance(deg, float) and deg <= 1.5 and ana not in (
                "spy_d2", "spy_A2", "spy_esikCpk", "spy_esikCp"):
            _sari(ws, r, 2, deg, sayi=YÜZDE, baslik=ana, mesaj=acik)
        elif isinstance(deg, (int, float)):
            fmt = ONDALIK if isinstance(deg, float) else CATI
            if "puan" in bir or "satir" in bir:
                fmt = CATI
            _sari(ws, r, 2, deg, sayi=fmt, baslik=ana, mesaj=acik)
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
        (32, "spy_doluKayit", "=HESAP!C7"),
        (33, "spy_xbarOrt", "=HESAP!C15"),
        (34, "spy_sigmaOrt", "=HESAP!C10"),
        (35, "spy_cpOrt", "=HESAP!C11"),
        (36, "spy_cpkOrt", "=HESAP!C12"),
        (37, "spy_cpkMin", "=HESAP!C13"),
        (38, "spy_aktifProses", "=HESAP!C16"),
        (39, "spy_rToplam", "=HESAP!C9"),
        (40, "spy_saglikSkor", "=HESAP!C20"),
        (41, "spy_alarmAdet", "=HESAP!C18"),
        (42, "spy_yetimAdet", "=HESAP!C19"),
        (43, "spy_gunlukXbarOrt", "=HESAP!C25"),
        (44, "spy_aylikCpkTrend", "=HESAP!C26"),
        (45, "spy_cpkSapma", "=HESAP!C27"),
        (46, "spy_tornadoZirve", "=HESAP!C39"),
        (47, "spy_senaryoKarsilastirma", "=HESAP!C35"),
        (48, "spy_hhi", "=HESAP!C40"),
        (49, "spy_kaliteSkor", "=HESAP!C41"),
        (50, "spy_senaryoMotor", "=HESAP!C44"),
        (51, "spy_tahminAralik", "=HESAP!C42"),
        (52, "spy_yuzdelikP90", "=HESAP!C43"),
        (53, "spy_sezonsalTrend",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"12_AY",TEXT(HESAP!C45,"0.000"),"Cpk",'
         'TEXT(HESAP!C46,"0"),"olcum sezonsal")'),
        (54, "spy_alarmListesi",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Alarm",HESAP!C18,"Yetim",HESAP!C19,'
         '"CpkAşım",HESAP!C22,"WE",HESAP!C23)'),
        (55, "spy_kararKapi",
         '=IF(HESAP!C7=0,"VERİ YOK",'
         'IF(OR(HESAP!C12<0,HESAP!C11<0,HESAP!C10<0),"KRİTİK",'
         'IF(HESAP!C47=1,"KRİTİK",'
         'IF(OR(HESAP!C22>0,HESAP!C12<spy_esikCpk),"DİKKAT","UYGUN"))))'),
        (56, "spy_kararGerekce",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Karar",spy_kararKapi,"| Cpk",'
         'TEXT(spy_cpkOrt,"0.000"),"| sağlık",TEXT(spy_saglikSkor,"0"),'
         '"| alarm",spy_alarmAdet)'),
        (57, "spy_sonrakiKimlik",
         '="PRS-"&TEXT(COUNTA(tblProsesler[VarlikId])+1,"00000")'),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 59, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (60, "spy_yorumGunluk",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Günlük X̄ ort",TEXT(spy_gunlukXbarOrt,"0.000"))'),
        (61, "spy_yorumAylik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Aylık Cpk proxy",TEXT(spy_aylikCpkTrend,"0.000"))'),
        (62, "spy_yorumSapma",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Cpk sapma",TEXT(spy_cpkSapma,"0.000"))'),
        (63, "spy_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(spy_tornadoZirve,"0.000"))'),
        (64, "spy_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(spy_senaryoKarsilastirma,"0.000"))'),
        (65, "spy_yorumHhi",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma HHI",TEXT(spy_hhi,"0.0%"))'),
        (66, "spy_yorumKalite",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite skoru",spy_kaliteSkor)'),
        (67, "spy_yorumSenMotor",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motoru",spy_senaryoMotor)'),
        (68, "spy_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(spy_tahminAralik,"0.000"))'),
        (69, "spy_yorumP90",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 Cpk",TEXT(spy_yuzdelikP90,"0.000"))'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 24, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. VARLIKLAR sayfasına proses kartı ekleyin (LSL/USL/hedef/altgrup n); sonraki kimliği kullanın.",
        "2. GUNLUK_GIRIS sayfasına günlük satır yazın: varlık, tarih, X̄, R, n, vardiya (≤6 alan).",
        "3. σ = R/d2; Cp = (USL−LSL)/(6σ); Cpk = min((USL−X̄)/(3σ),(X̄−LSL)/(3σ)).",
        "4. DONEM_ANALIZ 12 aylık sezonsal Cp/Cpk zincirini gösterir.",
        "5. ALARMLAR sayfasında varlık + eşik + tarih listesi; eşikler AYARLAR'dadır.",
        "6. raporTarihi AYARLAR'dadır; TODAY kullanılmaz.",
        "7. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı hücreler açıktır.",
        "8. RAPOR sayfasını PDF olarak yazdırabilirsiniz (IATF/müşteri kanıtı).",
        f"9. Sürüm {SURUM} | Kod SPY-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
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
            f"E{r}", CellIsRule(operator="lessThan", formula=["spy_esikCpk"], fill=sari))
        ws.conditional_formatting.add(
            f"F{r}", CellIsRule(operator="lessThan", formula=["spy_esikSaglik"], fill=kirmizi))

    ws = wb["GUNLUK_GIRIS"]
    ws.conditional_formatting.add(
        f"N{ILK}:N{SON}",
        CellIsRule(operator="equal", formula=['"ALARM"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"P{ILK}:P{SON}",
        CellIsRule(operator="equal", formula=['"YETIM"'], fill=sari))
    ws.conditional_formatting.add(
        f"J{ILK}:J{SON}",
        CellIsRule(operator="lessThan", formula=["spy_esikCpk"], fill=sari))
    ws.conditional_formatting.add(
        f"M{ILK}:M{SON}",
        CellIsRule(operator="lessThan", formula=["spy_esikSaglik"], fill=kirmizi))

    ws = wb["VARLIKLAR"]
    ws.conditional_formatting.add(
        f"I{ILK}:I{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"L{ILK}:L{SON}",
        CellIsRule(operator="lessThan", formula=["spy_esikSaglik"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"K{ILK}:K{SON}",
        CellIsRule(operator="lessThan", formula=["spy_esikCpk"], fill=sari))

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
            f"D{r}", CellIsRule(operator="lessThan", formula=["spy_esikCpk"], fill=sari))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "spy_d2", "spy_A2", "spy_ucSigma", "spy_altiSigma",
        "spy_esikCpk", "spy_esikCp", "spy_esikSaglik",
        "spy_aylikCarpan", "spy_senaryoTemkinli", "spy_senaryoBaz", "spy_senaryoIyimser",
        "spy_tornadoOran", "spy_yuzdelikOran", "spy_percentileOran", "spy_tahminCarpan",
        "spy_olcekHedef",
        "spy_kaliteTaban", "spy_kaliteCeza", "spy_kaliteSapma", "spy_dosyaSurumu",
        "spy_firmaUnvan", "spy_alarmEsikMetin",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "spy_doluKayit": 32, "spy_xbarOrt": 33, "spy_sigmaOrt": 34,
        "spy_cpOrt": 35, "spy_cpkOrt": 36, "spy_cpkMin": 37,
        "spy_aktifProses": 38, "spy_rToplam": 39, "spy_saglikSkor": 40,
        "spy_alarmAdet": 41, "spy_yetimAdet": 42, "spy_gunlukXbarOrt": 43,
        "spy_aylikCpkTrend": 44, "spy_cpkSapma": 45, "spy_tornadoZirve": 46,
        "spy_senaryoKarsilastirma": 47, "spy_hhi": 48, "spy_kaliteSkor": 49,
        "spy_senaryoMotor": 50, "spy_tahminAralik": 51, "spy_yuzdelikP90": 52,
        "spy_sezonsalTrend": 53, "spy_alarmListesi": 54, "spy_kararKapi": 55,
        "spy_kararGerekce": 56, "spy_sonrakiKimlik": 57,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "spy_yorumGunluk": 60, "spy_yorumAylik": 61, "spy_yorumSapma": 62,
        "spy_yorumTornado": 63, "spy_yorumSenaryo": 64, "spy_yorumHhi": 65,
        "spy_yorumKalite": 66, "spy_yorumSenMotor": 67, "spy_yorumTahmin": 68,
        "spy_yorumP90": 69,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$A$7:$A$8")
    ad_ekle(wb, "ListeVardiya", "LISTELER!$C$7:$C$12")
    ad_ekle(wb, "ListeProsesTipi", "LISTELER!$E$7:$E$13")


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
    dosya_ad = "SpcProsesYetenekAnalizi.xlsx"
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
        f.write("SatirId,VarlikId,Tarih,Xbar_birim,R_birim,AltgrupN_adet,Vardiya\n")
        for s in ORNEK:
            f.write(f"{s['id']},{s['varlik']},{s['tarih']:%d.%m.%Y},"
                    f"{s['xbar']},{s['r']},{s['n']},{s['vardiya']}\n")

    print(f"Dosya oluşturuldu: {hedef}")
    print(f"Kopya: {cikti}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
