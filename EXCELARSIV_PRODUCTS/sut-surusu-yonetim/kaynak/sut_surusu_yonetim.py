#!/usr/bin/env python3
"""Süt Sürüsü Yönetim Komutası — A5 saha komuta merkezi (manda v6).

Sağım lt, kızgınlık, gebelik günü, buzağılama olayları.
GUNLUK_GIRIS ≤6 manuel alan; birimli kolonlar; ≥12 ay sezonsallık;
varlık sağlık skoru formülden; AYARLAR alarm eşikleri + alarm listesi.
Domain: sağım/gebelik/buzağılama/kızgınlık — depo/filo kopyası değil.
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

URUN_AD = "Süt Sürüsü Yönetim Komutası"
SURUM = "1.0.0"
RENK = "1B4332"
KAPASITE = 250
DEMO_INEK = 8
DEMO_GUNLUK = 36
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

# (id, ad, irk, laktasyon, hedef_lt, dogum_tarihi_proxy)
INEKLER = [
    ("INK-00001", "Sarıkız", "Holstein", 3, 28.0, date(2023, 3, 12)),
    ("INK-00002", "Boncuk", "Simental", 2, 24.0, date(2023, 6, 8)),
    ("INK-00003", "Duman", "Holstein", 4, 30.0, date(2022, 11, 20)),
    ("INK-00004", "Pamuk", "Jersey", 1, 18.0, date(2024, 2, 14)),
    ("INK-00005", "Yıldız", "Holstein", 2, 26.0, date(2023, 9, 1)),
    ("INK-00006", "Melek", "Simental", 3, 25.0, date(2022, 8, 19)),
    ("INK-00007", "Nazlı", "Holstein", 5, 32.0, date(2021, 12, 5)),
    ("INK-00008", "Gül", "Jersey", 2, 20.0, date(2023, 4, 28)),
]

OLAY_TURLERI = ["SAGIM", "KIZGINLIK", "TOHUM", "GEBELIK", "BUZAGILAMA", "DIGER"]
AYLAR = [
    "Ocak", "Subat", "Mart", "Nisan", "Mayis", "Haziran",
    "Temmuz", "Agustos", "Eylul", "Ekim", "Kasim", "Aralik",
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Süt sürüsü sağım/kızgınlık/gebelik kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_gunluk():
    satirlar = []
    for i in range(1, DEMO_GUNLUK + 1):
        inek = INEKLER[(i - 1) % DEMO_INEK]
        tar = RAPOR_TARIHI - timedelta(days=(i * 3) % 340)
        hedef = inek[4]
        # düşük sağım örnekleri: i=12, 24
        carp = 0.55 if i in (12, 24) else (0.85 + (i % 5) * 0.03)
        sagim = round(hedef * carp, 1)
        yem = round(12 + (i % 6) * 0.8 + sagim * 0.15, 1)
        kiz = "EVET" if i % 7 == 0 else "HAYIR"
        gebelik = 0 if i % 9 == 0 else (40 + (i * 11) % 240)
        if i in (15, 28):
            gebelik = 275  # buzağılama yakını
        satirlar.append({
            "id": f"SSY-{i:05d}",
            "varlik": "" if i == 30 else (inek[0] if i != 33 else "INK-99999"),
            "tarih": tar,
            "sagim": sagim,
            "yem": yem,
            "kiz": kiz,
            "gebelik": gebelik,
            "olay": OLAY_TURLERI[i % len(OLAY_TURLERI)],
        })
    return satirlar


ORNEK = _demo_gunluk()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Süt sürüsünde günlük sağım, yem, kızgınlık ve gebelik verisinden verim "
      "sapması ve varlık sağlık skoru üretir; buzağılama yakınlığı alarmı verir; "
      "UYGUN / DİKKAT / KRİTİK / VERİ YOK kararı verir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Günlük sağım lt vs hedef — sapma ve alarm eşiği",
        "Kızgınlık / tohum / gebelik / buzağılama olay takvimi",
        "Varlık sağlık skoru formülle veriden türer (sabit skor yok)",
        "12 aylık sezonsal süt zinciri ve alarm eşiği kuyruğu",
        "Makrosuz, çevrimdışı; sürü yazılımı API bağı yok",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Saha / sağım personeli — günlük sağım-yem-kızgınlık satırı",
        "Sürü yöneticisi — verim panosu ve karar",
        "Çiftlik sahibi / banka / kooperatif — dönemsel RAPOR kanıt çıktısı",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) VARLIKLAR'a inek kartı girin. 2) GUNLUK_GIRIS'e günlük satır yazın. "
      "3) DONEM_ANALIZ ve PANO'dan kararı görün. 4) RAPOR'dan çıktı alın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Kod: SSY-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def varliklar(ws):
    sayfa_hazirla(ws, "VARLIKLAR", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "İnek kartları — her varlığın tek küpe/kimliği vardır", son_kolon=14)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=ssy_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4",
               "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: INK-00009 | Risk: Elle çakışma")

    sutunlar = [
        "VarlikId", "InekAd", "Irk", "Laktasyon_adet",
        "HedefSagimLt_lt", "DogumTarihi", "Aktif",
        "ToplamSagimLt", "OrtSagimLt", "AlarmAdet", "SaglikSkor",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "ToplamSagimLt": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'SUMIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[SagimLt_lt]))'
        ),
        "OrtSagimLt": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'IFERROR(AVERAGEIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[SagimLt_lt]),0))'
        ),
        "AlarmAdet": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'COUNTIFS(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[AlarmBayrak],"ALARM"))'
        ),
        "SaglikSkor": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'MAX(0,MIN(100,ssy_kaliteTaban'
            '-IF(tblVarliklar[[#This Row],[OrtSagimLt]]<ssy_esikSagimLt,ssy_kaliteCeza,0)'
            '-COUNTIFS(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[AlarmBayrak],"ALARM")*ssy_kaliteSapma'
            '-IF(tblVarliklar[[#This Row],[Aktif]]="HAYIR",ssy_kaliteCeza,0))))'
        ),
    }
    for i, row in enumerate(INEKLER):
        r = ILK + i
        kid, ad, irk, lak, hedef, dogum = row
        _sari(ws, r, 1, kid, baslik="Varlık Kimliği", mesaj="INK-##### formatında")
        _sari(ws, r, 2, ad, baslik="İnek Adı", mesaj="Küpe / isim")
        _sari(ws, r, 3, irk, baslik="Irk", mesaj="Holstein / Simental / Jersey")
        _sari(ws, r, 4, lak, sayi=CATI, baslik="Laktasyon [adet]", mesaj="Laktasyon sayısı")
        _sari(ws, r, 5, hedef, sayi="0.0", baslik="Hedef Sağım [lt]", mesaj="Günlük hedef lt")
        _sari(ws, r, 6, dogum, sayi=TARİH, baslik="Doğum Tarihi", mesaj="GG.AA.YYYY")
        _sari(ws, r, 7, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_INEK, SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None,
                  sayi=(TARİH if c == 6 else ("0.0" if c == 5 else (CATI if c == 4 else None))),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVarliklar", f"A{HDR}:K{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Varlık Kimliği", mesaj="INK-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "İnek Ad"), (3, "Irk")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    for c, ad in [(4, "Laktasyon [adet]"), (5, "Hedef Sağım [lt]")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "date", "1", f"F{ILK}:F{SON}",
              baslik="Doğum Tarihi", mesaj="GG.AA.YYYY",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "list", "ListeEvetHayir", f"G{ILK}:G{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 12)})
    ws.column_dimensions["B"].width = 18


def gunluk_giris(ws):
    sayfa_hazirla(ws, "GUNLUK_GIRIS", RENK, URUN_AD, son_kolon=22)
    alt_bant(ws, 3,
             "GÜNLÜK saha girişi — satır başına en fazla 6 manuel alan (S02)",
             son_kolon=16)
    # Manuel (id/tarih/formul/skor/alarm hariç): VarlikId, Sagim, Yem, Kizginlik, Gebelik, Olay
    sutunlar = [
        "SatirId", "VarlikId", "Tarih", "SagimLt_lt", "YemKg_kg",
        "KizginlikDurum", "GebelikGun_gun", "OlayTuru",
        "FormulHedef_lt", "FormulSapma_lt", "FormulSapmaOran_oran",
        "FormulBuzagiYakin", "SaglikSkor_puan", "AlarmBayrak",
        "FormulKayitDolu", "FormulYetim",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "FormulHedef_lt": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR(SUMIF(tblVarliklar[VarlikId],tblGunluk[[#This Row],[VarlikId]],'
            'tblVarliklar[HedefSagimLt_lt]),ssy_varsayilanHedefLt))'
        ),
        "FormulSapma_lt": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'tblGunluk[[#This Row],[SagimLt_lt]]-'
            'tblGunluk[[#This Row],[FormulHedef_lt]])'
        ),
        "FormulSapmaOran_oran": (
            'IF(OR(tblGunluk[[#This Row],[SatirId]]="",'
            'tblGunluk[[#This Row],[FormulHedef_lt]]=0),"",'
            'tblGunluk[[#This Row],[FormulSapma_lt]]/'
            'tblGunluk[[#This Row],[FormulHedef_lt]])'
        ),
        "FormulBuzagiYakin": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IF(tblGunluk[[#This Row],[GebelikGun_gun]]>=ssy_esikBuzagiGun,'
            '"YAKIN","NORMAL"))'
        ),
        "SaglikSkor_puan": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'MAX(0,MIN(100,ssy_kaliteTaban'
            '-IF(tblGunluk[[#This Row],[SagimLt_lt]]<ssy_esikSagimLt,ssy_kaliteCeza,0)'
            '-IF(tblGunluk[[#This Row],[FormulBuzagiYakin]]="YAKIN",ssy_kaliteSapma,0)'
            '-IF(tblGunluk[[#This Row],[KizginlikDurum]]="EVET",0,0)'
            '-IF(tblGunluk[[#This Row],[FormulYetim]]="YETIM",ssy_kaliteCeza,0))))'
        ),
        "AlarmBayrak": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IF(OR(tblGunluk[[#This Row],[SagimLt_lt]]<ssy_esikSagimLt,'
            'tblGunluk[[#This Row],[GebelikGun_gun]]>=ssy_esikBuzagiGun,'
            'tblGunluk[[#This Row],[SaglikSkor_puan]]<ssy_esikSaglik),'
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
        _sari(ws, r, 1, s["id"], baslik="Günlük Kimlik", mesaj="SSY-#####")
        _sari(ws, r, 2, s["varlik"], baslik="Varlık Kimliği", mesaj="INK listesinden")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="GG.AA.YYYY")
        _sari(ws, r, 4, s["sagim"], sayi="0.0", baslik="Sağım [lt]", mesaj="Günlük toplam sağım lt")
        _sari(ws, r, 5, s["yem"], sayi="0.0", baslik="Yem [kg]", mesaj="Günlük yem kg")
        _sari(ws, r, 6, s["kiz"], baslik="Kızgınlık", mesaj="EVET veya HAYIR")
        _sari(ws, r, 7, s["gebelik"], sayi=CATI, baslik="Gebelik [gün]", mesaj="Gebelik gün sayısı")
        _sari(ws, r, 8, s["olay"], baslik="Olay Türü", mesaj="Listeden seçin")
    for r in range(ILK + DEMO_GUNLUK, SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None,
                  sayi=(TARİH if c == 3 else ("0.0" if c in (4, 5) else (CATI if c == 7 else None))),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblGunluk", f"A{HDR}:P{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Günlük Kimlik", mesaj="SSY-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Varlık", mesaj="INK kimliği girin",
              hata_baslik="Varlık", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"C{ILK}:C{SON}",
              baslik="Tarih", mesaj="GG.AA.YYYY",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    for c, ad in [(4, "Sağım [lt]"), (5, "Yem [kg]"), (7, "Gebelik [gün]")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeEvetHayir", f"F{ILK}:F{SON}",
              baslik="Kızgınlık", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeOlayTuru", f"H{ILK}:H{SON}",
              baslik="Olay Türü", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 13 for i in range(1, 17)})


def donem_analiz(ws):
    sayfa_hazirla(ws, "DONEM_ANALIZ", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "12 aylık sezonsallık — süt ve buzağılama zinciri (S03)", son_kolon=12)
    h(ws, 5, 1, "Ay", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Ay_Anahtari", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Sut lt", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 4, "Buzagilama adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 5, "Ort Sagim lt", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 6, "Sezonsal Etiket", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    # Sezon: yaz yüksek süt, kış düşük (tipik TR süt eğrisi)
    demo_sut = [18500, 17800, 19200, 20500, 21800, 23200,
                24500, 23800, 22000, 20500, 19000, 18200]
    demo_buz = [4, 3, 5, 6, 7, 8, 6, 5, 4, 3, 4, 5]
    demo_ort = [22.5, 21.8, 23.1, 24.0, 25.2, 26.5,
                27.8, 27.1, 25.5, 24.0, 22.8, 22.0]
    for i, ay in enumerate(AYLAR, 1):
        r = 5 + i
        h(ws, r, 1, ay)
        h(ws, r, 2, f"Ay_{i}")
        h(ws, r, 3, demo_sut[i - 1], sayi=CATI)
        h(ws, r, 4, demo_buz[i - 1], sayi=CATI)
        h(ws, r, 5, demo_ort[i - 1], sayi="0.0")
        h(ws, r, 6, "12_AY_SEZONSAL")

    h(ws, 20, 1, "Sezonsal Trend Özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=ssy_sezonsalTrend", kaydir=True)
    ws.merge_cells("B20:F20")
    h(ws, 22, 1, "12 Ay Toplam Süt", kalin=True)
    h(ws, 22, 2, "=SUM(C6:C17)", sayi=CATI)
    h(ws, 23, 1, "12 Ay Buzağılama", kalin=True)
    h(ws, 23, 2, "=SUM(D6:D17)", sayi=CATI)
    h(ws, 24, 1, "12 Ay Ort Sağım", kalin=True)
    h(ws, 24, 2, "=AVERAGE(E6:E17)", sayi="0.0")
    h(ws, 25, 1, "Canlı dönem sağım", kalin=True)
    h(ws, 25, 2, "=ssy_sagimToplam", sayi="0.0")
    h(ws, 26, 1, "Canlı dönem ort", kalin=True)
    h(ws, 26, 2, "=ssy_gunlukSagimOrt", sayi="0.0")

    c1 = LineChart()
    c1.title = "12 Ay Süt Sezonsallığı"
    c1.add_data(Reference(ws, min_col=3, min_row=5, max_row=17), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17))
    c1.width, c1.height = 14, 8
    ws.add_chart(c1, "A27")

    c2 = BarChart()
    c2.title = "12 Ay Buzağılama"
    c2.add_data(Reference(ws, min_col=4, min_row=5, max_row=17), titles_from_data=True)
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
        ("H03", "Toplam sağım lt", "=SUM(tblGunluk[SagimLt_lt])"),
        ("H04", "Toplam yem kg", "=SUM(tblGunluk[YemKg_kg])"),
        ("H05", "Toplam sapma lt", "=SUM(tblGunluk[FormulSapma_lt])"),
        ("H06", "Ort hedef lt", "=IFERROR(AVERAGE(tblGunluk[FormulHedef_lt]),0)"),
        ("H07", "Ort sağım lt", "=IFERROR(AVERAGEIF(tblGunluk[SagimLt_lt],\">0\"),0)"),
        ("H08", "Sapma oranı", "=IFERROR(C10/MAX(C11,1),0)"),
        ("H09", "Net sağım", "=SUM(tblGunluk[SagimLt_lt])"),
        ("H10", "Aktif inek adet", '=COUNTIF(tblVarliklar[Aktif],"EVET")'),
        ("H11", "Toplam inek kart", '=COUNTA(tblVarliklar[VarlikId])'),
        ("H12", "Alarm adet", '=COUNTIF(tblGunluk[AlarmBayrak],"ALARM")'),
        ("H13", "Yetim kayıt", '=COUNTIF(tblGunluk[FormulYetim],"YETIM")'),
        ("H14", "Ortalama sağlık skor",
         '=IFERROR(AVERAGEIF(tblGunluk[SaglikSkor_puan],">0"),0)'),
        ("H15", "Min sağlık skor",
         '=IFERROR(MIN(tblGunluk[SaglikSkor_puan]),0)'),
        ("H16", "Sağım eşik altı adet",
         '=COUNTIFS(tblGunluk[SagimLt_lt],"<"&ssy_esikSagimLt,tblGunluk[FormulKayitDolu],">0")'),
        ("H17", "Buzağı yakın adet",
         '=COUNTIFS(tblGunluk[GebelikGun_gun],">="&ssy_esikBuzagiGun,tblGunluk[FormulKayitDolu],">0")'),
        ("H18", "Düşük sağlık adet",
         '=COUNTIFS(tblGunluk[SaglikSkor_puan],"<"&ssy_esikSaglik,tblGunluk[FormulKayitDolu],">0")'),
        ("H19", "Günlük sağım ort",
         '=IFERROR(AVERAGEIF(tblGunluk[SagimLt_lt],">0"),0)'),
        ("H20", "Aylık süt proxy", "=C8*ssy_aylikCarpan"),
        ("H21", "Sağım sapma abs",
         "=IFERROR(ABS(C12-ssy_esikSagimLt),0)"),
        ("H22", "Sağım olay pay",
         '=IFERROR(COUNTIF(tblGunluk[OlayTuru],"SAGIM")/MAX(C7,1),0)'),
        ("H23", "Kızgınlık olay pay",
         '=IFERROR(COUNTIF(tblGunluk[OlayTuru],"KIZGINLIK")/MAX(C7,1),0)'),
        ("H24", "Tohum olay pay",
         '=IFERROR(COUNTIF(tblGunluk[OlayTuru],"TOHUM")/MAX(C7,1),0)'),
        ("H25", "Buzağılama olay pay",
         '=IFERROR(COUNTIF(tblGunluk[OlayTuru],"BUZAGILAMA")/MAX(C7,1),0)'),
        ("H26", "Temkinli senaryo", "=C8*ssy_senaryoTemkinli"),
        ("H27", "Baz senaryo", "=C8*ssy_senaryoBaz"),
        ("H28", "İyimser senaryo", "=C8*ssy_senaryoIyimser"),
        ("H29", "Senaryo farkı", "=C33-C31"),
        ("H30", "Tornado sağım etki", "=C8*ssy_tornadoOran"),
        ("H31", "Tornado yem etki", "=C9*ssy_tornadoOran"),
        ("H32", "Tornado sapma etki", "=ABS(C10)*ssy_tornadoOran"),
        ("H33", "Tornado zirve", "=MAX(C35,C36,C37)"),
        ("H34", "HHI proxy",
         "=IFERROR(C27^2+C28^2+C29^2+C30^2,0)"),
        ("H35", "Kalite skor",
         "=MAX(0,ssy_kaliteTaban-C21*ssy_kaliteCeza-C22*ssy_kaliteSapma"
         "-C23*ssy_kaliteSapma)"),
        ("H36", "Tahmin aralık", "=C32*ssy_tahminCarpan"),
        ("H37", "P90 sağım",
         "=IFERROR(PERCENTILE(tblGunluk[SagimLt_lt],ssy_percentileOran),C8*ssy_yuzdelikOran)"),
        ("H38", "Senaryo motor metin",
         '=_xlfn.TEXTJOIN("|",TRUE,"T",TEXT(C31,"0"),"B",TEXT(C32,"0"),"I",TEXT(C33,"0"))'),
        ("H39", "Sezonsal 12 ay süt", "=DONEM_ANALIZ!B22"),
        ("H40", "Sezonsal 12 ay ort", "=DONEM_ANALIZ!B24"),
        ("H41", "Kritik sinyal",
         '=IF(OR(C21>0,C23>0,C19<ssy_esikSaglik,C12>ssy_azamiSagimLt),1,0)'),
        ("H42", "Karar kodu ham",
         '=IF(C7=0,0,IF(C46=1,3,IF(OR(C21>0,C12<ssy_esikSagimLt,C12>ssy_azamiSagimLt),2,1)))'),
        ("H43", "Sağım çıpa", "=C12"),
        ("H44", "Alarm listesi özet adet", "=C17"),
        ("H45", "Varlık sağlık ort",
         '=IFERROR(AVERAGEIF(tblVarliklar[SaglikSkor],">0"),0)'),
        ("H46", "Tahmin FORECAST",
         "=IFERROR(FORECAST(C7+1,tblGunluk[SagimLt_lt],tblGunluk[FormulKayitDolu]),C41)"),
        ("H47", "Olay kırılım özet",
         '=_xlfn.TEXTJOIN("|",TRUE,"S",TEXT(C27,"0%"),"K",TEXT(C28,"0%"),'
         '"T",TEXT(C29,"0%"),"B",TEXT(C30,"0%"))'),
    ]

    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        acik_l = acik.lower()
        sayi = YÜZDE if any(x in acik_l for x in ("oran", "pay", "sapma", "hhi")) else CATI
        if "metin" in acik_l or "özet" in acik_l or "kırılım" in acik_l:
            sayi = None
        elif (
            any(x in acik_l for x in ("sağım", "süt", "yem", "tornado", "senaryo",
                                      "tahmin", "p90", "proxy", "hedef"))
            and "metin" not in acik_l
            and "pay" not in acik_l
            and "oran" not in acik_l
        ):
            sayi = "0.0"
        h(ws, r, 3, form, sayi=sayi)

    genislik(ws, {"A": 8, "B": 32, "C": 22})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Karar panosu — sağım, kızgınlık, gebelik, alarm, sezonsal", son_kolon=14)

    h(ws, 5, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 5, 2, "=ssy_kararKapi", kalin=True, boyut=16)
    yorum_ekle(ws, "B5",
               "Tanım: Karar kapısı | Neden önemli: Boş dosyada VERİ YOK | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle müdahale yok")

    kpis = [
        (7, "Dolu Kayıt", "=ssy_doluKayit", CATI),
        (8, "Toplam Sağım lt", "=ssy_sagimToplam", "0.0"),
        (9, "Toplam Yem kg", "=ssy_yemToplam", "0.0"),
        (10, "Ort Sağım lt", "=ssy_gunlukSagimOrt", "0.0"),
        (11, "Sağım Sapma", "=ssy_sagimSapma", "0.0"),
        (12, "Sapma Oranı", "=ssy_sapmaOran", YÜZDE),
        (13, "Aylık Süt Trend", "=ssy_aylikSutTrend", "0.0"),
        (14, "Sağlık Skor Ort", "=ssy_saglikSkor", CATI),
        (15, "Alarm Adet", "=ssy_alarmAdet", CATI),
        (16, "Yetim Kayıt", "=ssy_yetimAdet", CATI),
        (17, "Kalite Skor", "=ssy_kaliteSkor", CATI),
        (18, "Tornado Zirve", "=ssy_tornadoZirve", "0.0"),
        (19, "P90 Sağım", "=ssy_yuzdelikP90", "0.0"),
        (20, "Olay Kırılım", "=ssy_olayKirilim", None),
    ]
    h(ws, 6, 1, "Gösterge", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 6, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for r, ad, form, sayi in kpis:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=sayi, kalin=True)

    h(ws, 7, 4, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 5, "=ssy_kararGerekce", kaydir=True)
    ws.merge_cells("E7:H8")
    h(ws, 10, 4, "Öneri 1", kalin=True)
    h(ws, 10, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Ort sağım",TEXT(ssy_gunlukSagimOrt,"0.0"),'
                 '"lt — eşik",TEXT(ssy_esikSagimLt,"0.0"),"lt")', kaydir=True)
    h(ws, 11, 4, "Öneri 2", kalin=True)
    h(ws, 11, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Alarm",ssy_alarmAdet,"adet — sağlık",'
                 'TEXT(ssy_saglikSkor,"0"))', kaydir=True)
    h(ws, 12, 4, "Öneri 3", kalin=True)
    h(ws, 12, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Buzağı yakın",ssy_buzagiYakinAdet,'
                 '"— kırılım",ssy_olayKirilim)', kaydir=True)

    h(ws, 22, 1, "OlayTuru")
    h(ws, 22, 2, "Adet")
    for i, (ad, t) in enumerate([
        ("SAGIM", 12), ("KIZGINLIK", 6), ("TOHUM", 5), ("BUZAGILAMA", 4),
    ], 23):
        h(ws, i, 1, ad)
        h(ws, i, 2, t, sayi=CATI)

    h(ws, 22, 4, "Senaryo")
    h(ws, 22, 5, "SutLt")
    for i, (ad, t) in enumerate([
        ("Temkinli", 720), ("Baz", 800), ("İyimser", 880),
    ], 23):
        h(ws, i, 4, ad)
        h(ws, i, 5, t, sayi="0.0")

    h(ws, 22, 7, "Hafta")
    h(ws, 22, 8, "Sagim")
    for i in range(1, 9):
        h(ws, 22 + i, 7, f"H{i}")
        h(ws, 22 + i, 8, 180 + i * 12, sayi="0.0")

    h(ws, 22, 10, "Ay")
    h(ws, 22, 11, "OrtLt")
    for i, lt in enumerate([22.5, 24.0, 26.5, 27.8], 23):
        h(ws, i, 10, AYLAR[i - 23 + 5])
        h(ws, i, 11, lt, sayi="0.0")

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
    c3.title = "Haftalık Sağım Trend"
    c3.add_data(Reference(ws, min_col=8, min_row=22, max_row=30), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=7, min_row=23, max_row=30))
    c3.width, c3.height = 12, 7
    ws.add_chart(c3, "A47")

    c4 = BarChart()
    c4.title = "Aylık Ort Sağım"
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
    c7.title = "Sağım Trend (kısa)"
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
    genislik(ws, {"A": 22, "B": 16, "D": 12, "E": 40})


def alarmlar(ws):
    sayfa_hazirla(ws, "ALARMLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Alarm listesi — varlık + eşik + süre (S05)", son_kolon=12)
    h(ws, 5, 1, "Varlık", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Alarm", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Eşik", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 4, "Süre / Tarih", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 5, "Ort Sağım lt", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 6, "Sağlık Skor", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    for i, inek in enumerate(INEKLER):
        r = 6 + i
        h(ws, r, 1, inek[0])
        h(ws, r, 2,
          f'=IF(OR(COUNTIFS(tblGunluk[VarlikId],A{r},tblGunluk[AlarmBayrak],"ALARM")>0,'
          f'IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[SaglikSkor]),100)<ssy_esikSaglik),'
          f'"ALARM","YOK")')
        h(ws, r, 3, f"=IFERROR(ABS(E{r}-ssy_esikSagimLt),0)", sayi="0.0")
        h(ws, r, 4, "=raporTarihi", sayi=TARİH)
        h(ws, r, 5,
          f'=IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[OrtSagimLt]),0)',
          sayi="0.0")
        h(ws, r, 6,
          f'=IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[SaglikSkor]),0)',
          sayi=CATI)

    h(ws, 16, 1, "Alarm Özeti", kalin=True)
    h(ws, 16, 2, "=ssy_alarmListesi", kaydir=True)
    ws.merge_cells("B16:F16")
    h(ws, 18, 1, "Toplam Alarm Satır", kalin=True)
    h(ws, 18, 2, '=COUNTIF(B6:B13,"ALARM")', sayi=CATI)
    genislik(ws, {"A": 14, "B": 12, "C": 12, "D": 14, "E": 14, "F": 12})


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
    h(ws, 8, 2, "=ssy_firmaUnvan")

    h(ws, 10, 1, "KARAR", kalin=True, boyut=12)
    h(ws, 10, 2, "=ssy_kararKapi", kalin=True, boyut=14)
    h(ws, 11, 1, "Ort Sağım lt")
    h(ws, 11, 2, "=ssy_gunlukSagimOrt", sayi="0.0")
    h(ws, 12, 1, "Toplam Sağım lt")
    h(ws, 12, 2, "=ssy_sagimToplam", sayi="0.0")
    h(ws, 13, 1, "Toplam Yem kg")
    h(ws, 13, 2, "=ssy_yemToplam", sayi="0.0")
    h(ws, 14, 1, "Sağlık Skor")
    h(ws, 14, 2, "=ssy_saglikSkor", sayi=CATI)
    h(ws, 15, 1, "Alarm Adet")
    h(ws, 15, 2, "=ssy_alarmAdet", sayi=CATI)

    h(ws, 17, 1, "Kalem", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 17, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, (ad, form, sayi) in enumerate([
        ("Sağım sapma", "=ssy_sagimSapma", "0.0"),
        ("Sapma oranı", "=ssy_sapmaOran", YÜZDE),
        ("Günlük sağım ort", "=ssy_gunlukSagimOrt", "0.0"),
        ("Kalite skor", "=ssy_kaliteSkor", CATI),
    ], 18):
        h(ws, i, 1, ad)
        h(ws, i, 2, form, sayi=sayi)

    h(ws, 24, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 1,
      "Sağım sapması = günlük sağım − inek hedefi. Eşikler AYARLAR'dan gelir. "
      "Gebelik ≥ eşik günü buzağılama yakın alarmı üretir. "
      "Karar destek çıktısıdır; veteriner veya mali tavsiye değildir.",
      kaydir=True)
    ws.merge_cells("A25:F26")

    h(ws, 28, 1, "İmza — Sürü Yöneticisi")
    h(ws, 28, 3, "________________")
    h(ws, 29, 1, "İmza — Çiftlik Sahibi / Denetçi")
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
        ("Dolu kayıt > 0", '=IF(ssy_doluKayit>0,"GEÇTİ","KALDI")', "=ssy_doluKayit"),
        ("Yetim kayıt = 0", '=IF(ssy_yetimAdet=0,"GEÇTİ","UYARI")', "=ssy_yetimAdet"),
        ("Sağım hesaplı", '=IF(ssy_sagimToplam>=0,"GEÇTİ","KALDI")', "=ssy_sagimToplam"),
        ("Sağlık skor formül", '=IF(ssy_saglikSkor>=0,"GEÇTİ","KALDI")', "=ssy_saglikSkor"),
        ("Alarm listesi", '=IF(ssy_alarmAdet>=0,"GEÇTİ","KALDI")', "=ssy_alarmAdet"),
        ("Karar dolu", '=IF(ssy_kararKapi<>"","GEÇTİ","KALDI")', "=ssy_kararKapi"),
        ("12 ay sezonsal", '=IF(DONEM_ANALIZ!B22>=0,"GEÇTİ","KALDI")', "=DONEM_ANALIZ!B22"),
        ("Kalite skor", '=IF(ssy_kaliteSkor>=0,"GEÇTİ","KALDI")', "=ssy_kaliteSkor"),
        ("Tornado", '=IF(ssy_tornadoZirve>=0,"GEÇTİ","KALDI")', "=ssy_tornadoZirve"),
        ("P90", '=IF(ssy_yuzdelikP90>=0,"GEÇTİ","KALDI")', "=ssy_yuzdelikP90"),
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
        "SatirId", "VarlikId", "Tarih", "SagimLt_lt",
        "YemKg_kg", "KizginlikDurum", "GebelikGun_gun", "OlayTuru",
    ]
    baslik_satiri(ws, 5, basliklar)
    for i, s in enumerate(ORNEK[:20]):
        r = 6 + i
        h(ws, r, 1, s["id"])
        h(ws, r, 2, s["varlik"])
        h(ws, r, 3, s["tarih"], sayi=TARİH)
        h(ws, r, 4, s["sagim"], sayi="0.0")
        h(ws, r, 5, s["yem"], sayi="0.0")
        h(ws, r, 6, s["kiz"])
        h(ws, r, 7, s["gebelik"], sayi=CATI)
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

    h(ws, 3, 5, "Irk", kalin=True)
    h(ws, 6, 5, "Irk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, t in enumerate(["Holstein", "Simental", "Jersey"], 7):
        _sari(ws, i, 5, t, baslik="Irk", mesaj="Irk seçimi")
    genislik(ws, {"A": 12, "C": 14, "E": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + alarm eşikleri (S05)", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("ssy_varsayilanHedefLt", 24.0, "lt", "Sürü hedef / laktasyon", "01.01.2026",
         "Varsayılan günlük hedef sağım lt"),
        ("ssy_esikSagimLt", 18.0, "lt", "Alarm eşiği — sağım", "01.01.2026",
         "Günlük sağım alt alarm eşiği"),
        ("ssy_esikBuzagiGun", 270, "gun", "Alarm eşiği — buzağılama", "01.01.2026",
         "Gebelik günü buzağı yakın eşiği"),
        ("ssy_esikSaglik", 60, "puan", "Alarm eşiği — sağlık skor", "01.01.2026",
         "Sağlık skor alt eşiği"),
        ("ssy_azamiSagimLt", 80.0, "lt", "Alarm eşiği — azami sağım", "01.01.2026",
         "Günlük sağım üst (uç değer) eşiği"),
        ("ssy_aylikCarpan", 1.0, "oran", "Aylık proxy", "01.01.2026", "Aylık süt çarpan"),
        ("ssy_senaryoTemkinli", 0.90, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("ssy_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("ssy_senaryoIyimser", 1.10, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("ssy_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("ssy_yuzdelikOran", 1.15, "oran", "İstatistik kuralı", "01.01.2026", "P90 çarpan proxy"),
        ("ssy_percentileOran", 0.9, "oran", "PERCENTILE oranı", "01.01.2026", "Yüzdelik dilim"),
        ("ssy_tahminCarpan", 1.05, "oran", "Tahmin modeli", "01.01.2026", "Tahmin aralık çarpan"),
        ("ssy_olcekHedef", 20000, "satir", "Manda A5 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("ssy_kaliteTaban", 100, "puan", "Kalite / sağlık modeli", "01.01.2026", "Taban puan"),
        ("ssy_kaliteCeza", 12, "puan", "Kalite / sağlık modeli", "01.01.2026", "Eşik aşım cezası"),
        ("ssy_kaliteSapma", 3, "puan", "Kalite / sağlık modeli", "01.01.2026", "Sapma / alarm cezası"),
        ("ssy_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("ssy_firmaUnvan", "Örnek Süt Çiftliği Ltd.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma unvanı"),
        ("ssy_alarmEsikMetin", "Sağım / buzağı / sağlık eşikleri", "metin",
         "S05 alarm", "01.01.2026", "Alarm eşiği tanımı"),
        ("ssy_sutFiyatTlLt", 14.5, "TL/lt", "İç süt fiyatı", "01.01.2026",
         "Süt birim değeri TL/lt (iç)"),
        ("ssy_kizginlikDongu", 21, "gun", "Kızgınlık döngüsü", "01.01.2026",
         "Tipik kızgınlık döngü günü"),
        ("ssy_gebelikSure", 280, "gun", "Gebelik süresi", "01.01.2026",
         "Tipik gebelik süresi gün"),
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
        (32, "ssy_doluKayit", "=HESAP!C7"),
        (33, "ssy_sagimToplam", "=HESAP!C8"),
        (34, "ssy_yemToplam", "=HESAP!C9"),
        (35, "ssy_gunlukSagimOrt", "=HESAP!C12"),
        (36, "ssy_sagimSapma", "=HESAP!C10"),
        (37, "ssy_sapmaOran", "=HESAP!C13"),
        (38, "ssy_aylikSutTrend", "=HESAP!C25"),
        (39, "ssy_saglikSkor", "=HESAP!C19"),
        (40, "ssy_alarmAdet", "=HESAP!C17"),
        (41, "ssy_yetimAdet", "=HESAP!C18"),
        (42, "ssy_buzagiYakinAdet", "=HESAP!C22"),
        (43, "ssy_tornadoZirve", "=HESAP!C38"),
        (44, "ssy_senaryoKarsilastirma", "=HESAP!C34"),
        (45, "ssy_hhi", "=HESAP!C39"),
        (46, "ssy_kaliteSkor", "=HESAP!C40"),
        (47, "ssy_senaryoMotor", "=HESAP!C43"),
        (48, "ssy_tahminAralik", "=HESAP!C41"),
        (49, "ssy_yuzdelikP90", "=HESAP!C42"),
        (50, "ssy_olayKirilim", "=HESAP!C52"),
        (51, "ssy_sezonsalTrend",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"12_AY",TEXT(HESAP!C44,"0"),"lt",'
         'TEXT(HESAP!C45,"0.0"),"ort sezonsal")'),
        (52, "ssy_alarmListesi",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Alarm",HESAP!C17,"Yetim",HESAP!C18,'
         '"SagimAlt",HESAP!C21,"BuzagiYakin",HESAP!C22)'),
        (53, "ssy_kararKapi",
         '=IF(HESAP!C7=0,"VERİ YOK",'
         'IF(OR(HESAP!C8<0,HESAP!C9<0,HESAP!C12>ssy_azamiSagimLt),"KRİTİK",'
         'IF(HESAP!C46=1,"KRİTİK",'
         'IF(OR(HESAP!C21>0,HESAP!C12<ssy_esikSagimLt),"DİKKAT","UYGUN"))))'),
        (54, "ssy_kararGerekce",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Karar",ssy_kararKapi,"| ort sağım",'
         'TEXT(ssy_gunlukSagimOrt,"0.0"),"| sağlık",TEXT(ssy_saglikSkor,"0"),'
         '"| alarm",ssy_alarmAdet)'),
        (55, "ssy_sonrakiKimlik",
         '="INK-"&TEXT(COUNTA(tblVarliklar[VarlikId])+1,"00000")'),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 57, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (58, "ssy_yorumGunluk",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Günlük sağım ort",TEXT(ssy_gunlukSagimOrt,"0.0"),"lt")'),
        (59, "ssy_yorumAylik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Aylık süt proxy",TEXT(ssy_aylikSutTrend,"0"),"lt")'),
        (60, "ssy_yorumSapma",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Sağım sapma",TEXT(ssy_sagimSapma,"0.0"),"lt")'),
        (61, "ssy_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(ssy_tornadoZirve,"0"),"lt")'),
        (62, "ssy_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(ssy_senaryoKarsilastirma,"0"))'),
        (63, "ssy_yorumHhi",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma HHI",TEXT(ssy_hhi,"0.0%"))'),
        (64, "ssy_yorumKalite",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite skoru",ssy_kaliteSkor)'),
        (65, "ssy_yorumSenMotor",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motoru",ssy_senaryoMotor)'),
        (66, "ssy_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(ssy_tahminAralik,"0"),"lt")'),
        (67, "ssy_yorumP90",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 sağım",TEXT(ssy_yuzdelikP90,"0"),"lt")'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 24, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. VARLIKLAR sayfasına inek kartı ekleyin; sonraki kimliği kullanın.",
        "2. GUNLUK_GIRIS: varlık, tarih, sağım lt, yem kg, kızgınlık, gebelik gün, olay (≤6 manuel alan).",
        "3. Sağım sapması = günlük sağım − inek hedefi; sağlık skoru formülle türer.",
        "4. DONEM_ANALIZ 12 aylık sezonsal süt-buzağılama zincirini gösterir.",
        "5. ALARMLAR sayfasında varlık + eşik + tarih listesi; eşikler AYARLAR'dadır.",
        "6. raporTarihi AYARLAR'dadır; TODAY kullanılmaz.",
        "7. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı hücreler açıktır.",
        "8. RAPOR sayfasını PDF olarak yazdırabilirsiniz.",
        f"9. Sürüm {SURUM} | Kod SSY-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
    ]
    for i, m in enumerate(maddeler, 5):
        h(ws, i, 1, m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    genislik(ws, {"A": 90})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 30):
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
            f"E{r}", CellIsRule(operator="lessThan", formula=["ssy_esikSagimLt"], fill=sari))
        ws.conditional_formatting.add(
            f"F{r}", CellIsRule(operator="lessThan", formula=["ssy_esikSaglik"], fill=kirmizi))

    ws = wb["GUNLUK_GIRIS"]
    ws.conditional_formatting.add(
        f"N{ILK}:N{SON}",
        CellIsRule(operator="equal", formula=['"ALARM"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"P{ILK}:P{SON}",
        CellIsRule(operator="equal", formula=['"YETIM"'], fill=sari))
    ws.conditional_formatting.add(
        f"D{ILK}:D{SON}",
        CellIsRule(operator="lessThan", formula=["ssy_esikSagimLt"], fill=sari))
    ws.conditional_formatting.add(
        f"M{ILK}:M{SON}",
        CellIsRule(operator="lessThan", formula=["ssy_esikSaglik"], fill=kirmizi))

    ws = wb["VARLIKLAR"]
    ws.conditional_formatting.add(
        f"G{ILK}:G{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"K{ILK}:K{SON}",
        CellIsRule(operator="lessThan", formula=["ssy_esikSaglik"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"I{ILK}:I{SON}",
        CellIsRule(operator="lessThan", formula=["ssy_esikSagimLt"], fill=sari))

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
            f"E{r}", CellIsRule(operator="lessThan", formula=["ssy_esikSagimLt"], fill=sari))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "ssy_varsayilanHedefLt", "ssy_esikSagimLt", "ssy_esikBuzagiGun",
        "ssy_esikSaglik", "ssy_azamiSagimLt", "ssy_aylikCarpan", "ssy_senaryoTemkinli",
        "ssy_senaryoBaz",
        "ssy_senaryoIyimser", "ssy_tornadoOran", "ssy_yuzdelikOran", "ssy_percentileOran",
        "ssy_tahminCarpan", "ssy_olcekHedef", "ssy_kaliteTaban", "ssy_kaliteCeza",
        "ssy_kaliteSapma", "ssy_dosyaSurumu", "ssy_firmaUnvan", "ssy_alarmEsikMetin",
        "ssy_sutFiyatTlLt", "ssy_kizginlikDongu", "ssy_gebelikSure",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "ssy_doluKayit": 32, "ssy_sagimToplam": 33, "ssy_yemToplam": 34,
        "ssy_gunlukSagimOrt": 35, "ssy_sagimSapma": 36, "ssy_sapmaOran": 37,
        "ssy_aylikSutTrend": 38, "ssy_saglikSkor": 39, "ssy_alarmAdet": 40,
        "ssy_yetimAdet": 41, "ssy_buzagiYakinAdet": 42, "ssy_tornadoZirve": 43,
        "ssy_senaryoKarsilastirma": 44, "ssy_hhi": 45, "ssy_kaliteSkor": 46,
        "ssy_senaryoMotor": 47, "ssy_tahminAralik": 48, "ssy_yuzdelikP90": 49,
        "ssy_olayKirilim": 50, "ssy_sezonsalTrend": 51, "ssy_alarmListesi": 52,
        "ssy_kararKapi": 53, "ssy_kararGerekce": 54, "ssy_sonrakiKimlik": 55,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "ssy_yorumGunluk": 58, "ssy_yorumAylik": 59, "ssy_yorumSapma": 60,
        "ssy_yorumTornado": 61, "ssy_yorumSenaryo": 62, "ssy_yorumHhi": 63,
        "ssy_yorumKalite": 64, "ssy_yorumSenMotor": 65, "ssy_yorumTahmin": 66,
        "ssy_yorumP90": 67,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$A$7:$A$8")
    ad_ekle(wb, "ListeOlayTuru", "LISTELER!$C$7:$C$12")
    ad_ekle(wb, "ListeIrk", "LISTELER!$E$7:$E$9")


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
    dosya_ad = "SutSurusuYonetim.xlsx"
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
        f.write("SatirId,VarlikId,Tarih,SagimLt_lt,YemKg_kg,KizginlikDurum,GebelikGun_gun,OlayTuru\n")
        for s in ORNEK:
            f.write(f"{s['id']},{s['varlik']},{s['tarih']:%d.%m.%Y},"
                    f"{s['sagim']},{s['yem']},{s['kiz']},{s['gebelik']},{s['olay']}\n")

    print(f"Dosya oluşturuldu: {hedef}")
    print(f"Kopya: {cikti}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
