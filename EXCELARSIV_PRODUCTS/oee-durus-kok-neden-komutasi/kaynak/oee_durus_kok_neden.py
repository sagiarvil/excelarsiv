#!/usr/bin/env python3
"""OEE & Duruş Kök-Neden Komutası — A5 saha komuta merkezi (manda v6).

OEE = kullanılabilirlik × performans × kalite.
Kullanılabilirlik = çalışma_dk / plan_dk.
Performans = üretim_adet / teorik_adet.
Kalite = (üretim − fire) / üretim.
Duruş kök-neden Pareto + varlık sağlık skoru formülden;
GUNLUK_GIRIS ≤6 manuel alan; ≥12 ay sezonsallık; AYARLAR alarm eşikleri.
Domain: OEE/duruş/kök neden — rename YASAK; SPC/filo/GES kopyası değil.
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

URUN_AD = "OEE & Duruş Kök-Neden Komutası"
SURUM = "1.0.0"
RENK = "1A365D"
KAPASITE = 250
DEMO_MAKINE = 8
DEMO_GUNLUK = 36
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

# (id, ad, hat, ideal_cevrim_sn, vardiya_dk, hedef_oee)
MAKINELER = [
    ("MAK-00001", "Pres Hat 1", "HAT-A", 12.0, 480, 0.75),
    ("MAK-00002", "Torna CNC 2", "HAT-A", 45.0, 480, 0.72),
    ("MAK-00003", "Montaj Hatti B", "HAT-B", 18.0, 480, 0.78),
    ("MAK-00004", "Kaynak Robot 4", "HAT-B", 30.0, 480, 0.70),
    ("MAK-00005", "Enjeksiyon 5", "HAT-C", 22.0, 480, 0.74),
    ("MAK-00006", "Paketleme 6", "HAT-C", 8.0, 480, 0.80),
    ("MAK-00007", "Kesim Lazer 7", "HAT-A", 35.0, 480, 0.71),
    ("MAK-00008", "Kalip Pres 8", "HAT-B", 15.0, 480, 0.76),
]

KOK_NEDENLER = ["ARIZA", "AYAR", "MALZEME", "OPERATOR", "KALIP", "BEKLEME", "DIGER"]
AYLAR = [
    "Ocak", "Subat", "Mart", "Nisan", "Mayis", "Haziran",
    "Temmuz", "Agustos", "Eylul", "Ekim", "Kasim", "Aralik",
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: OEE ve duruş kök-neden analizini etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_gunluk():
    satirlar = []
    for i in range(1, DEMO_GUNLUK + 1):
        mak = MAKINELER[(i - 1) % DEMO_MAKINE]
        tar = RAPOR_TARIHI - timedelta(days=(i * 3) % 340)
        plan = 480
        # düşük OEE örnekleri: i=12, 24
        durus = 180 if i in (12, 24) else (40 + (i % 7) * 8)
        ideal = mak[3]
        calisma = max(1, plan - durus)
        teorik = round(calisma * 60 / ideal)
        uretim = round(teorik * (0.55 if i in (12, 24) else (0.82 + (i % 5) * 0.02)))
        fire = round(uretim * (0.08 if i in (12, 24) else 0.02 + (i % 4) * 0.005))
        satirlar.append({
            "id": f"ODK-{i:05d}",
            "varlik": "" if i == 30 else (mak[0] if i != 33 else "MAK-99999"),
            "tarih": tar,
            "plan": plan,
            "durus": durus,
            "uretim": uretim,
            "fire": fire,
            "kok": KOK_NEDENLER[i % len(KOK_NEDENLER)],
        })
    return satirlar


ORNEK = _demo_gunluk()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Üretim hatlarında günlük plan süre, duruş, üretim ve fire verisinden OEE'yi "
      "hesaplar; duruş kök-neden Pareto kırılımı ve varlık sağlık skoru üretir; "
      "UYGUN / DİKKAT / KRİTİK / VERİ YOK kararı verir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "OEE = kullanılabilirlik × performans × kalite — tek formül, sabit skor yok",
        "Günlük saha girişi en fazla 6 manuel alan — 30 saniye hedefi",
        "Duruş kök-neden Pareto (ARIZA / AYAR / MALZEME / OPERATÖR / KALIP / BEKLEME)",
        "12 aylık sezonsal OEE-duruş zinciri ve alarm eşiği kuyruğu",
        "Makrosuz, çevrimdışı; MES / SCADA API bağı yok",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Vardiya / hat operatörü — günlük plan-duruş-üretim-fire-kök neden satırı",
        "Üretim müdürü / OEE sahibi — OEE panosu ve karar",
        "Kalite müdürü / fabrika yöneticisi — dönemsel RAPOR kanıt çıktısı",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) VARLIKLAR'a makine kartı girin. 2) GUNLUK_GIRIS'e günlük satır yazın. "
      "3) DONEM_ANALIZ ve PANO'dan kararı görün. 4) RAPOR'dan çıktı alın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Kod: ODK-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def varliklar(ws):
    sayfa_hazirla(ws, "VARLIKLAR", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Makine / hat kartları — her varlığın tek kimliği vardır", son_kolon=14)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=odk_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4",
               "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: MAK-00009 | Risk: Elle çakışma")

    sutunlar = [
        "VarlikId", "MakineAd", "Hat", "IdealCevrim_sn",
        "VardiyaSure_dk", "HedefOee_oran", "Aktif",
        "ToplamDurus_dk", "OrtOee", "ToplamUretim_adet", "SaglikSkor",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "ToplamDurus_dk": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'SUMIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[DurusSure_dk]))'
        ),
        "OrtOee": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'IFERROR(AVERAGEIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[FormulOee_oran]),0))'
        ),
        "ToplamUretim_adet": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'SUMIF(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[UretimAdet_adet]))'
        ),
        "SaglikSkor": (
            'IF(tblVarliklar[[#This Row],[VarlikId]]="","",'
            'MAX(0,MIN(100,odk_kaliteTaban'
            '-IF(tblVarliklar[[#This Row],[OrtOee]]<odk_esikOee,odk_kaliteCeza,0)'
            '-COUNTIFS(tblGunluk[VarlikId],tblVarliklar[[#This Row],[VarlikId]],'
            'tblGunluk[AlarmBayrak],"ALARM")*odk_kaliteSapma'
            '-IF(tblVarliklar[[#This Row],[Aktif]]="HAYIR",odk_kaliteCeza,0))))'
        ),
    }
    for i, row in enumerate(MAKINELER):
        r = ILK + i
        kid, ad, hat, ideal, vardiya, hedef = row
        _sari(ws, r, 1, kid, baslik="Varlık Kimliği", mesaj="MAK-##### formatında")
        _sari(ws, r, 2, ad, baslik="Makine Adı", mesaj="Makine / hat adı")
        _sari(ws, r, 3, hat, baslik="Hat", mesaj="Üretim hattı kodu")
        _sari(ws, r, 4, ideal, sayi="0.0", baslik="İdeal Çevrim [sn]", mesaj="Parça başına sn")
        _sari(ws, r, 5, vardiya, sayi=CATI, baslik="Vardiya [dk]", mesaj="Planlı vardiya dakikası")
        _sari(ws, r, 6, hedef, sayi=YÜZDE, baslik="Hedef OEE", mesaj="Hedef OEE oranı")
        _sari(ws, r, 7, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_MAKINE, SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None,
                  sayi=(YÜZDE if c == 6 else ("0.0" if c == 4 else (CATI if c == 5 else None))),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVarliklar", f"A{HDR}:K{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Varlık Kimliği", mesaj="MAK-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Makine Ad"), (3, "Hat")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    for c, ad in [(4, "İdeal Çevrim [sn]"), (5, "Vardiya [dk]"), (6, "Hedef OEE")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeEvetHayir", f"G{ILK}:G{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 12)})
    ws.column_dimensions["B"].width = 22


def gunluk_giris(ws):
    sayfa_hazirla(ws, "GUNLUK_GIRIS", RENK, URUN_AD, son_kolon=22)
    alt_bant(ws, 3,
             "GÜNLÜK saha girişi — satır başına en fazla 6 manuel alan (S02)",
             son_kolon=16)
    # Manuel (id/tarih/formul/skor/alarm hariç): VarlikId, Plan, Durus, Uretim, Fire, KokNeden = 6
    sutunlar = [
        "SatirId", "VarlikId", "Tarih", "PlanSure_dk", "DurusSure_dk",
        "UretimAdet_adet", "FireAdet_adet", "KokNedenKodu",
        "FormulIdealCevrim_sn", "FormulCalisma_dk", "FormulKullanilabilirlik_oran",
        "FormulTeorikAdet_adet", "FormulPerformans_oran", "FormulKalite_oran",
        "FormulOee_oran", "FormulDurusPay_oran", "SaglikSkor_puan",
        "AlarmBayrak", "FormulKayitDolu", "FormulYetim",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "FormulIdealCevrim_sn": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR(SUMIF(tblVarliklar[VarlikId],tblGunluk[[#This Row],[VarlikId]],'
            'tblVarliklar[IdealCevrim_sn]),odk_varsayilanCevrim))'
        ),
        "FormulCalisma_dk": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR(MAX(0,tblGunluk[[#This Row],[PlanSure_dk]]-'
            'tblGunluk[[#This Row],[DurusSure_dk]]),0))'
        ),
        "FormulKullanilabilirlik_oran": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR(tblGunluk[[#This Row],[FormulCalisma_dk]]/'
            'tblGunluk[[#This Row],[PlanSure_dk]],0))'
        ),
        "FormulTeorikAdet_adet": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR(tblGunluk[[#This Row],[FormulCalisma_dk]]*odk_dakikaSn/'
            'tblGunluk[[#This Row],[FormulIdealCevrim_sn]],0))'
        ),
        "FormulPerformans_oran": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR(tblGunluk[[#This Row],[UretimAdet_adet]]/'
            'tblGunluk[[#This Row],[FormulTeorikAdet_adet]],0))'
        ),
        "FormulKalite_oran": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR((tblGunluk[[#This Row],[UretimAdet_adet]]-'
            'tblGunluk[[#This Row],[FireAdet_adet]])/'
            'tblGunluk[[#This Row],[UretimAdet_adet]],0))'
        ),
        "FormulOee_oran": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR(N(tblGunluk[[#This Row],[FormulKullanilabilirlik_oran]])*'
            'N(tblGunluk[[#This Row],[FormulPerformans_oran]])*'
            'N(tblGunluk[[#This Row],[FormulKalite_oran]]),0))'
        ),
        "FormulDurusPay_oran": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR(tblGunluk[[#This Row],[DurusSure_dk]]/'
            'tblGunluk[[#This Row],[PlanSure_dk]],0))'
        ),
        "SaglikSkor_puan": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IFERROR(MAX(0,MIN(100,odk_kaliteTaban'
            '-IF(N(tblGunluk[[#This Row],[FormulOee_oran]])<odk_esikOee,odk_kaliteCeza,0)'
            '-IF(tblGunluk[[#This Row],[DurusSure_dk]]>odk_esikDurusDk,odk_kaliteSapma,0)'
            '-IF(tblGunluk[[#This Row],[FormulYetim]]="YETIM",odk_kaliteCeza,0))),0))'
        ),
        "AlarmBayrak": (
            'IF(tblGunluk[[#This Row],[SatirId]]="","",'
            'IF(OR(N(tblGunluk[[#This Row],[FormulOee_oran]])<odk_esikOee,'
            'tblGunluk[[#This Row],[DurusSure_dk]]>odk_esikDurusDk,'
            'N(tblGunluk[[#This Row],[SaglikSkor_puan]])<odk_esikSaglik),'
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
        _sari(ws, r, 1, s["id"], baslik="Günlük Kimlik", mesaj="ODK-#####")
        _sari(ws, r, 2, s["varlik"], baslik="Varlık Kimliği", mesaj="MAK listesinden")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="GG.AA.YYYY")
        _sari(ws, r, 4, s["plan"], sayi=CATI, baslik="Plan Süre [dk]", mesaj="Planlı üretim süresi")
        _sari(ws, r, 5, s["durus"], sayi=CATI, baslik="Duruş [dk]", mesaj="Toplam duruş dakikası")
        _sari(ws, r, 6, s["uretim"], sayi=CATI, baslik="Üretim [adet]", mesaj="Toplam üretilen adet")
        _sari(ws, r, 7, s["fire"], sayi=CATI, baslik="Fire [adet]", mesaj="Fire / hurda adedi")
        _sari(ws, r, 8, s["kok"], baslik="Kök Neden", mesaj="Duruş kök-neden kodu")
    for r in range(ILK + DEMO_GUNLUK, SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None,
                  sayi=(TARİH if c == 3 else (CATI if c in (4, 5, 6, 7) else None)),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblGunluk", f"A{HDR}:T{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Günlük Kimlik", mesaj="ODK-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Varlık", mesaj="MAK kimliği girin",
              hata_baslik="Varlık", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"C{ILK}:C{SON}",
              baslik="Tarih", mesaj="GG.AA.YYYY",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    for c, ad in [(4, "Plan [dk]"), (5, "Duruş [dk]"), (6, "Üretim [adet]"), (7, "Fire [adet]")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeKokNeden", f"H{ILK}:H{SON}",
              baslik="Kök Neden", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 13 for i in range(1, 21)})


def donem_analiz(ws):
    sayfa_hazirla(ws, "DONEM_ANALIZ", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "12 aylık sezonsallık — OEE ve duruş zinciri (S03)", son_kolon=12)
    h(ws, 5, 1, "Ay", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Ay_Anahtari", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Durus dk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 4, "Uretim adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 5, "OEE Ort", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 6, "Sezonsal Etiket", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    demo_durus = [9200, 8800, 7600, 7200, 6800, 6400,
                  7000, 7400, 7800, 8200, 8600, 9000]
    demo_uretim = [42000, 45000, 51000, 54000, 58000, 60000,
                   57000, 55000, 52000, 48000, 44000, 41000]
    demo_oee = [0.62, 0.65, 0.70, 0.72, 0.75, 0.78,
                0.74, 0.72, 0.70, 0.67, 0.64, 0.61]
    for i, ay in enumerate(AYLAR, 1):
        r = 5 + i
        h(ws, r, 1, ay)
        h(ws, r, 2, f"Ay_{i}")
        h(ws, r, 3, demo_durus[i - 1], sayi=CATI)
        h(ws, r, 4, demo_uretim[i - 1], sayi=CATI)
        h(ws, r, 5, demo_oee[i - 1], sayi=YÜZDE)
        h(ws, r, 6, "12_AY_SEZONSAL")

    h(ws, 20, 1, "Sezonsal Trend Özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=odk_sezonsalTrend", kaydir=True)
    ws.merge_cells("B20:F20")
    h(ws, 22, 1, "12 Ay Toplam Duruş", kalin=True)
    h(ws, 22, 2, "=SUM(C6:C17)", sayi=CATI)
    h(ws, 23, 1, "12 Ay Toplam Üretim", kalin=True)
    h(ws, 23, 2, "=SUM(D6:D17)", sayi=CATI)
    h(ws, 24, 1, "12 Ay Ort OEE", kalin=True)
    h(ws, 24, 2, "=AVERAGE(E6:E17)", sayi=YÜZDE)
    h(ws, 25, 1, "Canlı dönem duruş", kalin=True)
    h(ws, 25, 2, "=odk_toplamDurus", sayi=CATI)
    h(ws, 26, 1, "Canlı dönem OEE", kalin=True)
    h(ws, 26, 2, "=odk_oeeOrani", sayi=YÜZDE)

    c1 = LineChart()
    c1.title = "12 Ay Duruş Sezonsallığı"
    c1.add_data(Reference(ws, min_col=3, min_row=5, max_row=17), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17))
    c1.width, c1.height = 14, 8
    ws.add_chart(c1, "A27")

    c2 = BarChart()
    c2.title = "12 Ay OEE Dağılımı"
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
        ("H03", "Toplam plan dk", "=SUM(tblGunluk[PlanSure_dk])"),
        ("H04", "Toplam duruş dk", "=SUM(tblGunluk[DurusSure_dk])"),
        ("H05", "Toplam üretim adet", "=SUM(tblGunluk[UretimAdet_adet])"),
        ("H06", "Toplam fire adet", "=SUM(tblGunluk[FireAdet_adet])"),
        ("H07", "Toplam çalışma dk", "=SUM(tblGunluk[FormulCalisma_dk])"),
        ("H08", "Kullanılabilirlik", "=IFERROR(C12/MAX(C8,1),0)"),
        ("H09", "Toplam teorik adet", "=SUM(tblGunluk[FormulTeorikAdet_adet])"),
        ("H10", "Performans", "=IFERROR(C10/MAX(C14,1),0)"),
        ("H11", "Kalite oranı", "=IFERROR((C10-C11)/MAX(C10,1),0)"),
        ("H12", "OEE oranı", "=IFERROR(C13*C15*C16,0)"),
        ("H13", "Duruş payı", "=IFERROR(C9/MAX(C8,1),0)"),
        ("H14", "Aktif makine adet", '=COUNTIF(tblVarliklar[Aktif],"EVET")'),
        ("H15", "Toplam makine kart", '=COUNTA(tblVarliklar[VarlikId])'),
        ("H16", "Alarm adet", '=COUNTIF(tblGunluk[AlarmBayrak],"ALARM")'),
        ("H17", "Yetim kayıt", '=COUNTIF(tblGunluk[FormulYetim],"YETIM")'),
        ("H18", "Ortalama sağlık skor",
         '=IFERROR(AVERAGEIF(tblGunluk[SaglikSkor_puan],">0"),0)'),
        ("H19", "Min sağlık skor",
         '=IFERROR(MIN(tblGunluk[SaglikSkor_puan]),0)'),
        ("H20", "OEE eşik altı adet",
         '=COUNTIFS(tblGunluk[FormulOee_oran],"<"&odk_esikOee,tblGunluk[FormulKayitDolu],">0")'),
        ("H21", "Duruş eşik aşımı",
         '=COUNTIFS(tblGunluk[DurusSure_dk],">"&odk_esikDurusDk,tblGunluk[FormulKayitDolu],">0")'),
        ("H22", "Düşük sağlık adet",
         '=COUNTIFS(tblGunluk[SaglikSkor_puan],"<"&odk_esikSaglik,tblGunluk[FormulKayitDolu],">0")'),
        ("H23", "Günlük OEE ort",
         '=IFERROR(AVERAGEIF(tblGunluk[FormulOee_oran],">0"),0)'),
        ("H24", "Aylık duruş proxy", "=C9*odk_aylikCarpan"),
        ("H25", "OEE sapma",
         "=IFERROR(ABS(C17-odk_esikOee),0)"),
        ("H26", "ARIZA kök pay",
         '=IFERROR(COUNTIF(tblGunluk[KokNedenKodu],"ARIZA")/MAX(C7,1),0)'),
        ("H27", "AYAR kök pay",
         '=IFERROR(COUNTIF(tblGunluk[KokNedenKodu],"AYAR")/MAX(C7,1),0)'),
        ("H28", "MALZEME kök pay",
         '=IFERROR(COUNTIF(tblGunluk[KokNedenKodu],"MALZEME")/MAX(C7,1),0)'),
        ("H29", "OPERATOR kök pay",
         '=IFERROR(COUNTIF(tblGunluk[KokNedenKodu],"OPERATOR")/MAX(C7,1),0)'),
        ("H30", "Temkinli senaryo", "=C17*odk_senaryoTemkinli"),
        ("H31", "Baz senaryo", "=C17*odk_senaryoBaz"),
        ("H32", "İyimser senaryo", "=C17*odk_senaryoIyimser"),
        ("H33", "Senaryo farkı", "=C37-C35"),
        ("H34", "Tornado duruş etki", "=C9*odk_tornadoOran"),
        ("H35", "Tornado fire etki", "=C11*odk_tornadoOran"),
        ("H36", "Tornado OEE etki", "=C17*odk_tornadoOran"),
        ("H37", "Tornado zirve", "=MAX(C39,C40,C41)"),
        ("H38", "HHI proxy",
         "=IFERROR(C31^2+C32^2+C33^2+C34^2,0)"),
        ("H39", "Kalite skor",
         "=MAX(0,odk_kaliteTaban-C25*odk_kaliteCeza-C26*odk_kaliteSapma"
         "-C27*odk_kaliteSapma)"),
        ("H40", "Tahmin aralık", "=C36*odk_tahminCarpan"),
        ("H41", "P90 OEE",
         "=IFERROR(PERCENTILE(tblGunluk[FormulOee_oran],odk_percentileOran),C17*odk_yuzdelikOran)"),
        ("H42", "Senaryo motor metin",
         '=_xlfn.TEXTJOIN("|",TRUE,"T",TEXT(C35,"0.0%"),"B",TEXT(C36,"0.0%"),"I",TEXT(C37,"0.0%"))'),
        ("H43", "Sezonsal 12 ay duruş", "=DONEM_ANALIZ!B22"),
        ("H44", "Sezonsal 12 ay OEE", "=DONEM_ANALIZ!B24"),
        ("H45", "Kritik sinyal",
         '=IF(OR(C25>0,C27>0,C23<odk_esikSaglik),1,0)'),
        ("H46", "Karar kodu ham",
         '=IF(C7=0,0,IF(C50=1,3,IF(OR(C25>0,C17<odk_esikOee),2,1)))'),
        ("H47", "OEE çıpa", "=C17"),
        ("H48", "Alarm listesi özet adet", "=C21"),
        ("H49", "Varlık sağlık ort",
         '=IFERROR(AVERAGEIF(tblVarliklar[SaglikSkor],">0"),0)'),
        ("H50", "Tahmin FORECAST",
         "=IFERROR(FORECAST(C7+1,tblGunluk[FormulOee_oran],tblGunluk[FormulKayitDolu]),C45)"),
        ("H51", "Kök neden kırılım özet",
         '=_xlfn.TEXTJOIN("|",TRUE,"A",TEXT(C31,"0%"),"Y",TEXT(C32,"0%"),'
         '"M",TEXT(C33,"0%"),"O",TEXT(C34,"0%"))'),
    ]
    # H01=C6 … H51=C56
    # H01=C6 H02=C7 H03=C8 H04=C9 H05=C10 H06=C11 H07=C12 H08=C13 H09=C14
    # H10=C15 H11=C16 H12=C17 H13=C18 H14=C19 H15=C20 H16=C21 H17=C22 H18=C23
    # H19=C24 H20=C25 H21=C26 H22=C27 H23=C28 H24=C29 H25=C30
    # H26=C31 H27=C32 H28=C33 H29=C34 H30=C35 H31=C36 H32=C37 H33=C38
    # H34=C39 H35=C40 H36=C41 H37=C42 H38=C43 H39=C44 H40=C45 H41=C46
    # H42=C47 H43=C48 H44=C49 H45=C50 H46=C51 H47=C52 H48=C53 H49=C54
    # H50=C55 H51=C56

    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        acik_l = acik.lower()
        sayi = YÜZDE if any(x in acik_l for x in (
            "oee", "oran", "pay", "sapma", "hhi", "kullanılabilirlik",
            "performans", "kalite oranı", "senaryo",
        )) else CATI
        if "metin" in acik_l or "özet" in acik_l or "kırılım" in acik_l:
            sayi = None
        elif any(x in acik_l for x in ("adet", "dk", "proxy", "tornado", "tahmin", "p90")) and (
            "oran" not in acik_l and "oee" not in acik_l and "pay" not in acik_l
            and "senaryo" not in acik_l and "metin" not in acik_l
        ):
            sayi = CATI
        h(ws, r, 3, form, sayi=sayi)

    genislik(ws, {"A": 8, "B": 32, "C": 22})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Karar panosu — OEE, duruş, kök neden, sezonsal trend", son_kolon=14)

    h(ws, 5, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 5, 2, "=odk_kararKapi", kalin=True, boyut=16)
    yorum_ekle(ws, "B5",
               "Tanım: Karar kapısı | Neden önemli: Boş dosyada VERİ YOK | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle müdahale yok")

    kpis = [
        (7, "Dolu Kayıt", "=odk_doluKayit", CATI),
        (8, "Toplam Duruş dk", "=odk_toplamDurus", CATI),
        (9, "Toplam Üretim adet", "=odk_toplamUretim", CATI),
        (10, "OEE Oranı", "=odk_oeeOrani", YÜZDE),
        (11, "Kullanılabilirlik", "=odk_kullanilabilirlik", YÜZDE),
        (12, "Performans", "=odk_performans", YÜZDE),
        (13, "Kalite Oranı", "=odk_kaliteOrani", YÜZDE),
        (14, "Günlük OEE Ort", "=odk_gunlukOeeOrt", YÜZDE),
        (15, "Sağlık Skor Ort", "=odk_saglikSkor", CATI),
        (16, "Alarm Adet", "=odk_alarmAdet", CATI),
        (17, "Yetim Kayıt", "=odk_yetimAdet", CATI),
        (18, "Kalite Skor", "=odk_kaliteSkor", CATI),
        (19, "Tornado Zirve", "=odk_tornadoZirve", CATI),
        (20, "Kök Neden Kırılım", "=odk_kokNedenKirilim", None),
    ]
    h(ws, 6, 1, "Gösterge", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 6, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for r, ad, form, sayi in kpis:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=sayi, kalin=True)

    h(ws, 7, 4, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 5, "=odk_kararGerekce", kaydir=True)
    ws.merge_cells("E7:H8")
    h(ws, 10, 4, "Öneri 1", kalin=True)
    h(ws, 10, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"OEE",TEXT(odk_oeeOrani,"0.0%"),'
                 '"— eşik",TEXT(odk_esikOee,"0.0%"))', kaydir=True)
    h(ws, 11, 4, "Öneri 2", kalin=True)
    h(ws, 11, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Alarm",odk_alarmAdet,"adet — sağlık",'
                 'TEXT(odk_saglikSkor,"0"))', kaydir=True)
    h(ws, 12, 4, "Öneri 3", kalin=True)
    h(ws, 12, 5, '=_xlfn.TEXTJOIN(" ",TRUE,"Duruş",TEXT(odk_toplamDurus,"0"),'
                 '"dk — kök",odk_kokNedenKirilim)', kaydir=True)

    h(ws, 22, 1, "KokNeden")
    h(ws, 22, 2, "Adet")
    for i, (ad, t) in enumerate([
        ("ARIZA", 9), ("AYAR", 7), ("MALZEME", 6), ("OPERATOR", 5),
    ], 23):
        h(ws, i, 1, ad)
        h(ws, i, 2, t, sayi=CATI)

    h(ws, 22, 4, "Senaryo")
    h(ws, 22, 5, "OEE")
    for i, (ad, t) in enumerate([
        ("Temkinli", 0.58), ("Baz", 0.68), ("İyimser", 0.78),
    ], 23):
        h(ws, i, 4, ad)
        h(ws, i, 5, t, sayi=YÜZDE)

    h(ws, 22, 7, "Hafta")
    h(ws, 22, 8, "Durus")
    for i in range(1, 9):
        h(ws, 22 + i, 7, f"H{i}")
        h(ws, 22 + i, 8, 120 + i * 15, sayi=CATI)

    h(ws, 22, 10, "Ay")
    h(ws, 22, 11, "OEE")
    for i, oee in enumerate([0.62, 0.68, 0.72, 0.76], 23):
        h(ws, i, 10, AYLAR[i - 23 + 5])
        h(ws, i, 11, oee, sayi=YÜZDE)

    c1 = PieChart()
    c1.title = "Kök Neden Dağılımı"
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
    c3.title = "Haftalık Duruş Trend"
    c3.add_data(Reference(ws, min_col=8, min_row=22, max_row=30), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=7, min_row=23, max_row=30))
    c3.width, c3.height = 12, 7
    ws.add_chart(c3, "A47")

    c4 = BarChart()
    c4.title = "Aylık OEE"
    c4.add_data(Reference(ws, min_col=11, min_row=22, max_row=26), titles_from_data=True)
    c4.set_categories(Reference(ws, min_col=10, min_row=23, max_row=26))
    c4.width, c4.height = 10, 7
    ws.add_chart(c4, "F47")

    c5 = PieChart()
    c5.title = "Kök Neden (tekrar)"
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
    c7.title = "Duruş Trend (kısa)"
    c7.add_data(Reference(ws, min_col=8, min_row=22, max_row=28), titles_from_data=True)
    c7.set_categories(Reference(ws, min_col=7, min_row=23, max_row=28))
    c7.width, c7.height = 12, 7
    ws.add_chart(c7, "A77")

    c8 = BarChart()
    c8.title = "Kök Neden Adet"
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
    h(ws, 5, 5, "OEE Ort", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 6, "Sağlık Skor", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    for i, mak in enumerate(MAKINELER):
        r = 6 + i
        h(ws, r, 1, mak[0])
        h(ws, r, 2,
          f'=IF(OR(COUNTIFS(tblGunluk[VarlikId],A{r},tblGunluk[AlarmBayrak],"ALARM")>0,'
          f'IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[SaglikSkor]),100)<odk_esikSaglik),'
          f'"ALARM","YOK")')
        h(ws, r, 3, f"=IFERROR(ABS(E{r}-odk_esikOee),0)", sayi=YÜZDE)
        h(ws, r, 4, "=raporTarihi", sayi=TARİH)
        h(ws, r, 5,
          f'=IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[OrtOee]),0)',
          sayi=YÜZDE)
        h(ws, r, 6,
          f'=IFERROR(SUMIF(tblVarliklar[VarlikId],A{r},tblVarliklar[SaglikSkor]),0)',
          sayi=CATI)

    h(ws, 16, 1, "Alarm Özeti", kalin=True)
    h(ws, 16, 2, "=odk_alarmListesi", kaydir=True)
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
    h(ws, 8, 2, "=odk_firmaUnvan")

    h(ws, 10, 1, "KARAR", kalin=True, boyut=12)
    h(ws, 10, 2, "=odk_kararKapi", kalin=True, boyut=14)
    h(ws, 11, 1, "OEE Oranı")
    h(ws, 11, 2, "=odk_oeeOrani", sayi=YÜZDE)
    h(ws, 12, 1, "Toplam Duruş dk")
    h(ws, 12, 2, "=odk_toplamDurus", sayi=CATI)
    h(ws, 13, 1, "Toplam Üretim adet")
    h(ws, 13, 2, "=odk_toplamUretim", sayi=CATI)
    h(ws, 14, 1, "Sağlık Skor")
    h(ws, 14, 2, "=odk_saglikSkor", sayi=CATI)
    h(ws, 15, 1, "Alarm Adet")
    h(ws, 15, 2, "=odk_alarmAdet", sayi=CATI)

    h(ws, 17, 1, "Kalem", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 17, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, (ad, form, sayi) in enumerate([
        ("Kullanılabilirlik", "=odk_kullanilabilirlik", YÜZDE),
        ("Performans", "=odk_performans", YÜZDE),
        ("Kalite oranı", "=odk_kaliteOrani", YÜZDE),
        ("Günlük OEE ort", "=odk_gunlukOeeOrt", YÜZDE),
        ("Kalite skor", "=odk_kaliteSkor", CATI),
        ("Kök neden kırılım", "=odk_kokNedenKirilim", None),
    ], 18):
        h(ws, i, 1, ad)
        h(ws, i, 2, form, sayi=sayi)

    h(ws, 26, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1,
      "OEE = kullanılabilirlik × performans × kalite. Eşikler AYARLAR'dan gelir. "
      "Karar destek çıktısıdır; sertifikasyon veya kalite tavsiyesi değildir. "
      "Duruş kök-neden kodları saha sınıflandırmasıdır.",
      kaydir=True)
    ws.merge_cells("A27:F28")

    h(ws, 30, 1, "İmza — Üretim Müdürü / OEE Sahibi")
    h(ws, 30, 3, "________________")
    h(ws, 31, 1, "İmza — Kalite Müdürü / Fabrika Yöneticisi")
    h(ws, 31, 3, "________________")
    h(ws, 33, 1, f"SHA notu: sevk paketinde SHA-256 yer alır | {SURUM}", yazi=GRİ, boyut=9)

    baski_hazirla(ws, "A1:F35", f"{URUN_AD} | Rapor | {SURUM}")
    genislik(ws, {"A": 52, "B": 28, "C": 18})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Canlı sağlık kontrolleri — formülle", son_kolon=10)
    h(ws, 5, 1, "Kontrol", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Sonuç", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    kontroller_list = [
        ("Dolu kayıt > 0", '=IF(odk_doluKayit>0,"GEÇTİ","KALDI")', "=odk_doluKayit"),
        ("Yetim kayıt = 0", '=IF(odk_yetimAdet=0,"GEÇTİ","UYARI")', "=odk_yetimAdet"),
        ("OEE hesaplı", '=IF(odk_oeeOrani>=0,"GEÇTİ","KALDI")', "=odk_oeeOrani"),
        ("Sağlık skor formül", '=IF(odk_saglikSkor>=0,"GEÇTİ","KALDI")', "=odk_saglikSkor"),
        ("Alarm listesi", '=IF(odk_alarmAdet>=0,"GEÇTİ","KALDI")', "=odk_alarmAdet"),
        ("Karar dolu", '=IF(odk_kararKapi<>"","GEÇTİ","KALDI")', "=odk_kararKapi"),
        ("12 ay sezonsal", '=IF(DONEM_ANALIZ!B22>=0,"GEÇTİ","KALDI")', "=DONEM_ANALIZ!B22"),
        ("Kalite skor", '=IF(odk_kaliteSkor>=0,"GEÇTİ","KALDI")', "=odk_kaliteSkor"),
        ("Tornado", '=IF(odk_tornadoZirve>=0,"GEÇTİ","KALDI")', "=odk_tornadoZirve"),
        ("P90", '=IF(odk_yuzdelikP90>=0,"GEÇTİ","KALDI")', "=odk_yuzdelikP90"),
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
        "SatirId", "VarlikId", "Tarih", "PlanSure_dk", "DurusSure_dk",
        "UretimAdet_adet", "FireAdet_adet", "KokNedenKodu",
    ]
    baslik_satiri(ws, 5, basliklar)
    for i, s in enumerate(ORNEK[:20]):
        r = 6 + i
        h(ws, r, 1, s["id"])
        h(ws, r, 2, s["varlik"])
        h(ws, r, 3, s["tarih"], sayi=TARİH)
        h(ws, r, 4, s["plan"], sayi=CATI)
        h(ws, r, 5, s["durus"], sayi=CATI)
        h(ws, r, 6, s["uretim"], sayi=CATI)
        h(ws, r, 7, s["fire"], sayi=CATI)
        h(ws, r, 8, s["kok"])
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 9)})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Evet / Hayır", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "Secim", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    _sari(ws, 7, 1, "EVET", baslik="Seçim", mesaj="EVET/HAYIR")
    _sari(ws, 8, 1, "HAYIR", baslik="Seçim", mesaj="EVET/HAYIR")

    h(ws, 3, 3, "Kök Neden", kalin=True)
    h(ws, 6, 3, "Kod", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, t in enumerate(KOK_NEDENLER, 7):
        _sari(ws, i, 3, t, baslik="Kök Neden", mesaj="Duruş kök-neden kodu")

    h(ws, 3, 5, "Hat Tipi", kalin=True)
    h(ws, 6, 5, "Hat", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, t in enumerate(["HAT-A", "HAT-B", "HAT-C"], 7):
        _sari(ws, i, 5, t, baslik="Hat", mesaj="Hat kodu")
    genislik(ws, {"A": 12, "C": 14, "E": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + alarm eşikleri (S05)", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("odk_varsayilanCevrim", 20.0, "sn", "Makine kartı / iş etüdü", "01.01.2026",
         "Varsayılan ideal çevrim sn"),
        ("odk_esikOee", 0.65, "oran", "Alarm eşiği — OEE", "01.01.2026",
         "OEE alt alarm eşiği"),
        ("odk_esikDurusDk", 120.0, "dk", "Alarm eşiği — duruş", "01.01.2026",
         "Günlük duruş alarm eşiği"),
        ("odk_esikSaglik", 60, "puan", "Alarm eşiği — sağlık skor", "01.01.2026",
         "Sağlık skor alt eşiği"),
        ("odk_aylikCarpan", 1.0, "oran", "Aylık proxy", "01.01.2026", "Aylık duruş çarpan"),
        ("odk_senaryoTemkinli", 0.90, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("odk_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("odk_senaryoIyimser", 1.10, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("odk_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("odk_yuzdelikOran", 1.15, "oran", "İstatistik kuralı", "01.01.2026", "P90 çarpan proxy"),
        ("odk_percentileOran", 0.9, "oran", "PERCENTILE oranı", "01.01.2026", "Yüzdelik dilim"),
        ("odk_tahminCarpan", 1.05, "oran", "Tahmin modeli", "01.01.2026", "Tahmin aralık çarpan"),
        ("odk_olcekHedef", 20000, "satir", "Manda A5 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("odk_kaliteTaban", 100, "puan", "Kalite / sağlık modeli", "01.01.2026", "Taban puan"),
        ("odk_kaliteCeza", 12, "puan", "Kalite / sağlık modeli", "01.01.2026", "Eşik aşım cezası"),
        ("odk_kaliteSapma", 3, "puan", "Kalite / sağlık modeli", "01.01.2026", "Sapma / alarm cezası"),
        ("odk_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("odk_firmaUnvan", "Örnek Üretim A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma unvanı"),
        ("odk_alarmEsikMetin", "OEE / duruş / sağlık eşikleri", "metin",
         "S05 alarm", "01.01.2026", "Alarm eşiği tanımı"),
        ("odk_dakikaSn", 60, "sn/dk", "Birim dönüşüm", "01.01.2026",
         "Dakika → saniye çarpanı"),
        ("odk_oeeReferans", 0.75, "oran", "Hedef OEE referans", "01.01.2026",
         "Referans OEE hedefi"),
        ("odk_durusMaliyetTlDk", 18.5, "TL/dk", "İç maliyet modeli", "01.01.2026",
         "Duruş dakikası maliyeti"),
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
        (32, "odk_doluKayit", "=HESAP!C7"),
        (33, "odk_toplamDurus", "=HESAP!C9"),
        (34, "odk_toplamUretim", "=HESAP!C10"),
        (35, "odk_oeeOrani", "=HESAP!C17"),
        (36, "odk_kullanilabilirlik", "=HESAP!C13"),
        (37, "odk_performans", "=HESAP!C15"),
        (38, "odk_kaliteOrani", "=HESAP!C16"),
        (39, "odk_gunlukOeeOrt", "=HESAP!C28"),
        (40, "odk_aylikDurusTrend", "=HESAP!C29"),
        (41, "odk_saglikSkor", "=HESAP!C23"),
        (42, "odk_alarmAdet", "=HESAP!C21"),
        (43, "odk_yetimAdet", "=HESAP!C22"),
        (44, "odk_oeeSapma", "=HESAP!C30"),
        (45, "odk_tornadoZirve", "=HESAP!C42"),
        (46, "odk_senaryoKarsilastirma", "=HESAP!C38"),
        (47, "odk_hhi", "=HESAP!C43"),
        (48, "odk_kaliteSkor", "=HESAP!C44"),
        (49, "odk_senaryoMotor", "=HESAP!C47"),
        (50, "odk_tahminAralik", "=HESAP!C45"),
        (51, "odk_yuzdelikP90", "=HESAP!C46"),
        (52, "odk_kokNedenKirilim", "=HESAP!C56"),
        (53, "odk_sezonsalTrend",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"12_AY",TEXT(HESAP!C48,"0"),"dk",'
         'TEXT(HESAP!C49,"0.0%"),"OEE sezonsal")'),
        (54, "odk_alarmListesi",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Alarm",HESAP!C21,"Yetim",HESAP!C22,'
         '"OeeAlt",HESAP!C25,"DurusAsim",HESAP!C26)'),
        (55, "odk_kararKapi",
         '=IF(HESAP!C7=0,"VERİ YOK",'
         'IF(OR(HESAP!C9<0,HESAP!C10<0),"KRİTİK",'
         'IF(HESAP!C50=1,"KRİTİK",'
         'IF(OR(HESAP!C25>0,HESAP!C17<odk_esikOee),"DİKKAT","UYGUN"))))'),
        (56, "odk_kararGerekce",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Karar",odk_kararKapi,"| OEE",'
         'TEXT(odk_oeeOrani,"0.0%"),"| sağlık",TEXT(odk_saglikSkor,"0"),'
         '"| alarm",odk_alarmAdet)'),
        (57, "odk_sonrakiKimlik",
         '="MAK-"&TEXT(COUNTA(tblVarliklar[VarlikId])+1,"00000")'),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 59, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (60, "odk_yorumGunluk",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Günlük OEE ort",TEXT(odk_gunlukOeeOrt,"0.0%"))'),
        (61, "odk_yorumAylik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Aylık duruş proxy",TEXT(odk_aylikDurusTrend,"0"),"dk")'),
        (62, "odk_yorumSapma",
         '=_xlfn.TEXTJOIN(" ",TRUE,"OEE sapma",TEXT(odk_oeeSapma,"0.0%"))'),
        (63, "odk_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(odk_tornadoZirve,"0"))'),
        (64, "odk_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(odk_senaryoKarsilastirma,"0.0%"))'),
        (65, "odk_yorumHhi",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma HHI",TEXT(odk_hhi,"0.0%"))'),
        (66, "odk_yorumKalite",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite skoru",odk_kaliteSkor)'),
        (67, "odk_yorumSenMotor",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motoru",odk_senaryoMotor)'),
        (68, "odk_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(odk_tahminAralik,"0.0%"))'),
        (69, "odk_yorumP90",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 OEE",TEXT(odk_yuzdelikP90,"0.0%"))'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 24, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. VARLIKLAR sayfasına makine kartı ekleyin; sonraki kimliği kullanın.",
        "2. GUNLUK_GIRIS: varlık, tarih, plan dk, duruş dk, üretim adet, fire adet, kök neden (≤6 manuel alan).",
        "3. OEE = kullanılabilirlik × performans × kalite formülüyle hesaplanır.",
        "4. DONEM_ANALIZ 12 aylık sezonsal OEE-duruş zincirini gösterir.",
        "5. ALARMLAR sayfasında varlık + eşik + tarih listesi; eşikler AYARLAR'dadır.",
        "6. raporTarihi AYARLAR'dadır; TODAY kullanılmaz.",
        "7. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı hücreler açıktır.",
        "8. RAPOR sayfasını PDF olarak yazdırabilirsiniz.",
        f"9. Sürüm {SURUM} | Kod ODK-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
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
            f"E{r}", CellIsRule(operator="lessThan", formula=["odk_esikOee"], fill=sari))
        ws.conditional_formatting.add(
            f"F{r}", CellIsRule(operator="lessThan", formula=["odk_esikSaglik"], fill=kirmizi))

    ws = wb["GUNLUK_GIRIS"]
    ws.conditional_formatting.add(
        f"R{ILK}:R{SON}",
        CellIsRule(operator="equal", formula=['"ALARM"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"T{ILK}:T{SON}",
        CellIsRule(operator="equal", formula=['"YETIM"'], fill=sari))
    ws.conditional_formatting.add(
        f"O{ILK}:O{SON}",
        CellIsRule(operator="lessThan", formula=["odk_esikOee"], fill=sari))
    ws.conditional_formatting.add(
        f"Q{ILK}:Q{SON}",
        CellIsRule(operator="lessThan", formula=["odk_esikSaglik"], fill=kirmizi))

    ws = wb["VARLIKLAR"]
    ws.conditional_formatting.add(
        f"G{ILK}:G{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"K{ILK}:K{SON}",
        CellIsRule(operator="lessThan", formula=["odk_esikSaglik"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"I{ILK}:I{SON}",
        CellIsRule(operator="lessThan", formula=["odk_esikOee"], fill=sari))

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
    for r in range(11, 25):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["HESAP"]
    for r in range(6, 57):
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi))

    ws = wb["DONEM_ANALIZ"]
    for r in range(6, 18):
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"E{r}", CellIsRule(operator="lessThan", formula=["odk_esikOee"], fill=sari))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "odk_varsayilanCevrim", "odk_esikOee", "odk_esikDurusDk",
        "odk_esikSaglik", "odk_aylikCarpan", "odk_senaryoTemkinli", "odk_senaryoBaz",
        "odk_senaryoIyimser", "odk_tornadoOran", "odk_yuzdelikOran", "odk_percentileOran",
        "odk_tahminCarpan", "odk_olcekHedef", "odk_kaliteTaban", "odk_kaliteCeza",
        "odk_kaliteSapma", "odk_dosyaSurumu", "odk_firmaUnvan", "odk_alarmEsikMetin",
        "odk_dakikaSn", "odk_oeeReferans", "odk_durusMaliyetTlDk",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "odk_doluKayit": 32, "odk_toplamDurus": 33, "odk_toplamUretim": 34,
        "odk_oeeOrani": 35, "odk_kullanilabilirlik": 36, "odk_performans": 37,
        "odk_kaliteOrani": 38, "odk_gunlukOeeOrt": 39, "odk_aylikDurusTrend": 40,
        "odk_saglikSkor": 41, "odk_alarmAdet": 42, "odk_yetimAdet": 43,
        "odk_oeeSapma": 44, "odk_tornadoZirve": 45, "odk_senaryoKarsilastirma": 46,
        "odk_hhi": 47, "odk_kaliteSkor": 48, "odk_senaryoMotor": 49,
        "odk_tahminAralik": 50, "odk_yuzdelikP90": 51, "odk_kokNedenKirilim": 52,
        "odk_sezonsalTrend": 53, "odk_alarmListesi": 54, "odk_kararKapi": 55,
        "odk_kararGerekce": 56, "odk_sonrakiKimlik": 57,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "odk_yorumGunluk": 60, "odk_yorumAylik": 61, "odk_yorumSapma": 62,
        "odk_yorumTornado": 63, "odk_yorumSenaryo": 64, "odk_yorumHhi": 65,
        "odk_yorumKalite": 66, "odk_yorumSenMotor": 67, "odk_yorumTahmin": 68,
        "odk_yorumP90": 69,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$A$7:$A$8")
    ad_ekle(wb, "ListeKokNeden", "LISTELER!$C$7:$C$13")
    ad_ekle(wb, "ListeHat", "LISTELER!$E$7:$E$9")


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
    dosya_ad = "OeeDurusKokNedenKomutasi.xlsx"
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
        f.write("SatirId,VarlikId,Tarih,PlanSure_dk,DurusSure_dk,"
                "UretimAdet_adet,FireAdet_adet,KokNedenKodu\n")
        for s in ORNEK:
            f.write(f"{s['id']},{s['varlik']},{s['tarih']:%d.%m.%Y},"
                    f"{s['plan']},{s['durus']},{s['uretim']},{s['fire']},{s['kok']}\n")

    print(f"Dosya oluşturuldu: {hedef}")
    print(f"Kopya: {cikti}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
