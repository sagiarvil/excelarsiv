#!/usr/bin/env python3
"""Mizan Anomali & Denetim Öncesi Kontrol — A2 üretim betiği (manda v6).

DOMAIN: mizan satırı (hesap kodu/ad, borç, alacak, bakiye, dönem) ↔ beklenen
profil (beklenen bakiye/oran, sektör eşiği). MOTOR: borç=alacak denge,
sapma %, anomali skoru, karar TEMİZ / İNCELE / KRİTİK / VERİ YOK.
A2 eşleştirme: mizan (VERI_A) ↔ beklenen profil (VERI_B) hesap kodu anahtarı.
Kargo desi / e-ticaret kârlılık kopyası değildir.
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
from ortak.mutabakat_motoru import (  # noqa: E402
    dort_kova_butunluk,
    kok_neden_onerisi,
    normalize_formul,
)

URUN_AD = "Mizan Anomali & Denetim Öncesi Kontrol"
SURUM = "1.0.0"
RENK = "1B3A4B"
KAPASITE = 500
DEMO = 35
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

HESAP_SINIFLARI = ["DonenVarlik", "DuranVarlik", "KisaBorc", "UzunBorc", "OzKaynak"]
DONEMLER = ["2026-Q1", "2026-Q2", "2026-Q3", "2026-Q4", "2026-07"]
YORUM_AB = ["A", "B"]
DURUMLAR = ["Aktif", "Kapali", "Incelemede"]


def _ornek_satirlar():
    satirlar = []
    for i in range(1, DEMO + 1):
        kod = f"HK{1000 + i}"
        if i in (3, 7, 11):
            kod_a, kod_b = kod[:6], kod
        elif i == 15:
            kod_a, kod_b = kod + "X", kod
        else:
            kod_a, kod_b = kod, kod
        borc = round(5000 + (i % 9) * 1200 + i * 85, 2)
        alacak = round(4800 + (i % 7) * 1100 + i * 70, 2)
        if i % 6 == 0:
            alacak = borc
        bakiye = round(borc - alacak, 2)
        beklenen = round(abs(bakiye) * (0.92 + (i % 5) * 0.03), 2)
        if i % 8 == 0:
            beklenen = round(abs(bakiye) + 120, 2)
        elif i % 11 == 0:
            beklenen = round(abs(bakiye) + 0.03, 2)
        elif i == 15:
            beklenen = abs(bakiye)
        oran = round(0.05 + (i % 6) * 0.01, 4)
        esik = round(0.08 + (i % 4) * 0.02, 4)
        sinif = HESAP_SINIFLARI[i % len(HESAP_SINIFLARI)]
        donem = DONEMLER[i % len(DONEMLER)]
        tar = RAPOR_TARIHI - timedelta(days=(DEMO - i))
        satirlar.append({
            "kod_a": kod_a, "kod_b": kod_b, "ad": f"Hesap {i}",
            "borc": borc, "alacak": alacak, "bakiye": bakiye,
            "beklenen": beklenen, "oran": oran, "esik": esik,
            "sinif": sinif, "donem": donem, "tarih": tar,
            "yorum": "B" if i % 5 == 0 else "A",
        })
    return satirlar


ORNEK = _ornek_satirlar()


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        from openpyxl.comments import Comment
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Mizan anomali kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Mizan satırlarını beklenen profil ile eşleştirir; borç=alacak dengesi, "
      "sapma % ve anomali skoru üretir; TEMİZ / İNCELE / KRİTİK kararını verir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Mizan satırı↔beklenen profil eşleştirmesi",
        "Borç=alacak denge, sapma % ve anomali skoru motoru",
        "MAD-001..005 kök neden + TANIMSIZ kova",
        "Denetim öncesi TEMİZ / İNCELE / KRİTİK kararı",
        "Yöneticiye sunulabilir kanıt raporu",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "SMMM ve mali müşavirlik ofisleri",
        "İç denetim ve mali işler ekipleri",
        "Dönem sonu / denetim öncesi hazırlık yapanlar",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) VERI_A'ya mizan satırlarını girin. 2) VERI_B'ye beklenen profili "
      "yapıştırın. 3) MOTOR ve PANO'dan anomali ile kararı izleyin. "
      "4) AKSIYON'dan aksiyon metnini kopyalayın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def veri_a(ws):
    sayfa_hazirla(ws, "VERI_A", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Mizan — hesap kodu/ad, borç, alacak, dönem (sarıya girin)", son_kolon=12)
    sutunlar = ["HesapKodu", "HesapAdi", "Borc", "Alacak", "Bakiye",
                "Donem", "HesapSinifi", "YorumAB", "HedefBakiye", "KayitDolu"]
    baslik_satiri(ws, HDR, sutunlar)
    form_a = {
        "Bakiye": (
            'IF(OR(tblVeriA[[#This Row],[HesapKodu]]="",tblVeriA[[#This Row],[HesapKodu]]=0),"",'
            'IFERROR(ROUND(N(tblVeriA[[#This Row],[Borc]])-N(tblVeriA[[#This Row],[Alacak]]),2),0))'
        ),
        "HedefBakiye": (
            'IF(OR(tblVeriA[[#This Row],[HesapKodu]]="",tblVeriA[[#This Row],[HesapKodu]]=0),"",'
            'IFERROR(SUMIF(tblVeriB[HesapKodu],tblVeriA[[#This Row],[HesapKodu]],'
            'tblVeriB[BeklenenBakiye]),0))'
        ),
        "KayitDolu": 'IF(OR(tblVeriA[[#This Row],[HesapKodu]]="",tblVeriA[[#This Row],[HesapKodu]]=0),"",1)',
    }
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["kod_a"], baslik="Hesap Kodu", mesaj="Hesap planı kodu")
        _sari(ws, r, 2, s["ad"], baslik="Hesap Adı", mesaj="Hesap açıklaması")
        _sari(ws, r, 3, s["borc"], sayi=TL, baslik="Borç", mesaj="Dönem borç tutarı TL")
        _sari(ws, r, 4, s["alacak"], sayi=TL, baslik="Alacak", mesaj="Dönem alacak tutarı TL")
        _sari(ws, r, 6, s["donem"], baslik="Dönem", mesaj="YYYY-QA veya YYYY-AA")
        _sari(ws, r, 7, s["sinif"], baslik="Hesap Sınıfı", mesaj="Liste seçimi")
        _sari(ws, r, 8, s["yorum"], baslik="Yorum A/B", mesaj="Paket varyantı A veya B")
    for r in range(ILK + DEMO, SON + 1):
        for c in (1, 2, 3, 4, 6, 7, 8):
            _sari(ws, r, c, None,
                  sayi=(TL if c in (3, 4) else None),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriA", f"A{HDR}:{get_column_letter(10)}{SON}", sutunlar, formuller=form_a)
    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}", baslik="Hesap Kodu",
              mesaj="Hesap kodu girin", hata_baslik="Eksik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}", baslik="Hesap Adı",
              mesaj="Hesap adı", hata_baslik="Ad", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="80")
    for c, ad in [(3, "Borç"), (4, "Alacak")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj="0 veya üzeri", hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeHesapSiniflari", f"G{ILK}:G{SON}",
              baslik="Hesap Sınıfı", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçim yapın")
    dogrulama(ws, "list", "ListeYorumAB", f"H{ILK}:H{SON}",
              baslik="Yorum A/B", mesaj="A veya B seçin",
              hata_baslik="Liste", hata_mesaj="A veya B seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 11)})


def veri_b(ws):
    sayfa_hazirla(ws, "VERI_B", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Beklenen profil — bakiye/oran/sektör eşiği (sarıya girin)", son_kolon=10)
    sutunlar = ["HesapKodu", "BeklenenBakiye", "BeklenenOran", "SektorEsik",
                "HesapSinifi", "Donem", "GercekSapma", "Kaynak", "Not", "KayitDolu"]
    baslik_satiri(ws, HDR, sutunlar)
    form_b = {
        "GercekSapma": (
            'IF(OR(tblVeriB[[#This Row],[HesapKodu]]="",tblVeriB[[#This Row],[HesapKodu]]=0),"",'
            'IFERROR(ABS(IFERROR(SUMIF(tblVeriA[HesapKodu],tblVeriB[[#This Row],[HesapKodu]],'
            'tblVeriA[Borc]),0)-IFERROR(SUMIF(tblVeriA[HesapKodu],tblVeriB[[#This Row],[HesapKodu]],'
            'tblVeriA[Alacak]),0)-IFERROR(N(tblVeriB[[#This Row],[BeklenenBakiye]]),0)),0))'
        ),
        "KayitDolu": 'IF(OR(tblVeriB[[#This Row],[HesapKodu]]="",tblVeriB[[#This Row],[HesapKodu]]=0),"",1)',
    }
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["kod_b"], baslik="Hesap Kodu", mesaj="Profil hesap kodu")
        _sari(ws, r, 2, s["beklenen"], sayi=TL, baslik="Beklenen Bakiye", mesaj="TL beklenen")
        _sari(ws, r, 3, s["oran"], sayi=YÜZDE, baslik="Beklenen Oran", mesaj="0-1 oran")
        _sari(ws, r, 4, s["esik"], sayi=YÜZDE, baslik="Sektör Eşiği", mesaj="Sapma eşiği")
        _sari(ws, r, 5, s["sinif"], baslik="Hesap Sınıfı", mesaj="Liste seçimi")
        _sari(ws, r, 6, s["donem"], baslik="Dönem", mesaj="YYYY-QA")
        _sari(ws, r, 8, "Ic politika", baslik="Kaynak", mesaj="Profil kaynağı")
        _sari(ws, r, 9, "profil", baslik="Not", mesaj="Kısa not")
    for r in range(ILK + DEMO, SON + 1):
        for c in (1, 2, 3, 4, 5, 6, 8, 9):
            _sari(ws, r, c, None,
                  sayi=(YÜZDE if c in (3, 4) else (TL if c == 2 else None)),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriB", f"A{HDR}:J{SON}", sutunlar, formuller=form_b)
    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}", baslik="Hesap Kodu",
              mesaj="Hesap kodu girin", hata_baslik="Eksik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "decimal", "0", f"B{ILK}:B{SON}", baslik="Beklenen",
              mesaj="0 veya üzeri", hata_baslik="Bakiye", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual")
    for c, ad in [(3, "Oran"), (4, "Eşik")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj="0 veya üzeri", hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeHesapSiniflari", f"E{ILK}:E{SON}",
              baslik="Hesap Sınıfı", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "textLength", "0", f"F{ILK}:F{SON}", baslik="Dönem",
              mesaj="YYYY-QA", hata_baslik="Dönem", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="20")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 11)})
    ws.column_dimensions["H"].width = 24
    ws.column_dimensions["I"].width = 16


def anahtar(ws):
    sayfa_hazirla(ws, "ANAHTAR", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Normalize hesap kodu + uzunluk + ikincil + duplike (Ç7 / E01 / E05)", son_kolon=12)
    h(ws, 4, 1, "A Tarafı (Mizan)", kalin=True, yazi=KOYU_LACIVERT)
    sut_a = ["Ham", "Norm", "NormUzunluk", "IkincilAnahtar", "KaynakTaraf", "Duplike", "SatirNo"]
    baslik_satiri(ws, HDR, sut_a)
    form_a = {
        "Norm": normalize_formul("A6").replace("A6", "tblAnahtarA[[#This Row],[Ham]]").lstrip("="),
        "NormUzunluk": 'IF(tblAnahtarA[[#This Row],[Norm]]="",0,LEN(tblAnahtarA[[#This Row],[Norm]]))',
        "IkincilAnahtar": 'IF(tblAnahtarA[[#This Row],[Norm]]="",""'
                         ',LEFT(tblAnahtarA[[#This Row],[Norm]],mad_ikincilUzunluk))',
        "KaynakTaraf": '"A"',
        "Duplike": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",'
                   'IF(COUNTIF(tblAnahtarA[Norm],tblAnahtarA[[#This Row],[Norm]])>1,"EVET","HAYIR"))',
        "SatirNo": 'IF(tblAnahtarA[[#This Row],[Ham]]="","",ROW()-ROW($A$5))',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 1, f"=IF(VERI_A!A{r}=\"\",\"\",VERI_A!A{r})")
    tablo_ekle(ws, "tblAnahtarA", f"A{HDR}:G{SON}", sut_a, formuller=form_a)

    h(ws, 4, 9, "B Tarafı (Beklenen Profil)", kalin=True, yazi=KOYU_LACIVERT)
    sut_b = ["HamB", "NormB", "NormUzunlukB", "IkincilB", "KaynakB", "DuplikeB", "SatirNoB"]
    baslik_satiri(ws, HDR, sut_b, basla=9)
    form_b = {
        "NormB": 'UPPER(TRIM(SUBSTITUTE(tblAnahtarB[[#This Row],[HamB]]," ","")))',
        "NormUzunlukB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",0,LEN(tblAnahtarB[[#This Row],[NormB]]))',
        "IkincilB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",""'
                    ',LEFT(tblAnahtarB[[#This Row],[NormB]],mad_ikincilUzunluk))',
        "KaynakB": '"B"',
        "DuplikeB": 'IF(tblAnahtarB[[#This Row],[NormB]]="","",'
                    'IF(COUNTIF(tblAnahtarB[NormB],tblAnahtarB[[#This Row],[NormB]])>1,"EVET","HAYIR"))',
        "SatirNoB": 'IF(tblAnahtarB[[#This Row],[HamB]]="","",ROW()-ROW($I$5))',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 9, f"=IF(VERI_B!A{r}=\"\",\"\",VERI_B!A{r})")
    tablo_ekle(ws, "tblAnahtarB", f"I{HDR}:O{SON}", sut_b, formuller=form_b)
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 16)})


def motor(ws):
    sayfa_hazirla(ws, "MOTOR", RENK, URUN_AD, son_kolon=20)
    alt_bant(ws, 3, "Mizan motoru — denge, sapma %, anomali skoru, karar (D13)", son_kolon=16)
    sutunlar = [
        "A_Norm", "Borc", "Alacak", "Bakiye", "BeklenenBakiye", "DengeFarki",
        "SapmaPct", "AnomaliSkoru", "SektorEsik", "BeklenenOran", "AbsSapma",
        "KararSatir", "BorcAlacakOran", "EsikAsimi", "MotorAdim1", "MotorAdim2",
        "HesapSinifi", "Donem", "YorumAB", "MotorAdim3",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "A_Norm": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",tblAnahtarA[[#This Row],[Norm]])',
        "Borc": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblVeriA[Borc]),0))',
        "Alacak": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                  'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblVeriA[Alacak]),0))',
        "Bakiye": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                  'ROUND(tblMotor[[#This Row],[Borc]]-tblMotor[[#This Row],[Alacak]],2))',
        "BeklenenBakiye": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                          'IFERROR(SUMIF(tblAnahtarB[NormB],tblMotor[[#This Row],[A_Norm]],'
                          'tblVeriB[BeklenenBakiye]),0))',
        "DengeFarki": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                      'ROUND(tblMotor[[#This Row],[Borc]]-tblMotor[[#This Row],[Alacak]]'
                      '-tblMotor[[#This Row],[Bakiye]],2))',
        "SapmaPct": (
            'IF(OR(tblMotor[[#This Row],[A_Norm]]="",tblMotor[[#This Row],[BeklenenBakiye]]=0),0,'
            'ABS(ABS(tblMotor[[#This Row],[Bakiye]])-tblMotor[[#This Row],[BeklenenBakiye]])'
            '/ABS(tblMotor[[#This Row],[BeklenenBakiye]]))'
        ),
        "AnomaliSkoru": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'ROUND(tblMotor[[#This Row],[SapmaPct]]*100'
            '+IF(ABS(tblMotor[[#This Row],[DengeFarki]])>toleransKurus,mad_dengeCeza,0)'
            '+IF(tblMotor[[#This Row],[SapmaPct]]>tblMotor[[#This Row],[SektorEsik]],mad_esikCeza,0),1))'
        ),
        "SektorEsik": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                      'IFERROR(SUMIF(tblAnahtarB[NormB],tblMotor[[#This Row],[A_Norm]],'
                      'tblVeriB[SektorEsik]),mad_esikSapma))',
        "BeklenenOran": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                        'IFERROR(SUMIF(tblAnahtarB[NormB],tblMotor[[#This Row],[A_Norm]],'
                        'tblVeriB[BeklenenOran]),0))',
        "AbsSapma": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                    'ROUND(ABS(ABS(tblMotor[[#This Row],[Bakiye]])'
                    '-tblMotor[[#This Row],[BeklenenBakiye]]),2))',
        "KararSatir": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'IF(OR(tblMotor[[#This Row],[AnomaliSkoru]]>=mad_esikKritik,'
            'tblMotor[[#This Row],[SapmaPct]]>=mad_esikKritikOran),"KRİTİK",'
            'IF(OR(tblMotor[[#This Row],[AnomaliSkoru]]>=mad_esikIncele,'
            'tblMotor[[#This Row],[SapmaPct]]>=tblMotor[[#This Row],[SektorEsik]]),"İNCELE","TEMİZ")))'
        ),
        "BorcAlacakOran": (
            'IF(OR(tblMotor[[#This Row],[A_Norm]]="",tblMotor[[#This Row],[Alacak]]=0),0,'
            'tblMotor[[#This Row],[Borc]]/tblMotor[[#This Row],[Alacak]])'
        ),
        "EsikAsimi": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                     'IF(tblMotor[[#This Row],[SapmaPct]]>tblMotor[[#This Row],[SektorEsik]],1,0))',
        "MotorAdim1": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                      'tblMotor[[#This Row],[AbsSapma]])',
        "MotorAdim2": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                      'tblMotor[[#This Row],[AnomaliSkoru]])',
        "HesapSinifi": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                       'IFERROR(INDEX(tblVeriA[HesapSinifi],MATCH(tblMotor[[#This Row],[A_Norm]],'
                       'tblAnahtarA[Norm],0)),""))',
        "Donem": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                 'IFERROR(INDEX(tblVeriA[Donem],MATCH(tblMotor[[#This Row],[A_Norm]],'
                 'tblAnahtarA[Norm],0)),""))',
        "YorumAB": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                   'IFERROR(INDEX(tblVeriA[YorumAB],MATCH(tblMotor[[#This Row],[A_Norm]],'
                   'tblAnahtarA[Norm],0)),""))',
        "MotorAdim3": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                      'tblMotor[[#This Row],[EsikAsimi]]+tblMotor[[#This Row],[AnomaliSkoru]]/100)',
    }
    tablo_ekle(ws, "tblMotor", f"A{HDR}:T{SON}", sutunlar, formuller=form)
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 21)})


def eslestirme(ws):
    sayfa_hazirla(ws, "ESLESTIRME", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Mizan bakiye ↔ beklenen bakiye dört kova — her A kaydı tek kovaya (E02)", son_kolon=14)
    sutunlar = [
        "A_Norm", "B_Sayim", "B_Tutar", "A_Tutar", "Fark", "AbsFark",
        "Kova", "IkincilEslesme", "DuplikeBayrak", "FarkOran",
        "YuvarlakA", "YuvarlakB", "ToleransTest", "EsikUstu",
        "AnomaliSkoru", "MotorAdim",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    formuller = {
        "A_Norm": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",tblAnahtarA[[#This Row],[Norm]])',
        "B_Sayim": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                   'COUNTIF(tblAnahtarB[NormB],tblEslesme[[#This Row],[A_Norm]]))',
        "B_Tutar": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                   'IFERROR(SUMIF(tblAnahtarB[NormB],tblEslesme[[#This Row],[A_Norm]],'
                   'tblVeriB[BeklenenBakiye]),0))',
        "A_Tutar": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                   'IFERROR(ABS(SUMIF(tblAnahtarA[Norm],tblEslesme[[#This Row],[A_Norm]],'
                   'tblVeriA[Bakiye])),0))',
        "Fark": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                'ROUND(tblEslesme[[#This Row],[A_Tutar]],2)-ROUND(tblEslesme[[#This Row],[B_Tutar]],2))',
        "AbsFark": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",ABS(tblEslesme[[#This Row],[Fark]]))',
        "Kova": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            'IF(tblEslesme[[#This Row],[B_Sayim]]=0,"eslesmedi",'
            'IF(ROUND(tblEslesme[[#This Row],[A_Tutar]],2)=ROUND(tblEslesme[[#This Row],[B_Tutar]],2),"eslesti",'
            'IF(ABS(ROUND(tblEslesme[[#This Row],[A_Tutar]],2)-ROUND(tblEslesme[[#This Row],[B_Tutar]],2))'
            '<=toleransKurus,"tolerans_icinde","tutar_farki"))))'
        ),
        "IkincilEslesme": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            'IF(tblEslesme[[#This Row],[B_Sayim]]>0,"BIRINCIL",'
            'IF(COUNTIF(tblAnahtarB[IkincilB],tblAnahtarA[[#This Row],[IkincilAnahtar]])>0,'
            '"IKINCIL","YOK")))'
        ),
        "DuplikeBayrak": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",tblAnahtarA[[#This Row],[Duplike]])',
        "FarkOran": 'IF(OR(tblEslesme[[#This Row],[A_Norm]]="",tblEslesme[[#This Row],[A_Tutar]]=0),0,'
                    'tblEslesme[[#This Row],[AbsFark]]/ABS(tblEslesme[[#This Row],[A_Tutar]]))',
        "YuvarlakA": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",ROUND(tblEslesme[[#This Row],[A_Tutar]],2))',
        "YuvarlakB": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",ROUND(tblEslesme[[#This Row],[B_Tutar]],2))',
        "ToleransTest": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                        'IF(tblEslesme[[#This Row],[AbsFark]]<=toleransKurus,1,0))',
        "EsikUstu": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                    'IF(tblEslesme[[#This Row],[AbsFark]]>mad_esikFark,1,0))',
        "AnomaliSkoru": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                        'IFERROR(SUMIF(tblMotor[A_Norm],tblEslesme[[#This Row],[A_Norm]],'
                        'tblMotor[AnomaliSkoru]),0))',
        "MotorAdim": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                     'tblEslesme[[#This Row],[B_Sayim]]+tblEslesme[[#This Row],[ToleransTest]]'
                     '+tblEslesme[[#This Row],[EsikUstu]])',
    }
    tablo_ekle(ws, "tblEslesme", f"A{HDR}:P{SON}", sutunlar, formuller=formuller)
    h(ws, 2, 1, "Kova Bütünlük Denetimi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 2, 3,
      dort_kova_butunluk(
          'COUNTA(tblEslesme[A_Norm])',
          'COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde")'
          '+COUNTIF(tblEslesme[Kova],"tutar_farki")+COUNTIF(tblEslesme[Kova],"eslesmedi")',
      ))
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 13 for i in range(1, 17)})


def kok_neden(ws):
    sayfa_hazirla(ws, "KOK_NEDEN", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "MAD-001..005 kök neden — eşleşmeyen fark TANIMSIZ (E04)", son_kolon=10)
    h(ws, 4, 1, "Kural Önceliği", kalin=True, yazi=KOYU_LACIVERT)
    kurallar = kok_neden_onerisi([
        {"neden_kodu": "MAD-001", "test": "ABS(denge_farki)>toleransKurus",
         "aciklama": "Borç-alacak-bakiye denge bozukluğu"},
        {"neden_kodu": "MAD-002", "test": "sapma_pct>sektor_esik",
         "aciklama": "Beklenen bakiyeye göre sapma eşiği aşımı"},
        {"neden_kodu": "MAD-003", "test": "oran_sapmasi",
         "aciklama": "Beklenen oran sapması"},
        {"neden_kodu": "MAD-004", "test": "anomali_skoru>=esikKritik",
         "aciklama": "Anomali skoru kritik eşiğin üzerinde"},
        {"neden_kodu": "MAD-005", "test": "ABS(fark)>esikFark",
         "aciklama": "Mizan-beklenen mutlak tutar farkı"},
    ])
    baslik_satiri(ws, HDR, ["Sira", "NedenKodu", "Test", "Aciklama"])
    for i, k in enumerate(kurallar):
        r = ILK + i
        h(ws, r, 1, k["sira"], sayi=CATI)
        h(ws, r, 2, k["neden_kodu"], kalin=True,
          yazi=("B3261E" if k["neden_kodu"] == "TANIMSIZ" else KOYU_LACIVERT))
        h(ws, r, 3, k["test_formul_sablon"], kaydir=True)
        h(ws, r, 4, k["aciklama"], kaydir=True)
    h(ws, ILK + len(kurallar) + 1, 1,
      "TANIMSIZ kovası görünür tutulur — zorla sınıflandırma yok. Elle inceleyin.",
      yazi="B3261E", kaydir=True)
    ws.merge_cells(start_row=ILK + len(kurallar) + 1, start_column=1,
                   end_row=ILK + len(kurallar) + 1, end_column=6)

    h(ws, 14, 1, "Satır Sınıflandırması", kalin=True, yazi=KOYU_LACIVERT)
    sut = ["A_Norm", "Kova", "AbsFark", "SapmaPct", "AnomaliSkoru", "DengeFarki", "AbsSapma", "KokNeden"]
    baslik_satiri(ws, 15, sut)
    form = {
        "A_Norm": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",tblEslesme[[#This Row],[A_Norm]])',
        "Kova": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                'IFERROR(INDEX(tblEslesme[Kova],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),""))',
        "AbsFark": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                   'IFERROR(INDEX(tblEslesme[AbsFark],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),0))',
        "SapmaPct": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[SapmaPct]),0))',
        "AnomaliSkoru": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                        'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[AnomaliSkoru]),0))',
        "DengeFarki": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                      'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[DengeFarki]),0))',
        "AbsSapma": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[AbsSapma]),0))',
        "KokNeden": (
            'IF(OR(tblKok[[#This Row],[A_Norm]]="",tblKok[[#This Row],[Kova]]<>"tutar_farki"),"",'
            'IF(ABS(tblKok[[#This Row],[DengeFarki]])>toleransKurus,"MAD-001",'
            'IF(tblKok[[#This Row],[SapmaPct]]>mad_esikSapma,"MAD-002",'
            'IF(tblKok[[#This Row],[AnomaliSkoru]]>=mad_esikKritik,"MAD-004",'
            'IF(tblKok[[#This Row],[AbsFark]]>mad_esikFark,"MAD-005","TANIMSIZ")))))'
        ),
    }
    for i in range(KAPASITE):
        r = 16 + i
        h(ws, r, 1, f"=IF(ESLESTIRME!A{ILK + i}=\"\",\"\",ESLESTIRME!A{ILK + i})")
    tablo_ekle(ws, "tblKok", f"A15:H{15 + KAPASITE}", sut, formuller=form)
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 9)})
    ws.column_dimensions["D"].width = 36


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=20)
    sabitle(ws, "A3")
    h(ws, 3, 1, "Mizan Anomali Karar Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 3, 3, "Rapor Tarihi", yazi=GRİ)
    h(ws, 3, 4, "=raporTarihi", sayi=TARİH)

    h(ws, 4, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2,
      '=IF(COUNTA(tblVeriA[HesapKodu])=0,"VERİ YOK",'
      'IF(COUNTIF(tblMotor[KararSatir],"KRİTİK")>0,"KRİTİK",'
      'IF(COUNTIF(tblMotor[KararSatir],"İNCELE")>COUNTIF(tblMotor[KararSatir],"TEMİZ"),"İNCELE","TEMİZ")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem mizan kararı | Neden önemli: Boş veride VERİ YOK | "
               "Doğru kullanım: Otomatik | Örnek: TEMİZ | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Eşleşme Oranı", '=IFERROR(mad_eslesmeOrani,0)', YÜZDE),
        (7, "Toplam Borç", '=IFERROR(SUM(tblVeriA[Borc]),0)', TL),
        (8, "Toplam Alacak", '=IFERROR(SUM(tblVeriA[Alacak]),0)', TL),
        (9, "Mizan-Profil Fark", '=IFERROR(mad_mutabakatFarki,0)', TL),
        (10, "Ort. Sapma %", '=IFERROR(mad_modulT1,0)', YÜZDE),
        (11, "Ort. Anomali", '=IFERROR(mad_modulT2,0)', CATI),
        (12, "Toplam Abs Sapma", '=IFERROR(SUM(tblMotor[AbsSapma]),0)', TL),
        (13, "Eşleşti Adet", '=COUNTIF(tblEslesme[Kova],"eslesti")', CATI),
        (14, "Tutar Farkı Adet", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (15, "Eşleşmedi Adet", '=COUNTIF(tblEslesme[Kova],"eslesmedi")', CATI),
        (16, "Kova Bütünlük", "=mad_kovaButunluk", None),
        (17, "Anomali Sayısı", "=mad_modulO1", CATI),
        (18, "Tornado Zirve", "=mad_modulO2", TL),
        (19, "Senaryo Toplam", "=mad_modulO3", TL),
        (20, "Yoğunlaşma", "=mad_modulO6", YÜZDE),
        (21, "Kalite Skoru", "=mad_modulO8", CATI),
        (22, "Tahmin Aralık", "=mad_modulI1", YÜZDE),
        (23, "Sapma P90", "=mad_modulI2", YÜZDE),
        (24, "Duplike Uyarı", '=COUNTIF(tblAnahtarA[Duplike],"EVET")', CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (7, '=mad_modulT1Yorum'),
        (8, '=mad_modulT2Yorum'),
        (9, '=mad_modulO1Yorum'),
        (10, '=mad_modulO2Yorum'),
        (11, '=mad_modulO3Yorum'),
        (12, '=mad_modulO6Yorum'),
        (13, '=mad_modulO8Yorum'),
        (14, '=mad_modulMSenYorum'),
        (15, '=mad_modulI1Yorum'),
        (16, '=mad_modulI2Yorum'),
    ]
    for r, form in yorumlar:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 26, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1, "Kova")
    h(ws, 27, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Eşleşti", 20), ("Tolerans", 5), ("Tutar farkı", 7), ("Eşleşmedi", 3),
    ], 28):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 27, 4, "Neden")
    h(ws, 27, 5, "Adet")
    for i, (ad, adet) in enumerate([
        ("MAD-001", 2), ("MAD-002", 2), ("MAD-003", 1), ("MAD-004", 1),
        ("MAD-005", 1), ("Tanımsız", 2),
    ], 28):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 27, 7, "Gün")
    h(ws, 27, 8, "Sapma")
    for i in range(7):
        h(ws, 28 + i, 7, i + 1, sayi=CATI)
        h(ws, 28 + i, 8, round(40 + i * 8.5, 2), sayi=TL)

    h(ws, 27, 10, "Senaryo")
    h(ws, 27, 11, "Etki")
    for i, (ad, v) in enumerate([
        ("Eşik +2pp", 180), ("Sapma -10%", 220), ("Kritik +5", 95),
        ("Profil +1", 140), ("Denge sıkı", 160),
    ], 28):
        h(ws, i, 10, ad)
        h(ws, i, 11, v, sayi=TL)

    pie1 = PieChart()
    pie1.title = "Kova Dağılımı"
    pie1.add_data(Reference(ws, min_col=2, min_row=27, max_row=31), titles_from_data=True)
    pie1.set_categories(Reference(ws, min_col=1, min_row=28, max_row=31))
    pie1.width, pie1.height = 10, 7
    ws.add_chart(pie1, "A36")

    bar1 = BarChart()
    bar1.type = "col"
    bar1.title = "Kök Neden"
    bar1.add_data(Reference(ws, min_col=5, min_row=27, max_row=33), titles_from_data=True)
    bar1.set_categories(Reference(ws, min_col=4, min_row=28, max_row=33))
    bar1.width, bar1.height = 10, 7
    ws.add_chart(bar1, "F36")

    line1 = LineChart()
    line1.title = "Günlük Sapma"
    line1.add_data(Reference(ws, min_col=8, min_row=27, max_row=34), titles_from_data=True)
    line1.set_categories(Reference(ws, min_col=7, min_row=28, max_row=34))
    line1.width, line1.height = 10, 7
    ws.add_chart(line1, "A51")

    bar2 = BarChart()
    bar2.type = "bar"
    bar2.title = "Tornado Etki"
    bar2.add_data(Reference(ws, min_col=11, min_row=27, max_row=32), titles_from_data=True)
    bar2.set_categories(Reference(ws, min_col=10, min_row=28, max_row=32))
    bar2.width, bar2.height = 10, 7
    ws.add_chart(bar2, "F51")

    pie2 = PieChart()
    pie2.title = "Eşleşme Payı"
    pie2.add_data(Reference(ws, min_col=2, min_row=27, max_row=29), titles_from_data=True)
    pie2.set_categories(Reference(ws, min_col=1, min_row=28, max_row=29))
    pie2.width, pie2.height = 9, 6
    ws.add_chart(pie2, "A66")

    bar3 = BarChart()
    bar3.title = "Üst Nedenler"
    bar3.add_data(Reference(ws, min_col=5, min_row=27, max_row=30), titles_from_data=True)
    bar3.set_categories(Reference(ws, min_col=4, min_row=28, max_row=30))
    bar3.width, bar3.height = 9, 6
    ws.add_chart(bar3, "F66")

    line2 = LineChart()
    line2.title = "Sapma Eğilimi"
    line2.add_data(Reference(ws, min_col=8, min_row=27, max_row=31), titles_from_data=True)
    line2.set_categories(Reference(ws, min_col=7, min_row=28, max_row=31))
    line2.width, line2.height = 9, 6
    ws.add_chart(line2, "A79")

    bar4 = BarChart()
    bar4.title = "Senaryo Karşılaştırma"
    bar4.add_data(Reference(ws, min_col=11, min_row=27, max_row=30), titles_from_data=True)
    bar4.set_categories(Reference(ws, min_col=10, min_row=28, max_row=30))
    bar4.width, bar4.height = 9, 6
    ws.add_chart(bar4, "F79")

    baski_hazirla(ws, "A1:L35", f"{URUN_AD} | {SURUM}")
    genislik(ws, {get_column_letter(i): 16 for i in range(1, 13)})
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["D"].width = 18


def aksiyon(ws):
    sayfa_hazirla(ws, "AKSIYON", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Mizan aksiyon metni — tutar farkı satırları için kopyalanabilir", son_kolon=10)
    sut = ["A_Norm", "AbsSapma", "AnomaliSkoru", "Karar", "Oncelik", "AksiyonMetni"]
    baslik_satiri(ws, HDR, sut)
    form = {
        "A_Norm": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                  'IF(tblEslesme[[#This Row],[Kova]]="tutar_farki",tblEslesme[[#This Row],[A_Norm]],""))',
        "AbsSapma": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblMotor[A_Norm],tblAksiyon[[#This Row],[A_Norm]],tblMotor[AbsSapma]),0))',
        "AnomaliSkoru": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                        'IFERROR(SUMIF(tblMotor[A_Norm],tblAksiyon[[#This Row],[A_Norm]],'
                        'tblMotor[AnomaliSkoru]),0))',
        "Karar": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                 'IFERROR(INDEX(tblMotor[KararSatir],MATCH(tblAksiyon[[#This Row],[A_Norm]],tblMotor[A_Norm],0)),""))',
        "Oncelik": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                   'IF(OR(tblAksiyon[[#This Row],[Karar]]="KRİTİK",'
                   'tblAksiyon[[#This Row],[Karar]]="İNCELE"),"YUKSEK","NORMAL"))',
        "AksiyonMetni": (
            'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
            '_xlfn.TEXTJOIN(" ",TRUE,"Hesap",tblAksiyon[[#This Row],[A_Norm]],'
            '"karar",tblAksiyon[[#This Row],[Karar]],'
            '"sapma",TEXT(tblAksiyon[[#This Row],[AbsSapma]],"0.00"),'
            '"skor",TEXT(tblAksiyon[[#This Row],[AnomaliSkoru]],"0.0")))'
        ),
    }
    tablo_ekle(ws, "tblAksiyon", f"A{HDR}:F{SON}", sut, formuller=form)
    genislik(ws, {"A": 16, "B": 14, "C": 12, "D": 12, "E": 10, "F": 55})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Mizan Anomali Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Eşleşme Oranı", "=mad_eslesmeOrani", YÜZDE),
        (10, "Toplam Borç", "=SUM(tblVeriA[Borc])", TL),
        (11, "Mizan-Profil Fark", "=mad_mutabakatFarki", TL),
        (12, "Ort. Sapma %", "=mad_modulT1", YÜZDE),
        (13, "Ort. Anomali",
         '=IFERROR(IF(COUNTA(tblMotor[A_Norm])=0,0,AVERAGE(tblMotor[AnomaliSkoru])),0)', CATI),
        (14, "Tutar Farkı", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (15, "Kova Bütünlük", "=mad_kovaButunluk", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 17, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 1,
      "Bakiye = borç − alacak. Tolerans kuruşu AYARLAR'dan gelir. "
      "Sektör eşikleri kullanıcı girdisidir. Bu çıktı karar destektir; "
      "kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A18:H18")
    h(ws, 20, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 1, "1. KRİTİK satırlarda belge ve fiş zincirini doğrulayın.")
    h(ws, 22, 1, "2. İNCELE satırlarında beklenen profili güncelleyin veya düzeltin.")
    h(ws, 23, 1, "3. TANIMSIZ kök nedenleri elle inceleyin.")
    baski_hazirla(ws, "A1:H24", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 28, "B": 18, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "A satır sayısı", '=COUNTA(tblVeriA[HesapKodu])'),
        (6, "B satır sayısı", '=COUNTA(tblVeriB[HesapKodu])'),
        (7, "Eşleşti", '=COUNTIF(tblEslesme[Kova],"eslesti")'),
        (8, "Tolerans", '=COUNTIF(tblEslesme[Kova],"tolerans_icinde")'),
        (9, "Tutar farkı", '=COUNTIF(tblEslesme[Kova],"tutar_farki")'),
        (10, "Eşleşmedi", '=COUNTIF(tblEslesme[Kova],"eslesmedi")'),
        (11, "Kova toplam", "=B7+B8+B9+B10"),
        (12, "Bütünlük", '=IF(B11=B5,"BUTUN","KAYIP")'),
        (13, "Duplike A", '=COUNTIF(tblAnahtarA[Duplike],"EVET")'),
        (14, "TANIMSIZ", '=COUNTIF(tblKok[KokNeden],"TANIMSIZ")'),
        (15, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (16, "Kalite skoru", '=IF(B5=0,0,ROUND(100*(B7+B8)/MAX(B5,1),0))'),
        (17, "Borç toplam", "=SUM(tblVeriA[Borc])"),
        (18, "Alacak toplam", "=SUM(tblVeriA[Alacak])"),
        (19, "Denge kontrol", "=B17-B18"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, kalin=True, sayi=TL if r >= 17 else (CATI if r < 15 else None))
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO} satır örnek mizan vardır: mizan↔beklenen farkı, "
      "kısa/uzun hesap kodu (3,7,11), eşleşmeyen (15) ve Yorum B satırları kasıtlıdır. "
      "VERI_A ve VERI_B'yi temizleyip kendi mizan/profilinizi yapıştırabilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Karar seti: TEMİZ / İNCELE / KRİTİK / VERİ YOK")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 50000, sayi=CATI)
    genislik(ws, {"A": 70, "B": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=8)
    h(ws, 3, 1, "Hesap Sınıfı Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "HesapSinifi", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, p in enumerate(HESAP_SINIFLARI, 7):
        _sari(ws, i, 1, p, baslik="Hesap Sınıfı", mesaj="Listeye ekleyebilirsiniz")
    h(ws, 3, 3, "Dönem Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 3, "Donem", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, d in enumerate(DONEMLER, 7):
        _sari(ws, i, 3, d, baslik="Dönem", mesaj="Dönem değeri")
    h(ws, 3, 5, "Yorum A/B", kalin=True)
    h(ws, 6, 5, "Secim", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, y in enumerate(YORUM_AB, 7):
        _sari(ws, i, 5, y, baslik="Yorum", mesaj="A veya B")
    h(ws, 3, 7, "Durum", kalin=True)
    h(ws, 6, 7, "Durum", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, d in enumerate(DURUMLAR, 7):
        _sari(ws, i, 7, d, baslik="Durum", mesaj="Durum değeri")
    genislik(ws, {"A": 18, "C": 16, "E": 12, "G": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Tek kaynak parametreler — kaynak ve yürürlük zorunlu (D10)", son_kolon=10)
    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("toleransKurus", 0.05, "TL", "İş kuralı — kuruş toleransı", "01.01.2026", "Eşleşme toleransı"),
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026", "Rapor tarihi"),
        ("mad_esikEslesme", 0.85, "oran", "İç politika", "01.01.2026", "Eşleşme oranı eşiği"),
        ("mad_esikFark", 100, "TL", "İç politika", "01.01.2026", "Fark tutarı eşiği"),
        ("mad_esikSapma", 0.10, "oran", "Sektör sapma politikası", "01.01.2026", "Sapma % eşiği"),
        ("mad_esikIncele", 25, "skor", "Anomali politikası", "01.01.2026", "İNCELE skor eşiği"),
        ("mad_esikKritik", 55, "skor", "Anomali politikası", "01.01.2026", "KRİTİK skor eşiği"),
        ("mad_esikKritikOran", 0.35, "oran", "Kritik sapma oranı", "01.01.2026", "KRİTİK oran eşiği"),
        ("mad_dengeCeza", 25, "skor", "Anomali skor modeli", "01.01.2026", "Denge bozukluğu cezası"),
        ("mad_esikCeza", 30, "skor", "Anomali skor modeli", "01.01.2026", "Eşik aşımı cezası"),
        ("mad_azamiTutar", 100000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul tutar"),
        ("mad_ikincilUzunluk", 8, "karakter", "Anahtar stratejisi SPEC", "01.01.2026", "Ikincil anahtar"),
        ("mad_olcekHedef", 50000, "satir", "Manda A2 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("mad_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("mad_esikSenaryo", 0.10, "oran", "M_SEN parametre", "01.01.2026", "Eşik +% senaryo"),
        ("mad_fiyatModeli", "Tek seferlik alım — aylık ücret yok", "metin", "Ürün konumlandırma", "01.01.2026", "R3 ayrım"),
        ("mad_kvkkBeyani", "Veri cihazdan çıkmaz", "metin", "KVKK beyanı", "01.01.2026", "R3 ayrım"),
        ("mad_anomaliMotoru", "denge+sapma+skor motoru", "metin", "MOTOR", "01.01.2026", "R1 ayrım"),
        ("mad_normAnahtar", "UPPER TRIM SUBSTITUTE", "metin", "Ç7 anahtar", "01.01.2026", "Serbest alternatif"),
        ("mad_tanimsizKova", "TANIMSIZ görünür", "metin", "E04", "01.01.2026", "Serbest alternatif"),
        ("mad_kovaButunlukAd", "Dört kova bütünlük", "metin", "E02", "01.01.2026", "Serbest alternatif"),
        ("mad_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("mad_firmaUnvan", "Örnek Mali Müşavirlik", "metin", "Kullanıcı girişi", "01.01.2026", "Firma"),
        ("mad_kuralSeti", "MAD-001..005", "metin", "Kök neden kuralları", "01.01.2026", "Kural seti"),
        ("mad_eslesmeOraniAd", "Eşleşme oranı veriden", "metin", "E03", "01.01.2026", "Serbest alternatif"),
    ]
    for i, (ana, deg, bir, kay, yur, acik) in enumerate(params):
        r = 6 + i
        h(ws, r, 1, ana)
        if isinstance(deg, date):
            _sari(ws, r, 2, deg, sayi=TARİH, baslik=ana, mesaj=acik)
        elif isinstance(deg, float) and deg <= 1 and bir == "oran":
            _sari(ws, r, 2, deg, sayi=YÜZDE, baslik=ana, mesaj=acik)
        elif isinstance(deg, (int, float)):
            _sari(ws, r, 2, deg, sayi=TL if bir == "TL" else CATI, baslik=ana, mesaj=acik)
        else:
            _sari(ws, r, 2, deg, baslik=ana, mesaj=acik)
        h(ws, r, 3, bir)
        h(ws, r, 4, kay)
        h(ws, r, 5, yur)
        h(ws, r, 6, acik, kaydir=True)

    h(ws, 31, 8, "Motor Cikti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 31, 9, "Deger", kalin=True, yazi=KOYU_LACIVERT)
    motor_c = [
        (32, "mad_mutabakatFarki",
         '=IFERROR(SUM(tblEslesme[AbsFark]),0)'),
        (33, "mad_eslesmeOrani",
         '=IF(COUNTA(tblEslesme[A_Norm])=0,0,'
         '(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/COUNTA(tblEslesme[A_Norm]))'),
        (34, "mad_kovaButunluk", "=ESLESTIRME!C2"),
        (35, "mad_modulT1",
         '=IFERROR(IF(COUNTA(tblMotor[A_Norm])=0,0,AVERAGE(tblMotor[SapmaPct])),0)'),
        (36, "mad_modulT2",
         '=IFERROR(IF(COUNTA(tblMotor[A_Norm])=0,0,AVERAGE(tblMotor[AnomaliSkoru])),0)'),
        (37, "mad_modulO1",
         '=COUNTIF(tblEslesme[EsikUstu],1)+COUNTIF(tblMotor[KararSatir],"KRİTİK")'
         '+COUNTIF(tblMotor[KararSatir],"İNCELE")'),
        (38, "mad_modulO2",
         '=IFERROR(MAX(PANO!K28:K32),0)'),
        (39, "mad_modulO3",
         '=IFERROR(SUM(tblMotor[AbsSapma])*(1+mad_esikSenaryo),0)'),
        (40, "mad_modulO6",
         '=IFERROR(IF(COUNTA(tblVeriA[HesapSinifi])=0,0,'
         'COUNTIF(tblVeriA[HesapSinifi],INDEX(tblVeriA[HesapSinifi],1))'
         '/COUNTA(tblVeriA[HesapSinifi])),0)'),
        (41, "mad_modulO8",
         '=IF(COUNTA(tblVeriA[HesapKodu])=0,0,'
         'ROUND(100*(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/MAX(COUNTA(tblEslesme[A_Norm]),1),0))'),
        (42, "mad_modulMSen",
         '=IFERROR(SUM(tblMotor[AbsSapma])*mad_esikSenaryo,0)'),
        (43, "mad_modulI1",
         '=IFERROR(PERCENTILE(tblMotor[SapmaPct],mad_yuzdelikOran),0)'),
        (44, "mad_modulI2",
         '=IFERROR(PERCENTILE(tblMotor[SapmaPct],mad_yuzdelikOran),0)'),
    ]
    for r, ad, form in motor_c:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kalin=True)

    h(ws, 46, 8, "Canli Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorum_satir = [
        (47, "mad_modulT1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Ort sapma",TEXT(mad_modulT1,"0.0%")),"-")'),
        (48, "mad_modulT2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Ort anomali",TEXT(mad_modulT2,"0.0")),"-")'),
        (49, "mad_modulO1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Anomali adet",TEXT(mad_modulO1,"0")),"-")'),
        (50, "mad_modulO2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tornado",TEXT(mad_modulO2,"0.00"),"TL"),"-")'),
        (51, "mad_modulO3Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Senaryo toplam",TEXT(mad_modulO3,"0.00"),"TL"),"-")'),
        (52, "mad_modulO6Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma",TEXT(mad_modulO6,"0.0%")),"-")'),
        (53, "mad_modulO8Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Kalite",TEXT(mad_modulO8,"0")),"-")'),
        (54, "mad_modulMSenYorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Eşik senaryo etki",TEXT(mad_modulMSen,"0.00"),"TL"),"-")'),
        (55, "mad_modulI1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tahmin",TEXT(mad_modulI1,"0.0%")),"-")'),
        (56, "mad_modulI2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"P90",TEXT(mad_modulI2,"0.0%")),"-")'),
    ]
    for r, ad, form in yorum_satir:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 28, "B": 42, "C": 12, "D": 28, "E": 14, "F": 28, "H": 24, "I": 42})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    adimlar = [
        "VERI_A: Mizan satırlarını (hesap kodu, ad, borç, alacak, dönem, sınıf) sarıya girin.",
        "VERI_B: Beklenen bakiye, oran ve sektör eşiğini yapıştırın.",
        "ANAHTAR: Normalize hesap kodu, uzunluk, ikincil ve duplike kolonları otomatik dolar.",
        "MOTOR: Denge, sapma %, anomali skoru ve satır kararı hesaplanır.",
        "ESLESTIRME: Mizan bakiye ↔ beklenen bakiye dört kova denetimi.",
        "KOK_NEDEN: MAD-001..005 sınıflandırır; TANIMSIZ elle incelenir.",
        "PANO / KANIT_RAPORU: Kararı izleyin ve PDF yazdırın.",
    ]
    for i, m in enumerate(adimlar, 5):
        h(ws, i, 1, f"{i - 4}. {m}", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    h(ws, 13, 1, "Sık yapılan hatalar", kalin=True, yazi="B3261E")
    for i, m in enumerate([
        "Hesap kodunu boşluklu bırakmak (normalize edilir ama kontrol edin)",
        "Beklenen bakiyeyi 0 yazmak (sapma hesaplanamaz)",
        "Sektör eşiğini yüzde yerine tutar olarak girmek",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 31):
        dogrulama(ws, "textLength", "0", f"D{r}",
                  baslik="Kaynak", mesaj="Kaynak metni girin",
                  hata_baslik="Kaynak", hata_mesaj="Boş bırakmayın",
                  isaret="greaterThanOrEqual", f2="120")
    ws = wb["LISTELER"]
    for r in range(7, 12):
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Hesap Sınıfı", mesaj="Ad girin",
                  hata_baslik="Ad", hata_mesaj="En az 1 karakter",
                  isaret="greaterThan", f2="40")
    for r in range(7, 12):
        dogrulama(ws, "textLength", "1", f"C{r}",
                  baslik="Dönem", mesaj="Dönem girin",
                  hata_baslik="Dönem", hata_mesaj="Geçersiz",
                  isaret="greaterThan", f2="30")
    ws = wb["KANIT_RAPORU"]
    for r, ad in [(5, "Sürüm"), (21, "Aksiyon1"), (22, "Aksiyon2"), (23, "Aksiyon3")]:
        dogrulama(ws, "textLength", "0", f"B{r}" if r == 5 else f"A{r}",
                  baslik=ad, mesaj="Metin alanı",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="200")


def kosullu_bicim(wb):
    yesil = PatternFill("solid", fgColor="C6EFCE")
    kirmizi = PatternFill("solid", fgColor="FFC7CE")
    sari = PatternFill("solid", fgColor="FFEB9C")
    yf = Font(color="006100", name=FONT)
    kf = Font(color="9C0006", name=FONT)

    ws = wb["PANO"]
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"TEMİZ"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"İNCELE"'], fill=sari))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"KRİTİK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(6, 25):
        ws.conditional_formatting.add(f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(f"B{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["ESLESTIRME"]
    for col in ("G",):
        ws.conditional_formatting.add(f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=['"eslesti"'], fill=yesil, font=yf))
        ws.conditional_formatting.add(f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=['"tutar_farki"'], fill=sari))
        ws.conditional_formatting.add(f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=['"eslesmedi"'], fill=kirmizi, font=kf))
        ws.conditional_formatting.add(f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=['"tolerans_icinde"'], fill=sari))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add(f"L{ILK}:L{SON}",
        CellIsRule(operator="equal", formula=['"İNCELE"'], fill=sari))
    ws.conditional_formatting.add(f"L{ILK}:L{SON}",
        CellIsRule(operator="equal", formula=['"KRİTİK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"L{ILK}:L{SON}",
        CellIsRule(operator="equal", formula=['"TEMİZ"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(f"K{ILK}:K{SON}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=sari))

    ws = wb["KOK_NEDEN"]
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"TANIMSIZ"'], fill=kirmizi, font=kf))
    for kod in ("MAD-001", "MAD-002", "MAD-003", "MAD-004", "MAD-005"):
        ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
            CellIsRule(operator="equal", formula=[f'"{kod}"'], fill=sari))

    ws = wb["AKSIYON"]
    ws.conditional_formatting.add(f"E{ILK}:E{SON}",
        CellIsRule(operator="equal", formula=['"YUKSEK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"E{ILK}:E{SON}",
        CellIsRule(operator="equal", formula=['"NORMAL"'], fill=sari))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add("B12", CellIsRule(operator="equal", formula=['"BUTUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("B12", CellIsRule(operator="equal", formula=['"KAYIP"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add("B16", CellIsRule(operator="greaterThanOrEqual", formula=["90"], fill=yesil, font=yf))
    ws.conditional_formatting.add("B16", CellIsRule(operator="lessThan", formula=["70"], fill=kirmizi, font=kf))
    for r in range(5, 16):
        ws.conditional_formatting.add(f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"TEMİZ"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"İNCELE"'], fill=sari))
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))

    ws = wb["VERI_A"]
    ws.conditional_formatting.add(f"C{ILK}:C{SON}",
        CellIsRule(operator="greaterThan", formula=["100000"], fill=sari))
    ws.conditional_formatting.add(f"D{ILK}:D{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["VERI_B"]
    ws.conditional_formatting.add(f"B{ILK}:B{SON}",
        CellIsRule(operator="greaterThan", formula=["100000"], fill=sari))
    ws.conditional_formatting.add(f"D{ILK}:D{SON}",
        CellIsRule(operator="greaterThan", formula=["0.25"], fill=sari))

    ws = wb["ANAHTAR"]
    ws.conditional_formatting.add(f"F{ILK}:F{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"N{ILK}:N{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))


def adlari_bagla(wb):
    for i, ana in enumerate([
        "toleransKurus", "raporTarihi", "mad_esikEslesme", "mad_esikFark",
        "mad_esikSapma", "mad_esikIncele", "mad_esikKritik", "mad_esikKritikOran",
        "mad_dengeCeza", "mad_esikCeza",
        "mad_azamiTutar", "mad_ikincilUzunluk", "mad_olcekHedef", "mad_yuzdelikOran",
        "mad_esikSenaryo", "mad_fiyatModeli", "mad_kvkkBeyani",
        "mad_anomaliMotoru", "mad_normAnahtar", "mad_tanimsizKova",
        "mad_kovaButunlukAd", "mad_dosyaSurumu", "mad_firmaUnvan", "mad_kuralSeti",
        "mad_eslesmeOraniAd",
    ]):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "mad_mutabakatFarki": 32, "mad_eslesmeOrani": 33, "mad_kovaButunluk": 34,
        "mad_modulT1": 35, "mad_modulT2": 36, "mad_modulO1": 37, "mad_modulO2": 38,
        "mad_modulO3": 39, "mad_modulO6": 40, "mad_modulO8": 41, "mad_modulMSen": 42,
        "mad_modulI1": 43, "mad_modulI2": 44,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "mad_modulT1Yorum": 47, "mad_modulT2Yorum": 48, "mad_modulO1Yorum": 49,
        "mad_modulO2Yorum": 50, "mad_modulO3Yorum": 51, "mad_modulO6Yorum": 52,
        "mad_modulO8Yorum": 53, "mad_modulMSenYorum": 54, "mad_modulI1Yorum": 55,
        "mad_modulI2Yorum": 56,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeHesapSiniflari", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeDonemler", "LISTELER!$C$7:$C$11")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$E$7:$E$8")
    ad_ekle(wb, "ListeDurumlar", "LISTELER!$G$7:$G$9")


def main(cikti_yolu=None):
    wb = Workbook()
    siralar = [
        (kapak, "KAPAK"),
        (veri_a, "VERI_A"),
        (veri_b, "VERI_B"),
        (anahtar, "ANAHTAR"),
        (motor, "MOTOR"),
        (eslestirme, "ESLESTIRME"),
        (kok_neden, "KOK_NEDEN"),
        (pano, "PANO"),
        (aksiyon, "AKSIYON"),
        (kanit_raporu, "KANIT_RAPORU"),
        (kontroller, "KONTROLLER"),
        (ornek_veri, "ORNEK_VERI"),
        (listeler, "LISTELER"),
        (ayarlar, "AYARLAR"),
        (kilavuz, "KILAVUZ"),
    ]
    ilk = wb.active
    for i, (fn, ad) in enumerate(siralar):
        ws = ilk if i == 0 else wb.create_sheet()
        fn(ws)

    adlari_bagla(wb)
    _cok_dogrulama(wb)
    kosullu_bicim(wb)

    tablo_formullerini_hucrelere_yaz(wb, satir_basi=ILK, satir_sonu=SON)
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=16, satir_sonu=15 + KAPASITE)

    for ws in wb.worksheets:
        sayfa_koru(ws)

    wb.calculation.fullCalcOnLoad = True
    urun_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    dosya_ad = "MizanAnomaliDenetimOncesiKontrol.xlsx"
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
