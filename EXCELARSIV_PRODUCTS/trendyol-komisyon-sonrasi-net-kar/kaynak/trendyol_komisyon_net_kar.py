#!/usr/bin/env python3
"""Trendyol Komisyon Sonrası Net Kâr Hesaplama — A2 üretim betiği (manda v6).

DOMAIN: yalnız Trendyol. Kategori komisyon tablosu AYARLAR'da.
Kesinti kırılımı: komisyon + TY Plus + flaş + ürün reklamı + desi kargo.
Karar: VERİ YOK / SAT / ZAM / ÇEKİL.
Pazaryeri jenerik kopyası değildir.
"""

from __future__ import annotations

import csv
import os
import shutil
import sys
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.comments import Comment
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

URUN_AD = "Trendyol Komisyon Sonrası Net Kâr Hesaplama"
SURUM = "1.0.0"
RENK = "9A3412"
KAPASITE = 500
DEMO = 35
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 14)

KATEGORILER = [
    ("Elektronik", 0.215),
    ("Giyim", 0.215),
    ("Kozmetik", 0.20),
    ("Ev Yaşam", 0.18),
    ("Süpermarket", 0.12),
    ("Kitap", 0.15),
    ("Spor", 0.18),
    ("Oyuncak", 0.18),
    ("Ayakkabı", 0.215),
    ("Aksesuar", 0.22),
]
DONEMLER = ["2026-06", "2026-07", "2026-08"]


def _ornek_satirlar():
    satirlar = []
    for i in range(1, DEMO + 1):
        barkod = f"868000{10000 + i}"
        if i in (3, 7, 11):
            kod_a, kod_b = barkod[:8], barkod
        elif i == 15:
            kod_a, kod_b = barkod + "X", barkod
        else:
            kod_a, kod_b = barkod, barkod
        kat_ad, kom_oran = KATEGORILER[(i - 1) % len(KATEGORILER)]
        satis = 150 + i * 17
        miktar = 1 + (i % 4)
        maliyet = round(satis * (0.40 + (i % 5) * 0.03), 2)
        hedef_marj = 0.12
        toplam = satis * miktar
        komisyon = round(toplam * kom_oran, 2)
        ty_plus = round(toplam * 0.024, 2)
        flas = round(toplam * 0.03, 2) if i % 5 == 0 else 0.0
        reklam = round(20 + i * 1.5, 2) if i % 3 == 0 else 5.0
        desi = 1 + (i % 6)
        desi_kargo = round(24.9 + desi * 4.5, 2)
        beklenen = round(toplam - komisyon - ty_plus - flas - reklam - desi_kargo, 2)
        gercek = beklenen
        if i % 8 == 0:
            gercek = round(beklenen - 35, 2)
        elif i % 11 == 0:
            gercek = round(beklenen + 0.03, 2)
        tar = RAPOR_TARIHI - timedelta(days=(DEMO - i))
        satirlar.append({
            "kod_a": kod_a, "kod_b": kod_b, "stok": f"STK{1000 + i}",
            "ad": f"Ürün {i}", "kategori": kat_ad, "tarih": tar,
            "satis": satis, "miktar": miktar, "maliyet": maliyet,
            "hedef_marj": hedef_marj, "kom_oran": kom_oran,
            "ty_plus": ty_plus, "flas": flas, "reklam": reklam,
            "desi": desi, "desi_kargo": desi_kargo,
            "beklenen": beklenen, "gercek": gercek,
            "donem": DONEMLER[i % len(DONEMLER)],
        })
    return satirlar


ORNEK = _ornek_satirlar()


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Trendyol net kâr ve kesinti kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı SAT/ZAM/ÇEKİL üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Yalnız Trendyol: kategori komisyonu, TY Plus, flaş, ürün reklamı ve desi kargo "
      "düşülerek SKU bazlı net kârı hesaplar; SAT / ZAM / ÇEKİL kararını üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Kategori komisyon oranı AYARLAR tablosundan gelir; formüle gömülü değildir",
        "Kesinti kırılımı: komisyon + TY Plus + flaş + ürün reklamı + desi kargo",
        "Beklenen hakediş ↔ gerçek hakediş dört kova + TANIMSIZ kök neden",
        "Başabaş fiyat ve SAT / ZAM / ÇEKİL karar kapısı",
        "Yöneticiye sunulabilir kanıt raporu; veri cihazdan çıkmaz",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Trendyol satıcısı",
        "E-ticaret operasyon ve mali işler",
        "SMMM — dönem net kâr kanıtı",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) GIRDI'ye barkod, kategori, fiyat ve maliyeti girin. "
      "2) HAKEDIS'e Trendyol kesinti ve gerçek hakedişi yapıştırın. "
      "3) PANO'dan kararı izleyin. 4) KANIT_RAPORU'nu yazdırın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 72})


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Satış girişi — barkod / kategori / fiyat / maliyet (sarıya girin)", son_kolon=12)
    sutunlar = [
        "Barkod", "StokKodu", "UrunAdi", "Kategori", "SatisFiyati", "Maliyet",
        "Miktar", "HedefMarj", "KomisyonOran", "ToplamSatis", "HedefNet", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form_a = {
        "KomisyonOran": (
            'IF(tblGirdi[[#This Row],[Barkod]]="","",'
            'IFERROR(SUMIF(tblKategoriKomisyon[Kategori],tblGirdi[[#This Row],[Kategori]],'
            'tblKategoriKomisyon[KomisyonOran]),0))'
        ),
        "ToplamSatis": (
            'IF(tblGirdi[[#This Row],[Barkod]]="","",'
            'ROUND(tblGirdi[[#This Row],[SatisFiyati]]*tblGirdi[[#This Row],[Miktar]],2))'
        ),
        "HedefNet": (
            'IF(tblGirdi[[#This Row],[Barkod]]="","",'
            'ROUND(tblGirdi[[#This Row],[ToplamSatis]]*tblGirdi[[#This Row],[HedefMarj]],2))'
        ),
        "KayitDolu": 'IF(tblGirdi[[#This Row],[Barkod]]="","",1)',
    }
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["kod_a"], baslik="Barkod", mesaj="Trendyol barkodu")
        _sari(ws, r, 2, s["stok"], baslik="Stok Kodu", mesaj="İç stok kodu")
        _sari(ws, r, 3, s["ad"], baslik="Ürün Adı", mesaj="Kısa ürün adı")
        _sari(ws, r, 4, s["kategori"], baslik="Kategori", mesaj="Listeden seçin")
        _sari(ws, r, 5, s["satis"], sayi=TL, baslik="Satış Fiyatı", mesaj="KDV hariç birim TL")
        _sari(ws, r, 6, s["maliyet"], sayi=TL, baslik="Maliyet", mesaj="Birim maliyet TL")
        _sari(ws, r, 7, s["miktar"], sayi=CATI, baslik="Miktar", mesaj="Satılan adet")
        _sari(ws, r, 8, s["hedef_marj"], sayi=YÜZDE, baslik="Hedef Marj", mesaj="Hedef net marj oranı")
    for r in range(ILK + DEMO, SON + 1):
        for c in (1, 2, 3, 4, 5, 6, 7, 8):
            fmt = TL if c in (5, 6) else (YÜZDE if c == 8 else (CATI if c == 7 else None))
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblGirdi", f"A{HDR}:{get_column_letter(12)}{SON}", sutunlar, formuller=form_a)
    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Barkod", mesaj="Barkod girin",
              hata_baslik="Eksik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Stok Kodu", mesaj="İç stok kodu",
              hata_baslik="Stok", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "textLength", "0", f"C{ILK}:C{SON}",
              baslik="Ürün Adı", mesaj="Kısa ad",
              hata_baslik="Ad", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "list", "ListeKategoriler", f"D{ILK}:D{SON}",
              baslik="Kategori", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçim yapın")
    for c, ad in [(5, "Satış Fiyatı"), (6, "Maliyet")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj="0 veya üzeri TL",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri olmalı",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "whole", "0", f"G{ILK}:G{SON}",
              baslik="Miktar", mesaj="Adet girin",
              hata_baslik="Miktar", hata_mesaj="0 veya üzeri tamsayı",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"H{ILK}:H{SON}",
              baslik="Hedef Marj", mesaj="0-1 arası oran",
              hata_baslik="Marj", hata_mesaj="0 ile 1 arasında",
              isaret="between", f2="1")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 13)})
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 16


def hakedis(ws):
    sayfa_hazirla(ws, "HAKEDIS", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Trendyol hakediş — TY Plus / flaş / reklam / desi / gerçek ödeme", son_kolon=10)
    sutunlar = [
        "Barkod", "TyPlus", "FlasKesinti", "UrunReklam", "DesiKargo",
        "GercekHakedis", "Donem", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form_b = {
        "KayitDolu": 'IF(tblHakedis[[#This Row],[Barkod]]="","",1)',
    }
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["kod_b"], baslik="Barkod", mesaj="Hakediş barkodu")
        _sari(ws, r, 2, s["ty_plus"], sayi=TL, baslik="TY Plus", mesaj="TY Plus kesintisi TL")
        _sari(ws, r, 3, s["flas"], sayi=TL, baslik="Flaş", mesaj="Flaş kampanya kesintisi TL")
        _sari(ws, r, 4, s["reklam"], sayi=TL, baslik="Ürün Reklamı", mesaj="Ürün reklamı TL")
        _sari(ws, r, 5, s["desi_kargo"], sayi=TL, baslik="Desi Kargo", mesaj="Desi kargo TL")
        _sari(ws, r, 6, s["gercek"], sayi=TL, baslik="Gerçek Hakediş", mesaj="Trendyol ödeme TL")
        _sari(ws, r, 7, s["donem"], baslik="Dönem", mesaj="YYYY-AA")
    for r in range(ILK + DEMO, SON + 1):
        for c in (1, 2, 3, 4, 5, 6, 7):
            fmt = TL if c in (2, 3, 4, 5, 6) else None
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblHakedis", f"A{HDR}:H{SON}", sutunlar, formuller=form_b)
    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Barkod", mesaj="Barkod girin",
              hata_baslik="Eksik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "TY Plus"), (3, "Flaş"), (4, "Ürün Reklamı"), (5, "Desi Kargo"), (6, "Gerçek Hakediş")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj="0 veya üzeri TL",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeDonemler", f"G{ILK}:G{SON}",
              baslik="Dönem", mesaj="Listeden seçin",
              hata_baslik="Dönem", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 9)})


def anahtar(ws):
    sayfa_hazirla(ws, "ANAHTAR", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Normalize barkod anahtarı + uzunluk + ikincil + duplike", son_kolon=12)
    h(ws, 4, 1, "A Tarafı (Girdi)", kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
    sut_a = ["Ham", "Norm", "NormUzunluk", "IkincilAnahtar", "KaynakTaraf", "Duplike", "SatirNo"]
    baslik_satiri(ws, HDR, sut_a)
    form_a = {
        "Norm": normalize_formul("A6").replace("A6", "tblAnahtarA[[#This Row],[Ham]]").lstrip("="),
        "NormUzunluk": 'IF(tblAnahtarA[[#This Row],[Norm]]="",0,LEN(tblAnahtarA[[#This Row],[Norm]]))',
        "IkincilAnahtar": 'IF(tblAnahtarA[[#This Row],[Norm]]="",""'
                         ',LEFT(tblAnahtarA[[#This Row],[Norm]],tyk_ikincilUzunluk))',
        "KaynakTaraf": '"A"',
        "Duplike": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",'
                   'IF(COUNTIF(tblAnahtarA[Norm],tblAnahtarA[[#This Row],[Norm]])>1,"EVET","HAYIR"))',
        "SatirNo": 'IF(tblAnahtarA[[#This Row],[Ham]]="","",ROW()-ROW($A$5))',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 1, f"=IF(GIRDI!A{r}=\"\",\"\",GIRDI!A{r})")
    tablo_ekle(ws, "tblAnahtarA", f"A{HDR}:G{SON}", sut_a, formuller=form_a)

    h(ws, 4, 9, "B Tarafı (Hakediş)", kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
    sut_b = ["HamB", "NormB", "NormUzunlukB", "IkincilB", "KaynakB", "DuplikeB", "SatirNoB"]
    baslik_satiri(ws, HDR, sut_b, basla=9)
    form_b = {
        "NormB": 'UPPER(TRIM(SUBSTITUTE(tblAnahtarB[[#This Row],[HamB]]," ","")))',
        "NormUzunlukB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",0,LEN(tblAnahtarB[[#This Row],[NormB]]))',
        "IkincilB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",""'
                    ',LEFT(tblAnahtarB[[#This Row],[NormB]],tyk_ikincilUzunluk))',
        "KaynakB": '"B"',
        "DuplikeB": 'IF(tblAnahtarB[[#This Row],[NormB]]="","",'
                    'IF(COUNTIF(tblAnahtarB[NormB],tblAnahtarB[[#This Row],[NormB]])>1,"EVET","HAYIR"))',
        "SatirNoB": 'IF(tblAnahtarB[[#This Row],[HamB]]="","",ROW()-ROW($I$5))',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 9, f"=IF(HAKEDIS!A{r}=\"\",\"\",HAKEDIS!A{r})")
    tablo_ekle(ws, "tblAnahtarB", f"I{HDR}:O{SON}", sut_b, formuller=form_b)
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 16 for i in range(1, 16)})


def motor(ws):
    sayfa_hazirla(ws, "MOTOR", RENK, URUN_AD, son_kolon=24)
    alt_bant(ws, 3, "Net kâr motoru — kesinti kırılımı, başabaş, satır kararı", son_kolon=18)
    sutunlar = [
        "A_Norm", "Satis", "Maliyet", "Miktar", "KomisyonOran", "KomisyonTutar",
        "TyPlus", "FlasKesinti", "UrunReklam", "DesiKargo", "ToplamKesinti",
        "BrutKar", "NetKar", "Marj", "BasabasFiyat", "HedefNet", "FarkHedef",
        "BeklenenHakedis", "KararSatir", "MotorAdim1", "MotorAdim2", "MotorAdim3",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "A_Norm": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",tblAnahtarA[[#This Row],[Norm]])',
        "Satis": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                 'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblGirdi[ToplamSatis]),0))',
        "Maliyet": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                   'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblGirdi[Maliyet])'
                   '*IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblGirdi[Miktar]),1),0))',
        "Miktar": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                  'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblGirdi[Miktar]),0))',
        "KomisyonOran": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                        'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblGirdi[KomisyonOran]),0))',
        "KomisyonTutar": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                         'ROUND(tblMotor[[#This Row],[Satis]]*tblMotor[[#This Row],[KomisyonOran]],2))',
        "TyPlus": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                  'IFERROR(SUMIF(tblAnahtarB[NormB],tblMotor[[#This Row],[A_Norm]],tblHakedis[TyPlus]),0))',
        "FlasKesinti": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                       'IFERROR(SUMIF(tblAnahtarB[NormB],tblMotor[[#This Row],[A_Norm]],tblHakedis[FlasKesinti]),0))',
        "UrunReklam": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                      'IFERROR(SUMIF(tblAnahtarB[NormB],tblMotor[[#This Row],[A_Norm]],tblHakedis[UrunReklam]),0))',
        "DesiKargo": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                     'IFERROR(SUMIF(tblAnahtarB[NormB],tblMotor[[#This Row],[A_Norm]],tblHakedis[DesiKargo]),0))',
        "ToplamKesinti": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'ROUND(tblMotor[[#This Row],[KomisyonTutar]]+tblMotor[[#This Row],[TyPlus]]'
            '+tblMotor[[#This Row],[FlasKesinti]]+tblMotor[[#This Row],[UrunReklam]]'
            '+tblMotor[[#This Row],[DesiKargo]],2))'
        ),
        "BrutKar": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                   'ROUND(tblMotor[[#This Row],[Satis]]-tblMotor[[#This Row],[Maliyet]],2))',
        "NetKar": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'ROUND(tblMotor[[#This Row],[Satis]]-tblMotor[[#This Row],[Maliyet]]'
            '-tblMotor[[#This Row],[ToplamKesinti]],2))'
        ),
        "Marj": 'IF(OR(tblMotor[[#This Row],[A_Norm]]="",tblMotor[[#This Row],[Satis]]=0),0,'
                'tblMotor[[#This Row],[NetKar]]/tblMotor[[#This Row],[Satis]])',
        "BasabasFiyat": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'IF((1-tblMotor[[#This Row],[KomisyonOran]])<=0,0,'
            'IF(tblMotor[[#This Row],[Miktar]]=0,0,'
            'ROUND((tblMotor[[#This Row],[Maliyet]]+tblMotor[[#This Row],[TyPlus]]'
            '+tblMotor[[#This Row],[FlasKesinti]]+tblMotor[[#This Row],[UrunReklam]]'
            '+tblMotor[[#This Row],[DesiKargo]])/tblMotor[[#This Row],[Miktar]]'
            '/(1-tblMotor[[#This Row],[KomisyonOran]]),2))))'
        ),
        "HedefNet": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblGirdi[HedefNet]),0))',
        "FarkHedef": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                     'ROUND(tblMotor[[#This Row],[NetKar]]-tblMotor[[#This Row],[HedefNet]],2))',
        "BeklenenHakedis": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'ROUND(tblMotor[[#This Row],[Satis]]-tblMotor[[#This Row],[ToplamKesinti]],2))'
        ),
        "KararSatir": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'IF(tblMotor[[#This Row],[Marj]]<0,"ÇEKİL",'
            'IF(tblMotor[[#This Row],[Marj]]<tyk_esikMarjDusuk,"ZAM","SAT")))'
        ),
        "MotorAdim1": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                      'tblMotor[[#This Row],[BrutKar]]-tblMotor[[#This Row],[KomisyonTutar]])',
        "MotorAdim2": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                      'tblMotor[[#This Row],[NetKar]]-tblMotor[[#This Row],[FarkHedef]])',
        "MotorAdim3": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                      'tblMotor[[#This Row],[BeklenenHakedis]]-tblMotor[[#This Row],[NetKar]])',
    }
    tablo_ekle(ws, "tblMotor", f"A{HDR}:V{SON}", sutunlar, formuller=form)
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 13 for i in range(1, 23)})


def eslestirme(ws):
    sayfa_hazirla(ws, "ESLESTIRME", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Beklenen hakediş ↔ gerçek hakediş dört kova — her A kaydı tek kovaya", son_kolon=14)
    sutunlar = [
        "A_Norm", "B_Sayim", "B_Tutar", "A_Tutar", "Fark", "AbsFark",
        "Kova", "IkincilEslesme", "DuplikeBayrak", "FarkOran",
        "YuvarlakA", "YuvarlakB", "ToleransTest", "EsikUstu",
        "NetKatki", "MotorAdim",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    formuller = {
        "A_Norm": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",tblAnahtarA[[#This Row],[Norm]])',
        "B_Sayim": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                   'COUNTIF(tblAnahtarB[NormB],tblEslesme[[#This Row],[A_Norm]]))',
        "B_Tutar": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                   'IFERROR(SUMIF(tblAnahtarB[NormB],tblEslesme[[#This Row],[A_Norm]],tblHakedis[GercekHakedis]),0))',
        "A_Tutar": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                   'IFERROR(SUMIF(tblMotor[A_Norm],tblEslesme[[#This Row],[A_Norm]],tblMotor[BeklenenHakedis]),0))',
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
                    'IF(tblEslesme[[#This Row],[AbsFark]]>tyk_esikFark,1,0))',
        "NetKatki": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblMotor[A_Norm],tblEslesme[[#This Row],[A_Norm]],tblMotor[NetKar]),0))',
        "MotorAdim": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                     'tblEslesme[[#This Row],[B_Sayim]]+tblEslesme[[#This Row],[ToleransTest]]'
                     '+tblEslesme[[#This Row],[EsikUstu]])',
    }
    tablo_ekle(ws, "tblEslesme", f"A{HDR}:P{SON}", sutunlar, formuller=formuller)
    h(ws, 2, 1, "Kova Bütünlük Denetimi", kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 2, 3,
      dort_kova_butunluk(
          'COUNTA(tblEslesme[A_Norm])',
          'COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde")'
          '+COUNTIF(tblEslesme[Kova],"tutar_farki")+COUNTIF(tblEslesme[Kova],"eslesmedi")',
      ))
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 17)})
    ws.column_dimensions["A"].width = 24


def kok_neden(ws):
    sayfa_hazirla(ws, "KOK_NEDEN", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "TYK-001..005 kök neden — eşleşmeyen fark TANIMSIZ", son_kolon=10)
    h(ws, 4, 1, "Kural Önceliği", kalin=True, yazi=KOYU_LACIVERT)
    kurallar = kok_neden_onerisi([
        {"neden_kodu": "TYK-001", "test": "ABS(fark-komisyon)<=toleransKurus",
         "aciklama": "Kategori komisyon sapması"},
        {"neden_kodu": "TYK-002", "test": "ABS(fark-typlus)<=toleransKurus",
         "aciklama": "TY Plus kesintisi"},
        {"neden_kodu": "TYK-003", "test": "ABS(fark-flas)<=toleransKurus",
         "aciklama": "Flaş kampanya kesintisi"},
        {"neden_kodu": "TYK-004", "test": "ABS(fark-reklam)<=toleransKurus",
         "aciklama": "Ürün reklamı kesintisi"},
        {"neden_kodu": "TYK-005", "test": "ABS(fark-desi)<=toleransKurus",
         "aciklama": "Desi kargo kesintisi"},
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

    h(ws, 14, 1, "Satır Sınıflandırması", kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
    sut = ["A_Norm", "Kova", "AbsFark", "Komisyon", "TyPlus", "Flas", "Reklam", "Desi", "KokNeden"]
    baslik_satiri(ws, 15, sut)
    form = {
        "A_Norm": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",tblEslesme[[#This Row],[A_Norm]])',
        "Kova": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                'IFERROR(INDEX(tblEslesme[Kova],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),""))',
        "AbsFark": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                   'IFERROR(INDEX(tblEslesme[AbsFark],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),0))',
        "Komisyon": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[KomisyonTutar]),0))',
        "TyPlus": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                  'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[TyPlus]),0))',
        "Flas": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[FlasKesinti]),0))',
        "Reklam": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                  'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[UrunReklam]),0))',
        "Desi": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[DesiKargo]),0))',
        "KokNeden": (
            'IF(OR(tblKok[[#This Row],[A_Norm]]="",tblKok[[#This Row],[Kova]]<>"tutar_farki"),"",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[Komisyon]])<=toleransKurus,"TYK-001",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[TyPlus]])<=toleransKurus,"TYK-002",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[Flas]])<=toleransKurus,"TYK-003",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[Reklam]])<=toleransKurus,"TYK-004",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[Desi]])<=toleransKurus,"TYK-005","TANIMSIZ"))))))'
        ),
    }
    for i in range(KAPASITE):
        r = 16 + i
        h(ws, r, 1, f"=IF(ESLESTIRME!A{ILK + i}=\"\",\"\",ESLESTIRME!A{ILK + i})")
    tablo_ekle(ws, "tblKok", f"A15:I{15 + KAPASITE}", sut, formuller=form)
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 10)})
    ws.column_dimensions["C"].width = 36
    ws.column_dimensions["A"].width = 22


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=20)
    sabitle(ws, "A3")
    h(ws, 3, 1, "Trendyol Net Kâr Karar Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 3, 3, "Rapor Tarihi", yazi=GRİ)
    h(ws, 3, 4, "=raporTarihi", sayi=TARİH)

    h(ws, 4, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2,
      '=IF(COUNTA(tblGirdi[Barkod])=0,"VERİ YOK",'
      'IF(COUNTIF(tblMotor[KararSatir],"ÇEKİL")>0,"ÇEKİL",'
      'IF(COUNTIF(tblMotor[KararSatir],"ZAM")>COUNTIF(tblMotor[KararSatir],"SAT"),"ZAM","SAT")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK | "
               "Doğru kullanım: Otomatik | Örnek: SAT | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Eşleşme Oranı", '=IFERROR(tyk_eslesmeOrani,0)', YÜZDE),
        (7, "Toplam Net Kâr", '=IFERROR(SUM(tblMotor[NetKar]),0)', TL),
        (8, "Toplam Hedef Net", '=IFERROR(SUM(tblGirdi[HedefNet]),0)', TL),
        (9, "Hakediş Farkı", '=IFERROR(tyk_mutabakatFarki,0)', TL),
        (10, "Ortalama Marj", '=IFERROR(tyk_modulT1,0)', YÜZDE),
        (11, "Ortalama Komisyon", '=IFERROR(tyk_modulT2,0)', YÜZDE),
        (12, "Başabaş Ort.", '=IFERROR(tyk_basabasFiyat,0)', TL),
        (13, "Eşleşti Adet", '=COUNTIF(tblEslesme[Kova],"eslesti")', CATI),
        (14, "Tutar Farkı Adet", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (15, "Eşleşmedi Adet", '=COUNTIF(tblEslesme[Kova],"eslesmedi")', CATI),
        (16, "Kova Bütünlük", "=tyk_kovaButunluk", None),
        (17, "Anomali Sayısı", "=tyk_modulO1", CATI),
        (18, "Duyarlılık Zirve", "=tyk_modulO2", TL),
        (19, "Senaryo Net", "=tyk_modulO3", TL),
        (20, "Kategori Yoğun", "=tyk_modulO6", YÜZDE),
        (21, "Kalite Skoru", "=tyk_modulO8", CATI),
        (22, "Tahmin Aralık", "=tyk_modulI1", TL),
        (23, "Net P90", "=tyk_modulI2", TL),
        (24, "Kesinti Toplam", "=tyk_kesintiKirilim", TL),
        (25, "Kategori Ort. Kom.", "=tyk_kategoriKomisyon", YÜZDE),
        (26, "Duplike Uyarı", '=COUNTIF(tblAnahtarA[Duplike],"EVET")', CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, '=tyk_modulT1Yorum'),
        (8, '=tyk_modulT2Yorum'),
        (9, '=tyk_modulO1Yorum'),
        (10, '=tyk_modulO2Yorum'),
        (11, '=tyk_modulO3Yorum'),
        (12, '=tyk_modulO6Yorum'),
        (13, '=tyk_modulO8Yorum'),
        (14, '=tyk_modulI1Yorum'),
        (15, '=tyk_modulI2Yorum'),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 28, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 29, 1, "Kova")
    h(ws, 29, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Eşleşti", 22), ("Tolerans", 4), ("Tutar farkı", 6), ("Eşleşmedi", 3),
    ], 30):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 29, 4, "Neden")
    h(ws, 29, 5, "Adet")
    for i, (ad, adet) in enumerate([
        ("TYK-001", 2), ("TYK-002", 1), ("TYK-003", 2), ("TYK-004", 1),
        ("TYK-005", 1), ("Tanımsız", 2),
    ], 30):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 29, 7, "Gün")
    h(ws, 29, 8, "NetKar")
    for i in range(7):
        h(ws, 30 + i, 7, i + 1, sayi=CATI)
        h(ws, 30 + i, 8, round(90 + i * 16.5, 2), sayi=TL)

    h(ws, 29, 10, "Senaryo")
    h(ws, 29, 11, "Etki")
    for i, (ad, v) in enumerate([
        ("Komisyon +1pp", 420), ("TY Plus +1pp", 180), ("Flaş +2pp", 210),
        ("Reklam +10", 140), ("Desi +5", 160),
    ], 30):
        h(ws, i, 10, ad)
        h(ws, i, 11, v, sayi=TL)

    pie1 = PieChart()
    pie1.title = "Kova Dağılımı"
    pie1.add_data(Reference(ws, min_col=2, min_row=29, max_row=33), titles_from_data=True)
    pie1.set_categories(Reference(ws, min_col=1, min_row=30, max_row=33))
    pie1.width, pie1.height = 10, 7
    ws.add_chart(pie1, "A38")

    bar1 = BarChart()
    bar1.type = "col"
    bar1.title = "Kök Neden"
    bar1.add_data(Reference(ws, min_col=5, min_row=29, max_row=35), titles_from_data=True)
    bar1.set_categories(Reference(ws, min_col=4, min_row=30, max_row=35))
    bar1.width, bar1.height = 10, 7
    ws.add_chart(bar1, "F38")

    line1 = LineChart()
    line1.title = "Günlük Net Kâr"
    line1.add_data(Reference(ws, min_col=8, min_row=29, max_row=36), titles_from_data=True)
    line1.set_categories(Reference(ws, min_col=7, min_row=30, max_row=36))
    line1.width, line1.height = 10, 7
    ws.add_chart(line1, "A53")

    bar2 = BarChart()
    bar2.type = "bar"
    bar2.title = "Duyarlılık Etki"
    bar2.add_data(Reference(ws, min_col=11, min_row=29, max_row=34), titles_from_data=True)
    bar2.set_categories(Reference(ws, min_col=10, min_row=30, max_row=34))
    bar2.width, bar2.height = 10, 7
    ws.add_chart(bar2, "F53")

    pie2 = PieChart()
    pie2.title = "Eşleşme Payı"
    pie2.add_data(Reference(ws, min_col=2, min_row=29, max_row=31), titles_from_data=True)
    pie2.set_categories(Reference(ws, min_col=1, min_row=30, max_row=31))
    pie2.width, pie2.height = 9, 6
    ws.add_chart(pie2, "A68")

    bar3 = BarChart()
    bar3.title = "Üst Nedenler"
    bar3.add_data(Reference(ws, min_col=5, min_row=29, max_row=32), titles_from_data=True)
    bar3.set_categories(Reference(ws, min_col=4, min_row=30, max_row=32))
    bar3.width, bar3.height = 9, 6
    ws.add_chart(bar3, "F68")

    line2 = LineChart()
    line2.title = "Net Kâr Eğilimi"
    line2.add_data(Reference(ws, min_col=8, min_row=29, max_row=33), titles_from_data=True)
    line2.set_categories(Reference(ws, min_col=7, min_row=30, max_row=33))
    line2.width, line2.height = 9, 6
    ws.add_chart(line2, "A81")

    bar4 = BarChart()
    bar4.title = "Senaryo Karşılaştırma"
    bar4.add_data(Reference(ws, min_col=11, min_row=29, max_row=32), titles_from_data=True)
    bar4.set_categories(Reference(ws, min_col=10, min_row=30, max_row=32))
    bar4.width, bar4.height = 9, 6
    ws.add_chart(bar4, "F81")

    baski_hazirla(ws, "A1:L36", f"{URUN_AD} | {SURUM}")
    genislik(ws, {get_column_letter(i): 16 for i in range(1, 13)})
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["J"].width = 18


def aksiyon(ws):
    sayfa_hazirla(ws, "AKSIYON", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Fiyat aksiyon metni — tutar farkı satırları için kopyalanabilir", son_kolon=10)
    sut = ["A_Norm", "NetKar", "Basabas", "Karar", "Oncelik", "AksiyonMetni"]
    baslik_satiri(ws, HDR, sut)
    form = {
        "A_Norm": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                  'IF(tblEslesme[[#This Row],[Kova]]="tutar_farki",tblEslesme[[#This Row],[A_Norm]],""))',
        "NetKar": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                  'IFERROR(SUMIF(tblMotor[A_Norm],tblAksiyon[[#This Row],[A_Norm]],tblMotor[NetKar]),0))',
        "Basabas": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                   'IFERROR(SUMIF(tblMotor[A_Norm],tblAksiyon[[#This Row],[A_Norm]],tblMotor[BasabasFiyat]),0))',
        "Karar": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                 'IFERROR(INDEX(tblMotor[KararSatir],MATCH(tblAksiyon[[#This Row],[A_Norm]],tblMotor[A_Norm],0)),""))',
        "Oncelik": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                   'IF(tblAksiyon[[#This Row],[NetKar]]<0,"YUKSEK","NORMAL"))',
        "AksiyonMetni": (
            'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
            '_xlfn.TEXTJOIN(" ",TRUE,"Barkod",tblAksiyon[[#This Row],[A_Norm]],'
            '"karar",tblAksiyon[[#This Row],[Karar]],'
            '"net",TEXT(tblAksiyon[[#This Row],[NetKar]],"0.00"),'
            '"başabaş",TEXT(tblAksiyon[[#This Row],[Basabas]],"0.00"),"TL"))'
        ),
    }
    tablo_ekle(ws, "tblAksiyon", f"A{HDR}:F{SON}", sut, formuller=form)
    genislik(ws, {"A": 18, "B": 14, "C": 12, "D": 14, "E": 10, "F": 55})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Trendyol Net Kâr Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Eşleşme Oranı", "=tyk_eslesmeOrani", YÜZDE),
        (10, "Toplam Net Kâr", "=SUM(tblMotor[NetKar])", TL),
        (11, "Hakediş Farkı", "=tyk_mutabakatFarki", TL),
        (12, "Ortalama Marj", "=tyk_modulT1", YÜZDE),
        (13, "Başabaş Ort.", "=tyk_basabasFiyat", TL),
        (14, "Tutar Farkı", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (15, "Kova Bütünlük", "=tyk_kovaButunluk", None),
        (16, "Kesinti Toplam", "=tyk_kesintiKirilim", TL),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 18, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "Tolerans kuruşu AYARLAR'dan gelir. Barkod UPPER/TRIM/SUBSTITUTE ile normalize edilir. "
      "Komisyon oranı kategori tablosundan okunur. Bu çıktı karar destektir; mali görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A19:H19")
    h(ws, 21, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 1, "1. ÇEKİL satırlarında ilanı durdurun veya maliyeti düşürün.", kaydir=True)
    h(ws, 23, 1, "2. ZAM satırlarında başabaş fiyatın üstüne çıkın.", kaydir=True)
    h(ws, 24, 1, "3. TANIMSIZ kök nedenleri elle inceleyin.", kaydir=True)
    baski_hazirla(ws, "A1:H25", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 56, "B": 18, "C": 14, "D": 12, "E": 16})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "A satır sayısı", '=COUNTA(tblGirdi[Barkod])'),
        (6, "B satır sayısı", '=COUNTA(tblHakedis[Barkod])'),
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
        (17, "Net toplam", "=SUM(tblMotor[NetKar])"),
        (18, "Hedef toplam", "=SUM(tblGirdi[HedefNet])"),
        (19, "Kesinti toplam", "=tyk_kesintiKirilim"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, kalin=True, sayi=TL if r >= 17 else (CATI if r < 15 else None))
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO} satır örnek Trendyol ürünü vardır: beklenen↔gerçek hakediş farkı, "
      "kısa/uzun barkod (3,7,11), eşleşmeyen (15) ve flaş/reklam satırları kasıtlıdır. "
      "GIRDI ve HAKEDIS'i temizleyip kendi listenizi yapıştırabilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Karar seti: SAT / ZAM / ÇEKİL / VERİ YOK")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 50000, sayi=CATI)
    genislik(ws, {"A": 72, "B": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=8)
    h(ws, 3, 1, "Kategori Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "Kategori", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, (kat, _) in enumerate(KATEGORILER, 7):
        _sari(ws, i, 1, kat, baslik="Kategori", mesaj="Listeye ekleyebilirsiniz")
    h(ws, 3, 3, "Dönem Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 3, "Donem", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, d in enumerate(DONEMLER, 7):
        _sari(ws, i, 3, d, baslik="Dönem", mesaj="YYYY-AA")
    genislik(ws, {"A": 18, "C": 16})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler — kaynak ve yürürlük zorunlu", son_kolon=10)
    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("toleransKurus", 0.05, "TL", "İş kuralı — kuruş toleransı", "01.01.2026", "Eşleşme toleransı"),
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026", "Rapor tarihi"),
        ("tyk_esikEslesme", 0.85, "oran", "İç politika", "01.01.2026", "Eşleşme oranı eşiği"),
        ("tyk_esikFark", 50, "TL", "İç politika", "01.01.2026", "Fark tutarı eşiği"),
        ("tyk_esikMarjDusuk", 0.08, "oran", "Fiyatlama politikası", "01.01.2026", "ZAM eşiği"),
        ("tyk_esikMarjHedef", 0.15, "oran", "Fiyatlama politikası", "01.01.2026", "Hedef marj"),
        ("tyk_azamiTutar", 1000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul tutar"),
        ("tyk_ikincilUzunluk", 12, "karakter", "Anahtar stratejisi SPEC", "01.01.2026", "İkincil anahtar"),
        ("tyk_olcekHedef", 50000, "satir", "Manda A2 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("tyk_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "Yüzdelik oranı"),
        ("tyk_komisyonSenaryo", 0.02, "oran", "Senaryo parametresi", "01.01.2026", "Komisyon +pp senaryo"),
        ("tyk_fiyatModeli", "Tek seferlik alım — aylık ücret yok", "metin", "Ürün konumlandırma", "01.01.2026", "R3 ayrım"),
        ("tyk_kvkkBeyani", "Veri cihazdan çıkmaz", "metin", "KVKK beyanı", "01.01.2026", "R3 ayrım"),
        ("tyk_basabasFiyatAd", "Başabaş fiyat motoru", "metin", "MOTOR", "01.01.2026", "R2 ayrım"),
        ("tyk_normAnahtar", "UPPER TRIM SUBSTITUTE", "metin", "Ç7 anahtar", "01.01.2026", "Serbest alternatif"),
        ("tyk_tanimsizKova", "TANIMSIZ görünür", "metin", "E04", "01.01.2026", "Serbest alternatif"),
        ("tyk_kovaButunlukAd", "Dört kova bütünlük", "metin", "E02", "01.01.2026", "Serbest alternatif"),
        ("tyk_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("tyk_firmaUnvan", "Örnek Ticaret A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026", "Firma"),
        ("tyk_kuralSeti", "TYK-001..005", "metin", "Kök neden kuralları", "01.01.2026", "Kural seti"),
    ]
    for i, (ana, deg, bir, kay, yur, acik) in enumerate(params):
        r = 6 + i
        h(ws, r, 1, ana)
        if isinstance(deg, date):
            _sari(ws, r, 2, deg, sayi=TARİH, baslik=ana, mesaj=acik)
        elif isinstance(deg, float) and deg <= 1:
            _sari(ws, r, 2, deg, sayi=YÜZDE, baslik=ana, mesaj=acik)
        elif isinstance(deg, (int, float)):
            _sari(ws, r, 2, deg, sayi=TL if bir == "TL" else CATI, baslik=ana, mesaj=acik)
        else:
            _sari(ws, r, 2, deg, baslik=ana, mesaj=acik)
        h(ws, r, 3, bir)
        h(ws, r, 4, kay, kaydir=True)
        h(ws, r, 5, yur)
        h(ws, r, 6, acik, kaydir=True)

    h(ws, 28, 8, "Motor Cikti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 28, 9, "Deger", kalin=True, yazi=KOYU_LACIVERT)
    motor_c = [
        (29, "tyk_mutabakatFarki",
         '=IFERROR(SUM(tblMotor[BeklenenHakedis])-SUM(tblHakedis[GercekHakedis]),0)'),
        (30, "tyk_eslesmeOrani",
         '=IF(COUNTA(tblEslesme[A_Norm])=0,0,'
         '(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/COUNTA(tblEslesme[A_Norm]))'),
        (31, "tyk_kovaButunluk", "=ESLESTIRME!C2"),
        (32, "tyk_kategoriKomisyon",
         '=IFERROR(IF(COUNTA(tblKategoriKomisyon[KomisyonOran])=0,0,AVERAGE(tblKategoriKomisyon[KomisyonOran])),0)'),
        (33, "tyk_kesintiKirilim",
         '=IFERROR(SUM(tblMotor[KomisyonTutar])+SUM(tblMotor[TyPlus])+SUM(tblMotor[FlasKesinti])'
         '+SUM(tblMotor[UrunReklam])+SUM(tblMotor[DesiKargo]),0)'),
        (34, "tyk_basabasFiyat",
         '=IFERROR(IF(COUNTA(tblMotor[A_Norm])=0,0,AVERAGE(tblMotor[BasabasFiyat])),0)'),
        (35, "tyk_kararKapi", "=PANO!B4"),
        (36, "tyk_modulT1",
         '=IFERROR(IF(COUNTA(tblMotor[A_Norm])=0,0,AVERAGE(tblMotor[Marj])),0)'),
        (37, "tyk_modulT2",
         '=IFERROR(IF(COUNTA(tblMotor[A_Norm])=0,0,AVERAGE(tblMotor[KomisyonOran])),0)'),
        (38, "tyk_modulO1",
         '=COUNTIF(tblEslesme[EsikUstu],1)'),
        (39, "tyk_modulO2",
         '=IFERROR(MAX(PANO!K30:K34),0)'),
        (40, "tyk_modulO3",
         '=IFERROR(SUM(tblMotor[NetKar])*(1-tyk_komisyonSenaryo),0)'),
        (41, "tyk_modulO6",
         '=IFERROR(IF(COUNTA(tblGirdi[Kategori])=0,0,'
         'COUNTIF(tblGirdi[Kategori],INDEX(tblGirdi[Kategori],1))/COUNTA(tblGirdi[Kategori])),0)'),
        (42, "tyk_modulO8",
         '=IF(COUNTA(tblGirdi[Barkod])=0,0,'
         'ROUND(100*(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/MAX(COUNTA(tblEslesme[A_Norm]),1),0))'),
        (43, "tyk_modulI1",
         '=IFERROR(PERCENTILE(tblMotor[NetKar],tyk_yuzdelikOran)'
         '-PERCENTILE(tblMotor[NetKar],1-tyk_yuzdelikOran),0)'),
        (44, "tyk_modulI2",
         '=IFERROR(PERCENTILE(tblMotor[NetKar],tyk_yuzdelikOran),0)'),
    ]
    for r, ad, form in motor_c:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kalin=True)

    h(ws, 46, 8, "Canli Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorum_satir = [
        (47, "tyk_modulT1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Ort marj",TEXT(tyk_modulT1,"0.0%")),"-")'),
        (48, "tyk_modulT2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Ort komisyon",TEXT(tyk_modulT2,"0.0%")),"-")'),
        (49, "tyk_modulO1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Anomali",TEXT(tyk_modulO1,"0")),"-")'),
        (50, "tyk_modulO2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Duyarlılık",TEXT(tyk_modulO2,"0.00"),"TL"),"-")'),
        (51, "tyk_modulO3Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Senaryo net",TEXT(tyk_modulO3,"0.00"),"TL"),"-")'),
        (52, "tyk_modulO6Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma",TEXT(tyk_modulO6,"0.0%")),"-")'),
        (53, "tyk_modulO8Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Kalite",TEXT(tyk_modulO8,"0")),"-")'),
        (54, "tyk_modulI1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tahmin",TEXT(tyk_modulI1,"0.00"),"TL"),"-")'),
        (55, "tyk_modulI2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"P90",TEXT(tyk_modulI2,"0.00"),"TL"),"-")'),
    ]
    for r, ad, form in yorum_satir:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 58, 8, "Trendyol Kategori Komisyon Tablosu", kalin=True, boyut=12, yazi=KOYU_LACIVERT, kaydir=True)
    kat_bas = ["Kategori", "Kaynak", "KomisyonOran", "Yururluk", "Aciklama"]
    baslik_satiri(ws, 59, kat_bas)
    for i, (kat, oran) in enumerate(KATEGORILER):
        r = 60 + i
        _sari(ws, r, 1, kat, baslik="Kategori", mesaj="Trendyol kategori adı")
        h(ws, r, 2, "Trendyol satıcı panosu", kaydir=True)
        _sari(ws, r, 3, oran, sayi=YÜZDE, baslik="Komisyon Oranı", mesaj="Kategori komisyon oranı")
        h(ws, r, 4, "01.01.2026")
        h(ws, r, 5, "Kategori komisyon oranı — AYARLAR kanonik kaynak", kaydir=True)
    kat_son = 59 + len(KATEGORILER)
    tablo_ekle(ws, "tblKategoriKomisyon", f"A59:E{kat_son}", kat_bas)

    genislik(ws, {"A": 28, "B": 42, "C": 12, "D": 36, "E": 14, "F": 42, "H": 24, "I": 42})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    adimlar = [
        "GIRDI: Barkod, kategori, satış fiyatı ve maliyeti sarı hücrelere girin.",
        "HAKEDIS: TY Plus, flaş, ürün reklamı, desi kargo ve gerçek hakedişi yapıştırın.",
        "AYARLAR: Kategori komisyon tablosunu Trendyol panosundaki oranlarla güncelleyin.",
        "ANAHTAR: Normalize barkod, uzunluk, ikincil ve duplike kolonları otomatik dolar.",
        "MOTOR: Kesinti kırılımı, net kâr, başabaş fiyat ve SAT/ZAM/ÇEKİL hesaplanır.",
        "ESLESTIRME: Beklenen hakediş ↔ gerçek hakediş dört kova denetimi.",
        "KOK_NEDEN: TYK-001..005 sınıflandırır; TANIMSIZ elle incelenir.",
        "PANO / KANIT_RAPORU: Kararı izleyin ve yazdırın.",
    ]
    for i, m in enumerate(adimlar, 5):
        h(ws, i, 1, f"{i - 4}. {m}", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    h(ws, 14, 1, "Sık yapılan hatalar", kalin=True, yazi="B3261E")
    for i, m in enumerate([
        "Barkodu boşluklu bırakmak (normalize edilir ama kontrol edin)",
        "Komisyon oranını satıra yazmak — oran AYARLAR kategori tablosundan gelir",
        "Gerçek hakedişe KDV dahil tutar girmek",
    ], 15):
        h(ws, i, 1, "• " + m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    h(ws, 19, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 26):
        dogrulama(ws, "textLength", "0", f"D{r}",
                  baslik="Kaynak", mesaj="Kaynak metni girin",
                  hata_baslik="Kaynak", hata_mesaj="Boş bırakmayın",
                  isaret="greaterThanOrEqual", f2="120")
    for r in range(60, 70):
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Kategori", mesaj="Kategori adı",
                  hata_baslik="Kategori", hata_mesaj="En az 1 karakter",
                  isaret="greaterThan", f2="40")
        dogrulama(ws, "decimal", "0", f"C{r}",
                  baslik="Komisyon", mesaj="0-1 arası oran",
                  hata_baslik="Oran", hata_mesaj="0 ile 1 arasında",
                  isaret="between", f2="1")
    ws = wb["LISTELER"]
    for r in range(7, 17):
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Kategori", mesaj="Ad girin",
                  hata_baslik="Ad", hata_mesaj="En az 1 karakter",
                  isaret="greaterThan", f2="40")
    for r in range(7, 10):
        dogrulama(ws, "textLength", "1", f"C{r}",
                  baslik="Dönem", mesaj="YYYY-AA",
                  hata_baslik="Dönem", hata_mesaj="Geçersiz",
                  isaret="greaterThan", f2="20")
    ws = wb["KANIT_RAPORU"]
    for r, ad in [(5, "Sürüm"), (22, "Aksiyon1"), (23, "Aksiyon2"), (24, "Aksiyon3")]:
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
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"SAT"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"ZAM"'], fill=sari))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"ÇEKİL"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(6, 27):
        ws.conditional_formatting.add(f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(f"B{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["ESLESTIRME"]
    ws.conditional_formatting.add(f"G{ILK}:G{SON}",
        CellIsRule(operator="equal", formula=['"eslesti"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(f"G{ILK}:G{SON}",
        CellIsRule(operator="equal", formula=['"tutar_farki"'], fill=sari))
    ws.conditional_formatting.add(f"G{ILK}:G{SON}",
        CellIsRule(operator="equal", formula=['"eslesmedi"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"G{ILK}:G{SON}",
        CellIsRule(operator="equal", formula=['"tolerans_icinde"'], fill=sari))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add(f"S{ILK}:S{SON}",
        CellIsRule(operator="equal", formula=['"ZAM"'], fill=sari))
    ws.conditional_formatting.add(f"S{ILK}:S{SON}",
        CellIsRule(operator="equal", formula=['"ÇEKİL"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"S{ILK}:S{SON}",
        CellIsRule(operator="equal", formula=['"SAT"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(f"M{ILK}:M{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["KOK_NEDEN"]
    ws.conditional_formatting.add(f"I16:I{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"TANIMSIZ"'], fill=kirmizi, font=kf))
    for kod in ("TYK-001", "TYK-002", "TYK-003", "TYK-004", "TYK-005"):
        ws.conditional_formatting.add(f"I16:I{15 + KAPASITE}",
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
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"SAT"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"ZAM"'], fill=sari))
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"ÇEKİL"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))

    ws = wb["GIRDI"]
    ws.conditional_formatting.add(f"E{ILK}:E{SON}",
        CellIsRule(operator="greaterThan", formula=["10000"], fill=sari))
    ws.conditional_formatting.add(f"F{ILK}:F{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["HAKEDIS"]
    ws.conditional_formatting.add(f"E{ILK}:E{SON}",
        CellIsRule(operator="greaterThan", formula=["100"], fill=sari))
    ws.conditional_formatting.add(f"D{ILK}:D{SON}",
        CellIsRule(operator="greaterThan", formula=["500"], fill=sari))

    ws = wb["ANAHTAR"]
    ws.conditional_formatting.add(f"F{ILK}:F{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"N{ILK}:N{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))


def adlari_bagla(wb):
    for i, ana in enumerate([
        "toleransKurus", "raporTarihi", "tyk_esikEslesme", "tyk_esikFark",
        "tyk_esikMarjDusuk", "tyk_esikMarjHedef", "tyk_azamiTutar",
        "tyk_ikincilUzunluk", "tyk_olcekHedef", "tyk_yuzdelikOran",
        "tyk_komisyonSenaryo", "tyk_fiyatModeli", "tyk_kvkkBeyani",
        "tyk_basabasFiyatAd", "tyk_normAnahtar", "tyk_tanimsizKova",
        "tyk_kovaButunlukAd", "tyk_dosyaSurumu", "tyk_firmaUnvan", "tyk_kuralSeti",
    ]):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "tyk_mutabakatFarki": 29, "tyk_eslesmeOrani": 30, "tyk_kovaButunluk": 31,
        "tyk_kategoriKomisyon": 32, "tyk_kesintiKirilim": 33, "tyk_basabasFiyat": 34,
        "tyk_kararKapi": 35,
        "tyk_modulT1": 36, "tyk_modulT2": 37, "tyk_modulO1": 38, "tyk_modulO2": 39,
        "tyk_modulO3": 40, "tyk_modulO6": 41, "tyk_modulO8": 42,
        "tyk_modulI1": 43, "tyk_modulI2": 44,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "tyk_modulT1Yorum": 47, "tyk_modulT2Yorum": 48, "tyk_modulO1Yorum": 49,
        "tyk_modulO2Yorum": 50, "tyk_modulO3Yorum": 51, "tyk_modulO6Yorum": 52,
        "tyk_modulO8Yorum": 53, "tyk_modulI1Yorum": 54, "tyk_modulI2Yorum": 55,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeKategoriler", "LISTELER!$A$7:$A$16")
    ad_ekle(wb, "ListeDonemler", "LISTELER!$C$7:$C$9")


def _csv_yaz(yol):
    with open(yol, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            "taraf", "barkod", "stok_kodu", "urun_adi", "kategori", "satis_fiyati",
            "maliyet", "miktar", "hedef_marj", "ty_plus", "flas", "urun_reklam",
            "desi_kargo", "gercek_hakedis", "donem", "not",
        ])
        for s in ORNEK:
            w.writerow([
                "A", s["kod_a"], s["stok"], s["ad"], s["kategori"], s["satis"],
                s["maliyet"], s["miktar"], s["hedef_marj"], "", "", "", "", "", "", "girdi",
            ])
            w.writerow([
                "B", s["kod_b"], "", "", "", "", "", "", "",
                s["ty_plus"], s["flas"], s["reklam"], s["desi_kargo"],
                s["gercek"], s["donem"], "hakedis",
            ])


def main(cikti_yolu=None):
    wb = Workbook()
    siralar = [
        (kapak, "KAPAK"),
        (girdi, "GIRDI"),
        (hakedis, "HAKEDIS"),
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
    dosya_ad = "TrendyolKomisyonSonrasiNetKar.xlsx"
    hedef = cikti_yolu or os.path.join(urun_dir, dosya_ad)
    os.makedirs(os.path.dirname(hedef) or ".", exist_ok=True)
    wb.save(hedef)

    cikti = os.path.join(KOK, "cikti", dosya_ad)
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(hedef) != os.path.abspath(cikti):
        shutil.copy2(hedef, cikti)

    _csv_yaz(os.path.join(urun_dir, "ornek_veri.csv"))

    print(f"Dosya oluşturuldu: {hedef}")
    print(f"Kopya: {cikti}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
