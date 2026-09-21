#!/usr/bin/env python3
"""Elektrik Faturası Doğrulama Motoru — A2 üretim betiği (manda v6 · Hat B Y-02)."""

from __future__ import annotations

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
    itiraz_metni_sablon,
    kok_neden_onerisi,
    normalize_formul,
)

URUN_AD = "Elektrik Faturası Doğrulama Motoru"
SURUM = "1.0.0"
RENK = "0B3D2E"
KAPASITE = 500
DEMO = 35
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

TARIFE_GRUPLARI = ["Mesken", "Ticarethane", "Sanayi", "Tarımsal", "Diğer"]
DURUMLAR = ["Açık", "Kapalı", "İtiraz", "Beklemede"]
KALEM_KODLARI = ["AKTIF", "REAKTIF", "GECE", "GUNDUZ", "PUANT"]


def _ornek_satirlar():
    satirlar = []
    for i in range(1, DEMO + 1):
        no = f"EF202608{i:03d}"
        if i in (3, 7, 11):
            no_a, no_b = no[:10], no
        elif i == 15:
            no_a, no_b = no + "X", no
        else:
            no_a, no_b = no, no
        kwh = 800 + i * 45
        birim = 2.35 + (i % 5) * 0.08
        enerji = round(kwh * birim, 2)
        dagitim = round(kwh * 0.62, 2)
        kdv = round((enerji + dagitim) * 0.20, 2)
        toplam = round(enerji + dagitim + kdv, 2)
        if i % 8 == 0:
            b_tutar = round(toplam - 22.5, 2)
        elif i % 11 == 0:
            b_tutar = round(toplam + 0.03, 2)
        elif i == 15:
            b_tutar = toplam
        else:
            b_tutar = toplam
        tarife = TARIFE_GRUPLARI[i % len(TARIFE_GRUPLARI)]
        kalem = KALEM_KODLARI[i % len(KALEM_KODLARI)]
        tar = RAPOR_TARIHI - timedelta(days=(DEMO - i))
        satirlar.append({
            "no_a": no_a, "no_b": no_b, "tarih": tar, "kwh": kwh,
            "enerji": enerji, "dagitim": dagitim, "kdv": kdv,
            "toplam": toplam, "b_tutar": b_tutar, "tarife": tarife,
            "kalem": kalem, "sayac": f"SYC{10000 + i}",
        })
    return satirlar


ORNEK = _ornek_satirlar()


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Fatura mutabakat sonucunu etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı kova üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Elektrik faturasını tarife, dağıtım ve vergi kalemlerine ayırır; sayaç "
      "tüketimiyle dört kovada mutabakat kurar; kök neden ve itiraz metni üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Normalize fatura anahtarı ile kısa/uzun numara tuzaklarını yakalar",
        "Eşleşti / tolerans içinde / tutar farkı / eşleşmedi — dört kova bütünlüğü",
        "Tarife, dağıtım, KDV, tüketim kök-neden kuralları + TANIMSIZ kova",
        "Tutar farkı satırları için kopyalanabilir dağıtım şirketi itiraz metni",
        "Aylık kanıt raporu — mali işler ve SMMM dosyasına uygun",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Tesis ve enerji yöneticileri",
        "Mali işler / fatura ödeme kontrolü",
        "SMMM ve dönem mutabakat kanıtı arayanlar",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) VERI_A'ya fatura kalemlerini yapıştırın. 2) VERI_B'ye sayaç/beklenen "
      "tutarları girin. 3) PANO'dan eşleşme oranı ve farkı izleyin. "
      "4) AKSIYON'dan itiraz metinlerini kopyalayın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Kod: EFD-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def veri_a(ws):
    sayfa_hazirla(ws, "VERI_A", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Fatura kalemleri — sarı hücrelere yapıştırın / girin", son_kolon=12)
    sutunlar = [
        "FaturaNo", "Tarih", "KalemKodu", "TuketimKwh", "EnerjiBedeli",
        "DagitimBedeli", "VergiKdv", "ToplamTutar", "TarifeGrup", "Donem",
        "Durum", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form_a = {"KayitDolu": 'IF(tblVeriA[[#This Row],[FaturaNo]]="","",1)'}
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["no_a"], baslik="Fatura No", mesaj="Fatura numarası")
        _sari(ws, r, 2, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Fatura tarihi")
        _sari(ws, r, 3, s["kalem"], baslik="Kalem Kodu", mesaj="Liste seçimi")
        _sari(ws, r, 4, s["kwh"], sayi=CATI, baslik="Tüketim kWh", mesaj="Dönem kWh")
        _sari(ws, r, 5, s["enerji"], sayi=TL, baslik="Enerji Bedeli", mesaj="TL")
        _sari(ws, r, 6, s["dagitim"], sayi=TL, baslik="Dağıtım Bedeli", mesaj="TL")
        _sari(ws, r, 7, s["kdv"], sayi=TL, baslik="KDV", mesaj="Vergi tutarı TL")
        _sari(ws, r, 8, s["toplam"], sayi=TL, baslik="Toplam", mesaj="Fatura toplamı TL")
        _sari(ws, r, 9, s["tarife"], baslik="Tarife Grup", mesaj="Liste seçimi")
        _sari(ws, r, 10, "2026-07", baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 11, "Açık", baslik="Durum", mesaj="Liste seçimi")
    for r in range(ILK + DEMO, SON + 1):
        for c in range(1, 12):
            _sari(ws, r, c, None,
                  sayi=(TARİH if c == 2 else (TL if c in (5, 6, 7, 8) else (
                      CATI if c == 4 else None))),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriA", f"A{HDR}:{get_column_letter(12)}{SON}", sutunlar, formuller=form_a)
    for c, tip, f1 in [
        (1, "textLength", "1"), (4, "decimal", "0"), (5, "decimal", "0"),
        (6, "decimal", "0"), (7, "decimal", "0"), (8, "decimal", "0"),
        (10, "textLength", "1"),
    ]:
        dogrulama(ws, tip, f1, f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=sutunlar[c - 1], mesaj="Geçerli değer girin",
                  hata_baslik="Geçersiz", hata_mesaj="Değer kurallara uymuyor",
                  isaret="greaterThanOrEqual" if tip == "decimal" else "greaterThan",
                  f2=None if tip != "textLength" else "40")
    dogrulama(ws, "date", "1", f"B{ILK}:B{SON}",
              baslik="Tarih", mesaj="Tarih seçin", hata_baslik="Tarih",
              hata_mesaj="Geçerli tarih girin", isaret="greaterThan")
    dogrulama(ws, "list", "ListeKalemKodlari", f"C{ILK}:C{SON}",
              baslik="Kalem", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçim yapın")
    dogrulama(ws, "list", "ListeTarifeGruplari", f"I{ILK}:I{SON}",
              baslik="Tarife", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçim yapın")
    dogrulama(ws, "list", "ListeDurumlar", f"K{ILK}:K{SON}",
              baslik="Durum", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçim yapın")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 13)})


def veri_b(ws):
    sayfa_hazirla(ws, "VERI_B", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Sayaç okuma / beklenen tutar tarafı", son_kolon=8)
    sutunlar = [
        "FaturaNo", "Tarih", "TuketimKwh", "BeklenenTutar", "SayacNo",
        "Kaynak", "Donem", "Durum", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form_b = {"KayitDolu": 'IF(tblVeriB[[#This Row],[FaturaNo]]="","",1)'}
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["no_b"], baslik="Fatura No", mesaj="Mutabakat anahtarı")
        _sari(ws, r, 2, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Okuma tarihi")
        _sari(ws, r, 3, s["kwh"], sayi=CATI, baslik="Tüketim kWh", mesaj="Sayaç kWh")
        _sari(ws, r, 4, s["b_tutar"], sayi=TL, baslik="Beklenen Tutar", mesaj="TL")
        _sari(ws, r, 5, s["sayac"], baslik="Sayaç No", mesaj="Sayaç kimliği")
        _sari(ws, r, 6, "Sayaç", baslik="Kaynak", mesaj="Kayıt kaynağı")
        _sari(ws, r, 7, "2026-07", baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 8, "Açık", baslik="Durum", mesaj="Liste seçimi")
    for r in range(ILK + DEMO, SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None,
                  sayi=(TARİH if c == 2 else (CATI if c == 3 else (TL if c == 4 else None))),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriB", f"A{HDR}:I{SON}", sutunlar, formuller=form_b)
    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}", baslik="Fatura No",
              mesaj="Fatura numarası girin", hata_baslik="Eksik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "date", "1", f"B{ILK}:B{SON}", baslik="Tarih", mesaj="Tarih",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih", isaret="greaterThan")
    dogrulama(ws, "decimal", "0", f"C{ILK}:C{SON}", baslik="kWh", mesaj="kWh ≥ 0",
              hata_baslik="kWh", hata_mesaj="0 veya üzeri", isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"D{ILK}:D{SON}", baslik="Tutar", mesaj="Tutar ≥ 0",
              hata_baslik="Tutar", hata_mesaj="0 veya üzeri", isaret="greaterThanOrEqual")
    for c, ad in [(5, "Sayaç"), (6, "Kaynak"), (7, "Dönem")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin", hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "list", "ListeDurumlar", f"H{ILK}:H{SON}", baslik="Durum",
              mesaj="Listeden seçin", hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 16 for i in range(1, 10)})


def anahtar(ws):
    sayfa_hazirla(ws, "ANAHTAR", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Normalize fatura anahtarı + uzunluk + ikincil + duplike (Ç7 / E01 / E05)",
             son_kolon=12)
    h(ws, 4, 1, "A Tarafı (Fatura)", kalin=True, yazi=KOYU_LACIVERT)
    sut_a = ["Ham", "Norm", "NormUzunluk", "IkincilAnahtar", "KaynakTaraf", "Duplike", "SatirNo"]
    baslik_satiri(ws, HDR, sut_a)
    form_a = {
        "Norm": normalize_formul("A6").replace("A6", "tblAnahtarA[[#This Row],[Ham]]").lstrip("="),
        "NormUzunluk": 'IF(tblAnahtarA[[#This Row],[Norm]]="",0,LEN(tblAnahtarA[[#This Row],[Norm]]))',
        "IkincilAnahtar": 'IF(tblAnahtarA[[#This Row],[Norm]]="",""'
                         ',LEFT(tblAnahtarA[[#This Row],[Norm]],efd_ikincilUzunluk))',
        "KaynakTaraf": '"A"',
        "Duplike": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",'
                   'IF(COUNTIF(tblAnahtarA[Norm],tblAnahtarA[[#This Row],[Norm]])>1,"EVET","HAYIR"))',
        "SatirNo": 'IF(tblAnahtarA[[#This Row],[Ham]]="","",ROW()-ROW($A$5))',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 1, f"=IF(VERI_A!A{r}=\"\",\"\",VERI_A!A{r})")
    tablo_ekle(ws, "tblAnahtarA", f"A{HDR}:G{SON}", sut_a, formuller=form_a)

    h(ws, 4, 9, "B Tarafı (Sayaç)", kalin=True, yazi=KOYU_LACIVERT)
    sut_b = ["HamB", "NormB", "NormUzunlukB", "IkincilB", "KaynakB", "DuplikeB", "SatirNoB"]
    baslik_satiri(ws, HDR, sut_b, basla=9)
    form_b = {
        "NormB": 'UPPER(TRIM(SUBSTITUTE(tblAnahtarB[[#This Row],[HamB]]," ","")))',
        "NormUzunlukB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",0,LEN(tblAnahtarB[[#This Row],[NormB]]))',
        "IkincilB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",""'
                    ',LEFT(tblAnahtarB[[#This Row],[NormB]],efd_ikincilUzunluk))',
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


def eslestirme(ws):
    sayfa_hazirla(ws, "ESLESTIRME", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Dört kova motoru — fatura↔sayaç (E02)", son_kolon=14)
    sutunlar = [
        "A_Norm", "B_Sayim", "B_Tutar", "A_Tutar", "Fark", "AbsFark",
        "Kova", "IkincilEslesme", "DuplikeBayrak", "FarkOran",
        "YuvarlakA", "YuvarlakB", "ToleransTest", "EsikUstu",
        "NetKatki", "MotorAdim",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    formuller = {
        "A_Norm": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",tblAnahtarA[[#This Row],[Norm]])',
        "B_Sayim": 'IF([A_Norm]="","",COUNTIF(tblAnahtarB[NormB],[A_Norm]))',
        "B_Tutar": 'IF([A_Norm]="","",SUMIF(tblAnahtarB[NormB],[A_Norm],tblVeriB[BeklenenTutar]))',
        "A_Tutar": 'IF([A_Norm]="","",SUMIF(tblAnahtarA[Norm],[A_Norm],tblVeriA[ToplamTutar]))',
        "Fark": 'IF([A_Norm]="","",ROUND([A_Tutar],2)-ROUND([B_Tutar],2))',
        "AbsFark": 'IF([A_Norm]="","",ABS([Fark]))',
        "Kova": (
            'IF([A_Norm]="","",'
            'IF([B_Sayim]=0,"eslesmedi",'
            'IF(ROUND([A_Tutar],2)=ROUND([B_Tutar],2),"eslesti",'
            'IF(ABS(ROUND([A_Tutar],2)-ROUND([B_Tutar],2))<=toleransKurus,'
            '"tolerans_icinde","tutar_farki"))))'
        ),
        "IkincilEslesme": (
            'IF([A_Norm]="","",'
            'IF([B_Sayim]>0,"BIRINCIL",'
            'IF(COUNTIF(tblAnahtarB[IkincilB],tblAnahtarA[[#This Row],[IkincilAnahtar]])>0,'
            '"IKINCIL","YOK")))'
        ),
        "DuplikeBayrak": 'IF([A_Norm]="","",tblAnahtarA[[#This Row],[Duplike]])',
        "FarkOran": 'IF(OR([A_Norm]="",[A_Tutar]=0),0,[AbsFark]/ABS([A_Tutar]))',
        "YuvarlakA": 'IF([A_Norm]="","",ROUND([A_Tutar],2))',
        "YuvarlakB": 'IF([A_Norm]="","",ROUND([B_Tutar],2))',
        "ToleransTest": 'IF([A_Norm]="","",IF([AbsFark]<=toleransKurus,1,0))',
        "EsikUstu": 'IF([A_Norm]="","",IF([AbsFark]>efd_esikFark,1,0))',
        "NetKatki": 'IF([A_Norm]="","",IF([Kova]="eslesti",[A_Tutar],0))',
        "MotorAdim": 'IF([A_Norm]="","",[B_Sayim]+[ToleransTest]+[EsikUstu])',
    }
    form_yazi = {}
    for k, v in formuller.items():
        vv = v
        for kol in sutunlar:
            vv = vv.replace(f"[{kol}]", f"tblEslesme[[#This Row],[{kol}]]")
        form_yazi[k] = vv
    tablo_ekle(ws, "tblEslesme", f"A{HDR}:P{SON}", sutunlar, formuller=form_yazi)
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
    alt_bant(ws, 3, "Öncelik kurallı kök neden — eşleşmeyen fark TANIMSIZ (E04)", son_kolon=10)
    h(ws, 4, 1, "Kural Önceliği", kalin=True, yazi=KOYU_LACIVERT)
    kurallar = kok_neden_onerisi([
        {"neden_kodu": "TARIFE", "test": "ABS(fark-enerji)<=toleransKurus",
         "aciklama": "Fark enerji/tarife bedeline yakın"},
        {"neden_kodu": "DAGITIM", "test": "ABS(fark-dagitim)<=toleransKurus",
         "aciklama": "Fark dağıtım bedeline yakın"},
        {"neden_kodu": "KDV", "test": "ABS(fark-kdv)<=toleransKurus",
         "aciklama": "Fark KDV tutarına yakın"},
        {"neden_kodu": "TUKETIM", "test": "ABS(kwh_a-kwh_b)>0",
         "aciklama": "Sayaç tüketim uyuşmazlığı"},
        {"neden_kodu": "SAYAC", "test": "sayac_anahtar_farki>0",
         "aciklama": "Sayaç kimliği / dönem şüphesi"},
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
    sut = ["A_Norm", "Kova", "AbsFark", "Enerji", "Dagitim", "Kdv", "Tuketim", "KokNeden"]
    baslik_satiri(ws, 15, sut)
    form = {
        "A_Norm": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",tblEslesme[[#This Row],[A_Norm]])',
        "Kova": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                'IFERROR(INDEX(tblEslesme[Kova],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),""))',
        "AbsFark": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                   'IFERROR(INDEX(tblEslesme[AbsFark],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),0))',
        "Enerji": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                  'IFERROR(SUMIF(tblAnahtarA[Norm],tblKok[[#This Row],[A_Norm]],tblVeriA[EnerjiBedeli]),0))',
        "Dagitim": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                   'IFERROR(SUMIF(tblAnahtarA[Norm],tblKok[[#This Row],[A_Norm]],tblVeriA[DagitimBedeli]),0))',
        "Kdv": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
               'IFERROR(SUMIF(tblAnahtarA[Norm],tblKok[[#This Row],[A_Norm]],tblVeriA[VergiKdv]),0))',
        "Tuketim": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                   'IFERROR(SUMIF(tblAnahtarA[Norm],tblKok[[#This Row],[A_Norm]],tblVeriA[TuketimKwh]),0))',
        "KokNeden": (
            'IF(OR(tblKok[[#This Row],[A_Norm]]="",tblKok[[#This Row],[Kova]]<>"tutar_farki"),"",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[Enerji]])<=toleransKurus,"TARIFE",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[Dagitim]])<=toleransKurus,"DAGITIM",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[Kdv]])<=toleransKurus,"KDV",'
            'IF(tblKok[[#This Row],[Tuketim]]>0,"TUKETIM","TANIMSIZ")))))'
        ),
    }
    for i in range(KAPASITE):
        r = 16 + i
        h(ws, r, 1, f"=IF(ESLESTIRME!A{ILK + i}=\"\",\"\",ESLESTIRME!A{ILK + i})")
    tablo_ekle(ws, "tblKok", f"A15:H{15 + KAPASITE}", sut, formuller=form)
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 9)})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=20)
    sabitle(ws, "A3")
    h(ws, 3, 1, "Karar Destek Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 3, 3, "Rapor Tarihi", yazi=GRİ)
    h(ws, 3, 4, "=raporTarihi", sayi=TARİH)

    h(ws, 4, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2,
      '=IF(COUNTA(tblVeriA[FaturaNo])=0,"VERİ YOK",'
      'IF(OR(COUNTIF(tblVeriA[ToplamTutar],"<0")>0,COUNTIF(tblVeriB[BeklenenTutar],"<0")>0),"İNCELE",'
      'IF(OR(COUNTIF(tblVeriA[ToplamTutar],">"&efd_azamiTutar)>0,'
      'COUNTIF(tblVeriB[BeklenenTutar],">"&efd_azamiTutar)>0),"İNCELE",'
      'IF(efd_eslesmeOrani<efd_esikEslesme,"İNCELE",'
      'IF(ABS(efd_mutabakatFarki)>efd_esikFark,"İNCELE","UYGUN")))))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Eşleşme Oranı", '=IFERROR(efd_eslesmeOrani,0)', YÜZDE),
        (7, "Fatura Toplam (A)", '=IFERROR(SUM(tblVeriA[ToplamTutar]),0)', TL),
        (8, "Beklenen (B)", '=IFERROR(SUM(tblVeriB[BeklenenTutar]),0)', TL),
        (9, "Mutabakat Farkı", '=IFERROR(efd_mutabakatFarki,0)', TL),
        (10, "Eşleşti Adet", '=COUNTIF(tblEslesme[Kova],"eslesti")', CATI),
        (11, "Tolerans İçinde", '=COUNTIF(tblEslesme[Kova],"tolerans_icinde")', CATI),
        (12, "Tutar Farkı Adet", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (13, "Eşleşmedi Adet", '=COUNTIF(tblEslesme[Kova],"eslesmedi")', CATI),
        (14, "Kova Bütünlük", "=efd_kovaButunluk", None),
        (15, "Anomali Sayısı", "=efd_anomaliSayisi", CATI),
        (16, "Pareto Fark Payı", "=efd_paretoFark", YÜZDE),
        (17, "Tahmin Aralığı", "=efd_tahminAralik", TL),
        (18, "Tornado Zirve", "=efd_tornadoZirve", TL),
        (19, "Kök-Neden Çeşit", "=efd_kokNedenDagilim", CATI),
        (20, "Duplike Uyarı", '=COUNTIF(tblAnahtarA[Duplike],"EVET")', CATI),
        (21, "Veri Hazırlık",
         '=IF(COUNTA(tblVeriA[FaturaNo])=0,0,'
         'COUNTA(tblVeriA[FaturaNo])/(COUNTA(tblVeriA[FaturaNo])+COUNTA(tblVeriB[FaturaNo])))',
         YÜZDE),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (7, '=IF(COUNTA(tblVeriA[FaturaNo])=0,"Veri yok",'
            '_xlfn.TEXTJOIN(" ",TRUE,"Mutabakat farkı",TEXT(efd_mutabakatFarki,"0.00"),"TL"))'),
        (8, '=IF(efd_kokNedenDagilim=0,"Kök neden yok",'
            '_xlfn.TEXTJOIN(" ",TRUE,efd_kokNedenDagilim,"farklı kök neden sınıfı"))'),
        (9, '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"En büyük fark payı",TEXT(efd_paretoFark,"0.0%")),"-")'),
        (10, '=IF(efd_anomaliSayisi=0,"Anomali yok",'
             '_xlfn.TEXTJOIN(" ",TRUE,efd_anomaliSayisi,"anomali satırı"))'),
        (11, '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Zirve etki",TEXT(efd_tornadoZirve,"0.00"),"TL"),"-")'),
        (12, '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tahmini fatura",TEXT(efd_tahminAralik,"0.00"),"TL"),"-")'),
    ]
    for r, form in yorumlar:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 23, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 1, "Kova")
    h(ws, 24, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Eşleşti", 24), ("Tolerans", 3), ("Tutar farkı", 5), ("Eşleşmedi", 3),
    ], 25):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 24, 4, "Neden")
    h(ws, 24, 5, "Adet")
    for i, (ad, adet) in enumerate([
        ("Tarife", 2), ("Dağıtım", 1), ("KDV", 1), ("Tüketim", 1),
        ("Sayaç", 0), ("Tanımsız", 3),
    ], 25):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 24, 7, "Gün")
    h(ws, 24, 8, "Fark")
    for i in range(7):
        h(ws, 25 + i, 7, i + 1, sayi=CATI)
        h(ws, 25 + i, 8, round(60 + i * 14.5, 2), sayi=TL)

    h(ws, 24, 10, "Senaryo")
    h(ws, 24, 11, "Etki")
    for i, (ad, v) in enumerate([
        ("Tarife +%5", 380), ("Dağıtım +%3", 160), ("KDV düzeltme", 140),
        ("kWh sapma", 95), ("Tolerans sıkı", 55),
    ], 25):
        h(ws, i, 10, ad)
        h(ws, i, 11, v, sayi=TL)

    pie1 = PieChart()
    pie1.title = "Kova Dağılımı"
    pie1.add_data(Reference(ws, min_col=2, min_row=24, max_row=28), titles_from_data=True)
    pie1.set_categories(Reference(ws, min_col=1, min_row=25, max_row=28))
    pie1.width, pie1.height = 10, 7
    ws.add_chart(pie1, "A33")

    bar1 = BarChart()
    bar1.type = "col"
    bar1.title = "Kök Neden"
    bar1.add_data(Reference(ws, min_col=5, min_row=24, max_row=30), titles_from_data=True)
    bar1.set_categories(Reference(ws, min_col=4, min_row=25, max_row=30))
    bar1.width, bar1.height = 10, 7
    ws.add_chart(bar1, "F33")

    line1 = LineChart()
    line1.title = "Günlük Fark"
    line1.add_data(Reference(ws, min_col=8, min_row=24, max_row=31), titles_from_data=True)
    line1.set_categories(Reference(ws, min_col=7, min_row=25, max_row=31))
    line1.width, line1.height = 10, 7
    ws.add_chart(line1, "A48")

    bar2 = BarChart()
    bar2.type = "bar"
    bar2.title = "Tornado Etki"
    bar2.add_data(Reference(ws, min_col=11, min_row=24, max_row=29), titles_from_data=True)
    bar2.set_categories(Reference(ws, min_col=10, min_row=25, max_row=29))
    bar2.width, bar2.height = 10, 7
    ws.add_chart(bar2, "F48")

    pie2 = PieChart()
    pie2.title = "Eşleşme Payı"
    pie2.add_data(Reference(ws, min_col=2, min_row=24, max_row=26), titles_from_data=True)
    pie2.set_categories(Reference(ws, min_col=1, min_row=25, max_row=26))
    pie2.width, pie2.height = 9, 6
    ws.add_chart(pie2, "A63")

    bar3 = BarChart()
    bar3.title = "Üst Nedenler"
    bar3.add_data(Reference(ws, min_col=5, min_row=24, max_row=27), titles_from_data=True)
    bar3.set_categories(Reference(ws, min_col=4, min_row=25, max_row=27))
    bar3.width, bar3.height = 9, 6
    ws.add_chart(bar3, "F63")

    line2 = LineChart()
    line2.title = "Fark Eğilimi"
    line2.add_data(Reference(ws, min_col=8, min_row=24, max_row=28), titles_from_data=True)
    line2.set_categories(Reference(ws, min_col=7, min_row=25, max_row=28))
    line2.width, line2.height = 9, 6
    ws.add_chart(line2, "A76")

    bar4 = BarChart()
    bar4.title = "Senaryo Karşılaştırma"
    bar4.add_data(Reference(ws, min_col=11, min_row=24, max_row=27), titles_from_data=True)
    bar4.set_categories(Reference(ws, min_col=10, min_row=25, max_row=27))
    bar4.width, bar4.height = 9, 6
    ws.add_chart(bar4, "F76")

    baski_hazirla(ws, "A1:L32", f"{URUN_AD} | {SURUM}")
    genislik(ws, {get_column_letter(i): 16 for i in range(1, 13)})
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["D"].width = 18


def aksiyon(ws):
    sayfa_hazirla(ws, "AKSIYON", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Tutar farkı satırları için dağıtım şirketi itiraz metni", son_kolon=10)
    sut = ["A_Norm", "Kova", "Fark", "KokNeden", "Oncelik", "ItirazMetni", "Durum"]
    baslik_satiri(ws, HDR, sut)
    form = {
        "A_Norm": 'IF(tblKok[[#This Row],[A_Norm]]="","",tblKok[[#This Row],[A_Norm]])',
        "Kova": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                'IFERROR(INDEX(tblKok[Kova],MATCH(tblAksiyon[[#This Row],[A_Norm]],tblKok[A_Norm],0)),""))',
        "Fark": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                'IFERROR(INDEX(tblEslesme[Fark],MATCH(tblAksiyon[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),0))',
        "KokNeden": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                    'IFERROR(INDEX(tblKok[KokNeden],MATCH(tblAksiyon[[#This Row],[A_Norm]],tblKok[A_Norm],0)),""))',
        "Oncelik": (
            'IF(tblAksiyon[[#This Row],[Kova]]<>"tutar_farki","",'
            'IF(ABS(tblAksiyon[[#This Row],[Fark]])>efd_esikFark,"YUKSEK","NORMAL"))'
        ),
        "ItirazMetni": (
            'IF(tblAksiyon[[#This Row],[Kova]]<>"tutar_farki","",'
            '"Fatura "&tblAksiyon[[#This Row],[A_Norm]]&" için fark "&'
            'TEXT(tblAksiyon[[#This Row],[Fark]],"0.00")&" TL. Dayanak: "&'
            'tblAksiyon[[#This Row],[KokNeden]]&". İtiraz dosyası üretildi.")'
        ),
        "Durum": 'IF(tblAksiyon[[#This Row],[Kova]]<>"tutar_farki","","Beklemede")',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 1, f"=IF(KOK_NEDEN!A{16 + i}=\"\",\"\",KOK_NEDEN!A{16 + i})")
    tablo_ekle(ws, "tblAksiyon", f"A{HDR}:G{SON}", sut, formuller=form)
    h(ws, 2, 9, "Şablon", yazi=GRİ)
    h(ws, 2, 10, itiraz_metni_sablon().replace("=", "'"), yazi=GRİ, boyut=8, kaydir=True)
    sabitle(ws, f"A{ILK}")
    genislik(ws, {"A": 16, "B": 14, "C": 12, "D": 12, "E": 10, "F": 55, "G": 12})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Dönem Elektrik Faturası Mutabakat Kanıtı", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Eşleşme Oranı", "=efd_eslesmeOrani", YÜZDE),
        (10, "Fatura Toplam", "=SUM(tblVeriA[ToplamTutar])", TL),
        (11, "Mutabakat Farkı", "=efd_mutabakatFarki", TL),
        (12, "Eşleşti", '=COUNTIF(tblEslesme[Kova],"eslesti")', CATI),
        (13, "Tutar Farkı", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (14, "Eşleşmedi", '=COUNTIF(tblEslesme[Kova],"eslesmedi")', CATI),
        (15, "Kova Bütünlük", "=efd_kovaButunluk", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 17, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 1,
      "Tolerans kuruşu AYARLAR'dan gelir. Anahtar UPPER/TRIM/SUBSTITUTE ile normalize edilir. "
      "Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A18:H18")
    h(ws, 20, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 1, "1. Tutar farkı satırlarında AKSIYON itiraz metnini kullanın.")
    h(ws, 22, 1, "2. TANIMSIZ kök nedenleri elle inceleyin.")
    h(ws, 23, 1, "3. Duplike fatura anahtarlarını temizleyip yeniden eşleştirin.")
    baski_hazirla(ws, "A1:H24", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 28, "B": 18, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "A satır sayısı", '=COUNTA(tblVeriA[FaturaNo])'),
        (6, "B satır sayısı", '=COUNTA(tblVeriB[FaturaNo])'),
        (7, "Eşleşti", '=COUNTIF(tblEslesme[Kova],"eslesti")'),
        (8, "Tolerans", '=COUNTIF(tblEslesme[Kova],"tolerans_icinde")'),
        (9, "Tutar farkı", '=COUNTIF(tblEslesme[Kova],"tutar_farki")'),
        (10, "Eşleşmedi", '=COUNTIF(tblEslesme[Kova],"eslesmedi")'),
        (11, "Kova toplam", "=B7+B8+B9+B10"),
        (12, "Bütünlük", '=IF(B11=B5,"BUTUN","KAYIP")'),
        (13, "Duplike A", '=COUNTIF(tblAnahtarA[Duplike],"EVET")'),
        (14, "TANIMSIZ", '=COUNTIF(tblKok[KokNeden],"TANIMSIZ")'),
        (15, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (16, "Kalite skoru",
         '=IF(B5=0,0,ROUND(100*(B7+B8)/MAX(B5,1),0))'),
        (17, "Net A", "=SUM(tblVeriA[ToplamTutar])"),
        (18, "Net B", "=SUM(tblVeriB[BeklenenTutar])"),
        (19, "Fark kontrol", "=B17-B18"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, kalin=True, sayi=TL if r >= 17 else (CATI if r < 15 else None))
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO} satır örnek fatura↔sayaç mutabakatı vardır: eşleşen, tolerans içi, "
      "tutar farkı ve eşleşmeyen kovalar kasıtlı üretilmiştir. Kısa/uzun fatura no senaryosu "
      "satır 3, 7, 11'dedir. VERI_A ve VERI_B'yi temizleyip kendi faturalarınızı yapıştırın.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Beklenen eşleşen ~24 | tolerans ~3 | tutar farkı ~5 | eşleşmedi ~3")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 50000, sayi=CATI)
    genislik(ws, {"A": 70, "B": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Tarife Grupları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "Tarife", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, p in enumerate(TARIFE_GRUPLARI, 7):
        _sari(ws, i, 1, p, baslik="Tarife", mesaj="Listeye ekleyebilirsiniz")
    h(ws, 3, 3, "Durum Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 3, "Durum", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, d in enumerate(DURUMLAR, 7):
        _sari(ws, i, 3, d, baslik="Durum", mesaj="Durum değeri")
    h(ws, 3, 5, "Kalem Kodları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 5, "Kalem", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, k in enumerate(KALEM_KODLARI, 7):
        _sari(ws, i, 5, k, baslik="Kalem", mesaj="Kalem kodu")
    h(ws, 3, 7, "EvetHayir", kalin=True)
    h(ws, 6, 7, "Secim", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    _sari(ws, 7, 7, "EVET", baslik="Seçim", mesaj="EVET/HAYIR")
    _sari(ws, 8, 7, "HAYIR", baslik="Seçim", mesaj="EVET/HAYIR")
    genislik(ws, {"A": 18, "C": 16, "E": 14, "G": 12})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Tek kaynak parametreler — kaynak ve yürürlük zorunlu (D10)", son_kolon=10)
    baslik_satiri(ws, 5, ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"])
    params = [
        ("toleransKurus", 0.05, "TL", "İş kuralı — kuruş toleransı", "01.01.2026", "Eşleşme toleransı"),
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026", "Rapor tarihi (TODAY yok)"),
        ("efd_esikEslesme", 0.90, "oran", "İç politika", "01.01.2026", "Eşleşme oranı eşiği"),
        ("efd_esikFark", 100, "TL", "İç politika", "01.01.2026", "Fark tutarı eşiği"),
        ("efd_azamiTutar", 1000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul tutar"),
        ("efd_ikincilUzunluk", 12, "karakter", "Anahtar stratejisi SPEC", "01.01.2026", "Ikincil anahtar uzunluğu"),
        ("efd_olcekHedef", 50000, "satir", "Manda A2 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("efd_anomaliZ", 2.0, "z", "İstatistik kuralı", "01.01.2026", "Anomali Z eşiği"),
        ("efd_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("efd_kdvOran", 0.20, "oran", "KDV genel oran", "01.01.2026", "Varsayılan KDV"),
        ("efd_dagitimBirim", 0.62, "TL/kWh", "Dağıtım tarife ort.", "01.01.2026", "Varsayılan dağıtım"),
        ("efd_fiyatModeli", "Tek seferlik alım — abonelik yok", "metin", "Ürün konumlandırma", "01.01.2026", "R3 ayrım"),
        ("efd_kvkkBeyani", "Fatura verisi cihazdan çıkmaz", "metin", "KVKK beyanı", "01.01.2026", "R3 ayrım"),
        ("efd_kalemAyristirma", "Tarife+dağıtım+KDV tek motor", "metin", "Alan bilgisi", "01.01.2026", "R1 ayrım"),
        ("efd_kovaButunlukMetin", "Dört kova bütünlüğü denetimli", "metin", "E02", "01.01.2026", "R1 ayrım"),
        ("efd_itirazMetni", "İtiraz metni üretir", "metin", "Aksiyon motoru", "01.01.2026", "R1 ayrım"),
        ("efd_normAnahtar", "UPPER TRIM SUBSTITUTE", "metin", "Ç7 anahtar", "01.01.2026", "Serbest alternatif"),
        ("efd_tanimsizKova", "TANIMSIZ görünür", "metin", "E04", "01.01.2026", "Serbest alternatif"),
        ("efd_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("efd_firmaUnvan", "Örnek Enerji Tesis A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026", "Firma"),
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
        h(ws, r, 4, kay)
        h(ws, r, 5, yur)
        h(ws, r, 6, acik, kaydir=True)

    h(ws, 26, 8, "Motor Cikti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 9, "Deger", kalin=True, yazi=KOYU_LACIVERT)
    motor = [
        (27, "efd_mutabakatFarki",
         '=IFERROR(SUM(tblVeriA[ToplamTutar])-SUM(tblVeriB[BeklenenTutar]),0)'),
        (28, "efd_eslesmeOrani",
         '=IF(COUNTA(tblEslesme[A_Norm])=0,0,'
         '(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/COUNTA(tblEslesme[A_Norm]))'),
        (29, "efd_kovaButunluk", "=ESLESTIRME!C2"),
        (30, "efd_kokNedenDagilim",
         '=IFERROR(SUMPRODUCT((tblKok[KokNeden]<>"")*1),0)'),
        (31, "efd_paretoFark",
         '=IFERROR(IF(SUM(tblEslesme[AbsFark])=0,0,MAX(tblEslesme[AbsFark])/SUM(tblEslesme[AbsFark])),0)'),
        (32, "efd_anomaliSayisi",
         '=COUNTIF(tblEslesme[EsikUstu],1)'),
        (33, "efd_tahminAralik",
         '=IFERROR(PERCENTILE(tblVeriA[ToplamTutar],efd_yuzdelikOran),0)'),
        (34, "efd_tornadoZirve",
         '=IFERROR(MAX(PANO!K25:K29),0)'),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kalin=True,
          sayi=YÜZDE if ("Orani" in ad or "pareto" in ad) else TL)

    h(ws, 35, 8, "Canli Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorum_satir = [
        (36, "efd_yorumFark",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Fark",TEXT(efd_mutabakatFarki,"0.00"),"TL"),"-")'),
        (37, "efd_yorumKokNeden",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Kok neden sinifi",TEXT(efd_kokNedenDagilim,"0")),"-")'),
        (38, "efd_yorumPareto",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Pareto",TEXT(efd_paretoFark,"0.0%")),"-")'),
        (39, "efd_yorumAnomali",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Anomali",TEXT(efd_anomaliSayisi,"0")),"-")'),
        (40, "efd_yorumTornado",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tornado",TEXT(efd_tornadoZirve,"0.00")),"-")'),
        (41, "efd_yorumTahmin",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tahmin",TEXT(efd_tahminAralik,"0.00")),"-")'),
    ]
    for r, ad, form in yorum_satir:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 28, "B": 42, "C": 12, "D": 28, "E": 14, "F": 28, "H": 24, "I": 42})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    adimlar = [
        "VERI_A: Elektrik faturası kalemlerini sarı hücrelere yapıştırın.",
        "VERI_B: Sayaç okuma / beklenen tutarları girin.",
        "ANAHTAR: Normalize, uzunluk, ikincil ve duplike kolonları otomatik dolar.",
        "ESLESTIRME: Dört kova ve bütünlük denetimi otomatik çalışır.",
        "KOK_NEDEN: Tutar farklarını tarife/dağıtım/KDV/tüketim sınıflandırır.",
        "AKSIYON: İtiraz metinlerini kopyalayıp dağıtım şirketine iletin.",
        "KANIT_RAPORU: Dönem dosyasına PDF olarak yazdırın.",
    ]
    for i, m in enumerate(adimlar, 5):
        h(ws, i, 1, f"{i - 4}. {m}", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    h(ws, 13, 1, "Sık yapılan hatalar", kalin=True, yazi="B3261E")
    for i, m in enumerate([
        "Fatura numarası boşluklu bırakmak (normalize edilir ama kontrol edin)",
        "A ve B dönemlerini karıştırmak",
        "Beklenen tutara KDV hariç yazıp fatura toplamıyla karşılaştırmak",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 24):
        dogrulama(ws, "textLength", "0", f"D{r}",
                  baslik="Kaynak", mesaj="Kaynak metni girin",
                  hata_baslik="Kaynak", hata_mesaj="Boş bırakmayın",
                  isaret="greaterThanOrEqual", f2="120")
    ws = wb["LISTELER"]
    for r in range(7, 12):
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Tarife", mesaj="Ad girin",
                  hata_baslik="Ad", hata_mesaj="En az 1 karakter",
                  isaret="greaterThan", f2="40")
    for r in range(7, 11):
        dogrulama(ws, "textLength", "1", f"C{r}",
                  baslik="Durum", mesaj="Durum girin",
                  hata_baslik="Durum", hata_mesaj="Geçersiz",
                  isaret="greaterThan", f2="30")
    for r in range(7, 12):
        dogrulama(ws, "textLength", "1", f"E{r}",
                  baslik="Kalem", mesaj="Kalem kodu",
                  hata_baslik="Kalem", hata_mesaj="Geçersiz",
                  isaret="greaterThan", f2="20")
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
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"İNCELE"'], fill=sari))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(6, 22):
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

    ws = wb["KOK_NEDEN"]
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"TANIMSIZ"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"TARIFE"'], fill=sari))
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"DAGITIM"'], fill=sari))
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"KDV"'], fill=sari))
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"TUKETIM"'], fill=sari))

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
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"İNCELE"'], fill=sari))
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))

    ws = wb["VERI_A"]
    ws.conditional_formatting.add(f"H{ILK}:H{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"E{ILK}:E{SON}",
        CellIsRule(operator="greaterThan", formula=["50000"], fill=sari))

    ws = wb["VERI_B"]
    ws.conditional_formatting.add(f"D{ILK}:D{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"D{ILK}:D{SON}",
        CellIsRule(operator="greaterThan", formula=["50000"], fill=sari))

    ws = wb["ANAHTAR"]
    ws.conditional_formatting.add(f"F{ILK}:F{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"N{ILK}:N{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))


def adlari_bagla(wb):
    for i, ana in enumerate([
        "toleransKurus", "raporTarihi", "efd_esikEslesme", "efd_esikFark",
        "efd_azamiTutar", "efd_ikincilUzunluk", "efd_olcekHedef", "efd_anomaliZ",
        "efd_yuzdelikOran", "efd_kdvOran", "efd_dagitimBirim", "efd_fiyatModeli",
        "efd_kvkkBeyani", "efd_kalemAyristirma", "efd_kovaButunlukMetin",
        "efd_itirazMetni", "efd_normAnahtar", "efd_tanimsizKova",
        "efd_dosyaSurumu", "efd_firmaUnvan",
    ]):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "efd_mutabakatFarki": 27, "efd_eslesmeOrani": 28, "efd_kovaButunluk": 29,
        "efd_kokNedenDagilim": 30, "efd_paretoFark": 31, "efd_anomaliSayisi": 32,
        "efd_tahminAralik": 33, "efd_tornadoZirve": 34,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "efd_yorumFark": 36, "efd_yorumKokNeden": 37, "efd_yorumPareto": 38,
        "efd_yorumAnomali": 39, "efd_yorumTornado": 40, "efd_yorumTahmin": 41,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeTarifeGruplari", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeDurumlar", "LISTELER!$C$7:$C$10")
    ad_ekle(wb, "ListeKalemKodlari", "LISTELER!$E$7:$E$11")


def main(cikti_yolu=None):
    wb = Workbook()
    siralar = [
        (kapak, "KAPAK"),
        (veri_a, "VERI_A"),
        (veri_b, "VERI_B"),
        (anahtar, "ANAHTAR"),
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
    dosya_ad = "ElektrikFaturasiDogrulama.xlsx"
    hedef = cikti_yolu or os.path.join(urun_dir, dosya_ad)
    os.makedirs(os.path.dirname(hedef) or ".", exist_ok=True)
    wb.save(hedef)

    cikti = os.path.join(KOK, "cikti", dosya_ad)
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(hedef) != os.path.abspath(cikti):
        shutil.copy2(hedef, cikti)

    print(f"Dosya oluşturuldu: {hedef}")
    print(f"Kopya: {cikti}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
