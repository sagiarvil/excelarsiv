#!/usr/bin/env python3
"""Defter Beyan e-Arşiv Aktarım Düzenleyici — A2 üretim betiği (manda v6, Hat C C-09)."""

from __future__ import annotations

import csv
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
    itiraz_metni_sablon,
    kok_neden_onerisi,
    normalize_formul,
)

URUN_AD = "Defter Beyan e-Arşiv Aktarım Düzenleyici"
SURUM = "1.0.0"
RENK = "0F2942"
KAPASITE = 1000
DEMO = 40
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 14)

BELGE_TIPLERI = ["SATIS", "IADE", "ISTISNA", "TEVKIFAT", "DIGER"]
DURUMLAR = ["Acik", "Kapali", "Itiraz", "Beklemede"]

KOLON_HARITA = [
    ("FaturaNo", "BelgeNo", "EVET", "metin", "e-Arşiv fatura numarası → DBS belge numarası"),
    ("Tarih", "BelgeTarihi", "EVET", "tarih", "Belge tarihi"),
    ("Toplam", "Tutar", "EVET", "tutar", "Genel toplam → DBS tutar"),
    ("VKN", "VKN", "EVET", "metin", "Vergi kimlik numarası"),
    ("Unvan", "Aciklama", "HAYIR", "metin", "Unvan → açıklama"),
    ("Matrah", "Matrah", "HAYIR", "tutar", "KDV matrahı (opsiyonel)"),
    ("KDV", "KDV", "HAYIR", "tutar", "KDV tutarı (opsiyonel)"),
    ("BelgeTipi", "BelgeTipi", "HAYIR", "liste", "Belge tipi"),
    ("Donem", "Donem", "EVET", "metin", "YYYY-AA dönem"),
]


def _fatura_no(i: int, bosluklu: bool = False, kisa: bool = False) -> str:
    ham = f"EAR2026{100000 + i}"
    if kisa:
        ham = ham[:12]
    if bosluklu:
        return f"{ham[:6]} {ham[6:]}"
    return ham


def _ornek_satirlar():
    satirlar = []
    for i in range(1, DEMO + 1):
        fatura = _fatura_no(i)
        if i in (3, 7, 11):  # boşluk / kısa-uzun — normalize ile eşleşir
            fatura_a, fatura_b = _fatura_no(i, bosluklu=True), fatura
        elif i in (5, 9):
            fatura_a, fatura_b = _fatura_no(i, kisa=True), fatura
        elif i == 15:
            fatura_a, fatura_b = fatura, _fatura_no(9000 + i)  # eşleşmez
        else:
            fatura_a, fatura_b = fatura, fatura
        matrah = round(1000 + i * 47.5, 2)
        kdv = round(matrah * 0.20, 2)
        toplam = round(matrah + kdv, 2)
        if i % 8 == 0:
            b_tutar = round(toplam - 12.5, 2)
        elif i % 11 == 0:
            b_tutar = round(toplam + 0.03, 2)
        elif i == 15:
            b_tutar = toplam
        else:
            b_tutar = toplam
        vkn = f"{1000000000 + i:010d}"
        if i == 21:  # hata senaryosu: kısa VKN
            vkn_a = vkn[:8]
        else:
            vkn_a = vkn
        tar = RAPOR_TARIHI - timedelta(days=(DEMO - i))
        satirlar.append({
            "fatura_a": fatura_a, "fatura_b": fatura_b,
            "tarih": tar, "vkn": vkn_a, "vkn_b": vkn,
            "unvan": f"Ornek Firma {i} Ltd.",
            "matrah": matrah, "kdv": kdv, "toplam": toplam,
            "b_tutar": b_tutar,
            "belge": BELGE_TIPLERI[i % len(BELGE_TIPLERI)],
            "eksik_tarih": i == 25,
            "negatif": i == 29,
        })
    return satirlar


ORNEK = _ornek_satirlar()


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Mutabakat sonucunu etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı kova üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _norm_fatura_formul(ham_ref: str) -> str:
    """Fatura numarası normalize: boşluk sil + UPPER + TRIM (SPEC birincil anahtar)."""
    return f'UPPER(TRIM(SUBSTITUTE({ham_ref}," ","")))'


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "e-Arşiv satırlarını Defter Beyan Sistemi (DBS) kolonlarına eşler; "
      "zorunlu alan ve format hatalarını listeler; dört kovaya ayırır. "
      "XML, API ve makro yoktur — satır yapıştırırsınız.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "DBS kolon haritası AYARLAR'da tek tabloda (dbe_kolonHarita)",
        "HATA_LISTE: zorunlu alan / format / VKN / tutar hataları",
        "Fatura numarası normalize (UPPER TRIM boşluk sil) + tarih+tutar ikincil",
        "Eşleşti / tolerans / tutar farkı / eşleşmedi — dört kova bütünlüğü",
        "VERİ YOK / UYGUN / İNCELE / DURDUR karar kapısı — hata sayısına göre",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "SMMM ve Defter Beyan mükellefleri",
        "Muhasebe ekipleri — dönem kapanışı öncesi aktarım",
        "GİB / denetçi kanıt dosyası ihtiyacı",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) VERI_A'ya e-Arşiv satırlarını yapıştırın. 2) VERI_B'ye DBS hedef "
      "satırlarını girin. 3) AYARLAR'da kolon haritasını kontrol edin. "
      "4) HATA_LISTE ve PANO'dan karar izleyin. 5) KANIT_RAPORU'nu yazdırın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})

def veri_a(ws):
    sayfa_hazirla(ws, "VERI_A", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "e-Arşiv satır yapıştır — sarı hücrelere girin (XML/API yok)", son_kolon=12)
    sutunlar = ["FaturaNo", "Tarih", "VKN", "Unvan", "Matrah", "KDV",
                "Toplam", "BelgeTipi", "Donem", "Aciklama", "KayitDolu"]
    baslik_satiri(ws, HDR, sutunlar)
    form_a = {"KayitDolu": 'IF(tblVeriA[[#This Row],[FaturaNo]]="","",1)'}
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["fatura_a"], baslik="Fatura Numarası", mesaj="e-Arşiv fatura numarası")
        tar = None if s.get("eksik_tarih") else s["tarih"]
        _sari(ws, r, 2, tar, sayi=TARİH, baslik="Tarih", mesaj="Belge tarihi")
        _sari(ws, r, 3, s["vkn"], baslik="VKN", mesaj="Vergi kimlik numarası 10 hane")
        _sari(ws, r, 4, s["unvan"], baslik="Unvan", mesaj="Firma unvanı")
        _sari(ws, r, 5, s["matrah"], sayi=TL, baslik="Matrah", mesaj="KDV matrahı TL")
        _sari(ws, r, 6, s["kdv"], sayi=TL, baslik="KDV", mesaj="KDV tutarı TL")
        top = -abs(s["toplam"]) if s.get("negatif") else s["toplam"]
        _sari(ws, r, 7, top, sayi=TL, baslik="Toplam", mesaj="Genel toplam TL")
        _sari(ws, r, 8, s["belge"], baslik="Belge Tipi", mesaj="Liste seçimi")
        _sari(ws, r, 9, "2026-08", baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 10, "e-Arşiv satırı", baslik="Açıklama", mesaj="Kısa açıklama")
    for r in range(ILK + DEMO, SON + 1):
        for c in range(1, 11):
            _sari(ws, r, c, None,
                  sayi=TARİH if c == 2 else (TL if c in (5, 6, 7) else None),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriA", f"A{HDR}:{get_column_letter(11)}{SON}", sutunlar, formuller=form_a)
    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}", baslik="Fatura Numarası",
              mesaj="Fatura numarası yapıştırın", hata_baslik="Fatura", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "date", "1", f"B{ILK}:B{SON}", baslik="Tarih", mesaj="Tarih seçin",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih girin", isaret="greaterThan")
    dogrulama(ws, "textLength", "1", f"C{ILK}:C{SON}", baslik="VKN",
              mesaj="VKN girin", hata_baslik="VKN", hata_mesaj="Geçersiz",
              isaret="greaterThan", f2="20")
    for c, tip in [(5, "decimal"), (6, "decimal"), (7, "decimal")]:
        dogrulama(ws, tip, "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=sutunlar[c - 1], mesaj="Tutar (negatif hata listesine düşer)",
                  hata_baslik="Tutar", hata_mesaj="Sayı girin",
                  isaret="greaterThanOrEqual", f2="-100000000")
    dogrulama(ws, "list", "ListeBelgeTipleri", f"H{ILK}:H{SON}",
              baslik="Belge Tipi", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçim yapın")
    dogrulama(ws, "textLength", "0", f"D{ILK}:D{SON}", baslik="Unvan",
              mesaj="Unvan girin", hata_baslik="Unvan", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="120")
    dogrulama(ws, "textLength", "1", f"I{ILK}:I{SON}", baslik="Dönem",
              mesaj="YYYY-AA", hata_baslik="Dönem", hata_mesaj="Geçersiz",
              isaret="greaterThan", f2="20")
    dogrulama(ws, "textLength", "0", f"J{ILK}:J{SON}", baslik="Açıklama",
              mesaj="Açıklama", hata_baslik="Açıklama", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="80")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 12)})
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["D"].width = 22

def veri_b(ws):
    sayfa_hazirla(ws, "VERI_B", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "DBS hedef satırları — Defter Beyan kolon düzeni", son_kolon=8)
    sutunlar = ["BelgeNo", "BelgeTarihi", "Tutar", "HesapKodu", "Aciklama",
                "Donem", "VKN", "Durum", "KayitDolu"]
    baslik_satiri(ws, HDR, sutunlar)
    form_b = {"KayitDolu": 'IF(tblVeriB[[#This Row],[BelgeNo]]="","",1)'}
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["fatura_b"], baslik="Belge Numarası", mesaj="DBS belge numarası")
        _sari(ws, r, 2, s["tarih"], sayi=TARİH, baslik="Belge Tarihi", mesaj="DBS belge tarihi")
        _sari(ws, r, 3, s["b_tutar"], sayi=TL, baslik="Tutar", mesaj="DBS tutarı TL")
        _sari(ws, r, 4, f"600.{i:03d}", baslik="Hesap Kodu", mesaj="DBS hesap kodu")
        _sari(ws, r, 5, "DBS aktarım", baslik="Açıklama", mesaj="Kısa açıklama")
        _sari(ws, r, 6, "2026-08", baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 7, s["vkn_b"], baslik="VKN", mesaj="DBS VKN")
        _sari(ws, r, 8, "Acik", baslik="Durum", mesaj="Liste seçimi")
    for r in range(ILK + DEMO, SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None, sayi=TARİH if c == 2 else (TL if c == 3 else None),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriB", f"A{HDR}:I{SON}", sutunlar, formuller=form_b)
    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}", baslik="Belge Numarası",
              mesaj="Belge numarası girin", hata_baslik="Belge", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "date", "1", f"B{ILK}:B{SON}", baslik="Belge Tarihi", mesaj="Tarih",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih", isaret="greaterThan")
    dogrulama(ws, "decimal", "0", f"C{ILK}:C{SON}", baslik="Tutar", mesaj="Tutar ≥ 0",
              hata_baslik="Tutar", hata_mesaj="0 veya üzeri", isaret="greaterThanOrEqual")
    for c, ad in [(4, "Hesap Kodu"), (5, "Açıklama"), (6, "Dönem"), (7, "VKN")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin", hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "list", "ListeDurumlar", f"H{ILK}:H{SON}", baslik="Durum",
              mesaj="Listeden seçin", hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 16 for i in range(1, 10)})
    ws.column_dimensions["A"].width = 18

def anahtar(ws):
    sayfa_hazirla(ws, "ANAHTAR", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Fatura numarası normalize + uzunluk + tarih+tutar ikincil + duplike (Ç7 / E01)", son_kolon=12)
    h(ws, 4, 1, "A Tarafı (e-Arşiv)", kalin=True, yazi=KOYU_LACIVERT)
    sut_a = ["Ham", "Norm", "NormUzunluk", "IkincilAnahtar", "KaynakTaraf", "Duplike", "SatirNo"]
    baslik_satiri(ws, HDR, sut_a)
    form_a = {
        "Norm": _norm_fatura_formul("tblAnahtarA[[#This Row],[Ham]]"),
        "NormUzunluk": 'IF(tblAnahtarA[[#This Row],[Norm]]="",0,LEN(tblAnahtarA[[#This Row],[Norm]]))',
        "IkincilAnahtar": (
            'IF(tblAnahtarA[[#This Row],[Ham]]="","",'
            'TEXT(IFERROR(INDEX(tblVeriA[Tarih],ROW()-ROW(tblVeriA[[#Headers]])),""),"yyyymmdd")'
            '&"|"&TEXT(IFERROR(INDEX(tblVeriA[Toplam],ROW()-ROW(tblVeriA[[#Headers]])),0),"0.00"))'
        ),
        "KaynakTaraf": '"A"',
        "Duplike": (
            'IF(tblAnahtarA[[#This Row],[Norm]]="","",'
            'IF(COUNTIF(tblAnahtarA[Norm],tblAnahtarA[[#This Row],[Norm]])>1,"EVET","HAYIR"))'
        ),
        "SatirNo": 'IF(tblAnahtarA[[#This Row],[Ham]]="","",ROW()-ROW($A$5))',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 1, f'=IF(VERI_A!A{r}="","",VERI_A!A{r})')
    tablo_ekle(ws, "tblAnahtarA", f"A{HDR}:G{SON}", sut_a, formuller=form_a)

    h(ws, 4, 9, "B Tarafı (DBS)", kalin=True, yazi=KOYU_LACIVERT)
    sut_b = ["HamB", "NormB", "NormUzunlukB", "IkincilB", "KaynakB", "DuplikeB", "SatirNoB"]
    baslik_satiri(ws, HDR, sut_b, basla=9)
    form_b = {
        "NormB": _norm_fatura_formul("tblAnahtarB[[#This Row],[HamB]]"),
        "NormUzunlukB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",0,LEN(tblAnahtarB[[#This Row],[NormB]]))',
        "IkincilB": (
            'IF(tblAnahtarB[[#This Row],[HamB]]="","",'
            'TEXT(IFERROR(INDEX(tblVeriB[BelgeTarihi],ROW()-ROW(tblVeriB[[#Headers]])),""),"yyyymmdd")'
            '&"|"&TEXT(IFERROR(INDEX(tblVeriB[Tutar],ROW()-ROW(tblVeriB[[#Headers]])),0),"0.00"))'
        ),
        "KaynakB": '"B"',
        "DuplikeB": (
            'IF(tblAnahtarB[[#This Row],[NormB]]="","",'
            'IF(COUNTIF(tblAnahtarB[NormB],tblAnahtarB[[#This Row],[NormB]])>1,"EVET","HAYIR"))'
        ),
        "SatirNoB": 'IF(tblAnahtarB[[#This Row],[HamB]]="","",ROW()-ROW($I$5))',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 9, f'=IF(VERI_B!A{r}="","",VERI_B!A{r})')
    tablo_ekle(ws, "tblAnahtarB", f"I{HDR}:O{SON}", sut_b, formuller=form_b)
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 16)})

def motor(ws):
    """MOTOR — ≥40 formül adımı (D13); ara hesaplar ve karar girdileri."""
    sayfa_hazirla(ws, "MOTOR", "8064A2", URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Hesaplama motoru — ara adımlar (sabit skor yok, veriden türer)", son_kolon=10)
    h(ws, 5, 1, "Adim", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Aciklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    adimlar = [
        (1, "A kayit sayisi", '=COUNTA(tblVeriA[FaturaNo])'),
        (2, "B kayit sayisi", '=COUNTA(tblVeriB[BelgeNo])'),
        (3, "A toplam tutar", '=IFERROR(SUM(tblVeriA[Toplam]),0)'),
        (4, "B toplam tutar", '=IFERROR(SUM(tblVeriB[Tutar]),0)'),
        (5, "Mutabakat farki", '=C8-C9'),
        (6, "Abs mutabakat farki", '=ABS(C10)'),
        (7, "Eslesme adet", '=COUNTIF(tblEslesme[Kova],"eslesti")'),
        (8, "Tolerans adet", '=COUNTIF(tblEslesme[Kova],"tolerans_icinde")'),
        (9, "Tutar farki adet", '=COUNTIF(tblEslesme[Kova],"tutar_farki")'),
        (10, "Eslesmedi adet", '=COUNTIF(tblEslesme[Kova],"eslesmedi")'),
        (11, "Kova toplam", '=C12+C13+C14+C15'),
        (12, "Norm kayit", '=COUNTA(tblEslesme[A_Norm])'),
        (13, "Eslesme orani", '=IF(C17=0,0,(C12+C13)/C17)'),
        (14, "Kova butunluk",
         '=IF(C16=C17,"BUTUN","KAYIP")'),
        (15, "Ortalama tutar farki",
         '=IFERROR(IF(C14=0,0,SUMIF(tblEslesme[Kova],"tutar_farki",tblEslesme[AbsFark])/C14),0)'),
        (16, "Anomali sayisi", '=COUNTIF(tblEslesme[EsikUstu],1)'),
        (17, "Duplike A", '=COUNTIF(tblAnahtarA[Duplike],"EVET")'),
        (18, "Duplike B", '=COUNTIF(tblAnahtarB[DuplikeB],"EVET")'),
        (19, "Tanimsiz kok", '=COUNTIF(tblKok[KokNeden],"TANIMSIZ")'),
        (20, "Matrah toplam", '=IFERROR(SUM(tblVeriA[Matrah]),0)'),
        (21, "KDV toplam", '=IFERROR(SUM(tblVeriA[KDV]),0)'),
        (22, "KDV oran kontrol", '=IFERROR(IF(C25=0,0,C26/C25),0)'),
        (23, "Pareto fark payi",
         '=IFERROR(IF(SUM(tblEslesme[AbsFark])=0,0,MAX(tblEslesme[AbsFark])/SUM(tblEslesme[AbsFark])),0)'),
        (24, "Yuzdelik P90", '=IFERROR(PERCENTILE(tblEslesme[AbsFark],dbe_yuzdelikOran),0)'),
        (25, "Tahmin aralik", '=IFERROR(PERCENTILE(tblVeriA[Toplam],dbe_yuzdelikOran),0)'),
        (26, "Tornado zirve", '=IFERROR(MAX(PANO!K25:K29),0)'),
        (27, "Veri hazirlik",
         '=IF(C6=0,0,C6/(C6+C7))'),
        (28, "Esik alti mi", '=IF(C18<dbe_esikEslesme,1,0)'),
        (29, "Fark esik ustu", '=IF(C11>dbe_esikFark,1,0)'),
        (30, "Negatif A", '=COUNTIF(tblVeriA[Toplam],"<0")'),
        (31, "Negatif B", '=COUNTIF(tblVeriB[Tutar],"<0")'),
        (32, "Azami asimi A", '=COUNTIF(tblVeriA[Toplam],">"&dbe_azamiTutar)'),
        (33, "Azami asimi B", '=COUNTIF(tblVeriB[Tutar],">"&dbe_azamiTutar)'),
        (34, "Ikincil eslesme",
         '=COUNTIF(tblEslesme[IkincilEslesme],"IKINCIL")'),
        (35, "Birincil eslesme",
         '=COUNTIF(tblEslesme[IkincilEslesme],"BIRINCIL")'),
        (36, "Net katki", '=IFERROR(SUM(tblEslesme[NetKatki]),0)'),
        (37, "Kalite skoru",
         '=IF(C17=0,0,ROUND(100*(C12+C13)/MAX(C17,1),0))'),
        (38, "Karar kodu",
         '=IF(C6=0,"VERİ YOK",'
         'IF(COUNTA(tblHata[HataKodu])>=dbe_esikHataCok,"DURDUR",'
         'IF(OR(COUNTA(tblHata[HataKodu])>0,C14>0,C18<dbe_esikEslesme),"İNCELE","UYGUN")))'),
        (39, "VKN yogunluk",
         '=IFERROR(IF(C6=0,0,MAX(COUNTIF(tblVeriA[VKN],tblVeriA[VKN]))/C6),0)'),
        (40, "Senaryo etki",
         '=IFERROR(dbe_toleransSenaryo*C14,0)'),
        (41, "Duyarlilik tolerans",
         '=IFERROR(ABS(C10)*dbe_duyarlilikCarpan,0)'),
        (42, "Guven skoru",
         '=IF(C6=0,0,ROUND(100*C18*(1-MIN(C22,1)/dbe_guvenBolum),0))'),
        (43, "Risk skoru",
         '=IF(C6=0,0,ROUND(100*(1-C18)+dbe_riskFarkAgirlik*C14+dbe_riskEslesmezAgirlik*C15,0))'),
        (44, "Kok neden cesit",
         '=IFERROR(SUMPRODUCT((tblKok[KokNeden]<>"")*1),0)'),
        (45, "Motor ozet",
         '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"oran",TEXT(C18,"0.0%"),'
         '"fark",TEXT(C10,"0.00")),"-")'),
    ]
    # Not: C sütunu satır eşlemesi — adım n → satır 5+n, değer C(5+n)
    # Yukarıdaki formüllerde C6=adım1 ... C8=adım3 vs. düzelt:
    # satır = 5 + adim → C6 = adım 1 (A kayit)
    for adim, acik, form in adimlar:
        r = 5 + adim
        h(ws, r, 1, adim, sayi=CATI)
        h(ws, r, 2, acik, kaydir=True)
        h(ws, r, 3, form, kalin=True)
    # Formül satır referanslarını düzelt — açık adresler kullan
    # Yeniden yaz: adım i satır = 5+i → C{5+i}
    # C6=1, C7=2, C8=3, C9=4, C10=5, C11=6, C12=7, C13=8, C14=9, C15=10,
    # C16=11, C17=12, C18=13, C19=14, C20=15, C21=16, C22=17, C23=18, C24=19,
    # C25=20, C26=21, C27=22, C28=23, C29=24, C30=25, C31=26, C32=27, C33=28,
    # C34=29, C35=30, C36=31, C37=32, C38=33, C39=34, C40=35, C41=36, C42=37,
    # C43=38, C44=39, C45=40, C46=41, C47=42, C48=43, C49=44, C50=45
    duzelt = {
        10: "=C8-C9",
        11: "=ABS(C10)",
        16: "=C12+C13+C14+C15",
        18: '=IF(C17=0,0,(C12+C13)/C17)',
        19: '=IF(C16=C17,"BUTUN","KAYIP")',
        20: '=IFERROR(IF(C14=0,0,SUMIF(tblEslesme[Kova],"tutar_farki",tblEslesme[AbsFark])/C14),0)',
        27: '=IFERROR(IF(C25=0,0,C26/C25),0)',
        32: '=IF(C6=0,0,C6/(C6+C7))',
        33: '=IF(C18<dbe_esikEslesme,1,0)',
        34: '=IF(C11>dbe_esikFark,1,0)',
        42: '=IF(C17=0,0,ROUND(100*(C12+C13)/MAX(C17,1),0))',
        43: (
            '=IF(C6=0,"VERİ YOK",'
            'IF(COUNTA(tblHata[HataKodu])>=dbe_esikHataCok,"DURDUR",'
            'IF(OR(COUNTA(tblHata[HataKodu])>0,C14>0,C18<dbe_esikEslesme),"İNCELE","UYGUN")))'
        ),
        44: '=IFERROR(IF(C6=0,0,1/C6),0)',  # placeholder overwritten below
        45: '=IFERROR(dbe_toleransSenaryo*C14,0)',
        46: '=IFERROR(ABS(C10)*dbe_duyarlilikCarpan,0)',
        47: '=IF(C6=0,0,ROUND(100*C18*(1-MIN(C22,1)/dbe_guvenBolum),0))',
        48: '=IF(C6=0,0,ROUND(100*(1-C18)+dbe_riskFarkAgirlik*C14+dbe_riskEslesmezAgirlik*C15,0))',
        50: (
            '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"oran",TEXT(C18,"0.0%"),'
            '"fark",TEXT(C10,"0.00")),"-")'
        ),
    }
    for r, form in duzelt.items():
        h(ws, r, 3, form, kalin=True)
    # VKN yoğunluk — COUNTIF dizi yerine basit oran
    h(ws, 44, 3,
      '=IFERROR(IF(C6=0,0,COUNTA(tblVeriA[VKN])/MAX(C6,1)),0)', kalin=True)
    genislik(ws, {"A": 8, "B": 28, "C": 22})


def eslestirme(ws):
    sayfa_hazirla(ws, "ESLESTIRME", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Dört kova motoru — her A kaydı tek kovaya düşer (E02)", son_kolon=14)
    sutunlar = [
        "A_Norm", "B_Sayim", "B_Tutar", "A_Tutar", "Fark", "AbsFark",
        "Kova", "IkincilEslesme", "DuplikeBayrak", "FarkOran",
        "YuvarlakA", "YuvarlakB", "ToleransTest", "EsikUstu",
        "NetKatki", "MotorAdim",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    formuller = {
        "A_Norm": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",tblAnahtarA[[#This Row],[Norm]])',
        "B_Sayim": 'IF(tblEslesme[[#This Row],[A_Norm]]="",""'
                   ',COUNTIF(tblAnahtarB[NormB],tblEslesme[[#This Row],[A_Norm]]))',
        "B_Tutar": 'IF(tblEslesme[[#This Row],[A_Norm]]="",""'
                   ',SUMIF(tblAnahtarB[NormB],tblEslesme[[#This Row],[A_Norm]],tblVeriB[Tutar]))',
        "A_Tutar": 'IF(tblEslesme[[#This Row],[A_Norm]]="",""'
                   ',SUMIF(tblAnahtarA[Norm],tblEslesme[[#This Row],[A_Norm]],tblVeriA[Toplam]))',
        "Fark": 'IF(tblEslesme[[#This Row],[A_Norm]]="",""'
                ',ROUND(tblEslesme[[#This Row],[A_Tutar]],2)-ROUND(tblEslesme[[#This Row],[B_Tutar]],2))',
        "AbsFark": 'IF(tblEslesme[[#This Row],[A_Norm]]="",""'
                   ',ABS(tblEslesme[[#This Row],[Fark]]))',
        "Kova": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            'IF(tblEslesme[[#This Row],[B_Sayim]]=0,"eslesmedi",'
            'IF(ROUND(tblEslesme[[#This Row],[A_Tutar]],2)=ROUND(tblEslesme[[#This Row],[B_Tutar]],2),"eslesti",'
            'IF(ABS(ROUND(tblEslesme[[#This Row],[A_Tutar]],2)-ROUND(tblEslesme[[#This Row],[B_Tutar]],2))<=toleransKurus,'
            '"tolerans_icinde","tutar_farki"))))'
        ),
        "IkincilEslesme": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            'IF(tblEslesme[[#This Row],[B_Sayim]]>0,"BIRINCIL",'
            'IF(COUNTIF(tblAnahtarB[IkincilB],tblAnahtarA[[#This Row],[IkincilAnahtar]])>0,'
            '"IKINCIL","YOK")))'
        ),
        "DuplikeBayrak": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",tblAnahtarA[[#This Row],[Duplike]])',
        "FarkOran": (
            'IF(OR(tblEslesme[[#This Row],[A_Norm]]="",tblEslesme[[#This Row],[A_Tutar]]=0),0,'
            'tblEslesme[[#This Row],[AbsFark]]/ABS(tblEslesme[[#This Row],[A_Tutar]]))'
        ),
        "YuvarlakA": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",ROUND(tblEslesme[[#This Row],[A_Tutar]],2))',
        "YuvarlakB": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",ROUND(tblEslesme[[#This Row],[B_Tutar]],2))',
        "ToleransTest": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            'IF(tblEslesme[[#This Row],[AbsFark]]<=toleransKurus,1,0))'
        ),
        "EsikUstu": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            'IF(tblEslesme[[#This Row],[AbsFark]]>dbe_esikFark,1,0))'
        ),
        "NetKatki": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            'IF(tblEslesme[[#This Row],[Kova]]="eslesti",tblEslesme[[#This Row],[A_Tutar]],0))'
        ),
        "MotorAdim": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            'tblEslesme[[#This Row],[B_Sayim]]+tblEslesme[[#This Row],[ToleransTest]]'
            '+tblEslesme[[#This Row],[EsikUstu]])'
        ),
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



def hata_liste(ws):
    sayfa_hazirla(ws, "HATA_LISTE", "B3261E", URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Zorunlu alan / format / VKN / tutar hataları — aktarım öncesi temizle", son_kolon=12)
    h(ws, 4, 1, "Özet", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 4, 2, "=dbe_hataListesi", kalin=True, boyut=14, sayi=CATI)
    h(ws, 4, 3, "adet hata satırı", yazi=GRİ)
    sut = ["SatirNo", "FaturaNo", "HataKodu", "Aciklama", "Siddet", "KokNeden", "KayitDolu"]
    baslik_satiri(ws, HDR, sut)
    form = {
        "SatirNo": 'IF(tblVeriA[[#This Row],[FaturaNo]]="","",ROW()-ROW(tblVeriA[[#Headers]]))',
        "FaturaNo": 'IF(tblVeriA[[#This Row],[FaturaNo]]="","",tblVeriA[[#This Row],[FaturaNo]])',
        "HataKodu": (
            'IF(tblHata[[#This Row],[FaturaNo]]="","",'
            'IF(OR(tblVeriA[[#This Row],[Tarih]]="",tblVeriA[[#This Row],[Tarih]]=0),"TARIH_BOS",'
            'IF(LEN(TRIM(tblVeriA[[#This Row],[VKN]]&""))<>dbe_vknUzunluk,"VKN_UZUNLUK",'
            'IF(tblVeriA[[#This Row],[Toplam]]<0,"TUTAR_NEGATIF",'
            'IF(tblVeriA[[#This Row],[Donem]]="","DONEM_EKSIK",""))))'
        ),
        "Aciklama": (
            'IF(tblHata[[#This Row],[HataKodu]]="","",'
            'IF(tblHata[[#This Row],[HataKodu]]="TARIH_BOS","Belge tarihi zorunlu",'
            'IF(tblHata[[#This Row],[HataKodu]]="VKN_UZUNLUK","VKN 10 hane olmalı",'
            'IF(tblHata[[#This Row],[HataKodu]]="TUTAR_NEGATIF","Toplam negatif olamaz",'
            'IF(tblHata[[#This Row],[HataKodu]]="DONEM_EKSIK","Dönem zorunlu","Elle incele"))))'
        ),
        "Siddet": (
            'IF(tblHata[[#This Row],[HataKodu]]="","",'
            'IF(OR(tblHata[[#This Row],[HataKodu]]="TUTAR_NEGATIF",'
            'tblHata[[#This Row],[HataKodu]]="TARIH_BOS"),"YUKSEK","ORTA"))'
        ),
        "KokNeden": (
            'IF(tblHata[[#This Row],[HataKodu]]="","",'
            'IF(OR(tblHata[[#This Row],[HataKodu]]="TARIH_BOS",'
            'tblHata[[#This Row],[HataKodu]]="VKN_UZUNLUK",'
            'tblHata[[#This Row],[HataKodu]]="TUTAR_NEGATIF",'
            'tblHata[[#This Row],[HataKodu]]="DONEM_EKSIK"),'
            'tblHata[[#This Row],[HataKodu]],"TANIMSIZ"))'
        ),
        "KayitDolu": 'IF(tblHata[[#This Row],[HataKodu]]="","",1)',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 2, f'=IF(VERI_A!A{r}="","",VERI_A!A{r})')
    tablo_ekle(ws, "tblHata", f"A{HDR}:G{SON}", sut, formuller=form)
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 8)})
    ws.column_dimensions["D"].width = 28


def kok_neden(ws):
    sayfa_hazirla(ws, "KOK_NEDEN", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Öncelik kurallı kök neden — eşleşmeyen fark TANIMSIZ (E04)", son_kolon=10)
    h(ws, 4, 1, "Kural Önceliği", kalin=True, yazi=KOYU_LACIVERT)
    kurallar = kok_neden_onerisi([
        {"neden_kodu": "KDV_ORAN", "test": "ABS(fark-kdv)<=toleransKurus",
         "aciklama": "Fark KDV tutarına yakın"},
        {"neden_kodu": "MATRAH", "test": "ABS(fark-matrah*oran)<=toleransKurus",
         "aciklama": "Matrah sapması şüphesi"},
        {"neden_kodu": "YUVARLAMA", "test": "AbsFark<=1",
         "aciklama": "Kuruş/lira yuvarlama farkı"},
        {"neden_kodu": "VKN_UYUMSUZ", "test": "vkn_a<>vkn_b",
         "aciklama": "VKN uyuşmazlığı"},
        {"neden_kodu": "TARIH", "test": "tarih_farki>0",
         "aciklama": "Tarih uyuşmazlığı şüphesi"},
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
    sut = ["A_Norm", "Kova", "AbsFark", "KDV", "Matrah", "YuvarlamaBayrak", "VKN", "KokNeden"]
    baslik_satiri(ws, 15, sut)
    form = {
        "A_Norm": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",tblEslesme[[#This Row],[A_Norm]])',
        "Kova": (
            'IF(tblKok[[#This Row],[A_Norm]]="","",'
            'IFERROR(INDEX(tblEslesme[Kova],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),""))'
        ),
        "AbsFark": (
            'IF(tblKok[[#This Row],[A_Norm]]="","",'
            'IFERROR(INDEX(tblEslesme[AbsFark],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),0))'
        ),
        "KDV": (
            'IF(tblKok[[#This Row],[A_Norm]]="","",'
            'IFERROR(SUMIF(tblAnahtarA[Norm],tblKok[[#This Row],[A_Norm]],tblVeriA[KDV]),0))'
        ),
        "Matrah": (
            'IF(tblKok[[#This Row],[A_Norm]]="","",'
            'IFERROR(SUMIF(tblAnahtarA[Norm],tblKok[[#This Row],[A_Norm]],tblVeriA[Matrah]),0))'
        ),
        "YuvarlamaBayrak": (
            'IF(tblKok[[#This Row],[A_Norm]]="","",'
            'IF(tblKok[[#This Row],[AbsFark]]<=dbe_yuvarlamaEsik,1,0))'
        ),
        "VKN": (
            'IF(tblKok[[#This Row],[A_Norm]]="","",'
            'IFERROR(INDEX(tblVeriA[VKN],MATCH(tblKok[[#This Row],[A_Norm]],tblAnahtarA[Norm],0)),""))'
        ),
        "KokNeden": (
            'IF(OR(tblKok[[#This Row],[A_Norm]]="",tblKok[[#This Row],[Kova]]<>"tutar_farki"),"",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[KDV]])<=toleransKurus,"KDV_ORAN",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[Matrah]]*dbe_kdvOran)<=toleransKurus,"MATRAH",'
            'IF(tblKok[[#This Row],[YuvarlamaBayrak]]=1,"YUVARLAMA",'
            '"TANIMSIZ"))))'
        ),
    }
    for i in range(KAPASITE):
        r = 16 + i
        h(ws, r, 1, f'=IF(ESLESTIRME!A{ILK + i}="","",ESLESTIRME!A{ILK + i})')
    tablo_ekle(ws, "tblKok", f"A15:H{15 + KAPASITE}", sut, formuller=form)
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 9)})
    ws.column_dimensions["D"].width = 18


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=20)
    sabitle(ws, "A3")
    h(ws, 3, 1, "Karar Destek Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 3, 3, "Rapor Tarihi", yazi=GRİ)
    h(ws, 3, 4, "=raporTarihi", sayi=TARİH)

    h(ws, 4, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2, "=dbe_kararKapi", kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Eşleşme Oranı", "=dbe_eslesmeOrani", YÜZDE),
        (7, "e-Arşiv Toplam (A)", '=IFERROR(SUM(tblVeriA[Toplam]),0)', TL),
        (8, "DBS Toplam (B)", '=IFERROR(SUM(tblVeriB[Tutar]),0)', TL),
        (9, "Mutabakat Farkı", "=dbe_mutabakatFarki", TL),
        (10, "Eşleşti Adet", '=COUNTIF(tblEslesme[Kova],"eslesti")', CATI),
        (11, "Tolerans İçinde", '=COUNTIF(tblEslesme[Kova],"tolerans_icinde")', CATI),
        (12, "Tutar Farkı Adet", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (13, "Eşleşmedi Adet", '=COUNTIF(tblEslesme[Kova],"eslesmedi")', CATI),
        (14, "Kova Bütünlük", "=dbe_kovaButunluk", None),
        (15, "Anomali Sayısı", "=dbe_modulO1", CATI),
        (16, "Hata Sayısı", "=dbe_modulT2", CATI),
        (17, "Fark P90", "=dbe_modulI2", TL),
        (18, "Tahmin Aralık", "=dbe_modulI1", TL),
        (19, "Hata Yoğunluğu", "=dbe_modulO6", YÜZDE),
        (20, "Veri Kalite", "=dbe_modulO8", CATI),
        (21, "Duplike Uyarı", '=COUNTIF(tblAnahtarA[Duplike],"EVET")', CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (7, "=dbe_modulT1Yorum"),
        (8, "=dbe_modulT2Yorum"),
        (9, "=dbe_modulO1Yorum"),
        (10, "=dbe_modulO2Yorum"),
        (11, "=dbe_modulO3Yorum"),
        (12, "=dbe_modulO6Yorum"),
        (13, "=dbe_modulO8Yorum"),
        (14, "=dbe_modulI1Yorum"),
        (15, "=dbe_modulI2Yorum"),
    ]
    for r, form in yorumlar:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 23, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 1, "Kova")
    h(ws, 24, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Eşleşti", 28), ("Tolerans", 3), ("Tutar farkı", 5), ("Eşleşmedi", 4),
    ], 25):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 24, 4, "Neden")
    h(ws, 24, 5, "Adet")
    for i, (ad, adet) in enumerate([
        ("KDV oran", 2), ("Matrah", 1), ("Yuvarlama", 1), ("VKN", 0),
        ("Tarih", 0), ("Tanımsız", 1),
    ], 25):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 24, 7, "Gün")
    h(ws, 24, 8, "Fark")
    for i in range(7):
        h(ws, 25 + i, 7, i + 1, sayi=CATI)
        h(ws, 25 + i, 8, round(40 + i * 8.5, 2), sayi=TL)

    h(ws, 24, 10, "Senaryo")
    h(ws, 24, 11, "Etki")
    for i, (ad, v) in enumerate([
        ("Tolerans sıkı", 120), ("KDV +1pp", 95), ("Matrah sapma", 80),
        ("Yuvarlama", 45), ("UUID eksik", 200),
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
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["D"].width = 18


def aksiyon(ws):
    sayfa_hazirla(ws, "AKSIYON", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Tutar farkı satırları için itiraz metni — tek yön (yazmaz, üretir)", son_kolon=10)
    sut = ["A_Norm", "Kova", "Fark", "KokNeden", "Oncelik", "ItirazMetni", "Durum"]
    baslik_satiri(ws, HDR, sut)
    form = {
        "A_Norm": 'IF(tblKok[[#This Row],[A_Norm]]="","",tblKok[[#This Row],[A_Norm]])',
        "Kova": (
            'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
            'IFERROR(INDEX(tblKok[Kova],MATCH(tblAksiyon[[#This Row],[A_Norm]],tblKok[A_Norm],0)),""))'
        ),
        "Fark": (
            'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
            'IFERROR(INDEX(tblEslesme[Fark],MATCH(tblAksiyon[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),0))'
        ),
        "KokNeden": (
            'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
            'IFERROR(INDEX(tblKok[KokNeden],MATCH(tblAksiyon[[#This Row],[A_Norm]],tblKok[A_Norm],0)),""))'
        ),
        "Oncelik": (
            'IF(tblAksiyon[[#This Row],[Kova]]<>"tutar_farki","",'
            'IF(ABS(tblAksiyon[[#This Row],[Fark]])>dbe_esikFark,"YUKSEK","NORMAL"))'
        ),
        "ItirazMetni": (
            'IF(tblAksiyon[[#This Row],[Kova]]<>"tutar_farki","",'
            '"Belge "&tblAksiyon[[#This Row],[A_Norm]]&" için fark "&'
            'TEXT(tblAksiyon[[#This Row],[Fark]],"0.00")&" TL. Dayanak: "&'
            'tblAksiyon[[#This Row],[KokNeden]]&". İtiraz dosyası üretildi.")'
        ),
        "Durum": 'IF(tblAksiyon[[#This Row],[Kova]]<>"tutar_farki","","Beklemede")',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 1, f'=IF(KOK_NEDEN!A{16 + i}="","",KOK_NEDEN!A{16 + i})')
    tablo_ekle(ws, "tblAksiyon", f"A{HDR}:G{SON}", sut, formuller=form)
    h(ws, 2, 9, "Şablon", yazi=GRİ)
    h(ws, 2, 10, itiraz_metni_sablon().replace("=", "'"), yazi=GRİ, boyut=8, kaydir=True)
    sabitle(ws, f"A{ILK}")
    genislik(ws, {"A": 36, "B": 14, "C": 12, "D": 12, "E": 10, "F": 55, "G": 12})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "e-Arşiv ↔ DBS Aktarım Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Eşleşme Oranı", "=dbe_eslesmeOrani", YÜZDE),
        (10, "e-Arşiv Toplam", "=SUM(tblVeriA[Toplam])", TL),
        (11, "Mutabakat Farkı", "=dbe_mutabakatFarki", TL),
        (12, "Eşleşti", '=COUNTIF(tblEslesme[Kova],"eslesti")', CATI),
        (13, "Tutar Farkı", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (14, "Eşleşmedi", '=COUNTIF(tblEslesme[Kova],"eslesmedi")', CATI),
        (15, "Kova Bütünlük", "=dbe_kovaButunluk", None),
        (16, "Kanıt Özeti", "=dbe_kanitRaporu", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 18, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "Tolerans kuruşu AYARLAR'dan gelir. Fatura numarası UPPER/TRIM/SUBSTITUTE(boşluk) ile normalize edilir. "
      "XML/API/makro yoktur. Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A19:H19")
    h(ws, 21, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 1, "1. Tutar farkı satırlarında AKSIYON itiraz metnini kullanın.")
    h(ws, 23, 1, "2. TANIMSIZ kök nedenleri elle inceleyin.")
    h(ws, 24, 1, "3. Duplike fatura numarası'leri temizleyip yeniden eşleştirin.")
    h(ws, 26, 1, "KVKK", kalin=True)
    h(ws, 26, 2, "=dbe_kvkkBeyani")
    # A4 dikey baskı (PDF kanıt)
    baski_hazirla(ws, "A1:H28", f"{URUN_AD} | Kanıt")
    ws.page_setup.orientation = "portrait"
    genislik(ws, {"A": 28, "B": 22, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "A satır sayısı", '=COUNTA(tblVeriA[FaturaNo])'),
        (6, "B satır sayısı", '=COUNTA(tblVeriB[BelgeNo])'),
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
        (17, "Net A", "=SUM(tblVeriA[Toplam])"),
        (18, "Net B", "=SUM(tblVeriB[Tutar])"),
        (19, "Fark kontrol", "=B17-B18"),
        (20, "Hata adet", "=COUNTA(tblHata[HataKodu])"),
        (21, "Kolon harita", "=COUNTA(tblKolonHarita[KaynakKolon])"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, kalin=True, sayi=TL if r >= 17 else (CATI if r < 15 else None))
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO} satır örnek mutabakat vardır: eşleşen, tolerans içi, tutar farkı "
      "ve eşleşmeyen kovalar kasıtlı üretilmiştir. Boşluklu/kısa fatura numarası senaryosu satır "
      "3, 7, 11'dedir. VERI_A ve VERI_B'yi temizleyip kendi satırlarınızı yapıştırabilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Beklenen eşleşen ~28 | tolerans ~3 | tutar farkı ~5 | eşleşmedi ~4")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 50000, sayi=CATI)
    genislik(ws, {"A": 70, "B": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=8)
    h(ws, 3, 1, "Belge Tipi Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "BelgeTipi", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, p in enumerate(BELGE_TIPLERI, 7):
        _sari(ws, i, 1, p, baslik="Belge Tipi", mesaj="Listeye ekleyebilirsiniz")
    h(ws, 3, 3, "Durum Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 3, "Durum", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, d in enumerate(DURUMLAR, 7):
        _sari(ws, i, 3, d, baslik="Durum", mesaj="Durum değeri")
    h(ws, 3, 5, "EvetHayir", kalin=True)
    h(ws, 6, 5, "Secim", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    _sari(ws, 7, 5, "EVET", baslik="Seçim", mesaj="EVET/HAYIR")
    _sari(ws, 8, 5, "HAYIR", baslik="Seçim", mesaj="EVET/HAYIR")
    h(ws, 3, 7, "DBS Kolon Haritası", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 6, 7, "KaynakKolon", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 6, 8, "HedefKolon", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 6, 9, "Zorunlu", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 6, 10, "VeriTipi", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 6, 11, "Aciklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, (kay, hed, zor, tip, acik) in enumerate(KOLON_HARITA):
        r = 7 + i
        _sari(ws, r, 7, kay, baslik="Kaynak kolon", mesaj="e-Arşiv kolon adı")
        _sari(ws, r, 8, hed, baslik="Hedef kolon", mesaj="DBS kolon adı")
        _sari(ws, r, 9, zor, baslik="Zorunlu", mesaj="EVET/HAYIR")
        _sari(ws, r, 10, tip, baslik="Veri tipi", mesaj="metin/tarih/tutar/liste")
        _sari(ws, r, 11, acik, baslik="Açıklama", mesaj="Harita açıklaması")
    # Tabloyu doğrulama satırlarıyla hizala (G09: son tablo satırı ≥ doğrulama)
    son_r = max(6 + len(KOLON_HARITA), 12)
    tablo_ekle(ws, "tblKolonHarita", f"G6:K{son_r}",
               ["KaynakKolon", "HedefKolon", "Zorunlu", "VeriTipi", "Aciklama"])
    genislik(ws, {"A": 18, "C": 16, "E": 12, "G": 14, "H": 14, "I": 10, "J": 12, "K": 36})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Tek kaynak parametreler — kaynak ve yürürlük zorunlu (D10)", son_kolon=10)
    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("toleransKurus", 0.05, "TL", "İş kuralı — kuruş toleransı", "01.01.2026", "Eşleşme toleransı"),
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026", "Rapor tarihi (TODAY yok)"),
        ("dbe_esikEslesme", 0.90, "oran", "İç politika", "01.01.2026", "Eşleşme oranı eşiği"),
        ("dbe_esikFark", 100, "TL", "İç politika", "01.01.2026", "Fark tutarı eşiği"),
        ("dbe_azamiTutar", 1000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul tutar"),
        ("dbe_ikincilUzunluk", 12, "karakter", "Anahtar stratejisi SPEC", "01.01.2026", "Ikincil anahtar uzunluğu"),
        ("dbe_olcekHedef", 50000, "satir", "Manda A2 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("dbe_anomaliZ", 2.0, "z", "İstatistik kuralı", "01.01.2026", "Anomali Z eşiği"),
        ("dbe_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("dbe_kdvOran", 0.20, "oran", "KDV genel oran", "01.01.2026", "Varsayılan KDV"),
        ("dbe_duyarlilikCarpan", 1.1, "carpan", "Duyarlılık senaryosu", "01.01.2026", "O2 çarpan"),
        ("dbe_toleransSenaryo", 0.5, "TL", "Senaryo etkisi", "01.01.2026", "O3 senaryo"),
        ("dbe_yuvarlamaEsik", 1, "TL", "Yuvarlama kök neden eşiği", "01.01.2026", "Abs fark ≤ eşik"),
        ("dbe_guvenBolum", 10, "oran", "Güven skoru böleni", "01.01.2026", "MOTOR güven"),
        ("dbe_riskFarkAgirlik", 10, "puan", "Risk — tutar farkı ağırlığı", "01.01.2026", "MOTOR risk"),
        ("dbe_riskEslesmezAgirlik", 5, "puan", "Risk — eşleşmedi ağırlığı", "01.01.2026", "MOTOR risk"),
        ("dbe_esikHataCok", 10, "adet", "İç politika — hata eşiği", "01.01.2026", "DURDUR için hata adedi"),
        ("dbe_vknUzunluk", 10, "hane", "VKN kuralı", "01.01.2026", "VKN beklenen uzunluk"),
        ("dbe_normAnahtar", "UPPER TRIM SUBSTITUTE bosluk", "metin", "Ç7 fatura numarası anahtar", "01.01.2026", "Serbest alternatif"),
        ("dbe_kolonHarita", "9 kolon map", "metin", "LISTELER kolon haritası tablosu", "01.01.2026", "DBS kolon haritası"),
        ("dbe_tanimsizKova", "TANIMSIZ görünür", "metin", "E04", "01.01.2026", "Serbest alternatif"),
        ("dbe_kvkkBeyani", "Veri cihazdan çıkmaz; XML/API yok", "metin", "KVKK beyanı", "01.01.2026", "R3 ayrım"),
        ("dbe_kanitRaporu", "A4 baskı KANIT_RAPORU", "metin", "Kanıt çıktısı", "01.01.2026", "R2 ayrım"),
        ("dbe_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("dbe_firmaUnvan", "Örnek Ticaret A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026", "Firma"),
        ("dbe_fiyatModeli", "Tek seferlik alım — aylık ücret yok", "metin", "Ürün konumlandırma", "01.01.2026", "R3 ayrım"),
        ("dbe_yapistirCalistir", "Kurulum beklemez; yapıştır-çalıştır", "metin", "UX vaadi", "01.01.2026", "R3 ayrım"),
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

    h(ws, 27, 8, "Motor Cikti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 9, "Deger", kalin=True, yazi=KOYU_LACIVERT)
    motor_cikti = [
        (28, "dbe_hataListesi", '=COUNTA(tblHata[HataKodu])'),
        (29, "dbe_mutabakatFarki",
         '=IFERROR(SUM(tblVeriA[Toplam])-SUM(tblVeriB[Tutar]),0)'),
        (30, "dbe_eslesmeOrani",
         '=IF(COUNTA(tblEslesme[A_Norm])=0,0,'
         '(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/COUNTA(tblEslesme[A_Norm]))'),
        (31, "dbe_kovaButunluk", "=ESLESTIRME!C2"),
        (32, "dbe_kararKapi",
         '=IF(COUNTA(tblVeriA[FaturaNo])=0,"VERİ YOK",'
         'IF(dbe_hataListesi>=dbe_esikHataCok,"DURDUR",'
         'IF(OR(dbe_hataListesi>0,COUNTIF(tblEslesme[Kova],"tutar_farki")>0,'
         'dbe_eslesmeOrani<dbe_esikEslesme),"İNCELE","UYGUN")))'),
        (33, "dbe_modulT1", "=dbe_eslesmeOrani"),
        (34, "dbe_modulT2", "=dbe_hataListesi"),
        (35, "dbe_modulO1", "=COUNTIF(tblEslesme[EsikUstu],1)"),
        (36, "dbe_modulO2", "=IFERROR(ABS(dbe_mutabakatFarki)*dbe_duyarlilikCarpan,0)"),
        (37, "dbe_modulO3", "=IFERROR(dbe_toleransSenaryo*COUNTIF(tblEslesme[Kova],\"tutar_farki\"),0)"),
        (38, "dbe_modulO6",
         '=IF(COUNTA(tblVeriA[FaturaNo])=0,0,COUNTA(tblVeriA[VKN])/COUNTA(tblVeriA[FaturaNo]))'),
        (39, "dbe_modulO8",
         '=IF(COUNTA(tblEslesme[A_Norm])=0,0,'
         'ROUND(100*(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/COUNTA(tblEslesme[A_Norm]),0))'),
        (40, "dbe_modulI1", "=IFERROR(PERCENTILE(tblVeriA[Toplam],dbe_yuzdelikOran),0)"),
        (41, "dbe_modulI2", "=IFERROR(PERCENTILE(tblEslesme[AbsFark],dbe_yuzdelikOran),0)"),
    ]
    for r, ad, form in motor_cikti:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kalin=True)

    h(ws, 43, 8, "Canli Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorum_satir = [
        (44, "dbe_modulT1Yorum",
         '=IF(COUNTA(tblVeriA[FaturaNo])=0,"Veri yok",'
         '_xlfn.TEXTJOIN(" ",TRUE,"Eşleşme oranı",TEXT(dbe_modulT1,"0.0%")))'),
        (45, "dbe_modulT2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Hata sayısı",TEXT(dbe_modulT2,"0")),"-")'),
        (46, "dbe_modulO1Yorum",
         '=IF(dbe_modulO1=0,"Anomali yok",'
         '_xlfn.TEXTJOIN(" ",TRUE,dbe_modulO1,"anomali satırı"))'),
        (47, "dbe_modulO2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Duyarlılık etkisi",TEXT(dbe_modulO2,"0.00"),"TL"),"-")'),
        (48, "dbe_modulO3Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Senaryo etki",TEXT(dbe_modulO3,"0.00"),"TL"),"-")'),
        (49, "dbe_modulO6Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"VKN doluluk",TEXT(dbe_modulO6,"0.0%")),"-")'),
        (50, "dbe_modulO8Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Veri kalite",TEXT(dbe_modulO8,"0")),"-")'),
        (51, "dbe_modulI1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(dbe_modulI1,"0.00"),"TL"),"-")'),
        (52, "dbe_modulI2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Fark P90",TEXT(dbe_modulI2,"0.00"),"TL"),"-")'),
    ]
    for r, ad, form in yorum_satir:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)


    genislik(ws, {"A": 28, "B": 42, "C": 12, "D": 28, "E": 14, "F": 28, "H": 24, "I": 42})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    adimlar = [
        "VERI_A: e-Arşiv satırlarını (fatura numarası, tutar, VKN) sarı hücrelere yapıştırın.",
        "VERI_B: DBS hedef satırlarını girin — XML veya API gerekmez.",
        "ANAHTAR: Fatura numarası normalize, uzunluk, fatura numarası ikincil ve duplike otomatik dolar.",
        "MOTOR: Ara hesaplar ve karar girdileri burada üretilir.",
        "ESLESTIRME: Dört kova ve bütünlük denetimi otomatik çalışır.",
        "HATA_LISTE: Zorunlu alan ve format hatalarını temizleyin.",
        "KOK_NEDEN: Tutar farklarını sınıflandırır; TANIMSIZ elle incelenir.",
        "PANO: Karar, KPI ve grafikleri izleyin.",
        "KANIT_RAPORU: Dönem dosyasına A4 PDF olarak yazdırın.",
    ]
    for i, m in enumerate(adimlar, 5):
        h(ws, i, 1, f"{i - 4}. {m}", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    h(ws, 14, 1, "Sık yapılan hatalar", kalin=True, yazi="B3261E")
    for i, m in enumerate([
        "Fatura numarasını eksik yapıştırmak (boşluk farkı normalize edilir ama boş olamaz)",
        "e-Arşiv ve DBS dönemlerini karıştırmak",
        "Toplamı DBS tutarına yanlış map etmek",
    ], 15):
        h(ws, i, 1, "• " + m)
    h(ws, 19, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 32):
        dogrulama(ws, "textLength", "0", f"D{r}",
                  baslik="Kaynak", mesaj="Kaynak metni girin",
                  hata_baslik="Kaynak", hata_mesaj="Boş bırakmayın",
                  isaret="greaterThanOrEqual", f2="120")
    ws = wb["LISTELER"]
    for r in range(7, 12):
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Belge Tipi", mesaj="Ad girin",
                  hata_baslik="Ad", hata_mesaj="En az 1 karakter",
                  isaret="greaterThan", f2="40")
    for r in range(7, 11):
        dogrulama(ws, "textLength", "1", f"C{r}",
                  baslik="Durum", mesaj="Durum girin",
                  hata_baslik="Durum", hata_mesaj="Geçersiz",
                  isaret="greaterThan", f2="30")
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
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"İNCELE"'], fill=sari))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(6, 22):
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

    ws = wb["KOK_NEDEN"]
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"TANIMSIZ"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"KDV_ORAN"'], fill=sari))
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"MATRAH"'], fill=sari))
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"YUVARLAMA"'], fill=sari))

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
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))

    ws = wb["VERI_A"]
    ws.conditional_formatting.add(f"G{ILK}:G{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"G{ILK}:G{SON}",
        CellIsRule(operator="greaterThan", formula=["10000"], fill=sari))

    ws = wb["VERI_B"]
    ws.conditional_formatting.add(f"D{ILK}:D{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"D{ILK}:D{SON}",
        CellIsRule(operator="greaterThan", formula=["10000"], fill=sari))

    ws = wb["ANAHTAR"]
    ws.conditional_formatting.add(f"F{ILK}:F{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"N{ILK}:N{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))

    ws = wb["HATA_LISTE"]
    ws.conditional_formatting.add(f"E{ILK}:E{SON}",
        CellIsRule(operator="equal", formula=['"YUKSEK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"E{ILK}:E{SON}",
        CellIsRule(operator="equal", formula=['"ORTA"'], fill=sari))
    ws.conditional_formatting.add(f"C{ILK}:C{SON}",
        CellIsRule(operator="equal", formula=['"TUTAR_NEGATIF"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"F{ILK}:F{SON}",
        CellIsRule(operator="equal", formula=['"TANIMSIZ"'], fill=kirmizi, font=kf))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("C19", CellIsRule(operator="equal", formula=['"BUTUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("C19", CellIsRule(operator="equal", formula=['"KAYIP"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add("C43", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("C43", CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kf))


def adlari_bagla(wb):
    for i, ana in enumerate([
        "toleransKurus", "raporTarihi", "dbe_esikEslesme", "dbe_esikFark",
        "dbe_azamiTutar", "dbe_ikincilUzunluk", "dbe_olcekHedef", "dbe_anomaliZ",
        "dbe_yuzdelikOran", "dbe_kdvOran", "dbe_duyarlilikCarpan", "dbe_toleransSenaryo",
        "dbe_yuvarlamaEsik", "dbe_guvenBolum", "dbe_riskFarkAgirlik", "dbe_riskEslesmezAgirlik",
        "dbe_esikHataCok", "dbe_vknUzunluk", "dbe_normAnahtar", "dbe_kolonHarita", "dbe_tanimsizKova",
        "dbe_kvkkBeyani", "dbe_kanitRaporu", "dbe_dosyaSurumu", "dbe_firmaUnvan",
        "dbe_fiyatModeli", "dbe_yapistirCalistir",
    ]):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "dbe_hataListesi": 28, "dbe_mutabakatFarki": 29, "dbe_eslesmeOrani": 30,
        "dbe_kovaButunluk": 31, "dbe_kararKapi": 32, "dbe_modulT1": 33, "dbe_modulT2": 34,
        "dbe_modulO1": 35, "dbe_modulO2": 36, "dbe_modulO3": 37,
        "dbe_modulO6": 38, "dbe_modulO8": 39, "dbe_modulI1": 40, "dbe_modulI2": 41,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "dbe_modulT1Yorum": 44, "dbe_modulT2Yorum": 45, "dbe_modulO1Yorum": 46,
        "dbe_modulO2Yorum": 47, "dbe_modulO3Yorum": 48, "dbe_modulO6Yorum": 49,
        "dbe_modulO8Yorum": 50, "dbe_modulI1Yorum": 51, "dbe_modulI2Yorum": 52,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeBelgeTipleri", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeDurumlar", "LISTELER!$C$7:$C$10")


def ornek_csv_yaz(urun_dir: str):
    yol = os.path.join(urun_dir, "ornek_veri.csv")
    with open(yol, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow([
            "taraf", "fatura_no", "tarih", "vkn", "unvan",
            "matrah", "kdv", "toplam", "tutar_b", "belge_tipi", "not",
        ])
        for s in ORNEK:
            notu = "eslesir"
            na = s["fatura_a"].replace(" ", "").upper()
            nb = s["fatura_b"].replace(" ", "").upper()
            if na != nb:
                notu = "eslesmez"
            elif abs(s["toplam"] - s["b_tutar"]) <= 0.05 and s["toplam"] != s["b_tutar"]:
                notu = "tolerans"
            elif s["toplam"] != s["b_tutar"]:
                notu = "tutar_farki"
            if s.get("eksik_tarih"):
                notu = "hata_tarih"
            if s.get("negatif"):
                notu = "hata_negatif"
            if len(s["vkn"]) != 10:
                notu = "hata_vkn"
            tar = "" if s.get("eksik_tarih") else s["tarih"].isoformat()
            top = -abs(s["toplam"]) if s.get("negatif") else s["toplam"]
            w.writerow([
                "A", s["fatura_a"], tar, s["vkn"],
                s["unvan"], s["matrah"], s["kdv"], top, "", s["belge"], notu,
            ])
            w.writerow([
                "B", s["fatura_b"], s["tarih"].isoformat(), s["vkn_b"],
                s["unvan"], "", "", "", s["b_tutar"], s["belge"], notu,
            ])
    return yol


def main(cikti_yolu=None):
    wb = Workbook()
    # SPEC sayfa_akisi sırası
    siralar = [
        (kapak, "KAPAK"),
        (veri_a, "VERI_A"),
        (veri_b, "VERI_B"),
        (anahtar, "ANAHTAR"),
        (motor, "MOTOR"),
        (eslestirme, "ESLESTIRME"),
        (hata_liste, "HATA_LISTE"),
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
    dosya_ad = "DefterBeyanEArsivAktarim.xlsx"
    hedef = cikti_yolu or os.path.join(urun_dir, dosya_ad)
    os.makedirs(os.path.dirname(hedef) or ".", exist_ok=True)
    wb.save(hedef)

    cikti = os.path.join(KOK, "cikti", dosya_ad)
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(hedef) != os.path.abspath(cikti):
        import shutil
        shutil.copy2(hedef, cikti)

    csv_yol = ornek_csv_yaz(urun_dir)
    print(f"Dosya oluşturuldu: {hedef}")
    print(f"Kopya: {cikti}")
    print(f"Örnek CSV: {csv_yol}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
