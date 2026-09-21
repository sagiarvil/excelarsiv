#!/usr/bin/env python3
"""Toplu e-Fatura Satır Defteri ve PDF Kanıtı — A2 üretim betiği (manda v6, Ç15)."""

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
)

URUN_AD = "Toplu e-Fatura Satır Defteri ve PDF Kanıtı"
SURUM = "1.0.0"
RENK = "0B1F3A"
KAPASITE = 1000
DEMO = 40
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 14)

BELGE_TIPLERI = ["SATIS", "IADE", "ISTISNA", "TEVKIFAT", "DIGER"]
DURUMLAR = ["Acik", "Kapali", "Itiraz", "Beklemede"]


def _uuid(i: int, tireli: bool = True) -> str:
    """Demo UUID — tireli / tiresiz senaryo için."""
    ham = f"{i:08x}-{(1000 + i):04x}-4{(i % 1000):03x}-a{(i % 1000):03x}-{(i * 17):012x}"[:36]
    if len(ham) < 36:
        ham = (ham + "0" * 36)[:36]
    return ham if tireli else ham.replace("-", "").upper()


def _ornek_satirlar():
    satirlar = []
    for i in range(1, DEMO + 1):
        uuid_a = _uuid(i, tireli=True)
        if i in (3, 7, 11):  # kısa/uzun / tire farkı — normalize ile eşleşir
            uuid_b = _uuid(i, tireli=False)
        elif i == 15:
            uuid_b = _uuid(9000 + i, tireli=True)  # eşleşmez
        else:
            uuid_b = uuid_a
        fatura_no = f"EFT2026{100000 + i}"
        if i in (5, 9):
            fatura_a, fatura_b = fatura_no[:12], fatura_no
        else:
            fatura_a, fatura_b = fatura_no, fatura_no
        matrah = round(1000 + i * 47.5, 2)
        kdv = round(matrah * 0.20, 2)
        toplam = round(matrah + kdv, 2)
        if i % 8 == 0:  # tutar farkı
            b_tutar = round(toplam - 12.5, 2)
        elif i % 11 == 0:  # tolerans içinde
            b_tutar = round(toplam + 0.03, 2)
        elif i == 15:
            b_tutar = toplam
        else:
            b_tutar = toplam
        vkn = f"{1000000000 + i:010d}"
        tar = RAPOR_TARIHI - timedelta(days=(DEMO - i))
        satirlar.append({
            "uuid_a": uuid_a, "uuid_b": uuid_b,
            "fatura_a": fatura_a, "fatura_b": fatura_b,
            "tarih": tar, "vkn": vkn,
            "unvan": f"Ornek Firma {i} Ltd.",
            "matrah": matrah, "kdv": kdv, "toplam": toplam,
            "b_tutar": b_tutar,
            "belge": BELGE_TIPLERI[i % len(BELGE_TIPLERI)],
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


def _norm_uuid_formul(ham_ref: str) -> str:
    """UUID normalize: tire sil + UPPER + TRIM (Ç15 / SPEC birincil anahtar)."""
    return f'UPPER(TRIM(SUBSTITUTE({ham_ref},"-","")))'


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "e-Fatura satırlarını muhasebe kayıtlarıyla UUID ve fatura numarası üzerinden "
      "eşleştirir; dört kovaya ayırır; PDF baskıya hazır kanıt raporu üretir. "
      "XML dosya okuma, API ve makro yoktur — satır yapıştırırsınız.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "UUID normalize (tire sil, büyük harf) + fatura numarası ikincil anahtar",
        "Eşleşti / tolerans içinde / tutar farkı / eşleşmedi — dört kova bütünlüğü",
        "KDV, matrah, yuvarlama, VKN kök-neden kuralları + TANIMSIZ kova",
        "Tutar farkı satırları için kopyalanabilir itiraz metni",
        "KANIT_RAPORU A4 baskı alanı — vergi dairesi / denetçi dosyasına uygun",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "SMMM ve mali müşavirler",
        "Muhasebe ekipleri",
        "İşletme sahipleri — KDV beyanı öncesi mutabakat",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) VERI_A'ya e-Fatura satırlarını yapıştırın. 2) VERI_B'ye muhasebe "
      "satırlarını girin. 3) PANO'dan eşleşme oranı ve kararı izleyin. "
      "4) KANIT_RAPORU'nu PDF yazdırın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def veri_a(ws):
    sayfa_hazirla(ws, "VERI_A", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "e-Fatura satır yapıştır — sarı hücrelere girin (XML/API yok)", son_kolon=12)
    sutunlar = ["UUID", "FaturaNo", "Tarih", "VKN", "Unvan", "Matrah", "KDV",
                "Toplam", "BelgeTipi", "Donem", "Durum", "KayitDolu"]
    baslik_satiri(ws, HDR, sutunlar)
    form_a = {"KayitDolu": 'IF(tblVeriA[[#This Row],[UUID]]="","",1)'}
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["uuid_a"], baslik="UUID", mesaj="e-Fatura UUID (tireli veya tiresiz)")
        _sari(ws, r, 2, s["fatura_a"], baslik="Fatura No", mesaj="Fatura numarası")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Fatura tarihi")
        _sari(ws, r, 4, s["vkn"], baslik="VKN", mesaj="Vergi kimlik numarası")
        _sari(ws, r, 5, s["unvan"], baslik="Unvan", mesaj="Firma unvanı")
        _sari(ws, r, 6, s["matrah"], sayi=TL, baslik="Matrah", mesaj="KDV matrahı TL")
        _sari(ws, r, 7, s["kdv"], sayi=TL, baslik="KDV", mesaj="KDV tutarı TL")
        _sari(ws, r, 8, s["toplam"], sayi=TL, baslik="Toplam", mesaj="Genel toplam TL")
        _sari(ws, r, 9, s["belge"], baslik="Belge Tipi", mesaj="Liste seçimi")
        _sari(ws, r, 10, "2026-08", baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 11, "Acik", baslik="Durum", mesaj="Liste seçimi")
    for r in range(ILK + DEMO, SON + 1):
        for c in range(1, 12):
            _sari(ws, r, c, None,
                  sayi=TARİH if c == 3 else (TL if c in (6, 7, 8) else None),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriA", f"A{HDR}:{get_column_letter(12)}{SON}", sutunlar, formuller=form_a)
    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}", baslik="UUID",
              mesaj="UUID yapıştırın", hata_baslik="UUID", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="80")
    dogrulama(ws, "textLength", "1", f"B{ILK}:B{SON}", baslik="Fatura No",
              mesaj="Fatura numarası", hata_baslik="Fatura", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "date", "1", f"C{ILK}:C{SON}", baslik="Tarih", mesaj="Tarih seçin",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih girin", isaret="greaterThan")
    dogrulama(ws, "textLength", "1", f"D{ILK}:D{SON}", baslik="VKN",
              mesaj="VKN girin", hata_baslik="VKN", hata_mesaj="Geçersiz",
              isaret="greaterThan", f2="20")
    for c, tip in [(6, "decimal"), (7, "decimal"), (8, "decimal")]:
        dogrulama(ws, tip, "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=sutunlar[c - 1], mesaj="Tutar ≥ 0",
                  hata_baslik="Tutar", hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeBelgeTipleri", f"I{ILK}:I{SON}",
              baslik="Belge Tipi", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçim yapın")
    dogrulama(ws, "list", "ListeDurumlar", f"K{ILK}:K{SON}",
              baslik="Durum", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçim yapın")
    dogrulama(ws, "textLength", "0", f"E{ILK}:E{SON}", baslik="Unvan",
              mesaj="Unvan girin", hata_baslik="Unvan", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="120")
    dogrulama(ws, "textLength", "1", f"J{ILK}:J{SON}", baslik="Dönem",
              mesaj="YYYY-AA", hata_baslik="Dönem", hata_mesaj="Geçersiz",
              isaret="greaterThan", f2="20")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 13)})
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["E"].width = 22


def veri_b(ws):
    sayfa_hazirla(ws, "VERI_B", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Muhasebe satır tarafı — beklenen kayıt tutarı", son_kolon=8)
    sutunlar = ["UUID", "FaturaNo", "Tarih", "Tutar", "HesapKodu", "Aciklama",
                "Donem", "Durum", "KayitDolu"]
    baslik_satiri(ws, HDR, sutunlar)
    form_b = {"KayitDolu": 'IF(tblVeriB[[#This Row],[UUID]]="","",1)'}
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["uuid_b"], baslik="UUID", mesaj="Muhasebe UUID")
        _sari(ws, r, 2, s["fatura_b"], baslik="Fatura No", mesaj="Muhasebe fatura numarası")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Kayıt tarihi")
        _sari(ws, r, 4, s["b_tutar"], sayi=TL, baslik="Tutar", mesaj="Muhasebe tutarı TL")
        _sari(ws, r, 5, f"320.{i:03d}", baslik="Hesap Kodu", mesaj="Muhasebe hesap kodu")
        _sari(ws, r, 6, "e-Fatura kaydı", baslik="Açıklama", mesaj="Kısa açıklama")
        _sari(ws, r, 7, "2026-08", baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 8, "Acik", baslik="Durum", mesaj="Liste seçimi")
    for r in range(ILK + DEMO, SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None, sayi=TARİH if c == 3 else (TL if c == 4 else None),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriB", f"A{HDR}:I{SON}", sutunlar, formuller=form_b)
    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}", baslik="UUID",
              mesaj="UUID girin", hata_baslik="UUID", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="80")
    dogrulama(ws, "textLength", "1", f"B{ILK}:B{SON}", baslik="Fatura No",
              mesaj="Fatura numarası", hata_baslik="Fatura", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "date", "1", f"C{ILK}:C{SON}", baslik="Tarih", mesaj="Tarih",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih", isaret="greaterThan")
    dogrulama(ws, "decimal", "0", f"D{ILK}:D{SON}", baslik="Tutar", mesaj="Tutar ≥ 0",
              hata_baslik="Tutar", hata_mesaj="0 veya üzeri", isaret="greaterThanOrEqual")
    for c, ad in [(5, "Hesap Kodu"), (6, "Açıklama"), (7, "Dönem")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin", hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "list", "ListeDurumlar", f"H{ILK}:H{SON}", baslik="Durum",
              mesaj="Listeden seçin", hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 16 for i in range(1, 10)})
    ws.column_dimensions["A"].width = 38


def anahtar(ws):
    sayfa_hazirla(ws, "ANAHTAR", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "UUID normalize + uzunluk + fatura numarası ikincil + duplike (Ç7 / E01)", son_kolon=12)
    h(ws, 4, 1, "A Tarafı (e-Fatura)", kalin=True, yazi=KOYU_LACIVERT)
    sut_a = ["Ham", "Norm", "NormUzunluk", "IkincilAnahtar", "KaynakTaraf", "Duplike", "SatirNo"]
    baslik_satiri(ws, HDR, sut_a)
    form_a = {
        "Norm": _norm_uuid_formul("tblAnahtarA[[#This Row],[Ham]]"),
        "NormUzunluk": 'IF(tblAnahtarA[[#This Row],[Norm]]="",0,LEN(tblAnahtarA[[#This Row],[Norm]]))',
        "IkincilAnahtar": (
            'IF(tblAnahtarA[[#This Row],[Ham]]="","",'
            'UPPER(TRIM(SUBSTITUTE(IFERROR(INDEX(tblVeriA[FaturaNo],'
            'ROW()-ROW(tblVeriA[[#Headers]])),"")," ",""))))'
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

    h(ws, 4, 9, "B Tarafı (Muhasebe)", kalin=True, yazi=KOYU_LACIVERT)
    sut_b = ["HamB", "NormB", "NormUzunlukB", "IkincilB", "KaynakB", "DuplikeB", "SatirNoB"]
    baslik_satiri(ws, HDR, sut_b, basla=9)
    form_b = {
        "NormB": _norm_uuid_formul("tblAnahtarB[[#This Row],[HamB]]"),
        "NormUzunlukB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",0,LEN(tblAnahtarB[[#This Row],[NormB]]))',
        "IkincilB": (
            'IF(tblAnahtarB[[#This Row],[HamB]]="","",'
            'UPPER(TRIM(SUBSTITUTE(IFERROR(INDEX(tblVeriB[FaturaNo],'
            'ROW()-ROW(tblVeriB[[#Headers]])),"")," ",""))))'
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
        (1, "A kayit sayisi", '=COUNTA(tblVeriA[UUID])'),
        (2, "B kayit sayisi", '=COUNTA(tblVeriB[UUID])'),
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
        (24, "Yuzdelik P90", '=IFERROR(PERCENTILE(tblEslesme[AbsFark],efs_yuzdelikOran),0)'),
        (25, "Tahmin aralik", '=IFERROR(PERCENTILE(tblVeriA[Toplam],efs_yuzdelikOran),0)'),
        (26, "Tornado zirve", '=IFERROR(MAX(PANO!K25:K29),0)'),
        (27, "Veri hazirlik",
         '=IF(C6=0,0,C6/(C6+C7))'),
        (28, "Esik alti mi", '=IF(C18<efs_esikEslesme,1,0)'),
        (29, "Fark esik ustu", '=IF(C11>efs_esikFark,1,0)'),
        (30, "Negatif A", '=COUNTIF(tblVeriA[Toplam],"<0")'),
        (31, "Negatif B", '=COUNTIF(tblVeriB[Tutar],"<0")'),
        (32, "Azami asimi A", '=COUNTIF(tblVeriA[Toplam],">"&efs_azamiTutar)'),
        (33, "Azami asimi B", '=COUNTIF(tblVeriB[Tutar],">"&efs_azamiTutar)'),
        (34, "Ikincil eslesme",
         '=COUNTIF(tblEslesme[IkincilEslesme],"IKINCIL")'),
        (35, "Birincil eslesme",
         '=COUNTIF(tblEslesme[IkincilEslesme],"BIRINCIL")'),
        (36, "Net katki", '=IFERROR(SUM(tblEslesme[NetKatki]),0)'),
        (37, "Kalite skoru",
         '=IF(C17=0,0,ROUND(100*(C12+C13)/MAX(C17,1),0))'),
        (38, "Karar kodu",
         '=IF(C6=0,"VERİ YOK",'
         'IF(OR(C35>0,C36>0,C37>0,C38>0),"İNCELE",'
         'IF(C18<efs_esikEslesme,"DURDUR",'
         'IF(C14>0,"İNCELE","UYGUN"))))'),
        (39, "VKN yogunluk",
         '=IFERROR(IF(C6=0,0,MAX(COUNTIF(tblVeriA[VKN],tblVeriA[VKN]))/C6),0)'),
        (40, "Senaryo etki",
         '=IFERROR(efs_toleransSenaryo*C14,0)'),
        (41, "Duyarlilik tolerans",
         '=IFERROR(ABS(C10)*efs_duyarlilikCarpan,0)'),
        (42, "Guven skoru",
         '=IF(C6=0,0,ROUND(100*C18*(1-MIN(C22,1)/efs_guvenBolum),0))'),
        (43, "Risk skoru",
         '=IF(C6=0,0,ROUND(100*(1-C18)+efs_riskFarkAgirlik*C14+efs_riskEslesmezAgirlik*C15,0))'),
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
        33: '=IF(C18<efs_esikEslesme,1,0)',
        34: '=IF(C11>efs_esikFark,1,0)',
        42: '=IF(C17=0,0,ROUND(100*(C12+C13)/MAX(C17,1),0))',
        43: (
            '=IF(C6=0,"VERİ YOK",'
            'IF(OR(C35>0,C36>0,C37>0,C38>0),"İNCELE",'
            'IF(C18<efs_esikEslesme,"DURDUR",'
            'IF(C14>0,"İNCELE","UYGUN"))))'
        ),
        44: '=IFERROR(IF(C6=0,0,1/C6),0)',  # placeholder overwritten below
        45: '=IFERROR(efs_toleransSenaryo*C14,0)',
        46: '=IFERROR(ABS(C10)*efs_duyarlilikCarpan,0)',
        47: '=IF(C6=0,0,ROUND(100*C18*(1-MIN(C22,1)/efs_guvenBolum),0))',
        48: '=IF(C6=0,0,ROUND(100*(1-C18)+efs_riskFarkAgirlik*C14+efs_riskEslesmezAgirlik*C15,0))',
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
            'IF(tblEslesme[[#This Row],[AbsFark]]>efs_esikFark,1,0))'
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
            'IF(tblKok[[#This Row],[AbsFark]]<=efs_yuvarlamaEsik,1,0))'
        ),
        "VKN": (
            'IF(tblKok[[#This Row],[A_Norm]]="","",'
            'IFERROR(INDEX(tblVeriA[VKN],MATCH(tblKok[[#This Row],[A_Norm]],tblAnahtarA[Norm],0)),""))'
        ),
        "KokNeden": (
            'IF(OR(tblKok[[#This Row],[A_Norm]]="",tblKok[[#This Row],[Kova]]<>"tutar_farki"),"",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[KDV]])<=toleransKurus,"KDV_ORAN",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[Matrah]]*efs_kdvOran)<=toleransKurus,"MATRAH",'
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
    h(ws, 4, 2, "=efs_kararKapi", kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Eşleşme Oranı", "=efs_eslesmeOrani", YÜZDE),
        (7, "e-Fatura Toplam (A)", '=IFERROR(SUM(tblVeriA[Toplam]),0)', TL),
        (8, "Muhasebe Toplam (B)", '=IFERROR(SUM(tblVeriB[Tutar]),0)', TL),
        (9, "Mutabakat Farkı", "=efs_mutabakatFarki", TL),
        (10, "Eşleşti Adet", '=COUNTIF(tblEslesme[Kova],"eslesti")', CATI),
        (11, "Tolerans İçinde", '=COUNTIF(tblEslesme[Kova],"tolerans_icinde")', CATI),
        (12, "Tutar Farkı Adet", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (13, "Eşleşmedi Adet", '=COUNTIF(tblEslesme[Kova],"eslesmedi")', CATI),
        (14, "Kova Bütünlük", "=efs_kovaButunluk", None),
        (15, "Anomali Sayısı", "=efs_modulO1", CATI),
        (16, "Ort. Tutar Farkı", "=efs_modulT2", TL),
        (17, "Fark P90", "=efs_modulI2", TL),
        (18, "Tahmin Aralık", "=efs_modulI1", TL),
        (19, "VKN Yoğunlaşma", "=efs_modulO6", YÜZDE),
        (20, "Veri Kalite", "=efs_modulO8", CATI),
        (21, "Duplike Uyarı", '=COUNTIF(tblAnahtarA[Duplike],"EVET")', CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (7, "=efs_modulT1Yorum"),
        (8, "=efs_modulT2Yorum"),
        (9, "=efs_modulO1Yorum"),
        (10, "=efs_modulO2Yorum"),
        (11, "=efs_modulO3Yorum"),
        (12, "=efs_modulO6Yorum"),
        (13, "=efs_modulO8Yorum"),
        (14, "=efs_modulI1Yorum"),
        (15, "=efs_modulI2Yorum"),
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
            'IF(ABS(tblAksiyon[[#This Row],[Fark]])>efs_esikFark,"YUKSEK","NORMAL"))'
        ),
        "ItirazMetni": (
            'IF(tblAksiyon[[#This Row],[Kova]]<>"tutar_farki","",'
            '"UUID "&tblAksiyon[[#This Row],[A_Norm]]&" için fark "&'
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
    bant(ws, 3, "e-Fatura ↔ Muhasebe Dönem Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Eşleşme Oranı", "=efs_eslesmeOrani", YÜZDE),
        (10, "e-Fatura Toplam", "=SUM(tblVeriA[Toplam])", TL),
        (11, "Mutabakat Farkı", "=efs_mutabakatFarki", TL),
        (12, "Eşleşti", '=COUNTIF(tblEslesme[Kova],"eslesti")', CATI),
        (13, "Tutar Farkı", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (14, "Eşleşmedi", '=COUNTIF(tblEslesme[Kova],"eslesmedi")', CATI),
        (15, "Kova Bütünlük", "=efs_kovaButunluk", None),
        (16, "Kanıt Özeti", "=efs_kanitRaporu", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 18, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "Tolerans kuruşu AYARLAR'dan gelir. UUID UPPER/TRIM/SUBSTITUTE(-) ile normalize edilir. "
      "XML/API/makro yoktur. Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A19:H19")
    h(ws, 21, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 1, "1. Tutar farkı satırlarında AKSIYON itiraz metnini kullanın.")
    h(ws, 23, 1, "2. TANIMSIZ kök nedenleri elle inceleyin.")
    h(ws, 24, 1, "3. Duplike UUID'leri temizleyip yeniden eşleştirin.")
    h(ws, 26, 1, "KVKK", kalin=True)
    h(ws, 26, 2, "=efs_kvkkBeyani")
    # A4 dikey baskı (PDF kanıt)
    baski_hazirla(ws, "A1:H28", f"{URUN_AD} | Kanıt")
    ws.page_setup.orientation = "portrait"
    genislik(ws, {"A": 28, "B": 22, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "A satır sayısı", '=COUNTA(tblVeriA[UUID])'),
        (6, "B satır sayısı", '=COUNTA(tblVeriB[UUID])'),
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
      "ve eşleşmeyen kovalar kasıtlı üretilmiştir. Tireli/tiresiz UUID senaryosu satır "
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
    genislik(ws, {"A": 18, "C": 16, "E": 12})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Tek kaynak parametreler — kaynak ve yürürlük zorunlu (D10)", son_kolon=10)
    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("toleransKurus", 0.05, "TL", "İş kuralı — kuruş toleransı", "01.01.2026", "Eşleşme toleransı"),
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026", "Rapor tarihi (TODAY yok)"),
        ("efs_esikEslesme", 0.90, "oran", "İç politika", "01.01.2026", "Eşleşme oranı eşiği"),
        ("efs_esikFark", 100, "TL", "İç politika", "01.01.2026", "Fark tutarı eşiği"),
        ("efs_azamiTutar", 1000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul tutar"),
        ("efs_ikincilUzunluk", 12, "karakter", "Anahtar stratejisi SPEC", "01.01.2026", "Ikincil anahtar uzunluğu"),
        ("efs_olcekHedef", 50000, "satir", "Manda A2 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("efs_anomaliZ", 2.0, "z", "İstatistik kuralı", "01.01.2026", "Anomali Z eşiği"),
        ("efs_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("efs_kdvOran", 0.20, "oran", "KDV genel oran", "01.01.2026", "Varsayılan KDV"),
        ("efs_duyarlilikCarpan", 1.1, "carpan", "Duyarlılık senaryosu", "01.01.2026", "O2 çarpan"),
        ("efs_toleransSenaryo", 0.5, "TL", "Senaryo etkisi", "01.01.2026", "O3 senaryo"),
        ("efs_yuvarlamaEsik", 1, "TL", "Yuvarlama kök neden eşiği", "01.01.2026", "Abs fark ≤ eşik"),
        ("efs_guvenBolum", 10, "oran", "Güven skoru böleni", "01.01.2026", "MOTOR güven"),
        ("efs_riskFarkAgirlik", 10, "puan", "Risk — tutar farkı ağırlığı", "01.01.2026", "MOTOR risk"),
        ("efs_riskEslesmezAgirlik", 5, "puan", "Risk — eşleşmedi ağırlığı", "01.01.2026", "MOTOR risk"),
        ("efs_normAnahtar", "UPPER TRIM SUBSTITUTE tire", "metin", "Ç7 UUID anahtar", "01.01.2026", "Serbest alternatif"),
        ("efs_tanimsizKova", "TANIMSIZ görünür", "metin", "E04", "01.01.2026", "Serbest alternatif"),
        ("efs_kvkkBeyani", "Veri cihazdan çıkmaz; XML/API yok", "metin", "KVKK beyanı", "01.01.2026", "R3 ayrım"),
        ("efs_kanitRaporu", "A4 baskı KANIT_RAPORU", "metin", "Kanıt çıktısı", "01.01.2026", "R2 ayrım"),
        ("efs_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("efs_firmaUnvan", "Örnek Ticaret A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026", "Firma"),
        ("efs_fiyatModeli", "Tek seferlik alım — aylık ücret yok", "metin", "Ürün konumlandırma", "01.01.2026", "R3 ayrım"),
        ("efs_yapistirCalistir", "Kurulum beklemez; yapıştır-çalıştır", "metin", "UX vaadi", "01.01.2026", "R3 ayrım"),
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

    h(ws, 28, 8, "Motor Cikti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 28, 9, "Deger", kalin=True, yazi=KOYU_LACIVERT)
    motor_cikti = [
        (29, "efs_mutabakatFarki",
         '=IFERROR(SUM(tblVeriA[Toplam])-SUM(tblVeriB[Tutar]),0)'),
        (30, "efs_eslesmeOrani",
         '=IF(COUNTA(tblEslesme[A_Norm])=0,0,'
         '(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/COUNTA(tblEslesme[A_Norm]))'),
        (31, "efs_kovaButunluk", "=ESLESTIRME!C2"),
        (32, "efs_kararKapi",
         '=IF(COUNTA(tblVeriA[UUID])=0,"VERİ YOK",'
         'IF(OR(COUNTIF(tblVeriA[Toplam],"<0")>0,COUNTIF(tblVeriB[Tutar],"<0")>0),"İNCELE",'
         'IF(OR(COUNTIF(tblVeriA[Toplam],">"&efs_azamiTutar)>0,'
         'COUNTIF(tblVeriB[Tutar],">"&efs_azamiTutar)>0),"İNCELE",'
         'IF(efs_eslesmeOrani<efs_esikEslesme,"DURDUR",'
         'IF(COUNTIF(tblEslesme[Kova],"tutar_farki")>0,"İNCELE","UYGUN")))))'),
        (33, "efs_modulT1", "=efs_eslesmeOrani"),
        (34, "efs_modulT2",
         '=IFERROR(IF(COUNTIF(tblEslesme[Kova],"tutar_farki")=0,0,'
         'SUMIF(tblEslesme[Kova],"tutar_farki",tblEslesme[AbsFark])'
         '/COUNTIF(tblEslesme[Kova],"tutar_farki")),0)'),
        (35, "efs_modulO1", "=COUNTIF(tblEslesme[EsikUstu],1)"),
        (36, "efs_modulO2", "=IFERROR(ABS(efs_mutabakatFarki)*efs_duyarlilikCarpan,0)"),
        (37, "efs_modulO3", "=IFERROR(efs_toleransSenaryo*COUNTIF(tblEslesme[Kova],\"tutar_farki\"),0)"),
        (38, "efs_modulO6",
         '=IF(COUNTA(tblVeriA[UUID])=0,0,COUNTA(tblVeriA[VKN])/COUNTA(tblVeriA[UUID]))'),
        (39, "efs_modulO8",
         '=IF(COUNTA(tblEslesme[A_Norm])=0,0,'
         'ROUND(100*(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/COUNTA(tblEslesme[A_Norm]),0))'),
        (40, "efs_modulI1", "=IFERROR(PERCENTILE(tblVeriA[Toplam],efs_yuzdelikOran),0)"),
        (41, "efs_modulI2", "=IFERROR(PERCENTILE(tblEslesme[AbsFark],efs_yuzdelikOran),0)"),
    ]
    for r, ad, form in motor_cikti:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kalin=True)

    h(ws, 43, 8, "Canli Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorum_satir = [
        (44, "efs_modulT1Yorum",
         '=IF(COUNTA(tblVeriA[UUID])=0,"Veri yok",'
         '_xlfn.TEXTJOIN(" ",TRUE,"Eşleşme oranı",TEXT(efs_modulT1,"0.0%")))'),
        (45, "efs_modulT2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Ort. tutar farkı",TEXT(efs_modulT2,"0.00"),"TL"),"-")'),
        (46, "efs_modulO1Yorum",
         '=IF(efs_modulO1=0,"Anomali yok",'
         '_xlfn.TEXTJOIN(" ",TRUE,efs_modulO1,"anomali satırı"))'),
        (47, "efs_modulO2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Duyarlılık etkisi",TEXT(efs_modulO2,"0.00"),"TL"),"-")'),
        (48, "efs_modulO3Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Senaryo etki",TEXT(efs_modulO3,"0.00"),"TL"),"-")'),
        (49, "efs_modulO6Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"VKN doluluk",TEXT(efs_modulO6,"0.0%")),"-")'),
        (50, "efs_modulO8Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Veri kalite",TEXT(efs_modulO8,"0")),"-")'),
        (51, "efs_modulI1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(efs_modulI1,"0.00"),"TL"),"-")'),
        (52, "efs_modulI2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Fark P90",TEXT(efs_modulI2,"0.00"),"TL"),"-")'),
    ]
    for r, ad, form in yorum_satir:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 28, "B": 42, "C": 12, "D": 28, "E": 14, "F": 28, "H": 24, "I": 42})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    adimlar = [
        "VERI_A: e-Fatura satırlarını (UUID, fatura numarası, tutar) sarı hücrelere yapıştırın.",
        "VERI_B: Muhasebe satırlarını girin — XML veya API gerekmez.",
        "ANAHTAR: UUID normalize, uzunluk, fatura numarası ikincil ve duplike otomatik dolar.",
        "MOTOR: Ara hesaplar ve karar girdileri burada üretilir.",
        "ESLESTIRME: Dört kova ve bütünlük denetimi otomatik çalışır.",
        "KOK_NEDEN: Tutar farklarını sınıflandırır; TANIMSIZ elle incelenir.",
        "PANO: Karar, KPI ve grafikleri izleyin.",
        "KANIT_RAPORU: Dönem dosyasına A4 PDF olarak yazdırın.",
    ]
    for i, m in enumerate(adimlar, 5):
        h(ws, i, 1, f"{i - 4}. {m}", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    h(ws, 14, 1, "Sık yapılan hatalar", kalin=True, yazi="B3261E")
    for i, m in enumerate([
        "UUID'yi eksik yapıştırmak (tire farkı normalize edilir ama boş olamaz)",
        "A ve B dönemlerini karıştırmak",
        "Toplam yerine yalnızca matrahı B tarafına yazmak",
    ], 15):
        h(ws, i, 1, "• " + m)
    h(ws, 19, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 30):
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
    ws.conditional_formatting.add(f"H{ILK}:H{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"F{ILK}:F{SON}",
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

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("C19", CellIsRule(operator="equal", formula=['"BUTUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("C19", CellIsRule(operator="equal", formula=['"KAYIP"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add("C43", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("C43", CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kf))


def adlari_bagla(wb):
    for i, ana in enumerate([
        "toleransKurus", "raporTarihi", "efs_esikEslesme", "efs_esikFark",
        "efs_azamiTutar", "efs_ikincilUzunluk", "efs_olcekHedef", "efs_anomaliZ",
        "efs_yuzdelikOran", "efs_kdvOran", "efs_duyarlilikCarpan", "efs_toleransSenaryo",
        "efs_yuvarlamaEsik", "efs_guvenBolum", "efs_riskFarkAgirlik", "efs_riskEslesmezAgirlik",
        "efs_normAnahtar", "efs_tanimsizKova", "efs_kvkkBeyani", "efs_kanitRaporu",
        "efs_dosyaSurumu", "efs_firmaUnvan", "efs_fiyatModeli", "efs_yapistirCalistir",
    ]):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "efs_mutabakatFarki": 29, "efs_eslesmeOrani": 30, "efs_kovaButunluk": 31,
        "efs_kararKapi": 32, "efs_modulT1": 33, "efs_modulT2": 34,
        "efs_modulO1": 35, "efs_modulO2": 36, "efs_modulO3": 37,
        "efs_modulO6": 38, "efs_modulO8": 39, "efs_modulI1": 40, "efs_modulI2": 41,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "efs_modulT1Yorum": 44, "efs_modulT2Yorum": 45, "efs_modulO1Yorum": 46,
        "efs_modulO2Yorum": 47, "efs_modulO3Yorum": 48, "efs_modulO6Yorum": 49,
        "efs_modulO8Yorum": 50, "efs_modulI1Yorum": 51, "efs_modulI2Yorum": 52,
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
            "taraf", "uuid", "fatura_no", "tarih", "vkn", "unvan",
            "matrah", "kdv", "toplam", "tutar_b", "belge_tipi", "not",
        ])
        for s in ORNEK:
            notu = "eslesir"
            if s["uuid_a"].replace("-", "").upper() != s["uuid_b"].replace("-", "").upper():
                notu = "eslesmez"
            elif abs(s["toplam"] - s["b_tutar"]) <= 0.05 and s["toplam"] != s["b_tutar"]:
                notu = "tolerans"
            elif s["toplam"] != s["b_tutar"]:
                notu = "tutar_farki"
            w.writerow([
                "A", s["uuid_a"], s["fatura_a"], s["tarih"].isoformat(), s["vkn"],
                s["unvan"], s["matrah"], s["kdv"], s["toplam"], "", s["belge"], notu,
            ])
            w.writerow([
                "B", s["uuid_b"], s["fatura_b"], s["tarih"].isoformat(), s["vkn"],
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
    dosya_ad = "EFaturaSatirDefteriPdfKaniti.xlsx"
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
