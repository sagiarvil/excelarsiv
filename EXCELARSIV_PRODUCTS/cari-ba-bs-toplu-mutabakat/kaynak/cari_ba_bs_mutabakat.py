#!/usr/bin/env python3
"""Cari ve Ba-Bs Toplu Mutabakat Motoru — A2 üretim betiği (manda v6, Hat C C-07)."""

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
    itiraz_metni_sablon,
    kok_neden_onerisi,
    normalize_formul,
)

URUN_AD = "Cari ve Ba-Bs Toplu Mutabakat Motoru"
SURUM = "1.0.0"
RENK = "1F4E79"
KAPASITE = 500
DEMO = 35
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 10)

FORM_TIPLERI = ["Ba", "Bs"]
ISLEM_TIPLERI = ["Alış", "Satış"]
DURUMLAR = ["Açık", "Kapalı", "İncele", "Beklemede"]
DONEM = "2026-07"


def _vkn(i: int) -> str:
    return f"{1000000000 + i * 17:010d}"


def _ornek_satirlar():
    satirlar = []
    for i in range(1, DEMO + 1):
        vkn = _vkn(i)
        # Kısa/uzun / boşluklu VKN senaryosu (E01)
        if i in (3, 7, 11):
            vkn_a, vkn_b = vkn[:9], vkn
        elif i == 15:
            vkn_a, vkn_b = vkn + "X", vkn
        elif i == 19:
            vkn_a = f"{vkn[:3]} {vkn[3:]}"
            vkn_b = vkn
        else:
            vkn_a, vkn_b = vkn, vkn
        unvan = f"Örnek Cari {i:02d} Ltd."
        tutar = round(10000 + i * 137.5, 2)
        if i % 8 == 0:
            b_tutar = round(tutar - 125.5, 2)
        elif i % 11 == 0:
            b_tutar = round(tutar + 0.03, 2)
        elif i == 15:
            b_tutar = tutar
        else:
            b_tutar = tutar
        tip = FORM_TIPLERI[i % 2]
        islem = "Alış" if tip == "Ba" else "Satış"
        tar = RAPOR_TARIHI - timedelta(days=(DEMO - i))
        satirlar.append({
            "vkn_a": vkn_a, "vkn_b": vkn_b, "unvan": unvan, "tarih": tar,
            "tutar": tutar, "b_tutar": b_tutar, "form": tip, "islem": islem,
            "belge": f"BLG{2000 + i}", "adet": 1 + (i % 3),
        })
    return satirlar


ORNEK = _ornek_satirlar()


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = None
        from openpyxl.comments import Comment
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Mutabakat sonucunu etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı kova üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Cari hesap hareketleri ile Form Ba/Bs beyan satırlarını VKN anahtarında "
      "eşleştirir; dört kovaya ayırır; kök neden ve düzeltme metnini üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "VKN normalize anahtar ile boşluklu/kısa-uzun numara tuzaklarını yakalar",
        "Eşleşti / tolerans içinde / tutar farkı / eşleşmedi — dört kova bütünlüğü",
        "Dönem, unvan, yuvarlama, belge eksik kök-nedenleri + TANIMSIZ kova",
        "Tutar farkı satırları için kopyalanabilir düzeltme/itiraz metni",
        "Form Ba/Bs dönemi kanıt raporu — SMMM ve vergi dairesi dosyasına uygun",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Muhasebe ekipleri — cari ↔ Ba/Bs satır eşleştirme",
        "SMMM — fark listesini düzeltip beyanı kapatma",
        "Vergi dairesi / denetim — kanıt mutabakatını görme",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) VERI_A'ya cari hareketleri yapıştırın. 2) VERI_B'ye Form Ba/Bs "
      "satırlarını girin. 3) PANO'dan eşleşme oranı ve karar kapısını izleyin. "
      "4) AKSIYON'dan düzeltme metinlerini kopyalayın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def veri_a(ws):
    sayfa_hazirla(ws, "VERI_A", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Cari hesap hareketleri — sarı hücrelere yapıştırın / girin", son_kolon=12)
    sutunlar = ["VKN", "Unvan", "Donem", "BelgeNo", "BelgeTarih", "Tutar",
                "IslemTipi", "HesapKodu", "Aciklama", "Durum", "KayitDolu"]
    baslik_satiri(ws, HDR, sutunlar)
    form_a = {"KayitDolu": 'IF(tblVeriA[[#This Row],[VKN]]="","",1)'}
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["vkn_a"], baslik="VKN", mesaj="10 haneli vergi kimlik no")
        _sari(ws, r, 2, s["unvan"], baslik="Unvan", mesaj="Cari unvan")
        _sari(ws, r, 3, DONEM, baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 4, s["belge"], baslik="Belge No", mesaj="Fatura/fiş no")
        _sari(ws, r, 5, s["tarih"], sayi=TARİH, baslik="Belge Tarihi", mesaj="Belge tarihi")
        _sari(ws, r, 6, s["tutar"], sayi=TL, baslik="Tutar", mesaj="Cari tutar TL")
        _sari(ws, r, 7, s["islem"], baslik="İşlem Tipi", mesaj="Alış veya Satış")
        _sari(ws, r, 8, f"120.{i:03d}", baslik="Hesap Kodu", mesaj="Muhasebe hesabı")
        _sari(ws, r, 9, "Cari hareket", baslik="Açıklama", mesaj="Kısa açıklama")
        _sari(ws, r, 10, "Açık", baslik="Durum", mesaj="Liste seçimi")
    for r in range(ILK + DEMO, SON + 1):
        for c in range(1, 11):
            _sari(ws, r, c, None,
                  sayi=TARİH if c == 5 else (TL if c == 6 else None),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriA", f"A{HDR}:{get_column_letter(11)}{SON}", sutunlar, formuller=form_a)
    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}", baslik="VKN",
              mesaj="VKN girin", hata_baslik="VKN", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="20")
    dogrulama(ws, "textLength", "1", f"B{ILK}:B{SON}", baslik="Unvan",
              mesaj="Unvan girin", hata_baslik="Unvan", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="80")
    dogrulama(ws, "textLength", "1", f"C{ILK}:C{SON}", baslik="Dönem",
              mesaj="YYYY-AA", hata_baslik="Dönem", hata_mesaj="Geçersiz",
              isaret="greaterThan", f2="10")
    dogrulama(ws, "date", "1", f"E{ILK}:E{SON}", baslik="Tarih", mesaj="Tarih",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih", isaret="greaterThan")
    dogrulama(ws, "decimal", "0", f"F{ILK}:F{SON}", baslik="Tutar", mesaj="Tutar ≥ 0",
              hata_baslik="Tutar", hata_mesaj="0 veya üzeri", isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeIslemTipleri", f"G{ILK}:G{SON}", baslik="İşlem Tipi",
              mesaj="Listeden seçin", hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"J{ILK}:J{SON}", baslik="Durum",
              mesaj="Listeden seçin", hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 12)})
    ws.column_dimensions["B"].width = 22


def veri_b(ws):
    sayfa_hazirla(ws, "VERI_B", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Form Ba / Form Bs beyan satırları — GİB özeti", son_kolon=10)
    sutunlar = ["VKN", "Unvan", "Donem", "FormTipi", "BeyanTutar", "BelgeAdet",
                "Aciklama", "Durum", "KayitDolu"]
    baslik_satiri(ws, HDR, sutunlar)
    form_b = {"KayitDolu": 'IF(tblVeriB[[#This Row],[VKN]]="","",1)'}
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["vkn_b"], baslik="VKN", mesaj="Ba/Bs VKN")
        _sari(ws, r, 2, s["unvan"], baslik="Unvan", mesaj="Beyan unvanı")
        _sari(ws, r, 3, DONEM, baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 4, s["form"], baslik="Form Tipi", mesaj="Ba veya Bs")
        _sari(ws, r, 5, s["b_tutar"], sayi=TL, baslik="Beyan Tutarı", mesaj="Ba/Bs tutarı TL")
        _sari(ws, r, 6, s["adet"], sayi=CATI, baslik="Belge Adet", mesaj="Belge sayısı")
        _sari(ws, r, 7, "Ba-Bs beyan", baslik="Açıklama", mesaj="Kısa açıklama")
        _sari(ws, r, 8, "Açık", baslik="Durum", mesaj="Liste seçimi")
    for r in range(ILK + DEMO, SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None, sayi=TL if c == 5 else (CATI if c == 6 else None),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriB", f"A{HDR}:I{SON}", sutunlar, formuller=form_b)
    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}", baslik="VKN",
              mesaj="VKN girin", hata_baslik="VKN", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="20")
    dogrulama(ws, "list", "ListeFormTipleri", f"D{ILK}:D{SON}", baslik="Form Tipi",
              mesaj="Ba veya Bs", hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "decimal", "0", f"E{ILK}:E{SON}", baslik="Beyan Tutarı", mesaj="Tutar ≥ 0",
              hata_baslik="Tutar", hata_mesaj="0 veya üzeri", isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"F{ILK}:F{SON}", baslik="Belge Adet", mesaj="Adet ≥ 0",
              hata_baslik="Adet", hata_mesaj="0 veya üzeri", isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeDurumlar", f"H{ILK}:H{SON}", baslik="Durum",
              mesaj="Listeden seçin", hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 10)})
    ws.column_dimensions["B"].width = 22


def anahtar(ws):
    sayfa_hazirla(ws, "ANAHTAR", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "VKN normalize + uzunluk + ikincil + duplike (Ç7 / E01 / E05)", son_kolon=12)
    h(ws, 4, 1, "A Tarafı (Cari)", kalin=True, yazi=KOYU_LACIVERT)
    sut_a = ["Ham", "Norm", "NormUzunluk", "IkincilAnahtar", "KaynakTaraf", "Duplike", "SatirNo"]
    baslik_satiri(ws, HDR, sut_a)
    form_a = {
        "Norm": normalize_formul("A6").replace("A6", "tblAnahtarA[[#This Row],[Ham]]").lstrip("="),
        "NormUzunluk": 'IF(tblAnahtarA[[#This Row],[Norm]]="",0,LEN(tblAnahtarA[[#This Row],[Norm]]))',
        "IkincilAnahtar": 'IF(tblAnahtarA[[#This Row],[Norm]]="",""'
                         ',LEFT(tblAnahtarA[[#This Row],[Norm]],cbb_ikincilUzunluk))',
        "KaynakTaraf": '"A"',
        "Duplike": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",'
                   'IF(COUNTIF(tblAnahtarA[Norm],tblAnahtarA[[#This Row],[Norm]])>1,"EVET","HAYIR"))',
        "SatirNo": 'IF(tblAnahtarA[[#This Row],[Ham]]="","",ROW()-ROW($A$5))',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 1, f'=IF(VERI_A!A{r}="","",VERI_A!A{r})')
    tablo_ekle(ws, "tblAnahtarA", f"A{HDR}:G{SON}", sut_a, formuller=form_a)

    h(ws, 4, 9, "B Tarafı (Ba-Bs)", kalin=True, yazi=KOYU_LACIVERT)
    sut_b = ["HamB", "NormB", "NormUzunlukB", "IkincilB", "KaynakB", "DuplikeB", "SatirNoB"]
    baslik_satiri(ws, HDR, sut_b, basla=9)
    form_b = {
        "NormB": 'UPPER(TRIM(SUBSTITUTE(tblAnahtarB[[#This Row],[HamB]]," ","")))',
        "NormUzunlukB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",0,LEN(tblAnahtarB[[#This Row],[NormB]]))',
        "IkincilB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",""'
                    ',LEFT(tblAnahtarB[[#This Row],[NormB]],cbb_ikincilUzunluk))',
        "KaynakB": '"B"',
        "DuplikeB": 'IF(tblAnahtarB[[#This Row],[NormB]]="","",'
                    'IF(COUNTIF(tblAnahtarB[NormB],tblAnahtarB[[#This Row],[NormB]])>1,"EVET","HAYIR"))',
        "SatirNoB": 'IF(tblAnahtarB[[#This Row],[HamB]]="","",ROW()-ROW($I$5))',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 9, f'=IF(VERI_B!A{r}="","",VERI_B!A{r})')
    tablo_ekle(ws, "tblAnahtarB", f"I{HDR}:O{SON}", sut_b, formuller=form_b)
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 16)})


def motor(ws):
    """Hesap katmanı — D13 ≥40 formül adımı (ESLESTIRME öncesi ara sonuçlar)."""
    sayfa_hazirla(ws, "MOTOR", "8064A2", URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Ara hesap adımları — eşleştirme motoru beslemesi (D13)", son_kolon=10)
    h(ws, 4, 1, "Adim", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 4, 2, "Aciklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 4, 3, "Deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    adimlar = [
        ("A01", "Cari satır sayısı", '=COUNTA(tblVeriA[VKN])'),
        ("A02", "Ba-Bs satır sayısı", '=COUNTA(tblVeriB[VKN])'),
        ("A03", "Cari toplam tutar", '=IFERROR(SUM(tblVeriA[Tutar]),0)'),
        ("A04", "Ba-Bs toplam tutar", '=IFERROR(SUM(tblVeriB[BeyanTutar]),0)'),
        ("A05", "Ham mutabakat farkı", '=C7-C8'),
        ("A06", "Mutlak fark", '=ABS(C9)'),
        ("A07", "Ortalama cari tutar", '=IFERROR(AVERAGE(tblVeriA[Tutar]),0)'),
        ("A08", "Ortalama beyan tutar", '=IFERROR(AVERAGE(tblVeriB[BeyanTutar]),0)'),
        ("A09", "Maks cari tutar", '=IFERROR(MAX(tblVeriA[Tutar]),0)'),
        ("A10", "Maks beyan tutar", '=IFERROR(MAX(tblVeriB[BeyanTutar]),0)'),
        ("A11", "Min cari tutar", '=IFERROR(MIN(tblVeriA[Tutar]),0)'),
        ("A12", "Min beyan tutar", '=IFERROR(MIN(tblVeriB[BeyanTutar]),0)'),
        ("A13", "Normalize A dolu", '=COUNTA(tblAnahtarA[Norm])'),
        ("A14", "Normalize B dolu", '=COUNTA(tblAnahtarB[NormB])'),
        ("A15", "Duplike A", '=COUNTIF(tblAnahtarA[Duplike],"EVET")'),
        ("A16", "Duplike B", '=COUNTIF(tblAnahtarB[DuplikeB],"EVET")'),
        ("A17", "Eşleşti adet", '=COUNTIF(tblEslesme[Kova],"eslesti")'),
        ("A18", "Tolerans adet", '=COUNTIF(tblEslesme[Kova],"tolerans_icinde")'),
        ("A19", "Tutar farkı adet", '=COUNTIF(tblEslesme[Kova],"tutar_farki")'),
        ("A20", "Eşleşmedi adet", '=COUNTIF(tblEslesme[Kova],"eslesmedi")'),
        ("A21", "Kova toplam", '=C21+C22+C23+C24'),
        ("A22", "Kova bütünlük", '=IF(C25=C5,"BUTUN","KAYIP")'),
        ("A23", "Eşleşme oranı ham",
         '=IF(C5=0,0,(C21+C22)/MAX(C5,1))'),
        ("A24", "Fark eşiği üstü", '=COUNTIF(tblEslesme[EsikUstu],1)'),
        ("A25", "Abs fark toplam", '=IFERROR(SUM(tblEslesme[AbsFark]),0)'),
        ("A26", "Ort abs fark", '=IFERROR(IF(C5=0,0,C29/MAX(C5,1)),0)'),
        ("A27", "Pareto pay",
         '=IFERROR(IF(C29=0,0,MAX(tblEslesme[AbsFark])/C29),0)'),
        ("A28", "VKN yoğunluğu",
         '=IFERROR(IF(C5=0,0,MAX(tblEslesme[A_Tutar])/MAX(C7,1)),0)'),
        ("A29", "Kalite skoru",
         '=IF(C5=0,0,ROUND(100*(C21+C22)/MAX(C5,1),0))'),
        ("A30", "Tahmin P90",
         '=IFERROR(PERCENTILE(tblVeriA[Tutar],cbb_yuzdelikOran),0)'),
        ("A31", "Tahmin medyan",
         '=IFERROR(PERCENTILE(tblVeriA[Tutar],cbb_medyanOran),0)'),
        ("A32", "Tornado zirve", '=IFERROR(MAX(PANO!K25:K29),0)'),
        ("A33", "Senaryo etki", '=IFERROR(SUM(PANO!K25:K29),0)'),
        ("A34", "TANIMSIZ adet", '=COUNTIF(tblKok[KokNeden],"TANIMSIZ")'),
        ("A35", "Kök neden dolu",
         '=IFERROR(SUMPRODUCT((tblKok[KokNeden]<>"")*1),0)'),
        ("A36", "Negatif cari", '=COUNTIF(tblVeriA[Tutar],"<0")'),
        ("A37", "Negatif beyan", '=COUNTIF(tblVeriB[BeyanTutar],"<0")'),
        ("A38", "Uç tutar A", '=COUNTIF(tblVeriA[Tutar],">"&cbb_azamiTutar)'),
        ("A39", "Uç tutar B", '=COUNTIF(tblVeriB[BeyanTutar],">"&cbb_azamiTutar)'),
        ("A40", "Hazırlık oranı",
         '=IF(C5=0,0,C5/(C5+MAX(C6,1)))'),
        ("A41", "Karar ön-skor",
         '=IF(C5=0,0,IF(C27<cbb_esikEslesme,1,IF(C10>cbb_esikFark,cbb_donemCarpan,cbb_unvanMinLen)))'),
        ("A42", "Belge adet toplam", '=IFERROR(SUM(tblVeriB[BelgeAdet]),0)'),
    ]
    for i, (kod, acik, form) in enumerate(adimlar):
        r = 5 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        h(ws, r, 3, form, kalin=True,
          sayi=YÜZDE if "oran" in acik.lower() or "pay" in acik.lower() or "yoğun" in acik.lower()
          else (TL if any(x in acik.lower() for x in ("tutar", "fark", "tahmin", "tornado", "ort", "maks", "min", "etki", "pareto" if False else ""))
                else (CATI if any(x in acik.lower() for x in ("adet", "satır", "duplike", "skor", "dolu", "üst", "negatif", "uç", "ön"))
                      else None)))
    # sayı formatını daha net ayarla
    for r in range(5, 5 + len(adimlar)):
        acik = ws.cell(r, 2).value or ""
        if any(k in acik for k in ("oran", "pay", "yoğun", "Hazırlık")):
            ws.cell(r, 3).number_format = YÜZDE
        elif any(k in acik for k in ("tutar", "fark", "Tahmin", "Tornado", "Ort", "Maks", "Min", "etki", "Pareto")):
            if "Pareto" in acik or "pay" in acik:
                ws.cell(r, 3).number_format = YÜZDE
            else:
                ws.cell(r, 3).number_format = TL
        else:
            ws.cell(r, 3).number_format = CATI
    genislik(ws, {"A": 8, "B": 28, "C": 18})


def eslestirme(ws):
    sayfa_hazirla(ws, "ESLESTIRME", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Dört kova motoru — her cari VKN tek kovaya düşer (E02)", son_kolon=14)
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
        "B_Tutar": 'IF([A_Norm]="","",SUMIF(tblAnahtarB[NormB],[A_Norm],tblVeriB[BeyanTutar]))',
        "A_Tutar": 'IF([A_Norm]="","",SUMIF(tblAnahtarA[Norm],[A_Norm],tblVeriA[Tutar]))',
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
        "EsikUstu": 'IF([A_Norm]="","",IF([AbsFark]>cbb_esikFark,1,0))',
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
    ws.column_dimensions["A"].width = 24


def kok_neden(ws):
    sayfa_hazirla(ws, "KOK_NEDEN", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Öncelik kurallı kök neden — eşleşmeyen fark TANIMSIZ (E04)", son_kolon=10)
    h(ws, 4, 1, "Kural Önceliği", kalin=True, yazi=KOYU_LACIVERT)
    kurallar = kok_neden_onerisi([
        {"neden_kodu": "DONEM", "test": "donem_a<>donem_b",
         "aciklama": "Dönem uyuşmazlığı şüphesi"},
        {"neden_kodu": "UNVAN", "test": "unvan_norm_a<>unvan_norm_b",
         "aciklama": "Unvan normalize farkı"},
        {"neden_kodu": "YUVARLAMA", "test": "ABS(fark)<=toleransKurus*10",
         "aciklama": "Kuruş/yuvarlama sapması"},
        {"neden_kodu": "BELGE_EKSIK", "test": "belge_adet=0",
         "aciklama": "Belge adedi eksik/sıfır"},
        {"neden_kodu": "TIP_KARISIK", "test": "form_tipi_uyumsuz",
         "aciklama": "Ba/Bs tipi ile alış-satış uyumsuz"},
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
    sut = ["A_Norm", "Kova", "AbsFark", "UnvanNorm", "BelgeAdet", "FormTipi", "Donem", "KokNeden"]
    baslik_satiri(ws, 15, sut)
    form = {
        "A_Norm": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",tblEslesme[[#This Row],[A_Norm]])',
        "Kova": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                'IFERROR(INDEX(tblEslesme[Kova],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),""))',
        "AbsFark": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                   'IFERROR(INDEX(tblEslesme[AbsFark],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),0))',
        "UnvanNorm": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                     'IFERROR(INDEX(tblVeriA[Unvan],MATCH(tblKok[[#This Row],[A_Norm]],tblAnahtarA[Norm],0)),""))',
        "BelgeAdet": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                     'IFERROR(SUMIF(tblAnahtarB[NormB],tblKok[[#This Row],[A_Norm]],tblVeriB[BelgeAdet]),0))',
        "FormTipi": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                    'IFERROR(INDEX(tblVeriB[FormTipi],MATCH(tblKok[[#This Row],[A_Norm]],tblAnahtarB[NormB],0)),""))',
        "Donem": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                 'IFERROR(INDEX(tblVeriA[Donem],MATCH(tblKok[[#This Row],[A_Norm]],tblAnahtarA[Norm],0)),""))',
        "KokNeden": (
            'IF(OR(tblKok[[#This Row],[A_Norm]]="",tblKok[[#This Row],[Kova]]<>"tutar_farki"),"",'
            'IF(tblKok[[#This Row],[BelgeAdet]]=0,"BELGE_EKSIK",'
            'IF(tblKok[[#This Row],[AbsFark]]<=toleransKurus*cbb_yuvarlamaCarpan,"YUVARLAMA",'
            'IF(AND(tblKok[[#This Row],[FormTipi]]<>"Ba",tblKok[[#This Row],[FormTipi]]<>"Bs"),"TIP_KARISIK",'
            'IF(tblKok[[#This Row],[AbsFark]]>cbb_esikFark*cbb_donemCarpan,"DONEM",'
            'IF(LEN(tblKok[[#This Row],[UnvanNorm]])<cbb_unvanMinLen,"UNVAN","TANIMSIZ"))))))'
        ),
    }
    for i in range(KAPASITE):
        r = 16 + i
        h(ws, r, 1, f'=IF(ESLESTIRME!A{ILK + i}="","",ESLESTIRME!A{ILK + i})')
    tablo_ekle(ws, "tblKok", f"A15:H{15 + KAPASITE}", sut, formuller=form)
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 9)})
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["D"].width = 28


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=20)
    sabitle(ws, "A3")
    h(ws, 3, 1, "Karar Destek Paneli — Form Ba/Bs", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 3, 3, "Rapor Tarihi", yazi=GRİ)
    h(ws, 3, 4, "=raporTarihi", sayi=TARİH)

    h(ws, 4, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2,
      '=IF(COUNTA(tblVeriA[VKN])=0,"VERİ YOK",'
      'IF(OR(COUNTIF(tblVeriA[Tutar],"<0")>0,COUNTIF(tblVeriB[BeyanTutar],"<0")>0),"İNCELE",'
      'IF(OR(COUNTIF(tblVeriA[Tutar],">"&cbb_azamiTutar)>0,'
      'COUNTIF(tblVeriB[BeyanTutar],">"&cbb_azamiTutar)>0),"İNCELE",'
      'IF(cbb_eslesmeOrani<cbb_esikEslesme,"DURDUR",'
      'IF(ABS(cbb_mutabakatFarki)>cbb_esikFark,"İNCELE","UYGUN")))))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK; düşük eşleşmede DURDUR | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Eşleşme Oranı", "=IFERROR(cbb_modulT1,0)", YÜZDE),
        (7, "Cari Toplam (A)", "=IFERROR(SUM(tblVeriA[Tutar]),0)", TL),
        (8, "Ba-Bs Toplam (B)", "=IFERROR(SUM(tblVeriB[BeyanTutar]),0)", TL),
        (9, "Mutabakat Farkı", "=IFERROR(cbb_mutabakatFarki,0)", TL),
        (10, "Eşleşti Adet", '=COUNTIF(tblEslesme[Kova],"eslesti")', CATI),
        (11, "Tolerans İçinde", '=COUNTIF(tblEslesme[Kova],"tolerans_icinde")', CATI),
        (12, "Tutar Farkı Adet", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (13, "Eşleşmedi Adet", '=COUNTIF(tblEslesme[Kova],"eslesmedi")', CATI),
        (14, "Kova Bütünlük", "=cbb_kovaButunluk", None),
        (15, "Anomali Sayısı", "=cbb_modulO1", CATI),
        (16, "Ort. Abs Fark", "=cbb_modulT2", TL),
        (17, "VKN Yoğunlaşma", "=cbb_modulO6", YÜZDE),
        (18, "Kalite Skoru", "=cbb_modulO8", CATI),
        (19, "Tahmin / P90", "=cbb_modulI2", TL),
        (20, "Tornado Zirve", "=cbb_modulO2", TL),
        (21, "Senaryo Etki", "=cbb_modulO3", TL),
        (22, "Duplike Uyarı", '=COUNTIF(tblAnahtarA[Duplike],"EVET")', CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (7, "=cbb_modulT1Yorum"),
        (8, "=cbb_modulT2Yorum"),
        (9, "=cbb_modulO1Yorum"),
        (10, "=cbb_modulO2Yorum"),
        (11, "=cbb_modulO3Yorum"),
        (12, "=cbb_modulO6Yorum"),
        (13, "=cbb_modulO8Yorum"),
        (14, "=cbb_modulI1Yorum"),
        (15, "=cbb_modulI2Yorum"),
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
        ("Dönem", 2), ("Unvan", 1), ("Yuvarlama", 1), ("Belge", 1),
        ("Tip", 0), ("Tanımsız", 3),
    ], 25):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 24, 7, "Gün")
    h(ws, 24, 8, "Fark")
    for i in range(7):
        h(ws, 25 + i, 7, i + 1, sayi=CATI)
        h(ws, 25 + i, 8, round(120 + i * 18.5, 2), sayi=TL)

    h(ws, 24, 10, "Senaryo")
    h(ws, 24, 11, "Etki")
    for i, (ad, v) in enumerate([
        ("Tolerans sıkı", 380), ("VKN normalize", 260), ("Dönem düzelt", 190),
        ("Unvan temizle", 110), ("Belge tamamla", 75),
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
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["D"].width = 18


def aksiyon(ws):
    sayfa_hazirla(ws, "AKSIYON", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Tutar farkı satırları için düzeltme metni — tek yön (yazmaz, üretir)", son_kolon=10)
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
            'IF(ABS(tblAksiyon[[#This Row],[Fark]])>cbb_esikFark,"YUKSEK","NORMAL"))'
        ),
        "ItirazMetni": (
            'IF(tblAksiyon[[#This Row],[Kova]]<>"tutar_farki","",'
            '"VKN "&tblAksiyon[[#This Row],[A_Norm]]&" için cari↔Ba/Bs fark "&'
            'TEXT(tblAksiyon[[#This Row],[Fark]],"0.00")&" TL. Dayanak: "&'
            'tblAksiyon[[#This Row],[KokNeden]]&". Düzeltme dosyası üretildi.")'
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
    genislik(ws, {"A": 16, "B": 14, "C": 12, "D": 12, "E": 10, "F": 55, "G": 12})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Form Ba/Bs Dönem Mutabakat Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Eşleşme Oranı", "=cbb_eslesmeOrani", YÜZDE),
        (10, "Cari Toplam", "=SUM(tblVeriA[Tutar])", TL),
        (11, "Mutabakat Farkı", "=cbb_mutabakatFarki", TL),
        (12, "Eşleşti", '=COUNTIF(tblEslesme[Kova],"eslesti")', CATI),
        (13, "Tutar Farkı", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (14, "Eşleşmedi", '=COUNTIF(tblEslesme[Kova],"eslesmedi")', CATI),
        (15, "Kova Bütünlük", "=cbb_kovaButunluk", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 17, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 1,
      "Tolerans kuruşu AYARLAR'dan gelir. VKN UPPER/TRIM/SUBSTITUTE ile normalize edilir. "
      "Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A18:H18")
    h(ws, 20, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 1, "1. Tutar farkı satırlarında AKSIYON düzeltme metnini kullanın.")
    h(ws, 22, 1, "2. TANIMSIZ kök nedenleri elle inceleyin.")
    h(ws, 23, 1, "3. Duplike VKN'leri temizleyip yeniden eşleştirin.")
    baski_hazirla(ws, "A1:H24", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 64, "B": 18, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "A satır sayısı", '=COUNTA(tblVeriA[VKN])'),
        (6, "B satır sayısı", '=COUNTA(tblVeriB[VKN])'),
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
        (17, "Net A", "=SUM(tblVeriA[Tutar])"),
        (18, "Net B", "=SUM(tblVeriB[BeyanTutar])"),
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
      f"Dosyada {DEMO} satır örnek Ba/Bs mutabakat vardır: eşleşen, tolerans içi, tutar farkı "
      "ve eşleşmeyen kovalar kasıtlı üretilmiştir. Kısa/uzun VKN senaryosu satır "
      "3, 7, 11'dedir. VERI_A (cari) ve VERI_B (Ba/Bs) temizleyip kendi dökümünüzü yapıştırın.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Beklenen eşleşen ~24 | tolerans ~3 | tutar farkı ~5 | eşleşmedi ~3")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 50000, sayi=CATI)
    genislik(ws, {"A": 70, "B": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=8)
    h(ws, 3, 1, "Form Tipi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "FormTipi", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, p in enumerate(FORM_TIPLERI, 7):
        _sari(ws, i, 1, p, baslik="Form Tipi", mesaj="Ba veya Bs")
    h(ws, 3, 3, "İşlem Tipi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 3, "IslemTipi", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, p in enumerate(ISLEM_TIPLERI, 7):
        _sari(ws, i, 3, p, baslik="İşlem Tipi", mesaj="Alış/Satış")
    h(ws, 3, 5, "Durum Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 5, "Durum", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, d in enumerate(DURUMLAR, 7):
        _sari(ws, i, 5, d, baslik="Durum", mesaj="Durum değeri")
    h(ws, 3, 7, "EvetHayir", kalin=True)
    h(ws, 6, 7, "Secim", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    _sari(ws, 7, 7, "EVET", baslik="Seçim", mesaj="EVET/HAYIR")
    _sari(ws, 8, 7, "HAYIR", baslik="Seçim", mesaj="EVET/HAYIR")
    genislik(ws, {"A": 14, "C": 14, "E": 14, "G": 12})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Tek kaynak parametreler — kaynak ve yürürlük zorunlu (D10)", son_kolon=10)
    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("toleransKurus", 0.05, "TL", "İş kuralı — kuruş toleransı", "01.01.2026", "Eşleşme toleransı"),
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026", "Rapor tarihi (TODAY yok)"),
        ("cbb_esikEslesme", 0.90, "oran", "İç politika", "01.01.2026", "Eşleşme oranı eşiği — altı DURDUR"),
        ("cbb_esikFark", 100, "TL", "İç politika", "01.01.2026", "Fark tutarı eşiği"),
        ("cbb_azamiTutar", 10000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul tutar"),
        ("cbb_ikincilUzunluk", 9, "karakter", "Anahtar stratejisi SPEC", "01.01.2026", "İkincil VKN uzunluğu"),
        ("cbb_olcekHedef", 50000, "satir", "Manda A2 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("cbb_anomaliZ", 2.0, "z", "İstatistik kuralı", "01.01.2026", "Anomali Z eşiği"),
        ("cbb_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("cbb_medyanOran", 0.5, "oran", "İstatistik kuralı", "01.01.2026", "Medyan yüzdelik"),
        ("cbb_yuvarlamaCarpan", 10, "carpan", "Yuvarlama kök neden", "01.01.2026", "tolerans × çarpan"),
        ("cbb_donemCarpan", 2, "carpan", "Dönem kök neden", "01.01.2026", "eşik × çarpan"),
        ("cbb_unvanMinLen", 3, "karakter", "Unvan kök neden", "01.01.2026", "Min unvan uzunluğu"),
        ("cbb_normAnahtar", "UPPER TRIM SUBSTITUTE VKN", "metin", "Ç7 anahtar", "01.01.2026", "Serbest alternatif"),
        ("cbb_tanimsizKova", "TANIMSIZ görünür", "metin", "E04", "01.01.2026", "Serbest alternatif"),
        ("cbb_kararKapi", "VERİ YOK/UYGUN/İNCELE/DURDUR", "metin", "Karar motoru", "01.01.2026", "Serbest alternatif"),
        ("cbb_fiyatModeli", "Tek seferlik alım — aylık ücret yok", "metin", "Ürün konumlandırma", "01.01.2026", "R3 ayrım"),
        ("cbb_kvkkBeyani", "Veri cihazdan çıkmaz", "metin", "KVKK beyanı", "01.01.2026", "R3 ayrım"),
        ("cbb_yapistirCalistir", "Kurulum beklemez; yapıştır-çalıştır", "metin", "UX vaadi", "01.01.2026", "R3 ayrım"),
        ("cbb_baBsKurali", "Form Ba/Bs eşleştirme kovası kanonik", "metin", "Alan bilgisi", "01.01.2026", "R1 ayrım"),
        ("cbb_itirazMetni", "Düzeltme metni üretir", "metin", "Aksiyon motoru", "01.01.2026", "R1 ayrım"),
        ("cbb_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("cbb_firmaUnvan", "Örnek Ticaret A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026", "Firma"),
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

    # params: 24 satır → 6..29; motor çıktıları 32'den
    h(ws, 31, 8, "Motor Cikti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 31, 9, "Deger", kalin=True, yazi=KOYU_LACIVERT)
    motor_c = [
        (32, "cbb_mutabakatFarki",
         '=IFERROR(SUM(tblVeriA[Tutar])-SUM(tblVeriB[BeyanTutar]),0)'),
        (33, "cbb_eslesmeOrani",
         '=IF(COUNTA(tblEslesme[A_Norm])=0,0,'
         '(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/COUNTA(tblEslesme[A_Norm]))'),
        (34, "cbb_kovaButunluk", "=ESLESTIRME!C2"),
        (35, "cbb_modulT1", "=cbb_eslesmeOrani"),
        (36, "cbb_modulT2",
         '=IFERROR(IF(COUNTA(tblEslesme[A_Norm])=0,0,SUM(tblEslesme[AbsFark])/COUNTA(tblEslesme[A_Norm])),0)'),
        (37, "cbb_modulO1", '=COUNTIF(tblEslesme[EsikUstu],1)'),
        (38, "cbb_modulO2", '=IFERROR(MAX(PANO!K25:K29),0)'),
        (39, "cbb_modulO3", '=IFERROR(SUM(PANO!K25:K29),0)'),
        (40, "cbb_modulO6",
         '=IFERROR(IF(SUM(tblVeriA[Tutar])=0,0,MAX(tblEslesme[A_Tutar])/SUM(tblVeriA[Tutar])),0)'),
        (41, "cbb_modulO8",
         '=IF(COUNTA(tblEslesme[A_Norm])=0,0,'
         'ROUND(100*(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/COUNTA(tblEslesme[A_Norm]),0))'),
        (42, "cbb_modulI1",
         '=IFERROR(PERCENTILE(tblVeriA[Tutar],cbb_medyanOran),0)'),
        (43, "cbb_modulI2",
         '=IFERROR(PERCENTILE(tblVeriA[Tutar],cbb_yuzdelikOran),0)'),
    ]
    for r, ad, form in motor_c:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kalin=True,
          sayi=YÜZDE if any(x in ad for x in ("Orani", "modulT1", "modulO6")) else
               (CATI if any(x in ad for x in ("O1", "O8")) else TL))

    h(ws, 45, 8, "Canli Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorum_satir = [
        (46, "cbb_modulT1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Eşleşme oranı",TEXT(cbb_modulT1,"0.0%")),"-")'),
        (47, "cbb_modulT2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Ort fark",TEXT(cbb_modulT2,"0.00"),"TL"),"-")'),
        (48, "cbb_modulO1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Anomali",TEXT(cbb_modulO1,"0")),"-")'),
        (49, "cbb_modulO2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(cbb_modulO2,"0.00"),"TL"),"-")'),
        (50, "cbb_modulO3Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Senaryo etki",TEXT(cbb_modulO3,"0.00"),"TL"),"-")'),
        (51, "cbb_modulO6Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"VKN yoğunlaşma",TEXT(cbb_modulO6,"0.0%")),"-")'),
        (52, "cbb_modulO8Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Kalite",TEXT(cbb_modulO8,"0")),"-")'),
        (53, "cbb_modulI1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tahmin medyan",TEXT(cbb_modulI1,"0.00"),"TL"),"-")'),
        (54, "cbb_modulI2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"P90",TEXT(cbb_modulI2,"0.00"),"TL"),"-")'),
    ]
    for r, ad, form in yorum_satir:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 28, "B": 42, "C": 12, "D": 28, "E": 14, "F": 28, "H": 24, "I": 42})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    adimlar = [
        "VERI_A: Cari hesap hareketlerini sarı hücrelere yapıştırın (VKN zorunlu).",
        "VERI_B: Form Ba/Bs beyan satırlarını girin.",
        "ANAHTAR: VKN normalize, uzunluk, ikincil ve duplike kolonları otomatik dolar.",
        "MOTOR: Ara hesap adımlarını inceler; PANO'ya besler.",
        "ESLESTIRME: Dört kova ve bütünlük denetimi otomatik çalışır.",
        "KOK_NEDEN: Tutar farklarını sınıflandırır; TANIMSIZ elle incelenir.",
        "PANO: Karar kapısı (VERİ YOK / UYGUN / İNCELE / DURDUR) ve KPI'lar.",
        "AKSIYON: Düzeltme metinlerini kopyalayıp dosyaya ekleyin.",
        "KANIT_RAPORU: Dönem dosyasına PDF olarak yazdırın.",
    ]
    for i, m in enumerate(adimlar, 5):
        h(ws, i, 1, f"{i - 4}. {m}", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    h(ws, 15, 1, "Sık yapılan hatalar", kalin=True, yazi="B3261E")
    for i, m in enumerate([
        "VKN'yi boşluklu veya eksik haneli bırakmak (normalize edilir ama kontrol edin)",
        "A ve B dönemlerini karıştırmak",
        "Ba tutarını Bs satırına yazmak",
    ], 16):
        h(ws, i, 1, "• " + m)
    h(ws, 20, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 30):
        dogrulama(ws, "textLength", "0", f"D{r}",
                 baslik="Kaynak", mesaj="Kaynak metni girin",
                 hata_baslik="Kaynak", hata_mesaj="Boş bırakmayın",
                 isaret="greaterThanOrEqual", f2="120")
    ws = wb["LISTELER"]
    for r in range(7, 9):
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Form Tipi", mesaj="Ba/Bs",
                  hata_baslik="Ad", hata_mesaj="En az 1 karakter",
                  isaret="greaterThan", f2="10")
    for r in range(7, 9):
        dogrulama(ws, "textLength", "1", f"C{r}",
                  baslik="İşlem Tipi", mesaj="Alış/Satış",
                  hata_baslik="Tip", hata_mesaj="Geçersiz",
                  isaret="greaterThan", f2="20")
    for r in range(7, 11):
        dogrulama(ws, "textLength", "1", f"E{r}",
                  baslik="Durum", mesaj="Durum girin",
                  hata_baslik="Durum", hata_mesaj="Geçersiz",
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
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"İNCELE"'], fill=sari))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(6, 23):
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
        CellIsRule(operator="equal", formula=['"DONEM"'], fill=sari))
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"UNVAN"'], fill=sari))
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"YUVARLAMA"'], fill=sari))
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"BELGE_EKSIK"'], fill=sari))

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
    ws.conditional_formatting.add(f"F{ILK}:F{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"F{ILK}:F{SON}",
        CellIsRule(operator="greaterThan", formula=["100000"], fill=sari))

    ws = wb["VERI_B"]
    ws.conditional_formatting.add(f"E{ILK}:E{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"E{ILK}:E{SON}",
        CellIsRule(operator="greaterThan", formula=["100000"], fill=sari))

    ws = wb["ANAHTAR"]
    ws.conditional_formatting.add(f"F{ILK}:F{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"N{ILK}:N{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("C26", CellIsRule(operator="equal", formula=['"BUTUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("C26", CellIsRule(operator="equal", formula=['"KAYIP"'], fill=kirmizi, font=kf))


def adlari_bagla(wb):
    for i, ana in enumerate([
        "toleransKurus", "raporTarihi", "cbb_esikEslesme", "cbb_esikFark",
        "cbb_azamiTutar", "cbb_ikincilUzunluk", "cbb_olcekHedef", "cbb_anomaliZ",
        "cbb_yuzdelikOran", "cbb_medyanOran", "cbb_yuvarlamaCarpan", "cbb_donemCarpan",
        "cbb_unvanMinLen", "cbb_normAnahtar", "cbb_tanimsizKova", "cbb_kararKapi",
        "cbb_fiyatModeli", "cbb_kvkkBeyani", "cbb_yapistirCalistir", "cbb_baBsKurali",
        "cbb_itirazMetni", "cbb_dosyaSurumu", "cbb_firmaUnvan",
    ]):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "cbb_mutabakatFarki": 32, "cbb_eslesmeOrani": 33, "cbb_kovaButunluk": 34,
        "cbb_modulT1": 35, "cbb_modulT2": 36, "cbb_modulO1": 37,
        "cbb_modulO2": 38, "cbb_modulO3": 39, "cbb_modulO6": 40,
        "cbb_modulO8": 41, "cbb_modulI1": 42, "cbb_modulI2": 43,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "cbb_modulT1Yorum": 46, "cbb_modulT2Yorum": 47, "cbb_modulO1Yorum": 48,
        "cbb_modulO2Yorum": 49, "cbb_modulO3Yorum": 50, "cbb_modulO6Yorum": 51,
        "cbb_modulO8Yorum": 52, "cbb_modulI1Yorum": 53, "cbb_modulI2Yorum": 54,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeFormTipleri", "LISTELER!$A$7:$A$8")
    ad_ekle(wb, "ListeIslemTipleri", "LISTELER!$C$7:$C$8")
    ad_ekle(wb, "ListeDurumlar", "LISTELER!$E$7:$E$10")


def ornek_csv_yaz(yol):
    satirlar = ["taraf,vkn,unvan,donem,belge_no,tarih,tutar,form_tipi,islem_tipi,belge_adet,not"]
    for s in ORNEK:
        satirlar.append(
            f"A,{s['vkn_a']},{s['unvan']},{DONEM},{s['belge']},{s['tarih'].isoformat()},"
            f"{s['tutar']},,{s['islem']},,cari"
        )
        satirlar.append(
            f"B,{s['vkn_b']},{s['unvan']},{DONEM},,{s['tarih'].isoformat()},"
            f"{s['b_tutar']},{s['form']},,{s['adet']},ba_bs"
        )
    with open(yol, "w", encoding="utf-8") as f:
        f.write("\n".join(satirlar) + "\n")


def main(cikti_yolu=None):
    wb = Workbook()
    # SPEC sayfa_akisi: KAPAK → girdiler → MOTOR → ESLESTIRME → … → PANO → rapor → sonda ayarlar
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
    dosya_ad = "CariBaBsTopluMutabakat.xlsx"
    hedef = cikti_yolu or os.path.join(urun_dir, dosya_ad)
    os.makedirs(os.path.dirname(hedef) or ".", exist_ok=True)
    wb.save(hedef)

    csv_yol = os.path.join(urun_dir, "ornek_veri.csv")
    ornek_csv_yaz(csv_yol)

    cikti = os.path.join(KOK, "cikti", dosya_ad)
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(hedef) != os.path.abspath(cikti):
        import shutil
        shutil.copy2(hedef, cikti)

    print(f"Dosya oluşturuldu: {hedef}")
    print(f"CSV: {csv_yol}")
    print(f"Kopya: {cikti}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
