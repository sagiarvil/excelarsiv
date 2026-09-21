#!/usr/bin/env python3
"""Kargo/Desi Maliyet Optimizasyonu — A2 üretim betiği (manda v6).

DOMAIN: paket ölçüleri → hacimsel desi, faturalanan birim (max desi/kg),
birim kargo, sipariş toplamı, alternatif taşıyıcı farkı ve karar
(DESİ DÜŞÜR / TAŞIYICI DEĞİŞTİR / UYGUN / VERİ YOK).
A2 eşleştirme: sipariş (VERI_A) ↔ taşıyıcı tarife (VERI_B) sipariş anahtarı.
Pazaryeri hakediş / e-ticaret kârlılık kopyası değildir.
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

URUN_AD = "Kargo/Desi Maliyet Optimizasyonu"
SURUM = "1.0.0"
RENK = "1B3A4B"
KAPASITE = 500
DEMO = 35
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

TASIYICILAR = ["Yurtici", "Aras", "MNG", "PTT", "Sürat"]
BOLGELER = ["IcAnadolu", "Marmara", "Ege", "Akdeniz", "Karadeniz"]
YORUM_AB = ["A", "B"]
DURUMLAR = ["Aktif", "Kampanya", "Durduruldu"]


def _ornek_satirlar():
    satirlar = []
    for i in range(1, DEMO + 1):
        kod = f"SIP{20000 + i}"
        if i in (3, 7, 11):
            kod_a, kod_b = kod[:8], kod
        elif i == 15:
            kod_a, kod_b = kod + "X", kod
        else:
            kod_a, kod_b = kod, kod
        en = 20 + (i % 8) * 5
        boy = 15 + (i % 6) * 4
        yuk = 10 + (i % 5) * 3
        agirlik = round(0.4 + (i % 9) * 0.35, 2)
        desi_carp = 3000
        hacim_desi = round((en * boy * yuk) / desi_carp, 2)
        fat_birim = max(hacim_desi, agirlik)
        tarif = round(8.5 + (i % 5) * 1.2, 2)
        alt_tarif = round(tarif * (0.85 if i % 4 == 0 else 1.05), 2)
        adet = 1 + (i % 3)
        iade_oran = 0.02 + (i % 5) * 0.01
        hedef = round(fat_birim * tarif * adet, 2)
        gercek = round(fat_birim * tarif * adet * (1 + iade_oran * 0.5), 2)
        if i % 8 == 0:
            gercek = round(gercek + 18, 2)
        elif i % 11 == 0:
            gercek = round(hedef + 0.03, 2)
        elif i == 15:
            gercek = hedef
        tasiyici = TASIYICILAR[i % len(TASIYICILAR)]
        bolge = BOLGELER[i % len(BOLGELER)]
        tar = RAPOR_TARIHI - timedelta(days=(DEMO - i))
        satirlar.append({
            "kod_a": kod_a, "kod_b": kod_b, "tarih": tar,
            "en": en, "boy": boy, "yuk": yuk, "agirlik": agirlik,
            "desi_carp": desi_carp, "adet": adet, "bolge": bolge,
            "tasiyici": tasiyici, "tarif": tarif, "alt_tarif": alt_tarif,
            "iade_oran": iade_oran, "hedef": hedef, "gercek": gercek,
            "yorum": "B" if i % 5 == 0 else "A",
            "hacim_desi": hacim_desi, "fat_birim": fat_birim,
        })
    return satirlar


ORNEK = _ornek_satirlar()


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        from openpyxl.comments import Comment
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Desi ve kargo kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Paket ölçüleri ve taşıyıcı tarifesinden hacimsel desi, faturalanan birim "
      "(max desi/kg), birim kargo ve sipariş toplamını hesaplar; alternatif "
      "taşıyıcı farkıyla DESİ DÜŞÜR / TAŞIYICI DEĞİŞTİR / UYGUN kararını üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Sipariş↔tarife eşleştirmesi ile gerçek kargo maliyeti",
        "Hacimsel desi, faturalanan birim ve birim kargo motoru",
        "KDO-001..005 kök neden + TANIMSIZ kova",
        "Desi israfı ve alternatif taşıyıcı farkını yakalayan karar",
        "Yöneticiye sunulabilir kanıt raporu",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "E-ticaret / pazaryeri operasyon ekipleri",
        "Lojistik ve mali işler",
        "SMMM ve dönem kargo kanıtı arayanlar",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) VERI_A'ya paket ölçülerini girin. 2) VERI_B'ye taşıyıcı tarifelerini "
      "yapıştırın. 3) MOTOR ve PANO'dan desi ile kararı izleyin. "
      "4) AKSIYON'dan aksiyon metnini kopyalayın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def veri_a(ws):
    sayfa_hazirla(ws, "VERI_A", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Sipariş / paket — en×boy×yükseklik / ağırlık / hedef maliyet (sarıya girin)", son_kolon=12)
    sutunlar = ["SiparisNo", "EnCm", "BoyCm", "YukseklikCm", "AgirlikKg",
                "DesiCarpani", "Adet", "Bolge", "YorumAB", "HedefMaliyet", "KayitDolu"]
    baslik_satiri(ws, HDR, sutunlar)
    form_a = {
        "HedefMaliyet": (
            'IF(tblVeriA[[#This Row],[SiparisNo]]="","",'
            'ROUND(MAX((tblVeriA[[#This Row],[EnCm]]*tblVeriA[[#This Row],[BoyCm]]'
            '*tblVeriA[[#This Row],[YukseklikCm]])/MAX(tblVeriA[[#This Row],[DesiCarpani]],1),'
            'tblVeriA[[#This Row],[AgirlikKg]])'
            '*IFERROR(SUMIF(tblVeriB[SiparisNo],tblVeriA[[#This Row],[SiparisNo]],tblVeriB[TarifTlDesi]),0)'
            '*tblVeriA[[#This Row],[Adet]],2))'
        ),
        "KayitDolu": 'IF(tblVeriA[[#This Row],[SiparisNo]]="","",1)',
    }
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["kod_a"], baslik="Sipariş No", mesaj="Sipariş / barkod")
        _sari(ws, r, 2, s["en"], sayi=CATI, baslik="En (cm)", mesaj="Paket eni cm")
        _sari(ws, r, 3, s["boy"], sayi=CATI, baslik="Boy (cm)", mesaj="Paket boyu cm")
        _sari(ws, r, 4, s["yuk"], sayi=CATI, baslik="Yükseklik (cm)", mesaj="Paket yüksekliği cm")
        _sari(ws, r, 5, s["agirlik"], sayi=CATI, baslik="Ağırlık (kg)", mesaj="Brüt ağırlık kg")
        _sari(ws, r, 6, s["desi_carp"], sayi=CATI, baslik="Desi Çarpanı", mesaj="Genelde 3000")
        _sari(ws, r, 7, s["adet"], sayi=CATI, baslik="Adet", mesaj="Paket adedi")
        _sari(ws, r, 8, s["bolge"], baslik="Bölge", mesaj="Liste seçimi")
        _sari(ws, r, 9, s["yorum"], baslik="Yorum A/B", mesaj="Paket varyantı A veya B")
    for r in range(ILK + DEMO, SON + 1):
        for c in range(1, 10):
            _sari(ws, r, c, None,
                  sayi=(CATI if c in (2, 3, 4, 5, 6, 7) else None),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriA", f"A{HDR}:{get_column_letter(11)}{SON}", sutunlar, formuller=form_a)
    for c, tip, f1 in [
        (1, "textLength", "1"),
        (2, "decimal", "0"), (3, "decimal", "0"), (4, "decimal", "0"),
        (5, "decimal", "0"), (6, "decimal", "1"), (7, "whole", "1"),
    ]:
        dogrulama(ws, tip, f1, f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=sutunlar[c - 1], mesaj="Geçerli değer girin",
                  hata_baslik="Geçersiz", hata_mesaj="Değer kurallara uymuyor",
                  isaret="greaterThanOrEqual" if tip != "textLength" else "greaterThan",
                  f2=None if tip != "textLength" else "40")
    dogrulama(ws, "list", "ListeBolgeler", f"H{ILK}:H{SON}",
              baslik="Bölge", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçim yapın")
    dogrulama(ws, "list", "ListeYorumAB", f"I{ILK}:I{SON}",
              baslik="Yorum A/B", mesaj="A veya B seçin",
              hata_baslik="Liste", hata_mesaj="A veya B seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 12)})


def veri_b(ws):
    sayfa_hazirla(ws, "VERI_B", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Taşıyıcı tarife — TL/desi veya kg / iade / alternatif (sarıya girin)", son_kolon=10)
    sutunlar = ["SiparisNo", "Tasiyici", "TarifTlDesi", "TarifTlKg", "IadeOran",
                "AlternatifTarif", "GercekMaliyet", "Bolge", "Donem", "KayitDolu"]
    baslik_satiri(ws, HDR, sutunlar)
    form_b = {
        "GercekMaliyet": (
            'IF(tblVeriB[[#This Row],[SiparisNo]]="","",'
            'IFERROR(MAX('
            'IFERROR(SUMIF(tblVeriA[SiparisNo],tblVeriB[[#This Row],[SiparisNo]],tblVeriA[EnCm])'
            '*SUMIF(tblVeriA[SiparisNo],tblVeriB[[#This Row],[SiparisNo]],tblVeriA[BoyCm])'
            '*SUMIF(tblVeriA[SiparisNo],tblVeriB[[#This Row],[SiparisNo]],tblVeriA[YukseklikCm])'
            '/MAX(SUMIF(tblVeriA[SiparisNo],tblVeriB[[#This Row],[SiparisNo]],tblVeriA[DesiCarpani]),1),0),'
            'IFERROR(SUMIF(tblVeriA[SiparisNo],tblVeriB[[#This Row],[SiparisNo]],tblVeriA[AgirlikKg]),0))'
            '*tblVeriB[[#This Row],[TarifTlDesi]]'
            '*IFERROR(SUMIF(tblVeriA[SiparisNo],tblVeriB[[#This Row],[SiparisNo]],tblVeriA[Adet]),1)'
            '*(1+tblVeriB[[#This Row],[IadeOran]]*kdo_iadeEtkiCarpani),0))'
        ),
        "KayitDolu": 'IF(tblVeriB[[#This Row],[SiparisNo]]="","",1)',
    }
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["kod_b"], baslik="Sipariş No", mesaj="Taşıyıcı sipariş no")
        _sari(ws, r, 2, s["tasiyici"], baslik="Taşıyıcı", mesaj="Liste seçimi")
        _sari(ws, r, 3, s["tarif"], sayi=TL, baslik="Tarife TL/Desi", mesaj="TL/desi veya TL/kg")
        _sari(ws, r, 4, s["tarif"], sayi=TL, baslik="Tarife TL/Kg", mesaj="Ağırlık tarife")
        _sari(ws, r, 5, s["iade_oran"], sayi=YÜZDE, baslik="İade %", mesaj="İade oranı")
        _sari(ws, r, 6, s["alt_tarif"], sayi=TL, baslik="Alternatif Tarife", mesaj="Rakip TL/desi")
        _sari(ws, r, 8, s["bolge"], baslik="Bölge", mesaj="Liste seçimi")
        _sari(ws, r, 9, "2026-08", baslik="Dönem", mesaj="YYYY-AA")
    for r in range(ILK + DEMO, SON + 1):
        for c in (1, 2, 3, 4, 5, 6, 8, 9):
            _sari(ws, r, c, None,
                  sayi=(YÜZDE if c == 5 else (TL if c in (3, 4, 6) else None)),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriB", f"A{HDR}:J{SON}", sutunlar, formuller=form_b)
    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}", baslik="Sipariş No",
              mesaj="Sipariş no girin", hata_baslik="Eksik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    for c, ad in [(3, "Tarife Desi"), (4, "Tarife Kg"), (6, "Alternatif")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj="0 veya üzeri", hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"E{ILK}:E{SON}", baslik="İade",
              mesaj="0-1 arası oran", hata_baslik="İade", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeTasiyicilar", f"B{ILK}:B{SON}",
              baslik="Taşıyıcı", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeBolgeler", f"H{ILK}:H{SON}",
              baslik="Bölge", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "textLength", "0", f"I{ILK}:I{SON}", baslik="Dönem",
              mesaj="YYYY-AA", hata_baslik="Dönem", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="20")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 11)})


def anahtar(ws):
    sayfa_hazirla(ws, "ANAHTAR", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Normalize sipariş anahtarı + uzunluk + ikincil + duplike (Ç7 / E01 / E05)", son_kolon=12)
    h(ws, 4, 1, "A Tarafı (Sipariş/Paket)", kalin=True, yazi=KOYU_LACIVERT)
    sut_a = ["Ham", "Norm", "NormUzunluk", "IkincilAnahtar", "KaynakTaraf", "Duplike", "SatirNo"]
    baslik_satiri(ws, HDR, sut_a)
    form_a = {
        "Norm": normalize_formul("A6").replace("A6", "tblAnahtarA[[#This Row],[Ham]]").lstrip("="),
        "NormUzunluk": 'IF(tblAnahtarA[[#This Row],[Norm]]="",0,LEN(tblAnahtarA[[#This Row],[Norm]]))',
        "IkincilAnahtar": 'IF(tblAnahtarA[[#This Row],[Norm]]="",""'
                         ',LEFT(tblAnahtarA[[#This Row],[Norm]],kdo_ikincilUzunluk))',
        "KaynakTaraf": '"A"',
        "Duplike": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",'
                   'IF(COUNTIF(tblAnahtarA[Norm],tblAnahtarA[[#This Row],[Norm]])>1,"EVET","HAYIR"))',
        "SatirNo": 'IF(tblAnahtarA[[#This Row],[Ham]]="","",ROW()-ROW($A$5))',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 1, f"=IF(VERI_A!A{r}=\"\",\"\",VERI_A!A{r})")
    tablo_ekle(ws, "tblAnahtarA", f"A{HDR}:G{SON}", sut_a, formuller=form_a)

    h(ws, 4, 9, "B Tarafı (Taşıyıcı Tarife)", kalin=True, yazi=KOYU_LACIVERT)
    sut_b = ["HamB", "NormB", "NormUzunlukB", "IkincilB", "KaynakB", "DuplikeB", "SatirNoB"]
    baslik_satiri(ws, HDR, sut_b, basla=9)
    form_b = {
        "NormB": 'UPPER(TRIM(SUBSTITUTE(tblAnahtarB[[#This Row],[HamB]]," ","")))',
        "NormUzunlukB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",0,LEN(tblAnahtarB[[#This Row],[NormB]]))',
        "IkincilB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",""'
                    ',LEFT(tblAnahtarB[[#This Row],[NormB]],kdo_ikincilUzunluk))',
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
    alt_bant(ws, 3, "Desi motoru — hacimsel desi, faturalanan birim, birim/toplam kargo, karar (D13)", son_kolon=16)
    sutunlar = [
        "A_Norm", "EnCm", "BoyCm", "YukCm", "AgirlikKg", "DesiCarpani",
        "HacimselDesi", "FaturalananBirim", "TarifTl", "Adet", "BirimKargo",
        "SiparisToplam", "AltTarif", "AltToplam", "Tasarruf", "KararSatir",
        "IadeOran", "DesiKgOran", "MotorAdim1", "MotorAdim2",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "A_Norm": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",tblAnahtarA[[#This Row],[Norm]])',
        "EnCm": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblVeriA[EnCm]),0))',
        "BoyCm": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                 'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblVeriA[BoyCm]),0))',
        "YukCm": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                 'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblVeriA[YukseklikCm]),0))',
        "AgirlikKg": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                     'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblVeriA[AgirlikKg]),0))',
        "DesiCarpani": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                       'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblVeriA[DesiCarpani]),kdo_varsayilanDesiCarpani))',
        "HacimselDesi": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'ROUND(tblMotor[[#This Row],[EnCm]]*tblMotor[[#This Row],[BoyCm]]'
            '*tblMotor[[#This Row],[YukCm]]/MAX(tblMotor[[#This Row],[DesiCarpani]],1),2))'
        ),
        "FaturalananBirim": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'MAX(tblMotor[[#This Row],[HacimselDesi]],tblMotor[[#This Row],[AgirlikKg]]))'
        ),
        "TarifTl": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                   'IFERROR(SUMIF(tblAnahtarB[NormB],tblMotor[[#This Row],[A_Norm]],tblVeriB[TarifTlDesi]),0))',
        "Adet": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblVeriA[Adet]),1))',
        "BirimKargo": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'ROUND(tblMotor[[#This Row],[FaturalananBirim]]*tblMotor[[#This Row],[TarifTl]],2))'
        ),
        "SiparisToplam": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'ROUND(tblMotor[[#This Row],[BirimKargo]]*tblMotor[[#This Row],[Adet]]'
            '*(1+tblMotor[[#This Row],[IadeOran]]*kdo_iadeEtkiCarpani),2))'
        ),
        "AltTarif": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblAnahtarB[NormB],tblMotor[[#This Row],[A_Norm]],tblVeriB[AlternatifTarif]),0))',
        "AltToplam": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'ROUND(tblMotor[[#This Row],[FaturalananBirim]]*tblMotor[[#This Row],[AltTarif]]'
            '*tblMotor[[#This Row],[Adet]]*(1+tblMotor[[#This Row],[IadeOran]]*kdo_iadeEtkiCarpani),2))'
        ),
        "Tasarruf": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                    'ROUND(tblMotor[[#This Row],[SiparisToplam]]-tblMotor[[#This Row],[AltToplam]],2))',
        "KararSatir": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'IF(tblMotor[[#This Row],[DesiKgOran]]>kdo_esikDesiOran,"DESİ DÜŞÜR",'
            'IF(tblMotor[[#This Row],[Tasarruf]]>kdo_esikTasarruf,"TAŞIYICI DEĞİŞTİR","UYGUN")))'
        ),
        "IadeOran": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblAnahtarB[NormB],tblMotor[[#This Row],[A_Norm]],tblVeriB[IadeOran]),0))',
        "DesiKgOran": (
            'IF(OR(tblMotor[[#This Row],[A_Norm]]="",tblMotor[[#This Row],[AgirlikKg]]=0),0,'
            'tblMotor[[#This Row],[HacimselDesi]]/tblMotor[[#This Row],[AgirlikKg]])'
        ),
        "MotorAdim1": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                      'tblMotor[[#This Row],[HacimselDesi]]-tblMotor[[#This Row],[AgirlikKg]])',
        "MotorAdim2": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                      'tblMotor[[#This Row],[SiparisToplam]]-tblMotor[[#This Row],[AltToplam]])',
    }
    tablo_ekle(ws, "tblMotor", f"A{HDR}:T{SON}", sutunlar, formuller=form)
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 21)})


def eslestirme(ws):
    sayfa_hazirla(ws, "ESLESTIRME", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Hedef maliyet ↔ gerçek maliyet dört kova — her A kaydı tek kovaya (E02)", son_kolon=14)
    sutunlar = [
        "A_Norm", "B_Sayim", "B_Tutar", "A_Tutar", "Fark", "AbsFark",
        "Kova", "IkincilEslesme", "DuplikeBayrak", "FarkOran",
        "YuvarlakA", "YuvarlakB", "ToleransTest", "EsikUstu",
        "SiparisToplam", "MotorAdim",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    formuller = {
        "A_Norm": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",tblAnahtarA[[#This Row],[Norm]])',
        "B_Sayim": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                   'COUNTIF(tblAnahtarB[NormB],tblEslesme[[#This Row],[A_Norm]]))',
        "B_Tutar": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                   'IFERROR(SUMIF(tblMotor[A_Norm],tblEslesme[[#This Row],[A_Norm]],tblMotor[SiparisToplam]),0))',
        "A_Tutar": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                   'IFERROR(SUMIF(tblAnahtarA[Norm],tblEslesme[[#This Row],[A_Norm]],tblVeriA[HedefMaliyet]),0))',
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
                    'IF(tblEslesme[[#This Row],[AbsFark]]>kdo_esikFark,1,0))',
        "SiparisToplam": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                         'IFERROR(SUMIF(tblMotor[A_Norm],tblEslesme[[#This Row],[A_Norm]],tblMotor[SiparisToplam]),0))',
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
    alt_bant(ws, 3, "KDO-001..005 kök neden — eşleşmeyen fark TANIMSIZ (E04)", son_kolon=10)
    h(ws, 4, 1, "Kural Önceliği", kalin=True, yazi=KOYU_LACIVERT)
    kurallar = kok_neden_onerisi([
        {"neden_kodu": "KDO-001", "test": "desi_kg_oran>esik",
         "aciklama": "Hacimsel desi israfı (desi >> kg)"},
        {"neden_kodu": "KDO-002", "test": "ABS(fark-iade)<=toleransKurus",
         "aciklama": "Yüksek iade oranı etkisi"},
        {"neden_kodu": "KDO-003", "test": "bolge_tarife_sapmasi",
         "aciklama": "Bölge tarife sapması"},
        {"neden_kodu": "KDO-004", "test": "tasarruf>esik",
         "aciklama": "Alternatif taşıyıcı daha ucuz"},
        {"neden_kodu": "KDO-005", "test": "ABS(fark)>esikFark",
         "aciklama": "Hedef-gerçek maliyet sapması"},
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
    sut = ["A_Norm", "Kova", "AbsFark", "DesiKgOran", "Tasarruf", "Iade", "BirimKargo", "KokNeden"]
    baslik_satiri(ws, 15, sut)
    form = {
        "A_Norm": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",tblEslesme[[#This Row],[A_Norm]])',
        "Kova": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                'IFERROR(INDEX(tblEslesme[Kova],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),""))',
        "AbsFark": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                   'IFERROR(INDEX(tblEslesme[AbsFark],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),0))',
        "DesiKgOran": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                      'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[DesiKgOran]),0))',
        "Tasarruf": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[Tasarruf]),0))',
        "Iade": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[IadeOran]),0))',
        "BirimKargo": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                      'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[BirimKargo]),0))',
        "KokNeden": (
            'IF(OR(tblKok[[#This Row],[A_Norm]]="",tblKok[[#This Row],[Kova]]<>"tutar_farki"),"",'
            'IF(tblKok[[#This Row],[DesiKgOran]]>kdo_esikDesiOran,"KDO-001",'
            'IF(tblKok[[#This Row],[Iade]]>kdo_esikIade,"KDO-002",'
            'IF(tblKok[[#This Row],[Tasarruf]]>kdo_esikTasarruf,"KDO-004",'
            'IF(tblKok[[#This Row],[AbsFark]]>kdo_esikFark,"KDO-005","TANIMSIZ")))))'
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
    h(ws, 3, 1, "Kargo/Desi Karar Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 3, 3, "Rapor Tarihi", yazi=GRİ)
    h(ws, 3, 4, "=raporTarihi", sayi=TARİH)

    h(ws, 4, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2,
      '=IF(COUNTA(tblVeriA[SiparisNo])=0,"VERİ YOK",'
      'IF(COUNTIF(tblMotor[KararSatir],"DESİ DÜŞÜR")>COUNTIF(tblMotor[KararSatir],"UYGUN"),"DESİ DÜŞÜR",'
      'IF(COUNTIF(tblMotor[KararSatir],"TAŞIYICI DEĞİŞTİR")>0,"TAŞIYICI DEĞİŞTİR","UYGUN")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kargo kararı | Neden önemli: Boş veride VERİ YOK | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Eşleşme Oranı", '=IFERROR(kdo_eslesmeOrani,0)', YÜZDE),
        (7, "Toplam Kargo", '=IFERROR(SUM(tblMotor[SiparisToplam]),0)', TL),
        (8, "Toplam Hedef", '=IFERROR(SUM(tblVeriA[HedefMaliyet]),0)', TL),
        (9, "Hedef-Gerçek Fark", '=IFERROR(kdo_mutabakatFarki,0)', TL),
        (10, "Ort. Birim Kargo", '=IFERROR(kdo_modulT1,0)', TL),
        (11, "Ort. Desi/Kg", '=IFERROR(kdo_modulT2,0)', CATI),
        (12, "Ort. Faturalanan", '=IFERROR(AVERAGE(tblMotor[FaturalananBirim]),0)', CATI),
        (13, "Eşleşti Adet", '=COUNTIF(tblEslesme[Kova],"eslesti")', CATI),
        (14, "Tutar Farkı Adet", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (15, "Eşleşmedi Adet", '=COUNTIF(tblEslesme[Kova],"eslesmedi")', CATI),
        (16, "Kova Bütünlük", "=kdo_kovaButunluk", None),
        (17, "Anomali Sayısı", "=kdo_modulO1", CATI),
        (18, "Tornado Zirve", "=kdo_modulO2", TL),
        (19, "Senaryo Toplam", "=kdo_modulO3", TL),
        (20, "Yoğunlaşma", "=kdo_modulO6", YÜZDE),
        (21, "Kalite Skoru", "=kdo_modulO8", CATI),
        (22, "Tahmin Aralık", "=kdo_modulI1", TL),
        (23, "Maliyet P90", "=kdo_modulI2", TL),
        (24, "Duplike Uyarı", '=COUNTIF(tblAnahtarA[Duplike],"EVET")', CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (7, '=kdo_modulT1Yorum'),
        (8, '=kdo_modulT2Yorum'),
        (9, '=kdo_modulO1Yorum'),
        (10, '=kdo_modulO2Yorum'),
        (11, '=kdo_modulO3Yorum'),
        (12, '=kdo_modulO6Yorum'),
        (13, '=kdo_modulO8Yorum'),
        (14, '=kdo_modulMSenYorum'),
        (15, '=kdo_modulI1Yorum'),
        (16, '=kdo_modulI2Yorum'),
    ]
    for r, form in yorumlar:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 26, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1, "Kova")
    h(ws, 27, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Eşleşti", 22), ("Tolerans", 4), ("Tutar farkı", 6), ("Eşleşmedi", 3),
    ], 28):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 27, 4, "Neden")
    h(ws, 27, 5, "Adet")
    for i, (ad, adet) in enumerate([
        ("KDO-001", 2), ("KDO-002", 1), ("KDO-003", 2), ("KDO-004", 1),
        ("KDO-005", 1), ("Tanımsız", 2),
    ], 28):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 27, 7, "Gün")
    h(ws, 27, 8, "Kargo")
    for i in range(7):
        h(ws, 28 + i, 7, i + 1, sayi=CATI)
        h(ws, 28 + i, 8, round(80 + i * 12.5, 2), sayi=TL)

    h(ws, 27, 10, "Senaryo")
    h(ws, 27, 11, "Etki")
    for i, (ad, v) in enumerate([
        ("Tarife +10%", 320), ("Desi -20%", 280), ("İade +2pp", 90),
        ("Adet +1", 150), ("Alt taşıyıcı", 210),
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
    line1.title = "Günlük Kargo"
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
    line2.title = "Kargo Eğilimi"
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
    alt_bant(ws, 3, "Kargo aksiyon metni — tutar farkı satırları için kopyalanabilir", son_kolon=10)
    sut = ["A_Norm", "SiparisToplam", "Tasarruf", "Karar", "Oncelik", "AksiyonMetni"]
    baslik_satiri(ws, HDR, sut)
    form = {
        "A_Norm": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                  'IF(tblEslesme[[#This Row],[Kova]]="tutar_farki",tblEslesme[[#This Row],[A_Norm]],""))',
        "SiparisToplam": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                         'IFERROR(SUMIF(tblMotor[A_Norm],tblAksiyon[[#This Row],[A_Norm]],tblMotor[SiparisToplam]),0))',
        "Tasarruf": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblMotor[A_Norm],tblAksiyon[[#This Row],[A_Norm]],tblMotor[Tasarruf]),0))',
        "Karar": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                 'IFERROR(INDEX(tblMotor[KararSatir],MATCH(tblAksiyon[[#This Row],[A_Norm]],tblMotor[A_Norm],0)),""))',
        "Oncelik": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                   'IF(OR(tblAksiyon[[#This Row],[Karar]]="DESİ DÜŞÜR",'
                   'tblAksiyon[[#This Row],[Karar]]="TAŞIYICI DEĞİŞTİR"),"YUKSEK","NORMAL"))',
        "AksiyonMetni": (
            'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
            '_xlfn.TEXTJOIN(" ",TRUE,"Sipariş",tblAksiyon[[#This Row],[A_Norm]],'
            '"karar",tblAksiyon[[#This Row],[Karar]],'
            '"toplam",TEXT(tblAksiyon[[#This Row],[SiparisToplam]],"0.00"),'
            '"tasarruf",TEXT(tblAksiyon[[#This Row],[Tasarruf]],"0.00"),"TL"))'
        ),
    }
    tablo_ekle(ws, "tblAksiyon", f"A{HDR}:F{SON}", sut, formuller=form)
    genislik(ws, {"A": 16, "B": 14, "C": 12, "D": 18, "E": 10, "F": 55})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Kargo/Desi Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Eşleşme Oranı", "=kdo_eslesmeOrani", YÜZDE),
        (10, "Toplam Kargo", "=SUM(tblMotor[SiparisToplam])", TL),
        (11, "Hedef-Gerçek Fark", "=kdo_mutabakatFarki", TL),
        (12, "Ort. Birim Kargo", "=kdo_modulT1", TL),
        (13, "Ort. Faturalanan",
         '=IFERROR(IF(COUNTA(tblMotor[A_Norm])=0,0,AVERAGE(tblMotor[FaturalananBirim])),0)', CATI),
        (14, "Tutar Farkı", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (15, "Kova Bütünlük", "=kdo_kovaButunluk", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 17, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 1,
      "Desi çarpanı varsayılan 3000'tir. Faturalanan birim max(desi, kg). "
      "Tolerans kuruşu AYARLAR'dan gelir. Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A18:H18")
    h(ws, 20, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 1, "1. DESİ DÜŞÜR satırlarında paket boyutunu küçültün.")
    h(ws, 22, 1, "2. TAŞIYICI DEĞİŞTİR satırlarında alternatif tarifeyi deneyin.")
    h(ws, 23, 1, "3. TANIMSIZ kök nedenleri elle inceleyin.")
    baski_hazirla(ws, "A1:H24", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 28, "B": 18, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "A satır sayısı", '=COUNTA(tblVeriA[SiparisNo])'),
        (6, "B satır sayısı", '=COUNTA(tblVeriB[SiparisNo])'),
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
        (17, "Kargo toplam", "=SUM(tblMotor[SiparisToplam])"),
        (18, "Hedef toplam", "=SUM(tblVeriA[HedefMaliyet])"),
        (19, "Fark kontrol", "=B18-B17"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, kalin=True, sayi=TL if r >= 17 else (CATI if r < 15 else None))
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO} satır örnek sipariş vardır: hedef↔gerçek maliyet farkı, "
      "kısa/uzun sipariş numarası (3,7,11), eşleşmeyen (15) ve Yorum B satırları kasıtlıdır. "
      "VERI_A ve VERI_B'yi temizleyip kendi paket/tarifelerinizi yapıştırabilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Karar seti: DESİ DÜŞÜR / TAŞIYICI DEĞİŞTİR / UYGUN / VERİ YOK")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 50000, sayi=CATI)
    genislik(ws, {"A": 70, "B": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=8)
    h(ws, 3, 1, "Taşıyıcı Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "Tasiyici", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, p in enumerate(TASIYICILAR, 7):
        _sari(ws, i, 1, p, baslik="Taşıyıcı", mesaj="Listeye ekleyebilirsiniz")
    h(ws, 3, 3, "Bölge Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 3, "Bolge", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, d in enumerate(BOLGELER, 7):
        _sari(ws, i, 3, d, baslik="Bölge", mesaj="Bölge değeri")
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
        ("kdo_esikEslesme", 0.85, "oran", "İç politika", "01.01.2026", "Eşleşme oranı eşiği"),
        ("kdo_esikFark", 50, "TL", "İç politika", "01.01.2026", "Fark tutarı eşiği"),
        ("kdo_esikDesiOran", 2.5, "oran", "Desi politikası", "01.01.2026", "DESİ DÜŞÜR eşiği"),
        ("kdo_esikTasarruf", 15, "TL", "Taşıyıcı politikası", "01.01.2026", "TAŞIYICI DEĞİŞTİR eşiği"),
        ("kdo_esikIade", 0.08, "oran", "İade politikası", "01.01.2026", "KDO-002 eşiği"),
        ("kdo_iadeEtkiCarpani", 0.5, "oran", "İade maliyet modeli", "01.01.2026", "İade maliyeti çarpanı"),
        ("kdo_varsayilanDesiCarpani", 3000, "sayi", "Taşıyıcı teamülü", "01.01.2026", "Varsayılan desi çarpanı"),
        ("kdo_azamiTutar", 1000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul tutar"),
        ("kdo_ikincilUzunluk", 12, "karakter", "Anahtar stratejisi SPEC", "01.01.2026", "Ikincil anahtar"),
        ("kdo_olcekHedef", 50000, "satir", "Manda A2 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("kdo_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("kdo_tarifeSenaryo", 0.10, "oran", "M_SEN parametre", "01.01.2026", "Tarife +% senaryo"),
        ("kdo_fiyatModeli", "Tek seferlik alım — aylık ücret yok", "metin", "Ürün konumlandırma", "01.01.2026", "R3 ayrım"),
        ("kdo_kvkkBeyani", "Veri cihazdan çıkmaz", "metin", "KVKK beyanı", "01.01.2026", "R3 ayrım"),
        ("kdo_faturalananBirim", "max(desi, kg) motoru", "metin", "MOTOR", "01.01.2026", "R1 ayrım"),
        ("kdo_normAnahtar", "UPPER TRIM SUBSTITUTE", "metin", "Ç7 anahtar", "01.01.2026", "Serbest alternatif"),
        ("kdo_tanimsizKova", "TANIMSIZ görünür", "metin", "E04", "01.01.2026", "Serbest alternatif"),
        ("kdo_kovaButunlukAd", "Dört kova bütünlük", "metin", "E02", "01.01.2026", "Serbest alternatif"),
        ("kdo_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("kdo_firmaUnvan", "Örnek Lojistik A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026", "Firma"),
        ("kdo_kuralSeti", "KDO-001..005", "metin", "Kök neden kuralları", "01.01.2026", "Kural seti"),
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
        (32, "kdo_mutabakatFarki",
         '=IFERROR(SUM(tblVeriA[HedefMaliyet])-SUM(tblMotor[SiparisToplam]),0)'),
        (33, "kdo_eslesmeOrani",
         '=IF(COUNTA(tblEslesme[A_Norm])=0,0,'
         '(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/COUNTA(tblEslesme[A_Norm]))'),
        (34, "kdo_kovaButunluk", "=ESLESTIRME!C2"),
        (35, "kdo_modulT1",
         '=IFERROR(IF(COUNTA(tblMotor[A_Norm])=0,0,AVERAGE(tblMotor[BirimKargo])),0)'),
        (36, "kdo_modulT2",
         '=IFERROR(IF(COUNTA(tblMotor[A_Norm])=0,0,AVERAGE(tblMotor[DesiKgOran])),0)'),
        (37, "kdo_modulO1",
         '=COUNTIF(tblEslesme[EsikUstu],1)'),
        (38, "kdo_modulO2",
         '=IFERROR(MAX(PANO!K28:K32),0)'),
        (39, "kdo_modulO3",
         '=IFERROR(SUM(tblMotor[SiparisToplam])*(1+kdo_tarifeSenaryo),0)'),
        (40, "kdo_modulO6",
         '=IFERROR(IF(COUNTA(tblVeriB[Tasiyici])=0,0,'
         'COUNTIF(tblVeriB[Tasiyici],INDEX(tblVeriB[Tasiyici],1))/COUNTA(tblVeriB[Tasiyici])),0)'),
        (41, "kdo_modulO8",
         '=IF(COUNTA(tblVeriA[SiparisNo])=0,0,'
         'ROUND(100*(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/MAX(COUNTA(tblEslesme[A_Norm]),1),0))'),
        (42, "kdo_modulMSen",
         '=IFERROR(SUM(tblMotor[SiparisToplam])*kdo_tarifeSenaryo,0)'),
        (43, "kdo_modulI1",
         '=IFERROR(PERCENTILE(tblMotor[SiparisToplam],kdo_yuzdelikOran),0)'),
        (44, "kdo_modulI2",
         '=IFERROR(PERCENTILE(tblMotor[SiparisToplam],kdo_yuzdelikOran),0)'),
    ]
    for r, ad, form in motor_c:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kalin=True)

    h(ws, 46, 8, "Canli Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorum_satir = [
        (47, "kdo_modulT1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Ort birim kargo",TEXT(kdo_modulT1,"0.00"),"TL"),"-")'),
        (48, "kdo_modulT2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Ort desi/kg",TEXT(kdo_modulT2,"0.00")),"-")'),
        (49, "kdo_modulO1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Anomali",TEXT(kdo_modulO1,"0")),"-")'),
        (50, "kdo_modulO2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tornado",TEXT(kdo_modulO2,"0.00"),"TL"),"-")'),
        (51, "kdo_modulO3Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Senaryo toplam",TEXT(kdo_modulO3,"0.00"),"TL"),"-")'),
        (52, "kdo_modulO6Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma",TEXT(kdo_modulO6,"0.0%")),"-")'),
        (53, "kdo_modulO8Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Kalite",TEXT(kdo_modulO8,"0")),"-")'),
        (54, "kdo_modulMSenYorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tarife senaryo etki",TEXT(kdo_modulMSen,"0.00"),"TL"),"-")'),
        (55, "kdo_modulI1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tahmin",TEXT(kdo_modulI1,"0.00"),"TL"),"-")'),
        (56, "kdo_modulI2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"P90",TEXT(kdo_modulI2,"0.00"),"TL"),"-")'),
    ]
    for r, ad, form in yorum_satir:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 28, "B": 42, "C": 12, "D": 28, "E": 14, "F": 28, "H": 24, "I": 42})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    adimlar = [
        "VERI_A: Sipariş ölçülerini (en, boy, yükseklik, ağırlık, desi çarpanı, adet) sarıya girin.",
        "VERI_B: Taşıyıcı tarife, iade oranı ve alternatif tarifeyi yapıştırın.",
        "ANAHTAR: Normalize sipariş numarası, uzunluk, ikincil ve duplike kolonları otomatik dolar.",
        "MOTOR: Hacimsel desi, faturalanan birim, birim/toplam kargo ve satır kararı hesaplanır.",
        "ESLESTIRME: Hedef maliyet ↔ gerçek maliyet dört kova denetimi.",
        "KOK_NEDEN: KDO-001..005 sınıflandırır; TANIMSIZ elle incelenir.",
        "PANO / KANIT_RAPORU: Kararı izleyin ve PDF yazdırın.",
    ]
    for i, m in enumerate(adimlar, 5):
        h(ws, i, 1, f"{i - 4}. {m}", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    h(ws, 13, 1, "Sık yapılan hatalar", kalin=True, yazi="B3261E")
    for i, m in enumerate([
        "Sipariş numarasını boşluklu bırakmak (normalize edilir ama kontrol edin)",
        "Desi çarpanını 0 yazmak (bölme hatası — varsayılan 3000 kullanın)",
        "Tarifeyi yüzde yerine TL/desi olarak girmemek",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 29):
        dogrulama(ws, "textLength", "0", f"D{r}",
                  baslik="Kaynak", mesaj="Kaynak metni girin",
                  hata_baslik="Kaynak", hata_mesaj="Boş bırakmayın",
                  isaret="greaterThanOrEqual", f2="120")
    ws = wb["LISTELER"]
    for r in range(7, 12):
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Taşıyıcı", mesaj="Ad girin",
                  hata_baslik="Ad", hata_mesaj="En az 1 karakter",
                  isaret="greaterThan", f2="40")
    for r in range(7, 12):
        dogrulama(ws, "textLength", "1", f"C{r}",
                  baslik="Bölge", mesaj="Bölge girin",
                  hata_baslik="Bölge", hata_mesaj="Geçersiz",
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
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"DESİ DÜŞÜR"'], fill=sari))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"TAŞIYICI DEĞİŞTİR"'], fill=kirmizi, font=kf))
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
    ws.conditional_formatting.add(f"P{ILK}:P{SON}",
        CellIsRule(operator="equal", formula=['"DESİ DÜŞÜR"'], fill=sari))
    ws.conditional_formatting.add(f"P{ILK}:P{SON}",
        CellIsRule(operator="equal", formula=['"TAŞIYICI DEĞİŞTİR"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"P{ILK}:P{SON}",
        CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(f"O{ILK}:O{SON}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=sari))

    ws = wb["KOK_NEDEN"]
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"TANIMSIZ"'], fill=kirmizi, font=kf))
    for kod in ("KDO-001", "KDO-002", "KDO-003", "KDO-004", "KDO-005"):
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
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"DESİ DÜŞÜR"'], fill=sari))
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))

    ws = wb["VERI_A"]
    ws.conditional_formatting.add(f"B{ILK}:B{SON}",
        CellIsRule(operator="greaterThan", formula=["100"], fill=sari))
    ws.conditional_formatting.add(f"E{ILK}:E{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["VERI_B"]
    ws.conditional_formatting.add(f"C{ILK}:C{SON}",
        CellIsRule(operator="greaterThan", formula=["50"], fill=sari))
    ws.conditional_formatting.add(f"E{ILK}:E{SON}",
        CellIsRule(operator="greaterThan", formula=["0.15"], fill=sari))

    ws = wb["ANAHTAR"]
    ws.conditional_formatting.add(f"F{ILK}:F{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"N{ILK}:N{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))


def adlari_bagla(wb):
    for i, ana in enumerate([
        "toleransKurus", "raporTarihi", "kdo_esikEslesme", "kdo_esikFark",
        "kdo_esikDesiOran", "kdo_esikTasarruf", "kdo_esikIade", "kdo_iadeEtkiCarpani",
        "kdo_varsayilanDesiCarpani", "kdo_azamiTutar",
        "kdo_ikincilUzunluk", "kdo_olcekHedef", "kdo_yuzdelikOran",
        "kdo_tarifeSenaryo", "kdo_fiyatModeli", "kdo_kvkkBeyani",
        "kdo_faturalananBirim", "kdo_normAnahtar", "kdo_tanimsizKova",
        "kdo_kovaButunlukAd", "kdo_dosyaSurumu", "kdo_firmaUnvan", "kdo_kuralSeti",
    ]):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "kdo_mutabakatFarki": 32, "kdo_eslesmeOrani": 33, "kdo_kovaButunluk": 34,
        "kdo_modulT1": 35, "kdo_modulT2": 36, "kdo_modulO1": 37, "kdo_modulO2": 38,
        "kdo_modulO3": 39, "kdo_modulO6": 40, "kdo_modulO8": 41, "kdo_modulMSen": 42,
        "kdo_modulI1": 43, "kdo_modulI2": 44,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "kdo_modulT1Yorum": 47, "kdo_modulT2Yorum": 48, "kdo_modulO1Yorum": 49,
        "kdo_modulO2Yorum": 50, "kdo_modulO3Yorum": 51, "kdo_modulO6Yorum": 52,
        "kdo_modulO8Yorum": 53, "kdo_modulMSenYorum": 54, "kdo_modulI1Yorum": 55,
        "kdo_modulI2Yorum": 56,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeTasiyicilar", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeBolgeler", "LISTELER!$C$7:$C$11")
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
    dosya_ad = "KargoDesiMaliyetOptimizasyonu.xlsx"
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
