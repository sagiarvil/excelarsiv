#!/usr/bin/env python3
"""E-Ticaret Gerçek Kârlılık & Fiyatlama — A2 üretim betiği (manda v6).

DOMAIN: ürün/SKU bazlı gerçek net kâr, katkı payı, başabaş fiyat ve
fiyatlama kararı (ZAM YAP / TUT / İNDİRİMİ KES / VERİ YOK).
A2 eşleştirme: katalog (VERI_A) ↔ platform maliyet (VERI_B) SKU anahtarı.
Pazaryeri hakediş mutabakat kopyası değildir.
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

URUN_AD = "E-Ticaret Gerçek Kârlılık & Fiyatlama"
SURUM = "1.0.0"
RENK = "0B3D2E"
KAPASITE = 500
DEMO = 35
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

PLATFORMLAR = ["Trendyol", "Hepsiburada", "Amazon", "N11", "KendiSite"]
YORUM_AB = ["A", "B"]
DURUMLAR = ["Aktif", "Kampanya", "Durduruldu"]


def _ornek_satirlar():
    satirlar = []
    for i in range(1, DEMO + 1):
        kod = f"SKU{10000 + i}"
        if i in (3, 7, 11):
            kod_a, kod_b = kod[:8], kod
        elif i == 15:
            kod_a, kod_b = kod + "X", kod
        else:
            kod_a, kod_b = kod, kod
        satis = 200 + i * 23
        maliyet = round(satis * (0.42 + (i % 5) * 0.02), 2)
        kom_oran = 0.10 + (i % 4) * 0.02
        kargo = 29.9 if i % 3 else 39.9
        iade_oran = 0.02 + (i % 6) * 0.01
        reklam = round(satis * 0.05, 2) if i % 4 == 0 else round(satis * 0.02, 2)
        roas = 4.0 if i % 5 else 2.2
        hedef_katki = 0.18
        # hedef net (A) ve gerçek net (B) — kasıtlı farklar
        hedef_net = round(satis * hedef_katki, 2)
        kom = round(satis * kom_oran, 2)
        iade_m = round(satis * iade_oran, 2)
        gercek_net = round(satis - maliyet - kom - kargo - iade_m - reklam, 2)
        if i % 8 == 0:
            gercek_net = round(gercek_net - 25, 2)
        elif i % 11 == 0:
            gercek_net = round(hedef_net + 0.03, 2)
        elif i == 15:
            gercek_net = hedef_net
        plat = PLATFORMLAR[i % len(PLATFORMLAR)]
        tar = RAPOR_TARIHI - timedelta(days=(DEMO - i))
        satirlar.append({
            "kod_a": kod_a, "kod_b": kod_b, "ad": f"Ürün {i}",
            "tarih": tar, "satis": satis, "maliyet": maliyet,
            "hedef_katki": hedef_katki, "hedef_net": hedef_net,
            "kdv": 0.20, "kom_oran": kom_oran, "kargo": kargo,
            "iade_oran": iade_oran, "reklam": reklam, "roas": roas,
            "gercek_net": gercek_net, "plat": plat,
            "yorum": "B" if i % 5 == 0 else "A",
        })
    return satirlar


ORNEK = _ornek_satirlar()


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        from openpyxl.comments import Comment
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Gerçek kâr ve fiyatlama kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Ürün/SKU bazında satış fiyatı, maliyet, komisyon, kargo, iade ve reklam "
      "sonrası gerçek net kârı hesaplar; başabaş fiyat ve ZAM YAP / TUT / "
      "İNDİRİMİ KES kararını üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Katalog↔platform SKU eşleştirmesi ile gerçek net kâr",
        "Brüt/net kâr, katkı payı ve başabaş fiyat motoru",
        "Komisyon, iade, ROAS, kargo kök-neden kuralları + TANIMSIZ",
        "Kampanya/indirim zararını yakalayan fiyatlama kararı",
        "Yöneticiye sunulabilir kanıt raporu",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "E-ticaret / pazaryeri satıcıları",
        "Mali işler ve fiyatlama ekipleri",
        "SMMM ve dönem kâr kanıtı arayanlar",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) VERI_A'ya ürün katalogunu girin. 2) VERI_B'ye platform maliyetlerini "
      "yapıştırın. 3) MOTOR ve PANO'dan net kâr ile kararı izleyin. "
      "4) AKSIYON'dan fiyat aksiyon metnini kopyalayın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def veri_a(ws):
    sayfa_hazirla(ws, "VERI_A", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Ürün kataloğu — satış fiyatı / maliyet / hedef katkı (sarıya girin)", son_kolon=12)
    sutunlar = ["UrunKodu", "UrunAdi", "SatisFiyati", "Maliyet", "HedefKatkiOran",
                "HedefNet", "KdvOrani", "Platform", "YorumAB", "Durum", "KayitDolu"]
    baslik_satiri(ws, HDR, sutunlar)
    form_a = {
        "HedefNet": (
            'IF(tblVeriA[[#This Row],[UrunKodu]]="","",'
            'ROUND(tblVeriA[[#This Row],[SatisFiyati]]*tblVeriA[[#This Row],[HedefKatkiOran]],2))'
        ),
        "KayitDolu": 'IF(tblVeriA[[#This Row],[UrunKodu]]="","",1)',
    }
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["kod_a"], baslik="Ürün Kodu", mesaj="SKU / barkod")
        _sari(ws, r, 2, s["ad"], baslik="Ürün Adı", mesaj="Kısa ürün adı")
        _sari(ws, r, 3, s["satis"], sayi=TL, baslik="Satış Fiyatı", mesaj="KDV hariç satış TL")
        _sari(ws, r, 4, s["maliyet"], sayi=TL, baslik="Maliyet", mesaj="Birim maliyet TL")
        _sari(ws, r, 5, s["hedef_katki"], sayi=YÜZDE, baslik="Hedef Katkı", mesaj="Hedef katkı oranı")
        # HedefNet hesaplanan
        _sari(ws, r, 7, s["kdv"], sayi=YÜZDE, baslik="KDV Oranı", mesaj="KDV oranı")
        _sari(ws, r, 8, s["plat"], baslik="Platform", mesaj="Liste seçimi")
        _sari(ws, r, 9, s["yorum"], baslik="Yorum A/B", mesaj="Fiyat varyantı A veya B")
        _sari(ws, r, 10, "Aktif" if i % 7 else "Kampanya", baslik="Durum", mesaj="Liste seçimi")
    for r in range(ILK + DEMO, SON + 1):
        for c in range(1, 11):
            if c == 6:
                continue
            _sari(ws, r, c, None,
                  sayi=(TL if c in (3, 4) else (YÜZDE if c in (5, 7) else None)),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriA", f"A{HDR}:{get_column_letter(11)}{SON}", sutunlar, formuller=form_a)
    for c, tip, f1 in [
        (1, "textLength", "1"), (2, "textLength", "1"),
        (3, "decimal", "0"), (4, "decimal", "0"),
        (5, "decimal", "0"), (7, "decimal", "0"),
    ]:
        dogrulama(ws, tip, f1, f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=sutunlar[c - 1], mesaj="Geçerli değer girin",
                  hata_baslik="Geçersiz", hata_mesaj="Değer kurallara uymuyor",
                  isaret="greaterThanOrEqual" if tip == "decimal" else "greaterThan",
                  f2=None if tip != "textLength" else "40")
    dogrulama(ws, "list", "ListePlatformlar", f"H{ILK}:H{SON}",
              baslik="Platform", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçim yapın")
    dogrulama(ws, "list", "ListeYorumAB", f"I{ILK}:I{SON}",
              baslik="Yorum A/B", mesaj="A veya B seçin",
              hata_baslik="Liste", hata_mesaj="A veya B seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"J{ILK}:J{SON}",
              baslik="Durum", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçim yapın")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 12)})


def veri_b(ws):
    sayfa_hazirla(ws, "VERI_B", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Platform gerçekleşen maliyet — komisyon / kargo / iade / reklam", son_kolon=10)
    sutunlar = ["UrunKodu", "KomisyonOran", "Kargo", "IadeOran", "ReklamTutar",
                "Roas", "GercekNet", "Platform", "Donem", "KayitDolu"]
    baslik_satiri(ws, HDR, sutunlar)
    form_b = {
        "GercekNet": (
            'IF(tblVeriB[[#This Row],[UrunKodu]]="","",'
            'IFERROR(SUMIF(tblVeriA[UrunKodu],tblVeriB[[#This Row],[UrunKodu]],tblVeriA[SatisFiyati])'
            '-SUMIF(tblVeriA[UrunKodu],tblVeriB[[#This Row],[UrunKodu]],tblVeriA[Maliyet])'
            '-SUMIF(tblVeriA[UrunKodu],tblVeriB[[#This Row],[UrunKodu]],tblVeriA[SatisFiyati])'
            '*tblVeriB[[#This Row],[KomisyonOran]]'
            '-tblVeriB[[#This Row],[Kargo]]'
            '-SUMIF(tblVeriA[UrunKodu],tblVeriB[[#This Row],[UrunKodu]],tblVeriA[SatisFiyati])'
            '*tblVeriB[[#This Row],[IadeOran]]'
            '-tblVeriB[[#This Row],[ReklamTutar]],0))'
        ),
        "KayitDolu": 'IF(tblVeriB[[#This Row],[UrunKodu]]="","",1)',
    }
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["kod_b"], baslik="Ürün Kodu", mesaj="Platform SKU")
        _sari(ws, r, 2, s["kom_oran"], sayi=YÜZDE, baslik="Komisyon %", mesaj="Komisyon oranı")
        _sari(ws, r, 3, s["kargo"], sayi=TL, baslik="Kargo", mesaj="Birim kargo TL")
        _sari(ws, r, 4, s["iade_oran"], sayi=YÜZDE, baslik="İade %", mesaj="İade oranı")
        _sari(ws, r, 5, s["reklam"], sayi=TL, baslik="Reklam Tutarı", mesaj="Reklam TL")
        _sari(ws, r, 6, s["roas"], sayi=CATI, baslik="ROAS", mesaj="Reklam getirisi")
        _sari(ws, r, 8, s["plat"], baslik="Platform", mesaj="Liste seçimi")
        _sari(ws, r, 9, "2026-08", baslik="Dönem", mesaj="YYYY-AA")
    for r in range(ILK + DEMO, SON + 1):
        for c in (1, 2, 3, 4, 5, 6, 8, 9):
            _sari(ws, r, c, None,
                  sayi=(YÜZDE if c in (2, 4) else (TL if c in (3, 5) else (CATI if c == 6 else None))),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriB", f"A{HDR}:J{SON}", sutunlar, formuller=form_b)
    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}", baslik="Ürün Kodu",
              mesaj="SKU girin", hata_baslik="Eksik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Komisyon"), (4, "İade")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj="0-1 arası oran", hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    for c, ad in [(3, "Kargo"), (5, "Reklam"), (6, "ROAS")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj="0 veya üzeri", hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListePlatformlar", f"H{ILK}:H{SON}",
              baslik="Platform", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "textLength", "0", f"I{ILK}:I{SON}", baslik="Dönem",
              mesaj="YYYY-AA", hata_baslik="Dönem", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="20")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 11)})


def anahtar(ws):
    sayfa_hazirla(ws, "ANAHTAR", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Normalize SKU anahtarı + uzunluk + ikincil + duplike (Ç7 / E01 / E05)", son_kolon=12)
    h(ws, 4, 1, "A Tarafı (Katalog)", kalin=True, yazi=KOYU_LACIVERT)
    sut_a = ["Ham", "Norm", "NormUzunluk", "IkincilAnahtar", "KaynakTaraf", "Duplike", "SatirNo"]
    baslik_satiri(ws, HDR, sut_a)
    form_a = {
        "Norm": normalize_formul("A6").replace("A6", "tblAnahtarA[[#This Row],[Ham]]").lstrip("="),
        "NormUzunluk": 'IF(tblAnahtarA[[#This Row],[Norm]]="",0,LEN(tblAnahtarA[[#This Row],[Norm]]))',
        "IkincilAnahtar": 'IF(tblAnahtarA[[#This Row],[Norm]]="",""'
                         ',LEFT(tblAnahtarA[[#This Row],[Norm]],egk_ikincilUzunluk))',
        "KaynakTaraf": '"A"',
        "Duplike": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",'
                   'IF(COUNTIF(tblAnahtarA[Norm],tblAnahtarA[[#This Row],[Norm]])>1,"EVET","HAYIR"))',
        "SatirNo": 'IF(tblAnahtarA[[#This Row],[Ham]]="","",ROW()-ROW($A$5))',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 1, f"=IF(VERI_A!A{r}=\"\",\"\",VERI_A!A{r})")
    tablo_ekle(ws, "tblAnahtarA", f"A{HDR}:G{SON}", sut_a, formuller=form_a)

    h(ws, 4, 9, "B Tarafı (Platform)", kalin=True, yazi=KOYU_LACIVERT)
    sut_b = ["HamB", "NormB", "NormUzunlukB", "IkincilB", "KaynakB", "DuplikeB", "SatirNoB"]
    baslik_satiri(ws, HDR, sut_b, basla=9)
    form_b = {
        "NormB": 'UPPER(TRIM(SUBSTITUTE(tblAnahtarB[[#This Row],[HamB]]," ","")))',
        "NormUzunlukB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",0,LEN(tblAnahtarB[[#This Row],[NormB]]))',
        "IkincilB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",""'
                    ',LEFT(tblAnahtarB[[#This Row],[NormB]],egk_ikincilUzunluk))',
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
    alt_bant(ws, 3, "Gerçek kârlılık motoru — brüt/net, katkı, başabaş, satır kararı (D13)", son_kolon=16)
    sutunlar = [
        "A_Norm", "Satis", "Maliyet", "KomisyonOran", "KomisyonTutar", "Kargo",
        "IadeOran", "IadeMaliyet", "ReklamHarcama", "BrutKar", "NetKar",
        "KatkiPayi", "BasabasFiyat", "HedefNet", "FarkHedef", "KararSatir",
        "RoasEtki", "IndirimBayrak", "MotorAdim1", "MotorAdim2",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "A_Norm": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",tblAnahtarA[[#This Row],[Norm]])',
        "Satis": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                 'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblVeriA[SatisFiyati]),0))',
        "Maliyet": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                   'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblVeriA[Maliyet]),0))',
        "KomisyonOran": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                        'IFERROR(SUMIF(tblAnahtarB[NormB],tblMotor[[#This Row],[A_Norm]],tblVeriB[KomisyonOran]),0))',
        "KomisyonTutar": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                         'ROUND(tblMotor[[#This Row],[Satis]]*tblMotor[[#This Row],[KomisyonOran]],2))',
        "Kargo": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                 'IFERROR(SUMIF(tblAnahtarB[NormB],tblMotor[[#This Row],[A_Norm]],tblVeriB[Kargo]),0))',
        "IadeOran": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblAnahtarB[NormB],tblMotor[[#This Row],[A_Norm]],tblVeriB[IadeOran]),0))',
        "IadeMaliyet": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                       'ROUND(tblMotor[[#This Row],[Satis]]*tblMotor[[#This Row],[IadeOran]],2))',
        "ReklamHarcama": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                         'IFERROR(SUMIF(tblAnahtarB[NormB],tblMotor[[#This Row],[A_Norm]],tblVeriB[ReklamTutar]),0))',
        "BrutKar": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                   'ROUND(tblMotor[[#This Row],[Satis]]-tblMotor[[#This Row],[Maliyet]],2))',
        "NetKar": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'ROUND(tblMotor[[#This Row],[Satis]]-tblMotor[[#This Row],[Maliyet]]'
            '-tblMotor[[#This Row],[KomisyonTutar]]-tblMotor[[#This Row],[Kargo]]'
            '-tblMotor[[#This Row],[IadeMaliyet]]-tblMotor[[#This Row],[ReklamHarcama]],2))'
        ),
        "KatkiPayi": 'IF(OR(tblMotor[[#This Row],[A_Norm]]="",tblMotor[[#This Row],[Satis]]=0),0,'
                     'tblMotor[[#This Row],[NetKar]]/tblMotor[[#This Row],[Satis]])',
        "BasabasFiyat": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'IF((1-tblMotor[[#This Row],[KomisyonOran]]-tblMotor[[#This Row],[IadeOran]])<=0,0,'
            'ROUND((tblMotor[[#This Row],[Maliyet]]+tblMotor[[#This Row],[Kargo]]'
            '+tblMotor[[#This Row],[ReklamHarcama]])'
            '/(1-tblMotor[[#This Row],[KomisyonOran]]-tblMotor[[#This Row],[IadeOran]]),2)))'
        ),
        "HedefNet": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblAnahtarA[Norm],tblMotor[[#This Row],[A_Norm]],tblVeriA[HedefNet]),0))',
        "FarkHedef": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                     'ROUND(tblMotor[[#This Row],[NetKar]]-tblMotor[[#This Row],[HedefNet]],2))',
        "KararSatir": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'IF(tblMotor[[#This Row],[KatkiPayi]]<egk_esikKatkiDusuk,"ZAM YAP",'
            'IF(AND(tblMotor[[#This Row],[IndirimBayrak]]=1,'
            'tblMotor[[#This Row],[KatkiPayi]]<egk_esikKatkiHedef),"İNDİRİMİ KES","TUT")))'
        ),
        "RoasEtki": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblAnahtarB[NormB],tblMotor[[#This Row],[A_Norm]],tblVeriB[Roas]),0))',
        "IndirimBayrak": (
            'IF(tblMotor[[#This Row],[A_Norm]]="","",'
            'IF(OR(IFERROR(INDEX(tblVeriA[YorumAB],MATCH(tblMotor[[#This Row],[A_Norm]],tblAnahtarA[Norm],0)),"")="B",'
            'IFERROR(INDEX(tblVeriA[Durum],MATCH(tblMotor[[#This Row],[A_Norm]],tblAnahtarA[Norm],0)),"")="Kampanya"),1,0))'
        ),
        "MotorAdim1": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                      'tblMotor[[#This Row],[BrutKar]]-tblMotor[[#This Row],[KomisyonTutar]])',
        "MotorAdim2": 'IF(tblMotor[[#This Row],[A_Norm]]="","",'
                      'tblMotor[[#This Row],[NetKar]]-tblMotor[[#This Row],[FarkHedef]])',
    }
    tablo_ekle(ws, "tblMotor", f"A{HDR}:T{SON}", sutunlar, formuller=form)
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 21)})


def eslestirme(ws):
    sayfa_hazirla(ws, "ESLESTIRME", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Hedef net ↔ gerçek net dört kova — her A kaydı tek kovaya (E02)", son_kolon=14)
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
                   'IFERROR(SUMIF(tblMotor[A_Norm],tblEslesme[[#This Row],[A_Norm]],tblMotor[NetKar]),0))',
        "A_Tutar": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                   'IFERROR(SUMIF(tblAnahtarA[Norm],tblEslesme[[#This Row],[A_Norm]],tblVeriA[HedefNet]),0))',
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
                    'IF(tblEslesme[[#This Row],[AbsFark]]>egk_esikFark,1,0))',
        "NetKatki": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblMotor[A_Norm],tblEslesme[[#This Row],[A_Norm]],tblMotor[NetKar]),0))',
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
    alt_bant(ws, 3, "EGK-001..005 kök neden — eşleşmeyen fark TANIMSIZ (E04)", son_kolon=10)
    h(ws, 4, 1, "Kural Önceliği", kalin=True, yazi=KOYU_LACIVERT)
    kurallar = kok_neden_onerisi([
        {"neden_kodu": "EGK-001", "test": "ABS(fark-komisyon)<=toleransKurus",
         "aciklama": "Yüksek komisyon baskısı"},
        {"neden_kodu": "EGK-002", "test": "ABS(fark-iade)<=toleransKurus",
         "aciklama": "Yüksek iade oranı"},
        {"neden_kodu": "EGK-003", "test": "ABS(fark-reklam)<=toleransKurus",
         "aciklama": "Zayıf ROAS / yüksek reklam"},
        {"neden_kodu": "EGK-004", "test": "ABS(fark-kargo)<=toleransKurus",
         "aciklama": "Yüksek kargo maliyeti"},
        {"neden_kodu": "EGK-005", "test": "maliyet/satis>esik",
         "aciklama": "Maliyet baskısı / düşük brüt"},
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
    sut = ["A_Norm", "Kova", "AbsFark", "Komisyon", "Kargo", "Iade", "Reklam", "KokNeden"]
    baslik_satiri(ws, 15, sut)
    form = {
        "A_Norm": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",tblEslesme[[#This Row],[A_Norm]])',
        "Kova": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                'IFERROR(INDEX(tblEslesme[Kova],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),""))',
        "AbsFark": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                   'IFERROR(INDEX(tblEslesme[AbsFark],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),0))',
        "Komisyon": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[KomisyonTutar]),0))',
        "Kargo": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                 'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[Kargo]),0))',
        "Iade": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[IadeMaliyet]),0))',
        "Reklam": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                  'IFERROR(SUMIF(tblMotor[A_Norm],tblKok[[#This Row],[A_Norm]],tblMotor[ReklamHarcama]),0))',
        "KokNeden": (
            'IF(OR(tblKok[[#This Row],[A_Norm]]="",tblKok[[#This Row],[Kova]]<>"tutar_farki"),"",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[Komisyon]])<=toleransKurus,"EGK-001",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[Iade]])<=toleransKurus,"EGK-002",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[Reklam]])<=toleransKurus,"EGK-003",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[Kargo]])<=toleransKurus,"EGK-004",'
            'IF(tblKok[[#This Row],[AbsFark]]>egk_esikFark,"EGK-005","TANIMSIZ"))))))'
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
    h(ws, 3, 1, "Fiyatlama Karar Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 3, 3, "Rapor Tarihi", yazi=GRİ)
    h(ws, 3, 4, "=raporTarihi", sayi=TARİH)

    h(ws, 4, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2,
      '=IF(COUNTA(tblVeriA[UrunKodu])=0,"VERİ YOK",'
      'IF(COUNTIF(tblMotor[KararSatir],"ZAM YAP")>COUNTIF(tblMotor[KararSatir],"TUT"),"ZAM YAP",'
      'IF(COUNTIF(tblMotor[KararSatir],"İNDİRİMİ KES")>0,"İNDİRİMİ KES","TUT")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem fiyatlama kararı | Neden önemli: Boş veride VERİ YOK | "
               "Doğru kullanım: Otomatik | Örnek: TUT | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Eşleşme Oranı", '=IFERROR(egk_eslesmeOrani,0)', YÜZDE),
        (7, "Toplam Net Kâr", '=IFERROR(SUM(tblMotor[NetKar]),0)', TL),
        (8, "Toplam Hedef Net", '=IFERROR(SUM(tblVeriA[HedefNet]),0)', TL),
        (9, "Hedef-Gerçek Fark", '=IFERROR(egk_mutabakatFarki,0)', TL),
        (10, "Ortalama Katkı", '=IFERROR(egk_modulT1,0)', YÜZDE),
        (11, "Ortalama Komisyon", '=IFERROR(egk_modulT2,0)', YÜZDE),
        (12, "Başabaş Ort.", '=IFERROR(AVERAGE(tblMotor[BasabasFiyat]),0)', TL),
        (13, "Eşleşti Adet", '=COUNTIF(tblEslesme[Kova],"eslesti")', CATI),
        (14, "Tutar Farkı Adet", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (15, "Eşleşmedi Adet", '=COUNTIF(tblEslesme[Kova],"eslesmedi")', CATI),
        (16, "Kova Bütünlük", "=egk_kovaButunluk", None),
        (17, "Anomali Sayısı", "=egk_modulO1", CATI),
        (18, "Tornado Zirve", "=egk_modulO2", TL),
        (19, "Senaryo Net", "=egk_modulO3", TL),
        (20, "Yoğunlaşma", "=egk_modulO6", YÜZDE),
        (21, "Kalite Skoru", "=egk_modulO8", CATI),
        (22, "Tahmin Aralık", "=egk_modulI1", TL),
        (23, "Net P90", "=egk_modulI2", TL),
        (24, "Duplike Uyarı", '=COUNTIF(tblAnahtarA[Duplike],"EVET")', CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (7, '=egk_modulT1Yorum'),
        (8, '=egk_modulT2Yorum'),
        (9, '=egk_modulO1Yorum'),
        (10, '=egk_modulO2Yorum'),
        (11, '=egk_modulO3Yorum'),
        (12, '=egk_modulO6Yorum'),
        (13, '=egk_modulO8Yorum'),
        (14, '=egk_modulMSenYorum'),
        (15, '=egk_modulI1Yorum'),
        (16, '=egk_modulI2Yorum'),
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
        ("EGK-001", 2), ("EGK-002", 1), ("EGK-003", 2), ("EGK-004", 1),
        ("EGK-005", 1), ("Tanımsız", 2),
    ], 28):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 27, 7, "Gün")
    h(ws, 27, 8, "NetKar")
    for i in range(7):
        h(ws, 28 + i, 7, i + 1, sayi=CATI)
        h(ws, 28 + i, 8, round(120 + i * 18.5, 2), sayi=TL)

    h(ws, 27, 10, "Senaryo")
    h(ws, 27, 11, "Etki")
    for i, (ad, v) in enumerate([
        ("Komisyon +1pp", 380), ("Kargo +5", 160), ("İade +2pp", 210),
        ("Reklam +10", 140), ("Zam +5%", 420),
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
    line1.title = "Günlük Net Kâr"
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
    line2.title = "Net Kâr Eğilimi"
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
            '_xlfn.TEXTJOIN(" ",TRUE,"SKU",tblAksiyon[[#This Row],[A_Norm]],'
            '"karar",tblAksiyon[[#This Row],[Karar]],'
            '"net",TEXT(tblAksiyon[[#This Row],[NetKar]],"0.00"),'
            '"başabaş",TEXT(tblAksiyon[[#This Row],[Basabas]],"0.00"),"TL"))'
        ),
    }
    tablo_ekle(ws, "tblAksiyon", f"A{HDR}:F{SON}", sut, formuller=form)
    genislik(ws, {"A": 16, "B": 14, "C": 12, "D": 14, "E": 10, "F": 55})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Gerçek Kârlılık Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Eşleşme Oranı", "=egk_eslesmeOrani", YÜZDE),
        (10, "Toplam Net Kâr", "=SUM(tblMotor[NetKar])", TL),
        (11, "Hedef-Gerçek Fark", "=egk_mutabakatFarki", TL),
        (12, "Ortalama Katkı", "=egk_modulT1", YÜZDE),
        (13, "Başabaş Ort.", '=IFERROR(IF(COUNTA(tblMotor[A_Norm])=0,0,AVERAGE(tblMotor[BasabasFiyat])),0)', TL),
        (14, "Tutar Farkı", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (15, "Kova Bütünlük", "=egk_kovaButunluk", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 17, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 1,
      "Tolerans kuruşu AYARLAR'dan gelir. SKU UPPER/TRIM/SUBSTITUTE ile normalize edilir. "
      "Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A18:H18")
    h(ws, 20, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 1, "1. ZAM YAP satırlarında başabaş fiyatın üstüne çıkın.")
    h(ws, 22, 1, "2. İNDİRİMİ KES satırlarında kampanya/varyant B'yi kapatın.")
    h(ws, 23, 1, "3. TANIMSIZ kök nedenleri elle inceleyin.")
    baski_hazirla(ws, "A1:H24", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 28, "B": 18, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "A satır sayısı", '=COUNTA(tblVeriA[UrunKodu])'),
        (6, "B satır sayısı", '=COUNTA(tblVeriB[UrunKodu])'),
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
        (18, "Hedef toplam", "=SUM(tblVeriA[HedefNet])"),
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
      f"Dosyada {DEMO} satır örnek ürün vardır: hedef↔gerçek net kâr farkı, "
      "kısa/uzun SKU (3,7,11), eşleşmeyen (15) ve kampanya/Yorum B satırları kasıtlıdır. "
      "VERI_A ve VERI_B'yi temizleyip kendi kataloğunuzu yapıştırabilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Karar seti: ZAM YAP / TUT / İNDİRİMİ KES / VERİ YOK")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 50000, sayi=CATI)
    genislik(ws, {"A": 70, "B": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=8)
    h(ws, 3, 1, "Platform Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "Platform", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, p in enumerate(PLATFORMLAR, 7):
        _sari(ws, i, 1, p, baslik="Platform", mesaj="Listeye ekleyebilirsiniz")
    h(ws, 3, 3, "Durum Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 3, "Durum", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, d in enumerate(DURUMLAR, 7):
        _sari(ws, i, 3, d, baslik="Durum", mesaj="Durum değeri")
    h(ws, 3, 5, "Yorum A/B", kalin=True)
    h(ws, 6, 5, "Secim", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, y in enumerate(YORUM_AB, 7):
        _sari(ws, i, 5, y, baslik="Yorum", mesaj="A veya B")
    genislik(ws, {"A": 18, "C": 16, "E": 12})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Tek kaynak parametreler — kaynak ve yürürlük zorunlu (D10)", son_kolon=10)
    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("toleransKurus", 0.05, "TL", "İş kuralı — kuruş toleransı", "01.01.2026", "Eşleşme toleransı"),
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026", "Rapor tarihi"),
        ("egk_esikEslesme", 0.85, "oran", "İç politika", "01.01.2026", "Eşleşme oranı eşiği"),
        ("egk_esikFark", 50, "TL", "İç politika", "01.01.2026", "Fark tutarı eşiği"),
        ("egk_esikKatkiDusuk", 0.08, "oran", "Fiyatlama politikası", "01.01.2026", "ZAM YAP eşiği"),
        ("egk_esikKatkiHedef", 0.15, "oran", "Fiyatlama politikası", "01.01.2026", "Hedef katkı"),
        ("egk_azamiTutar", 1000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul tutar"),
        ("egk_ikincilUzunluk", 12, "karakter", "Anahtar stratejisi SPEC", "01.01.2026", "Ikincil anahtar"),
        ("egk_olcekHedef", 50000, "satir", "Manda A2 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("egk_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("egk_komisyonSenaryo", 0.02, "oran", "M_SEN parametre", "01.01.2026", "Komisyon +pp senaryo"),
        ("egk_fiyatModeli", "Tek seferlik alım — aylık ücret yok", "metin", "Ürün konumlandırma", "01.01.2026", "R3 ayrım"),
        ("egk_kvkkBeyani", "Veri cihazdan çıkmaz", "metin", "KVKK beyanı", "01.01.2026", "R3 ayrım"),
        ("egk_basabasFiyat", "Başabaş fiyat motoru", "metin", "MOTOR", "01.01.2026", "R1 ayrım"),
        ("egk_normAnahtar", "UPPER TRIM SUBSTITUTE", "metin", "Ç7 anahtar", "01.01.2026", "Serbest alternatif"),
        ("egk_tanimsizKova", "TANIMSIZ görünür", "metin", "E04", "01.01.2026", "Serbest alternatif"),
        ("egk_kovaButunlukAd", "Dört kova bütünlük", "metin", "E02", "01.01.2026", "Serbest alternatif"),
        ("egk_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("egk_firmaUnvan", "Örnek Ticaret A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026", "Firma"),
        ("egk_kuralSeti", "EGK-001..005", "metin", "Kök neden kuralları", "01.01.2026", "Kural seti"),
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
    motor_c = [
        (29, "egk_mutabakatFarki",
         '=IFERROR(SUM(tblVeriA[HedefNet])-SUM(tblMotor[NetKar]),0)'),
        (30, "egk_eslesmeOrani",
         '=IF(COUNTA(tblEslesme[A_Norm])=0,0,'
         '(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/COUNTA(tblEslesme[A_Norm]))'),
        (31, "egk_kovaButunluk", "=ESLESTIRME!C2"),
        (32, "egk_modulT1",
         '=IFERROR(IF(COUNTA(tblMotor[A_Norm])=0,0,AVERAGE(tblMotor[KatkiPayi])),0)'),
        (33, "egk_modulT2",
         '=IFERROR(IF(COUNTA(tblMotor[A_Norm])=0,0,AVERAGE(tblMotor[KomisyonOran])),0)'),
        (34, "egk_modulO1",
         '=COUNTIF(tblEslesme[EsikUstu],1)'),
        (35, "egk_modulO2",
         '=IFERROR(MAX(PANO!K28:K32),0)'),
        (36, "egk_modulO3",
         '=IFERROR(SUM(tblMotor[NetKar])*(1-egk_komisyonSenaryo),0)'),
        (37, "egk_modulO6",
         '=IFERROR(IF(COUNTA(tblVeriA[Platform])=0,0,'
         'COUNTIF(tblVeriA[Platform],INDEX(tblVeriA[Platform],1))/COUNTA(tblVeriA[Platform])),0)'),
        (38, "egk_modulO8",
         '=IF(COUNTA(tblVeriA[UrunKodu])=0,0,'
         'ROUND(100*(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/MAX(COUNTA(tblEslesme[A_Norm]),1),0))'),
        (39, "egk_modulMSen",
         '=IFERROR(SUM(tblMotor[KomisyonTutar])*egk_komisyonSenaryo,0)'),
        (40, "egk_modulI1",
         '=IFERROR(PERCENTILE(tblMotor[NetKar],egk_yuzdelikOran),0)'),
        (41, "egk_modulI2",
         '=IFERROR(PERCENTILE(tblMotor[NetKar],0.9),0)'),
    ]
    for r, ad, form in motor_c:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kalin=True)

    h(ws, 43, 8, "Canli Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorum_satir = [
        (44, "egk_modulT1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Ort katkı",TEXT(egk_modulT1,"0.0%")),"-")'),
        (45, "egk_modulT2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Ort komisyon",TEXT(egk_modulT2,"0.0%")),"-")'),
        (46, "egk_modulO1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Anomali",TEXT(egk_modulO1,"0")),"-")'),
        (47, "egk_modulO2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tornado",TEXT(egk_modulO2,"0.00"),"TL"),"-")'),
        (48, "egk_modulO3Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Senaryo net",TEXT(egk_modulO3,"0.00"),"TL"),"-")'),
        (49, "egk_modulO6Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma",TEXT(egk_modulO6,"0.0%")),"-")'),
        (50, "egk_modulO8Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Kalite",TEXT(egk_modulO8,"0")),"-")'),
        (51, "egk_modulMSenYorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Komisyon senaryo etki",TEXT(egk_modulMSen,"0.00"),"TL"),"-")'),
        (52, "egk_modulI1Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tahmin",TEXT(egk_modulI1,"0.00"),"TL"),"-")'),
        (53, "egk_modulI2Yorum",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"P90",TEXT(egk_modulI2,"0.00"),"TL"),"-")'),
    ]
    for r, ad, form in yorum_satir:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 28, "B": 42, "C": 12, "D": 28, "E": 14, "F": 28, "H": 24, "I": 42})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    adimlar = [
        "VERI_A: Ürün kataloğunu (kod, fiyat, maliyet, hedef katkı) sarı hücrelere girin.",
        "VERI_B: Platform komisyon, kargo, iade, reklam ve ROAS değerlerini yapıştırın.",
        "ANAHTAR: Normalize SKU, uzunluk, ikincil ve duplike kolonları otomatik dolar.",
        "MOTOR: Brüt/net kâr, katkı payı, başabaş fiyat ve satır kararı hesaplanır.",
        "ESLESTIRME: Hedef net ↔ gerçek net dört kova denetimi.",
        "KOK_NEDEN: EGK-001..005 sınıflandırır; TANIMSIZ elle incelenir.",
        "PANO / KANIT_RAPORU: Kararı izleyin ve PDF yazdırın.",
    ]
    for i, m in enumerate(adimlar, 5):
        h(ws, i, 1, f"{i - 4}. {m}", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    h(ws, 13, 1, "Sık yapılan hatalar", kalin=True, yazi="B3261E")
    for i, m in enumerate([
        "SKU boşluklu bırakmak (normalize edilir ama kontrol edin)",
        "Komisyon oranını yüzde yerine tutar yazmak",
        "ROAS yerine yalnızca reklam tutarını doldurup diğerini boş bırakmak",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 26):
        dogrulama(ws, "textLength", "0", f"D{r}",
                  baslik="Kaynak", mesaj="Kaynak metni girin",
                  hata_baslik="Kaynak", hata_mesaj="Boş bırakmayın",
                  isaret="greaterThanOrEqual", f2="120")
    ws = wb["LISTELER"]
    for r in range(7, 12):
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Platform", mesaj="Ad girin",
                  hata_baslik="Ad", hata_mesaj="En az 1 karakter",
                  isaret="greaterThan", f2="40")
    for r in range(7, 10):
        dogrulama(ws, "textLength", "1", f"C{r}",
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
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"TUT"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"ZAM YAP"'], fill=sari))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"İNDİRİMİ KES"'], fill=kirmizi, font=kf))
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
        CellIsRule(operator="equal", formula=['"ZAM YAP"'], fill=sari))
    ws.conditional_formatting.add(f"P{ILK}:P{SON}",
        CellIsRule(operator="equal", formula=['"İNDİRİMİ KES"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"P{ILK}:P{SON}",
        CellIsRule(operator="equal", formula=['"TUT"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(f"K{ILK}:K{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["KOK_NEDEN"]
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"TANIMSIZ"'], fill=kirmizi, font=kf))
    for kod in ("EGK-001", "EGK-002", "EGK-003", "EGK-004", "EGK-005"):
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
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"TUT"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"ZAM YAP"'], fill=sari))
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))

    ws = wb["VERI_A"]
    ws.conditional_formatting.add(f"C{ILK}:C{SON}",
        CellIsRule(operator="greaterThan", formula=["10000"], fill=sari))
    ws.conditional_formatting.add(f"D{ILK}:D{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["VERI_B"]
    ws.conditional_formatting.add(f"C{ILK}:C{SON}",
        CellIsRule(operator="greaterThan", formula=["100"], fill=sari))
    ws.conditional_formatting.add(f"E{ILK}:E{SON}",
        CellIsRule(operator="greaterThan", formula=["500"], fill=sari))

    ws = wb["ANAHTAR"]
    ws.conditional_formatting.add(f"F{ILK}:F{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"N{ILK}:N{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))


def adlari_bagla(wb):
    for i, ana in enumerate([
        "toleransKurus", "raporTarihi", "egk_esikEslesme", "egk_esikFark",
        "egk_esikKatkiDusuk", "egk_esikKatkiHedef", "egk_azamiTutar",
        "egk_ikincilUzunluk", "egk_olcekHedef", "egk_yuzdelikOran",
        "egk_komisyonSenaryo", "egk_fiyatModeli", "egk_kvkkBeyani",
        "egk_basabasFiyat", "egk_normAnahtar", "egk_tanimsizKova",
        "egk_kovaButunlukAd", "egk_dosyaSurumu", "egk_firmaUnvan", "egk_kuralSeti",
    ]):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "egk_mutabakatFarki": 29, "egk_eslesmeOrani": 30, "egk_kovaButunluk": 31,
        "egk_modulT1": 32, "egk_modulT2": 33, "egk_modulO1": 34, "egk_modulO2": 35,
        "egk_modulO3": 36, "egk_modulO6": 37, "egk_modulO8": 38, "egk_modulMSen": 39,
        "egk_modulI1": 40, "egk_modulI2": 41,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "egk_modulT1Yorum": 44, "egk_modulT2Yorum": 45, "egk_modulO1Yorum": 46,
        "egk_modulO2Yorum": 47, "egk_modulO3Yorum": 48, "egk_modulO6Yorum": 49,
        "egk_modulO8Yorum": 50, "egk_modulMSenYorum": 51, "egk_modulI1Yorum": 52,
        "egk_modulI2Yorum": 53,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListePlatformlar", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeDurumlar", "LISTELER!$C$7:$C$9")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$E$7:$E$8")


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
    dosya_ad = "ETicaretGercekKarlilikFiyatlama.xlsx"
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
