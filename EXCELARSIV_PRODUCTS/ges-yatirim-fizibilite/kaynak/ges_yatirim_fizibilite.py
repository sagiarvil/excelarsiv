#!/usr/bin/env python3
"""
GES Yatırım Fizibilite Simülatörü — üretim betiği (A6 / manda v6).
Çatı/saha GES, tüketim eşleştirmeli gelir, CAPEX/OPEX, NPV/IRR, tornado.
Karar: YATIRIM YAPILIR / DİKKAT / YATIRIM YAPMA / VERİ YOK.
Kod: GYF-PRO · önek gyf_*
"""

from __future__ import annotations

import hashlib
import os
import shutil
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Font, PatternFill, Protection
from openpyxl.utils import get_column_letter

from excel_uretim.ortak import (
    CATI,
    DIS_REF_YEŞIL,
    GIRIS_SARI,
    GRİ,
    KOYU_LACIVERT,
    KRITIK,
    NORMAL,
    ORTA_MAVI,
    SIFRE,
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
    giris_hucreleri,
    h,
    sabitle,
    sayfa_hazirla,
    sayfa_koru,
    tablo_ekle,
    tablo_formullerini_hucrelere_yaz,
    yorum_ekle,
)
from ortak.mevzuat_motoru import (
    KURAL_BASLIK,
    MOTOR_BASLIK,
    VAKA_BASLIK,
    belirsizlik_beyanlari,
)

URUN_AD = "GES Yatırım Fizibilite Simülatörü"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 11)

# Demo: 100 kWp çatı — baz NPV pozitif, kötümser negatif
DEMO = {
    "sirket_adi": "Örnek Enerji Yatırım A.Ş.",
    "donem": "2026-Q3",
    "hesapYili": 2026,
    "proje_tipi": "CATI",
    "kurulu_guc": 100,
    "spesifik_uretim": 1400,
    "degradasyon": 0.005,
    "oz_tuketim": 0.70,
    "yillik_tuketim": 200_000,
    "perakende_tarife": 3.50,
    "satis_tarife": 1.20,
    "capex_birim": 12_000,
    "opex_oran": 0.015,
    "wacc": 0.18,
    "npv_donem": 25,
    "kredi_orani": 0.28,
    "oz_kaynak_oran": 0.40,
}


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: GES fizibilite ve yatırım kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: NPV/IRR ve karar bozulur | Kaynak: EPDK tarife / saha ölçümü",
    )


def kural_cek(kural_id: str) -> str:
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-gyf_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(gyf_yuvarMax,IFERROR(N(G12),0)))"


def _guvenli_formul(formul: str, birim: str) -> str:
    if not formul.startswith("="):
        formul = "=" + formul
    inner = formul[1:]
    if inner.upper().startswith("IFERROR("):
        return formul
    yedek = '""' if birim in ("metin", "kod") else "0"
    return f"=IFERROR({inner},{yedek})"


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=40)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=40)
    h(ws, 4, 1,
      "Çatı veya saha GES kapasitesi, tüketim eşleştirme, CAPEX/OPEX ve tarife girdilerini "
      "işleyin; NPV/IRR aralığı ile YATIRIM YAPILIR / YATIRIM YAPMA kararını üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Tüketim eşleştirmeli gelir: öz tüketim tasarrufu + şebeke satış",
        "CAPEX / OPEX ayrımı + yıllık degradasyonlu nakit akışı",
        "Öz kaynak / kredi finansman senaryosu (F01)",
        "NPV / IRR aralık + tornado; kötümser YATIRIM YAPMA yolu (F04–F05)",
        "Varsayım kaynak+güven ve KANIT_RAPORU (banka/yatırımcı)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Tesis sahibi / mühendis — kapasite ve tüketim eşleştirmesini girmek",
        "Yatırımcı / CFO — NPV/IRR ile yatırım kararı vermek",
        "Banka kredi komitesi — imzalı fizibilite kanıtını dosyalamak",
    ], 14):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) HIZLI_BASLANGIC → 2) GIRDI sarı hücreler → 3) PANO/KARAR → 4) KANIT_RAPORU yazdır.",
      yazi="333333", kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 21, 1, f"Sürüm {SURUM} | 2026 | ExcelArşiv | Lisans: Tek kullanıcı | Kod: GYF-PRO",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 72})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Üç adımda sonuç", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Girdiler", "GIRDI'de proje tipi, kWp, özgül üretim, öz tüketim ve tarifeleri sarı hücrelere girin."),
        ("Adım 2 — Karar", "PANO ve KARAR'da NPV/IRR aralığı ve YATIRIM YAPILIR / YATIRIM YAPMA görün."),
        ("Adım 3 — Kanıt", "KANIT_RAPORU'nu yazdırın; FINANSMAN ve SENARYO ile doğrulayın."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, metin, kaydir=True, yazi="333333")
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=12)
    h(ws, 9, 1, "Formüller", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 1,
      "Üretim = kWp × özgül × (1−degradasyon)^yıl | Gelir = öz tüketim×perakende + satış×şebeke | "
      "NPV = Σ(net/(1+WACC)^t) − CAPEX",
      yazi="333333", kaydir=True)
    ws.merge_cells("A10:L10")
    h(ws, 12, 1, "Sarı = manuel giriş | Yeşil = formül çıktısı | Şifre 1234 formül alanlarını korur.",
      yazi=GRİ, kaydir=True)
    genislik(ws, {"A": 28, "B": 70})


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "GES proje girdileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Tarifeler KURALLAR'dan ayna ile doğrulanır.",
      yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Şirket / Proje", DEMO["sirket_adi"], None, "sirket_adi", "Unvan veya proje kodu."),
        (7, "Dönem", DEMO["donem"], None, "donem", "Rapor dönemi."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili", "2025, 2026 veya 2027."),
        (9, "Proje Tipi", DEMO["proje_tipi"], None, "proje_tipi", "CATI veya SAHA."),
        (10, "Kurulu güç [kWp]", DEMO["kurulu_guc"], CATI, "kurulu_guc", "DC kurulu güç."),
        (11, "Özgül üretim [kWh/kWp]", DEMO["spesifik_uretim"], CATI, "spesifik_uretim", "Yıllık özgül üretim."),
        (12, "Degradasyon [%/yıl]", DEMO["degradasyon"], YÜZDE, "degradasyon", "Yıllık panel kaybı."),
        (13, "Öz tüketim oranı [%]", DEMO["oz_tuketim"], YÜZDE, "oz_tuketim", "Üretimin öz tüketim payı."),
        (14, "Yıllık tüketim [kWh]", DEMO["yillik_tuketim"], CATI, "yillik_tuketim", "Tesis yıllık elektrik tüketimi."),
        (15, "Perakende tarife [₺/kWh]", DEMO["perakende_tarife"], "0.00", "perakende_tarife", "Öz tüketim birim değeri."),
        (16, "Şebeke satış tarife [₺/kWh]", DEMO["satis_tarife"], "0.00", "satis_tarife", "YEKDEM/şebeke satış."),
        (17, "CAPEX birim [₺/kWp]", DEMO["capex_birim"], TL, "capex_birim", "Yatırım birim maliyeti."),
        (18, "OPEX oranı [% CAPEX]", DEMO["opex_oran"], YÜZDE, "opex_oran", "Yıllık işletme oranı."),
        (19, "WACC [%]", DEMO["wacc"], YÜZDE, "wacc", "İskonto oranı (nominal)."),
        (20, "NPV dönem [yıl]", DEMO["npv_donem"], CATI, "npv_donem", "Nakit akışı yılı."),
        (21, "Kredi faiz oranı [%]", DEMO["kredi_orani"], YÜZDE, "kredi_orani", "Kredi senaryosu faiz."),
        (22, "Öz kaynak oranı [%]", DEMO["oz_kaynak_oran"], YÜZDE, "oz_kaynak_oran", "Öz kaynak payı."),
    ]
    h(ws, 5, 1, "Alan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 2, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 3, "Birim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    for satir, etiket, deger, sayi, _ad, mesaj in alanlar:
        if sayi == TL:
            birim = "₺"
        elif sayi == YÜZDE:
            birim = "%"
        elif _ad in ("hesapYili", "npv_donem"):
            birim = "yıl"
        elif _ad in ("kurulu_guc",):
            birim = "kWp"
        elif _ad in ("spesifik_uretim", "yillik_tuketim"):
            birim = "kWh"
        elif _ad in ("perakende_tarife", "satis_tarife"):
            birim = "₺/kWh"
        elif _ad == "capex_birim":
            birim = "₺/kWp"
        else:
            birim = "metin"
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim, yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 24, 1, "CAPEX / YATIRIM toplam (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, "=IFERROR(gyf_capexToplam,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 25, 1, "OPEX / ISLETME yıllık (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=IFERROR(gyf_opexYillik,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Yıllık gelir (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 2, "=IFERROR(gyf_yillikGelir,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 27, 1, "Giriş doluluk", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B22)=0,gyf_girisBeklenen=0),0,ROUND(COUNTA(B6:B22)/gyf_girisBeklenen*100,0)),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)

    h(ws, 29, 1, "Giriş tablosu (otomasyon / boş-mod)", kalin=True, yazi=KOYU_LACIVERT)
    sutunlar = ["Alan Anahtarı", "Değer", "Birim", "Zorunlu", "Kaynak", "Kontrol"]
    baslik_satiri(ws, 30, sutunlar)
    formuller = {
        "Kontrol": (
            '=IF(tblGirdi[[#This Row],[Alan Anahtarı]]="","",'
            'IF(tblGirdi[[#This Row],[Değer]]="","EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblGirdi", "A30:F1029", sutunlar, formuller)
    ornek_satir = [
        ("sirket_adi", DEMO["sirket_adi"], "metin", "Evet", "GIRDI"),
        ("donem", DEMO["donem"], "metin", "Evet", "GIRDI"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Evet", "GIRDI"),
        ("proje_tipi", DEMO["proje_tipi"], "kod", "Evet", "GIRDI"),
        ("kurulu_guc", DEMO["kurulu_guc"], "kWp", "Evet", "GIRDI"),
        ("spesifik_uretim", DEMO["spesifik_uretim"], "kWh/kWp", "Evet", "GIRDI"),
        ("degradasyon", DEMO["degradasyon"], "oran", "Evet", "GIRDI"),
        ("oz_tuketim", DEMO["oz_tuketim"], "oran", "Evet", "GIRDI"),
        ("yillik_tuketim", DEMO["yillik_tuketim"], "kWh", "Evet", "GIRDI"),
        ("perakende_tarife", DEMO["perakende_tarife"], "TL/kWh", "Evet", "GIRDI"),
        ("satis_tarife", DEMO["satis_tarife"], "TL/kWh", "Evet", "GIRDI"),
        ("capex_birim", DEMO["capex_birim"], "TL/kWp", "Evet", "GIRDI"),
        ("opex_oran", DEMO["opex_oran"], "oran", "Evet", "GIRDI"),
        ("wacc", DEMO["wacc"], "oran", "Evet", "GIRDI"),
        ("npv_donem", DEMO["npv_donem"], "yıl", "Evet", "GIRDI"),
        ("kredi_orani", DEMO["kredi_orani"], "oran", "Evet", "GIRDI"),
        ("oz_kaynak_oran", DEMO["oz_kaynak_oran"], "oran", "Evet", "GIRDI"),
    ]
    tl_a = {"capex_birim"}
    oran_a = {"degradasyon", "oz_tuketim", "opex_oran", "wacc", "kredi_orani", "oz_kaynak_oran"}
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 31):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in tl_a:
            ws.cell(i, 2).number_format = TL
        elif a in oran_a:
            ws.cell(i, 2).number_format = YÜZDE
        elif a in ("perakende_tarife", "satis_tarife"):
            ws.cell(i, 2).number_format = "0.00"
        elif a in ("hesapYili", "kurulu_guc", "spesifik_uretim", "yillik_tuketim", "npv_donem"):
            ws.cell(i, 2).number_format = CATI
        ws.cell(i, 3).value = b
        ws.cell(i, 4).value = z
        ws.cell(i, 5).value = k
        for kolon in range(1, 6):
            ws.cell(i, kolon).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
            ws.cell(i, kolon).protection = Protection(locked=False)
            _kim(ws, f"{get_column_letter(kolon)}{i}", a, "Tablo girişi; PANO ile senkron tutun")

    dogrulama(ws, "list", "ListeYillar", "B8",
              baslik="Hesap Yılı", mesaj="2025, 2026 veya 2027 seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Yalnızca listeden seçin.", bos=False)
    dogrulama(ws, "list", "ListeProjeTipi", "B9",
              baslik="Proje Tipi", mesaj="CATI veya SAHA seçin.",
              hata_baslik="Geçersiz", hata_mesaj="CATI veya SAHA.", bos=False)
    dogrulama(ws, "list", "ListeYillar", "B33",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeProjeTipi", "B34",
              baslik="Proje Tipi", mesaj="CATI veya SAHA.",
              hata_baslik="Geçersiz", hata_mesaj="CATI/SAHA.")
    for aralik, baslik, mesaj in [
        ("B10", "Kurulu güç", "0 ile 1e6 kWp."),
        ("B11", "Özgül üretim", "0 ile 3000."),
        ("B14", "Yıllık tüketim", "0 ile 1e9 kWh."),
        ("B15", "Perakende tarife", "0 ile 100 ₺/kWh."),
        ("B16", "Satış tarife", "0 ile 100 ₺/kWh."),
        ("B17", "CAPEX birim", "0 ile 1e7 ₺/kWp."),
        ("B35", "Kurulu güç", "Tablo: kWp."),
        ("B36", "Özgül", "Tablo: kWh/kWp."),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj=mesaj,
                  hata_baslik="Geçersiz tutar", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000")
    for aralik, baslik in [("B12", "Degradasyon"), ("B13", "Öz tüketim"), ("B18", "OPEX"),
                           ("B19", "WACC"), ("B21", "Kredi"), ("B22", "Öz kaynak"),
                           ("B37", "Degradasyon"), ("B38", "Öz tüketim"), ("B43", "OPEX"),
                           ("B44", "WACC"), ("B46", "Kredi"), ("B47", "Öz kaynak")]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj="0 ile 1 arasında oran.",
                  hata_baslik="Geçersiz", hata_mesaj="0-1 arası.",
                  isaret="between", f2="1")
    for aralik, baslik in [("B20", "NPV dönem"), ("B45", "NPV dönem")]:
        dogrulama(ws, "whole", "1", aralik,
                  baslik=baslik, mesaj="1-25 arası.",
                  hata_baslik="Geçersiz", hata_mesaj="1-25.",
                  isaret="between", f2="25")

    giris_hucreleri(ws, 6, 22, [2])
    giris_hucreleri(ws, 31, 47, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 40, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1031, "Sarı alanlar giriş; tarife ve formüller kilitlidir.")


def varsayimlar(ws):
    sayfa_hazirla(ws, "VARSAYIMLAR", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Model varsayımları — her satırda kaynak ve güven", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:H3")
    sutunlar = ["varsayim_id", "aciklama", "deger", "birim", "kaynak", "guven", "duyarlilik", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    satirlar = [
        ("V-01", "Özgül üretim kaynağı", "PVGIS/ölçüm", "metin", "PVGIS / saha ölçümü", "sert", "yüksek"),
        ("V-02", "Degradasyon", "0.5%/yıl", "oran", "Panel üretici veri sayfası", "sert", "orta"),
        ("V-03", "Öz tüketim oranı", "tesis profili", "oran", "Sayaç / fatura analizi", "yumuşak", "yüksek"),
        ("V-04", "Perakende tarife", "EPDK dağıtım", "TL/kWh", "EPDK tarife tablosu", "sert", "yüksek"),
        ("V-05", "Şebeke satış tarife", "YEKDEM/mahalle", "TL/kWh", "YEKDEM / iç politika", "yumuşak", "yüksek"),
        ("V-06", "CAPEX birim", "teklif ortalaması", "TL/kWp", "EPC teklifleri", "yumuşak", "yüksek"),
        ("V-07", "OPEX oranı", "% CAPEX", "oran", "İşletme bütçesi", "yumuşak", "orta"),
        ("V-08", "WACC / iskonto", "nominal", "oran", "Finans komitesi", "yumuşak", "yüksek"),
        ("V-09", "Para dünyası", "nominal", "metin", "MOTOR para_dunyasi", "sert", "orta"),
        ("V-10", "Kötümser üretim şoku", "−25%", "oran", "Stres testi", "yumuşak", "yüksek"),
    ]
    for i, row in enumerate(satirlar, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v, yazi="333333",
              zemin=GIRIS_SARI if k in (3, 5, 6) else None)
            if k in (3, 5, 6):
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", row[1], "Varsayım; kaynak ve guven zorunlu")
    formuller = {
        "kontrol": (
            '=IF(tblVarsayimlar[[#This Row],[varsayim_id]]="","",'
            'IF(OR(tblVarsayimlar[[#This Row],[kaynak]]="",'
            'tblVarsayimlar[[#This Row],[guven]]=""),"EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblVarsayimlar", "A5:H15", sutunlar, formuller)
    h(ws, 17, 10, "CAPEX / YATIRIM tutarı (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 11, "=IFERROR(gyf_capexToplam,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 10, "OPEX / ISLETME yıllık (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 11, "=IFERROR(gyf_opexYillik,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 14, "B": 28, "C": 18, "D": 10, "E": 28, "F": 12, "G": 12, "H": 10, "J": 32, "K": 18})


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Tarife / teşvik kural tablosu (yıl yan yana)", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["GYF-001", "EPDK tarife", "perakende_tarife", "öz tüketim × tarife", "perakende_tl_kwh",
         3.20, 3.50, 3.80, "01.01.2025", "EPDK dağıtım bedelli"],
        ["GYF-002", "YEKDEM / satış", "satis_tarife", "şebeke satış × tarife", "satis_tl_kwh",
         1.10, 1.20, 1.30, "01.01.2025", "YEKDEM referans"],
        ["GYF-003", "İç politika", "hedef_geri_odeme", "geri ödeme ≤ eşik", "geri_odeme_yil",
         8, 8, 7, "01.01.2025", "Hedef geri ödeme yılı"],
        ["GYF-004", "İç politika", "min_oz_tuketim", "öz tüketim ≥ eşik", "min_oz_tuketim",
         0.40, 0.40, 0.45, "01.01.2025", "Asgari öz tüketim"],
        ["GYF-005", "Uygulama", "yuvarlama", "her hesap", "yuvarlama_ondalik",
         2, 2, 2, "01.01.2025", "Yuvarlama ondalık"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = None
            if k in (6, 7, 8):
                if s[0] in ("GYF-001", "GYF-002"):
                    sayi = "0.00"
                elif s[0] == "GYF-003":
                    sayi = CATI
                elif s[0] == "GYF-004":
                    sayi = YÜZDE
                else:
                    sayi = CATI
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-gyf_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A14:B14")
    h(ws, 15, 1, "Perakende (aktif yıl)", yazi="333333", kaydir=True)
    ws.merge_cells("A15:B15")
    h(ws, 15, 3, kural_cek("GYF-001"), sayi="0.00", yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Satış tarife (aktif yıl)", yazi="333333", kaydir=True)
    ws.merge_cells("A16:B16")
    h(ws, 16, 3, kural_cek("GYF-002"), sayi="0.00", yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 3,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"GYF-001=",TEXT(C15,"0.00"),"GYF-002=",TEXT(C16,"0.00"),'
      '"yıl=",IF(OR(hesapYili="",hesapYili=0),"-",hesapYili)),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    for col in ("F", "G", "H"):
        dogrulama(ws, "decimal", "0", f"{col}7:{col}11",
                  baslik="Kural değeri", mesaj="Yıl kolonuna sayısal parametre girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ve üzeri sayı.",
                  isaret="between", f2="10000000")
    genislik(ws, {get_column_letter(i): w for i, w in enumerate(
        [18, 22, 16, 28, 16, 12, 12, 12, 12, 18, 14], 1)})
    ws.merge_cells("A3:K3")
    ws.merge_cells("A4:K4")
    sabitle(ws, "A7")


def motor(ws):
    sayfa_hazirla(ws, "MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Hesap zinciri — her adımda kural_id · para_dunyasi=nominal", kalin=True,
      boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Tarifeler tblKurallar'dan; girdiler GIRDI'den.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")
    basliklar = [*MOTOR_BASLIK, "deger"]
    baslik_satiri(ws, 6, basliklar)

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif kural yılı", "GYF-005")
    ekle("Perakende tarife kural", km("GYF-001"), "TL/kWh", "EPDK perakende", "GYF-001")
    ekle("Satış tarife kural", km("GYF-002"), "TL/kWh", "YEKDEM/satış", "GYF-002")
    ekle("Hedef geri ödeme", km("GYF-003"), "yıl", "Geri ödeme eşiği", "GYF-003")
    ekle("Min öz tüketim", km("GYF-004"), "oran", "Asgari öz tüketim", "GYF-004")
    ekle("Yuvarlama ondalık", km("GYF-005"), "adet", "Yuvarlama basamağı", "GYF-005")
    ekle("Kurulu güç kWp", "=GIRDI!B10", "kWp", "DC güç", "GYF-005")
    ekle("Özgül üretim", "=GIRDI!B11", "kWh/kWp", "Yıllık özgül", "GYF-005")
    ekle("Degradasyon", "=GIRDI!B12", "oran", "Yıllık kayıp", "GYF-005")
    ekle("Öz tüketim oranı", "=GIRDI!B13", "oran", "Öz tüketim payı", "GYF-004")
    ekle("Yıllık tüketim", "=GIRDI!B14", "kWh", "Tesis tüketimi", "GYF-004")
    ekle("Perakende giriş", "=GIRDI!B15", "TL/kWh", "Giriş perakende", "GYF-001")
    ekle("Satış giriş", "=GIRDI!B16", "TL/kWh", "Giriş satış", "GYF-002")
    ekle("CAPEX birim", "=GIRDI!B17", "TL/kWp", "Birim yatırım", "GYF-005")
    ekle("OPEX oranı", "=GIRDI!B18", "oran", "İşletme oranı", "GYF-005")
    ekle("WACC", "=GIRDI!B19", "oran", "İskonto nominal", "GYF-005")
    ekle("NPV dönem", "=GIRDI!B20", "adet", "Akış yılı", "GYF-005")
    ekle("Kredi faiz", "=GIRDI!B21", "oran", "Kredi senaryosu", "GYF-005")
    ekle("Öz kaynak oranı", "=GIRDI!B22", "oran", "Öz kaynak payı", "GYF-005")
    ekle("Yıllık üretim Y1", f"=ROUND(G13*G14,{_YV})", "kWh", "kWp×özgül", "GYF-005")
    ekle("Öz tüketim kWh", f"=ROUND(G26*G16,{_YV})", "kWh", "Üretim×öz oran", "GYF-004")
    ekle("Şebeke satış kWh", "=G26-G27", "kWh", "Kalan üretim", "GYF-002")
    ekle("Tüketim eşleşme kWh", f"=ROUND(MIN(G27,G17),{_YV})", "kWh", "min(öz,tüketim)", "GYF-004")
    ekle("Öz tasarruf geliri", f"=ROUND(G29*G18,{_YV})", "TL", "eşleşme×perakende", "GYF-001")
    ekle("Şebeke satış geliri", f"=ROUND(G28*G19,{_YV})", "TL", "satış×tarife", "GYF-002")
    ekle("Yıllık gelir baz", "=G30+G31", "TL", "Toplam gelir Y1", "GYF-001")
    ekle("CAPEX / YATIRIM", f"=ROUND(G13*G20,{_YV})", "TL", "kWp×birim CAPEX", "GYF-005")
    ekle("OPEX / ISLETME yıllık", f"=ROUND(G33*G21,{_YV})", "TL", "CAPEX×oran", "GYF-005")
    ekle("Net nakit Y1", "=G32-G34", "TL", "Gelir−OPEX", "GYF-005")
    ekle("Geri ödeme yılı",
         "=IF(OR(G35<=0,G33<=0),gyf_geriOdemeSonsuz,ROUND(G33/G35,1))",
         "yıl", "CAPEX/net", "GYF-003")
    ekle("NPV (baz)",
         f'=IF(OR(G23<=0,G22=""),0,IF(G22=0,ROUND(G35*G23-G33,{_YV}),'
         f'ROUND(G35*(1-POWER(1+G22,-G23))/G22-G33,{_YV})))',
         "TL", "Anüite NPV − CAPEX", "GYF-005")
    ekle("IRR yaklaşımı",
         '=IF(OR(G33<=0,G35<=0,G23<=0),0,ROUND(G35/G33-gyf_bazCarpan/G23,4))',
         "oran", "Kabaca IRR", "GYF-005")
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(OR(G33<=0,G13<=0,G32<=0),"YAPMA",'
         'IF(OR(G37<0,G36>G10*gyf_dikkatCarpan),"YAPMA",'
         'IF(OR(G36>G10,G37<gyf_npvDikkat),"DIKKAT","YAPILIR"))))',
         "kod", "Karar kodu", "GYF-003")
    ekle("Karar metni",
         '=IF(G39="VERI_YOK","VERİ YOK",'
         'IF(G39="YAPILIR","YATIRIM YAPILIR",'
         'IF(G39="DIKKAT","DİKKAT","YATIRIM YAPMA")))',
         "metin", "Kullanıcıya karar", "GYF-003")
    ekle("Gerekçe",
         '=IF(G39="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G39="YAPILIR","NPV pozitif ve geri ödeme hedef altında; GES yatırımı uygun.",'
         'IF(G39="DIKKAT","Geri ödeme veya NPV dikkat bandında; tarife/üretimi gözden geçirin.",'
         '"NPV negatif veya geri ödeme aşırı; kötümser/bazda YATIRIM YAPMA.")))',
         "metin", "Gerekçe cümlesi", "GYF-003")
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B22)=0,gyf_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B22)/gyf_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "GYF-005")
    ekle("Risk skoru",
         '=IF(G39="VERI_YOK",0,IF(G39="YAPILIR",gyf_riskDusuk,'
         'IF(G39="DIKKAT",gyf_riskOrta,gyf_riskYuksek)))',
         "puan", "Yatırım riski", "GYF-005")
    ekle("İyimser üretim", f"=ROUND(G26*gyf_senaryoIyi,{_YV})", "kWh", "İyimser üretim", "GYF-005")
    ekle("Kötümser üretim", f"=ROUND(G26*gyf_senaryoKotu,{_YV})", "kWh", "Kötümser üretim", "GYF-005")
    ekle("İyimser gelir",
         f"=ROUND(G44*G16*G18+(G44-G44*G16)*G19,{_YV})",
         "TL", "İyimser gelir", "GYF-001")
    ekle("Kötümser gelir",
         f"=ROUND(G45*G16*G18*gyf_tarifeKotu+(G45-G45*G16)*G19*gyf_tarifeKotu,{_YV})",
         "TL", "Kötümser gelir", "GYF-002")
    ekle("İyimser NPV",
         f'=IF(G22=0,ROUND((G46-G34)*G23-G33*gyf_capexIyi,{_YV}),'
         f'ROUND((G46-G34)*(1-POWER(1+G22,-G23))/G22-G33*gyf_capexIyi,{_YV}))',
         "TL", "İyimser NPV", "GYF-005")
    ekle("Kötümser NPV",
         f'=IF(G22=0,ROUND((G47-G34*gyf_opexKotu)*G23-G33*gyf_capexKotu,{_YV}),'
         f'ROUND((G47-G34*gyf_opexKotu)*(1-POWER(1+G22,-G23))/G22-G33*gyf_capexKotu,{_YV}))',
         "TL", "Kötümser NPV", "GYF-005")
    ekle("Kötümser karar",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(G49<0,"YATIRIM YAPMA","DİKKAT"))',
         "metin", "Kötümser YATIRIM YAPMA yolu", "GYF-003")
    ekle("Tornado: tarife etkisi",
         f"=ABS(ROUND(G32*gyf_tornadoTarife-G32,{_YV}))",
         "TL", "Tarife ± etki", "GYF-001")
    ekle("Tornado: üretim etkisi",
         f"=ABS(ROUND(G32*gyf_tornadoUretim-G32,{_YV}))",
         "TL", "Üretim ± etki", "GYF-005")
    ekle("Tornado: CAPEX etkisi",
         f"=ABS(ROUND(G33*gyf_tornadoCapex-G33,{_YV}))",
         "TL", "CAPEX ± etki", "GYF-005")
    ekle("Tornado: OPEX etkisi",
         f"=ABS(ROUND(G34*gyf_tornadoOpex-G34,{_YV}))",
         "TL", "OPEX ± etki", "GYF-005")
    ekle("Tornado: WACC etkisi",
         f"=ABS(ROUND(G37*gyf_tornadoWacc-G37,{_YV}))",
         "TL", "WACC ± etki", "GYF-005")
    ekle("Öz tüketim uyumu",
         "=IF(G17=0,0,G29/G17)",
         "oran", "Eşleşme/tüketim", "GYF-004")
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"GYF-001→perakende","GYF-002→satış","GYF-003→geri ödeme","GYF-004→öz tüketim")',
         "metin", "Kanıt atıfları", "GYF-001")
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"CAPEX=",TEXT(G33,"₺ #,##0"),"NPV=",TEXT(G37,"₺ #,##0"),'
         '"IRR=",TEXT(G38,"0.0%"),"Karar=",G40)',
         "metin", "Rapor özeti", "GYF-005")
    ekle("Yıllık gelir (pano)", "=G32", "TL", "Gelir ayna", "GYF-001")
    ekle("CAPEX (pano)", "=G33", "TL", "CAPEX ayna", "GYF-005")
    ekle("OPEX (pano)", "=G34", "TL", "OPEX ayna", "GYF-005")
    ekle("NPV (pano)", "=G37", "TL", "NPV ayna", "GYF-005")
    ekle("IRR (pano)", "=G38", "oran", "IRR ayna", "GYF-005")
    ekle("Tüketim eşleşme (pano)", "=G29", "kWh", "Eşleşme ayna", "GYF-004")

    assert len(adimlar) >= 40, len(adimlar)

    for i, (ad, formul, birim, ne, kid) in enumerate(adimlar, 1):
        r = 6 + i
        guvenli = _guvenli_formul(formul, birim)
        h(ws, r, 1, i, sayi=CATI, hiza="center")
        h(ws, r, 2, ad, kaydir=True)
        h(ws, r, 3, guvenli, yazi=DIS_REF_YEŞIL, kaydir=True)
        h(ws, r, 4, birim, hiza="center")
        h(ws, r, 5, ne, kaydir=True, yazi="333333")
        h(ws, r, 6, kid, hiza="center", kalin=True, yazi="B08948")
        h(ws, r, 7, guvenli, yazi=DIS_REF_YEŞIL)
        if birim == "TL":
            ws.cell(r, 7).number_format = TL
            ws.cell(r, 3).number_format = TL
        elif birim in ("oran",):
            ws.cell(r, 7).number_format = YÜZDE
            ws.cell(r, 3).number_format = YÜZDE
        elif birim in ("TL/kWh",):
            ws.cell(r, 7).number_format = "0.00"
            ws.cell(r, 3).number_format = "0.00"
        elif birim in ("puan", "yıl", "adet", "kWp", "kWh", "kWh/kWp", "TL/kWp"):
            ws.cell(r, 7).number_format = CATI if birim != "TL/kWp" else TL
            ws.cell(r, 3).number_format = ws.cell(r, 7).number_format

    genislik(ws, {"A": 8, "B": 36, "C": 55, "D": 12, "E": 28, "F": 12, "G": 22})
    sabitle(ws, "A7")
    return len(adimlar)


def finansman(ws):
    sayfa_hazirla(ws, "FINANSMAN", "C65911", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Finansman senaryoları — OZ_KAYNAK ve KREDI", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:E3")
    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Pay", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Tutar [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Faiz / maliyet", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "NPV etkisi", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    h(ws, 6, 1, "OZ_KAYNAK", kalin=True, yazi="333333")
    h(ws, 6, 2, "=GIRDI!B22", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 3, "=IFERROR(gyf_capexToplam*B6,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "=IFERROR(GIRDI!B22*gyf_sifirCarpan,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 5, "=IFERROR(MOTOR!G37,0)", sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 7, 1, "KREDI", kalin=True, yazi="333333")
    h(ws, 7, 2, "=gyf_bazCarpan-GIRDI!B22", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 3, "=IFERROR(gyf_capexToplam*B7,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 4, "=GIRDI!B21", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 5,
      "=IFERROR(MOTOR!G37-C7*D7*GIRDI!B20,0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 9, 1, "CAPEX / YATIRIM toplam", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 9, 2, "=IFERROR(gyf_capexToplam,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "OPEX / ISLETME yıllık", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 2, "=IFERROR(gyf_opexYillik,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Finansman karşılaştırma", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"OZ_KAYNAK NPV ",TEXT(E6,"₺ #,##0"),'
      '" · KREDI NPV ",TEXT(E7,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B11:H11")

    h(ws, 13, 1, "FinansDemo", kalin=True, yazi=GRİ)
    h(ws, 14, 1, "OZ_KAYNAK")
    h(ws, 14, 2, 480_000, sayi=TL, yazi=GRİ)
    h(ws, 15, 1, "KREDI")
    h(ws, 15, 2, 720_000, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Finansman Payları"
    g.add_data(Reference(ws, min_col=2, min_row=13, max_row=15), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=14, max_row=15))
    g.height, g.width = 7, 10
    ws.add_chart(g, "D13")
    genislik(ws, {"A": 32, "B": 14, "C": 16, "D": 14, "E": 16})


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve duyarlılık (tornado)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)

    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Çarpan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Yıllık gelir", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "NPV", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(gyf_senaryoIyi+N(GIRDI!B10)*gyf_sifirCarpan,0)",
         "=IFERROR(MOTOR!G46,0)", "=IFERROR(MOTOR!G48,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D6<0,"YATIRIM YAPMA","YATIRIM YAPILIR"))'),
        ("Baz", "=IFERROR(gyf_bazCarpan+N(GIRDI!B10)*gyf_sifirCarpan,1)",
         "=IFERROR(gyf_yillikGelir,0)", "=IFERROR(gyf_npv,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(gyf_kararMetni,"-"))'),
        ("Kötümser", "=IFERROR(gyf_senaryoKotu+N(GIRDI!B10)*gyf_sifirCarpan,0)",
         "=IFERROR(MOTOR!G47,0)", "=IFERROR(MOTOR!G49,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(MOTOR!G50,"YATIRIM YAPMA"))'),
    ]
    for i, (ad, carp, gelir, npv, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, gelir, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, npv, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 10, 1, "Senaryo NPV bant (iyimser−kötümser)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 2, "=IFERROR(D6-D8,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser NPV ",TEXT(D6,"₺ #,##0")," · Baz ",TEXT(D7,"₺ #,##0"),'
      '" · Kötümser ",TEXT(D8,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B11:H11")

    h(ws, 13, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 14, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 14, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 14, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    tornado = [
        (15, "Tornado tarife etkisi", "=IFERROR(MOTOR!G51,0)"),
        (16, "Tornado üretim etkisi", "=IFERROR(MOTOR!G52,0)"),
        (17, "Tornado CAPEX etkisi", "=IFERROR(MOTOR!G53,0)"),
        (18, "Tornado OPEX etkisi", "=IFERROR(MOTOR!G54,0)"),
        (19, "Tornado WACC etkisi", "=IFERROR(MOTOR!G55,0)"),
    ]
    for r, ad, formul in tornado:
        h(ws, r, 1, ad, yazi="333333")
        h(ws, r, 2, formul, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 3,
          f'=IFERROR(IF(B{r}>=MAX(B15:B19),gyf_siraBir,IF(B{r}=LARGE(B15:B19,2),gyf_siraIki,'
          f'IF(B{r}=LARGE(B15:B19,3),gyf_siraUc,4))),4)',
          sayi=CATI, yazi=DIS_REF_YEŞIL)

    h(ws, 21, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 2,
      '=IFERROR(CONCATENATE("1)",INDEX(A15:A19,MATCH(gyf_siraBir,C15:C19,0)),'
      '" 2)",INDEX(A15:A19,MATCH(gyf_siraIki,C15:C19,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 22, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B15:B19),"₺ #,##0")," — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 24, 1, "Tahmin alt", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, "=IFERROR(gyf_npv*gyf_tahminAlt,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 24, 3, "=IFERROR(gyf_npv,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 24, 4, "=IFERROR(gyf_npv*gyf_tahminUst,0)", sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 26, 1, "Tahmin (PERCENTILE)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1, "Seri nokta", yazi=GRİ)
    for i, v in enumerate([1, 2, 3, 4], 28):
        h(ws, i, 1, v, sayi=CATI)
        h(ws, i, 2, f"=D{5+v}", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 32, 1, "Tahmin P90", yazi="333333")
    h(ws, 32, 2,
      '=IFERROR(PERCENTILE(B28:B31,gyf_yuzdelikOran),0)',
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Tahmin yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"P90 ",TEXT(B32,"₺ #,##0")," · Alt ",TEXT(B24,"₺ #,##0"),'
      '" · Üst ",TEXT(D24,"₺ #,##0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 5, 7, "NpvDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([1_800_000, 850_000, -180_000], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 14, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([220_000, 180_000, 150_000, 90_000, 70_000], 15):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)
    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([700_000, 850_000, 1_100_000, 850_000], 28):
        h(ws, i, 4, v, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo NPV Karşılaştırması"
    g1.add_data(Reference(ws, min_col=7, min_row=5, max_row=8), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=6, max_row=8))
    g1.height, g1.width = 8, 14
    ws.add_chart(g1, "F5")

    g2 = BarChart()
    g2.type = "bar"
    g2.title = "Duyarlılık Tornado"
    g2.add_data(Reference(ws, min_col=5, min_row=14, max_row=19), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=15, max_row=19))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "F18")

    g3 = LineChart()
    g3.title = "NPV Trend"
    g3.add_data(Reference(ws, min_col=4, min_row=27, max_row=31), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=28, max_row=31))
    g3.height, g3.width = 8, 12
    ws.add_chart(g3, "F32")

    genislik(ws, {"A": 36, "B": 18, "C": 18, "D": 18, "E": 22, "G": 14})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=40)
    baski_hazirla(ws, "A1:F34", f"{URUN_AD} · PANO")
    h(ws, 3, 1, "Karar panosu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpis = [
        (2, "Yıllık gelir", "=IFERROR(gyf_yillikGelir,0)", TL),
        (3, "CAPEX", "=IFERROR(gyf_capexToplam,0)", TL),
        (4, "OPEX", "=IFERROR(gyf_opexYillik,0)", TL),
        (5, "Üretim Y1", "=IFERROR(MOTOR!G26,0)", CATI),
        (6, "Eşleşme", "=IFERROR(gyf_tuketimEslesme,0)", CATI),
        (7, "NPV", "=IFERROR(gyf_npv,0)", TL),
        (8, "IRR", "=IFERROR(gyf_irr,0)", YÜZDE),
        (9, "Geri ödeme", "=IFERROR(MOTOR!G36,0)", CATI),
        (10, "Güven", "=IFERROR(MOTOR!G42,0)", CATI),
        (11, "Risk", "=IFERROR(MOTOR!G43,0)", CATI),
        (12, "İyimser NPV", "=IFERROR(MOTOR!G48,0)", TL),
        (13, "Kötümser NPV", "=IFERROR(MOTOR!G49,0)", TL),
        (14, "Tahmin P90", "=IFERROR(gyf_tahminAralik,0)", TL),
    ]
    for col, ad, formul, sayi in kpis:
        h(ws, 3, col, ad, kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
        h(ws, 4, col, formul, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "Karar özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=gyf_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=gyf_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 10, 1, "Kod", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 2, "Ad", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 3, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 4, "Yorum", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    moduller = [
        ("T1", "Yıllık üretim / degradasyon",
         "=IFERROR(MOTOR!G26,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Üretim Y1 ",TEXT(C11,"#,##0")," kWh — degradasyon ",TEXT(GIRDI!B12,"0.0%"))'),
        ("T2", "Tüketim eşleştirme geliri",
         "=IFERROR(gyf_yillikGelir,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Gelir ",TEXT(C12,"₺ #,##0")," — eşleşme ",TEXT(gyf_tuketimEslesme,"#,##0")," kWh")'),
        ("O1", "Senaryo NPV sapması",
         "=IFERROR(ABS(SENARYO!D6-SENARYO!D8),0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"İyimser−kötümser NPV bant ",TEXT(C13,"₺ #,##0"))'),
        ("O2", "Duyarlılık / tornado",
         '=IFERROR(gyf_duyarlilikSira,"-")',
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado sıra: ",IFERROR(gyf_yorumDuyarlilik,"-"))'),
        ("O3", "Senaryo motoru",
         '=IFERROR(gyf_senaryoKarsilastirma,"-")',
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo: ",IFERROR(gyf_yorumSenaryo,"-"))'),
        ("O8", "Veri kalite skoru",
         "=IFERROR(MOTOR!G42,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Güven skoru ",TEXT(C16,"0")," / 100")'),
        ("I1", "Tahmin + aralık",
         "=IFERROR(gyf_tahminAralik,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 tahmin ",TEXT(C17,"₺ #,##0"))'),
        ("I2", "NPV yüzdelik P90",
         "=IFERROR(SENARYO!B32,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"PERCENTILE P90 ",TEXT(C18,"₺ #,##0"))'),
    ]
    for i, (kod, ad, deger, yorum) in enumerate(moduller, 11):
        h(ws, i, 1, kod, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, ad, yazi="333333")
        h(ws, i, 3, deger, sayi=TL if kod not in ("O2", "O3", "O8", "T1") else (CATI if kod in ("O8", "T1") else None),
          yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.merge_cells(start_row=i, start_column=4, end_row=i, end_column=10)

    h(ws, 20, 1, "CAPEX/OPEX özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=IFERROR(gyf_capexOpex,\"-\")", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B20:H20")

    h(ws, 22, 1, "PanoDemo", kalin=True, yazi=GRİ)
    for i, (ad, v) in enumerate([("Gelir", 393_000), ("CAPEX", 1_200_000), ("NPV", 850_000)], 23):
        h(ws, i, 1, ad)
        h(ws, i, 2, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Pano Özet"
    g.add_data(Reference(ws, min_col=2, min_row=22, max_row=25), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=23, max_row=25))
    g.height, g.width = 7, 10
    ws.add_chart(g, "D22")
    genislik(ws, {"A": 28, "B": 28, "C": 18, "D": 50})


def karar(ws):
    sayfa_hazirla(ws, "KARAR", "B3261E", URUN_AD, son_kolon=30)
    baski_hazirla(ws, "A1:F28", f"{URUN_AD} · KARAR")
    h(ws, 3, 1, "Yatırım kararı", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Karar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 2, "=gyf_kararMetni", boyut=16, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=gyf_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B6:H6")
    h(ws, 8, 1, "NPV (baz)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 8, 2, "=IFERROR(gyf_npv,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 9, 1, "IRR", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 9, 2, "=IFERROR(gyf_irr,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "İyimser NPV", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 2, "=IFERROR(MOTOR!G48,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Kötümser karar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(MOTOR!G50,\"YATIRIM YAPMA\")", yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 12, 1, "Kötümser NPV", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=IFERROR(MOTOR!G49,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "Finansman", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 2, "=IFERROR(gyf_finansmanKarsilastirma,\"-\")", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B14:H14")
    h(ws, 16, 1, "Aksiyon 1", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 2,
      '=_xlfn.TEXTJOIN(" ",TRUE,"Tarife duyarlılığını kontrol edin; etki ",TEXT(IFERROR(MOTOR!G51,0),"₺ #,##0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    h(ws, 17, 1, "Aksiyon 2", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 2,
      '=_xlfn.TEXTJOIN(" ",TRUE,"Öz tüketim oranını gözden geçirin; eşleşme ",TEXT(IFERROR(gyf_tuketimEslesme,0),"#,##0")," kWh")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    h(ws, 18, 1, "Aksiyon 3", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 2,
      '=_xlfn.TEXTJOIN(" ",TRUE,"OZ_KAYNAK vs KREDI karşılaştırmasını FINANSMAN sayfasında inceleyin")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    h(ws, 20, 1, "KararDemo", kalin=True, yazi=GRİ)
    for i, (ad, v) in enumerate([("İyimser", 1_800_000), ("Baz", 850_000), ("Kötümser", -180_000)], 21):
        h(ws, i, 1, ad)
        h(ws, i, 2, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Karar NPV Bandı"
    g.add_data(Reference(ws, min_col=2, min_row=20, max_row=23), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=21, max_row=23))
    g.height, g.width = 7, 10
    ws.add_chart(g, "D20")
    genislik(ws, {"A": 22, "B": 70})


def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", "548235", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Altın vakalar — beklenen vs hesaplanan", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:H3")
    baslik_satiri(ws, 5, [*VAKA_BASLIK])
    # Vaka1: 100kWp × 1400 = 140000 kWh; gelir≈393400; CAPEX=1.2M
    satirlar = [
        ["V1", "Demo çatı baz CAPEX", "100kWp×12000", 1_200_000,
         "=IFERROR(gyf_capexToplam,0)", "=E6-D6",
         '=IF(ABS(F6)<=gyf_tolerans,"TUTARLI","KIRIK")', "Demo"],
        ["V2", "Demo yıllık üretim", "100×1400", 140_000,
         "=IFERROR(MOTOR!G26,0)", "=E7-D7",
         '=IF(ABS(F7)<=gyf_tolerans,"TUTARLI","KIRIK")', "Demo"],
        ["V3", "Öz tüketim kWh", "140000×0.7", 98_000,
         "=IFERROR(MOTOR!G27,0)", "=E8-D8",
         '=IF(ABS(F8)<=gyf_tolerans,"TUTARLI","KIRIK")', "Demo"],
        ["V4", "Kötümser karar yolu", "NPV<0 → YAPMA", 1,
         '=IF(ISNUMBER(SEARCH("YAPMA",IFERROR(MOTOR!G50,""))),1,0)', "=E9-D9",
         '=IF(E9=1,"TUTARLI","KIRIK")', "F05"],
    ]
    for i, s in enumerate(satirlar, 6):
        for k, v in enumerate(s, 1):
            sayi = TL if k in (4, 5, 6) and i == 6 else (CATI if k in (4, 5, 6) else None)
            if i > 6 and k in (4, 5, 6):
                sayi = CATI
            h(ws, i, k, v, sayi=sayi, yazi=DIS_REF_YEŞIL if k in (5, 6, 7) else "333333")

    h(ws, 12, 1, "Vaka durum özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,G6,G7,G8,G9),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "Vaka fark toplamı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2, "=IFERROR(SUM(F6:F9),0)", sayi=CATI, yazi=DIS_REF_YEŞIL)

    h(ws, 15, 1, "VakaDemo", kalin=True, yazi=GRİ)
    for i, v in enumerate([1_200_000, 140_000, 98_000, 1], 16):
        h(ws, i, 1, f"V{i-15}")
        h(ws, i, 2, v, sayi=CATI if i > 16 else TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Vaka Beklenen"
    g.add_data(Reference(ws, min_col=2, min_row=15, max_row=19), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=16, max_row=19))
    g.height, g.width = 7, 10
    ws.add_chart(g, "D15")
    genislik(ws, {"A": 14, "B": 28, "C": 22, "D": 14, "E": 14, "F": 12, "G": 12, "H": 10})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", "0F2742", URUN_AD, son_kolon=30)
    baski_hazirla(ws, "A1:F40", f"{URUN_AD} · KANIT")
    h(ws, 3, 1, "Bankaya / yatırımcıya fizibilite kanıtı", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Şirket", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Dönem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=GIRDI!B7", yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Proje tipi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=GIRDI!B9", yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Kurulu güç [kWp]", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 8, 2, "=GIRDI!B10", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "CAPEX / YATIRIM", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 2, "=IFERROR(gyf_capexToplam,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "OPEX / ISLETME", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(gyf_opexYillik,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Yıllık gelir", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=IFERROR(gyf_yillikGelir,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "Tüketim eşleşme", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2, "=IFERROR(gyf_tuketimEslesme,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "NPV baz", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 2, "=IFERROR(gyf_npv,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "IRR", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 2, "=IFERROR(gyf_irr,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "NPV aralık", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 2,
      '=_xlfn.TEXTJOIN(" … ",TRUE,TEXT(IFERROR(MOTOR!G48,0),"₺ #,##0"),TEXT(IFERROR(MOTOR!G49,0),"₺ #,##0"))',
      yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Finansman", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2, "=IFERROR(gyf_finansmanKarsilastirma,\"-\")", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B19:H19")
    h(ws, 20, 1, "Karar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=gyf_kararMetni", kalin=True, boyut=14, yazi=DIS_REF_YEŞIL)
    h(ws, 21, 1, "Kanıt özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 2, "=IFERROR(gyf_kanitRaporu,\"-\")", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B21:H21")
    h(ws, 23, 1, "İmza — Yatırımcı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "________________________", yazi=GRİ)
    h(ws, 24, 1, "İmza — Banka", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, "________________________", yazi=GRİ)
    h(ws, 26, 1, f"Sürüm {SURUM} | GYF-PRO | Tarih", yazi=GRİ)
    h(ws, 26, 2, "=raporTarihi", sayi=TARİH, yazi=GRİ)
    genislik(ws, {"A": 28, "B": 60})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "833C0C", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Dosya sağlık kontrolleri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    kontroller_list = [
        (5, "Giriş dolu mu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"BOŞ","DOLU")'),
        (6, "CAPEX pozitif mi",
         '=IF(IFERROR(gyf_capexToplam,0)>0,"TEMİZ","NEGATİF")'),
        (7, "OPEX ayrık mı",
         '=IF(AND(IFERROR(gyf_capexToplam,0)>0,IFERROR(gyf_opexYillik,0)>=0),"TEMİZ","EKSİK")'),
        (8, "NPV hesaplı mı",
         '=IF(ISNUMBER(gyf_npv),"TEMİZ","HATALI")'),
        (9, "Kötümser YAPMA yolu",
         '=IF(ISNUMBER(SEARCH("YAPMA",IFERROR(MOTOR!G50,""))),"TEMİZ","EKSİK")'),
        (10, "Varsayım kaynak/güven",
         '=IF(COUNTIF(tblVarsayimlar[kontrol],"EKSİK")=0,"TEMİZ","EKSİK")'),
        (11, "Vaka tutarlılığı",
         '=IF(COUNTIF(VAKALAR!G6:G9,"KIRIK")=0,"TUTARLI","KIRIK")'),
    ]
    for r, ad, formul in kontroller_list:
        h(ws, r, 1, ad, yazi="333333")
        h(ws, r, 2, formul, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Sağlık skoru", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(ROUND((COUNTIF(B5:B11,"TEMİZ")+COUNTIF(B5:B11,"DOLU")+COUNTIF(B5:B11,"TUTARLI"))'
      '/7*100,0),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    genislik(ws, {"A": 32, "B": 20})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek senaryolar (CSV ile uyumlu)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = ["Senaryo", "Proje_tipi", "Kurulu_guc_kWp", "Spesifik_uretim", "Oz_tuketim",
                "Perakende_tarife", "Satis_tarife", "Capex_birim", "Opex_oran", "WACC", "Yil", "Durum"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Durum": (
            '=IF(tblOrnek[[#This Row],[Senaryo]]="","",'
            'IF(tblOrnek[[#This Row],[Kurulu_guc_kWp]]="","EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:L1004", sutunlar, formuller)
    rows = [
        ("Demo cati", "CATI", 100, 1400, 0.70, 3.50, 1.20, 12_000, 0.015, 0.18, 2026),
        ("Yuksek uretim", "CATI", 100, 1600, 0.75, 3.80, 1.40, 11_000, 0.012, 0.16, 2026),
        ("Kotumser saha", "SAHA", 250, 1200, 0.40, 3.20, 1.00, 14_000, 0.020, 0.20, 2026),
        ("2025 gecis", "CATI", 80, 1350, 0.65, 3.00, 1.10, 12_500, 0.016, 0.18, 2025),
    ]
    for i, row in enumerate(rows, 6):
        for k, v in enumerate(row, 1):
            sayi = None
            if k in (3, 4, 11):
                sayi = CATI
            if k in (5, 9, 10):
                sayi = YÜZDE
            if k in (6, 7):
                sayi = "0.00"
            if k == 8:
                sayi = TL
            h(ws, i, k, v, sayi=sayi, zemin=GIRIS_SARI if k <= 11 else None, yazi="333333")
            if k <= 11:
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", sutunlar[k - 1], "Örnek satır")
    h(ws, 22, 1, "OrnekCapexDemo", kalin=True, yazi=GRİ)
    for i, v in enumerate([1_200_000, 1_100_000, 3_500_000, 1_000_000], 23):
        h(ws, i, 1, rows[i - 23][0])
        h(ws, i, 2, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Örnek CAPEX"
    g.add_data(Reference(ws, min_col=2, min_row=22, max_row=26), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=23, max_row=26))
    g.height, g.width = 7, 12
    ws.add_chart(g, "D22")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 13)})
    genislik(ws, {"A": 22, "L": 12})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", "7030A0", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Doğrulama listeleri", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Yıl", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, y in enumerate([2025, 2026, 2027], 6):
        h(ws, i, 1, y, sayi=CATI)
    h(ws, 5, 3, "ProjeTipi", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "CATI")
    h(ws, 7, 3, "SAHA")
    h(ws, 5, 5, "EvetHayır", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 5, "Evet")
    h(ws, 7, 5, "Hayır")
    genislik(ws, {"A": 24, "C": 12, "E": 12})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "833C0C", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Parametreler (tek kaynak) — tarife/teşvik", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = ["anahtar", "deger", "birim", "aciklama", "kaynak", "yururluk_tarihi", "dogrulama_tarihi", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    params = [
        ("hesapYili", DEMO["hesapYili"], "yıl", "Aktif hesap yılı", "Kullanıcı / GIRDI", "01.01.2025", "11.08.2026"),
        ("raporTarihi", RAPOR_TARIH, "tarih", "Rapor tarihi", "Üretim", "01.01.2025", "11.08.2026"),
        ("gyf_tolerans", 1, "adet", "Vaka tutarlılık toleransı", "Uygulama notu", "01.01.2025", "11.08.2026"),
        ("gyf_azamiEsik", 100_000_000_000, "TL", "Azami giriş eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("gyf_senaryoIyi", 1.12, "çarpan", "İyimser üretim çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("gyf_senaryoKotu", 0.75, "çarpan", "Kötümser üretim çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("gyf_tarifeKotu", 0.85, "çarpan", "Kötümser tarife çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("gyf_capexIyi", 0.95, "çarpan", "İyimser CAPEX çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("gyf_capexKotu", 1.15, "çarpan", "Kötümser CAPEX çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("gyf_opexKotu", 1.40, "çarpan", "Kötümser OPEX çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("gyf_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("gyf_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("gyf_tornadoTarife", 1.10, "çarpan", "Tornado tarife şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("gyf_tornadoUretim", 1.12, "çarpan", "Tornado üretim şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("gyf_tornadoCapex", 1.10, "çarpan", "Tornado CAPEX şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("gyf_tornadoOpex", 1.20, "çarpan", "Tornado OPEX şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("gyf_tornadoWacc", 0.90, "çarpan", "Tornado WACC şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("gyf_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "11.08.2026"),
        ("gyf_parmakIzi", "GYF-PRO-1.0.0", "metin", "Dosya parmak izi", "Üretim", "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv", "01.01.2025", "11.08.2026"),
        ("gyf_girisBeklenen", 17, "adet", "Zorunlu giriş sayısı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("gyf_riskDusuk", 20, "puan", "YAPILIR risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("gyf_riskOrta", 55, "puan", "DİKKAT risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("gyf_riskYuksek", 85, "puan", "YAPMA risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("gyf_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("gyf_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("gyf_bazCarpan", 1, "çarpan", "Baz senaryo çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("gyf_sifirCarpan", 0, "çarpan", "Nötr çarpan", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("gyf_siraBir", 1, "adet", "Tornado sıra 1", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("gyf_siraIki", 2, "adet", "Tornado sıra 2", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("gyf_siraUc", 3, "adet", "Tornado sıra 3", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("gyf_dikkatCarpan", 1.25, "çarpan", "Geri ödeme dikkat çarpanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("gyf_npvDikkat", 100_000, "TL", "NPV dikkat eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("gyf_geriOdemeSonsuz", 99, "yıl", "Geri ödeme tanımsız işaret", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("vaka_capex", 1_200_000, "TL", "Vaka CAPEX", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_uretim", 140_000, "kWh", "Vaka üretim", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_oz", 98_000, "kWh", "Vaka öz tüketim", "Altın vaka", "01.01.2025", "11.08.2026"),
    ]
    for i, row in enumerate(params, 6):
        for k, v in enumerate(row, 1):
            sayi = None
            if k == 2:
                if row[2] == "tarih":
                    sayi = TARİH
                elif row[2] in ("oran", "çarpan"):
                    sayi = YÜZDE
                elif row[2] == "TL":
                    sayi = TL
                elif row[2] in ("yıl", "adet", "puan", "kWh"):
                    sayi = CATI
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k == 2 and sayi is None and isinstance(v, (int, float)) and not isinstance(v, bool):
                ws.cell(i, k).number_format = CATI
            if k == 2:
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"B{i}", row[0], row[3])

    for i in range(6, 6 + len(params)):
        h(ws, i, 8, f'=IF(A{i}="","",IF(B{i}="","EKSİK","TAM"))', yazi=DIS_REF_YEŞIL, hiza="center")

    son = 5 + len(params)
    # A kolonuna yazma — D10 anahtar tablosu sanır; dipnot J kolonunda
    h(ws, son + 2, 10, f"Sürüm {SURUM} | Şifre koruması: {SIFRE} (formül alanları)", yazi=GRİ, boyut=9)
    dogrulama(ws, "list", "ListeYillar", "B6",
              baslik="Hesap Yılı", mesaj="Aktif yılı seçin.",
              hata_baslik="Geçersiz", hata_mesaj="2025-2027.")
    dogrulama(ws, "decimal", "0", "B10",
              baslik="İyimser çarpan", mesaj="0-2.",
              hata_baslik="Geçersiz", hata_mesaj="0-2.",
              isaret="between", f2="2")
    dogrulama(ws, "decimal", "0", "B11",
              baslik="Kötümser çarpan", mesaj="0-2.",
              hata_baslik="Geçersiz", hata_mesaj="0-2.",
              isaret="between", f2="2")
    genislik(ws, {"A": 28, "B": 24, "C": 10, "D": 32, "E": 22, "F": 14, "G": 14, "H": 12, "J": 48})
    ws.sheet_view.zoomScale = 90
    return len(params)


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar kuralları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "VERİ YOK — giriş tablosu boşsa hesap yapılmaz.",
        "YATIRIM YAPILIR — NPV pozitif ve geri ödeme hedef altında.",
        "DİKKAT — geri ödeme veya NPV dikkat bandında; tarife/üretim gözden geçirilmeli.",
        "YATIRIM YAPMA — NPV negatif veya geri ödeme aşırı (kötümser yol dahil).",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 10, 1, "Belirsizlik beyanları", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari(["yekdem_tarife_yenileme"]), 11):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)
    h(ws, 13, 1,
      "YEKDEM/şebeke satış tarifesi dönemsel yenilenir; bu dosya kullanıcı girişli tarife kullanır "
      "(dış API yok — Ç15). Tarife değişince GIRDI ve KURALLAR güncellenir.",
      yazi=KRITIK, kaydir=True)
    ws.merge_cells("A13:J13")

    h(ws, 15, 1, "Dış veri formatı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 2,
      "Özgül üretim: PVGIS veya saha ölçümü (kWh/kWp). Tarife: EPDK tebliğ / fatura. "
      "CAPEX: EPC teklif. OPEX: işletme bütçesi. API bağlantısı yoktur.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B15:J15")

    h(ws, 17, 1, "Kullanım sırası", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 1,
      "GIRDI → VARSAYIMLAR → KURALLAR → MOTOR → FINANSMAN → SENARYO → PANO → KARAR → "
      "VAKALAR → KANIT_RAPORU → AYARLAR.",
      kaydir=True, yazi="333333")
    ws.merge_cells("A18:J18")
    h(ws, 20, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv | GYF-PRO", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 36, "B": 70})


def kosullu_bicimlendirme(wb):
    yesil = Font(color=NORMAL, bold=True)
    kirmizi = Font(color=KRITIK, bold=True)
    amber = Font(color="B7791F", bold=True)
    y_fill = PatternFill("solid", fgColor="E2EFDA")
    k_fill = PatternFill("solid", fgColor="FDE9E9")
    a_fill = PatternFill("solid", fgColor="FFF2CC")

    ws = wb["KARAR"]
    ws.conditional_formatting.add("B5", FormulaRule(formula=['ISNUMBER(SEARCH("VERİ YOK",B5))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B5", FormulaRule(formula=['ISNUMBER(SEARCH("YATIRIM YAPMA",B5))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B5", FormulaRule(formula=['ISNUMBER(SEARCH("YATIRIM YAPILIR",B5))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B5", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B5))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B8", CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=yesil))
    ws.conditional_formatting.add("B8", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("B11", FormulaRule(formula=['ISNUMBER(SEARCH("YAPMA",B11))'], font=kirmizi))
    ws.conditional_formatting.add("B12", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("B10", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))

    ws = wb["PANO"]
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("VERİ YOK",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("YAPMA",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("YAPILIR",B6))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B6))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B4", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("G4", CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=yesil))
    ws.conditional_formatting.add("G4", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("J4", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("J4", CellIsRule(operator="between", formula=["50", "79"], font=amber))
    ws.conditional_formatting.add("J4", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))
    ws.conditional_formatting.add("K4", CellIsRule(operator="greaterThanOrEqual", formula=["70"], font=kirmizi))
    ws.conditional_formatting.add("K4", CellIsRule(operator="between", formula=["40", "69"], font=amber))
    ws.conditional_formatting.add("K4", CellIsRule(operator="lessThan", formula=["40"], font=yesil))
    ws.conditional_formatting.add("M4", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))

    ws = wb["GIRDI"]
    for col in ("B",):
        ws.conditional_formatting.add(f"{col}10:{col}17", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(f"{col}10:{col}11", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("F31:F47", FormulaRule(formula=['F31="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F31:F47", FormulaRule(formula=['F31="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F6:F9", CellIsRule(operator="notEqual", formula=["0"], font=amber))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("D6:D8", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("D6:D8", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("E6:E8", FormulaRule(formula=['ISNUMBER(SEARCH("YAPMA",E6))'], font=kirmizi))
    ws.conditional_formatting.add("E6:E8", FormulaRule(formula=['ISNUMBER(SEARCH("YAPILIR",E6))'], font=yesil))
    ws.conditional_formatting.add("E6:E8", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",E6))'], font=amber))
    ws.conditional_formatting.add("B15:B19", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add("B5:B11", FormulaRule(
        formula=['OR(B5="KIRIK",B5="HATALI",B5="NEGATİF",B5="BOŞ",B5="EKSİK",B5="AŞIYOR")'],
        font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B5:B11", FormulaRule(
        formula=['OR(B5="TEMİZ",B5="DOLU",B5="NORMAL",B5="TUTARLI")'],
        font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B12", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("B12", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("G7:G70", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("F7:F70", FormulaRule(formula=['F7="GYF-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F70", FormulaRule(formula=['F7="GYF-003"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["ORNEK_VERI"]
    for harf in ("C", "D", "H"):
        ws.conditional_formatting.add(f"{harf}6:{harf}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B20", FormulaRule(formula=['ISNUMBER(SEARCH("YAPMA",B20))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B20", FormulaRule(formula=['ISNUMBER(SEARCH("YAPILIR",B20))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B20", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B20))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B15", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("B16", CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=yesil))

    ws = wb["FINANSMAN"]
    ws.conditional_formatting.add("E6:E7", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("E6:E7", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))

    ws = wb["VARSAYIMLAR"]
    ws.conditional_formatting.add("H6:H15", FormulaRule(formula=['H6="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("H6:H15", FormulaRule(formula=['H6="TAM"'], font=yesil, fill=y_fill))


def ad_tanimlari(wb):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")

    # MOTOR satır numaraları: adım i → satır 6+i → G kolonu
    # 33 CAPEX, 34 OPEX, 32 gelir, 37 NPV, 38 IRR, 40 karar, 41 gerekçe
    # 29 eşleşme, 57 kanıt, 58-63 panolar
    # MOTOR: adım n → satır 6+n; pano aynaları 53–58 → G59–G64; kanıt=52→G58
    ad_ekle(wb, "gyf_yillikGelir", "MOTOR!$G$59")
    ad_ekle(wb, "gyf_capexToplam", "MOTOR!$G$60")
    ad_ekle(wb, "gyf_opexYillik", "MOTOR!$G$61")
    ad_ekle(wb, "gyf_npv", "MOTOR!$G$62")
    ad_ekle(wb, "gyf_irr", "MOTOR!$G$63")
    ad_ekle(wb, "gyf_tuketimEslesme", "MOTOR!$G$64")
    ad_ekle(wb, "gyf_kararMetni", "MOTOR!$G$40")
    ad_ekle(wb, "gyf_kararGerekce", "MOTOR!$G$41")
    ad_ekle(wb, "gyf_kanitRaporu", "MOTOR!$G$58")
    ad_ekle(wb, "gyf_finansmanKarsilastirma", "FINANSMAN!$B$11")

    ad_ekle(wb, "gyf_senaryoKarsilastirma", "SENARYO!$B$10")
    ad_ekle(wb, "gyf_yorumSenaryo", "SENARYO!$B$11")
    ad_ekle(wb, "gyf_duyarlilikSira", "SENARYO!$B$21")
    ad_ekle(wb, "gyf_yorumDuyarlilik", "SENARYO!$B$22")
    ad_ekle(wb, "gyf_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "gyf_yorumTahmin", "SENARYO!$B$33")
    ad_ekle(wb, "gyf_vakaDurum", "VAKALAR!$B$12")

    ayar_map = {
        "gyf_tolerans": 8,
        "gyf_azamiEsik": 9,
        "gyf_senaryoIyi": 10,
        "gyf_senaryoKotu": 11,
        "gyf_tarifeKotu": 12,
        "gyf_capexIyi": 13,
        "gyf_capexKotu": 14,
        "gyf_opexKotu": 15,
        "gyf_tahminAlt": 16,
        "gyf_tahminUst": 17,
        "gyf_tornadoTarife": 18,
        "gyf_tornadoUretim": 19,
        "gyf_tornadoCapex": 20,
        "gyf_tornadoOpex": 21,
        "gyf_tornadoWacc": 22,
        "gyf_yuzdelikOran": 23,
        "gyf_parmakIzi": 24,
        "gyf_girisBeklenen": 26,
        "gyf_riskDusuk": 27,
        "gyf_riskOrta": 28,
        "gyf_riskYuksek": 29,
        "gyf_yilTaban": 30,
        "gyf_yuvarMax": 31,
        "gyf_bazCarpan": 32,
        "gyf_sifirCarpan": 33,
        "gyf_siraBir": 34,
        "gyf_siraIki": 35,
        "gyf_siraUc": 36,
        "gyf_dikkatCarpan": 37,
        "gyf_npvDikkat": 38,
        "gyf_geriOdemeSonsuz": 39,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "gyf_modulT1": (11, "C"), "gyf_modulT1Yorum": (11, "D"),
        "gyf_modulT2": (12, "C"), "gyf_modulT2Yorum": (12, "D"),
        "gyf_modulO1": (13, "C"), "gyf_modulO1Yorum": (13, "D"),
        "gyf_modulO8": (16, "C"), "gyf_modulO8Yorum": (16, "D"),
        "gyf_modulI2": (18, "C"), "gyf_modulI2Yorum": (18, "D"),
    }
    for ad, (satir, kol) in modul_map.items():
        ad_ekle(wb, ad, f"PANO!${kol}${satir}")

    ad_ekle(wb, "ListeYillar", "LISTELER!$A$6:$A$8")
    ad_ekle(wb, "ListeProjeTipi", "LISTELER!$C$6:$C$7")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$6:$E$7")


def main(cikti_yolu=None):
    wb = Workbook()
    wb.remove(wb.active)

    siralar = [
        ("KAPAK", kapak),
        ("HIZLI_BASLANGIC", hizli_baslangic),
        ("GIRDI", girdi),
        ("VARSAYIMLAR", varsayimlar),
        ("KURALLAR", kurallar),
        ("MOTOR", motor),
        ("FINANSMAN", finansman),
        ("SENARYO", senaryo),
        ("PANO", pano),
        ("KARAR", karar),
        ("VAKALAR", vakalar),
        ("KANIT_RAPORU", kanit_raporu),
        ("KONTROLLER", kontroller),
        ("ORNEK_VERI", ornek_veri),
        ("LISTELER", listeler),
    ]
    for ad, fn in siralar:
        ws = wb.create_sheet(ad)
        fn(ws)

    ws_ay = wb.create_sheet("AYARLAR")
    ayarlar(ws_ay)
    ws_kil = wb.create_sheet("KILAVUZ")
    kilavuz(ws_kil)

    # CAPEX/OPEX özet metni — F03 ve serbest alternatif karşılığı
    ws_f = wb["FINANSMAN"]
    h(ws_f, 12, 1, "CAPEX OPEX özet", kalin=True, yazi=KOYU_LACIVERT)
    h(ws_f, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"CAPEX ",TEXT(B9,"₺ #,##0")," · OPEX ",TEXT(B10,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    ad_tanimlari(wb)
    # gyf_capexOpex → FINANSMAN B12
    # ad_tanimlari set B9 earlier — replace
    if "gyf_capexOpex" in wb.defined_names:
        del wb.defined_names["gyf_capexOpex"]
    ad_ekle(wb, "gyf_capexOpex", "FINANSMAN!$B$12")

    kosullu_bicimlendirme(wb)
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    wb.calculation.fullCalcOnLoad = True

    kok = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    varsayilan = os.path.join(kok, "GesYatirimFizibilite.xlsx")
    hedef = cikti_yolu or varsayilan
    os.makedirs(os.path.dirname(os.path.abspath(hedef)) or ".", exist_ok=True)
    wb.save(hedef)

    hsh = hashlib.sha256()
    with open(hedef, "rb") as f:
        for b in iter(lambda: f.read(65536), b""):
            hsh.update(b)
    dig = hsh.hexdigest()

    wb2 = __import__("openpyxl").load_workbook(hedef)
    wb2["AYARLAR"]["B24"] = dig[:16]
    for wsx in wb2.worksheets:
        sayfa_koru(wsx)
    wb2.save(hedef)

    print(f"Dosya: {hedef}")
    print(f"SHA-256: {dig}")
    print(f"Şifre: {SIFRE}")
    return hedef


if __name__ == "__main__":
    yol = sys.argv[1] if len(sys.argv) > 1 else None
    uretilen = main(yol)
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
    cikti = os.path.join(repo, "cikti", "GesYatirimFizibilite.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
