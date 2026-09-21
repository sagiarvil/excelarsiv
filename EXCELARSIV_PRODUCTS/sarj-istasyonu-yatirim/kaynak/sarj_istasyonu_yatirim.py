#!/usr/bin/env python3
"""
Şarj İstasyonu Yatırım Simülatörü — üretim betiği (A6 / manda v6).
DC hızlı şarj CAPEX/OPEX, kullanım×tarife, teşvik net CAPEX, NPV/IRR, tornado.
Karar: YATIRIM YAPILIR / DİKKAT / YATIRIM YAPMA / VERİ YOK.
Kod: SIY-PRO · önek siy_* — GES kopyası değildir.
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
from ortak.mevzuat_motoru import KURAL_BASLIK, MOTOR_BASLIK, VAKA_BASLIK, belirsizlik_beyanlari

URUN_AD = "Şarj İstasyonu Yatırım Simülatörü"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 11)

# Demo: 4×150 kW DC — baz NPV pozitif, kötümser negatif; teşvik %15
DEMO = {
    "sirket_adi": "Örnek Şarj Yatırım A.Ş.",
    "donem": "2026-Q3",
    "hesapYili": 2026,
    "soket_adet": 4,
    "guc_kw": 150,
    "gunluk_seans": 28,
    "kwh_seans": 35,
    "kullanim_oran": 0.85,
    "perakende_tl_kwh": 12.00,
    "toptan_tl_kwh": 4.50,
    "ag_ucreti_tl_kwh": 0.80,
    "capex_birim": 1_200_000,
    "tesvik_oran": 0.15,
    "opex_bakim": 180_000,
    "opex_kira": 240_000,
    "wacc": 0.18,
    "npv_donem": 15,
    "kredi_orani": 0.32,
    "oz_kaynak_oran": 0.40,
}

_YV = "MAX(0,MIN(siy_yuvarMax,IFERROR(N(G12),0)))"


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: Şarj istasyonu NPV/IRR kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: CAPEX/OPEX ve karar bozulur | Kaynak: tarife / teşvik / kullanım",
    )


def kural_cek(kural_id: str) -> str:
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-siy_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


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
      "Soket, kullanım, tarife ve CAPEX/OPEX girdilerini işleyin; teşvik net CAPEX ile "
      "NPV/IRR aralığı ve YATIRIM YAPILIR / YATIRIM YAPMA kararını üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Soket×seans×kWh kullanım gelir modeli (DC hızlı şarj)",
        "Perakende / toptan / ağ tarifesi marjı + teşvik net CAPEX",
        "CAPEX/OPEX ayrımı + öz kaynak / kredi finansman senaryosu",
        "NPV / IRR + tornado duyarlılık; kötümser YATIRIM YAPMA yolu",
        "Altın vakalar ve KANIT_RAPORU (A4 imza)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "EV şarj yatırımcısı — teşvik açılışında fizibiliteyi yeniden hesaplamak",
        "CFO — NPV/IRR ile istasyon yatırımı kararı vermek",
        "Banka / teşvik komitesi — imzalı kanıt raporunu dosyalamak",
    ], 14):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) HIZLI_BASLANGIC → 2) GIRDI sarı hücreler → 3) PANO/KARAR → 4) KANIT_RAPORU yazdır.",
      yazi="333333", kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 21, 1, f"Sürüm {SURUM} | 2026 | ExcelArşiv | Lisans: Tek kullanıcı | Kod: SIY-PRO",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 72})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Üç adımda sonuç", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    for i, (baslik, metin) in enumerate([
        ("Adım 1 — Girdiler", "GIRDI sayfasında soket, kullanım, tarife ve CAPEX/OPEX sarı hücrelere girin."),
        ("Adım 2 — Karar", "PANO ve KARAR'da net CAPEX, NPV/IRR ve YATIRIM YAPILIR / YATIRIM YAPMA görün."),
        ("Adım 3 — Kanıt", "KANIT_RAPORU'nu yazdırın; VAKALAR ve SENARYO ile doğrulayın."),
    ], 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, metin, kaydir=True, yazi="333333")
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=12)
    h(ws, 9, 1, "Formüller", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 1,
      "Yıllık kWh = seans×kWh×365×kullanım | Gelir = kWh×perakende | "
      "Net CAPEX = soket×birim×(1−teşvik) | NPV = net nakit anüitesi − net CAPEX",
      yazi="333333", kaydir=True)
    ws.merge_cells("A10:L10")
    h(ws, 12, 1, "Sarı = manuel giriş | Yeşil = formül çıktısı | Şifre 1234 formül alanlarını korur.",
      yazi=GRİ, kaydir=True)
    genislik(ws, {"A": 28, "B": 70})


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Şarj istasyonu yatırım girdileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Tarife/teşvik KURALLAR'dan ayna edilir.",
      yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Şirket / Proje", DEMO["sirket_adi"], None, "sirket_adi", "Unvan veya proje kodu."),
        (7, "Dönem", DEMO["donem"], None, "donem", "Rapor dönemi."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili", "2025, 2026 veya 2027."),
        (9, "Soket adedi", DEMO["soket_adet"], CATI, "soket_adet", "DC hızlı şarj soket sayısı."),
        (10, "Güç [kW]", DEMO["guc_kw"], CATI, "guc_kw", "Soket başına güç."),
        (11, "Günlük seans", DEMO["gunluk_seans"], CATI, "gunluk_seans", "Ortalama günlük seans."),
        (12, "Seans kWh", DEMO["kwh_seans"], CATI, "kwh_seans", "Seans başına enerji."),
        (13, "Kullanım oranı", DEMO["kullanim_oran"], YÜZDE, "kullanim_oran", "Doluluk / kullanım çarpanı."),
        (14, "Perakende [₺/kWh]", DEMO["perakende_tl_kwh"], TL, "perakende", "EV kullanıcı tarifesi."),
        (15, "Toptan [₺/kWh]", DEMO["toptan_tl_kwh"], TL, "toptan", "Toptan elektrik maliyeti."),
        (16, "Ağ ücreti [₺/kWh]", DEMO["ag_ucreti_tl_kwh"], TL, "ag_ucreti", "Ağ kullanım ücreti."),
        (17, "CAPEX birim [₺]", DEMO["capex_birim"], TL, "capex_birim", "Soket başına yatırım."),
        (18, "Teşvik oranı", DEMO["tesvik_oran"], YÜZDE, "tesvik_oran", "CAPEX teşvik oranı."),
        (19, "OPEX bakım [₺/yıl]", DEMO["opex_bakim"], TL, "opex_bakim", "Yıllık bakım."),
        (20, "OPEX kira [₺/yıl]", DEMO["opex_kira"], TL, "opex_kira", "Yıllık kira."),
        (21, "WACC [%]", DEMO["wacc"], YÜZDE, "wacc", "İskonto oranı."),
        (22, "NPV dönem [yıl]", DEMO["npv_donem"], CATI, "npv_donem", "Nakit akışı yılı."),
        (23, "Kredi faiz [%]", DEMO["kredi_orani"], YÜZDE, "kredi_orani", "Kredi senaryosu faiz."),
        (24, "Öz kaynak oranı", DEMO["oz_kaynak_oran"], YÜZDE, "oz_kaynak_oran", "Öz kaynak payı."),
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
        elif _ad in ("soket_adet", "gunluk_seans", "kwh_seans", "guc_kw"):
            birim = "adet/kW"
        else:
            birim = "metin"
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim, yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 26, 1, "Teşvik oranı (kural ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 2, kural_cek("SIY-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 27, 1, "Net CAPEX (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 2, "=IFERROR(siy_netCapex,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 28, 1, "Giriş doluluk", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 28, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B24)=0,siy_girisBeklenen=0),0,ROUND(COUNTA(B6:B24)/siy_girisBeklenen*100,0)),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)

    h(ws, 30, 1, "Giriş tablosu (otomasyon / boş-mod)", kalin=True, yazi=KOYU_LACIVERT)
    sutunlar = ["Alan Anahtarı", "Değer", "Birim", "Zorunlu", "Kaynak", "Kontrol"]
    baslik_satiri(ws, 31, sutunlar)
    formuller = {
        "Kontrol": (
            '=IF(tblGirdi[[#This Row],[Alan Anahtarı]]="","",'
            'IF(tblGirdi[[#This Row],[Değer]]="","EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblGirdi", "A31:F1030", sutunlar, formuller)
    ornek_satir = [
        ("sirket_adi", DEMO["sirket_adi"], "metin", "Evet", "GIRDI"),
        ("donem", DEMO["donem"], "metin", "Evet", "GIRDI"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Evet", "GIRDI"),
        ("soket_adet", DEMO["soket_adet"], "adet", "Evet", "GIRDI"),
        ("guc_kw", DEMO["guc_kw"], "kW", "Evet", "GIRDI"),
        ("gunluk_seans", DEMO["gunluk_seans"], "adet", "Evet", "GIRDI"),
        ("kwh_seans", DEMO["kwh_seans"], "kWh", "Evet", "GIRDI"),
        ("kullanim_oran", DEMO["kullanim_oran"], "oran", "Evet", "GIRDI"),
        ("perakende", DEMO["perakende_tl_kwh"], "TL", "Evet", "GIRDI"),
        ("toptan", DEMO["toptan_tl_kwh"], "TL", "Evet", "GIRDI"),
        ("ag_ucreti", DEMO["ag_ucreti_tl_kwh"], "TL", "Evet", "GIRDI"),
        ("capex_birim", DEMO["capex_birim"], "TL", "Evet", "GIRDI"),
        ("tesvik_oran", DEMO["tesvik_oran"], "oran", "Evet", "GIRDI"),
        ("opex_bakim", DEMO["opex_bakim"], "TL", "Evet", "GIRDI"),
        ("opex_kira", DEMO["opex_kira"], "TL", "Evet", "GIRDI"),
        ("wacc", DEMO["wacc"], "oran", "Evet", "GIRDI"),
        ("npv_donem", DEMO["npv_donem"], "yıl", "Evet", "GIRDI"),
        ("kredi_orani", DEMO["kredi_orani"], "oran", "Evet", "GIRDI"),
        ("oz_kaynak_oran", DEMO["oz_kaynak_oran"], "oran", "Evet", "GIRDI"),
    ]
    tl_a = {"perakende", "toptan", "ag_ucreti", "capex_birim", "opex_bakim", "opex_kira"}
    oran_a = {"kullanim_oran", "tesvik_oran", "wacc", "kredi_orani", "oz_kaynak_oran"}
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 32):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in tl_a:
            ws.cell(i, 2).number_format = TL
        elif a in oran_a:
            ws.cell(i, 2).number_format = YÜZDE
        elif a in ("hesapYili", "soket_adet", "guc_kw", "gunluk_seans", "kwh_seans", "npv_donem"):
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
    dogrulama(ws, "list", "ListeYillar", "B34",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    for aralik, baslik, mesaj in [
        ("B9", "Soket", "1 ile 100 arasında."),
        ("B10", "Güç", "1 ile 1000 kW."),
        ("B11", "Seans", "0 ile 1000."),
        ("B12", "kWh", "0 ile 500."),
        ("B14", "Perakende", "0 ile 1e6 TL/kWh."),
        ("B15", "Toptan", "0 ile 1e6."),
        ("B16", "Ağ", "0 ile 1e6."),
        ("B17", "CAPEX birim", "0 ile 1e12."),
        ("B19", "Bakım", "0 ile 1e12."),
        ("B20", "Kira", "0 ile 1e12."),
        ("B35", "Soket", "Tablo: adet."),
        ("B36", "Güç", "Tablo: kW."),
        ("B37", "Seans", "Tablo: adet."),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj=mesaj,
                  hata_baslik="Geçersiz tutar", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000000")
    for aralik, baslik in [("B13", "Kullanım"), ("B18", "Teşvik"), ("B21", "WACC"),
                           ("B23", "Kredi"), ("B24", "Öz kaynak"),
                           ("B39", "Kullanım"), ("B44", "Teşvik"), ("B47", "WACC"),
                           ("B49", "Kredi"), ("B50", "Öz kaynak")]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj="0 ile 1 arasında oran.",
                  hata_baslik="Geçersiz", hata_mesaj="0-1 arası.",
                  isaret="between", f2="1")
    for aralik, baslik in [("B22", "NPV dönem"), ("B48", "NPV dönem")]:
        dogrulama(ws, "whole", "1", aralik,
                  baslik=baslik, mesaj="1-25 arası.",
                  hata_baslik="Geçersiz", hata_mesaj="1-25.",
                  isaret="between", f2="25")

    giris_hucreleri(ws, 6, 24, [2])
    giris_hucreleri(ws, 32, 50, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 40, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1032, "Sarı alanlar giriş; tarife formülleri kilitlidir.")


def varsayimlar(ws):
    sayfa_hazirla(ws, "VARSAYIMLAR", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Model varsayımları — her satırda kaynak ve güven", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = ["varsayim_id", "aciklama", "deger", "birim", "kaynak", "guven", "duyarlilik", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    satirlar = [
        ("V-01", "Seans başına kWh", str(DEMO["kwh_seans"]), "kWh", "İşletme planı", "yumuşak", "yüksek"),
        ("V-02", "Günlük seans", str(DEMO["gunluk_seans"]), "adet", "Trafik tahmini", "yumuşak", "yüksek"),
        ("V-03", "Kullanım oranı", "0.85", "oran", "Saha gözlemi", "yumuşak", "yüksek"),
        ("V-04", "Perakende tarife", "12", "TL/kWh", "EPDK / perakende", "sert", "orta"),
        ("V-05", "Toptan enerji", "4.50", "TL/kWh", "Tedarik sözleşmesi", "yumuşak", "yüksek"),
        ("V-06", "Teşvik oranı", "0.15", "oran", "Teşvik programı — yenilenebilir", "yumuşak", "yüksek"),
        ("V-07", "WACC", "0.18", "oran", "Finans komitesi", "yumuşak", "yüksek"),
        ("V-08", "Kötümser kullanım şoku", "-25%", "oran", "Stres testi", "yumuşak", "yüksek"),
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
    tablo_ekle(ws, "tblVarsayimlar", "A5:H13", sutunlar, formuller)
    h(ws, 15, 10, "CAPEX / YATIRIM net (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 11, "=IFERROR(siy_netCapex,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 10, "OPEX / ISLETME yıllık (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 11, "=IFERROR(siy_opexYillik,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 12, "B": 28, "C": 18, "D": 10, "E": 28, "F": 12, "G": 12, "H": 10, "J": 32, "K": 18})


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Tarife / teşvik kural tablosu (yıl yan yana)", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["SIY-001", "Teşvik programı", "tesvik_orani", "CAPEX × (1−oran)", "tesvik_orani",
         0.10, 0.15, 0.15, "01.01.2025", "Şarj teşvik oranı"],
        ["SIY-002", "Ağ ücreti tavan", "ag_tavan", "₺/kWh tavan", "ag_tavan",
         1.00, 1.20, 1.20, "01.01.2025", "Ağ ücret tavanı"],
        ["SIY-003", "Bakım oranı", "bakim_oran", "CAPEX × oran", "bakim_oran",
         0.03, 0.035, 0.04, "01.01.2025", "Yıllık bakım oranı"],
        ["SIY-004", "Hedef marj", "hedef_marj", "₺/kWh ≥ eşik", "hedef_marj",
         4.00, 4.50, 5.00, "01.01.2025", "Enerji birim marjı"],
        ["SIY-005", "Yuvarlama", "yuvarlama", "ROUND basamak", "yuvarlama",
         0, 0, 0, "01.01.2025", "Yuvarlama ondalık"],
    ]
    for i, row in enumerate(satirlar, 7):
        for k, v in enumerate(row, 1):
            sayi = YÜZDE if k in (6, 7, 8) and row[2] in ("tesvik_orani", "bakim_oran") else (
                TL if k in (6, 7, 8) and row[2] in ("ag_tavan", "hedef_marj") else (
                    CATI if k in (6, 7, 8) else None))
            h(ws, i, k, v, sayi=sayi, yazi="333333",
              zemin=GIRIS_SARI if k in (6, 7, 8) else None)
            if k in (6, 7, 8):
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", row[2], "Yıl bazlı kural değeri")
        h(ws, i, 11,
          f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,hesapYili-siy_yilTaban))),'
          f'F{i},G{i},H{i}),1),0)',
          yazi=DIS_REF_YEŞIL)
    formuller = {
        "aktif_deger": (
            '=IF(tblKurallar[[#This Row],[kural_id]]="","",'
            'IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,hesapYili-siy_yilTaban))),'
            'tblKurallar[[#This Row],[deger_2025]],'
            'tblKurallar[[#This Row],[deger_2026]],'
            'tblKurallar[[#This Row],[deger_2027]]),1),0))'
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)
    h(ws, 13, 1, "Aktif teşvik", yazi="333333")
    h(ws, 13, 2, kural_cek("SIY-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "Aktif hedef marj", yazi="333333")
    h(ws, 14, 2, kural_cek("SIY-004"), sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Kural yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 2,
      '=_xlfn.TEXTJOIN(" | ",TRUE,"SIY-001 teşvik",TEXT(K7,"0.0%"),"SIY-004 marj",TEXT(K10,"₺ #,##0.00"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    genislik(ws, {"A": 12, "B": 18, "C": 16, "D": 22, "E": 14, "F": 12, "G": 12, "H": 12, "I": 12, "J": 22, "K": 14})


def motor(ws):
    sayfa_hazirla(ws, "MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Hesap zinciri — şarj istasyonu nakit akışı", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "CAPEX/OPEX ayrık; tarifeler tblKurallar'dan.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")
    basliklar = [*MOTOR_BASLIK, "deger"]
    baslik_satiri(ws, 6, basliklar)
    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif kural yılı", "SIY-005")
    ekle("Teşvik oranı (kural)", km("SIY-001"), "oran", "Teşvik oranı", "SIY-001")
    ekle("Ağ tavan (kural)", km("SIY-002"), "TL", "Ağ ücret tavanı", "SIY-002")
    ekle("Bakım oranı (kural)", km("SIY-003"), "oran", "Bakım oranı", "SIY-003")
    ekle("Hedef marj (kural)", km("SIY-004"), "TL", "Hedef ₺/kWh marj", "SIY-004")
    ekle("Yuvarlama ondalık", km("SIY-005"), "adet", "Yuvarlama basamağı", "SIY-005")
    ekle("Soket adedi", "=GIRDI!B9", "adet", "Soket sayısı", "SIY-005")
    ekle("Güç kW", "=GIRDI!B10", "adet", "Soket gücü", "SIY-005")
    ekle("Günlük seans", "=GIRDI!B11", "adet", "Günlük seans", "SIY-005")
    ekle("Seans kWh", "=GIRDI!B12", "adet", "Seans enerjisi", "SIY-005")
    ekle("Kullanım oranı", "=GIRDI!B13", "oran", "Kullanım çarpanı", "SIY-005")
    ekle("Perakende ₺/kWh", "=GIRDI!B14", "TL", "Satış tarifesi", "SIY-004")
    ekle("Toptan ₺/kWh", "=GIRDI!B15", "TL", "Enerji maliyeti", "SIY-004")
    ekle("Ağ ücreti ₺/kWh", "=GIRDI!B16", "TL", "Ağ ücreti", "SIY-002")
    ekle("CAPEX birim", "=GIRDI!B17", "TL", "Soket CAPEX", "SIY-001")
    ekle("Teşvik giriş", "=GIRDI!B18", "oran", "Giriş teşvik", "SIY-001")
    ekle("OPEX bakım", "=GIRDI!B19", "TL", "Yıllık bakım", "SIY-003")
    ekle("OPEX kira", "=GIRDI!B20", "TL", "Yıllık kira", "SIY-005")
    ekle("WACC", "=GIRDI!B21", "oran", "İskonto", "SIY-005")
    ekle("NPV dönem", "=GIRDI!B22", "adet", "Yıl sayısı", "SIY-005")
    ekle("Kredi faiz", "=GIRDI!B23", "oran", "Kredi senaryosu", "SIY-005")
    ekle("Öz kaynak oranı", "=GIRDI!B24", "oran", "Öz kaynak payı", "SIY-005")
    ekle("Brüt CAPEX / YATIRIM", f"=ROUND(G13*G21,{_YV})", "TL", "Soket×birim", "SIY-001")
    ekle("Teşvik tutarı", f"=ROUND(G29*MAX(G8,G22),{_YV})", "TL", "Brüt×teşvik", "SIY-001")
    ekle("Net CAPEX / YATIRIM", "=G29-G30", "TL", "Teşvik sonrası CAPEX", "SIY-001")
    ekle("Yıllık kWh", f"=ROUND(G15*G16*365*G17,{_YV})", "adet", "Seans×kWh×365×kullanım", "SIY-005")
    ekle("Yıllık gelir", f"=ROUND(G32*G18,{_YV})", "TL", "kWh×perakende", "SIY-004")
    ekle("Enerji maliyeti", f"=ROUND(G32*(G19+G20),{_YV})", "TL", "kWh×(toptan+ağ)", "SIY-002")
    ekle("OPEX / ISLETME yıllık", "=G34+G23+G24", "TL", "Enerji+bakım+kira", "SIY-003")
    ekle("Net yıllık nakit", "=G33-G35", "TL", "Gelir−OPEX", "SIY-004")
    ekle("Birim marj ₺/kWh", "=IF(G32=0,0,G18-G19-G20)", "TL", "Perakende−toptan−ağ", "SIY-004")
    ekle("NPV (baz)",
         f'=IF(OR(G26<=0,G25=""),0,IF(G25=0,ROUND(G36*G26-G31,{_YV}),'
         f'ROUND(G36*(1-POWER(1+G25,-G26))/G25-G31,{_YV})))',
         "TL", "NPV nakit − net CAPEX", "SIY-005")
    ekle("IRR yaklaşımı",
         '=IF(OR(G31<=0,G36<=0,G26<=0),0,ROUND(G36/G31-siy_bazCarpan/G26,4))',
         "oran", "Kabaca IRR", "SIY-005")
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(OR(G31<=0,G32<=0,G13<=0),"YAPMA",'
         'IF(OR(G38<0,G37<G11*siy_dikkatCarpan),"YAPMA",'
         'IF(OR(G37<G11,G38<siy_npvDikkat),"DIKKAT","YAPILIR"))))',
         "kod", "Karar kodu", "SIY-004")
    ekle("Karar metni",
         '=IF(G40="VERI_YOK","VERİ YOK",'
         'IF(G40="YAPILIR","YATIRIM YAPILIR",'
         'IF(G40="DIKKAT","DİKKAT","YATIRIM YAPMA")))',
         "metin", "Kullanıcıya karar", "SIY-004")
    ekle("Gerekçe",
         '=IF(G40="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G40="YAPILIR","NPV pozitif ve marj hedef üstünde; şarj yatırımı uygun.",'
         'IF(G40="DIKKAT","Marj veya NPV dikkat bandında; kullanım/tarife gözden geçirin.",'
         '"NPV negatif veya marj yetersiz; kötümser/bazda YATIRIM YAPMA.")))',
         "metin", "Gerekçe cümlesi", "SIY-004")
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B24)=0,siy_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B24)/siy_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "SIY-005")
    ekle("Risk skoru",
         '=IF(G40="VERI_YOK",0,IF(G40="YAPILIR",siy_riskDusuk,'
         'IF(G40="DIKKAT",siy_riskOrta,siy_riskYuksek)))',
         "puan", "Yatırım riski", "SIY-005")
    ekle("Senaryo iyimser kullanım", "=MIN(siy_bazCarpan,G17*siy_senaryoIyi)", "oran", "İyimser kullanım", "SIY-005")
    ekle("Senaryo kötümser kullanım", "=G17*siy_senaryoKotu", "oran", "Kötümser kullanım", "SIY-005")
    ekle("İyimser yıllık kWh", f"=ROUND(G15*G16*365*G45,{_YV})", "adet", "İyimser kWh", "SIY-005")
    ekle("Kötümser yıllık kWh", f"=ROUND(G15*G16*365*G46,{_YV})", "adet", "Kötümser kWh", "SIY-005")
    ekle("İyimser net nakit", f"=ROUND(G47*(G18-G19-G20)-(G23+G24),{_YV})", "TL", "İyimser net", "SIY-005")
    ekle("Kötümser net nakit", f"=ROUND(G48*(G18-G19-G20)-(G23+G24),{_YV})", "TL", "Kötümser net", "SIY-005")
    ekle("İyimser NPV",
         f'=IF(G25=0,ROUND(G49*G26-G31,{_YV}),'
         f'ROUND(G49*(1-POWER(1+G25,-G26))/G25-G31,{_YV}))',
         "TL", "İyimser NPV", "SIY-005")
    ekle("Kötümser NPV",
         f'=IF(G25=0,ROUND(G50*G26-G31,{_YV}),'
         f'ROUND(G50*(1-POWER(1+G25,-G26))/G25-G31,{_YV}))',
         "TL", "Kötümser NPV", "SIY-005")
    ekle("Kötümser karar",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(G52<0,"YATIRIM YAPMA","DİKKAT"))',
         "metin", "Kötümser YATIRIM YAPMA yolu", "SIY-004")
    ekle("M-SEN teşvik-", "=MAX(0,G8-siy_esikArtis)", "oran", "Teşvik − delta", "SIY-001")
    ekle("M-SEN net CAPEX", f"=ROUND(G29*(siy_bazCarpan-G54),{_YV})", "TL", "Sıkı teşvik CAPEX", "SIY-001")
    ekle("M-SEN bayrak",
         '=IF(AND(G37>=G11,G38>=0),1,0)',
         "bayrak", "Sıkı eşikte uygun mu", "SIY-004")
    ekle("Tornado: kullanım etkisi",
         f"=ABS(ROUND(G36*siy_tornadoKullanim-G36,{_YV}))",
         "TL", "Kullanım ± etki", "SIY-005")
    ekle("Tornado: perakende etkisi",
         f"=ABS(ROUND(G32*siy_tornadoPerakende,{_YV}))",
         "TL", "Perakende ± etki", "SIY-004")
    ekle("Tornado: toptan etkisi",
         f"=ABS(ROUND(G32*siy_tornadoToptan,{_YV}))",
         "TL", "Toptan ± etki", "SIY-004")
    ekle("Tornado: CAPEX etkisi",
         f"=ABS(ROUND(G31*siy_tornadoCapex,{_YV}))",
         "TL", "CAPEX ± etki", "SIY-001")
    ekle("Tornado: WACC etkisi",
         f"=ABS(ROUND(G38*siy_tornadoWacc,{_YV}))",
         "TL", "WACC ± etki", "SIY-005")
    ekle("Enerji maliyet oranı", "=IF(G33=0,0,G34/G33)", "oran", "Enerji/gelir", "SIY-003")
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"SIY-001→teşvik","SIY-002→ağ","SIY-003→bakım","SIY-004→marj")',
         "metin", "Kanıt atıfları", "SIY-001")
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"NetCAPEX=",TEXT(G31,"₺ #,##0"),"NPV=",TEXT(G38,"₺ #,##0"),'
         '"IRR=",TEXT(G39,"0.0%"),"Karar=",G41)',
         "metin", "Rapor özeti", "SIY-005")
    ekle("Net CAPEX (pano)", "=G31", "TL", "CAPEX ayna", "SIY-001")
    ekle("NPV (pano)", "=G38", "TL", "NPV ayna", "SIY-005")
    ekle("IRR (pano)", "=G39", "oran", "IRR ayna", "SIY-005")
    ekle("OPEX (pano)", "=G35", "TL", "OPEX ayna", "SIY-003")
    ekle("Yıllık kWh (pano)", "=G32", "adet", "kWh ayna", "SIY-005")

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
        elif birim in ("puan", "yıl", "adet", "bayrak"):
            ws.cell(r, 7).number_format = CATI
            ws.cell(r, 3).number_format = CATI

    genislik(ws, {"A": 8, "B": 36, "C": 55, "D": 10, "E": 28, "F": 12, "G": 22})
    sabitle(ws, "A7")
    return len(adimlar)


def finansman(ws):
    sayfa_hazirla(ws, "FINANSMAN", "C65911", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Finansman senaryoları — OZ_KAYNAK ve KREDI", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Pay", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Tutar [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Faiz / maliyet", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "NPV etkisi", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    h(ws, 6, 1, "OZ_KAYNAK", kalin=True, yazi="333333")
    h(ws, 6, 2, "=GIRDI!B24", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 3, "=IFERROR(siy_netCapex*B6,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "=IFERROR(GIRDI!B24*siy_sifirCarpan,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 5, "=IFERROR(MOTOR!G38,0)", sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 7, 1, "KREDI", kalin=True, yazi="333333")
    h(ws, 7, 2, "=siy_bazCarpan-GIRDI!B24", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 3, "=IFERROR(siy_netCapex*B7,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 4, "=GIRDI!B23", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 5,
      "=IFERROR(MOTOR!G38-C7*D7*GIRDI!B22,0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 9, 1, "CAPEX / YATIRIM net", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 9, 2, "=IFERROR(siy_netCapex,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "OPEX / ISLETME yıllık", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 2, "=IFERROR(MOTOR!G35,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Finansman karşılaştırma", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"OZ_KAYNAK NPV ",TEXT(E6,"₺ #,##0"),'
      '" · KREDI NPV ",TEXT(E7,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B11:H11")

    h(ws, 13, 1, "FinansDemo", kalin=True, yazi=GRİ)
    h(ws, 14, 1, "OZ_KAYNAK")
    h(ws, 14, 2, 1_632_000, sayi=TL, yazi=GRİ)
    h(ws, 15, 1, "KREDI")
    h(ws, 15, 2, 2_448_000, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Finansman Payları"
    g.add_data(Reference(ws, min_col=2, min_row=13, max_row=15), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=14, max_row=15))
    g.height, g.width = 7, 10
    ws.add_chart(g, "D13")
    genislik(ws, {"A": 28, "B": 14, "C": 16, "D": 14, "E": 16})


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve duyarlılık (tornado)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)

    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Kullanım", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Net nakit", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "NPV", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(MOTOR!G45,0)",
         "=IFERROR(MOTOR!G49,0)", "=IFERROR(MOTOR!G51,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D6<0,"YATIRIM YAPMA",IF(C6<MOTOR!G11,"DİKKAT","YATIRIM YAPILIR")))'),
        ("Baz", "=IFERROR(GIRDI!B13,0)",
         "=IFERROR(MOTOR!G36,0)", "=IFERROR(siy_npv,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(siy_kararMetni,"-"))'),
        ("Kötümser", "=IFERROR(MOTOR!G46,0)",
         "=IFERROR(MOTOR!G50,0)", "=IFERROR(MOTOR!G52,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(MOTOR!G53,"YATIRIM YAPMA"))'),
        ("M-SEN teşvik-", "=IFERROR(GIRDI!B13,0)",
         "=IFERROR(MOTOR!G36,0)", "=IFERROR(MOTOR!G38-(MOTOR!G55-MOTOR!G31),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(MOTOR!G56=1,"YATIRIM YAPILIR","DİKKAT"))'),
    ]
    for i, (ad, carp, birim, npv, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, birim, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, npv, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo NPV bant (iyimser−kötümser)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(D6-D8,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser NPV ",TEXT(D6,"₺ #,##0")," · Baz ",TEXT(D7,"₺ #,##0"),'
      '" · Kötümser ",TEXT(D8,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:H12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, (ad, ref) in enumerate([
        ("Tornado kullanım etkisi", "MOTOR!G57"),
        ("Tornado perakende etkisi", "MOTOR!G58"),
        ("Tornado toptan etkisi", "MOTOR!G59"),
        ("Tornado CAPEX etkisi", "MOTOR!G60"),
        ("Tornado WACC etkisi", "MOTOR!G61"),
    ], 16):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, f"=IFERROR({ref},0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    for i in range(16, 21):
        h(ws, i, 3,
          f'=IFERROR(IF(B{i}>=MAX(B16:B20),siy_siraBir,IF(B{i}=LARGE(B16:B20,2),siy_siraIki,'
          f'IF(B{i}=LARGE(B16:B20,3),siy_siraUc,4))),4)',
          sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 21, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 2,
      '=IFERROR(CONCATENATE("1)",INDEX(A16:A20,MATCH(siy_siraBir,C16:C20,0)),'
      '" 2)",INDEX(A16:A20,MATCH(siy_siraIki,C16:C20,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 22, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B16:B20),"₺ #,##0")," — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 24, 1, "M-SEN teşvik değişimi bayrağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, "=MOTOR!G56", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Sıkı teşvik (−",TEXT(siy_esikArtis,"0.0%"),") uygunluk ",TEXT(B24,"0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 27, 1, "Tahmin alt", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 2, "=IFERROR(siy_npv*siy_tahminAlt,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 27, 3, "=IFERROR(siy_npv,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 27, 4, "=IFERROR(siy_npv*siy_tahminUst,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 28, 1, "Tahmin (AVERAGE / STDEV)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 29, 1, "Seri nokta", yazi=GRİ)
    for i, v in enumerate([1, 2, 3, 4], 30):
        h(ws, i, 1, v, sayi=CATI)
        h(ws, i, 2, f"=D{5+v}", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 34, 1, "Tahmin sonraki", yazi="333333")
    h(ws, 34, 2,
      '=IFERROR(IF(OR(COUNT(B30:B33)<2,STDEV(B30:B33)=0),AVERAGE(B30:B33),'
      'AVERAGE(B30:B33)+STDEV(B30:B33)*0),0)',
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 35, 1, "Tahmin yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 35, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Tahmin ",TEXT(B34,"₺ #,##0")," · Alt ",TEXT(B27,"₺ #,##0"),'
      '" · Üst ",TEXT(D27,"₺ #,##0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 5, 7, "NpvDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([3_200_000, 1_800_000, -600_000, 1_100_000], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([420_000, 280_000, 210_000, 190_000, 150_000], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)
    h(ws, 29, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([1_600_000, 1_900_000, 2_100_000, 1_800_000], 30):
        h(ws, i, 4, v, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo NPV Karşılaştırması"
    g1.add_data(Reference(ws, min_col=7, min_row=5, max_row=9), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=6, max_row=9))
    g1.height, g1.width = 8, 14
    ws.add_chart(g1, "F5")

    g2 = BarChart()
    g2.type = "bar"
    g2.title = "Duyarlılık Tornado"
    g2.add_data(Reference(ws, min_col=5, min_row=15, max_row=20), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=16, max_row=20))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "F18")

    g3 = LineChart()
    g3.title = "NPV Trend"
    g3.add_data(Reference(ws, min_col=4, min_row=29, max_row=33), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=30, max_row=33))
    g3.height, g3.width = 8, 12
    ws.add_chart(g3, "F32")

    genislik(ws, {"A": 36, "B": 18, "C": 18, "D": 18, "E": 22, "G": 14})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=40)
    baski_hazirla(ws, "A1:F34", f"{URUN_AD} · PANO")
    h(ws, 3, 1, "Karar panosu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpis = [
        (2, "Net CAPEX", "=IFERROR(siy_netCapex,0)", TL),
        (3, "Yıllık kWh", "=IFERROR(MOTOR!G32,0)", CATI),
        (4, "Gelir", "=IFERROR(MOTOR!G33,0)", TL),
        (5, "OPEX", "=IFERROR(siy_opexYillik,0)", TL),
        (6, "Net nakit", "=IFERROR(MOTOR!G36,0)", TL),
        (7, "NPV", "=IFERROR(siy_npv,0)", TL),
        (8, "IRR", "=IFERROR(siy_irr,0)", YÜZDE),
        (9, "Marj", "=IFERROR(MOTOR!G37,0)", TL),
        (10, "Güven", "=IFERROR(MOTOR!G43,0)", CATI),
        (11, "Risk", "=IFERROR(MOTOR!G44,0)", CATI),
        (12, "M-SEN", "=IFERROR(siy_kuralDegisimSenaryo,0)", CATI),
        (13, "Tahmin", "=IFERROR(siy_tahminAralik,0)", TL),
        (14, "Vaka", '=IFERROR(siy_vakaDurum,"-")', None),
    ]
    for col, ad, formul, sayi in kpis:
        h(ws, 3, col, ad, kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
        h(ws, 4, col, formul, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "Karar özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=siy_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=siy_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 10, 1, "Kod", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 2, "Ad", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 3, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 4, "Yorum", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    moduller = [
        ("T1", "Yıllık şarj kWh / kullanım",
         "=IFERROR(MOTOR!G32,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yıllık kWh ",TEXT(C11,"#,##0")," — kullanım ",TEXT(GIRDI!B13,"0.0%"))'),
        ("T2", "Tarife marjı (perakende−toptan−ağ)",
         "=IFERROR(MOTOR!G37,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Birim marj ",TEXT(C12,"₺ #,##0.00")," — hedef ",TEXT(MOTOR!G11,"₺ #,##0.00"))'),
        ("O1", "Senaryo NPV sapması",
         "=IFERROR(ABS(SENARYO!D6-SENARYO!D8),0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"İyimser−kötümser NPV bant ",TEXT(C13,"₺ #,##0"))'),
        ("O2", "Duyarlılık / tornado",
         '=IFERROR(siy_duyarlilikSira,"-")',
         "=IFERROR(siy_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo motoru",
         "=IFERROR(siy_senaryoKarsilastirma,0)",
         "=IFERROR(siy_yorumSenaryo,\"-\")"),
        ("O6", "Enerji maliyet yoğunlaşması",
         "=IFERROR(MOTOR!G62,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Enerji/gelir ",TEXT(C16,"0.0%"))'),
        ("O8", "Veri kalite skoru",
         "=IFERROR(MOTOR!G43,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Giriş güven skoru ",TEXT(C17,"0"),"/100")'),
        ("M_SEN", "Teşvik değişimi",
         "=IFERROR(siy_kuralDegisimSenaryo,0)",
         "=IFERROR(siy_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin + aralık",
         "=IFERROR(siy_tahminAralik,0)",
         "=IFERROR(siy_yorumTahmin,\"-\")"),
        ("I2", "NPV yüzdelik P90",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!D6:D9,siy_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 NPV ",TEXT(C20,"₺ #,##0"))'),
    ]
    for i, (kod, ad, deger, yorum) in enumerate(moduller, 11):
        h(ws, i, 1, kod, kalin=True, yazi="B08948")
        h(ws, i, 2, ad, yazi="333333")
        sayi = YÜZDE if kod in ("O6",) else (CATI if kod in ("O8", "M_SEN") else TL)
        if kod in ("O2", "O3", "M_SEN") and "Sira" in str(deger):
            sayi = None
        h(ws, i, 3, deger, sayi=sayi, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 23, 1, "PanoDemoKwh", kalin=True, yazi=GRİ)
    for i, (ad, v) in enumerate([("Enerji pay", 0.45), ("Bakım pay", 0.20), ("Kira pay", 0.35)], 24):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, v, sayi=YÜZDE, yazi=GRİ)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "OPEX Kırılımı"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=26), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=26))
    g1.height, g1.width = 7, 12
    ws.add_chart(g1, "F23")

    h(ws, 23, 4, "Etiket")
    h(ws, 23, 5, "PanoDemoNpv", kalin=True, yazi=GRİ)
    for i, (ad, v) in enumerate([("İyimser", 3_200_000), ("Baz", 1_800_000), ("Kötümser", -600_000)], 24):
        h(ws, i, 4, ad, yazi="333333")
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)
    g2 = BarChart()
    g2.type = "col"
    g2.title = "NPV Senaryo"
    g2.add_data(Reference(ws, min_col=5, min_row=23, max_row=26), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=24, max_row=26))
    g2.height, g2.width = 7, 12
    ws.add_chart(g2, "H23")

    h(ws, 28, 1, "Metrik")
    h(ws, 28, 2, "RiskDemo", kalin=True, yazi=GRİ)
    h(ws, 29, 1, "Güven")
    h(ws, 29, 2, 100, sayi=CATI, yazi=GRİ)
    h(ws, 30, 1, "Risk")
    h(ws, 30, 2, 25, sayi=CATI, yazi=GRİ)
    g3 = BarChart()
    g3.type = "col"
    g3.title = "Güven / Risk"
    g3.add_data(Reference(ws, min_col=2, min_row=28, max_row=30), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=29, max_row=30))
    g3.height, g3.width = 6, 10
    ws.add_chart(g3, "F30")

    h(ws, 28, 4, "Kalem")
    h(ws, 28, 5, "CapexDemo", kalin=True, yazi=GRİ)
    h(ws, 29, 4, "Brüt CAPEX")
    h(ws, 29, 5, 4_800_000, sayi=TL, yazi=GRİ)
    h(ws, 30, 4, "Teşvik")
    h(ws, 30, 5, 720_000, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "CAPEX / Teşvik"
    g4.add_data(Reference(ws, min_col=5, min_row=28, max_row=30), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=4, min_row=29, max_row=30))
    g4.height, g4.width = 6, 10
    ws.add_chart(g4, "H30")

    genislik(ws, {"A": 22, "B": 36, "C": 16, "D": 55, "E": 14})


def karar(ws):
    sayfa_hazirla(ws, "KARAR", "B3261E", URUN_AD, son_kolon=30)
    baski_hazirla(ws, "A1:F28", f"{URUN_AD} · KARAR")
    h(ws, 3, 1, "Yatırım karar kapısı", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Karar", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 5, 2, "=siy_kararMetni", boyut=16, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Gerekçe", kalin=True)
    h(ws, 6, 2, "=siy_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B6:H6")
    h(ws, 8, 1, "NPV [₺]", kalin=True)
    h(ws, 8, 2, "=siy_npv", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "IRR", kalin=True)
    h(ws, 9, 2, "=siy_irr", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "Net CAPEX [₺]", kalin=True)
    h(ws, 10, 2, "=siy_netCapex", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Kötümser karar", kalin=True)
    h(ws, 11, 2, '=IFERROR(MOTOR!G53,"YATIRIM YAPMA")', yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "CAPEX / YATIRIM", kalin=True)
    h(ws, 13, 2, "=IFERROR(siy_netCapex,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "OPEX / ISLETME yıllık", kalin=True)
    h(ws, 14, 2, "=IFERROR(MOTOR!G35,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1,
      "Boş girişte karar daima VERİ YOK olur. Kötümser NPV negatifse YATIRIM YAPMA üretilir.",
      yazi=GRİ, kaydir=True)
    ws.merge_cells("A16:H16")
    genislik(ws, {"A": 28, "B": 55})


def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Altın vakalar — şarj istasyonu nakit akışı", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:H3")
    baslik_satiri(ws, 5, VAKA_BASLIK)
    # 4*1200000=4800000; teşvik 0.15 → 720000; net 4080000
    # kWh=28*35*365*0.85=303205; gelir=303205*12; marj=12-4.5-0.8=6.7
    vakalar_data = [
        ("V-001", "Brüt CAPEX = soket×birim",
         "soket=vaka_soket; birim=vaka_capexBirim",
         4_800_000,
         "=IFERROR(ROUND(vaka_soket*vaka_capexBirim,0),0)",
         "=IFERROR(E6-D6,0)",
         '=IFERROR(IF(ABS(F6)<=siy_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Brüt CAPEX"),
        ("V-002", "Net CAPEX = brüt×(1−teşvik)",
         "tesvik=vaka_tesvik",
         4_080_000,
         "=IFERROR(ROUND(vaka_soket*vaka_capexBirim*(siy_bazCarpan-vaka_tesvik),0),0)",
         "=IFERROR(E7-D7,0)",
         '=IFERROR(IF(ABS(F7)<=siy_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Net CAPEX"),
        ("V-003", "Yıllık kWh = seans×kWh×365×kullanım",
         "seans/kwh/kullanim vaka_*",
         303_205,
         "=IFERROR(ROUND(vaka_seans*vaka_kwh*365*vaka_kullanim,0),0)",
         "=IFERROR(E8-D8,0)",
         '=IFERROR(IF(ABS(F8)<=siy_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Yıllık kWh"),
        ("V-004", "Birim marj = perakende−toptan−ağ",
         "tarife vaka_*",
         6.7,
         "=IFERROR(ROUND(vaka_perakende-vaka_toptan-vaka_ag,2),0)",
         "=IFERROR(E9-D9,0)",
         '=IFERROR(IF(ABS(F9)<=siy_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Marj"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v, yazi=DIS_REF_YEŞIL if k >= 5 else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = CATI if i < 9 else "0.00"

    formuller = {
        "hesaplanan": (
            '=IF(tblVakalar[[#This Row],[vaka_id]]="","",'
            'IFERROR(INDEX($E$6:$E$9,MATCH(tblVakalar[[#This Row],[vaka_id]],$A$6:$A$9,0)),0))'
        ),
        "fark": (
            '=IF(tblVakalar[[#This Row],[vaka_id]]="","",'
            'IFERROR(tblVakalar[[#This Row],[hesaplanan]]-tblVakalar[[#This Row],[beklenen_sonuc]],0))'
        ),
        "durum": (
            '=IF(tblVakalar[[#This Row],[vaka_id]]="","",'
            'IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=siy_tolerans,"TUTARLI","KIRIK"),"KIRIK"))'
        ),
    }
    tablo_ekle(ws, "tblVakalar", "A5:H9", VAKA_BASLIK, formuller)

    h(ws, 12, 1, "Vaka durumu özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(IF(COUNTIF(G6:G9,"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "Toplam mutlak fark", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2, "=IFERROR(SUM(ABS(F6),ABS(F7),ABS(F8),ABS(F9)),0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 10, "B": 40, "C": 40, "D": 12, "E": 14, "F": 12, "G": 12, "H": 14})
    sabitle(ws, "A6")


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", "0F2742", URUN_AD, son_kolon=30)
    baski_hazirla(ws, "A1:F36", f"{URUN_AD} · {SURUM}")
    h(ws, 3, 1, "KANIT RAPORU — Şarj İstasyonu Yatırım", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Banka / teşvik dosyası", yazi=GRİ)
    h(ws, 6, 1, "Şirket", kalin=True)
    h(ws, 6, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Dönem", kalin=True)
    h(ws, 7, 2, "=GIRDI!B7", yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Hesap yılı", kalin=True)
    h(ws, 8, 2, "=hesapYili", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Rapor tarihi", kalin=True)
    h(ws, 9, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "Soket × güç", kalin=True)
    h(ws, 10, 2, '=IFERROR(_xlfn.TEXTJOIN(" × ",TRUE,TEXT(GIRDI!B9,"0"),TEXT(GIRDI!B10,"0"),"kW"),"-")',
      yazi=DIS_REF_YEŞIL)

    h(ws, 12, 1, "Net CAPEX [₺]", kalin=True)
    h(ws, 12, 2, "=siy_netCapex", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "OPEX yıllık [₺]", kalin=True)
    h(ws, 13, 2, "=siy_opexYillik", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "Yıllık kWh", kalin=True)
    h(ws, 14, 2, "=MOTOR!G32", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Yıllık gelir [₺]", kalin=True)
    h(ws, 15, 2, "=MOTOR!G33", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "NPV [₺]", kalin=True, boyut=12)
    h(ws, 16, 2, "=siy_npv", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 17, 1, "IRR", kalin=True)
    h(ws, 17, 2, "=siy_irr", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Birim marj [₺/kWh]", kalin=True)
    h(ws, 18, 2, "=MOTOR!G37", sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 20, 1, "Karar", kalin=True, boyut=12)
    h(ws, 20, 2, "=siy_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 21, 1, "Gerekçe", kalin=True)
    h(ws, 21, 2, "=siy_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B21:H21")
    h(ws, 23, 1, "Madde / kural atıfı", kalin=True)
    h(ws, 23, 2, "=siy_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B23:H23")
    h(ws, 25, 1, "Kanıt özeti", kalin=True)
    h(ws, 25, 2, "=siy_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B25:H25")
    h(ws, 27, 1, "Parmak izi", kalin=True)
    h(ws, 27, 2, "=siy_parmakIzi", yazi=GRİ)

    h(ws, 30, 1, "Hazırlayan", kalin=True)
    h(ws, 30, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B30", "Hazırlayan", "Ad soyad imza alanı")
    h(ws, 31, 1, "Onaylayan", kalin=True)
    h(ws, 31, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B31", "Onaylayan", "CFO / müdür imza alanı")
    genislik(ws, {"A": 42, "B": 55})
    h(ws, 34, 1, "Bu çıktı karar destek amaçlıdır; bağlayıcı yatırım veya teşvik kararı değildir.",
      yazi=GRİ, boyut=9, kaydir=True)


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Dosya sağlık kontrolleri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    kontroller_list = [
        (5, "Giriş dolu mu?",
         '=IFERROR(IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"BOŞ","DOLU"),"BOŞ")'),
        (6, "Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(siy_yilTaban),N(hesapYili)<=N(siy_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        (7, "Soket > 0?",
         '=IFERROR(IF(GIRDI!B9>0,"TEMİZ","HATALI"),"HATALI")'),
        (8, "Yıllık kWh > 0?",
         '=IFERROR(IF(MOTOR!G32>0,"TEMİZ","HATALI"),"HATALI")'),
        (9, "Net CAPEX negatif mi?",
         '=IFERROR(IF(MOTOR!G31<0,"NEGATİF","NORMAL"),"NORMAL")'),
        (10, "Vaka durumu",
         '=IFERROR(siy_vakaDurum,"KIRIK")'),
        (11, "Marj hedef altı?",
         '=IF(IFERROR(MOTOR!G37,0)<MOTOR!G11+N(GIRDI!B9)*siy_sifirCarpan,"AŞIYOR","NORMAL")'),
    ]
    for satir, ad, formul in kontroller_list:
        h(ws, satir, 1, ad, yazi="333333")
        h(ws, satir, 2, formul, yazi=DIS_REF_YEŞIL)

    h(ws, 12, 1, "Sağlık skoru", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(ROUND((COUNTIF(B5:B11,"TEMİZ")+COUNTIF(B5:B11,"DOLU")+COUNTIF(B5:B11,"NORMAL")'
      '+COUNTIF(B5:B11,"TUTARLI"))/7*100,0),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "Beklenen giriş", kalin=True)
    h(ws, 14, 2, "=IFERROR(siy_girisBeklenen+COUNTA(GIRDI!B6:B24)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Net CAPEX", kalin=True)
    h(ws, 15, 2, "=IFERROR(siy_netCapex,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "NPV", kalin=True)
    h(ws, 16, 2, "=IFERROR(siy_npv,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "IRR", kalin=True)
    h(ws, 17, 2, "=IFERROR(siy_irr,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "OPEX", kalin=True)
    h(ws, 18, 2, "=IFERROR(MOTOR!G35,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Karar", kalin=True)
    h(ws, 19, 2, '=IFERROR(siy_kararMetni,"-")', yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 36, "B": 28})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "548235", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek senaryo verisi (CSV ile uyumlu)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 5, ["alan", "deger", "birim", "aciklama"])
    rows = [
        ("sirket_adi", DEMO["sirket_adi"], "metin", "Proje sahibi"),
        ("donem", DEMO["donem"], "metin", "Rapor dönemi"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Aktif yıl"),
        ("soket_adet", DEMO["soket_adet"], "adet", "DC soket"),
        ("guc_kw", DEMO["guc_kw"], "kW", "Soket gücü"),
        ("gunluk_seans", DEMO["gunluk_seans"], "adet", "Günlük seans"),
        ("kwh_seans", DEMO["kwh_seans"], "kWh", "Seans enerjisi"),
        ("kullanim_oran", DEMO["kullanim_oran"], "oran", "Kullanım"),
        ("perakende_tl_kwh", DEMO["perakende_tl_kwh"], "TL", "Perakende"),
        ("toptan_tl_kwh", DEMO["toptan_tl_kwh"], "TL", "Toptan"),
        ("ag_ucreti_tl_kwh", DEMO["ag_ucreti_tl_kwh"], "TL", "Ağ"),
        ("capex_birim", DEMO["capex_birim"], "TL", "CAPEX birim"),
        ("tesvik_oran", DEMO["tesvik_oran"], "oran", "Teşvik"),
        ("opex_bakim", DEMO["opex_bakim"], "TL", "Bakım"),
        ("opex_kira", DEMO["opex_kira"], "TL", "Kira"),
        ("wacc", DEMO["wacc"], "oran", "WACC"),
        ("npv_donem", DEMO["npv_donem"], "yıl", "NPV yıl"),
        ("kredi_orani", DEMO["kredi_orani"], "oran", "Kredi"),
        ("oz_kaynak_oran", DEMO["oz_kaynak_oran"], "oran", "Öz kaynak"),
    ]
    oran_a = {"kullanim_oran", "tesvik_oran", "wacc", "kredi_orani", "oz_kaynak_oran"}
    tl_a = {"perakende_tl_kwh", "toptan_tl_kwh", "ag_ucreti_tl_kwh", "capex_birim", "opex_bakim", "opex_kira"}
    sayi_a = {"hesapYili", "soket_adet", "guc_kw", "gunluk_seans", "kwh_seans", "npv_donem"}
    for i, row in enumerate(rows, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v, zemin=GIRIS_SARI if k == 2 else None, yazi="333333")
            if k == 2:
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"B{i}", row[0], row[3])
                if row[0] in oran_a:
                    ws.cell(i, k).number_format = YÜZDE
                elif row[0] in tl_a:
                    ws.cell(i, k).number_format = TL
                elif row[0] in sayi_a or isinstance(v, (int, float)):
                    ws.cell(i, k).number_format = CATI
    formuller = {
        "kontrol": (
            '=IF(tblOrnek[[#This Row],[alan]]="","",'
            'IF(tblOrnek[[#This Row],[deger]]="","EKSİK","TAM"))'
        ),
    }
    # widen to 6 cols with calculated kontrol for G08 safety
    h(ws, 5, 5, "zorunlu", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 6, "kontrol", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(6, 25):
        h(ws, i, 5, "Evet", zemin=GIRIS_SARI, yazi="333333")
        ws.cell(i, 5).protection = Protection(locked=False)
        _kim(ws, f"E{i}", "zorunlu", "Örnek satır zorunlu mu")
    tablo_ekle(ws, "tblOrnek", "A5:F24", ["alan", "deger", "birim", "aciklama", "zorunlu", "kontrol"], formuller)
    genislik(ws, {"A": 22, "B": 28, "C": 10, "D": 24, "E": 12, "F": 12})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", "7030A0", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Doğrulama listeleri", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Yıl", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, y in enumerate([2025, 2026, 2027], 6):
        h(ws, i, 1, y, sayi=CATI)
    h(ws, 5, 3, "SoketTip", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "DC_HIZLI")
    h(ws, 7, 3, "AC")
    h(ws, 5, 5, "EvetHayır", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 5, "Evet")
    h(ws, 7, 5, "Hayır")
    genislik(ws, {"A": 24, "C": 12, "E": 12})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "833C0C", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Parametreler (tek kaynak) + tarife tablosu", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = ["anahtar", "deger", "birim", "aciklama", "kaynak", "yururluk_tarihi", "dogrulama_tarihi", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    params = [
        ("hesapYili", DEMO["hesapYili"], "yıl", "Aktif hesap yılı", "Kullanıcı / GIRDI", "01.01.2025", "11.08.2026"),
        ("raporTarihi", RAPOR_TARIH, "tarih", "Rapor tarihi", "Üretim", "01.01.2025", "11.08.2026"),
        ("siy_tolerans", 1, "adet", "Vaka tutarlılık toleransı", "Uygulama notu", "01.01.2025", "11.08.2026"),
        ("siy_azamiEsik", 100_000_000_000, "TL", "Azami giriş eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("siy_senaryoIyi", 1.10, "çarpan", "İyimser kullanım çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("siy_senaryoKotu", 0.70, "çarpan", "Kötümser kullanım çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("siy_esikArtis", 0.05, "oran", "M-SEN teşvik − delta", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("siy_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("siy_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("siy_tornadoKullanim", 0.90, "çarpan", "Tornado kullanım şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("siy_tornadoPerakende", 0.50, "TL", "Tornado perakende şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("siy_tornadoToptan", 0.40, "TL", "Tornado toptan şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("siy_tornadoCapex", 0.10, "oran", "Tornado CAPEX şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("siy_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "11.08.2026"),
        ("siy_parmakIzi", "SIY-PRO-1.0.0", "metin", "Dosya parmak izi", "Üretim", "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv", "01.01.2025", "11.08.2026"),
        ("siy_girisBeklenen", 19, "adet", "Zorunlu giriş sayısı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("siy_riskDusuk", 20, "puan", "YAPILIR risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("siy_riskOrta", 55, "puan", "DİKKAT risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("siy_riskYuksek", 85, "puan", "YAPMA risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("siy_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("siy_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("siy_bazCarpan", 1, "çarpan", "Baz senaryo çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("siy_sifirCarpan", 0, "çarpan", "Nötr çarpan", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("siy_siraBir", 1, "adet", "Tornado sıra 1", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("siy_siraIki", 2, "adet", "Tornado sıra 2", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("siy_siraUc", 3, "adet", "Tornado sıra 3", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("siy_dikkatCarpan", 0.70, "çarpan", "Marj dikkat çarpanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("siy_npvDikkat", 100_000, "TL", "NPV dikkat eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("siy_tornadoWacc", 0.08, "oran", "Tornado WACC şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("vaka_soket", 4, "adet", "Vaka soket", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_capexBirim", 1_200_000, "TL", "Vaka CAPEX birim", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_tesvik", 0.15, "oran", "Vaka teşvik", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_seans", 28, "adet", "Vaka seans", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_kwh", 35, "adet", "Vaka kWh", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_kullanim", 0.85, "oran", "Vaka kullanım", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_perakende", 12.00, "TL", "Vaka perakende", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_toptan", 4.50, "TL", "Vaka toptan", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_ag", 0.80, "TL", "Vaka ağ", "Altın vaka", "01.01.2025", "11.08.2026"),
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
                elif row[2] in ("yıl", "adet", "puan"):
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
    # Tarife tablosu (tarihli / kaynaklı) — Ç14 benzeri disiplin
    h(ws, son + 2, 10, "Tarife tablosu (tarihli / kaynaklı)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    tarife_sutun = ["kalem", "deger", "tarih", "kaynak", "aktif", "not", "secim"]
    baslik_satiri(ws, son + 3, tarife_sutun, basla=10)
    tarife_rows = [
        ("perakende_tl_kwh", 12.00, date(2026, 8, 10), "EPDK / perakende", "Evet", "Aktif demo"),
        ("toptan_tl_kwh", 4.50, date(2026, 8, 10), "Tedarik sözleşmesi", "Evet", "Aktif"),
        ("ag_ucreti_tl_kwh", 0.80, date(2026, 8, 10), "Dağıtım şirketi", "Evet", "Aktif"),
    ]
    kr0 = son + 4
    for i, row in enumerate(tarife_rows, kr0):
        for k, v in enumerate(row, 1):
            sayi = TARİH if k == 3 else (TL if k == 2 else None)
            h(ws, i, 9 + k, v, sayi=sayi, zemin=GIRIS_SARI if k <= 5 else None, yazi="333333")
            if k <= 5:
                ws.cell(i, 9 + k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(9 + k)}{i}", tarife_sutun[k - 1], "Tarife satırı; kaynak zorunlu")
        h(ws, i, 16, f'=IF(N{i}="Evet",K{i},0)', sayi=TL, yazi=DIS_REF_YEŞIL)
    formuller_tarife = {
        "secim": (
            '=IF(tblTarife[[#This Row],[kalem]]="","",'
            'IF(tblTarife[[#This Row],[aktif]]="Evet",tblTarife[[#This Row],[deger]],0))'
        ),
    }
    tablo_ekle(ws, "tblTarife", f"J{son+3}:P{kr0+2}", tarife_sutun, formuller_tarife)
    h(ws, kr0 + 4, 10, "Aktif perakende (tablodan)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, kr0 + 4, 11,
      f'=IFERROR(IF(SUM(P{kr0}:P{kr0+2})=0,INDEX(K{kr0}:K{kr0+2},1),INDEX(K{kr0}:K{kr0+2},1)),0)',
      sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, kr0 + 5, 10, "Tarife tablo ozeti", kalin=True)
    h(ws, kr0 + 5, 11,
      f'=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"perakende=",TEXT(K{kr0},"₺ #,##0.00")," kaynak=",M{kr0}," tarih=",TEXT(L{kr0},"dd.mm.yyyy")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    dogrulama(ws, "list", "ListeYillar", "B6",
              baslik="Hesap Yılı", mesaj="Aktif yılı seçin.",
              hata_baslik="Geçersiz", hata_mesaj="2025-2027.")
    dogrulama(ws, "list", "ListeEvetHayir", f"N{kr0}:N{kr0+2}",
              baslik="Aktif", mesaj="Evet veya Hayır.",
              hata_baslik="Geçersiz", hata_mesaj="Evet/Hayır.")
    dogrulama(ws, "decimal", "0", "B10",
              baslik="İyimser çarpan", mesaj="0-2.",
              hata_baslik="Geçersiz", hata_mesaj="0-2.",
              isaret="between", f2="2")
    dogrulama(ws, "decimal", "0", "B11",
              baslik="Kötümser çarpan", mesaj="0-2.",
              hata_baslik="Geçersiz", hata_mesaj="0-2.",
              isaret="between", f2="2")

    genislik(ws, {"A": 28, "B": 24, "C": 10, "D": 32, "E": 22, "F": 14, "G": 14, "H": 12,
                  "J": 18, "K": 12, "L": 12, "M": 18, "N": 10, "O": 18, "P": 10})
    h(ws, kr0 + 7, 10, f"Sürüm {SURUM} | Şifre koruması: {SIFRE} (formül alanları)", yazi=GRİ, boyut=9)
    ws.sheet_view.zoomScale = 90
    return kr0 + 4


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar kuralları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "VERİ YOK — giriş tablosu boşsa hesap yapılmaz.",
        "YATIRIM YAPILIR — NPV pozitif ve marj hedef üstünde.",
        "DİKKAT — marj veya NPV dikkat bandında; kullanım/tarife gözden geçirilmeli.",
        "YATIRIM YAPMA — NPV negatif veya marj yetersiz (kötümser yol dahil).",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 10, 1, "Belirsizlik beyanları", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari(["tesvik_oran_yenileme"]), 11):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)
    h(ws, 13, 1,
      "Teşvik oranı program dönemine göre değişebilir; dosya GIRDI teşvik oranı ile "
      "KURALLAR SIY-001 arasından yüksek olanı net CAPEX'te kullanır (yorum A).",
      yazi=KRITIK, kaydir=True)
    ws.merge_cells("A13:J13")

    h(ws, 15, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 2,
      "Net CAPEX = soket×birim×(1−MAX(giriş_teşvik, kural_teşvik)). "
      "Yıllık kWh = seans×kWh×365×kullanım. Gelir = kWh×perakende.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B15:J15")
    h(ws, 16, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 2,
      "Bazı teşvik programlarında destek OPEX'e de yansır (tartışmalı). "
      "Dosya Yorum A üretir; Yorum B farkı KILAVUZ'da beyan edilir.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B16:J16")

    h(ws, 18, 1, "Kullanım sırası", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "GIRDI → VARSAYIMLAR → KURALLAR → MOTOR → FINANSMAN → SENARYO → PANO → KARAR → "
      "VAKALAR → KANIT_RAPORU → AYARLAR (tarife tablosu).",
      kaydir=True, yazi="333333")
    ws.merge_cells("A19:J19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv | SIY-PRO", yazi=GRİ, boyut=9)
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

    ws = wb["GIRDI"]
    for col in ("B",):
        ws.conditional_formatting.add(f"{col}9:{col}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(f"{col}9:{col}12", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("F32:F50", FormulaRule(formula=['F32="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F32:F50", FormulaRule(formula=['F32="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F6:F9", CellIsRule(operator="notEqual", formula=["0"], font=amber))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("YAPMA",E6))'], font=kirmizi))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("YAPILIR",E6))'], font=yesil))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",E6))'], font=amber))
    ws.conditional_formatting.add("B16:B20", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

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
    ws.conditional_formatting.add("G7:G72", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("F7:F72", FormulaRule(formula=['F7="SIY-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F72", FormulaRule(formula=['F7="SIY-004"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["ORNEK_VERI"]
    for harf in ("B",):
        ws.conditional_formatting.add(f"{harf}6:{harf}24", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B20", FormulaRule(formula=['ISNUMBER(SEARCH("YAPMA",B20))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B20", FormulaRule(formula=['ISNUMBER(SEARCH("YAPILIR",B20))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B20", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B20))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B16", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("B17", CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=yesil))

    ws = wb["FINANSMAN"]
    ws.conditional_formatting.add("E6:E7", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("C6:C7", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))

    ws = wb["VARSAYIMLAR"]
    ws.conditional_formatting.add("H6:H13", FormulaRule(formula=['H6="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("H6:H13", FormulaRule(formula=['H6="TAM"'], font=yesil, fill=y_fill))

    ws = wb["AYARLAR"]
    ws.conditional_formatting.add("H6:H50", FormulaRule(formula=['H6="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("H6:H50", FormulaRule(formula=['H6="TAM"'], font=yesil, fill=y_fill))


def ad_tanimlari(wb, tarife_aktif_satir: int):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "siy_netCapex", "MOTOR!$G$65")
    ad_ekle(wb, "siy_npv", "MOTOR!$G$66")
    ad_ekle(wb, "siy_irr", "MOTOR!$G$67")
    ad_ekle(wb, "siy_opexYillik", "MOTOR!$G$68")
    ad_ekle(wb, "siy_yillikKwh", "MOTOR!$G$69")
    ad_ekle(wb, "siy_kararMetni", "MOTOR!$G$41")
    ad_ekle(wb, "siy_kararGerekce", "MOTOR!$G$42")
    ad_ekle(wb, "siy_maddeAtifMetni", "MOTOR!$G$63")
    ad_ekle(wb, "siy_kanitRaporu", "MOTOR!$G$64")
    ad_ekle(wb, "siy_kuralYilMatris", "KURALLAR!$B$15")
    ad_ekle(wb, "siy_parmakIzi", "AYARLAR!$B$20")
    ad_ekle(wb, "siy_aktifTarife", f"AYARLAR!$K${tarife_aktif_satir}")
    ad_ekle(wb, "siy_tarifeOzeti", f"AYARLAR!$K${tarife_aktif_satir + 1}")
    ad_ekle(wb, "siy_finansmanKarsilastirma", "FINANSMAN!$B$11")
    ad_ekle(wb, "siy_kullanimGelir", "MOTOR!$G$33")
    ad_ekle(wb, "siy_capexOpex", "MOTOR!$G$31")

    ad_ekle(wb, "siy_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "siy_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "siy_duyarlilikSira", "SENARYO!$B$21")
    ad_ekle(wb, "siy_yorumDuyarlilik", "SENARYO!$B$22")
    ad_ekle(wb, "siy_kuralDegisimSenaryo", "SENARYO!$B$24")
    ad_ekle(wb, "siy_yorumKuralDegisim", "SENARYO!$B$25")
    ad_ekle(wb, "siy_tahminAralik", "SENARYO!$B$34")
    ad_ekle(wb, "siy_yorumTahmin", "SENARYO!$B$35")

    ad_ekle(wb, "siy_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "siy_vakaFark", "VAKALAR!$B$13")

    ayar_map = {
        "siy_tolerans": 8,
        "siy_azamiEsik": 9,
        "siy_senaryoIyi": 10,
        "siy_senaryoKotu": 11,
        "siy_esikArtis": 12,
        "siy_tahminAlt": 13,
        "siy_tahminUst": 14,
        "siy_tornadoKullanim": 15,
        "siy_tornadoPerakende": 16,
        "siy_tornadoToptan": 17,
        "siy_tornadoCapex": 18,
        "siy_yuzdelikOran": 19,
        "siy_girisBeklenen": 22,
        "siy_riskDusuk": 23,
        "siy_riskOrta": 24,
        "siy_riskYuksek": 25,
        "siy_yilTaban": 26,
        "siy_yuvarMax": 27,
        "siy_bazCarpan": 28,
        "siy_sifirCarpan": 29,
        "siy_siraBir": 30,
        "siy_siraIki": 31,
        "siy_siraUc": 32,
        "siy_dikkatCarpan": 33,
        "siy_npvDikkat": 34,
        "siy_tornadoWacc": 35,
        "vaka_soket": 36,
        "vaka_capexBirim": 37,
        "vaka_tesvik": 38,
        "vaka_seans": 39,
        "vaka_kwh": 40,
        "vaka_kullanim": 41,
        "vaka_perakende": 42,
        "vaka_toptan": 43,
        "vaka_ag": 44,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "siy_modulT1": (11, "C"), "siy_modulT1Yorum": (11, "D"),
        "siy_modulT2": (12, "C"), "siy_modulT2Yorum": (12, "D"),
        "siy_modulO1": (13, "C"), "siy_modulO1Yorum": (13, "D"),
        "siy_modulO6": (16, "C"), "siy_modulO6Yorum": (16, "D"),
        "siy_modulO8": (17, "C"), "siy_modulO8Yorum": (17, "D"),
        "siy_modulI2": (20, "C"), "siy_modulI2Yorum": (20, "D"),
    }
    for ad, (satir, kol) in modul_map.items():
        ad_ekle(wb, ad, f"PANO!${kol}${satir}")

    ad_ekle(wb, "ListeYillar", "LISTELER!$A$6:$A$8")
    ad_ekle(wb, "ListeSoketTip", "LISTELER!$C$6:$C$7")
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
    tarife_aktif = ayarlar(ws_ay)
    ws_kil = wb.create_sheet("KILAVUZ")
    kilavuz(ws_kil)

    ad_tanimlari(wb, tarife_aktif)
    kosullu_bicimlendirme(wb)
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    wb.calculation.fullCalcOnLoad = True

    kok = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    varsayilan = os.path.join(kok, "SarjIstasyonuYatirim.xlsx")
    hedef = cikti_yolu or varsayilan
    os.makedirs(os.path.dirname(os.path.abspath(hedef)) or ".", exist_ok=True)
    wb.save(hedef)

    hsh = hashlib.sha256()
    with open(hedef, "rb") as f:
        for b in iter(lambda: f.read(65536), b""):
            hsh.update(b)
    dig = hsh.hexdigest()

    wb2 = __import__("openpyxl").load_workbook(hedef)
    wb2["AYARLAR"]["B20"] = dig[:16]
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
    cikti = os.path.join(repo, "cikti", "SarjIstasyonuYatirim.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
