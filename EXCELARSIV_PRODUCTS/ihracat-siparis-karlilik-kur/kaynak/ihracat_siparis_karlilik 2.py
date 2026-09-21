#!/usr/bin/env python3
"""
İhracat Sipariş Kârlılık & Kur Riski Komutası — üretim betiği (A3+A6 / manda v6).
Sipariş durum makinesi + net kâr (TL) + hedge boşluğu + NPV/IRR karar.
Karar: YATIRIM YAPILIR / DİKKAT / YATIRIM YAPMA / VERİ YOK.
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
    h,
    sabitle,
    sayfa_hazirla,
    sayfa_koru,
    tablo_ekle,
    tablo_formullerini_hucrelere_yaz,
    yorum_ekle,
)
from ortak.akis_motoru import (
    gecersiz_gecis_formul,
    yetim_kayit_formul,
    zincir_kirik_formul,
)
from ortak.mevzuat_motoru import MOTOR_BASLIK, VAKA_BASLIK

URUN_AD = "İhracat Sipariş Kârlılık & Kur Riski Komutası"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 11)

DURUM_HARITASI = [
    {"durum_id": "TEKLIF", "durum_adi": "Teklif", "sira": 1,
     "izinli_sonraki_durumlar": "ONAY|IPTAL", "kategori": "acik"},
    {"durum_id": "ONAY", "durum_adi": "Onay", "sira": 2,
     "izinli_sonraki_durumlar": "URETIM|IPTAL", "kategori": "acik"},
    {"durum_id": "URETIM", "durum_adi": "Üretim", "sira": 3,
     "izinli_sonraki_durumlar": "SEVK|IPTAL", "kategori": "acik"},
    {"durum_id": "SEVK", "durum_adi": "Sevk", "sira": 4,
     "izinli_sonraki_durumlar": "TAHSILAT|IPTAL", "kategori": "acik"},
    {"durum_id": "TAHSILAT", "durum_adi": "Tahsilat", "sira": 5,
     "izinli_sonraki_durumlar": "KAPALI", "kategori": "acik"},
    {"durum_id": "KAPALI", "durum_adi": "Kapalı", "sira": 6,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "IPTAL", "durum_adi": "İptal", "sira": 7,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
]

# Demo: FOB 200k USD; rezervasyon 34 / tahsilat 32; hedge %60; CAPEX işletme sermayesi
DEMO = {
    "sirket_adi": "Örnek İhracat Sanayi A.Ş.",
    "donem": "2026-Q3",
    "hesapYili": 2026,
    "parabirimi": "USD",
    "fob_doviz": 200_000,
    "miktar": 2_000,
    "uretim_maliyet_doviz": 110_000,
    "navlun_cikis_doviz": 8_000,
    "komisyon_oran_giris": 0.03,
    "hedge_oran": 0.60,
    "tahsilat_kur": 32.00,
    "hedef_marj_giris": 0.18,
    "yillik_tekrar": 4,
    "capex_tl": 600_000,
    "wacc": 0.20,
    "npv_donem": 3,
    "kredi_orani": 0.35,
    "oz_kaynak_oran": 0.45,
}

AKIS_ORNEK = [
    {"id": "SIP-00001", "kart": "MUS-001", "onceki": "", "durum": "ONAY",
     "fob": 80_000, "maliyet": 45_000, "tutar": 80_000, "hedge": 0.70},
    {"id": "SIP-00002", "kart": "MUS-002", "onceki": "TEKLIF", "durum": "URETIM",
     "fob": 60_000, "maliyet": 38_000, "tutar": 60_000, "hedge": 0.50},
    {"id": "SIP-00003", "kart": "MUS-001", "onceki": "SEVK", "durum": "TAHSILAT",
     "fob": 40_000, "maliyet": 22_000, "tutar": 40_000, "hedge": 0.80},
    {"id": "SIP-00004", "kart": "", "onceki": "TEKLIF", "durum": "ONAY",
     "fob": 20_000, "maliyet": 12_000, "tutar": 20_000, "hedge": 0.40},
]

HDR, ILK, SON = 5, 6, 1005


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: Sipariş kârı ve kur riskini etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: net kâr ve karar bozulur | Kaynak: İhracat / hazine kuralları",
    )


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    h(ws, r, c, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
    ws.cell(r, c).protection = Protection(locked=False)
    if baslik:
        _kim(ws, f"{get_column_letter(c)}{r}", baslik, mesaj or baslik)


def kural_cek(kural_id: str) -> str:
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-isk_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(isk_yuvarMax,IFERROR(N(G12),0)))"


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
      "İhracat siparişinin net kârını (TL), hedge boşluğunu ve kur riskini hesaplayın; "
      "YATIRIM YAPILIR / YATIRIM YAPMA kararını üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Sipariş durum makinesi (TEKLIF→…→KAPALI) + yetim / geçersiz geçiş / ZİNCİR KIRIK",
        "FOB gelir − üretim − navlun − komisyon + kur farkı = net kâr (TL)",
        "Hedge oranı vs asgari hedge → açık kur pozisyonu boşluğu",
        "İşletme sermayesi CAPEX/OPEX + öz kaynak / kredi + NPV/IRR tornado",
        "Kötümser YATIRIM YAPMA yolu + KANIT_RAPORU",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "İhracat operasyon — sipariş satırını durum makinesinde yönetmek",
        "CFO / hazine — net kâr + kur riski ile sipariş kabul kararı",
        "SMMM / denetçi — imzalı kanıt raporunu dosyalamak",
    ], 14):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) HIZLI_BASLANGIC → 2) GIRDI + AKIS sarı hücreler → 3) PANO/KARAR → 4) KANIT_RAPORU yazdır.",
      yazi="333333", kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 21, 1, f"Sürüm {SURUM} | 2026 | ExcelArşiv | Lisans: Tek kullanıcı | Kod: ISK-PRO",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 72})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Üç adımda sonuç", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    for i, (baslik, metin) in enumerate([
        ("Adım 1 — Girdiler", "GIRDI'de FOB, maliyet, hedge ve kurları sarı hücrelere girin; AKIS'te sipariş satırlarını işleyin."),
        ("Adım 2 — Karar", "PANO ve KARAR'da net kâr, hedge boşluğu, NPV/IRR ve YATIRIM YAPILIR / YATIRIM YAPMA görün."),
        ("Adım 3 — Kanıt", "KANIT_RAPORU'nu yazdırın; KUYRUKLAR ve SENARYO ile doğrulayın."),
    ], 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, metin, kaydir=True, yazi="333333")
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=12)
    h(ws, 9, 1, "Formüller", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 1,
      "Gelir_TL = FOB×tahsilat_kur | Maliyet_TL = (üretim+navlun)×rezervasyon_kur | "
      "Kur farkı = FOB×(tahsilat−rezervasyon) | Net = Gelir − maliyet − komisyon + kur farkı",
      yazi="333333", kaydir=True)
    ws.merge_cells("A10:L10")
    h(ws, 12, 1, "Sarı = manuel giriş | Yeşil = formül çıktısı | Şifre 1234 formül alanlarını korur.",
      yazi=GRİ, kaydir=True)
    genislik(ws, {"A": 28, "B": 70})


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "İhracat sipariş girdileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Oranlar KURALLAR'dan; kur AYARLAR kur tablosundan.",
      yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Şirket / Sipariş", DEMO["sirket_adi"], None, "sirket_adi", "Unvan veya sipariş kodu."),
        (7, "Dönem", DEMO["donem"], None, "donem", "Rapor dönemi."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili", "2025, 2026 veya 2027."),
        (9, "Döviz Cinsi", DEMO["parabirimi"], None, "parabirimi", "USD veya EUR."),
        (10, "FOB [döviz]", DEMO["fob_doviz"], CATI, "fob_doviz", "İhracat FOB tutarı."),
        (11, "Miktar [adet]", DEMO["miktar"], CATI, "miktar", "Sipariş miktarı."),
        (12, "Üretim maliyeti [döviz]", DEMO["uretim_maliyet_doviz"], CATI, "uretim_maliyet",
         "Üretim / malzeme maliyeti döviz."),
        (13, "Çıkış navlunu [döviz]", DEMO["navlun_cikis_doviz"], CATI, "navlun_cikis",
         "İhracat navlun döviz."),
        (14, "Komisyon oranı (giriş)", DEMO["komisyon_oran_giris"], YÜZDE, "komisyon_oran",
         "Satış komisyonu; kural ayna ile doğrulanır."),
        (15, "Hedge oranı", DEMO["hedge_oran"], YÜZDE, "hedge_oran", "Kapsanan döviz oranı."),
        (16, "Tahsilat kuru", DEMO["tahsilat_kur"], "0.00", "tahsilat_kur", "Tahsilatta beklenen kur."),
        (17, "Hedef marj (giriş)", DEMO["hedef_marj_giris"], YÜZDE, "hedef_marj", "Hedef net marj."),
        (18, "Yıllık tekrar [adet]", DEMO["yillik_tekrar"], CATI, "yillik_tekrar", "Yıllık sipariş adedi."),
        (19, "CAPEX / işletme sermayesi [₺]", DEMO["capex_tl"], TL, "capex_tl", "Sipariş işletme sermayesi."),
        (20, "WACC [%]", DEMO["wacc"], YÜZDE, "wacc", "İskonto oranı."),
        (21, "NPV dönem [yıl]", DEMO["npv_donem"], CATI, "npv_donem", "Nakit akışı yılı."),
        (22, "Kredi faiz oranı [%]", DEMO["kredi_orani"], YÜZDE, "kredi_orani", "Kredi senaryosu faiz."),
        (23, "Öz kaynak oranı [%]", DEMO["oz_kaynak_oran"], YÜZDE, "oz_kaynak_oran", "Öz kaynak payı."),
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
        elif _ad == "tahsilat_kur":
            birim = "kur"
        elif _ad in ("miktar", "yillik_tekrar", "fob_doviz", "uretim_maliyet", "navlun_cikis"):
            birim = "adet/döviz"
        else:
            birim = "metin"
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim, yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 25, 1, "Aktif kur / rezervasyon (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=IFERROR(isk_aktifKur,0)", sayi="0.0000", yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 26, 1, "Asgari hedge (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 2, kural_cek("ISK-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 27, 1, "Net kâr TL (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 2, "=IFERROR(isk_netKarTl,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 28, 1, "Giriş doluluk", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 28, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B23)=0,isk_girisBeklenen=0),0,ROUND(COUNTA(B6:B23)/isk_girisBeklenen*100,0)),0)',
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
        ("parabirimi", DEMO["parabirimi"], "kod", "Evet", "GIRDI"),
        ("fob_doviz", DEMO["fob_doviz"], "döviz", "Evet", "GIRDI"),
        ("miktar", DEMO["miktar"], "adet", "Evet", "GIRDI"),
        ("uretim_maliyet", DEMO["uretim_maliyet_doviz"], "döviz", "Evet", "GIRDI"),
        ("navlun_cikis", DEMO["navlun_cikis_doviz"], "döviz", "Evet", "GIRDI"),
        ("komisyon_oran", DEMO["komisyon_oran_giris"], "oran", "Evet", "GIRDI"),
        ("hedge_oran", DEMO["hedge_oran"], "oran", "Evet", "GIRDI"),
        ("tahsilat_kur", DEMO["tahsilat_kur"], "kur", "Evet", "GIRDI"),
        ("hedef_marj", DEMO["hedef_marj_giris"], "oran", "Evet", "GIRDI"),
        ("yillik_tekrar", DEMO["yillik_tekrar"], "adet", "Evet", "GIRDI"),
        ("capex_tl", DEMO["capex_tl"], "TL", "Evet", "GIRDI"),
        ("wacc", DEMO["wacc"], "oran", "Evet", "GIRDI"),
        ("npv_donem", DEMO["npv_donem"], "yıl", "Evet", "GIRDI"),
        ("kredi_orani", DEMO["kredi_orani"], "oran", "Evet", "GIRDI"),
        ("oz_kaynak_oran", DEMO["oz_kaynak_oran"], "oran", "Evet", "GIRDI"),
    ]
    tl_a = {"capex_tl"}
    oran_a = {"komisyon_oran", "hedge_oran", "hedef_marj", "wacc", "kredi_orani", "oz_kaynak_oran"}
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 32):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in tl_a:
            ws.cell(i, 2).number_format = TL
        elif a in oran_a:
            ws.cell(i, 2).number_format = YÜZDE
        elif a == "tahsilat_kur":
            ws.cell(i, 2).number_format = "0.00"
        elif a in ("hesapYili", "miktar", "yillik_tekrar", "npv_donem", "fob_doviz",
                   "uretim_maliyet", "navlun_cikis"):
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
    dogrulama(ws, "list", "ListeDoviz", "B9",
              baslik="Döviz", mesaj="USD veya EUR seçin.",
              hata_baslik="Geçersiz", hata_mesaj="USD veya EUR.", bos=False)
    for aralik, baslik in [
        ("B10", "FOB"), ("B11", "Miktar"), ("B12", "Üretim"), ("B13", "Navlun"),
        ("B16", "Tahsilat kur"), ("B18", "Yıllık tekrar"), ("B19", "CAPEX"),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj="0 veya üzeri.",
                  hata_baslik=baslik, hata_mesaj="Negatif olamaz.",
                  isaret="greaterThanOrEqual", f2="1E12")
    genislik(ws, {"A": 36, "B": 28, "C": 14})


def kartlar(ws):
    sayfa_hazirla(ws, "KARTLAR", "2E75B6", URUN_AD, son_kolon=20)
    alt_bant(ws, 3, "Alıcı / müşteri kartları — AKIS kaynak kimliği", son_kolon=12)
    sutunlar = ["KartId", "Unvan", "Ulke", "ParaBirimi", "RiskNotu", "Aktif", "Kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    rows = [
        ("MUS-001", "Nordic Trade GmbH", "DE", "EUR", "Düşük", "Evet"),
        ("MUS-002", "Atlantic Buyers LLC", "US", "USD", "Orta", "Evet"),
        ("MUS-003", "Gulf Export Co", "AE", "USD", "Yüksek", "Evet"),
    ]
    for i, row in enumerate(rows, 6):
        for k, v in enumerate(row, 1):
            _sari(ws, i, k, v, baslik=sutunlar[k - 1], mesaj="Kart alanı")
    for r in range(9, 1006):
        for c in range(1, 7):
            _sari(ws, r, c, None, baslik=sutunlar[c - 1], mesaj="Manuel kart girişi")
    formuller = {
        "Kontrol": (
            '=IF(tblKartlar[[#This Row],[KartId]]="","",'
            'IF(tblKartlar[[#This Row],[Unvan]]="","EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblKartlar", "A5:G1005", sutunlar, formuller)
    dogrulama(ws, "list", "ListeDoviz", "D6:D1005",
              baslik="Para", mesaj="USD/EUR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeEvetHayir", "F6:F1005",
              baslik="Aktif", mesaj="Evet/Hayır",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    genislik(ws, {"A": 14, "B": 28, "C": 10, "D": 12, "E": 12, "F": 10})
    sabitle(ws, "A6")


def akis(ws):
    sayfa_hazirla(ws, "AKIS", "548235", URUN_AD, son_kolon=24)
    alt_bant(ws, 3,
             "Sipariş satırları — durum makinesi + yetim + geçersiz geçiş + zincir",
             son_kolon=20)
    sutunlar = [
        "IslemId", "KaynakKartId", "OncekiDurum", "Durum",
        "FobDoviz", "MaliyetDoviz", "Tutar", "HedgeOran",
        "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)

    yetim_f = yetim_kayit_formul(
        "tblAkis[[#This Row],[KaynakKartId]]", "tblKartlar[KartId]").lstrip("=")
    gecis_birlesik = (
        'tblAkis[[#This Row],[OncekiDurum]]&"|"&tblAkis[[#This Row],[Durum]]'
    )
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    zincir_f = zincir_kirik_formul(
        "tblAkis[[#This Row],[FobDoviz]]",
        "tblAkis[[#This Row],[Tutar]]",
        "zincirTolerans",
    ).lstrip("=")

    form = {
        "YetimBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{yetim_f})',
        "GecisUyarisi": (
            f'IF(OR(tblAkis[[#This Row],[IslemId]]="",'
            f'tblAkis[[#This Row],[OncekiDurum]]=""),"",{gecis_f})'
        ),
        "ZincirBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{zincir_f})',
        "KayitDolu": 'IF(tblAkis[[#This Row],[IslemId]]="","",1)',
    }

    for i, s in enumerate(AKIS_ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="İşlem Kimliği", mesaj="SIP-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Kart", mesaj="KARTLAR'daki KartId")
        _sari(ws, r, 3, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 4, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 5, s["fob"], sayi=CATI, baslik="FOB", mesaj="FOB döviz")
        _sari(ws, r, 6, s["maliyet"], sayi=CATI, baslik="Maliyet", mesaj="Maliyet döviz")
        _sari(ws, r, 7, s["tutar"], sayi=CATI, baslik="Tutar", mesaj="Açık tutar döviz")
        _sari(ws, r, 8, s["hedge"], sayi=YÜZDE, baslik="Hedge", mesaj="Hedge oranı")

    for r in range(ILK + len(AKIS_ORNEK), SON + 1):
        for c in range(1, 9):
            fmt = YÜZDE if c == 8 else (CATI if c in (5, 6, 7) else None)
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblAkis", f"A{HDR}:L{SON}", sutunlar, formuller=form)

    dogrulama(ws, "list", "ListeDurumlar", f"C{ILK}:C{SON}",
              baslik="Önceki Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"D{ILK}:D{SON}",
              baslik="Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    for c, ad in [(5, "FOB"), (6, "Maliyet"), (7, "Tutar")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 13)})


def kuyruklar(ws):
    sayfa_hazirla(ws, "KUYRUKLAR", "C65911", URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Durum bazlı canlı kuyruk — kapalı kayıtlar düşer", son_kolon=10)
    h(ws, 5, 1, "Kuyruk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Tutar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    for i, d in enumerate(["TEKLIF", "ONAY", "URETIM", "SEVK", "TAHSILAT"], 6):
        h(ws, i, 1, f"Kuyruk {d}")
        h(ws, i, 2,
          f'=COUNTIFS(tblAkis[Durum],"{d}",tblAkis[Tutar],">0")', sayi=CATI, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3,
          f'=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"{d}")', sayi=CATI, yazi=DIS_REF_YEŞIL)

    h(ws, 12, 1, "Açık toplam (kapalı düşer)")
    h(ws, 12, 2,
      '=COUNTIFS(tblAkis[Durum],"<>KAPALI",tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"<>KAPALI")-SUMIFS(tblAkis[Tutar],tblAkis[Durum],"IPTAL")',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "Kapalı / iptal adet")
    h(ws, 13, 2,
      '=COUNTIF(tblAkis[Durum],"KAPALI")+COUNTIF(tblAkis[Durum],"IPTAL")',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "Yetim adet")
    h(ws, 14, 2, '=COUNTIF(tblAkis[YetimBayrak],"YETİM")', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Geçersiz geçiş")
    h(ws, 15, 2, '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "ZİNCİR KIRIK")
    h(ws, 16, 2, '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Kuyruk özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 2,
      '=_xlfn.TEXTJOIN(" | ",TRUE,"Açık",B12,"Yetim",B14,"Geçiş",B15,"Zincir",B16)',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 20, 1, "KuyrukDemo", kalin=True, yazi=GRİ)
    for i, (ad, v) in enumerate([("ONAY", 1), ("URETIM", 1), ("TAHSILAT", 1), ("TEKLIF", 0)], 21):
        h(ws, i, 1, ad)
        h(ws, i, 2, v, sayi=CATI, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Açık Sipariş Kuyruğu"
    g.add_data(Reference(ws, min_col=2, min_row=20, max_row=24), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=21, max_row=24))
    g.height, g.width = 7, 12
    ws.add_chart(g, "E6")
    genislik(ws, {"A": 28, "B": 12, "C": 14})


def varsayimlar(ws):
    sayfa_hazirla(ws, "VARSAYIMLAR", "7030A0", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Varsayımlar — kaynak + güven (F02)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = ["varsayim_id", "ad", "deger", "birim", "kaynak", "guven", "tornado", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    rows = [
        ("V01", "Rezervasyon kuru", 34.0, "kur", "TCMB / kur tablosu", "sert", "Evet"),
        ("V02", "Tahsilat kuru", DEMO["tahsilat_kur"], "kur", "Hazine tahmini", "yumuşak", "Evet"),
        ("V03", "Hedge oranı", DEMO["hedge_oran"], "oran", "Hazine politikası", "sert", "Evet"),
        ("V04", "Asgari hedge", 0.80, "oran", "ISK-001", "sert", "Hayır"),
        ("V05", "Komisyon", DEMO["komisyon_oran_giris"], "oran", "ISK-004", "sert", "Hayır"),
        ("V06", "Kur şoku", 0.10, "oran", "ISK-003", "yumuşak", "Evet"),
        ("V07", "WACC", DEMO["wacc"], "oran", "CFO varsayımı", "yumuşak", "Hayır"),
        ("V08", "İşletme sermayesi CAPEX", DEMO["capex_tl"], "TL", "Nakit planı", "yumuşak", "Hayır"),
    ]
    for i, row in enumerate(rows, 6):
        for k, v in enumerate(row, 1):
            sayi = None
            if k == 3:
                if row[3] == "oran":
                    sayi = YÜZDE
                elif row[3] == "TL":
                    sayi = TL
                elif row[3] == "kur":
                    sayi = "0.00"
            h(ws, i, k, v, sayi=sayi, zemin=GIRIS_SARI if k in (3, 5, 6, 7) else None, yazi="333333")
            if k in (3, 5, 6, 7):
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", row[1], "Varsayım alanı")
    formuller = {
        "kontrol": (
            '=IF(tblVarsayimlar[[#This Row],[varsayim_id]]="","",'
            'IF(OR(tblVarsayimlar[[#This Row],[kaynak]]="",'
            'tblVarsayimlar[[#This Row],[guven]]=""),"EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblVarsayimlar", "A5:H13", sutunlar, formuller)
    genislik(ws, {"A": 42, "B": 28, "C": 14, "D": 10, "E": 22, "F": 12, "G": 10, "H": 10})


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kural motoru — ISK-001…005 (yıl yan yana)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = [
        "kural_id", "madde", "kural_adi", "kosul_metni", "parametre_adi",
        "deger_2025", "deger_2026", "deger_2027", "yururluk", "kaynak",
    ]
    baslik_satiri(ws, 5, sutunlar)
    data = [
        ("ISK-001", "Hazine politikası", "Asgari hedge oranı",
         "Açık döviz pozisyonu için asgari kapsama", "asgari_hedge",
         0.75, 0.80, 0.80, "01.01.2025", "İç hazine politikası"),
        ("ISK-002", "Satış politikası", "Hedef ihracat marjı",
         "Net kâr / gelir alt sınırı", "hedef_marj",
         0.18, 0.18, 0.18, "01.01.2025", "CFO marj politikası"),
        ("ISK-003", "Kur stres", "Kur şoku oranı",
         "Kötümser kur düşüşü", "kur_soku",
         0.08, 0.10, 0.12, "01.01.2025", "Hazine stres testi"),
        ("ISK-004", "Komisyon", "Satış komisyonu oranı",
         "FOB üzerinden komisyon", "komisyon",
         0.025, 0.03, 0.03, "01.01.2025", "Satış sözleşmesi"),
        ("ISK-005", "Yuvarlama", "Yuvarlama basamağı",
         "TL yuvarlama ondalığı", "yuvarlama",
         0, 0, 0, "01.01.2025", "Ürün mimarisi"),
    ]
    for i, row in enumerate(data, 6):
        for k, v in enumerate(row, 1):
            sayi = YÜZDE if k in (6, 7, 8) and row[0] != "ISK-005" else (CATI if k in (6, 7, 8) else None)
            h(ws, i, k, v, sayi=sayi, zemin=GIRIS_SARI if k in (6, 7, 8) else None, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", row[0], "Yıllık kural değeri")
    sutunlar_k = [*sutunlar, "kontrol"]
    baslik_satiri(ws, 5, sutunlar_k)
    # yeniden baslik — kontrol kolonu için
    h(ws, 5, 11, "kontrol", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    formuller = {
        "kontrol": (
            '=IF(tblKurallar[[#This Row],[kural_id]]="","",'
            'IF(tblKurallar[[#This Row],[kaynak]]="","EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A5:K10", sutunlar_k, formuller)
    h(ws, 12, 1, "Belirsizlik: hedge etkinlik yorumu — kullanıcı kendi banka teyidini girmelidir.",
      yazi=GRİ, kaydir=True)
    h(ws, 14, 1, "Yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"ISK-001",'
      'IFERROR(INDEX(tblKurallar[deger_2026],MATCH("ISK-001",tblKurallar[kural_id],0)),0),'
      '"ISK-002",'
      'IFERROR(INDEX(tblKurallar[deger_2026],MATCH("ISK-002",tblKurallar[kural_id],0)),0)),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B14:H14")
    h(ws, 17, 1, "Kural yıl matrisi (ad tanımı)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 2, "=IFERROR(B14,\"-\")", yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 44, "B": 18, "C": 24, "D": 36, "E": 16, "F": 12, "G": 12, "H": 12,
                  "I": 12, "J": 22, "K": 10})


def motor(ws):
    sayfa_hazirla(ws, "MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Hesap zinciri — sipariş kârı + kur riski + NPV", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar tblKurallar'dan; kur tblKur'dan çekilir.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")
    baslik_satiri(ws, 6, [*MOTOR_BASLIK, "deger"])

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif kural yılı", "ISK-005")
    ekle("Asgari hedge", km("ISK-001"), "oran", "Asgari hedge oranı", "ISK-001")
    ekle("Hedef marj kural", km("ISK-002"), "oran", "Hedef net marj", "ISK-002")
    ekle("Kur şoku oranı", km("ISK-003"), "oran", "Kötümser kur şoku", "ISK-003")
    ekle("Komisyon kural", km("ISK-004"), "oran", "Komisyon oranı", "ISK-004")
    ekle("Yuvarlama ondalık", km("ISK-005"), "adet", "Yuvarlama basamağı", "ISK-005")
    ekle("Aktif / rezervasyon kur", "=IFERROR(isk_aktifKur,0)", "kur", "Kur tablosu aktif kur", "ISK-005")
    ekle("FOB döviz", "=GIRDI!B10", "döviz", "FOB tutarı", "ISK-002")
    ekle("Miktar", "=GIRDI!B11", "adet", "Sipariş miktarı", "ISK-005")
    ekle("Üretim maliyet döviz", "=GIRDI!B12", "döviz", "Üretim maliyeti", "ISK-002")
    ekle("Navlun çıkış döviz", "=GIRDI!B13", "döviz", "Çıkış navlunu", "ISK-002")
    ekle("Komisyon giriş", "=GIRDI!B14", "oran", "Komisyon (giriş)", "ISK-004")
    ekle("Hedge oranı", "=GIRDI!B15", "oran", "Kapsanan oran", "ISK-001")
    ekle("Tahsilat kuru", "=GIRDI!B16", "kur", "Tahsilat kuru", "ISK-003")
    ekle("Hedef marj giriş", "=GIRDI!B17", "oran", "Hedef marj (giriş)", "ISK-002")
    ekle("Yıllık tekrar", "=GIRDI!B18", "adet", "Yıllık sipariş adedi", "ISK-005")
    ekle("CAPEX / YATIRIM", "=GIRDI!B19", "TL", "İşletme sermayesi", "ISK-005")
    ekle("WACC", "=GIRDI!B20", "oran", "İskonto oranı", "ISK-005")
    ekle("NPV dönem", "=GIRDI!B21", "adet", "Nakit akışı yılı", "ISK-005")
    ekle("Kredi faiz", "=GIRDI!B22", "oran", "Kredi senaryosu", "ISK-005")
    ekle("Öz kaynak oranı", "=GIRDI!B23", "oran", "Öz kaynak payı", "ISK-005")
    # G28 = rezervasyon kur (adım 7 → row 13 → G13? Let's count:
    # row = 6+i, i starts 1 → G7 = yıl, G8=hedge min, ... G13 = aktif kur (i=7)
    # G14=FOB, G15=miktar, G16=uretim, G17=navlun, G18=komisyon, G19=hedge, G20=tahsilat kur
    ekle("Gelir TL (tahsilat)", f"=ROUND(G14*G20,{_YV})", "TL", "FOB×tahsilat kur", "ISK-002")
    ekle("Maliyet TL (rezervasyon)", f"=ROUND((G16+G17)*G13,{_YV})", "TL",
         "(üretim+navlun)×rezervasyon", "ISK-002")
    ekle("Komisyon TL", f"=ROUND(G14*G13*G18,{_YV})", "TL", "FOB×rezervasyon×komisyon", "ISK-004")
    ekle("Kur farkı TL", f"=ROUND(G14*(G20-G13),{_YV})", "TL", "FOB×(tahsilat−rezervasyon)", "ISK-003")
    ekle("Net kâr TL", "=G28-G29-G30+G31", "TL", "Gelir−maliyet−komisyon+kur farkı", "ISK-002")
    ekle("Net marj", "=IF(G28=0,0,G32/G28)", "oran", "Net kâr / gelir", "ISK-002")
    ekle("Açık pozisyon döviz", "=G14*(isk_bazCarpan-G19)", "döviz", "Hedge edilmemiş FOB", "ISK-001")
    ekle("Hedge boşluğu", "=MAX(0,G8-G19)", "oran", "Asgari−gerçek hedge", "ISK-001")
    ekle("Kur riski TL", f"=ROUND(G34*G13*G10,{_YV})", "TL", "Açık×kur×şok", "ISK-003")
    ekle("OPEX / ISLETME yıllık", "=ABS(G29)*G22", "TL", "Maliyet×yıllık tekrar", "ISK-005")
    ekle("Yıllık net katkı", "=G32*G22", "TL", "Net kâr × yıllık tekrar", "ISK-002")
    ekle("NPV (baz)",
         f'=IF(OR(G25<=0,G24=""),0,IF(G24=0,ROUND(G38*G25-G23,{_YV}),'
         f'ROUND(G38*(1-POWER(1+G24,-G25))/G24-G23,{_YV})))',
         "TL", "NPV nakit − CAPEX", "ISK-005")
    ekle("IRR yaklaşımı",
         '=IF(OR(G23<=0,G38<=0,G25<=0),0,ROUND(G38/G23-isk_bazCarpan/G25,4))',
         "oran", "Kabaca IRR göstergesi", "ISK-005")
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(OR(G14<=0,G13<=0,G20<=0),"YAPMA",'
         'IF(OR(G39<0,G35>isk_hedgeDikkat,G33<G9*isk_dikkatCarpan),"YAPMA",'
         'IF(OR(G33<G9,G39<isk_npvDikkat,G35>0),"DIKKAT","YAPILIR"))))',
         "kod", "Karar kodu", "ISK-002")
    ekle("Karar metni",
         '=IF(G41="VERI_YOK","VERİ YOK",'
         'IF(G41="YAPILIR","YATIRIM YAPILIR",'
         'IF(G41="DIKKAT","DİKKAT","YATIRIM YAPMA")))',
         "metin", "Kullanıcıya karar", "ISK-002")
    ekle("Gerekçe",
         '=IF(G41="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G41="YAPILIR","NPV pozitif, marj hedef üstü, hedge yeterli; sipariş kabul uygun.",'
         'IF(G41="DIKKAT","Marj, hedge boşluğu veya NPV dikkat bandında; kuru gözden geçirin.",'
         '"NPV negatif, marj düşük veya kur riski yüksek; YATIRIM YAPMA.")))',
         "metin", "Gerekçe cümlesi", "ISK-002")
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B23)=0,isk_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B23)/isk_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "ISK-005")
    ekle("Risk skoru",
         '=IF(G41="VERI_YOK",0,IF(G41="YAPILIR",isk_riskDusuk,'
         'IF(G41="DIKKAT",isk_riskOrta,isk_riskYuksek)))',
         "puan", "Sipariş riski", "ISK-005")
    ekle("Senaryo iyimser net", f"=ROUND(G32*isk_senaryoIyi,{_YV})", "TL", "İyimser net kâr", "ISK-005")
    ekle("Senaryo kötümser net", f"=ROUND(G32*isk_senaryoKotu,{_YV})", "TL", "Kötümser net kâr", "ISK-005")
    ekle("İyimser NPV",
         f'=IF(G24=0,ROUND(G46*G22*G25-G23,{_YV}),'
         f'ROUND(G46*G22*(1-POWER(1+G24,-G25))/G24-G23,{_YV}))',
         "TL", "İyimser NPV", "ISK-005")
    ekle("Kötümser NPV",
         f'=IF(G24=0,ROUND(G47*G22*G25-G23,{_YV}),'
         f'ROUND(G47*G22*(1-POWER(1+G24,-G25))/G24-G23,{_YV}))',
         "TL", "Kötümser NPV", "ISK-005")
    ekle("Kötümser karar",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(G49<0,"YATIRIM YAPMA","DİKKAT"))',
         "metin", "Kötümser YATIRIM YAPMA yolu", "ISK-002")
    ekle("M-SEN kur şoku+", "=G10+isk_esikArtis", "oran", "Kur şoku + delta", "ISK-003")
    ekle("M-SEN kur riski", f"=ROUND(G34*G13*G51,{_YV})", "TL", "Sıkı kur riski", "ISK-003")
    ekle("M-SEN bayrak",
         '=IF(AND(G33>=G9,G39>=0,G35<=0),1,0)',
         "bayrak", "Sıkı eşikte uygun mu", "ISK-001")
    ekle("Tornado: kur etkisi",
         f"=ABS(ROUND(G14*G13*isk_tornadoKur-G14*G13,{_YV}))",
         "TL", "Kur ± etki", "ISK-003")
    ekle("Tornado: hedge etkisi",
         f"=ABS(ROUND(G14*G13*(G8-G19),{_YV}))",
         "TL", "Hedge boşluğu etki", "ISK-001")
    ekle("Tornado: maliyet etkisi",
         f"=ABS(ROUND((G16+G17)*G13*isk_tornadoMaliyet,{_YV}))",
         "TL", "Maliyet ± etki", "ISK-002")
    ekle("Kur kayıp yoğunluğu", "=IF(G28=0,0,G36/G28)", "oran", "Kur riski / gelir", "ISK-003")
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"ISK-001→hedge","ISK-002→marj","ISK-003→kur şoku","ISK-004→komisyon")',
         "metin", "Kanıt atıfları", "ISK-001")
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Net=",TEXT(G32,"₺ #,##0"),"NPV=",TEXT(G39,"₺ #,##0"),'
         '"IRR=",TEXT(G40,"0.0%"),"HedgeBoş=",TEXT(G35,"0.0%"),"Karar=",G42)',
         "metin", "Rapor özeti", "ISK-005")
    ekle("Net kâr (pano)", "=G32", "TL", "Net ayna", "ISK-002")
    ekle("NPV (pano)", "=G39", "TL", "NPV ayna", "ISK-005")
    ekle("IRR (pano)", "=G40", "oran", "IRR ayna", "ISK-005")
    ekle("Hedge boşluk (pano)", "=G35", "oran", "Hedge ayna", "ISK-001")
    ekle("Açık kuyruk adet",
         '=COUNTIFS(tblAkis[Durum],"<>KAPALI",tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")',
         "adet", "Açık sipariş", "ISK-005")
    ekle("Zincir kırık adet",
         '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")',
         "adet", "A04 bayrağı", "ISK-005")

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
        elif birim in ("oran", "kur"):
            ws.cell(r, 7).number_format = YÜZDE if birim == "oran" else "0.0000"
            ws.cell(r, 3).number_format = ws.cell(r, 7).number_format
        elif birim in ("puan", "yıl", "adet", "bayrak", "döviz"):
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
    h(ws, 6, 2, "=GIRDI!B23", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 3, "=GIRDI!B19*B6", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "=IFERROR(GIRDI!B23*isk_sifirCarpan,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 5, "=IFERROR(MOTOR!G39,0)", sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 7, 1, "KREDI", kalin=True, yazi="333333")
    h(ws, 7, 2, "=isk_bazCarpan-GIRDI!B23", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 3, "=GIRDI!B19*B7", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 4, "=GIRDI!B22", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 5, "=IFERROR(MOTOR!G39-C7*D7*GIRDI!B21,0)", sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 9, 1, "CAPEX / YATIRIM toplam", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 9, 2, "=GIRDI!B19", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "OPEX / ISLETME yıllık", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 2, "=IFERROR(MOTOR!G37,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Finansman karşılaştırma", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"OZ_KAYNAK NPV ",TEXT(E6,"₺ #,##0"),'
      '" · KREDI NPV ",TEXT(E7,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B11:H11")

    h(ws, 13, 1, "FinansDemo", kalin=True, yazi=GRİ)
    h(ws, 14, 1, "OZ_KAYNAK")
    h(ws, 14, 2, 270_000, sayi=TL, yazi=GRİ)
    h(ws, 15, 1, "KREDI")
    h(ws, 15, 2, 330_000, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Finansman Payları"
    g.add_data(Reference(ws, min_col=2, min_row=13, max_row=15), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=14, max_row=15))
    g.height, g.width = 7, 10
    ws.add_chart(g, "D13")
    genislik(ws, {"A": 48, "B": 14, "C": 16, "D": 14, "E": 16})


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve duyarlılık (tornado)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)

    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Çarpan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Net kâr", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "NPV", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(isk_senaryoIyi+N(GIRDI!B10)*isk_sifirCarpan,0)",
         "=IFERROR(MOTOR!G46,0)", "=IFERROR(MOTOR!G48,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D6<0,"YATIRIM YAPMA","YATIRIM YAPILIR"))'),
        ("Baz", "=IFERROR(isk_bazCarpan+N(GIRDI!B10)*isk_sifirCarpan,1)",
         "=IFERROR(isk_netKarTl,0)", "=IFERROR(isk_npv,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(isk_kararMetni,"-"))'),
        ("Kötümser", "=IFERROR(isk_senaryoKotu+N(GIRDI!B10)*isk_sifirCarpan,0)",
         "=IFERROR(MOTOR!G47,0)", "=IFERROR(MOTOR!G49,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(MOTOR!G50,"YATIRIM YAPMA"))'),
        ("M-SEN kur+", "=IFERROR(isk_bazCarpan+N(GIRDI!B10)*isk_sifirCarpan,1)",
         "=IFERROR(isk_netKarTl-MOTOR!G52,0)", "=IFERROR(MOTOR!G39,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(MOTOR!G53=1,"YATIRIM YAPILIR","DİKKAT"))'),
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
    h(ws, 16, 1, "Tornado kur etkisi", yazi="333333")
    h(ws, 16, 2, "=IFERROR(MOTOR!G54,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado hedge etkisi", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G55,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado maliyet etkisi", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G56,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    for r in (16, 17, 18):
        h(ws, r, 3,
          f'=IFERROR(IF(B{r}>=MAX(B16:B18),isk_siraBir,IF(B{r}=MEDIAN(B16:B18),isk_siraIki,isk_siraUc)),isk_siraUc)',
          sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(CONCATENATE("1)",INDEX(A16:A18,MATCH(isk_siraBir,C16:C18,0)),'
      '" 2)",INDEX(A16:A18,MATCH(isk_siraIki,C16:C18,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B16:B18),"₺ #,##0")," — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN kur değişimi bayrağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=MOTOR!G53", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Sıkı kur şoku (+",TEXT(isk_esikArtis,"0.0%"),") uygunluk ",TEXT(B22,"0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin alt", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=IFERROR(isk_netKarTl*isk_tahminAlt,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 3, "=IFERROR(isk_netKarTl,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 4, "=IFERROR(isk_netKarTl*isk_tahminUst,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Tahmin (AVERAGE / STDEV)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1, "Seri nokta", yazi=GRİ)
    for i, v in enumerate([1, 2, 3, 4], 28):
        h(ws, i, 1, v, sayi=CATI)
        h(ws, i, 2, f"=C{5+v}", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 32, 1, "Tahmin sonraki", yazi="333333")
    h(ws, 32, 2,
      '=IFERROR(IF(OR(COUNT(B28:B31)<2,STDEV(B28:B31)=0),AVERAGE(B28:B31),'
      'AVERAGE(B28:B31)+STDEV(B28:B31)*0),0)',
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Tahmin yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Tahmin ",TEXT(B32,"₺ #,##0")," · Alt ",TEXT(B25,"₺ #,##0"),'
      '" · Üst ",TEXT(D25,"₺ #,##0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 5, 7, "NpvDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([1_800_000, 900_000, -250_000, 700_000], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([220_000, 160_000, 95_000], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)
    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([1_100_000, 1_250_000, 980_000, 1_250_000], 28):
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
    g2.add_data(Reference(ws, min_col=5, min_row=15, max_row=18), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=16, max_row=18))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "F18")

    g3 = LineChart()
    g3.title = "Net Kâr Trendi"
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
        (2, "Net kâr", "=IFERROR(isk_netKarTl,0)", TL),
        (3, "Hedge boş", "=IFERROR(isk_hedgeBosluk,0)", YÜZDE),
        (4, "Kur riski", "=IFERROR(MOTOR!G36,0)", TL),
        (5, "Marj", "=IFERROR(MOTOR!G33,0)", YÜZDE),
        (6, "Açık kuyruk", "=IFERROR(MOTOR!G64,0)", CATI),
        (7, "NPV", "=IFERROR(isk_npv,0)", TL),
        (8, "IRR", "=IFERROR(isk_irr,0)", YÜZDE),
        (9, "Zincir", "=IFERROR(MOTOR!G65,0)", CATI),
        (10, "Güven", "=IFERROR(MOTOR!G44,0)", CATI),
        (11, "Risk", "=IFERROR(MOTOR!G45,0)", CATI),
        (12, "M-SEN", "=IFERROR(isk_kuralDegisimSenaryo,0)", CATI),
        (13, "Tahmin", "=IFERROR(isk_tahminAralik,0)", TL),
        (14, "Vaka", '=IFERROR(isk_vakaDurum,"-")', None),
    ]
    for col, ad, formul, sayi in kpis:
        h(ws, 3, col, ad, kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
        h(ws, 4, col, formul, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "Karar özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=isk_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=isk_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 10, 1, "Kod", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 2, "Ad", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 3, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 4, "Yorum", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    moduller = [
        ("T1", "Sipariş net kâr (TL)",
         "=IFERROR(isk_netKarTl,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Net ",TEXT(C11,"₺ #,##0")," — marj ",TEXT(MOTOR!G33,"0.0%"))'),
        ("T2", "Kur etkisi / hedge boşluğu",
         "=IFERROR(isk_hedgeBosluk,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Boşluk ",TEXT(C12,"0.0%")," · kur riski ",TEXT(MOTOR!G36,"₺ #,##0"))'),
        ("O1", "Açık sipariş kuyruğu",
         "=IFERROR(MOTOR!G64,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Açık sipariş ",TEXT(C13,"0")," · zincir ",TEXT(MOTOR!G65,"0"))'),
        ("O2", "Duyarlılık / tornado",
         '=IFERROR(isk_duyarlilikSira,"-")',
         "=IFERROR(isk_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo motoru",
         "=IFERROR(isk_senaryoKarsilastirma,0)",
         "=IFERROR(isk_yorumSenaryo,\"-\")"),
        ("O8", "Veri kalite skoru",
         "=IFERROR(MOTOR!G44,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Giriş güven skoru ",TEXT(C16,"0"),"/100")'),
        ("M_SEN", "Kur şoku senaryosu",
         "=IFERROR(isk_kuralDegisimSenaryo,0)",
         "=IFERROR(isk_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin + aralık",
         "=IFERROR(isk_tahminAralik,0)",
         "=IFERROR(isk_yorumTahmin,\"-\")"),
        ("I2", "Kur kayıp P90",
         "=IFERROR(IF(COUNT(SENARYO!C6:C9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!C6:C9,isk_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 net ",TEXT(C19,"₺ #,##0"))'),
    ]
    for i, (kod, ad, deger, yorum) in enumerate(moduller, 11):
        h(ws, i, 1, kod, kalin=True, yazi="B08948")
        h(ws, i, 2, ad, yazi="333333")
        sayi = YÜZDE if kod == "T2" else (CATI if kod in ("O1", "O8", "M_SEN") else TL)
        if kod in ("O2",):
            sayi = None
        h(ws, i, 3, deger, sayi=sayi, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "PanoDemoKar", kalin=True, yazi=GRİ)
    for i, (ad, v) in enumerate([("Gelir", 0.55), ("Maliyet", 0.30), ("Kur/komisyon", 0.15)], 23):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, v, sayi=YÜZDE, yazi=GRİ)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Kâr Kırılımı"
    g1.add_data(Reference(ws, min_col=2, min_row=22, max_row=25), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=23, max_row=25))
    g1.height, g1.width = 7, 12
    ws.add_chart(g1, "F22")

    h(ws, 22, 4, "Etiket")
    h(ws, 22, 5, "PanoDemoNpv", kalin=True, yazi=GRİ)
    for i, (ad, v) in enumerate([("İyimser", 1_800_000), ("Baz", 900_000), ("Kötümser", -250_000)], 23):
        h(ws, i, 4, ad, yazi="333333")
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)
    g2 = BarChart()
    g2.type = "col"
    g2.title = "NPV Senaryo"
    g2.add_data(Reference(ws, min_col=5, min_row=22, max_row=25), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=23, max_row=25))
    g2.height, g2.width = 7, 12
    ws.add_chart(g2, "H22")

    h(ws, 27, 1, "Metrik")
    h(ws, 27, 2, "RiskDemo", kalin=True, yazi=GRİ)
    h(ws, 28, 1, "Güven")
    h(ws, 28, 2, 100, sayi=CATI, yazi=GRİ)
    h(ws, 29, 1, "Risk")
    h(ws, 29, 2, 40, sayi=CATI, yazi=GRİ)
    g3 = BarChart()
    g3.type = "col"
    g3.title = "Güven / Risk"
    g3.add_data(Reference(ws, min_col=2, min_row=27, max_row=29), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=28, max_row=29))
    g3.height, g3.width = 6, 10
    ws.add_chart(g3, "F28")

    h(ws, 27, 4, "Kalem")
    h(ws, 27, 5, "KurDemo", kalin=True, yazi=GRİ)
    h(ws, 28, 4, "Kur farkı")
    h(ws, 28, 5, -400_000, sayi=TL, yazi=GRİ)
    h(ws, 29, 4, "Kur riski")
    h(ws, 29, 5, 272_000, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "Kur Etki"
    g4.add_data(Reference(ws, min_col=5, min_row=27, max_row=29), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=4, min_row=28, max_row=29))
    g4.height, g4.width = 6, 10
    ws.add_chart(g4, "H28")

    genislik(ws, {"A": 22, "B": 36, "C": 16, "D": 55, "E": 14})


def karar(ws):
    sayfa_hazirla(ws, "KARAR", "B3261E", URUN_AD, son_kolon=30)
    baski_hazirla(ws, "A1:F28", f"{URUN_AD} · KARAR")
    h(ws, 3, 1, "Sipariş / yatırım karar kapısı", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Karar", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 5, 2, "=isk_kararMetni", boyut=16, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Gerekçe", kalin=True)
    h(ws, 6, 2, "=isk_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B6:H6")
    h(ws, 8, 1, "NPV [₺]", kalin=True)
    h(ws, 8, 2, "=isk_npv", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "IRR", kalin=True)
    h(ws, 9, 2, "=isk_irr", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "Net kâr [₺]", kalin=True)
    h(ws, 10, 2, "=isk_netKarTl", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Kötümser karar", kalin=True)
    h(ws, 11, 2, '=IFERROR(MOTOR!G50,"YATIRIM YAPMA")', yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "CAPEX / YATIRIM", kalin=True)
    h(ws, 13, 2, "=GIRDI!B19", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "OPEX / ISLETME yıllık", kalin=True)
    h(ws, 14, 2, "=IFERROR(MOTOR!G37,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Hedge boşluğu", kalin=True)
    h(ws, 15, 2, "=isk_hedgeBosluk", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1,
      "Boş girişte karar daima VERİ YOK olur. Kötümser NPV negatifse YATIRIM YAPMA üretilir.",
      yazi=GRİ, kaydir=True)
    ws.merge_cells("A17:H17")
    genislik(ws, {"A": 28, "B": 55})


def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Altın vakalar — ihracat kâr / kur", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:H3")
    baslik_satiri(ws, 5, VAKA_BASLIK)
    # vaka_fob=200000, vaka_rez=34, vaka_tah=32 → gelir=6.4M, maliyet=(110k+8k)*34=4.012M
    # komisyon=200k*34*0.03=204000, kur_farki=200k*(32-34)=-400000
    # net = 6400000-4012000-204000-400000 = 1_784_000
    vakalar_data = [
        ("V-001", "Gelir TL = FOB × tahsilat kur",
         "fob=vaka_fob; tah=vaka_tah",
         6_400_000,
         "=IFERROR(ROUND(vaka_fob*vaka_tah,0),0)",
         "=IFERROR(E6-D6,0)",
         '=IFERROR(IF(ABS(F6)<=isk_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Gelir TL"),
        ("V-002", "Maliyet TL = (üretim+navlun)×rezervasyon",
         "uretim+navlun; rez=vaka_rez",
         4_012_000,
         "=IFERROR(ROUND((vaka_uretim+vaka_navlun)*vaka_rez,0),0)",
         "=IFERROR(E7-D7,0)",
         '=IFERROR(IF(ABS(F7)<=isk_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Maliyet"),
        ("V-003", "Kur farkı = FOB×(tah−rez)",
         "tah-rez",
         -400_000,
         "=IFERROR(ROUND(vaka_fob*(vaka_tah-vaka_rez),0),0)",
         "=IFERROR(E8-D8,0)",
         '=IFERROR(IF(ABS(F8)<=isk_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Kur farkı"),
        ("V-004", "Net kâr = gelir−maliyet−komisyon+kur farkı",
         "komisyon=vaka_kom",
         1_784_000,
         "=IFERROR(ROUND(vaka_fob*vaka_tah-(vaka_uretim+vaka_navlun)*vaka_rez"
         "-vaka_fob*vaka_rez*vaka_kom+vaka_fob*(vaka_tah-vaka_rez),0),0)",
         "=IFERROR(E9-D9,0)",
         '=IFERROR(IF(ABS(F9)<=isk_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Net kâr"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v, yazi=DIS_REF_YEŞIL if k >= 5 else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = CATI

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
            'IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=isk_tolerans,"TUTARLI","KIRIK"),"KIRIK"))'
        ),
    }
    tablo_ekle(ws, "tblVakalar", "A5:H9", VAKA_BASLIK, formuller)

    h(ws, 12, 1, "Vaka durumu özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(IF(COUNTIF(G6:G9,"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "Toplam mutlak fark", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2, "=IFERROR(SUM(ABS(F6),ABS(F7),ABS(F8),ABS(F9)),0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 10, "B": 44, "C": 36, "D": 12, "E": 14, "F": 12, "G": 12, "H": 14})
    sabitle(ws, "A6")


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", "0F2742", URUN_AD, son_kolon=30)
    baski_hazirla(ws, "A1:F36", f"{URUN_AD} · {SURUM}")
    h(ws, 3, 1, "KANIT RAPORU — İhracat Sipariş Kârlılık & Kur Riski", kalin=True, boyut=14,
      yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "CFO / hazine / SMMM dosyası", yazi=GRİ)
    h(ws, 6, 1, "Şirket", kalin=True)
    h(ws, 6, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Dönem", kalin=True)
    h(ws, 7, 2, "=GIRDI!B7", yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Hesap yılı", kalin=True)
    h(ws, 8, 2, "=hesapYili", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Rapor tarihi", kalin=True)
    h(ws, 9, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "Aktif kur", kalin=True)
    h(ws, 10, 2, "=isk_aktifKur", sayi="0.0000", yazi=DIS_REF_YEŞIL)

    h(ws, 12, 1, "Net kâr [₺]", kalin=True, boyut=12)
    h(ws, 12, 2, "=isk_netKarTl", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 13, 1, "Hedge boşluğu", kalin=True)
    h(ws, 13, 2, "=isk_hedgeBosluk", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "Kur riski [₺]", kalin=True)
    h(ws, 14, 2, "=MOTOR!G36", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "NPV [₺]", kalin=True)
    h(ws, 15, 2, "=isk_npv", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "IRR", kalin=True)
    h(ws, 16, 2, "=isk_irr", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)

    h(ws, 18, 1, "Karar", kalin=True, boyut=12)
    h(ws, 18, 2, "=isk_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Gerekçe", kalin=True)
    h(ws, 19, 2, "=isk_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B19:H19")
    h(ws, 21, 1, "Madde / kural atıfı", kalin=True)
    h(ws, 21, 2, "=isk_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B21:H21")
    h(ws, 23, 1, "Kanıt özeti", kalin=True)
    h(ws, 23, 2, "=isk_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B23:H23")
    h(ws, 25, 1, "Parmak izi", kalin=True)
    h(ws, 25, 2, "=isk_parmakIzi", yazi=GRİ)

    h(ws, 28, 1, "Hazırlayan", kalin=True)
    h(ws, 28, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B28", "Hazırlayan", "Ad soyad imza alanı")
    h(ws, 29, 1, "Onaylayan", kalin=True)
    h(ws, 29, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B29", "Onaylayan", "CFO / hazine imza alanı")
    genislik(ws, {"A": 42, "B": 55})
    h(ws, 32, 1, "Bu çıktı karar destek amaçlıdır; bağlayıcı yatırım veya döviz tavsiyesi değildir.",
      yazi=GRİ, boyut=9, kaydir=True)


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Dosya sağlık kontrolleri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    kontroller_list = [
        (5, "Giriş dolu mu?",
         '=IFERROR(IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"BOŞ","DOLU"),"BOŞ")'),
        (6, "Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(isk_yilTaban),N(hesapYili)<=N(isk_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        (7, "Kur > 0?",
         '=IFERROR(IF(isk_aktifKur>0,"TEMİZ","HATALI"),"HATALI")'),
        (8, "FOB > 0?",
         '=IFERROR(IF(MOTOR!G14>0,"TEMİZ","HATALI"),"HATALI")'),
        (9, "Net kâr kontrol",
         '=IFERROR(IF(MOTOR!G32<-isk_azamiEsik,"NEGATİF","NORMAL"),"NORMAL")'),
        (10, "Vaka durumu",
         '=IFERROR(isk_vakaDurum,"KIRIK")'),
        (11, "Hedge boşluğu?",
         '=IF(IFERROR(MOTOR!G35,0)>0,"AŞIYOR","NORMAL")'),
    ]
    for satir, ad, formul in kontroller_list:
        h(ws, satir, 1, ad, yazi="333333")
        h(ws, satir, 2, formul, yazi=DIS_REF_YEŞIL)

    h(ws, 13, 1, "Sağlık skoru", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2,
      '=IFERROR(ROUND((COUNTIF(B5:B11,"TEMİZ")+COUNTIF(B5:B11,"DOLU")+COUNTIF(B5:B11,"NORMAL")'
      '+COUNTIF(B5:B11,"TUTARLI"))/7*100,0),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Net kâr", kalin=True)
    h(ws, 15, 2, "=IFERROR(isk_netKarTl,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "NPV", kalin=True)
    h(ws, 16, 2, "=IFERROR(isk_npv,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "IRR", kalin=True)
    h(ws, 17, 2, "=IFERROR(isk_irr,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Karar", kalin=True)
    h(ws, 18, 2, '=IFERROR(isk_kararMetni,"-")', yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 36, "B": 28})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "548235", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek ihracat sipariş verileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:L3")
    sutunlar = [
        "Senaryo", "FOB", "Miktar", "Üretim", "Navlun", "RezKur", "TahKur",
        "Hedge", "Komisyon", "CAPEX", "Yıl", "Net Ornek",
    ]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Net Ornek": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(ROUND(tblOrnek[[#This Row],[FOB]]*tblOrnek[[#This Row],[TahKur]]"
            "-(tblOrnek[[#This Row],[Üretim]]+tblOrnek[[#This Row],[Navlun]])"
            "*tblOrnek[[#This Row],[RezKur]]"
            "-tblOrnek[[#This Row],[FOB]]*tblOrnek[[#This Row],[RezKur]]"
            "*tblOrnek[[#This Row],[Komisyon]]"
            "+tblOrnek[[#This Row],[FOB]]*(tblOrnek[[#This Row],[TahKur]]"
            "-tblOrnek[[#This Row],[RezKur]]),0),0))"
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:L1004", sutunlar, formuller)
    rows = [
        ("Demo ana", 200_000, 2_000, 110_000, 8_000, 34, 32, 0.60, 0.03, 600_000, 2026),
        ("Güçlü hedge", 200_000, 2_000, 110_000, 8_000, 34, 33, 0.90, 0.03, 500_000, 2026),
        ("Kur şoku", 200_000, 2_000, 110_000, 8_000, 34, 28, 0.40, 0.03, 700_000, 2026),
        ("2025 geçiş", 150_000, 1_500, 85_000, 6_000, 32, 31, 0.70, 0.025, 450_000, 2025),
    ]
    for i, row in enumerate(rows, 6):
        for k, v in enumerate(row, 1):
            sayi = None
            if k in (2, 3, 4, 5, 11):
                sayi = CATI
            if k in (6, 7):
                sayi = "0.00"
            if k in (8, 9):
                sayi = YÜZDE
            if k == 10:
                sayi = TL
            h(ws, i, k, v, sayi=sayi, zemin=GIRIS_SARI if k <= 11 else None, yazi="333333")
            if k <= 11:
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", sutunlar[k - 1], "Örnek satır")
    h(ws, 22, 1, "OrnekNetDemo", kalin=True, yazi=GRİ)
    for i, v in enumerate([1_784_000, 2_200_000, 400_000, 1_100_000], 23):
        h(ws, i, 1, rows[i - 23][0])
        h(ws, i, 2, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Örnek Net Kâr"
    g.add_data(Reference(ws, min_col=2, min_row=22, max_row=26), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=23, max_row=26))
    g.height, g.width = 7, 12
    ws.add_chart(g, "D22")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 13)})
    genislik(ws, {"A": 16, "L": 12})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", "7030A0", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Doğrulama listeleri", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Yıl", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, y in enumerate([2025, 2026, 2027], 6):
        h(ws, i, 1, y, sayi=CATI)
    h(ws, 5, 3, "Döviz", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "USD")
    h(ws, 7, 3, "EUR")
    h(ws, 5, 5, "EvetHayır", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 5, "Evet")
    h(ws, 7, 5, "Hayır")
    h(ws, 5, 7, "Durum", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, d in enumerate(DURUM_HARITASI, 6):
        h(ws, i, 7, d["durum_id"])
    h(ws, 5, 9, "IzinliGecis", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    izinler = [
        f'{d["durum_id"]}|{nxt}'
        for d in DURUM_HARITASI
        for nxt in (d["izinli_sonraki_durumlar"] or "").split("|")
        if nxt
    ]
    for i, g in enumerate(izinler, 6):
        h(ws, i, 9, g)
    genislik(ws, {"A": 12, "C": 12, "E": 12, "G": 12, "I": 22})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "833C0C", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Parametreler (tek kaynak) + kur tablosu + durum haritası", kalin=True,
      boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = ["anahtar", "deger", "birim", "aciklama", "kaynak", "yururluk_tarihi",
                "dogrulama_tarihi", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    params = [
        ("hesapYili", DEMO["hesapYili"], "yıl", "Aktif hesap yılı", "Kullanıcı / GIRDI",
         "01.01.2025", "11.08.2026"),
        ("raporTarihi", RAPOR_TARIH, "tarih", "Rapor tarihi", "Üretim",
         "01.01.2025", "11.08.2026"),
        ("zincirTolerans", 1, "adet", "FOB/tutar zincir toleransı", "A04",
         "01.01.2025", "11.08.2026"),
        ("isk_tolerans", 1, "adet", "Vaka tutarlılık toleransı", "Uygulama notu",
         "01.01.2025", "11.08.2026"),
        ("isk_azamiEsik", 100_000_000_000, "TL", "Azami giriş eşiği", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("isk_senaryoIyi", 1.12, "çarpan", "İyimser net çarpanı", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("isk_senaryoKotu", 0.55, "çarpan", "Kötümser net çarpanı", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("isk_esikArtis", 0.02, "oran", "M-SEN kur şoku + delta", "Senaryo motoru",
         "01.01.2025", "11.08.2026"),
        ("isk_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("isk_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("isk_tornadoKur", 1.08, "çarpan", "Tornado kur şoku", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("isk_tornadoMaliyet", 0.10, "oran", "Tornado maliyet şoku", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("isk_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik",
         "01.01.2025", "11.08.2026"),
        ("isk_parmakIzi", "ISK-PRO-1.0.0", "metin", "Dosya parmak izi", "Üretim",
         "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv",
         "01.01.2025", "11.08.2026"),
        ("isk_girisBeklenen", 18, "adet", "Zorunlu giriş sayısı", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("isk_riskDusuk", 25, "puan", "YAPILIR risk skoru", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("isk_riskOrta", 55, "puan", "DİKKAT risk skoru", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("isk_riskYuksek", 85, "puan", "YAPMA risk skoru", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("isk_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("isk_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("isk_bazCarpan", 1, "çarpan", "Baz senaryo çarpanı", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("isk_sifirCarpan", 0, "çarpan", "Nötr çarpan", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("isk_siraBir", 1, "adet", "Tornado sıra 1", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("isk_siraIki", 2, "adet", "Tornado sıra 2", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("isk_siraUc", 3, "adet", "Tornado sıra 3", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("isk_dikkatCarpan", 0.70, "çarpan", "Marj dikkat çarpanı", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("isk_npvDikkat", 80_000, "TL", "NPV dikkat eşiği", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("isk_hedgeDikkat", 0.25, "oran", "Hedge boşluğu kritik eşik", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("isk_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi",
         "01.01.2025", "11.08.2026"),
        ("vaka_fob", 200_000, "döviz", "Vaka FOB", "Altın vaka",
         "01.01.2025", "11.08.2026"),
        ("vaka_rez", 34, "kur", "Vaka rezervasyon kur", "Altın vaka",
         "01.01.2025", "11.08.2026"),
        ("vaka_tah", 32, "kur", "Vaka tahsilat kur", "Altın vaka",
         "01.01.2025", "11.08.2026"),
        ("vaka_uretim", 110_000, "döviz", "Vaka üretim", "Altın vaka",
         "01.01.2025", "11.08.2026"),
        ("vaka_navlun", 8_000, "döviz", "Vaka navlun", "Altın vaka",
         "01.01.2025", "11.08.2026"),
        ("vaka_kom", 0.03, "oran", "Vaka komisyon", "Altın vaka",
         "01.01.2025", "11.08.2026"),
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
                elif row[2] in ("yıl", "adet", "puan", "döviz"):
                    sayi = CATI
                elif row[2] == "kur":
                    sayi = "0.00"
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k == 2 and sayi is None and isinstance(v, (int, float)) and not isinstance(v, bool):
                ws.cell(i, k).number_format = CATI
            if k == 2:
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"B{i}", row[0], row[3])
        h(ws, i, 8, f'=IF(A{i}="","",IF(B{i}="","EKSİK","TAM"))', yazi=DIS_REF_YEŞIL, hiza="center")

    son = 5 + len(params)

    # Durum haritası — H kolonundan (A=anahtar parametre taramasını kirletmez)
    h(ws, 6, 10, "Durum haritası (izinli sonraki / kategori)", kalin=True,
      boyut=12, yazi=KOYU_LACIVERT)
    dh_bas = ["durum_id", "durum_adi", "sira", "izinli_sonraki_durumlar", "kategori"]
    baslik_satiri(ws, 7, dh_bas, basla=10)
    for i, d in enumerate(DURUM_HARITASI):
        r = 8 + i
        h(ws, r, 10, d["durum_id"])
        h(ws, r, 11, d["durum_adi"])
        h(ws, r, 12, d["sira"], sayi=CATI)
        h(ws, r, 13, d["izinli_sonraki_durumlar"])
        h(ws, r, 14, d["kategori"])
    tablo_ekle(ws, "tblDurumHaritasi", f"J7:N{7+len(DURUM_HARITASI)}", dh_bas)

    # Kur tablosu — parametre satırlarının altına, J kolonundan
    kr0 = son + 3
    h(ws, kr0, 10, "Kur tablosu (tarihli / kaynaklı)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kur_sutun = ["parabirimi", "kur", "tarih", "kaynak", "aktif", "not", "secim"]
    baslik_satiri(ws, kr0 + 1, kur_sutun, basla=10)
    kur_rows = [
        ("USD", 34.00, date(2026, 8, 10), "TCMB alış", "Evet", "Aktif rezervasyon kur"),
        ("EUR", 37.20, date(2026, 8, 10), "TCMB alış", "Hayır", "Yedek"),
        ("USD", 32.50, date(2025, 12, 31), "TCMB yılsonu", "Hayır", "2025 referans"),
    ]
    kr_data = kr0 + 2
    for i, row in enumerate(kur_rows, kr_data):
        for k, v in enumerate(row, 1):
            sayi = TARİH if k == 3 else ("0.00" if k == 2 else None)
            h(ws, i, 9 + k, v, sayi=sayi, zemin=GIRIS_SARI if k <= 5 else None, yazi="333333")
            if k <= 5:
                ws.cell(i, 9 + k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(9+k)}{i}", kur_sutun[k - 1], "Kur tablosu")
    formuller_kur = {
        "secim": (
            '=IF(tblKur[[#This Row],[aktif]]="Evet",tblKur[[#This Row],[kur]],"")'
        ),
    }
    tablo_ekle(ws, "tblKur", f"J{kr0+1}:P{kr_data+2}", kur_sutun, formuller_kur)
    h(ws, kr_data + 4, 10, "Aktif kur özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, kr_data + 4, 11,
      f'=IFERROR(SUMIF(N{kr_data}:N{kr_data+2},"Evet",L{kr_data}:L{kr_data+2}),0)',
      sayi="0.0000", yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, kr_data + 5, 10, "Kur tablo not", kalin=True)
    h(ws, kr_data + 5, 11, "Aktif satırdan çekilir; sabit kur gömülmez.", yazi=GRİ)

    genislik(ws, {"A": 22, "B": 22, "C": 12, "D": 36, "E": 18, "J": 14, "K": 12, "L": 12,
                  "M": 14, "N": 12, "O": 24, "P": 10})
    return kr_data + 4  # aktif kur özet satırı


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "595959", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kullanım kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. GIRDI sarı hücrelerine sipariş FOB, maliyet, hedge ve kurları girin.",
        "2. KARTLAR'da alıcı kimliklerini tanımlayın; AKIS'te sipariş satırlarını durumla işleyin.",
        "3. KUYRUKLAR açık siparişleri gösterir; KAPALI/IPTAL kuyruktan düşer.",
        "4. MOTOR net kâr, hedge boşluğu, kur riski ve NPV/IRR üretir.",
        "5. FINANSMAN'da OZ_KAYNAK ve KREDI senaryolarını karşılaştırın.",
        "6. SENARYO iyimser/baz/kötümser + tornado duyarlılık gösterir.",
        "7. KARAR boş girişte VERİ YOK; kötümserde YATIRIM YAPMA üretebilir.",
        "8. KANIT_RAPORU'nu yazdırıp imzalayın. Şifre: 1234 (formül koruması).",
    ]
    for i, m in enumerate(maddeler, 5):
        h(ws, i, 1, m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)
    genislik(ws, {"A": 90})


def kosullu_bicimlendirme(wb):
    kirmizi = Font(color="B3261E", bold=True)
    yesil = Font(color="1F7A4D", bold=True)
    k_fill = PatternFill("solid", fgColor="FDE9E9")
    y_fill = PatternFill("solid", fgColor="E2EFDA")

    ws = wb["KARAR"]
    ws.conditional_formatting.add("B5", FormulaRule(
        formula=['ISNUMBER(SEARCH("YAPMA",B5))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B5", FormulaRule(
        formula=['ISNUMBER(SEARCH("YAPILIR",B5))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B5", FormulaRule(
        formula=['ISNUMBER(SEARCH("VERİ YOK",B5))'], font=Font(color="595959", bold=True)))
    ws.conditional_formatting.add("B5", FormulaRule(
        formula=['ISNUMBER(SEARCH("DİKKAT",B5))'], font=Font(color="B7791F", bold=True)))
    ws.conditional_formatting.add("B8", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("B8", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("B10", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("B11", FormulaRule(
        formula=['ISNUMBER(SEARCH("YAPMA",B11))'], font=kirmizi))
    ws.conditional_formatting.add("B15", CellIsRule(operator="greaterThan", formula=["0"], font=kirmizi))

    ws = wb["PANO"]
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("YAPMA",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("YAPILIR",B6))'], font=yesil, fill=y_fill))
    for col in "BCDEFGHIJKLMN":
        ws.conditional_formatting.add(f"{col}4", CellIsRule(
            operator="lessThan", formula=["0"], font=kirmizi))
    for r in range(11, 20):
        ws.conditional_formatting.add(f"C{r}", CellIsRule(
            operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(f"A{r}", FormulaRule(
            formula=[f'AND(A{r}<>"",TRUE)'], font=Font(color="B08948", bold=True)))

    ws = wb["MOTOR"]
    for r in range(7, 70):
        ws.conditional_formatting.add(f"G{r}", CellIsRule(
            operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(f"G{r}", CellIsRule(
            operator="greaterThan", formula=["0"], font=yesil))

    ws = wb["AKIS"]
    ws.conditional_formatting.add("I6:I1005", FormulaRule(
        formula=['I6="YETİM"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("J6:J1005", FormulaRule(
        formula=['J6="GEÇERSİZ GEÇİŞ"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("K6:K1005", FormulaRule(
        formula=['K6="ZİNCİR KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("D6:D1005", FormulaRule(
        formula=['OR(D6="KAPALI",D6="IPTAL")'], font=Font(color="595959")))
    ws.conditional_formatting.add("D6:D1005", FormulaRule(
        formula=['OR(D6="ONAY",D6="TAHSILAT")'], font=yesil))

    ws = wb["KUYRUKLAR"]
    for r in range(6, 17):
        ws.conditional_formatting.add(f"B{r}", CellIsRule(
            operator="greaterThan", formula=["0"], font=Font(color="C65911", bold=True)))
        ws.conditional_formatting.add(f"C{r}", CellIsRule(
            operator="greaterThan", formula=["0"], font=Font(color="C65911")))

    ws = wb["FINANSMAN"]
    ws.conditional_formatting.add("E6:E7", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("E6:E7", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("C6:C7", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("E6:E9", FormulaRule(
        formula=['ISNUMBER(SEARCH("YAPMA",E6))'], font=kirmizi))
    ws.conditional_formatting.add("E6:E9", FormulaRule(
        formula=['ISNUMBER(SEARCH("YAPILIR",E6))'], font=yesil))
    ws.conditional_formatting.add("B16:B18", CellIsRule(
        operator="greaterThan", formula=["0"], font=Font(color="C65911", bold=True)))

    ws = wb["VARSAYIMLAR"]
    ws.conditional_formatting.add("H6:H13", FormulaRule(
        formula=['H6="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("H6:H13", FormulaRule(
        formula=['H6="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("G6:G9", FormulaRule(
        formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(
        formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("F6:F9", CellIsRule(
        operator="notEqual", formula=["0"], font=kirmizi))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add("B5:B11", FormulaRule(
        formula=['OR(B5="HATALI",B5="BOŞ",B5="KIRIK",B5="AŞIYOR")'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B5:B11", FormulaRule(
        formula=['OR(B5="TEMİZ",B5="DOLU",B5="NORMAL",B5="TUTARLI")'], font=yesil, fill=y_fill))

    ws = wb["GIRDI"]
    ws.conditional_formatting.add("B10:B23", CellIsRule(
        operator="lessThan", formula=["0"], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B28", CellIsRule(
        operator="lessThan", formula=["100"], font=Font(color="B7791F", bold=True)))

    ws = wb["AYARLAR"]
    ws.conditional_formatting.add("H6:H50", FormulaRule(
        formula=['H6="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("H6:H50", FormulaRule(
        formula=['H6="TAM"'], font=yesil, fill=y_fill))


def ad_tanimlari(wb, aktif_kur_satir: int):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "zincirTolerans", "AYARLAR!$B$8")
    ad_ekle(wb, "isk_netKarTl", "MOTOR!$G$60")
    ad_ekle(wb, "isk_npv", "MOTOR!$G$61")
    ad_ekle(wb, "isk_irr", "MOTOR!$G$62")
    ad_ekle(wb, "isk_hedgeBosluk", "MOTOR!$G$63")
    ad_ekle(wb, "isk_kararMetni", "MOTOR!$G$42")
    ad_ekle(wb, "isk_kararGerekce", "MOTOR!$G$43")
    ad_ekle(wb, "isk_maddeAtifMetni", "MOTOR!$G$58")
    ad_ekle(wb, "isk_kanitRaporu", "MOTOR!$G$59")
    ad_ekle(wb, "isk_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "isk_parmakIzi", "AYARLAR!$B$19")
    ad_ekle(wb, "isk_aktifKur", f"AYARLAR!$K${aktif_kur_satir}")
    ad_ekle(wb, "isk_kurTablosu", f"AYARLAR!$K${aktif_kur_satir + 1}")
    ad_ekle(wb, "isk_finansmanKarsilastirma", "FINANSMAN!$B$11")
    ad_ekle(wb, "isk_durumHaritasi", "AYARLAR!$B$35")
    ad_ekle(wb, "isk_zincirKirik", "MOTOR!$G$65")

    ad_ekle(wb, "isk_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "isk_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "isk_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "isk_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "isk_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "isk_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "isk_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "isk_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "isk_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "isk_vakaFark", "VAKALAR!$B$13")

    ayar_map = {
        "isk_tolerans": 9,
        "isk_azamiEsik": 10,
        "isk_senaryoIyi": 11,
        "isk_senaryoKotu": 12,
        "isk_esikArtis": 13,
        "isk_tahminAlt": 14,
        "isk_tahminUst": 15,
        "isk_tornadoKur": 16,
        "isk_tornadoMaliyet": 17,
        "isk_yuzdelikOran": 18,
        "isk_girisBeklenen": 21,
        "isk_riskDusuk": 22,
        "isk_riskOrta": 23,
        "isk_riskYuksek": 24,
        "isk_yilTaban": 25,
        "isk_yuvarMax": 26,
        "isk_bazCarpan": 27,
        "isk_sifirCarpan": 28,
        "isk_siraBir": 29,
        "isk_siraIki": 30,
        "isk_siraUc": 31,
        "isk_dikkatCarpan": 32,
        "isk_npvDikkat": 33,
        "isk_hedgeDikkat": 34,
        "vaka_fob": 36,
        "vaka_rez": 37,
        "vaka_tah": 38,
        "vaka_uretim": 39,
        "vaka_navlun": 40,
        "vaka_kom": 41,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "isk_modulT1": (11, "C"), "isk_modulT1Yorum": (11, "D"),
        "isk_modulT2": (12, "C"), "isk_modulT2Yorum": (12, "D"),
        "isk_modulO1": (13, "C"), "isk_modulO1Yorum": (13, "D"),
        "isk_modulO8": (16, "C"), "isk_modulO8Yorum": (16, "D"),
        "isk_modulI2": (19, "C"), "isk_modulI2Yorum": (19, "D"),
    }
    for ad, (satir, kol) in modul_map.items():
        ad_ekle(wb, ad, f"PANO!${kol}${satir}")

    ad_ekle(wb, "ListeYillar", "LISTELER!$A$6:$A$8")
    ad_ekle(wb, "ListeDoviz", "LISTELER!$C$6:$C$7")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$6:$E$7")
    ad_ekle(wb, "ListeDurumlar", "LISTELER!$G$6:$G$12")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$I$6:$I$20")


def main(cikti_yolu=None):
    wb = Workbook()
    wb.remove(wb.active)

    siralar = [
        ("KAPAK", kapak),
        ("HIZLI_BASLANGIC", hizli_baslangic),
        ("GIRDI", girdi),
        ("KARTLAR", kartlar),
        ("AKIS", akis),
        ("KUYRUKLAR", kuyruklar),
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
    aktif_kur_ozet_satir = ayarlar(ws_ay)
    ws_kil = wb.create_sheet("KILAVUZ")
    kilavuz(ws_kil)

    ad_tanimlari(wb, aktif_kur_ozet_satir)
    kosullu_bicimlendirme(wb)
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    wb.calculation.fullCalcOnLoad = True

    kok = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    varsayilan = os.path.join(kok, "IhracatSiparisKarlilikKur.xlsx")
    hedef = cikti_yolu or varsayilan
    os.makedirs(os.path.dirname(os.path.abspath(hedef)) or ".", exist_ok=True)
    wb.save(hedef)

    hsh = hashlib.sha256()
    with open(hedef, "rb") as f:
        for b in iter(lambda: f.read(65536), b""):
            hsh.update(b)
    dig = hsh.hexdigest()

    wb2 = __import__("openpyxl").load_workbook(hedef)
    wb2["AYARLAR"]["B19"] = dig[:16]
    for wsx in wb2.worksheets:
        sayfa_koru(wsx)
    wb2.save(hedef)

    print(f"Dosya: {hedef}")
    print(f"SHA-256: {dig}")
    print(f"Şifre: {SIFRE}")
    print(f"Aktif kur özet satır: {aktif_kur_ozet_satir}")
    return hedef


if __name__ == "__main__":
    yol = sys.argv[1] if len(sys.argv) > 1 else None
    uretilen = main(yol)
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
    cikti = os.path.join(repo, "cikti", "IhracatSiparisKarlilikKur.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
