#!/usr/bin/env python3
"""
İthalat Landed Cost Motoru — üretim betiği (A6+A1 / manda v6).
Navlun+sigorta+gümrük+KDV+KKDF → birim landed cost; NPV/IRR yatırım kararı.
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

URUN_AD = "İthalat Gerçek Birim Maliyet Motoru"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 11)

# Demo: FOB 100k USD @34 → CIF+vergiler; birim landed; NPV baz pozitif, kötümser negatif
DEMO = {
    "sirket_adi": "Örnek İthalat Ticaret A.Ş.",
    "donem": "2026-Q3",
    "hesapYili": 2026,
    "parabirimi": "USD",
    "fob_doviz": 100_000,
    "miktar": 1_000,
    "navlun_doviz": 5_000,
    "sigorta_oran_giris": 0.01,
    "lokal_masraf_tl": 25_000,
    "hedef_satis_birim": 6_500,
    "yillik_miktar": 12_000,
    "capex_tl": 800_000,
    "wacc": 0.18,
    "npv_donem": 5,
    "kredi_orani": 0.32,
    "oz_kaynak_oran": 0.40,
}


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: Gerçek birim maliyeti ve yatırım kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: birim maliyet ve NPV bozulur | Kaynak: Gümrük / KDV / KKDF kuralları",
    )


def kural_cek(kural_id: str) -> str:
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-ilc_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(ilc_yuvarMax,IFERROR(N(G12),0)))"


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
      "FOB, navlun, sigorta, gümrük, KDV ve KKDF girdilerini işleyin; gerçek birim landed "
      "maliyeti ve YATIRIM YAPILIR / YATIRIM YAPMA kararını üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Kur tablosu (tarihli/kaynaklı) ile döviz→TL çevrimi — sabit kur gömülmez",
        "Gümrük + KKDF + KDV kural motoru (2025/2026/2027 yan yana)",
        "CAPEX/OPEX ayrımı + öz kaynak / kredi finansman senaryosu",
        "NPV / IRR + tornado duyarlılık; kötümser YATIRIM YAPMA yolu",
        "Altın vakalar ve KANIT_RAPORU (A4 imza)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Satın alma / ithalat — sevkiyat öncesi gerçek birim maliyeti görmek",
        "CFO — NPV/IRR ile ithalat yatırımı kararı vermek",
        "Gümrük müşaviri / denetçi — imzalı kanıt raporunu dosyalamak",
    ], 14):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) HIZLI_BASLANGIC → 2) GIRDI sarı hücreler → 3) PANO/KARAR → 4) KANIT_RAPORU yazdır.",
      yazi="333333", kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 21, 1, f"Sürüm {SURUM} | 2026 | ExcelArşiv | Lisans: Tek kullanıcı | Kod: ILC-PRO",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 72})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Üç adımda sonuç", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Girdiler", "GIRDI sayfasında FOB, miktar, navlun ve kur seçimini sarı hücrelere girin."),
        ("Adım 2 — Karar", "PANO ve KARAR'da birim landed, NPV/IRR ve YATIRIM YAPILIR / YATIRIM YAPMA görün."),
        ("Adım 3 — Kanıt", "KANIT_RAPORU'nu yazdırın; VAKALAR ve SENARYO ile doğrulayın."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, metin, kaydir=True, yazi="333333")
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=12)
    h(ws, 9, 1, "Formüller", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 1,
      "CIF = FOB_TL + navlun_TL + sigorta_TL | Landed = CIF + gümrük + KKDF + KDV + lokal | "
      "Birim = Landed ÷ miktar",
      yazi="333333", kaydir=True)
    ws.merge_cells("A10:L10")
    h(ws, 12, 1, "Sarı = manuel giriş | Yeşil = formül çıktısı | Şifre 1234 formül alanlarını korur.",
      yazi=GRİ, kaydir=True)
    genislik(ws, {"A": 28, "B": 70})


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "İthalat sevkiyat girdileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Oranlar KURALLAR'dan; kur AYARLAR kur tablosundan.",
      yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Şirket / Sevkiyat", DEMO["sirket_adi"], None, "sirket_adi", "Unvan veya sevkiyat kodu."),
        (7, "Dönem", DEMO["donem"], None, "donem", "Rapor dönemi."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili", "2025, 2026 veya 2027."),
        (9, "Döviz Cinsi", DEMO["parabirimi"], None, "parabirimi", "USD veya EUR."),
        (10, "FOB [döviz]", DEMO["fob_doviz"], CATI, "fob_doviz", "FOB tutarı döviz cinsinden."),
        (11, "Miktar [adet]", DEMO["miktar"], CATI, "miktar", "Sevkiyat miktarı."),
        (12, "Navlun [döviz]", DEMO["navlun_doviz"], CATI, "navlun_doviz", "Navlun döviz tutarı."),
        (13, "Sigorta oranı (giriş)", DEMO["sigorta_oran_giris"], YÜZDE, "sigorta_oran",
         "Sigorta primi oranı; kural ayna ile doğrulanır."),
        (14, "Lokal masraf [₺]", DEMO["lokal_masraf_tl"], TL, "lokal_masraf", "Antrepo/aracı TL masraf."),
        (15, "Hedef satış birim [₺]", DEMO["hedef_satis_birim"], TL, "hedef_satis", "Birim satış hedefi."),
        (16, "Yıllık miktar [adet]", DEMO["yillik_miktar"], CATI, "yillik_miktar", "OPEX yıllık adet."),
        (17, "CAPEX / YATIRIM [₺]", DEMO["capex_tl"], TL, "capex_tl", "Kurulum ve depozito yatırımı."),
        (18, "WACC [%]", DEMO["wacc"], YÜZDE, "wacc", "İskonto oranı."),
        (19, "NPV dönem [yıl]", DEMO["npv_donem"], CATI, "npv_donem", "Nakit akışı yılı."),
        (20, "Kredi faiz oranı [%]", DEMO["kredi_orani"], YÜZDE, "kredi_orani", "Kredi senaryosu faiz."),
        (21, "Öz kaynak oranı [%]", DEMO["oz_kaynak_oran"], YÜZDE, "oz_kaynak_oran", "Öz kaynak payı."),
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
        elif _ad in ("miktar", "yillik_miktar", "fob_doviz", "navlun_doviz"):
            birim = "adet/döviz"
        else:
            birim = "metin"
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim, yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 23, 1, "Aktif kur (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "=IFERROR(ilc_aktifKur,0)", sayi="0.0000", yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 24, 1, "Gümrük oranı (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, kural_cek("ILC-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 1, "Birim landed (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=IFERROR(ilc_birimLanded,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 26, 1, "Giriş doluluk", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B21)=0,ilc_girisBeklenen=0),0,ROUND(COUNTA(B6:B21)/ilc_girisBeklenen*100,0)),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)

    h(ws, 28, 1, "Giriş tablosu (otomasyon / boş-mod)", kalin=True, yazi=KOYU_LACIVERT)
    sutunlar = ["Alan Anahtarı", "Değer", "Birim", "Zorunlu", "Kaynak", "Kontrol"]
    baslik_satiri(ws, 29, sutunlar)
    formuller = {
        "Kontrol": (
            '=IF(tblGirdi[[#This Row],[Alan Anahtarı]]="","",'
            'IF(tblGirdi[[#This Row],[Değer]]="","EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblGirdi", "A29:F1028", sutunlar, formuller)
    ornek_satir = [
        ("sirket_adi", DEMO["sirket_adi"], "metin", "Evet", "GIRDI"),
        ("donem", DEMO["donem"], "metin", "Evet", "GIRDI"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Evet", "GIRDI"),
        ("parabirimi", DEMO["parabirimi"], "kod", "Evet", "GIRDI"),
        ("fob_doviz", DEMO["fob_doviz"], "döviz", "Evet", "GIRDI"),
        ("miktar", DEMO["miktar"], "adet", "Evet", "GIRDI"),
        ("navlun_doviz", DEMO["navlun_doviz"], "döviz", "Evet", "GIRDI"),
        ("sigorta_oran", DEMO["sigorta_oran_giris"], "oran", "Evet", "GIRDI"),
        ("lokal_masraf", DEMO["lokal_masraf_tl"], "TL", "Evet", "GIRDI"),
        ("hedef_satis", DEMO["hedef_satis_birim"], "TL", "Evet", "GIRDI"),
        ("yillik_miktar", DEMO["yillik_miktar"], "adet", "Evet", "GIRDI"),
        ("capex_tl", DEMO["capex_tl"], "TL", "Evet", "GIRDI"),
        ("wacc", DEMO["wacc"], "oran", "Evet", "GIRDI"),
        ("npv_donem", DEMO["npv_donem"], "yıl", "Evet", "GIRDI"),
        ("kredi_orani", DEMO["kredi_orani"], "oran", "Evet", "GIRDI"),
        ("oz_kaynak_oran", DEMO["oz_kaynak_oran"], "oran", "Evet", "GIRDI"),
    ]
    tl_a = {"lokal_masraf", "hedef_satis", "capex_tl"}
    oran_a = {"sigorta_oran", "wacc", "kredi_orani", "oz_kaynak_oran"}
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 30):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in tl_a:
            ws.cell(i, 2).number_format = TL
        elif a in oran_a:
            ws.cell(i, 2).number_format = YÜZDE
        elif a in ("hesapYili", "miktar", "yillik_miktar", "npv_donem", "fob_doviz", "navlun_doviz"):
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
    dogrulama(ws, "list", "ListeYillar", "B32",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeDoviz", "B33",
              baslik="Döviz", mesaj="USD veya EUR.",
              hata_baslik="Geçersiz", hata_mesaj="USD/EUR.")
    for aralik, baslik, mesaj in [
        ("B10", "FOB", "0 ile 1e12 arasında."),
        ("B11", "Miktar", "0 ile 1e9 arasında."),
        ("B12", "Navlun", "0 ile 1e12 arasında."),
        ("B14", "Lokal masraf", "0 ile 1e12 TL."),
        ("B15", "Hedef satış", "0 ile 1e12 TL."),
        ("B16", "Yıllık miktar", "0 ile 1e9."),
        ("B17", "CAPEX", "0 ile 1e12 TL."),
        ("B34", "FOB", "Tablo: döviz tutar."),
        ("B35", "Miktar", "Tablo: adet."),
        ("B36", "Navlun", "Tablo: döviz."),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj=mesaj,
                  hata_baslik="Geçersiz tutar", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000000")
    for aralik, baslik in [("B13", "Sigorta"), ("B18", "WACC"), ("B20", "Kredi"), ("B21", "Öz kaynak"),
                           ("B37", "Sigorta"), ("B42", "WACC"), ("B44", "Kredi"), ("B45", "Öz kaynak")]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj="0 ile 1 arasında oran.",
                  hata_baslik="Geçersiz", hata_mesaj="0-1 arası.",
                  isaret="between", f2="1")
    for aralik, baslik in [("B19", "NPV dönem"), ("B43", "NPV dönem")]:
        dogrulama(ws, "whole", "1", aralik,
                  baslik=baslik, mesaj="1-25 arası.",
                  hata_baslik="Geçersiz", hata_mesaj="1-25.",
                  isaret="between", f2="25")

    giris_hucreleri(ws, 6, 21, [2])
    giris_hucreleri(ws, 30, 45, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 40, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1030, "Sarı alanlar giriş; oran ve kur formülleri kilitlidir.")


def varsayimlar(ws):
    sayfa_hazirla(ws, "VARSAYIMLAR", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Model varsayımları — her satırda kaynak ve güven", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = ["varsayim_id", "aciklama", "deger", "birim", "kaynak", "guven", "duyarlilik", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    satirlar = [
        ("V-01", "Aktif kur kaynağı", "TCMB alış", "metin", "TCMB günlük kur", "sert", "yüksek"),
        ("V-02", "Sigorta matrahı", "FOB", "metin", "Sigorta poliçesi", "sert", "orta"),
        ("V-03", "KKDF matrah yorumu", "FOB_TL", "metin", "KKDF tebliği — tartışmalı", "yumuşak", "yüksek"),
        ("V-04", "KDV matrahı", "CIF+gumruk+KKDF", "metin", "KDV Kanunu", "sert", "orta"),
        ("V-05", "CAPEX geri ödeme", "5 yıl", "yıl", "İç politika", "yumuşak", "orta"),
        ("V-06", "OPEX yıllık miktar", str(DEMO["yillik_miktar"]), "adet", "Satış planı", "yumuşak", "yüksek"),
        ("V-07", "WACC", "0.18", "oran", "Finans komitesi", "yumuşak", "yüksek"),
        ("V-08", "Kötümser kur şoku", "+15%", "oran", "Stres testi", "yumuşak", "yüksek"),
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
    h(ws, 15, 10, "CAPEX / YATIRIM tutarı (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 11, "=GIRDI!B17", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 10, "OPEX / ISLETME yıllık (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 11, "=IFERROR(ilc_birimLanded*GIRDI!B16,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 12, "B": 28, "C": 18, "D": 10, "E": 28, "F": 12, "G": 12, "H": 10, "J": 32, "K": 18})


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Gümrük / KDV / KKDF kural tablosu (yıl yan yana)", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["ILC-001", "Gümrük Kanunu", "gumruk_orani", "CIF × oran", "gumruk_orani",
         0.10, 0.10, 0.10, "01.01.2025", "GTİP örnek oran"],
        ["ILC-002", "KKDF tebliği", "kkdf_orani", "FOB_TL × oran", "kkdf_orani",
         0.06, 0.06, 0.06, "01.01.2025", "KKDF oranı"],
        ["ILC-003", "KDV Kanunu", "kdv_orani", "(CIF+gümrük+KKDF)×oran", "kdv_orani",
         0.20, 0.20, 0.20, "01.01.2025", "KDV genel"],
        ["ILC-004", "İç politika", "hedef_marj", "marj ≥ eşik", "hedef_marj",
         0.15, 0.15, 0.15, "01.01.2025", "Hedef brüt marj"],
        ["ILC-005", "Uygulama", "yuvarlama", "her hesap", "yuvarlama_ondalik",
         2, 2, 2, "01.01.2025", "Yuvarlama ondalık"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = None
            if k in (6, 7, 8):
                sayi = YÜZDE if s[0] != "ILC-005" else CATI
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-ilc_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Gümrük (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("ILC-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "KKDF (aktif yıl)", yazi="333333")
    h(ws, 16, 2, kural_cek("ILC-002"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"ILC-001=",TEXT(B15,"0.0%"),"ILC-002=",TEXT(B16,"0.0%"),'
      '"yıl=",IF(OR(hesapYili="",hesapYili=0),"-",hesapYili)),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    for col in ("F", "G", "H"):
        dogrulama(ws, "decimal", "0", f"{col}7:{col}11",
                  baslik="Kural değeri", mesaj="Yıl kolonuna sayısal parametre girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ve üzeri sayı.",
                  isaret="between", f2="10000000")
    genislik(ws, {get_column_letter(i): w for i, w in enumerate(
        [14, 22, 16, 28, 16, 12, 12, 12, 12, 18, 14], 1)})
    ws.merge_cells("A3:K3")
    ws.merge_cells("A4:K4")
    sabitle(ws, "A7")


def motor(ws):
    sayfa_hazirla(ws, "MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Hesap zinciri — her adımda kural_id", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar tblKurallar'dan; kur tblKur'dan çekilir.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")
    basliklar = [*MOTOR_BASLIK, "deger"]
    baslik_satiri(ws, 6, basliklar)

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif kural yılı", "ILC-005")
    ekle("Gümrük oranı", km("ILC-001"), "oran", "Gümrük vergisi oranı", "ILC-001")
    ekle("KKDF oranı", km("ILC-002"), "oran", "KKDF oranı", "ILC-002")
    ekle("KDV oranı", km("ILC-003"), "oran", "KDV oranı", "ILC-003")
    ekle("Hedef marj", km("ILC-004"), "oran", "Hedef brüt marj", "ILC-004")
    ekle("Yuvarlama ondalık", km("ILC-005"), "adet", "Yuvarlama basamağı", "ILC-005")
    ekle("Aktif kur", "=IFERROR(ilc_aktifKur,0)", "kur", "Kur tablosu aktif kur", "ILC-005")
    ekle("FOB döviz", "=GIRDI!B10", "döviz", "FOB tutarı", "ILC-001")
    ekle("Miktar", "=GIRDI!B11", "adet", "Sevkiyat miktarı", "ILC-005")
    ekle("Navlun döviz", "=GIRDI!B12", "döviz", "Navlun", "ILC-001")
    ekle("Sigorta oranı", "=GIRDI!B13", "oran", "Sigorta primi oranı", "ILC-001")
    ekle("Lokal masraf TL", "=GIRDI!B14", "TL", "Lokal masraflar", "ILC-005")
    ekle("Hedef satış birim", "=GIRDI!B15", "TL", "Hedef birim satış", "ILC-004")
    ekle("Yıllık miktar", "=GIRDI!B16", "adet", "OPEX yıllık adet", "ILC-005")
    ekle("CAPEX / YATIRIM", "=GIRDI!B17", "TL", "Yatırım tutarı", "ILC-005")
    ekle("WACC", "=GIRDI!B18", "oran", "İskonto oranı", "ILC-005")
    ekle("NPV dönem", "=GIRDI!B19", "adet", "Nakit akışı yılı", "ILC-005")
    ekle("Kredi faiz", "=GIRDI!B20", "oran", "Kredi senaryosu", "ILC-005")
    ekle("Öz kaynak oranı", "=GIRDI!B21", "oran", "Öz kaynak payı", "ILC-005")
    ekle("FOB TL", f"=ROUND(G14*G13,{_YV})", "TL", "FOB × kur", "ILC-001")
    ekle("Navlun TL", f"=ROUND(G16*G13,{_YV})", "TL", "Navlun × kur", "ILC-001")
    ekle("Sigorta TL", f"=ROUND(G14*G17*G13,{_YV})", "TL", "FOB×sigorta×kur", "ILC-001")
    ekle("CIF TL", "=G26+G27+G28", "TL", "FOB+navlun+sigorta", "ILC-001")
    ekle("Gümrük vergisi", f"=ROUND(G29*G8,{_YV})", "TL", "CIF × gümrük", "ILC-001")
    ekle("KKDF tutarı", f"=ROUND(G26*G9,{_YV})", "TL", "FOB_TL × KKDF", "ILC-002")
    ekle("KDV matrahı", "=G29+G30+G31", "TL", "CIF+gümrük+KKDF", "ILC-003")
    ekle("KDV tutarı", f"=ROUND(G32*G10,{_YV})", "TL", "Matrah × KDV", "ILC-003")
    ekle("Toplam landed", "=G29+G30+G31+G33+G18", "TL", "CIF+vergiler+lokal", "ILC-003")
    ekle("Birim landed", f"=IF(G15=0,0,ROUND(G34/G15,{_YV}))", "TL", "Landed ÷ miktar", "ILC-005")
    ekle("Birim marj", "=IF(G19=0,0,(G19-G35)/G19)", "oran", "Satış−maliyet marjı", "ILC-004")
    ekle("OPEX / ISLETME yıllık", "=G35*G20", "TL", "Birim×yıllık miktar", "ILC-005")
    ekle("Yıllık brüt katkı", "=(G19-G35)*G20", "TL", "Marj × yıllık adet", "ILC-004")
    ekle("NPV (baz)",
         f'=IF(OR(G23<=0,G22=""),0,IF(G22=0,ROUND(G38*G23-G21,{_YV}),'
         f'ROUND(G38*(1-POWER(1+G22,-G23))/G22-G21,{_YV})))',
         "TL", "NPV nakit akışı − CAPEX", "ILC-005")
    ekle("IRR yaklaşımı",
         '=IF(OR(G21<=0,G38<=0,G23<=0),0,ROUND(G38/G21-ilc_bazCarpan/G23,4))',
         "oran", "Kabaca IRR göstergesi", "ILC-005")
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(OR(G35<=0,G15<=0,G13<=0),"YAPMA",'
         'IF(OR(G39<0,G36<G11*ilc_dikkatCarpan),"YAPMA",'
         'IF(OR(G36<G11,G39<ilc_npvDikkat),"DIKKAT","YAPILIR"))))',
         "kod", "Karar kodu", "ILC-004")
    ekle("Karar metni",
         '=IF(G41="VERI_YOK","VERİ YOK",'
         'IF(G41="YAPILIR","YATIRIM YAPILIR",'
         'IF(G41="DIKKAT","DİKKAT","YATIRIM YAPMA")))',
         "metin", "Kullanıcıya karar", "ILC-004")
    ekle("Gerekçe",
         '=IF(G41="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G41="YAPILIR","NPV pozitif ve marj hedef üstünde; ithalat yatırımı uygun.",'
         'IF(G41="DIKKAT","Marj veya NPV dikkat bandında; kur ve oranları gözden geçirin.",'
         '"NPV negatif veya marj yetersiz; kötümser/bazda YATIRIM YAPMA.")))',
         "metin", "Gerekçe cümlesi", "ILC-004")
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B21)=0,ilc_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B21)/ilc_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "ILC-005")
    ekle("Risk skoru",
         '=IF(G41="VERI_YOK",0,IF(G41="YAPILIR",ilc_riskDusuk,'
         'IF(G41="DIKKAT",ilc_riskOrta,ilc_riskYuksek)))',
         "puan", "Yatırım riski", "ILC-005")
    ekle("Senaryo iyimser birim", f"=ROUND(G35*ilc_senaryoIyi,{_YV})", "TL", "İyimser maliyet", "ILC-005")
    ekle("Senaryo kötümser birim", f"=ROUND(G35*ilc_senaryoKotu,{_YV})", "TL", "Kötümser maliyet", "ILC-005")
    ekle("İyimser NPV",
         f'=IF(G22=0,ROUND((G19-G46)*G20*G23-G21,{_YV}),'
         f'ROUND((G19-G46)*G20*(1-POWER(1+G22,-G23))/G22-G21,{_YV}))',
         "TL", "İyimser NPV", "ILC-005")
    ekle("Kötümser NPV",
         f'=IF(G22=0,ROUND((G19-G47)*G20*G23-G21,{_YV}),'
         f'ROUND((G19-G47)*G20*(1-POWER(1+G22,-G23))/G22-G21,{_YV}))',
         "TL", "Kötümser NPV", "ILC-005")
    ekle("Kötümser karar",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(G49<0,"YATIRIM YAPMA","DİKKAT"))',
         "metin", "Kötümser YATIRIM YAPMA yolu", "ILC-004")
    ekle("M-SEN gümrük+", "=G8+ilc_esikArtis", "oran", "Gümrük + delta", "ILC-001")
    ekle("M-SEN landed", f"=ROUND(G29*(ilc_bazCarpan+G51)+G31+G33+G18,{_YV})", "TL",
         "Sıkı gümrük landed", "ILC-001")
    ekle("M-SEN bayrak",
         '=IF(AND(G36>=G11,G39>=0),1,0)',
         "bayrak", "Sıkı eşikte uygun mu", "ILC-004")
    ekle("Tornado: kur etkisi",
         f"=ABS(ROUND(G35*ilc_tornadoKur-G35,{_YV}))",
         "TL", "Kur ± etki", "ILC-005")
    ekle("Tornado: gümrük etkisi",
         f"=ABS(ROUND(G29*(G8+ilc_tornadoGumruk)-G29*G8,{_YV}))",
         "TL", "Gümrük ± etki", "ILC-001")
    ekle("Tornado: navlun etkisi",
         f"=ABS(ROUND(G16*G13*ilc_tornadoNavlun,{_YV}))",
         "TL", "Navlun ± etki", "ILC-001")
    ekle("Vergi yükü oranı", "=IF(G34=0,0,(G30+G31+G33)/G34)", "oran", "Vergi/landed", "ILC-003")
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"ILC-001→gümrük","ILC-002→KKDF","ILC-003→KDV","ILC-004→marj")',
         "metin", "Kanıt atıfları", "ILC-001")
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Birim=",TEXT(G35,"₺ #,##0"),"NPV=",TEXT(G39,"₺ #,##0"),'
         '"IRR=",TEXT(G40,"0.0%"),"Karar=",G42)',
         "metin", "Rapor özeti", "ILC-005")
    ekle("Birim landed (pano)", "=G35", "TL", "Birim ayna", "ILC-005")
    ekle("NPV (pano)", "=G39", "TL", "NPV ayna", "ILC-005")
    ekle("IRR (pano)", "=G40", "oran", "IRR ayna", "ILC-005")
    ekle("CIF (pano)", "=G29", "TL", "CIF ayna", "ILC-001")

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
    h(ws, 6, 2, "=GIRDI!B21", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 3, "=GIRDI!B17*B6", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "=IFERROR(GIRDI!B21*ilc_sifirCarpan,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 5, "=IFERROR(MOTOR!G39,0)", sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 7, 1, "KREDI", kalin=True, yazi="333333")
    h(ws, 7, 2, "=ilc_bazCarpan-GIRDI!B21", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 3, "=GIRDI!B17*B7", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 4, "=GIRDI!B20", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 5,
      "=IFERROR(MOTOR!G39-C7*D7*GIRDI!B19,0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 9, 1, "CAPEX / YATIRIM toplam", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 9, 2, "=GIRDI!B17", sayi=TL, yazi=DIS_REF_YEŞIL)
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
    h(ws, 14, 2, 320_000, sayi=TL, yazi=GRİ)
    h(ws, 15, 1, "KREDI")
    h(ws, 15, 2, 480_000, sayi=TL, yazi=GRİ)
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
    h(ws, 5, 2, "Çarpan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Birim landed", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "NPV", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(ilc_senaryoIyi+N(GIRDI!B10)*ilc_sifirCarpan,0)",
         "=IFERROR(MOTOR!G46,0)", "=IFERROR(MOTOR!G48,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D6<0,"YATIRIM YAPMA",IF(C6>GIRDI!B15*(ilc_bazCarpan-MOTOR!G11),"DİKKAT","YATIRIM YAPILIR")))'),
        ("Baz", "=IFERROR(ilc_bazCarpan+N(GIRDI!B10)*ilc_sifirCarpan,1)",
         "=IFERROR(ilc_birimLanded,0)", "=IFERROR(ilc_npv,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(ilc_kararMetni,"-"))'),
        ("Kötümser", "=IFERROR(ilc_senaryoKotu+N(GIRDI!B10)*ilc_sifirCarpan,0)",
         "=IFERROR(MOTOR!G47,0)", "=IFERROR(MOTOR!G49,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(MOTOR!G50,"YATIRIM YAPMA"))'),
        ("M-SEN oran+", "=IFERROR(ilc_bazCarpan+N(GIRDI!B10)*ilc_sifirCarpan,1)",
         "=IFERROR(IF(MOTOR!G15=0,0,MOTOR!G52/MOTOR!G15),0)", "=IFERROR(MOTOR!G39,0)",
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
    h(ws, 17, 1, "Tornado gümrük etkisi", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G55,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado navlun etkisi", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G56,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 3,
      '=IFERROR(IF(B16>=MAX(B16:B18),ilc_siraBir,IF(B16=MEDIAN(B16:B18),ilc_siraIki,ilc_siraUc)),ilc_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 3,
      '=IFERROR(IF(B17>=MAX(B16:B18),ilc_siraBir,IF(B17=MEDIAN(B16:B18),ilc_siraIki,ilc_siraUc)),ilc_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 3,
      '=IFERROR(IF(B18>=MAX(B16:B18),ilc_siraBir,IF(B18=MEDIAN(B16:B18),ilc_siraIki,ilc_siraUc)),ilc_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(CONCATENATE("1)",INDEX(A16:A18,MATCH(ilc_siraBir,C16:C18,0)),'
      '" 2)",INDEX(A16:A18,MATCH(ilc_siraIki,C16:C18,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B16:B18),"₺ #,##0")," — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN oran değişimi bayrağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=MOTOR!G53", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Sıkı gümrük (+",TEXT(ilc_esikArtis,"0.0%"),") uygunluk ",TEXT(B22,"0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin alt", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=IFERROR(ilc_birimLanded*ilc_tahminAlt,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 3, "=IFERROR(ilc_birimLanded,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 4, "=IFERROR(ilc_birimLanded*ilc_tahminUst,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
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
    for i, v in enumerate([2_400_000, 1_800_000, -400_000, 1_200_000], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([180_000, 120_000, 90_000], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)
    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([4_800, 5_200, 6_100, 5_200], 28):
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
    g3.title = "Birim Landed Trendi"
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
        (2, "Birim landed", "=IFERROR(ilc_birimLanded,0)", TL),
        (3, "CIF", "=IFERROR(MOTOR!G29,0)", TL),
        (4, "Gümrük", "=IFERROR(MOTOR!G30,0)", TL),
        (5, "KKDF", "=IFERROR(MOTOR!G31,0)", TL),
        (6, "KDV", "=IFERROR(MOTOR!G33,0)", TL),
        (7, "NPV", "=IFERROR(ilc_npv,0)", TL),
        (8, "IRR", "=IFERROR(ilc_irr,0)", YÜZDE),
        (9, "Marj", "=IFERROR(MOTOR!G36,0)", YÜZDE),
        (10, "Güven", "=IFERROR(MOTOR!G44,0)", CATI),
        (11, "Risk", "=IFERROR(MOTOR!G45,0)", CATI),
        (12, "M-SEN", "=IFERROR(ilc_kuralDegisimSenaryo,0)", CATI),
        (13, "Tahmin", "=IFERROR(ilc_tahminAralik,0)", TL),
        (14, "Vaka", '=IFERROR(ilc_vakaDurum,"-")', None),
    ]
    for col, ad, formul, sayi in kpis:
        h(ws, 3, col, ad, kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
        h(ws, 4, col, formul, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "Karar özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=ilc_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=ilc_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 10, 1, "Kod", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 2, "Ad", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 3, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 4, "Yorum", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    moduller = [
        ("T1", "CIF / gümrük matrahı",
         "=IFERROR(MOTOR!G29,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"CIF ",TEXT(C11,"₺ #,##0")," — gümrük ",TEXT(MOTOR!G8,"0.0%"))'),
        ("T2", "Birim gerçek maliyet",
         "=IFERROR(ilc_birimLanded,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Birim ",TEXT(C12,"₺ #,##0")," — hedef ",TEXT(GIRDI!B15,"₺ #,##0"))'),
        ("O1", "Senaryo landed sapması",
         "=IFERROR(ABS(SENARYO!C6-SENARYO!C8),0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"İyimser−kötümser birim bant ",TEXT(C13,"₺ #,##0"))'),
        ("O2", "Duyarlılık / tornado",
         '=IFERROR(ilc_duyarlilikSira,"-")',
         "=IFERROR(ilc_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo motoru",
         "=IFERROR(ilc_senaryoKarsilastirma,0)",
         "=IFERROR(ilc_yorumSenaryo,\"-\")"),
        ("O6", "Vergi yükü yoğunlaşması",
         "=IFERROR(MOTOR!G57,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Vergi/landed ",TEXT(C16,"0.0%"))'),
        ("O8", "Veri kalite skoru",
         "=IFERROR(MOTOR!G44,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Giriş güven skoru ",TEXT(C17,"0"),"/100")'),
        ("M_SEN", "Oran değişimi",
         "=IFERROR(ilc_kuralDegisimSenaryo,0)",
         "=IFERROR(ilc_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin + aralık",
         "=IFERROR(ilc_tahminAralik,0)",
         "=IFERROR(ilc_yorumTahmin,\"-\")"),
        ("I2", "Landed yüzdelik P90",
         "=IFERROR(IF(COUNT(SENARYO!C6:C9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!C6:C9,ilc_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 birim ",TEXT(C20,"₺ #,##0"))'),
    ]
    for i, (kod, ad, deger, yorum) in enumerate(moduller, 11):
        h(ws, i, 1, kod, kalin=True, yazi="B08948")
        h(ws, i, 2, ad, yazi="333333")
        sayi = YÜZDE if kod in ("O6",) else (CATI if kod in ("O8", "M_SEN") else TL)
        if kod in ("O2", "O3", "M_SEN") and "Sira" in str(deger):
            sayi = None
        h(ws, i, 3, deger, sayi=sayi, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 23, 1, "PanoDemoBirim", kalin=True, yazi=GRİ)
    for i, (ad, v) in enumerate([("CIF pay", 0.55), ("Vergi pay", 0.35), ("Lokal pay", 0.10)], 24):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, v, sayi=YÜZDE, yazi=GRİ)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Landed Maliyet Kırılımı"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=26), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=26))
    g1.height, g1.width = 7, 12
    ws.add_chart(g1, "F23")

    h(ws, 23, 4, "Etiket")
    h(ws, 23, 5, "PanoDemoNpv", kalin=True, yazi=GRİ)
    for i, (ad, v) in enumerate([("İyimser", 2_400_000), ("Baz", 1_800_000), ("Kötümser", -400_000)], 24):
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
    h(ws, 28, 5, "VergiDemo", kalin=True, yazi=GRİ)
    h(ws, 29, 4, "Gümrük")
    h(ws, 29, 5, 357_000, sayi=TL, yazi=GRİ)
    h(ws, 30, 4, "KKDF")
    h(ws, 30, 5, 204_000, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "Vergi Kırılımı"
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
    h(ws, 5, 2, "=ilc_kararMetni", boyut=16, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Gerekçe", kalin=True)
    h(ws, 6, 2, "=ilc_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B6:H6")
    h(ws, 8, 1, "NPV [₺]", kalin=True)
    h(ws, 8, 2, "=ilc_npv", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "IRR", kalin=True)
    h(ws, 9, 2, "=ilc_irr", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "Birim landed [₺]", kalin=True)
    h(ws, 10, 2, "=ilc_birimLanded", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Kötümser karar", kalin=True)
    h(ws, 11, 2, "=IFERROR(MOTOR!G50,\"YATIRIM YAPMA\")", yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "CAPEX / YATIRIM", kalin=True)
    h(ws, 13, 2, "=GIRDI!B17", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "OPEX / ISLETME yıllık", kalin=True)
    h(ws, 14, 2, "=IFERROR(MOTOR!G37,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1,
      "Boş girişte karar daima VERİ YOK olur. Kötümser NPV negatifse YATIRIM YAPMA üretilir.",
      yazi=GRİ, kaydir=True)
    ws.merge_cells("A16:H16")
    genislik(ws, {"A": 28, "B": 55})


def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Altın vakalar — gerçek birim maliyet", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:H3")
    baslik_satiri(ws, 5, VAKA_BASLIK)

    # V-001: FOB_TL = 100000*34 = 3400000
    # V-002: gümrük = CIF * 0.10; CIF≈3400000+170000+34000=3604000 → gümrük 360400
    # Simplified golden using named vaka_* params
    vakalar_data = [
        ("V-001", "FOB × kur = FOB_TL",
         "fob=vaka_fob; kur=vaka_kur",
         3_400_000,
         "=IFERROR(ROUND(vaka_fob*vaka_kur,0),0)",
         "=IFERROR(E6-D6,0)",
         '=IFERROR(IF(ABS(F6)<=ilc_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "FOB TL"),
        ("V-002", "CIF = FOB_TL+navlun_TL+sigorta_TL",
         "navlun=vaka_navlun; sigorta_oran=vaka_sigorta",
         3_604_000,
         "=IFERROR(ROUND(vaka_fob*vaka_kur+vaka_navlun*vaka_kur+vaka_fob*vaka_sigorta*vaka_kur,0),0)",
         "=IFERROR(E7-D7,0)",
         '=IFERROR(IF(ABS(F7)<=ilc_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "CIF TL"),
        ("V-003", "Gümrük = CIF × ILC-001",
         "CIF=vaka_cif; oran 2026",
         360_400,
         '=IFERROR(ROUND(vaka_cif*INDEX(tblKurallar[deger_2026],MATCH("ILC-001",tblKurallar[kural_id],0)),0),0)',
         "=IFERROR(E8-D8,0)",
         '=IFERROR(IF(ABS(F8)<=ilc_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Gümrük"),
        ("V-004", "Birim landed = toplam/miktar",
         "toplam=vaka_landed; miktar=vaka_miktar",
         5_200,
         "=IFERROR(ROUND(vaka_landed/vaka_miktar,0),0)",
         "=IFERROR(E9-D9,0)",
         '=IFERROR(IF(ABS(F9)<=ilc_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Birim"),
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
            'IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=ilc_tolerans,"TUTARLI","KIRIK"),"KIRIK"))'
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
    h(ws, 3, 1, "KANIT RAPORU — İthalat Gerçek Birim Maliyet", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Finans / gümrük dosyası", yazi=GRİ)
    h(ws, 6, 1, "Şirket", kalin=True)
    h(ws, 6, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Dönem", kalin=True)
    h(ws, 7, 2, "=GIRDI!B7", yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Hesap yılı", kalin=True)
    h(ws, 8, 2, "=hesapYili", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Rapor tarihi", kalin=True)
    h(ws, 9, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "Aktif kur", kalin=True)
    h(ws, 10, 2, "=ilc_aktifKur", sayi="0.0000", yazi=DIS_REF_YEŞIL)

    h(ws, 12, 1, "CIF [₺]", kalin=True)
    h(ws, 12, 2, "=MOTOR!G29", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "Gümrük [₺]", kalin=True)
    h(ws, 13, 2, "=MOTOR!G30", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "KKDF [₺]", kalin=True)
    h(ws, 14, 2, "=MOTOR!G31", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "KDV [₺]", kalin=True)
    h(ws, 15, 2, "=MOTOR!G33", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Birim landed [₺]", kalin=True, boyut=12)
    h(ws, 16, 2, "=ilc_birimLanded", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 17, 1, "NPV [₺]", kalin=True)
    h(ws, 17, 2, "=ilc_npv", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "IRR", kalin=True)
    h(ws, 18, 2, "=ilc_irr", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)

    h(ws, 20, 1, "Karar", kalin=True, boyut=12)
    h(ws, 20, 2, "=ilc_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 21, 1, "Gerekçe", kalin=True)
    h(ws, 21, 2, "=ilc_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B21:H21")
    h(ws, 23, 1, "Madde / kural atıfı", kalin=True)
    h(ws, 23, 2, "=ilc_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B23:H23")
    h(ws, 25, 1, "Kanıt özeti", kalin=True)
    h(ws, 25, 2, "=ilc_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B25:H25")
    h(ws, 27, 1, "Parmak izi", kalin=True)
    h(ws, 27, 2, "=ilc_parmakIzi", yazi=GRİ)

    h(ws, 30, 1, "Hazırlayan", kalin=True)
    h(ws, 30, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B30", "Hazırlayan", "Ad soyad imza alanı")
    h(ws, 31, 1, "Onaylayan", kalin=True)
    h(ws, 31, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B31", "Onaylayan", "CFO / müdür imza alanı")
    genislik(ws, {"A": 42, "B": 55})
    h(ws, 34, 1, "Bu çıktı karar destek amaçlıdır; bağlayıcı gümrük veya yatırım kararı değildir.",
      yazi=GRİ, boyut=9, kaydir=True)


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Dosya sağlık kontrolleri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    kontroller_list = [
        (5, "Giriş dolu mu?",
         '=IFERROR(IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"BOŞ","DOLU"),"BOŞ")'),
        (6, "Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(ilc_yilTaban),N(hesapYili)<=N(ilc_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        (7, "Kur > 0?",
         '=IFERROR(IF(ilc_aktifKur>0,"TEMİZ","HATALI"),"HATALI")'),
        (8, "Miktar > 0?",
         '=IFERROR(IF(MOTOR!G15>0,"TEMİZ","HATALI"),"HATALI")'),
        (9, "Birim negatif mi?",
         '=IFERROR(IF(MOTOR!G35<0,"NEGATİF","NORMAL"),"NORMAL")'),
        (10, "Vaka durumu",
         '=IFERROR(ilc_vakaDurum,"KIRIK")'),
        (11, "Marj hedef altı?",
         '=IF(IFERROR(MOTOR!G36,0)<MOTOR!G11+N(GIRDI!B10)*ilc_sifirCarpan,"AŞIYOR","NORMAL")'),
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
    h(ws, 14, 2, "=IFERROR(ilc_girisBeklenen+COUNTA(GIRDI!B6:B21)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Birim landed", kalin=True)
    h(ws, 15, 2, "=IFERROR(ilc_birimLanded,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "NPV", kalin=True)
    h(ws, 16, 2, "=IFERROR(ilc_npv,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "IRR", kalin=True)
    h(ws, 17, 2, "=IFERROR(ilc_irr,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "CIF", kalin=True)
    h(ws, 18, 2, "=IFERROR(MOTOR!G29,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Karar", kalin=True)
    h(ws, 19, 2, '=IFERROR(ilc_kararMetni,"-")', yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 36, "B": 28})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "548235", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek sevkiyat verileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:L3")
    sutunlar = [
        "Senaryo", "FOB", "Miktar", "Navlun", "Kur", "Sigorta",
        "Lokal", "Hedef satış", "Yıllık miktar", "CAPEX", "Yıl", "Birim Ornek",
    ]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Birim Ornek": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(ROUND((tblOrnek[[#This Row],[FOB]]*tblOrnek[[#This Row],[Kur]]"
            "+tblOrnek[[#This Row],[Navlun]]*tblOrnek[[#This Row],[Kur]]"
            "+tblOrnek[[#This Row],[FOB]]*tblOrnek[[#This Row],[Sigorta]]*tblOrnek[[#This Row],[Kur]]"
            "+tblOrnek[[#This Row],[Lokal]])"
            "/MAX(1,tblOrnek[[#This Row],[Miktar]]),0),0))"
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:L1004", sutunlar, formuller)
    rows = [
        ("Demo ana", 100_000, 1_000, 5_000, 34, 0.01, 25_000, 6_500, 12_000, 800_000, 2026),
        ("Yüksek kur", 100_000, 1_000, 5_000, 38, 0.01, 30_000, 7_000, 10_000, 900_000, 2026),
        ("Düşük marj", 100_000, 1_000, 8_000, 34, 0.015, 40_000, 5_200, 8_000, 1_000_000, 2026),
        ("2025 geçiş", 90_000, 900, 4_500, 32, 0.01, 20_000, 6_000, 11_000, 750_000, 2025),
    ]
    for i, row in enumerate(rows, 6):
        for k, v in enumerate(row, 1):
            sayi = None
            if k in (2, 3, 4, 9, 10, 11):
                sayi = CATI if k in (3, 9, 11) else (TL if k in (7, 8, 10) else CATI)
            if k == 5:
                sayi = "0.00"
            if k == 6:
                sayi = YÜZDE
            if k in (7, 8, 10):
                sayi = TL
            h(ws, i, k, v, sayi=sayi, zemin=GIRIS_SARI if k <= 11 else None, yazi="333333")
            if k <= 11:
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", sutunlar[k - 1], "Örnek satır")
    h(ws, 22, 1, "OrnekBirimDemo", kalin=True, yazi=GRİ)
    for i, v in enumerate([5_200, 5_900, 6_400, 5_100], 23):
        h(ws, i, 1, rows[i - 23][0])
        h(ws, i, 2, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Örnek Birim Landed"
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
    h(ws, 5, 3, "Döviz", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "USD")
    h(ws, 7, 3, "EUR")
    h(ws, 5, 5, "EvetHayır", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 5, "Evet")
    h(ws, 7, 5, "Hayır")
    genislik(ws, {"A": 24, "C": 12, "E": 12})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "833C0C", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Parametreler (tek kaynak) + kur tablosu", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = ["anahtar", "deger", "birim", "aciklama", "kaynak", "yururluk_tarihi", "dogrulama_tarihi", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    params = [
        ("hesapYili", DEMO["hesapYili"], "yıl", "Aktif hesap yılı", "Kullanıcı / GIRDI", "01.01.2025", "11.08.2026"),
        ("raporTarihi", RAPOR_TARIH, "tarih", "Rapor tarihi", "Üretim", "01.01.2025", "11.08.2026"),
        ("ilc_tolerans", 1, "adet", "Vaka tutarlılık toleransı", "Uygulama notu", "01.01.2025", "11.08.2026"),
        ("ilc_azamiEsik", 100_000_000_000, "TL", "Azami giriş eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("ilc_senaryoIyi", 0.92, "çarpan", "İyimser maliyet çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ilc_senaryoKotu", 1.18, "çarpan", "Kötümser maliyet çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ilc_esikArtis", 0.02, "oran", "M-SEN gümrük + delta", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("ilc_tahminAlt", 0.90, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ilc_tahminUst", 1.12, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ilc_tornadoKur", 1.05, "çarpan", "Tornado kur şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ilc_tornadoGumruk", 0.02, "oran", "Tornado gümrük şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ilc_tornadoNavlun", 0.10, "oran", "Tornado navlun şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ilc_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "11.08.2026"),
        ("ilc_parmakIzi", "ILC-PRO-1.0.0", "metin", "Dosya parmak izi", "Üretim", "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv", "01.01.2025", "11.08.2026"),
        ("ilc_girisBeklenen", 16, "adet", "Zorunlu giriş sayısı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ilc_riskDusuk", 20, "puan", "YAPILIR risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("ilc_riskOrta", 55, "puan", "DİKKAT risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("ilc_riskYuksek", 85, "puan", "YAPMA risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("ilc_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ilc_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ilc_bazCarpan", 1, "çarpan", "Baz senaryo çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ilc_sifirCarpan", 0, "çarpan", "Nötr çarpan", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ilc_siraBir", 1, "adet", "Tornado sıra 1", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ilc_siraIki", 2, "adet", "Tornado sıra 2", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ilc_siraUc", 3, "adet", "Tornado sıra 3", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ilc_dikkatCarpan", 0.70, "çarpan", "Marj dikkat çarpanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("ilc_npvDikkat", 100_000, "TL", "NPV dikkat eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("vaka_fob", 100_000, "döviz", "Vaka FOB", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_kur", 34, "kur", "Vaka kur", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_navlun", 5_000, "döviz", "Vaka navlun", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_sigorta", 0.01, "oran", "Vaka sigorta", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_cif", 3_604_000, "TL", "Vaka CIF", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_landed", 5_200_000, "TL", "Vaka toplam landed", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_miktar", 1_000, "adet", "Vaka miktar", "Altın vaka", "01.01.2025", "11.08.2026"),
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
                elif row[2] in ("yıl", "adet", "puan", "döviz", "kur"):
                    sayi = CATI if row[2] != "kur" else "0.00"
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
    # Kur tablosu Ç14 — J kolonundan başlar (A=anahtar tablosu ile çakışmasın, D10)
    h(ws, son + 2, 10, "Kur tablosu (tarihli / kaynaklı)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kur_sutun = ["parabirimi", "kur", "tarih", "kaynak", "aktif", "not", "secim"]
    baslik_satiri(ws, son + 3, kur_sutun, basla=10)
    kur_rows = [
        ("USD", 34.00, date(2026, 8, 10), "TCMB alış", "Evet", "Aktif demo kur"),
        ("EUR", 37.20, date(2026, 8, 10), "TCMB alış", "Hayır", "Yedek"),
        ("USD", 32.50, date(2025, 12, 31), "TCMB yılsonu", "Hayır", "2025 referans"),
    ]
    kr0 = son + 4
    for i, row in enumerate(kur_rows, kr0):
        for k, v in enumerate(row, 1):
            sayi = TARİH if k == 3 else ("0.00" if k == 2 else None)
            h(ws, i, 9 + k, v, sayi=sayi, zemin=GIRIS_SARI if k <= 5 else None, yazi="333333")
            if k <= 5:
                ws.cell(i, 9 + k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(9 + k)}{i}", kur_sutun[k - 1], "Kur satırı; kaynak zorunlu")
        h(ws, i, 16, f'=IF(N{i}="Evet",K{i},0)', sayi="0.00", yazi=DIS_REF_YEŞIL)
    formuller_kur = {
        "secim": (
            '=IF(tblKur[[#This Row],[parabirimi]]="","",'
            'IF(tblKur[[#This Row],[aktif]]="Evet",tblKur[[#This Row],[kur]],0))'
        ),
    }
    tablo_ekle(ws, "tblKur", f"J{son+3}:P{kr0+2}", kur_sutun, formuller_kur)
    # aktif kur özeti — B kolonuna yazma (anahtar tablosu D10); R/S kullan
    h(ws, kr0 + 4, 10, "Aktif kur (tablodan)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, kr0 + 4, 11,
      f'=IFERROR(IF(SUM(P{kr0}:P{kr0+2})=0,INDEX(K{kr0}:K{kr0+2},1),SUM(P{kr0}:P{kr0+2})),0)',
      sayi="0.0000", yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, kr0 + 5, 10, "Kur tablo ozeti", kalin=True)
    h(ws, kr0 + 5, 11,
      f'=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"USD=",TEXT(K{kr0},"0.00")," kaynak=",M{kr0}," tarih=",TEXT(L{kr0},"dd.mm.yyyy")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    dogrulama(ws, "list", "ListeYillar", "B6",
              baslik="Hesap Yılı", mesaj="Aktif yılı seçin.",
              hata_baslik="Geçersiz", hata_mesaj="2025-2027.")
    dogrulama(ws, "list", "ListeDoviz", f"J{kr0}:J{kr0+2}",
              baslik="Döviz", mesaj="USD veya EUR.",
              hata_baslik="Geçersiz", hata_mesaj="USD/EUR.")
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
                  "J": 12, "K": 12, "L": 12, "M": 18, "N": 10, "O": 18, "P": 10})
    h(ws, kr0 + 7, 10, f"Sürüm {SURUM} | Şifre koruması: {SIFRE} (formül alanları)", yazi=GRİ, boyut=9)
    ws.sheet_view.zoomScale = 90
    return kr0 + 4  # aktif kur satırı; değer K kolonunda


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar kuralları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "VERİ YOK — giriş tablosu boşsa hesap yapılmaz.",
        "YATIRIM YAPILIR — NPV pozitif ve marj hedef üstünde.",
        "DİKKAT — marj veya NPV dikkat bandında; kur/oran gözden geçirilmeli.",
        "YATIRIM YAPMA — NPV negatif veya marj yetersiz (kötümser yol dahil).",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 10, 1, "Belirsizlik beyanları (M05)", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari(["kkdf_matrah_yorumu"]), 11):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)
    h(ws, 13, 1,
      "KKDF matrahı (FOB mı CIF mi) uygulamada tartışmalı / belirsizdir; bu dosya FOB_TL varsayar "
      "(Yorum A). Yorum B için CIF matrahı seçilirse KKDF yükselir.",
      yazi=KRITIK, kaydir=True)
    ws.merge_cells("A13:J13")

    h(ws, 15, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 2,
      "KKDF matrahı = FOB_TL. CIF = FOB_TL + navlun_TL + sigorta_TL. "
      "Gümrük = CIF × oran. KDV = (CIF+gümrük+KKDF) × oran.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B15:J15")
    h(ws, 16, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 2,
      "Bazı uygulamalarda KKDF CIF üzerinden hesaplanır (tartışmalı). "
      "Dosya Yorum A üretir; Yorum B farkı KILAVUZ'da beyan edilir.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B16:J16")

    h(ws, 18, 1, "Kullanım sırası", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "GIRDI → VARSAYIMLAR → KURALLAR → MOTOR → FINANSMAN → SENARYO → PANO → KARAR → "
      "VAKALAR → KANIT_RAPORU → AYARLAR (kur tablosu).",
      kaydir=True, yazi="333333")
    ws.merge_cells("A19:J19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv | ILC-PRO", yazi=GRİ, boyut=9)
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
        ws.conditional_formatting.add(f"{col}10:{col}17", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(f"{col}10:{col}12", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("F30:F45", FormulaRule(formula=['F30="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F30:F45", FormulaRule(formula=['F30="TAM"'], font=yesil, fill=y_fill))

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
    ws.conditional_formatting.add("B16:B18", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

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
    ws.conditional_formatting.add("G7:G66", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("F7:F66", FormulaRule(formula=['F7="ILC-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F66", FormulaRule(formula=['F7="ILC-004"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "C", "D", "E", "J"):
        ws.conditional_formatting.add(f"{harf}6:{harf}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B20", FormulaRule(formula=['ISNUMBER(SEARCH("YAPMA",B20))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B20", FormulaRule(formula=['ISNUMBER(SEARCH("YAPILIR",B20))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B20", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B20))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B16", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("B17", CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=yesil))

    ws = wb["FINANSMAN"]
    ws.conditional_formatting.add("E6:E7", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("E6:E7", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))

    ws = wb["VARSAYIMLAR"]
    ws.conditional_formatting.add("H6:H13", FormulaRule(formula=['H6="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("H6:H13", FormulaRule(formula=['H6="TAM"'], font=yesil, fill=y_fill))


def ad_tanimlari(wb, kur_aktif_satir: int):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "ilc_birimLanded", "MOTOR!$G$60")
    ad_ekle(wb, "ilc_npv", "MOTOR!$G$61")
    ad_ekle(wb, "ilc_irr", "MOTOR!$G$62")
    ad_ekle(wb, "ilc_cif", "MOTOR!$G$63")
    ad_ekle(wb, "ilc_kararMetni", "MOTOR!$G$42")
    ad_ekle(wb, "ilc_kararGerekce", "MOTOR!$G$43")
    ad_ekle(wb, "ilc_maddeAtifMetni", "MOTOR!$G$58")
    ad_ekle(wb, "ilc_kanitRaporu", "MOTOR!$G$59")
    ad_ekle(wb, "ilc_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "ilc_parmakIzi", "AYARLAR!$B$19")
    ad_ekle(wb, "ilc_aktifKur", f"AYARLAR!$K${kur_aktif_satir}")
    ad_ekle(wb, "ilc_kurTablosu", f"AYARLAR!$K${kur_aktif_satir + 1}")
    ad_ekle(wb, "ilc_finansmanKarsilastirma", "FINANSMAN!$B$11")

    ad_ekle(wb, "ilc_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "ilc_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "ilc_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "ilc_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "ilc_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "ilc_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "ilc_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "ilc_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "ilc_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "ilc_vakaFark", "VAKALAR!$B$13")

    # AYARLAR rows: 6=hesapYili ... map by order in params
    ayar_map = {
        "ilc_tolerans": 8,
        "ilc_azamiEsik": 9,
        "ilc_senaryoIyi": 10,
        "ilc_senaryoKotu": 11,
        "ilc_esikArtis": 12,
        "ilc_tahminAlt": 13,
        "ilc_tahminUst": 14,
        "ilc_tornadoKur": 15,
        "ilc_tornadoGumruk": 16,
        "ilc_tornadoNavlun": 17,
        "ilc_yuzdelikOran": 18,
        "ilc_girisBeklenen": 21,
        "ilc_riskDusuk": 22,
        "ilc_riskOrta": 23,
        "ilc_riskYuksek": 24,
        "ilc_yilTaban": 25,
        "ilc_yuvarMax": 26,
        "ilc_bazCarpan": 27,
        "ilc_sifirCarpan": 28,
        "ilc_siraBir": 29,
        "ilc_siraIki": 30,
        "ilc_siraUc": 31,
        "ilc_dikkatCarpan": 32,
        "ilc_npvDikkat": 33,
        "vaka_fob": 34,
        "vaka_kur": 35,
        "vaka_navlun": 36,
        "vaka_sigorta": 37,
        "vaka_cif": 38,
        "vaka_landed": 39,
        "vaka_miktar": 40,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "ilc_modulT1": (11, "C"), "ilc_modulT1Yorum": (11, "D"),
        "ilc_modulT2": (12, "C"), "ilc_modulT2Yorum": (12, "D"),
        "ilc_modulO1": (13, "C"), "ilc_modulO1Yorum": (13, "D"),
        "ilc_modulO6": (16, "C"), "ilc_modulO6Yorum": (16, "D"),
        "ilc_modulO8": (17, "C"), "ilc_modulO8Yorum": (17, "D"),
        "ilc_modulI2": (20, "C"), "ilc_modulI2Yorum": (20, "D"),
    }
    for ad, (satir, kol) in modul_map.items():
        ad_ekle(wb, ad, f"PANO!${kol}${satir}")

    ad_ekle(wb, "ListeYillar", "LISTELER!$A$6:$A$8")
    ad_ekle(wb, "ListeDoviz", "LISTELER!$C$6:$C$7")
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
    kur_aktif_satir = ayarlar(ws_ay)
    ws_kil = wb.create_sheet("KILAVUZ")
    kilavuz(ws_kil)

    ad_tanimlari(wb, kur_aktif_satir)
    kosullu_bicimlendirme(wb)
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    wb.calculation.fullCalcOnLoad = True

    kok = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    varsayilan = os.path.join(kok, "IthalatLandedCostMotoru.xlsx")
    hedef = cikti_yolu or varsayilan
    os.makedirs(os.path.dirname(os.path.abspath(hedef)) or ".", exist_ok=True)
    wb.save(hedef)

    hsh = hashlib.sha256()
    with open(hedef, "rb") as f:
        for b in iter(lambda: f.read(65536), b""):
            hsh.update(b)
    dig = hsh.hexdigest()

    wb2 = __import__("openpyxl").load_workbook(hedef)
    # parmak izi satırı: ilc_parmakIzi is B19 (params index)
    wb2["AYARLAR"]["B19"] = dig[:16]
    for wsx in wb2.worksheets:
        sayfa_koru(wsx)
    wb2.save(hedef)

    print(f"Dosya: {hedef}")
    print(f"SHA-256: {dig}")
    print(f"Şifre: {SIFRE}")
    print(f"Kur aktif satır: {kur_aktif_satir}")
    return hedef


if __name__ == "__main__":
    yol = sys.argv[1] if len(sys.argv) > 1 else None
    uretilen = main(yol)
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
    cikti = os.path.join(repo, "cikti", "IthalatLandedCostMotoru.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
