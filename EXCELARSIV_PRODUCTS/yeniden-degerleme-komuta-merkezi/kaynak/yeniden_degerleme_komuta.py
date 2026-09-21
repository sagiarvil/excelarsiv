#!/usr/bin/env python3
"""
Yeniden Değerleme Komuta Merkezi (VUK 298/Ç + Geçici 32) — üretim betiği (A1 / manda v6).
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

URUN_AD = "Yeniden Değerleme Komuta Merkezi (298/Ç + Geç.32)"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 11)

# Demo: maliyet 1.000.000 × endeks 1,45 → fark 450.000 (vaka YD-V01)
DEMO = {
    "kurum_adi": "Örnek Sanayi A.Ş.",
    "donem": "2025",
    "hesapYili": 2025,
    "maliyet": 1_000_000,
    "birikmis_amortisman": 200_000,
    "endeks_orani": 1.45,
    "kalan_omur_yil": 10,
    "oncelik_yorumu": "A",
}



def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: Yeniden değerleme kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Fark ve karar bozulur | Kaynak: [Assumption] VUK 298/Ç · Geçici 32",
    )


def kural_cek(kural_id: str) -> str:
    """Oranları tblKurallar'dan çeker; yıl seçimi ydk_yilTaban ile (G07: adında rakam yok)."""
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-ydk_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


# Yuvarlama basamağı uç değerde (#VALUE!) kırmasın — tavan AYARLAR'dan (G07)
_YV = "MAX(0,MIN(ydk_yuvarMax,IFERROR(N(G12),0)))"


def _guvenli_formul(formul: str, birim: str) -> str:
    """Ö2 uç durumda MOTOR hücrelerinde hata değeri göstermemek için IFERROR."""
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
      "Sabit kıymet maliyet ve endeksini girin; VUK 298/Ç ile Geçici 32 yeniden değerleme "
      "farkını, fon/özkaynak ve vergi etkisini yan yana üretip GO/REVIEW kararı verin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Çok yıllı kural tablosu (2025/2026/2027) ile endeks/fon güncellemesine hazır motor",
        "298/Ç ve Geç.32 yan yana fark, fon, özkaynak ve vergi etkisi",
        "Baz / iyimser / kötümser + kural değişimi (fon +1 puan) senaryoları",
        "Tebliğ tarzı altın vakalar ve KANIT_RAPORU (A4 imza)",
        "Belirsizlik beyanı: Geç.32 vs 298/Ç öncelik (yorum A / B)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "SMMM / YMM — 298/Ç vs Geç.32 kararını müşteriye sunmak",
        "CFO — vergi/özkaynak etkisini bütçelemek",
        "Ortak / yönetim — yap/yapma kararını almak",
    ], 14):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) HIZLI_BASLANGIC → 2) GIRDI sarı hücreler → 3) PANO karar → 4) KANIT_RAPORU yazdır.",
      yazi="333333", kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 21, 1, f"Sürüm {SURUM} | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 72})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Üç adımda sonuç", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Girdiler",
         "GIRDI sayfasında maliyet, birikmiş amortisman, endeks ve kalan ömrü sarı hücrelere girin."),
        ("Adım 2 — Karar",
         "PANO'da fark, fon, özkaynak etkisi ve GO/REVIEW rozetini izleyin."),
        ("Adım 3 — Kanıt",
         "KANIT_RAPORU'nu yazdırıp imzalayın; VAKALAR tutarlılığını kontrol edin."),
    ]
    for i, (b, m) in enumerate(adimlar, 5):
        h(ws, i, 1, b, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        _kim(ws, f"A{i}", b, "Sırayla izleyin")
    h(ws, 9, 1, "Sık yapılan hatalar", kalin=True, yazi=KRITIK)
    for i, m in enumerate([
        "Endeksi yüzde yerine çarpan olarak girmemek (1,45 doğru; 45 yanlış)",
        "Net defter yerine brüt maliyeti karıştırmak",
        "Geç.32 fon oranını kural tablosu dışında sabitlemek",
    ], 10):
        h(ws, i, 1, "• " + m, yazi=KRITIK)
    h(ws, 14, 1,
      "Bu dosya karar destek aracıdır; beyanname veya mali müşavir görüşü yerine geçmez.",
      yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 72, "B": 70})
    for r in (10, 11, 12):
        h(ws, r, 1, ws.cell(r, 1).value, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Dönem ve kıymet", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Fark ve karar formülle üretilir.", yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Kurum Adı", DEMO["kurum_adi"], None, "kurum_adi",
         "Kurumun resmi unvanını yazın."),
        (7, "Dönem", DEMO["donem"], None, "donem",
         "Vergi dönemi etiketini yazın (ör. 2025)."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili",
         "2025, 2026 veya 2027 seçin."),
        (9, "Maliyet (iktisap) [₺]", DEMO["maliyet"], TL, "maliyet",
         "Sabit kıymetin iktisap maliyetini TL girin."),
        (10, "Birikmiş Amortisman [₺]", DEMO["birikmis_amortisman"], TL,
         "birikmis_amortisman", "Birikmiş amortisman tutarını TL girin."),
        (11, "Endeks Oranı [çarpan]", DEMO["endeks_orani"], None,
         "endeks_orani", "Yeniden değerleme endeks çarpanını girin (ör. 1,45)."),
        (12, "Kalan Ömür [yıl]", DEMO["kalan_omur_yil"], CATI, "kalan_omur_yil",
         "Kalan ekonomik ömrü yıl olarak girin."),
        (13, "Öncelik Yorumu (A/B)", DEMO["oncelik_yorumu"], None,
         "oncelik_yorumu", "A=298/Ç öncelik, B=Geç.32 öncelik."),
    ]
    h(ws, 5, 1, "Alan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 2, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 3, "Birim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    birim_map = {
        "hesapYili": "yıl", "kalan_omur_yil": "yıl", "endeks_orani": "çarpan",
        "oncelik_yorumu": "A/B", "kurum_adi": "metin", "donem": "metin",
    }
    for satir, etiket, deger, sayi, _ad, mesaj in alanlar:
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim_map.get(_ad, "₺"), yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)
    ws.cell(11, 2).number_format = "0.00"

    h(ws, 15, 1, "Net Defter Değeri [₺]", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 2, "=MAX(0,B9-B10)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 16, 1, "Değer Artış Farkı [₺]", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 2, "=IFERROR(ROUND(MAX(0,B9*(B11-1)),2),0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 17, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B13)=0,ydk_girisBeklenen=0),0,'
      'ROUND(COUNTA(B6:B13)/ydk_girisBeklenen*100,0)),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)

    h(ws, 20, 1, "Giriş tablosu (otomasyon / boş-mod)", kalin=True, yazi=KOYU_LACIVERT)
    sutunlar = ["Alan Anahtarı", "Değer", "Birim", "Zorunlu", "Kaynak", "Kontrol"]
    baslik_satiri(ws, 21, sutunlar)
    formuller = {
        "Kontrol": (
            '=IF(tblGirdi[[#This Row],[Alan Anahtarı]]="","",'
            'IF(tblGirdi[[#This Row],[Değer]]="","EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblGirdi", "A21:F1020", sutunlar, formuller)
    ornek_satir = [
        ("kurum_adi", DEMO["kurum_adi"], "metin", "Evet", "GIRDI"),
        ("donem", DEMO["donem"], "metin", "Evet", "GIRDI"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Evet", "GIRDI"),
        ("maliyet", DEMO["maliyet"], "TL", "Evet", "GIRDI"),
        ("birikmis_amortisman", DEMO["birikmis_amortisman"], "TL", "Evet", "GIRDI"),
        ("endeks_orani", DEMO["endeks_orani"], "çarpan", "Evet", "GIRDI"),
        ("kalan_omur_yil", DEMO["kalan_omur_yil"], "yıl", "Evet", "GIRDI"),
        ("oncelik_yorumu", DEMO["oncelik_yorumu"], "A/B", "Evet", "GIRDI"),
    ]
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 22):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in ("maliyet", "birikmis_amortisman"):
            ws.cell(i, 2).number_format = TL
        elif a == "endeks_orani":
            ws.cell(i, 2).number_format = "0.00"
        elif a in ("hesapYili", "kalan_omur_yil"):
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
    dogrulama(ws, "list", "ListeYorumAB", "B13",
              baslik="Öncelik", mesaj="A=298/Ç veya B=Geç.32 seçin.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B seçin.", bos=False)
    dogrulama(ws, "list", "ListeYillar", "B24",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeYorumAB", "B29",
              baslik="Öncelik", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.")
    for aralik, baslik in [
        ("B9", "Maliyet"), ("B10", "Birikmiş"), ("B25", "Maliyet"), ("B26", "Birikmiş"),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj="0 ile 1e12 arasında TL.",
                  hata_baslik="Geçersiz tutar", hata_mesaj="0 ile 1e12 arasında olmalı.",
                  isaret="between", f2="1000000000000")
    for aralik in ("B11", "B27"):
        dogrulama(ws, "decimal", "1", aralik,
                  baslik="Endeks", mesaj="1 ile 10 arasında çarpan.",
                  hata_baslik="Geçersiz endeks", hata_mesaj="1 ile 10 arasında olmalı.",
                  isaret="between", f2="10")
    for aralik in ("B12", "B28"):
        dogrulama(ws, "whole", "1", aralik,
                  baslik="Kalan ömür", mesaj="1 ile 100 arasında yıl.",
                  hata_baslik="Geçersiz ömür", hata_mesaj="1 ile 100 arasında olmalı.",
                  isaret="between", f2="100")

    giris_hucreleri(ws, 6, 13, [2])
    giris_hucreleri(ws, 22, 29, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 36, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1022, "Sarı alanlar giriş; fark ve kalite formülleri kilitlidir.")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Mevzuat kural tablosu (yıl yan yana)", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["YDK-001", "VUK 298/Ç", "endeks_taban", "endeks>=taban", "endeks_taban",
         1.0, 1.0, 1.0, "01.01.2025", "VUK md.298/Ç"],
        ["YDK-002", "Geçici 32", "gec32_fon_orani", "fark>0", "gec32_fon_orani",
         0.02, 0.02, 0.02, "01.01.2025", "VUK Geçici md.32"],
        ["YDK-003", "KVK md.32", "kv_oran", "vergi etkisi", "kv_oran",
         0.25, 0.25, 0.25, "01.01.2025", "KVK md.32"],
        ["YDK-004", "VUK 298/Ç", "c298_fon_orani", "298 yolu", "c298_fon_orani",
         0.0, 0.0, 0.0, "01.01.2025", "VUK md.298/Ç"],
        ["YDK-005", "Genel", "yuvarlama", "her hesap", "yuvarlama_ondalik",
         2, 2, 2, "01.01.2025", "Uygulama notu"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = None
            if k in (6, 7, 8):
                if s[0] == "YDK-005":
                    sayi = CATI
                elif s[0] == "YDK-001":
                    sayi = "0.00"
                elif s[0] in ("YDK-002", "YDK-003", "YDK-004"):
                    sayi = YÜZDE
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-ydk_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Geç.32 fon oranı (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("YDK-002"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "KV oranı (aktif yıl)", yazi="333333")
    h(ws, 16, 2, kural_cek("YDK-003"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=_xlfn.TEXTJOIN(" | ",TRUE,"YDK-002=",TEXT(B15,"0.0%"),"YDK-003=",TEXT(B16,"0.0%"),"yıl=",hesapYili)',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    for col in ("F", "G", "H"):
        dogrulama(ws, "decimal", "0", f"{col}7:{col}11",
                  baslik="Kural değeri", mesaj="Yıl kolonuna sayısal parametre girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ile 10 arasında.",
                  isaret="between", f2="10")
    genislik(ws, {get_column_letter(i): w for i, w in enumerate(
        [14, 14, 22, 22, 20, 12, 12, 12, 12, 18, 14], 1)})
    ws.merge_cells("A3:K3")
    ws.merge_cells("A4:K4")
    sabitle(ws, "A7")


def motor(ws):
    sayfa_hazirla(ws, "MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Hesap zinciri — her adımda kural_id", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar tblKurallar'dan çekilir; sabit oran yoktur.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")

    basliklar = [*MOTOR_BASLIK, "deger"]
    baslik_satiri(ws, 6, basliklar)

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    # G7 yil … G20 fark … ≥40 adım
    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif mevzuat yılı", "YDK-005")  # G7
    ekle("Endeks taban", km("YDK-001"), "çarpan", "Minimum endeks", "YDK-001")  # G8
    ekle("Geç.32 fon oranı", km("YDK-002"), "oran", "Fon vergi oranı", "YDK-002")  # G9
    ekle("KV oranı", km("YDK-003"), "oran", "Kurumlar vergisi oranı", "YDK-003")  # G10
    ekle("298/Ç fon oranı", km("YDK-004"), "oran", "298 yolu fon", "YDK-004")  # G11
    ekle("Yuvarlama ondalık", km("YDK-005"), "adet", "Yuvarlama basamağı", "YDK-005")  # G12
    ekle("Maliyet", "=GIRDI!B9", "TL", "İktisap maliyeti", "YDK-001")  # G13
    ekle("Birikmiş amortisman", "=GIRDI!B10", "TL", "Birikmiş amortisman", "YDK-001")  # G14
    ekle("Endeks oranı", "=GIRDI!B11", "çarpan", "Kullanıcı endeksi", "YDK-001")  # G15
    ekle("Kalan ömür", "=GIRDI!B12", "yıl", "Kalan ekonomik ömür", "YDK-003")  # G16
    ekle("Net defter değeri", "=MAX(0,G13-G14)", "TL", "Maliyet − birikmiş", "YDK-001")  # G17
    ekle("Endeks düzeltilmiş", "=MAX(G15,G8)", "çarpan", "Taban ile max", "YDK-001")  # G18
    ekle("Yeniden değerlenmiş maliyet", f"=ROUND(G13*G18,{_YV})", "TL", "Maliyet × endeks", "YDK-001")  # G19
    ekle("Değer artış farkı", f"=ROUND(MAX(0,G13*(G18-1)),{_YV})", "TL", "Maliyet×(endeks−1)", "YDK-001")  # G20
    ekle("Geç.32 fon vergisi", f"=ROUND(G20*G9,{_YV})", "TL", "Fark × fon oranı", "YDK-002")  # G21
    ekle("298/Ç fon vergisi", f"=ROUND(G20*G11,{_YV})", "TL", "Fark × 298 fon", "YDK-004")  # G22
    ekle("Özkaynak 298/Ç", "=G20-G22", "TL", "Fark − 298 fon", "YDK-004")  # G23
    ekle("Özkaynak Geç.32", "=G20-G21", "TL", "Fark − Geç.32 fon", "YDK-002")  # G24
    ekle("Ek amortisman yıllık", "=IF(G16=0,0,G20/G16)", "TL", "Fark / kalan ömür", "YDK-003")  # G25
    ekle("Vergi tasarrufu brüt", f"=ROUND(G20*G10,{_YV})", "TL", "Fark × KV oranı", "YDK-003")  # G26
    ekle("Net fayda 298/Ç", "=G23", "TL", "298 özkaynak net", "YDK-004")  # G27
    ekle("Net fayda Geç.32", "=G24+G26*ydk_tasarrufCarpan", "TL", "Özkaynak+vergi faydası", "YDK-002")  # G28
    ekle("Fayda farkı (298−Geç32)", "=G27-G28", "TL", "Yol karşılaştırması", "YDK-001")  # G29
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(G20<=0,"YOK",'
         'IF(ABS(G29)<=ydk_kararEsik,"REVIEW",'
         'IF(G29>0,"GO_298","GO_GEC32"))))',
         "kod", "Karar kodu", "YDK-001")  # G30
    ekle("Karar metni",
         '=IF(G30="VERI_YOK","VERİ YOK",'
         'IF(G30="YOK","YENİDEN DEĞERLEME YOK / UYGUN DEĞİL",'
         'IF(G30="REVIEW","REVIEW · DİKKAT",'
         'IF(G30="GO_298","GO · 298/Ç UYGUN","GO · GEÇ.32 UYGUN"))))',
         "metin", "Kullanıcıya karar", "YDK-001")  # G31
    ekle("Gerekçe",
         '=IF(G30="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G30="YOK","Endeks artışı yok; yeniden değerleme farkı sıfır.",'
         'IF(G30="REVIEW","298/Ç ve Geç.32 net faydaları eşik içinde; öncelik yorumunu gözden geçirin.",'
         'IF(G30="GO_298","298/Ç net faydası Geç.32 üstünde; fon yükü düşük yol öne çıkar.",'
         '"Geç.32 net faydası (özkaynak+vergi tasarrufu) 298/Ç üstünde."))))',
         "metin", "Gerekçe cümlesi", "YDK-001")  # G32
    ekle("Fark / maliyet", "=IF(G13=0,0,G20/G13)", "oran", "Artış yoğunluğu", "YDK-001")  # G33
    ekle("Fon / fark", "=IF(G20=0,0,G21/G20)", "oran", "Fon yük oranı", "YDK-002")  # G34
    ekle("Yorum A net (298 öncelik)", "=G27", "TL", "A yolu net", "YDK-004")  # G35
    ekle("Yorum B net (Geç.32 öncelik)", "=G28", "TL", "B yolu net", "YDK-002")  # G36
    ekle("Seçili yorum net",
         '=IF(GIRDI!B13="B",G36,G35)',
         "TL", "A/B seçimine göre", "YDK-001")  # G37
    ekle("Yorum farkı", "=ABS(G35-G36)", "TL", "Belirsizlik aralığı", "YDK-001")  # G38
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B13)=0,ydk_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B13)/ydk_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "YDK-005")  # G39
    ekle("Risk skoru",
         '=IF(G30="VERI_YOK",0,IF(G30="YOK",ydk_riskDusuk,'
         'IF(G30="REVIEW",ydk_riskOrta,ydk_riskYuksek)))',
         "puan", "Karar riski", "YDK-001")  # G40
    ekle("Senaryo iyimser maliyet", "=G13*ydk_senaryoIyi", "TL", "İyimser maliyet", "YDK-001")  # G41
    ekle("Senaryo kötümser maliyet", "=G13*ydk_senaryoKotu", "TL", "Kötümser maliyet", "YDK-001")  # G42
    ekle("İyimser fark", f"=ROUND(MAX(0,G41*(G18-1)),{_YV})", "TL", "İyimser fark", "YDK-001")  # G43
    ekle("Kötümser fark", f"=ROUND(MAX(0,G42*(G18-1)),{_YV})", "TL", "Kötümser fark", "YDK-001")  # G44
    ekle("M-SEN fon+1 puan fark net",
         f"=ROUND(G20*(G9+ydk_oranArtisPuan),{_YV})",
         "TL", "Fon +1 puan senaryosu", "YDK-002")  # G45
    ekle("M-SEN özkaynak Geç.32", "=G20-G45", "TL", "M-SEN net özkaynak", "YDK-002")  # G46
    ekle("Tahmin üst", "=G20*ydk_tahminUst", "TL", "Tahmin üst bant", "YDK-005")  # G47
    ekle("Tahmin alt", "=G20*ydk_tahminAlt", "TL", "Tahmin alt bant", "YDK-005")  # G48
    ekle("Tornado: maliyet etkisi",
         f"=ABS(ROUND(MAX(0,(G13*ydk_tornadoMaliyet)*(G18-1)),{_YV})-G20)",
         "TL", "Maliyet ± etki", "YDK-001")  # G49
    ekle("Tornado: endeks etkisi",
         f"=ABS(ROUND(MAX(0,G13*(G18*ydk_tornadoEndeks-1)),{_YV})-G20)",
         "TL", "Endeks ± etki", "YDK-001")  # G50
    ekle("Tornado: fon etkisi",
         f"=ABS(ROUND(G20*(G9+ydk_oranArtisPuan),{_YV})-G21)",
         "TL", "Fon ± etki", "YDK-002")  # G51
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"YDK-001→VUK 298/Ç","YDK-002→Geçici 32","YDK-003→KVK md.32")',
         "metin", "Kanıt atıfları", "YDK-001")  # G52
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Fark=",TEXT(G20,"₺ #,##0"),'
         '"Fon32=",TEXT(G21,"₺ #,##0"),"Özkaynak298=",TEXT(G23,"₺ #,##0"))',
         "metin", "Rapor özeti", "YDK-005")  # G53

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
        elif birim == "oran":
            ws.cell(r, 7).number_format = YÜZDE
            ws.cell(r, 3).number_format = YÜZDE
        elif birim == "çarpan":
            ws.cell(r, 7).number_format = "0.00"
            ws.cell(r, 3).number_format = "0.00"
        elif birim in ("puan", "yıl", "adet", "bayrak"):
            ws.cell(r, 7).number_format = CATI
            ws.cell(r, 3).number_format = CATI

    genislik(ws, {"A": 8, "B": 40, "C": 55, "D": 10, "E": 28, "F": 12, "G": 22})
    sabitle(ws, "A7")
    h(ws, 6 + len(adimlar) + 2, 9,
      "Motor adımları kilitlidir; oranlar yalnızca KURALLAR'dan gelir.", yazi=GRİ, kaydir=True)
    return len(adimlar)


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve duyarlılık", kalin=True, boyut=13, yazi=KOYU_LACIVERT)

    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Maliyet Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Değer Farkı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Geç.32 Fon", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    fon = 'IFERROR(INDEX(tblKurallar[aktif_deger],MATCH("YDK-002",tblKurallar[kural_id],0)),0)'
    senaryolar = [
        ("İyimser", "=IFERROR(ydk_senaryoIyi+N(GIRDI!B9)*0,0)",
         "=IFERROR(ROUND(MAX(0,IFERROR(ydk_maliyet,0)*B6*(IFERROR(GIRDI!B11,1)-1)),2),0)",
         f"=IFERROR(ROUND(C6*{fon},2),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(C6=0,"YOK","GO/REVIEW"))'),
        ("Baz", "=IFERROR(1+N(GIRDI!B9)*0,1)",
         "=IFERROR(ROUND(MAX(0,IFERROR(ydk_maliyet,0)*B7*(IFERROR(GIRDI!B11,1)-1)),2),0)",
         f"=IFERROR(ROUND(C7*{fon},2),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(C7=0,"YOK","GO/REVIEW"))'),
        ("Kötümser", "=IFERROR(ydk_senaryoKotu+N(GIRDI!B9)*0,0)",
         "=IFERROR(ROUND(MAX(0,IFERROR(ydk_maliyet,0)*B8*(IFERROR(GIRDI!B11,1)-1)),2),0)",
         f"=IFERROR(ROUND(C8*{fon},2),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(C8=0,"YOK","GO/REVIEW"))'),
        ("M-SEN fon+1 puan", "=IFERROR(1+N(GIRDI!B9)*0,1)",
         "=IFERROR(ydk_fark,0)",
         "=IFERROR(MOTOR!G45,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(C9=0,"YOK","GO/REVIEW"))'),
    ]
    for i, (ad, carp, fark, fo, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, fark, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, fo, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant genişliği (iyimser−kötümser fark)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(C6-C8,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser fark ",TEXT(C6,"₺ #,##0")," · Baz ",TEXT(C7,"₺ #,##0"),'
      '" · Kötümser ",TEXT(C8,"₺ #,##0")," · Bant ",TEXT(B11,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:H12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Tornado maliyet etkisi", yazi="333333")
    h(ws, 16, 2, "=IFERROR(MOTOR!G49,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado endeks etkisi", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G50,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado fon etkisi", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G51,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 3, '=IFERROR(IF(B16>=MAX(B16:B18),1,IF(B16>=MIN(B16:B18)+ABS(B16-B17),2,3)),3)', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 3, '=IFERROR(IF(B17>=MAX(B16:B18),1,IF(B17=MEDIAN(B16:B18),2,3)),3)', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 3, '=IFERROR(IF(B18>=MAX(B16:B18),1,IF(B18=MEDIAN(B16:B18),2,3)),3)', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2, '=IFERROR(CONCATENATE("1)",INDEX(A16:A18,MATCH(1,C16:C18,0))," 2)",INDEX(A16:A18,MATCH(2,C16:C18,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B16:B18),"₺ #,##0")," TL — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN kural değişimi fon", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=D9", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Fon oranı +1 puan senaryosunda fon ",TEXT(D9,"₺ #,##0")," TL")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı (fark)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=MOTOR!G48", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 3, "=ydk_fark", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 4, "=MOTOR!G47", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Tahmin (FORECAST)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1, "Seri nokta", yazi=GRİ)
    for i, v in enumerate([1, 2, 3, 4], 28):
        h(ws, i, 1, v, sayi=CATI)
        h(ws, i, 2, f"=C{5+v}", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 32, 1, "Tahmin sonraki", yazi="333333")
    h(ws, 32, 2, "=IFERROR(FORECAST.LINEAR(5,B28:B31,A28:A31),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Tahmin yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Tahmin ",TEXT(B32,"₺ #,##0")," · Alt ",TEXT(B25,"₺ #,##0"),'
      '" · Üst ",TEXT(D25,"₺ #,##0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 5, 7, "FarkDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([405000, 450000, 495000, 459000], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([22500, 45000, 4500], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)
    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([405000, 450000, 495000, 459000], 28):
        h(ws, i, 4, v, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo Fark Karşılaştırması"
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
    g3.title = "Senaryo Fark Trendi"
    g3.add_data(Reference(ws, min_col=4, min_row=27, max_row=31), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=28, max_row=31))
    g3.height, g3.width = 8, 12
    ws.add_chart(g3, "F32")

    genislik(ws, {"A": 36, "B": 18, "C": 18, "D": 18, "E": 22, "G": 14})


def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Tebliğ tarzı altın vakalar", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:H3")
    baslik_satiri(ws, 5, VAKA_BASLIK)

    # V1: maliyet 1M × (1.45-1) = 450000
    # V2: maliyet 2M × (1.20-1) = 400000
    # V3: maliyet 500k × (1.10-1) = 50000
    # V4: fon = 450000 * 0.02 = 9000
    vakalar_data = [
        ("YD-V01", "Tebliğ örnek 1 — pozitif fark",
         "maliyet=vaka_maliyet_1; endeks=vaka_endeks_1",
         450000,
         "=IFERROR(ROUND(MAX(0,vaka_maliyet_1*(vaka_endeks_1-INDEX(tblKurallar[deger_2025],MATCH(\"YDK-001\",tblKurallar[kural_id],0)))),2),0)",
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=ydk_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "VUK 298/Ç · Geçici 32 hesap örneği"),
        ("YD-V02", "Tebliğ örnek 2 — orta endeks",
         "maliyet=vaka_maliyet_2; endeks=vaka_endeks_2",
         400000,
         "=IFERROR(ROUND(MAX(0,vaka_maliyet_2*(vaka_endeks_2-INDEX(tblKurallar[deger_2025],MATCH(\"YDK-001\",tblKurallar[kural_id],0)))),2),0)",
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=ydk_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "VUK 298/Ç hesap örneği"),
        ("YD-V03", "Tebliğ örnek 3 — düşük endeks",
         "maliyet=vaka_maliyet_3; endeks=vaka_endeks_3",
         50000,
         "=IFERROR(ROUND(MAX(0,vaka_maliyet_3*(vaka_endeks_3-INDEX(tblKurallar[deger_2025],MATCH(\"YDK-001\",tblKurallar[kural_id],0)))),2),0)",
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=ydk_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "VUK 298/Ç hesap örneği"),
        ("YD-V04", "Geç.32 fon kontrolü",
         "fark=vaka_maliyet_1*(vaka_endeks_1-1); fon oranı YDK-002",
         9000,
         "=IFERROR(ROUND(MAX(0,vaka_maliyet_1*(vaka_endeks_1-1))*INDEX(tblKurallar[deger_2025],MATCH(\"YDK-002\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=ydk_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "VUK Geçici md.32"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v,
              yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = TL

    formuller = {
        "fark": "=IF(tblVakalar[[#This Row],[vaka_id]]=\"\",\"\",IFERROR(tblVakalar[[#This Row],[beklenen_sonuc]]-tblVakalar[[#This Row],[hesaplanan]],0))",
        "durum": '=IF(tblVakalar[[#This Row],[vaka_id]]="","",IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=ydk_tolerans,"TUTARLI","KIRIK"),"KIRIK"))',
    }
    tablo_ekle(ws, "tblVakalar", "A5:H9", VAKA_BASLIK, formuller)

    h(ws, 12, 1, "Vaka durum özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")',
      yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 13, 1, "Vaka fark toplamı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2, "=IFERROR(SUM(tblVakalar[fark]),0)", sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 5, 10, "BeklenenDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 11, "HesapDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, (b, he) in enumerate([(450000, 450000), (400000, 400000), (50000, 50000), (9000, 9000)], 6):
        h(ws, i, 10, b, sayi=TL, yazi=GRİ)
        h(ws, i, 11, he, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Vaka Beklenen vs Hesaplanan"
    g.add_data(Reference(ws, min_col=10, min_row=5, max_col=11, max_row=9), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=6, max_row=9))
    g.height, g.width = 8, 14
    ws.add_chart(g, "A15")

    genislik(ws, {"A": 12, "B": 32, "C": 40, "D": 14, "E": 14, "F": 12, "G": 12, "H": 28})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", "0F2742", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "KANIT RAPORU — Yeniden Değerleme (298/Ç + Geç.32)", kalin=True, boyut=14, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:F3")
    h(ws, 4, 1, "Rapor tarihi", yazi="333333")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 5, 1, "Kurum", yazi="333333")
    h(ws, 5, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Dönem / yıl", yazi="333333")
    h(ws, 6, 2, '=CONCATENATE(GIRDI!B7," / ",hesapYili)', yazi=DIS_REF_YEŞIL)

    h(ws, 8, 1, "Girdi özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, (et, form, sayi) in enumerate([
        ("Maliyet", "=GIRDI!B9", TL),
        ("Birikmiş amortisman", "=GIRDI!B10", TL),
        ("Endeks oranı", "=GIRDI!B11", "0.00"),
        ("Kalan ömür", "=GIRDI!B12", CATI),
        ("Değer artış farkı", "=ydk_fark", TL),
        ("Net defter", "=ydk_netDefter", TL),
    ], 9):
        h(ws, i, 1, et, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)

    h(ws, 16, 1, "Hesap zinciri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 17, 1, "Geç.32 fon", yazi="333333")
    h(ws, 17, 2, "=ydk_fonGec32", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 1, "Özkaynak 298/Ç", yazi="333333")
    h(ws, 18, 2, "=ydk_ozKaynak298", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "KARAR", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2, "=ydk_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Gerekçe", yazi="333333")
    h(ws, 20, 2, "=ydk_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B20:F20")

    h(ws, 22, 1, "Madde atıfları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=ydk_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B22:F22")

    h(ws, 24, 1, "Kanıt gövdesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, "=ydk_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B24:F24")

    h(ws, 26, 1, "Parmak izi (SHA-256 kısaltma)", yazi="333333")
    h(ws, 26, 2, "=ydk_parmakIzi", yazi=GRİ)

    h(ws, 28, 1, "Hazırlayan (imza)", yazi="333333")
    h(ws, 28, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B28", "İmza", "Raporu hazırlayan kişinin adını yazın.")
    h(ws, 29, 1, "Onaylayan (imza)", yazi="333333")
    h(ws, 29, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B29", "Onay", "Onaylayan kişinin adını yazın.")

    h(ws, 31, 1,
      "Uyarı: Bu çıktı karar destektir; kesin vergi tarhiyatı veya hukuki görüş değildir. "
      f"Sürüm {SURUM}.",
      yazi=GRİ, boyut=9, kaydir=True)
    ws.merge_cells("A31:F31")

    baski_hazirla(ws, "A1:F32", f"{URUN_AD} · {SURUM}")
    ws.page_setup.orientation = "portrait"
    genislik(ws, {"A": 28, "B": 22, "C": 14, "D": 14, "E": 14, "F": 14})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim paneli — Yeniden Değerleme", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpiler = [
        (2, "Maliyet", "=IFERROR(ydk_maliyet,0)", TL),
        (3, "Fark", "=IFERROR(ydk_fark,0)", TL),
        (4, "Fon Geç.32", "=IFERROR(ydk_fonGec32,0)", TL),
        (5, "Özkaynak 298", "=IFERROR(ydk_ozKaynak298,0)", TL),
        (6, "Güven", "=IFERROR(MOTOR!G39,0)", CATI),
        (7, "Risk", "=IFERROR(MOTOR!G40,0)", CATI),
        (8, "Fon oranı", "=IFERROR(MOTOR!G9,0)", YÜZDE),
        (9, "KV oranı", "=IFERROR(MOTOR!G10,0)", YÜZDE),
        (10, "Net Geç.32", "=IFERROR(MOTOR!G28,0)", TL),
        (11, "M-SEN fon", "=IFERROR(ydk_kuralDegisimSenaryo,0)", TL),
        (12, "Tahmin", "=IFERROR(SENARYO!B32,0)", TL),
        (13, "Vaka durum", "=IFERROR(ydk_vakaDurum,\"-\")", None),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=9, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=ydk_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=ydk_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 10, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Fark / maliyet",
         "=IFERROR(IF(ydk_maliyet=0,0,ydk_fark/ydk_maliyet),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Artış oranı ",TEXT(C11,"0.0%")," — endeks gücünü gösterir")'),
        ("T2", "Fon / fark",
         "=IFERROR(IF(ydk_fark=0,0,ydk_fonGec32/ydk_fark),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Fon yükü ",TEXT(C12,"0.0%")," — Geç.32 nakit etkisi")'),
        ("O1", "Senaryo fark sapması",
         "=IFERROR(IF(COUNT(SENARYO!C6:C9)<2,0,STDEV.P(SENARYO!C6:C9)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo fark sapması ",TEXT(C13,"₺ #,##0")," TL")'),
        ("O2", "Tornado max etki",
         "=IFERROR(ydk_duyarlilikSira,\"-\")",
         "=IFERROR(ydk_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo bant",
         "=IFERROR(ydk_senaryoKarsilastirma,0)",
         "=IFERROR(ydk_yorumSenaryo,\"-\")"),
        ("O6", "298 vs Geç.32 özkaynak",
         "=IFERROR(IF(ydk_ozKaynak298=0,0,MOTOR!G24/ydk_ozKaynak298),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Geç.32/298 özkaynak oranı ",TEXT(C16,"0.0%"))'),
        ("O8", "Kalite skoru",
         "=IFERROR(KONTROLLER!B12,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalite ",TEXT(C17,"0")," puan")'),
        ("M", "M-SEN fon",
         "=IFERROR(ydk_kuralDegisimSenaryo,0)",
         "=IFERROR(ydk_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin aralık",
         "=IFERROR(ydk_tahminAralik,0)",
         "=IFERROR(ydk_yorumTahmin,\"-\")"),
        ("I2", "Senaryo fark P90",
         "=IFERROR(IF(COUNT(SENARYO!C6:C9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!C6:C9,ydk_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"P90 fark ",TEXT(C20,"₺ #,##0")," TL")'),
    ]
    for i, (kod, ad, form, yorum) in enumerate(moduller):
        r = 11 + i
        h(ws, r, 1, kod, kalin=True, hiza="center", yazi="B08948")
        h(ws, r, 2, ad, kaydir=True)
        h(ws, r, 3, form, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[r].height = 28
        if kod in ("T1", "T2", "O6"):
            ws.cell(r, 3).number_format = YÜZDE
        elif kod != "O2":
            ws.cell(r, 3).number_format = TL

    h(ws, 23, 1, "Bileşen", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Maliyet", DEMO["maliyet"]),
        ("Birikmiş", DEMO["birikmis_amortisman"]),
        ("Net defter", 800_000),
        ("Yeniden değerli", 1_450_000),
        ("Fark", 450_000),
    ], 24):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)

    h(ws, 23, 4, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 5, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Fark", 450_000),
        ("Fon Geç.32", 9_000),
        ("Özkaynak 298", 450_000),
        ("Özkaynak Geç.32", 441_000),
        ("Vergi tasarrufu", 112_500),
    ], 24):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Kıymet Bileşenleri"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=28), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=28))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G9")

    g2 = BarChart()
    g2.type = "col"
    g2.title = "298/Ç vs Geç.32"
    g2.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "G23")

    g3 = LineChart()
    g3.title = "Etki Çizgisi"
    g3.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g3.height, g3.width = 7, 12
    ws.add_chart(g3, "P9")

    h(ws, 30, 1, "Metrik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 2, "Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Fark", 450_000),
        ("Fon", 9_000),
        ("Özkaynak", 450_000),
    ], 31):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "KPI Fark / Fon / Özkaynak"
    g4.add_data(Reference(ws, min_col=2, min_row=30, max_row=33), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=1, min_row=31, max_row=33))
    g4.height, g4.width = 7, 10
    ws.add_chart(g4, "P23")

    baski_hazirla(ws, "A1:F34", f"{URUN_AD} · PANO")
    genislik(ws, {"A": 28, "B": 16, "C": 18, "D": 55})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Canlı kontrol paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    kontroller_list = [
        ("Giriş tablosu boş mu?",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"BOŞ","DOLU")'),
        ("Negatif maliyet?",
         '=IF(GIRDI!B9<0,"NEGATİF","TEMİZ")'),
        ("Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(ydk_yilTaban),N(hesapYili)<=N(ydk_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        ("Fark sıfır mı?",
         '=IF(IFERROR(ydk_fark,0)=0,"SIFIR","NORMAL")'),
        ("Vaka kırık mı?",
         '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")'),
        ("Fark eşiği aşıyor mu?",
         '=IF(IFERROR(ydk_fark,0)>ydk_farkEsik,"AŞIYOR","NORMAL")'),
        ("Öncelik A/B seçili mi?",
         '=IF(OR(GIRDI!B13="A",GIRDI!B13="B"),"TEMİZ","EKSİK")'),
        ("Motor adım sayısı",
         "=COUNTA(MOTOR!B7:B60)"),
    ]
    for i, (ad, form) in enumerate(kontroller_list, 5):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Canlı formül | Değiştirmeyin")

    h(ws, 14, 1, "Toplam giriş alanı", yazi="333333")
    h(ws, 14, 2, "=IFERROR(ydk_girisBeklenen+COUNTA(GIRDI!B6:B13)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Dolu giriş", yazi="333333")
    h(ws, 15, 2, "=COUNTA(GIRDI!B6:B13)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Kalite skoru (modül)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=IFERROR(ROUND(IF(OR(B14=0,B14=\"\"),0,B15/B14*100),0),0)", sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)

    h(ws, 17, 1, "Fark kontrol", yazi="333333")
    h(ws, 17, 2, "=IFERROR(ydk_fark,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Fon kontrol", yazi="333333")
    h(ws, 18, 2, "=IFERROR(ydk_fonGec32,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Karar kontrol", yazi="333333")
    h(ws, 19, 2, "=IFERROR(ydk_kararMetni,\"-\")", yazi=DIS_REF_YEŞIL)

    genislik(ws, {"A": 40, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek senaryo kütüphanesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:I3")
    h(ws, 4, 1, "Bu sayfa demo değerleri saklar; GIRDI'ye kopyalanabilir.", yazi=GRİ)
    sutunlar = ["Senaryo", "Maliyet", "Birikmiş", "Endeks", "Kalan Ömür", "Öncelik",
                "Fark", "Fon Geç.32", "Özkaynak 298"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Fark": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(ROUND(MAX(0,tblOrnek[[#This Row],[Maliyet]]*"
            "(tblOrnek[[#This Row],[Endeks]]-1)),2),0))"
        ),
        "Fon Geç.32": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(ROUND(tblOrnek[[#This Row],[Fark]]*INDEX(tblKurallar[deger_2025],"
            "MATCH(\"YDK-002\",tblKurallar[kural_id],0)),2),0))"
        ),
        "Özkaynak 298": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(tblOrnek[[#This Row],[Fark]]-tblOrnek[[#This Row],[Fark]]*"
            "INDEX(tblKurallar[deger_2025],MATCH(\"YDK-004\",tblKurallar[kural_id],0)),0))"
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:I1004", sutunlar, formuller)
    ornekler = [
        ("Demo ana", 1_000_000, 200_000, 1.45, 10, "A"),
        ("Yüksek endeks", 1_000_000, 100_000, 1.80, 8, "A"),
        ("Düşük maliyet", 500_000, 50_000, 1.20, 5, "B"),
        ("Uzun ömür", 2_000_000, 400_000, 1.35, 20, "A"),
        ("Sıfır artış", 800_000, 100_000, 1.00, 10, "B"),
    ]
    for i, row in enumerate(ornekler, 6):
        for k, v in enumerate(row, 1):
            ws.cell(i, k).value = v
            if k in (2, 3):
                ws.cell(i, k).number_format = TL
            elif k == 4:
                ws.cell(i, k).number_format = "0.00"
            elif k == 5:
                ws.cell(i, k).number_format = CATI
            if k <= 6:
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", sutunlar[k - 1], "Örnek senaryo değeri")

    for col, ad in [("B", "Maliyet"), ("C", "Birikmiş")]:
        dogrulama(ws, "decimal", "0", f"{col}6:{col}1004",
                  baslik=ad, mesaj=f"{ad} tutarını TL girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ile 1e12 arasında.",
                  isaret="between", f2="1000000000000")
    dogrulama(ws, "decimal", "1", "D6:D1004",
              baslik="Endeks", mesaj="1-10 çarpan.",
              hata_baslik="Geçersiz", hata_mesaj="1 ile 10.",
              isaret="between", f2="10")
    dogrulama(ws, "list", "ListeYorumAB", "F6:F1004",
              baslik="Öncelik", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.")

    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6])
    h(ws, 5, 12, "FarkDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([450000, 800000, 100000, 700000, 0], 6):
        h(ws, i, 12, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Örnek Senaryo Farkları"
    g.add_data(Reference(ws, min_col=12, min_row=5, max_row=10), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=6, max_row=10))
    g.height, g.width = 8, 12
    ws.add_chart(g, "K5")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 13)})
    sabitle(ws, "A6")


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", "595959", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Doğrulama listeleri", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Yıllar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, y in enumerate([2025, 2026, 2027], 6):
        h(ws, i, 1, y, sayi=CATI, zemin=GIRIS_SARI)
        ws.cell(i, 1).protection = Protection(locked=False)
        _kim(ws, f"A{i}", "Yıl", "Hesap yılı listesi")
    h(ws, 5, 3, "Öncelik A/B", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate(["A", "B"], 6):
        h(ws, i, 3, v, zemin=GIRIS_SARI)
        ws.cell(i, 3).protection = Protection(locked=False)
        _kim(ws, f"C{i}", "Öncelik", "298/Ç veya Geç.32 önceliği")
    h(ws, 5, 5, "Evet/Hayır", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate(["Evet", "Hayır"], 6):
        h(ws, i, 5, v, zemin=GIRIS_SARI)
        ws.cell(i, 5).protection = Protection(locked=False)
        _kim(ws, f"E{i}", "Seçim", "Evet veya Hayır")
    genislik(ws, {"A": 12, "C": 12, "E": 12})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Parametreler (tek kaynak)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = ["anahtar", "deger", "birim", "aciklama", "kaynak", "yururluk_tarihi", "dogrulama_tarihi", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    params = [
        ("hesapYili", DEMO["hesapYili"], "yıl", "Aktif hesap yılı", "Kullanıcı / GIRDI", "01.01.2025", "11.08.2026"),
        ("raporTarihi", RAPOR_TARIH, "tarih", "Rapor tarihi (sabit)", "Üretim", "01.01.2025", "11.08.2026"),
        ("ydk_tolerans", 0.01, "TL", "Vaka tutarlılık toleransı", "Uygulama notu", "01.01.2025", "11.08.2026"),
        ("ydk_farkEsik", 100_000, "TL", "PANO fark uyarı eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("ydk_senaryoIyi", 0.90, "çarpan", "İyimser maliyet çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ydk_senaryoKotu", 1.10, "çarpan", "Kötümser maliyet çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ydk_oranArtisPuan", 0.01, "oran", "M-SEN fon +1 puan", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("ydk_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ydk_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ydk_tornadoMaliyet", 1.05, "çarpan", "Tornado maliyet şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ydk_tornadoEndeks", 1.05, "çarpan", "Tornado endeks şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ydk_tasarrufCarpan", 0.5, "çarpan", "Vergi tasarrufu ağırlığı", "Belirsizlik A/B", "01.01.2025", "11.08.2026"),
        ("ydk_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "11.08.2026"),
        ("ydk_parmakIzi", "YDK-PRO-1.0.0", "metin", "Dosya parmak izi etiketi", "Üretim", "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv", "01.01.2025", "11.08.2026"),
        ("ydk_girisBeklenen", 8, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ydk_riskDusuk", 15, "puan", "Düşük risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("ydk_riskOrta", 55, "puan", "Orta risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("ydk_riskYuksek", 85, "puan", "Yüksek risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("ydk_kararEsik", 25_000, "TL", "GO/REVIEW fayda fark eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("ydk_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı (hesapYili−taban)", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ydk_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("vaka_maliyet_1", 1_000_000, "TL", "Vaka1 maliyet", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_endeks_1", 1.45, "çarpan", "Vaka1 endeks", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_maliyet_2", 2_000_000, "TL", "Vaka2 maliyet", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_endeks_2", 1.20, "çarpan", "Vaka2 endeks", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_maliyet_3", 500_000, "TL", "Vaka3 maliyet", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_endeks_3", 1.10, "çarpan", "Vaka3 endeks", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
    ]
    for i, row in enumerate(params, 6):
        for k, v in enumerate(row, 1):
            sayi = None
            if k == 2:
                if row[2] == "tarih":
                    sayi = TARİH
                elif row[2] in ("oran", "çarpan"):
                    sayi = YÜZDE if row[2] == "oran" else "0.00"
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
        h(ws, i, 8,
          f'=IF(A{i}="","",IF(B{i}="","EKSİK","TAM"))',
          yazi=DIS_REF_YEŞIL, hiza="center")

    son_satir = 5 + len(params)
    h(ws, son_satir + 2, 2, "Parametre satırları yukarıda; ad tanımları B kolonuna bağlı.", kalin=True, yazi=KOYU_LACIVERT, kaydir=True)

    dogrulama(ws, "list", "ListeYillar", "B6",
              baslik="Hesap Yılı", mesaj="Aktif yılı seçin.",
              hata_baslik="Geçersiz", hata_mesaj="2025-2027.")
    dogrulama(ws, "decimal", "0", "B9",
              baslik="Fark eşiği", mesaj="Uyarı eşiği TL.",
              hata_baslik="Geçersiz", hata_mesaj="0 ve üzeri.",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", "B10",
              baslik="İyimser çarpan", mesaj="0-2 aralığı.",
              hata_baslik="Geçersiz", hata_mesaj="0-2.",
              isaret="between", f2="2")
    dogrulama(ws, "decimal", "0", "B11",
              baslik="Kötümser çarpan", mesaj="0-2 aralığı.",
              hata_baslik="Geçersiz", hata_mesaj="0-2.",
              isaret="between", f2="2")
    dogrulama(ws, "decimal", "0", "B12",
              baslik="Oran artış", mesaj="Puan cinsinden.",
              hata_baslik="Geçersiz", hata_mesaj="0-0.2.",
              isaret="between", f2="0.2")

    genislik(ws, {"A": 28, "B": 24, "C": 10, "D": 32, "E": 22, "F": 14, "G": 14, "H": 12})
    h(ws, son_satir + 4, 2, f"Sürüm {SURUM} | Şifre koruması: {SIFRE} (formül alanları)", yazi=GRİ, boyut=9, kaydir=True)


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar kuralları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "VERİ YOK — giriş tablosu boşsa hesap yapılmaz.",
        "GO · 298/Ç UYGUN — 298/Ç net faydası Geç.32 üstünde (eşik dışı).",
        "GO · GEÇ.32 UYGUN — Geç.32 net faydası (özkaynak+vergi) 298/Ç üstünde.",
        "REVIEW · DİKKAT — iki yol arasındaki fark eşik içinde; öncelik yorumunu gözden geçirin.",
        "YENİDEN DEĞERLEME YOK — endeks artışı/fark sıfır.",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 11, 1, "Belirsizlik beyanları (M05)", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari(["gec32_vs_298c_oncelik"]), 12):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 14, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 2,
      "298/Ç yolu öncelikli değerlendirilir. Fon yükü düşük; özkaynak artışı ana ölçüttür. "
      "Bu nokta mevzuatta tartışmalıdır; sonuç yorum B ile yan yana okunmalıdır.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B14:J14")
    h(ws, 15, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 2,
      "Geçici 32 yolu öncelikli değerlendirilir. Fon vergisi nakit çıkışıdır; ek amortisman "
      "vergi tasarrufu net faydaya eklenir. Dosya her iki yorumun sonucunu MOTOR'da üretir.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B15:J15")

    h(ws, 17, 1, "Kullanım sırası", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 1,
      "GIRDI → KURALLAR (gerekirse) → MOTOR (inceleme) → SENARYO → VAKALAR → PANO → KANIT_RAPORU → AYARLAR.",
      kaydir=True, yazi="333333")
    ws.merge_cells("A18:J18")
    h(ws, 20, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 18, "B": 70})


def kosullu_bicimlendirme(wb):
    yesil = Font(color=NORMAL, bold=True)
    kirmizi = Font(color=KRITIK, bold=True)
    amber = Font(color="B7791F", bold=True)
    y_fill = PatternFill("solid", fgColor="E2EFDA")
    k_fill = PatternFill("solid", fgColor="FDE9E9")
    a_fill = PatternFill("solid", fgColor="FFF2CC")

    ws = wb["PANO"]
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("UYGUN",B6))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B6))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("VERİ YOK",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("C4", CellIsRule(operator="greaterThan", formula=["0"], font=amber))
    ws.conditional_formatting.add("C4", CellIsRule(operator="equal", formula=["0"], font=yesil))
    ws.conditional_formatting.add("F4", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("F4", CellIsRule(operator="between", formula=["50", "79"], font=amber))
    ws.conditional_formatting.add("F4", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))
    ws.conditional_formatting.add("G4", CellIsRule(operator="greaterThanOrEqual", formula=["70"], font=kirmizi))
    ws.conditional_formatting.add("G4", CellIsRule(operator="between", formula=["40", "69"], font=amber))
    ws.conditional_formatting.add("G4", CellIsRule(operator="lessThan", formula=["40"], font=yesil))
    for col in range(2, 14):
        harf = get_column_letter(col)
        ws.conditional_formatting.add(f"{harf}4", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))

    ws = wb["GIRDI"]
    for col in ("B",):
        ws.conditional_formatting.add(f"{col}9:{col}12", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(f"{col}9:{col}10", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("F22:F29", FormulaRule(formula=['F22="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F22:F29", FormulaRule(formula=['F22="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F6:F9", CellIsRule(operator="notEqual", formula=["0"], font=amber))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("C6:C9", CellIsRule(operator="greaterThan", formula=["0"], font=amber))
    ws.conditional_formatting.add("C6:C9", CellIsRule(operator="equal", formula=["0"], font=yesil))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("GO",E6))'], font=yesil))
    ws.conditional_formatting.add("B16:B18", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add("B5:B11", FormulaRule(formula=['OR(B5="KIRIK",B5="HATALI",B5="NEGATİF",B5="BOŞ",B5="EKSİK",B5="AŞIYOR")'],
                                                        font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B5:B11", FormulaRule(formula=['OR(B5="TEMİZ",B5="DOLU",B5="NORMAL",B5="TUTARLI")'],
                                                        font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B12", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("B12", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("G7:G60", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("F7:F60", FormulaRule(formula=['F7="YDK-002"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F60", FormulaRule(formula=['F7="YDK-001"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "C", "D", "E", "G", "H", "I"):
        ws.conditional_formatting.add(f"{harf}6:{harf}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("G6:G20", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B19", FormulaRule(formula=['ISNUMBER(SEARCH("UYGUN",B19))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B19", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B19))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B17", CellIsRule(operator="greaterThan", formula=["0"], font=amber))


def ad_tanimlari(wb):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "ydk_maliyet", "GIRDI!$B$9")
    ad_ekle(wb, "ydk_netDefter", "GIRDI!$B$15")
    ad_ekle(wb, "ydk_fark", "MOTOR!$G$20")
    ad_ekle(wb, "ydk_fonGec32", "MOTOR!$G$21")
    ad_ekle(wb, "ydk_ozKaynak298", "MOTOR!$G$23")
    ad_ekle(wb, "ydk_kararMetni", "MOTOR!$G$31")
    ad_ekle(wb, "ydk_kararGerekce", "MOTOR!$G$32")
    ad_ekle(wb, "ydk_maddeAtifMetni", "MOTOR!$G$52")
    ad_ekle(wb, "ydk_kanitRaporu", "MOTOR!$G$53")
    ad_ekle(wb, "ydk_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "ydk_parmakIzi", "AYARLAR!$B$19")

    ad_ekle(wb, "ydk_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "ydk_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "ydk_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "ydk_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "ydk_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "ydk_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "ydk_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "ydk_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "ydk_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "ydk_vakaFark", "VAKALAR!$B$13")

    ayar_map = {
        "ydk_tolerans": 8,
        "ydk_farkEsik": 9,
        "ydk_senaryoIyi": 10,
        "ydk_senaryoKotu": 11,
        "ydk_oranArtisPuan": 12,
        "ydk_tahminAlt": 13,
        "ydk_tahminUst": 14,
        "ydk_tornadoMaliyet": 15,
        "ydk_tornadoEndeks": 16,
        "ydk_tasarrufCarpan": 17,
        "ydk_yuzdelikOran": 18,
        "ydk_girisBeklenen": 21,
        "ydk_riskDusuk": 22,
        "ydk_riskOrta": 23,
        "ydk_riskYuksek": 24,
        "ydk_kararEsik": 25,
        "ydk_yilTaban": 26,
        "ydk_yuvarMax": 27,
        "vaka_maliyet_1": 28,
        "vaka_endeks_1": 29,
        "vaka_maliyet_2": 30,
        "vaka_endeks_2": 31,
        "vaka_maliyet_3": 32,
        "vaka_endeks_3": 33,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "ydk_modulT1": (11, "C"), "ydk_modulT1Yorum": (11, "D"),
        "ydk_modulT2": (12, "C"), "ydk_modulT2Yorum": (12, "D"),
        "ydk_modulO1": (13, "C"), "ydk_modulO1Yorum": (13, "D"),
        "ydk_modulO6": (16, "C"), "ydk_modulO6Yorum": (16, "D"),
        "ydk_modulO8": (17, "C"), "ydk_modulO8Yorum": (17, "D"),
        "ydk_modulI2": (20, "C"), "ydk_modulI2Yorum": (20, "D"),
    }
    for ad, (satir, kol) in modul_map.items():
        ad_ekle(wb, ad, f"PANO!${kol}${satir}")

    ad_ekle(wb, "ListeYillar", "LISTELER!$A$6:$A$8")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$C$6:$C$7")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$6:$E$7")


def main(cikti_yolu=None):
    wb = Workbook()
    wb.remove(wb.active)

    siralar = [
        ("KAPAK", kapak),
        ("HIZLI_BASLANGIC", hizli_baslangic),
        ("PANO", pano),
        ("GIRDI", girdi),
        ("KURALLAR", kurallar),
        ("MOTOR", motor),
        ("SENARYO", senaryo),
        ("VAKALAR", vakalar),
        ("KANIT_RAPORU", kanit_raporu),
        ("KONTROLLER", kontroller),
        ("ORNEK_VERI", ornek_veri),
        ("LISTELER", listeler),
        ("AYARLAR", ayarlar),
        ("KILAVUZ", kilavuz),
    ]
    for ad, fn in siralar:
        ws = wb.create_sheet(ad)
        fn(ws)

    ad_tanimlari(wb)
    kosullu_bicimlendirme(wb)
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    wb.calculation.fullCalcOnLoad = True

    kok = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    varsayilan = os.path.join(kok, "YenidenDegerlemeKomutaMerkezi.xlsx")
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
    return hedef


if __name__ == "__main__":
    yol = sys.argv[1] if len(sys.argv) > 1 else None
    uretilen = main(yol)
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
    cikti = os.path.join(repo, "cikti", "YenidenDegerlemeKomutaMerkezi.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
