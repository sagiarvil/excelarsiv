#!/usr/bin/env python3
"""
Asgari Kurumlar Vergisi Simülasyon Motoru — üretim betiği (A1 / manda v6).
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

URUN_AD = "Asgari Kurumlar Vergisi Simülasyon Motoru"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 10)

# Demo: matrah 10.200.000 → asgari KV 1.020.000; ödenen 800.000 → fark 220.000
DEMO = {
    "kurum_adi": "Örnek Sanayi A.Ş.",
    "donem": "2025",
    "hesapYili": 2025,
    "ticari_kar": 10_000_000,
    "kanunen_kabul_edilmeyen": 500_000,
    "istisnalar": 200_000,
    "indirimler": 100_000,
    "odenen_kurumlar_vergisi": 800_000,
    "istisna_kirilim_yorumu": "A",
}


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: Asgari KV hesabını ve kararı etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Fark ve karar bozulur | Kaynak: [Assumption] KVK 32/C",
    )


def kural_cek(kural_id: str) -> str:
    """Oranları tblKurallar'dan çeker; yıl seçimi akv_yilTaban ile (G07: adında rakam yok)."""
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-akv_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


# Yuvarlama basamağı uç değerde (#VALUE!) kırmasın — tavan AYARLAR'dan (G07)
_YV = "MAX(0,MIN(akv_yuvarMax,IFERROR(N(G12),0)))"


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
      "Kurumlar matrahınızı girin; asgari kurumlar vergisi tabanı, ödenen KV ile farkı "
      "ve tamamlayıcı vergi kararını senaryolu olarak üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Çok yıllı kural tablosu (2025/2026/2027) ile oran güncellemesine hazır motor",
        "Asgari KV, ödenen KV farkı ve gerekçeli karar cümlesi",
        "Baz / iyimser / kötümser + kural değişimi (+1 puan) senaryoları",
        "Tebliğ tarzı altın vakalar ve KANIT_RAPORU (A4 imza)",
        "Belirsizlik beyanı: istisna kırılım sırası (yorum A / B)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "SMMM / YMM — müşteriye imzalı kanıt raporu",
        "CFO — asgari KV farkını bütçeye yazmak",
        "Ortak / yönetim — ek vergi riskini görmek",
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
        ("Adım 1 — Girdiler", "GIRDI sayfasında hesap yılı, matrah bileşenleri ve ödenen KV'yi sarı hücrelere girin."),
        ("Adım 2 — Karar", "PANO'da asgari KV, fark ve karar rozetini izleyin."),
        ("Adım 3 — Kanıt", "KANIT_RAPORU'nu yazdırıp imzalayın; VAKALAR tutarlılığını kontrol edin."),
    ]
    for i, (b, m) in enumerate(adimlar, 5):
        h(ws, i, 1, b, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        _kim(ws, f"A{i}", b, "Sırayla izleyin")
    h(ws, 9, 1, "Sık yapılan hatalar", kalin=True, yazi=KRITIK)
    for i, m in enumerate([
        "Matrahı KDV dahil girmek",
        "Ödenen KV yerine tahakkuk etmemiş tutarı yazmak",
        "Hesap yılını kural tablosuyla uyumsuz seçmek",
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
    h(ws, 3, 1, "Dönem ve kurum", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Matrah formülle üretilir.", yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Kurum Adı", DEMO["kurum_adi"], None, "kurum_adi",
         "Kurumun resmi unvanını yazın."),
        (7, "Dönem", DEMO["donem"], None, "donem",
         "Vergi dönemi etiketini yazın (ör. 2025)."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili",
         "2025, 2026 veya 2027 seçin."),
        (9, "Ticari Kâr [₺]", DEMO["ticari_kar"], TL, "ticari_kar",
         "Dönem ticari kârını TL girin."),
        (10, "Kanunen Kabul Edilmeyen [₺]", DEMO["kanunen_kabul_edilmeyen"], TL,
         "kanunen_kabul_edilmeyen", "KKEG tutarını TL girin."),
        (11, "İstisnalar [₺]", DEMO["istisnalar"], TL, "istisnalar",
         "Matrahtan düşülecek istisnaları TL girin."),
        (12, "İndirimler [₺]", DEMO["indirimler"], TL, "indirimler",
         "Matrahtan düşülecek indirimleri TL girin."),
        (13, "Ödenen Kurumlar Vergisi [₺]", DEMO["odenen_kurumlar_vergisi"], TL,
         "odenen_kurumlar_vergisi", "Normal KV olarak ödenen / tahakkuk eden tutarı girin."),
        (14, "İstisna Kırılım Yorumu (A/B)", DEMO["istisna_kirilim_yorumu"], None,
         "istisna_kirilim_yorumu", "Belirsizlikte yorum A veya B seçin."),
    ]
    h(ws, 5, 1, "Alan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 2, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 3, "Birim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    for satir, etiket, deger, sayi, _ad, mesaj in alanlar:
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, "₺" if sayi == TL else ("yıl" if _ad == "hesapYili" else "metin"), yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    # Matrah (formül — kilitli)
    h(ws, 16, 1, "Kurumlar Matrahı [₺]", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 2, "=B9+B10-B11-B12", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 17, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B14)=0,akv_girisBeklenen=0),0,ROUND(COUNTA(B6:B14)/akv_girisBeklenen*100,0)),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)

    # tblGirdi — D11 boşaltma + ölçek için (ayna satırlar)
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
        ("ticari_kar", DEMO["ticari_kar"], "TL", "Evet", "GIRDI"),
        ("kanunen_kabul_edilmeyen", DEMO["kanunen_kabul_edilmeyen"], "TL", "Evet", "GIRDI"),
        ("istisnalar", DEMO["istisnalar"], "TL", "Evet", "GIRDI"),
        ("indirimler", DEMO["indirimler"], "TL", "Evet", "GIRDI"),
        ("odenen_kurumlar_vergisi", DEMO["odenen_kurumlar_vergisi"], "TL", "Evet", "GIRDI"),
        ("istisna_kirilim_yorumu", DEMO["istisna_kirilim_yorumu"], "A/B", "Evet", "GIRDI"),
    ]
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 22):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if isinstance(d, (int, float)) and a != "hesapYili":
            ws.cell(i, 2).number_format = TL
        elif a == "hesapYili":
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
    dogrulama(ws, "list", "ListeYorumAB", "B14",
              baslik="Yorum", mesaj="İstisna kırılımında A veya B seçin.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B seçin.", bos=False)
    dogrulama(ws, "list", "ListeYillar", "B24",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeYorumAB", "B30",
              baslik="Yorum", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.")
    for aralik, baslik, mesaj in [
        ("B9", "Ticari Kâr", "0 ile 1e12 arasında TL."),
        ("B10", "KKEG", "0 ile 1e12 arasında TL."),
        ("B11", "İstisnalar", "0 ile 1e12 arasında TL."),
        ("B12", "İndirimler", "0 ile 1e12 arasında TL."),
        ("B13", "Ödenen KV", "0 ile 1e12 arasında TL."),
        ("B25", "Ticari Kâr", "Tablo: TL tutar."),
        ("B26", "KKEG", "Tablo: TL tutar."),
        ("B27", "İstisnalar", "Tablo: TL tutar."),
        ("B28", "İndirimler", "Tablo: TL tutar."),
        ("B29", "Ödenen KV", "Tablo: TL tutar."),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj=mesaj,
                  hata_baslik="Geçersiz tutar", hata_mesaj="0 ile 1.000.000.000.000 arasında olmalı.",
                  isaret="between", f2="1000000000000")

    giris_hucreleri(ws, 6, 14, [2])
    giris_hucreleri(ws, 22, 30, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 36, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1022, "Sarı alanlar giriş; matrah ve kalite formülleri kilitlidir.")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Mevzuat kural tablosu (yıl yan yana)", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["AKV-001", "KVK md.32", "kv_genel_oran", "kurumlar_matrah>0", "kv_genel_oran",
         0.25, 0.25, 0.25, "01.01.2025", "KVK md.32"],
        ["AKV-002", "KVK 32/C", "asgari_taban_matrah", "her zaman", "asgari_taban_matrah",
         0, 0, 0, "01.01.2025", "KVK 32/C"],
        ["AKV-003", "KVK 32/C-3", "asgari_oran", "kurumlar_matrah>0", "asgari_oran",
         0.10, 0.10, 0.10, "01.01.2025", "7524 s.K. md.36"],
        ["AKV-004", "KVK 32/C", "tamamlayici_hesap", "asgari>ödenen", "tamamlayici_bayrak",
         1, 1, 1, "01.01.2025", "KVK 32/C"],
        ["AKV-005", "Genel", "yuvarlama", "her hesap", "yuvarlama_ondalik",
         2, 2, 2, "01.01.2025", "Uygulama notu"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = YÜZDE if k in (6, 7, 8) and isinstance(v, float) and v < 1 else (
                CATI if k in (6, 7, 8) else None)
            if k in (6, 7, 8) and s[0] in ("AKV-002", "AKV-004", "AKV-005"):
                sayi = CATI
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-akv_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    # Yıl matrisi özeti (ad tanımı hedefi)
    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Asgari oran (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("AKV-003"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Genel KV oranı (aktif yıl)", yazi="333333")
    h(ws, 16, 2, kural_cek("AKV-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=_xlfn.TEXTJOIN(" | ",TRUE,"AKV-001=",TEXT(B16,"0.0%"),"AKV-003=",TEXT(B15,"0.0%"),"yıl=",hesapYili)',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    for col in ("F", "G", "H"):
        dogrulama(ws, "decimal", "0", f"{col}7:{col}11",
                  baslik="Kural değeri", mesaj="Yıl kolonuna sayısal parametre girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ile 1 arasında veya tamsayı.",
                  isaret="between", f2="1")
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

    km = kural_cek  # kısayol
    # deger formülleri — GIRDI hücrelerine ve önceki satırlara bağlı
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    # Satır = 6 + adım no. G kolonu canlı değer.
    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif mevzuat yılı", "AKV-005")  # G7
    ekle("Genel KV oranı", km("AKV-001"), "oran", "KVK md.32 oranı", "AKV-001")  # G8
    ekle("Asgari taban matrah", km("AKV-002"), "TL", "Taban matrah eşiği", "AKV-002")  # G9
    ekle("Asgari oran", km("AKV-003"), "oran", "Asgari KV oranı", "AKV-003")  # G10
    ekle("Tamamlayıcı bayrak", km("AKV-004"), "bayrak", "Tamamlayıcı hesap açık mı", "AKV-004")  # G11
    ekle("Yuvarlama ondalık", km("AKV-005"), "adet", "Yuvarlama basamağı", "AKV-005")  # G12
    ekle("Ticari kâr", "=GIRDI!B9", "TL", "Ham ticari kâr", "AKV-001")  # G13
    ekle("Kanunen kabul edilmeyen", "=GIRDI!B10", "TL", "KKEG eklemesi", "AKV-001")  # G14
    ekle("İstisnalar", "=GIRDI!B11", "TL", "İstisna düşümü", "AKV-002")  # G15
    ekle("İndirimler", "=GIRDI!B12", "TL", "İndirim düşümü", "AKV-002")  # G16
    ekle("Matrah ham", "=G13+G14-G15-G16", "TL", "Bileşenlerden matrah", "AKV-002")  # G17
    ekle("Matrah taban kontrol", "=MAX(G17,G9)", "TL", "Taban ile karşılaştır", "AKV-002")  # G18
    ekle("Kurumlar matrahı", f"=ROUND(G18,{_YV})", "TL", "Yuvarlanmış matrah", "AKV-005")  # G19
    ekle("Ödenen KV", "=GIRDI!B13", "TL", "Ödenen/tahakkuk KV", "AKV-001")  # G20
    ekle("Normal KV (referans)", f"=ROUND(G19*G8,{_YV})", "TL", "Genel oranlı KV", "AKV-001")  # G21
    ekle("Asgari KV", f"=ROUND(G19*G10,{_YV})", "TL", "Matrah × asgari oran", "AKV-003")  # G22
    ekle("Fark ham", "=G22-G20", "TL", "Asgari − ödenen", "AKV-004")  # G23
    ekle("Fark (sıfır altı yok)", "=MAX(0,G23)", "TL", "Tamamlayıcı tutar", "AKV-004")  # G24
    ekle("Tamamlayıcı uygulanır mı", '=IF(AND(G11=1,G24>0),1,0)', "bayrak", "Bayrak × fark", "AKV-004")  # G25
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",IF(G24=0,"EK_YOK","FARK"))',
         "kod", "Karar kodu", "AKV-004")  # G26
    ekle("Karar metni",
         '=IF(G26="VERI_YOK","VERİ YOK",IF(G26="EK_YOK","EK VERGİ YOK / UYGUN","FARK HESAPLA / DİKKAT"))',
         "metin", "Kullanıcıya karar", "AKV-004")  # G27
    ekle("Gerekçe",
         '=IF(G26="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G26="EK_YOK","Ödenen KV asgari tutarı karşılıyor; ek vergi yok.",'
         '"Ödenen KV asgari tutarın altında; tamamlayıcı fark hesaplandı."))',
         "metin", "Gerekçe cümlesi", "AKV-004")  # G28
    ekle("Ödenen / asgari oranı", "=IF(G22=0,0,G20/G22)", "oran", "Karşılama oranı", "AKV-003")  # G29
    ekle("Fark / matrah", "=IF(G19=0,0,G24/G19)", "oran", "Fark yoğunluğu", "AKV-003")  # G30
    ekle("Normal − asgari", "=G21-G22", "TL", "Oran farkı tutarı", "AKV-001")  # G31
    ekle("Yorum A etkisi",
         f"=ROUND((GIRDI!B9+GIRDI!B10-GIRDI!B11*akv_yorumAIstisnaCarpan-GIRDI!B12)*G10,{_YV})",
         "TL", "Yorum A alternatif matrah asgari", "AKV-003")  # G32
    ekle("Yorum B etkisi",
         f"=ROUND((GIRDI!B9+GIRDI!B10-GIRDI!B11-GIRDI!B12)*G10,{_YV})",
         "TL", "Yorum B (tam istisna)", "AKV-003")  # G33
    ekle("Seçili yorum asgari",
         '=IF(GIRDI!B14="B",G33,G32)',
         "TL", "A/B seçimine göre", "AKV-003")  # G34
    ekle("Yorum farkı", "=ABS(G32-G33)", "TL", "Belirsizlik aralığı", "AKV-003")  # G35
    ekle("Güven skoru",
         "=ROUND(IF(COUNTA(GIRDI!B6:B14)=0,0,COUNTA(GIRDI!B6:B14)/COUNTA(GIRDI!B6:B14)*100),0)",
         "puan", "Giriş bütünlüğü", "AKV-005")  # G36 — doluluk: aşağıdaki satırda düzeltilir
    ekle("Risk skoru",
         '=IF(G26="VERI_YOK",0,IF(G24=0,akv_riskDusuk,IF(G24/MAX(G22,1)>akv_riskEsikOran,akv_riskYuksek,akv_riskOrta)))',
         "puan", "Fark riski", "AKV-004")  # G37
    ekle("Senaryo iyimser matrah", "=G19*akv_senaryoIyi", "TL", "İyimser matrah", "AKV-002")  # G38
    ekle("Senaryo kötümser matrah", "=G19*akv_senaryoKotu", "TL", "Kötümser matrah", "AKV-002")  # G39
    ekle("İyimser asgari KV", f"=ROUND(G38*G10,{_YV})", "TL", "İyimser asgari", "AKV-003")  # G40
    ekle("Kötümser asgari KV", f"=ROUND(G39*G10,{_YV})", "TL", "Kötümser asgari", "AKV-003")  # G41
    ekle("Kural+1 puan asgari", f"=ROUND(G19*(G10+akv_oranArtisPuan),{_YV})", "TL",
         "M-SEN +1 puan", "AKV-003")  # G42
    ekle("M-SEN fark", "=MAX(0,G42-G20)", "TL", "Kural değişimi farkı", "AKV-004")  # G43
    ekle("Tahmin üst", "=G24*akv_tahminUst", "TL", "Tahmin üst bant", "AKV-005")  # G44
    ekle("Tahmin alt", "=G24*akv_tahminAlt", "TL", "Tahmin alt bant", "AKV-005")  # G45
    ekle("Tornado: matrah etkisi",
         f"=ABS(ROUND((G19*akv_tornadoMatrah)*G10,{_YV})-G22)",
         "TL", "Matrah ± etki", "AKV-003")  # G46
    ekle("Tornado: ödenen etkisi",
         "=ABS(MAX(0,G22-G20*akv_tornadoOdenen)-G24)",
         "TL", "Ödenen ± etki", "AKV-004")  # G47
    ekle("Tornado: oran etkisi",
         f"=ABS(ROUND(G19*(G10+akv_oranArtisPuan),{_YV})-G22)",
         "TL", "Oran ± etki", "AKV-003")  # G48
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"AKV-001→KVK md.32","AKV-003→KVK 32/C-3","AKV-004→tamamlayıcı")',
         "metin", "Kanıt atıfları", "AKV-001")  # G49
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Matrah=",TEXT(G19,"₺ #,##0"),"Asgari=",TEXT(G22,"₺ #,##0"),"Fark=",TEXT(G24,"₺ #,##0"))',
         "metin", "Rapor özeti", "AKV-005")  # G50

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
        elif birim in ("puan", "yıl", "adet", "bayrak"):
            ws.cell(r, 7).number_format = CATI
            ws.cell(r, 3).number_format = CATI

    # Güven: dolu/beklenen (beklenen = akv_girisBeklenen) — satır 42 = G36
    guven = (
        "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B14)=0,akv_girisBeklenen=0),0,"
        "COUNTA(GIRDI!B6:B14)/akv_girisBeklenen*100),0),0)"
    )
    ws.cell(42, 3).value = guven
    ws.cell(42, 7).value = guven

    # tblMotor YOK — Ö2 tablo genişletmesi benzersiz adım formüllerini ezer.
    genislik(ws, {"A": 8, "B": 40, "C": 55, "D": 10, "E": 28, "F": 12, "G": 22})
    sabitle(ws, "A7")
    h(ws, 6 + len(adimlar) + 2, 9,
      "Motor adımları kilitlidir; oranlar yalnızca KURALLAR'dan gelir.", yazi=GRİ, kaydir=True)
    return len(adimlar)


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve duyarlılık", kalin=True, boyut=13, yazi=KOYU_LACIVERT)

    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Matrah Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Asgari KV", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Fark", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(akv_senaryoIyi+N(GIRDI!B9)*0,0)",
         "=IFERROR(ROUND(IFERROR(akv_matrah,0)*B6*IFERROR(INDEX(tblKurallar[aktif_deger],MATCH(\"AKV-003\",tblKurallar[kural_id],0)),0),2),0)",
         "=IFERROR(MAX(0,C6-IFERROR(akv_odenenKv,0)),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(D6=0,"EK VERGİ YOK","FARK HESAPLA"))'),
        ("Baz", "=IFERROR(1+N(GIRDI!B9)*0,1)",
         "=IFERROR(ROUND(IFERROR(akv_matrah,0)*B7*IFERROR(INDEX(tblKurallar[aktif_deger],MATCH(\"AKV-003\",tblKurallar[kural_id],0)),0),2),0)",
         "=IFERROR(MAX(0,C7-IFERROR(akv_odenenKv,0)),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(D7=0,"EK VERGİ YOK","FARK HESAPLA"))'),
        ("Kötümser", "=IFERROR(akv_senaryoKotu+N(GIRDI!B9)*0,0)",
         "=IFERROR(ROUND(IFERROR(akv_matrah,0)*B8*IFERROR(INDEX(tblKurallar[aktif_deger],MATCH(\"AKV-003\",tblKurallar[kural_id],0)),0),2),0)",
         "=IFERROR(MAX(0,C8-IFERROR(akv_odenenKv,0)),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(D8=0,"EK VERGİ YOK","FARK HESAPLA"))'),
        ("M-SEN +1 puan", "=IFERROR(1+N(GIRDI!B9)*0,1)",
         "=IFERROR(ROUND(IFERROR(akv_matrah,0)*(IFERROR(INDEX(tblKurallar[aktif_deger],MATCH(\"AKV-003\",tblKurallar[kural_id],0)),0)+IFERROR(akv_oranArtisPuan,0)),2),0)",
         "=IFERROR(MAX(0,C9-IFERROR(akv_odenenKv,0)),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(D9=0,"EK VERGİ YOK","FARK HESAPLA"))'),
    ]
    for i, (ad, carp, asg, fark, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, asg, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, fark, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant genişliği (iyimser−kötümser fark)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(D6-D8,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser fark ",TEXT(D6,"₺ #,##0")," · Baz ",TEXT(D7,"₺ #,##0"),'
      '" · Kötümser ",TEXT(D8,"₺ #,##0")," · Bant ",TEXT(B11,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:H12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Tornado matrah etkisi", yazi="333333")
    h(ws, 16, 2, "=IFERROR(MOTOR!G46,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado ödenen etkisi", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G47,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado oran etkisi", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G48,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
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

    h(ws, 22, 1, "M-SEN kural değişimi farkı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=D9", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Asgari oran +1 puan senaryosunda fark ",TEXT(D9,"₺ #,##0")," TL")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı (fark)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=MOTOR!G45", sayi=TL, yazi=DIS_REF_YEŞIL)  # alt
    h(ws, 25, 3, "=akv_fark", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 4, "=MOTOR!G44", sayi=TL, yazi=DIS_REF_YEŞIL)  # üst
    h(ws, 26, 1, "Tahmin (FORECAST)", kalin=True, yazi=KOYU_LACIVERT)
    # Basit seri: senaryo farklarından tahmin
    h(ws, 27, 1, "Seri nokta", yazi=GRİ)
    for i, v in enumerate([1, 2, 3, 4], 28):
        h(ws, i, 1, v, sayi=CATI)
        h(ws, i, 2, f"=D{5+v}", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 32, 1, "Tahmin sonraki", yazi="333333")
    h(ws, 32, 2, "=IFERROR(FORECAST.LINEAR(5,B28:B31,A28:A31),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Tahmin yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Tahmin ",TEXT(B32,"₺ #,##0")," · Alt ",TEXT(B25,"₺ #,##0"),'
      '" · Üst ",TEXT(D25,"₺ #,##0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    # Sabit demo serileri (D09b)
    h(ws, 5, 7, "FarkDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([180000, 220000, 320000, 322000], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([102000, 44000, 102000], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)
    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([180000, 220000, 320000, 322000], 28):
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

    # Beklenen sonuçlar demo formüllerle tutarlı olmalı
    # Vaka1: matrah 10.2M, oran 0.10, ödenen 800k → fark 220000
    # Vaka2: ödenen = asgari → fark 0
    # Vaka3: matrah 5M, ödenen 400k → asgari 500k → fark 100k
    vakalar_data = [
        ("V-001", "Tebliğ örnek 1 — eksik ödeme",
         "matrah=vaka_matrah_1; ödenen=vaka_odenen_1",
         220000,
         "=IFERROR(MAX(0,ROUND(vaka_matrah_1*INDEX(tblKurallar[deger_2025],MATCH(\"AKV-003\",tblKurallar[kural_id],0)),2)-vaka_odenen_1),0)",
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=akv_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Asgari KV Tebliği örnek senaryo"),
        ("V-002", "Tebliğ örnek 2 — yeterli ödeme",
         "matrah=vaka_matrah_2; ödenen=vaka_odenen_2",
         0,
         "=IFERROR(MAX(0,ROUND(vaka_matrah_2*INDEX(tblKurallar[deger_2025],MATCH(\"AKV-003\",tblKurallar[kural_id],0)),2)-vaka_odenen_2),0)",
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=akv_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Asgari KV Tebliği örnek senaryo"),
        ("V-003", "Tebliğ örnek 3 — düşük matrah",
         "matrah=vaka_matrah_3; ödenen=vaka_odenen_3",
         100000,
         "=IFERROR(MAX(0,ROUND(vaka_matrah_3*INDEX(tblKurallar[deger_2025],MATCH(\"AKV-003\",tblKurallar[kural_id],0)),2)-vaka_odenen_3),0)",
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=akv_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Asgari KV Tebliği örnek senaryo"),
        ("V-004", "Genel oran referans",
         "matrah=vaka_matrah_1; genel KV",
         2550000,
         "=IFERROR(ROUND(vaka_matrah_1*INDEX(tblKurallar[deger_2025],MATCH(\"AKV-001\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=akv_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "KVK md.32 kontrol"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            sayi = TL if k in (4, 5, 6) and not (isinstance(v, str) and v.startswith("=")) else (
                TL if k in (4, 5, 6) else None)
            h(ws, i, k, v, sayi=sayi if k in (4, 5, 6) else None,
              yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = TL

    # beklenen_sonuc statik (tebliğ sonucu); hesaplanan formül
    formuller = {
        "fark": "=IF(tblVakalar[[#This Row],[vaka_id]]=\"\",\"\",IFERROR(tblVakalar[[#This Row],[beklenen_sonuc]]-tblVakalar[[#This Row],[hesaplanan]],0))",
        "durum": '=IF(tblVakalar[[#This Row],[vaka_id]]="","",IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=akv_tolerans,"TUTARLI","KIRIK"),"KIRIK"))',
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
    for i, (b, he) in enumerate([(220000, 220000), (0, 0), (100000, 100000), (2550000, 2550000)], 6):
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
    h(ws, 3, 1, "KANIT RAPORU — Asgari Kurumlar Vergisi", kalin=True, boyut=14, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:F3")
    h(ws, 4, 1, "Rapor tarihi", yazi="333333")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 5, 1, "Kurum", yazi="333333")
    h(ws, 5, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Dönem / yıl", yazi="333333")
    h(ws, 6, 2, '=CONCATENATE(GIRDI!B7," / ",hesapYili)', yazi=DIS_REF_YEŞIL)

    h(ws, 8, 1, "Girdi özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, (et, form, sayi) in enumerate([
        ("Ticari kâr", "=GIRDI!B9", TL),
        ("KKEG", "=GIRDI!B10", TL),
        ("İstisnalar", "=GIRDI!B11", TL),
        ("İndirimler", "=GIRDI!B12", TL),
        ("Matrah", "=akv_matrah", TL),
        ("Ödenen KV", "=akv_odenenKv", TL),
    ], 9):
        h(ws, i, 1, et, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)

    h(ws, 16, 1, "Hesap zinciri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 17, 1, "Asgari KV", yazi="333333")
    h(ws, 17, 2, "=akv_asgariKv", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 1, "Fark", yazi="333333")
    h(ws, 18, 2, "=akv_fark", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "KARAR", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2, "=akv_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Gerekçe", yazi="333333")
    h(ws, 20, 2, "=akv_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B20:F20")

    h(ws, 22, 1, "Madde atıfları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=akv_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B22:F22")

    h(ws, 24, 1, "Kanıt gövdesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, "=akv_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B24:F24")

    h(ws, 26, 1, "Parmak izi (SHA-256 kısaltma)", yazi="333333")
    h(ws, 26, 2, "=akv_parmakIzi", yazi=GRİ)

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
    # A4 dikey
    ws.page_setup.orientation = "portrait"
    genislik(ws, {"A": 28, "B": 22, "C": 14, "D": 14, "E": 14, "F": 14})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim paneli — Asgari KV", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpiler = [
        (2, "Matrah", "=IFERROR(akv_matrah,0)", TL),
        (3, "Asgari KV", "=IFERROR(akv_asgariKv,0)", TL),
        (4, "Ödenen KV", "=IFERROR(akv_odenenKv,0)", TL),
        (5, "Fark", "=IFERROR(akv_fark,0)", TL),
        (6, "Güven", "=IFERROR(MOTOR!G36,0)", CATI),
        (7, "Risk", "=IFERROR(MOTOR!G37,0)", CATI),
        (8, "Asgari oran", "=IFERROR(MOTOR!G10,0)", YÜZDE),
        (9, "Genel oran", "=IFERROR(MOTOR!G8,0)", YÜZDE),
        (10, "Karşılama", "=IFERROR(MOTOR!G29,0)", YÜZDE),
        (11, "M-SEN fark", "=IFERROR(akv_kuralDegisimSenaryo,0)", TL),
        (12, "Tahmin", "=IFERROR(SENARYO!B32,0)", TL),
        (13, "Vaka durum", "=IFERROR(akv_vakaDurum,\"-\")", None),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=9, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=akv_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=akv_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    # Modül paneli
    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 10, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "İstisna / matrah payı",
         "=IFERROR(IF(akv_matrah=0,0,GIRDI!B11/akv_matrah),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İstisna payı ",TEXT(C11,"0.0%")," — pay yüksekse yorum A/B farkı büyür")'),
        ("T2", "Ödenen / asgari",
         "=IFERROR(IF(akv_asgariKv=0,0,akv_odenenKv/akv_asgariKv),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Karşılama ",TEXT(C12,"0.0%")," — 100% altı tamamlayıcı riski")'),
        ("O1", "Fark sapması (senaryo)",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,STDEV.P(SENARYO!D6:D9)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo fark sapması ",TEXT(C13,"₺ #,##0")," TL")'),
        ("O2", "Tornado max etki",
         "=IFERROR(akv_duyarlilikSira,\"-\")",
         "=IFERROR(akv_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo bant",
         "=IFERROR(akv_senaryoKarsilastirma,0)",
         "=IFERROR(akv_yorumSenaryo,\"-\")"),
        ("O6", "İstisna yoğunluğu",
         "=IFERROR(IF((GIRDI!B11+GIRDI!B12)=0,0,GIRDI!B11/(GIRDI!B11+GIRDI!B12)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İstisna/(istisna+indirim) ",TEXT(C16,"0.0%"))'),
        ("O8", "Kalite skoru",
         "=IFERROR(KONTROLLER!B12,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalite ",TEXT(C17,"0")," puan")'),
        ("M", "M-SEN fark",
         "=IFERROR(akv_kuralDegisimSenaryo,0)",
         "=IFERROR(akv_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin aralık genişliği",
         "=IFERROR(akv_tahminAralik,0)",
         "=IFERROR(akv_yorumTahmin,\"-\")"),
        ("I2", "Senaryo fark P90",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!D6:D9,akv_yuzdelikOran)),0)",
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

    # Grafik kaynakları — sabit demo (D09b; pazaryeri yöntemi)
    h(ws, 23, 1, "Bileşen", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Ticari kâr", DEMO["ticari_kar"]),
        ("KKEG", DEMO["kanunen_kabul_edilmeyen"]),
        ("İstisnalar", DEMO["istisnalar"]),
        ("İndirimler", DEMO["indirimler"]),
        ("Matrah", 10_200_000),
    ], 24):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)

    h(ws, 23, 4, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 5, "Vergi", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Asgari KV", 1_020_000),
        ("Ödenen KV", 800_000),
        ("Fark", 220_000),
        ("Normal KV", 2_550_000),
        ("M-SEN asgari", 1_122_000),
    ], 24):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Matrah Bileşenleri"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=28), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=28))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G9")

    g2 = BarChart()
    g2.type = "col"
    g2.title = "Vergi Kalemleri"
    g2.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "G23")

    g3 = LineChart()
    g3.title = "Senaryo Fark Çizgisi"
    g3.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g3.height, g3.width = 7, 12
    ws.add_chart(g3, "P9")

    h(ws, 30, 1, "Metrik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 2, "Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Fark", 220_000),
        ("Asgari", 1_020_000),
        ("Ödenen", 800_000),
    ], 31):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "KPI Fark / Asgari / Ödenen"
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
        ("Negatif ticari kâr?",
         '=IF(GIRDI!B9<0,"NEGATİF","TEMİZ")'),
        ("Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(akv_yilTaban),N(hesapYili)<=N(akv_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        ("Matrah sıfır mı?",
         '=IF(IFERROR(akv_matrah,0)=0,"SIFIR","NORMAL")'),
        ("Vaka kırık mı?",
         '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")'),
        ("Fark eşiği aşıyor mu?",
         '=IF(IFERROR(akv_fark,0)>akv_farkEsik,"AŞIYOR","NORMAL")'),
        ("Yorum A/B seçili mi?",
         '=IF(OR(GIRDI!B14="A",GIRDI!B14="B"),"TEMİZ","EKSİK")'),
        ("Motor adım sayısı",
         "=COUNTA(MOTOR!B7:B50)"),
    ]
    for i, (ad, form) in enumerate(kontroller_list, 5):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Canlı formül | Değiştirmeyin")

    h(ws, 14, 1, "Toplam giriş alanı", yazi="333333")
    h(ws, 14, 2, "=IFERROR(akv_girisBeklenen+COUNTA(GIRDI!B6:B14)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Dolu giriş", yazi="333333")
    h(ws, 15, 2, "=COUNTA(GIRDI!B6:B14)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Kalite skoru", kalin=True, yazi=KOYU_LACIVERT)
    # B12 referenced from PANO O8 — use row 12
    h(ws, 12, 1, "Kalite skoru (modül)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=IFERROR(ROUND(IF(OR(B14=0,B14=\"\"),0,B15/B14*100),0),0)", sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)

    # Extra live formulas for G22 (≥5)
    h(ws, 17, 1, "Asgari KV kontrol", yazi="333333")
    h(ws, 17, 2, "=IFERROR(akv_asgariKv,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Fark kontrol", yazi="333333")
    h(ws, 18, 2, "=IFERROR(akv_fark,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Karar kontrol", yazi="333333")
    h(ws, 19, 2, "=IFERROR(akv_kararMetni,\"-\")", yazi=DIS_REF_YEŞIL)

    genislik(ws, {"A": 40, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek senaryo kütüphanesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:I3")
    h(ws, 4, 1, "Bu sayfa demo değerleri saklar; GIRDI'ye kopyalanabilir.", yazi=GRİ)
    sutunlar = ["Senaryo", "Ticari Kâr", "KKEG", "İstisna", "İndirim", "Ödenen KV",
                "Matrah", "Asgari KV", "Fark"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Matrah": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "tblOrnek[[#This Row],[Ticari Kâr]]+tblOrnek[[#This Row],[KKEG]]"
            "-tblOrnek[[#This Row],[İstisna]]-tblOrnek[[#This Row],[İndirim]])"
        ),
        "Asgari KV": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(ROUND(tblOrnek[[#This Row],[Matrah]]*INDEX(tblKurallar[deger_2025],"
            "MATCH(\"AKV-003\",tblKurallar[kural_id],0)),2),0))"
        ),
        "Fark": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(MAX(0,tblOrnek[[#This Row],[Asgari KV]]-tblOrnek[[#This Row],[Ödenen KV]]),0))"
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:I1004", sutunlar, formuller)
    ornekler = [
        ("Demo ana", 10_000_000, 500_000, 200_000, 100_000, 800_000),
        ("Yüksek ödeme", 10_000_000, 500_000, 200_000, 100_000, 1_200_000),
        ("Düşük matrah", 3_000_000, 100_000, 50_000, 50_000, 200_000),
        ("İstisna ağır", 8_000_000, 200_000, 1_500_000, 100_000, 500_000),
        ("Sıfır fark", 5_000_000, 0, 0, 0, 500_000),
    ]
    for i, row in enumerate(ornekler, 6):
        for k, v in enumerate(row, 1):
            ws.cell(i, k).value = v
            if k > 1:
                ws.cell(i, k).number_format = TL
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
        ws.cell(i, 1).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
        ws.cell(i, 1).protection = Protection(locked=False)
        for kolon in range(1, 7):
            _kim(ws, f"{get_column_letter(kolon)}{i}", sutunlar[kolon - 1], "Örnek senaryo değeri")

    for col, ad in [("B", "Ticari Kâr"), ("C", "KKEG"), ("D", "İstisna"),
                    ("E", "İndirim"), ("F", "Ödenen KV")]:
        dogrulama(ws, "decimal", "0", f"{col}6:{col}1004",
                  baslik=ad, mesaj=f"{ad} tutarını TL girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ile 1e12 arasında.",
                  isaret="between", f2="1000000000000")

    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6])
    h(ws, 5, 12, "FarkDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([220000, 0, 100000, 150000, 0], 6):
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
    h(ws, 5, 3, "Yorum A/B", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate(["A", "B"], 6):
        h(ws, i, 3, v, zemin=GIRIS_SARI)
        ws.cell(i, 3).protection = Protection(locked=False)
        _kim(ws, f"C{i}", "Yorum", "İstisna kırılım yorumu")
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
        ("hesapYili", DEMO["hesapYili"], "yıl", "Aktif hesap yılı", "Kullanıcı / GIRDI", "01.01.2025", "10.08.2026"),
        ("raporTarihi", RAPOR_TARIH, "tarih", "Rapor tarihi (sabit)", "Üretim", "01.01.2025", "10.08.2026"),
        ("akv_tolerans", 0.01, "TL", "Vaka tutarlılık toleransı", "Uygulama notu", "01.01.2025", "10.08.2026"),
        ("akv_farkEsik", 100_000, "TL", "PANO fark uyarı eşiği", "İç politika", "01.01.2025", "10.08.2026"),
        ("akv_senaryoIyi", 0.90, "çarpan", "İyimser matrah çarpanı", "Model varsayımı", "01.01.2025", "10.08.2026"),
        ("akv_senaryoKotu", 1.10, "çarpan", "Kötümser matrah çarpanı", "Model varsayımı", "01.01.2025", "10.08.2026"),
        ("akv_oranArtisPuan", 0.01, "oran", "M-SEN +1 puan", "Senaryo motoru", "01.01.2025", "10.08.2026"),
        ("akv_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "10.08.2026"),
        ("akv_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "10.08.2026"),
        ("akv_tornadoMatrah", 1.05, "çarpan", "Tornado matrah şoku", "Model varsayımı", "01.01.2025", "10.08.2026"),
        ("akv_tornadoOdenen", 0.95, "çarpan", "Tornado ödenen şoku", "Model varsayımı", "01.01.2025", "10.08.2026"),
        ("akv_yorumAIstisnaCarpan", 0.5, "çarpan", "Yorum A istisna çarpanı", "Belirsizlik A", "01.01.2025", "10.08.2026"),
        ("akv_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "10.08.2026"),
        ("akv_parmakIzi", "AKV-PRO-1.0.0", "metin", "Dosya parmak izi etiketi", "Üretim", "01.01.2025", "10.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv", "01.01.2025", "10.08.2026"),
        ("akv_girisBeklenen", 9, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi", "01.01.2025", "10.08.2026"),
        ("akv_riskDusuk", 15, "puan", "Düşük risk skoru", "İç politika", "01.01.2025", "10.08.2026"),
        ("akv_riskOrta", 55, "puan", "Orta risk skoru", "İç politika", "01.01.2025", "10.08.2026"),
        ("akv_riskYuksek", 85, "puan", "Yüksek risk skoru", "İç politika", "01.01.2025", "10.08.2026"),
        ("akv_riskEsikOran", 0.5, "oran", "Yüksek risk fark/asgari eşiği", "İç politika", "01.01.2025", "10.08.2026"),
        ("akv_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı (hesapYili−taban)", "Ürün mimarisi", "01.01.2025", "10.08.2026"),
        ("akv_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi", "01.01.2025", "10.08.2026"),
        ("vaka_matrah_1", 10_200_000, "TL", "Vaka1 matrah", "Tebliğ örneği", "01.01.2025", "10.08.2026"),
        ("vaka_odenen_1", 800_000, "TL", "Vaka1 ödenen", "Tebliğ örneği", "01.01.2025", "10.08.2026"),
        ("vaka_matrah_2", 10_200_000, "TL", "Vaka2 matrah", "Tebliğ örneği", "01.01.2025", "10.08.2026"),
        ("vaka_odenen_2", 1_020_000, "TL", "Vaka2 ödenen", "Tebliğ örneği", "01.01.2025", "10.08.2026"),
        ("vaka_matrah_3", 5_000_000, "TL", "Vaka3 matrah", "Tebliğ örneği", "01.01.2025", "10.08.2026"),
        ("vaka_odenen_3", 400_000, "TL", "Vaka3 ödenen", "Tebliğ örneği", "01.01.2025", "10.08.2026"),
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

    # tblAyarlar YOK — Ö2 genişletmesi named range hedeflerini bozar; D10 başlık satırı yeter.
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
        "EK VERGİ YOK / UYGUN — ödenen KV ≥ asgari KV (fark = 0).",
        "FARK HESAPLA / DİKKAT — ödenen KV < asgari KV; tamamlayıcı fark üretilir.",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 9, 1, "Belirsizlik beyanları (M05)", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari(["istisna_kirilim_sirasi"]), 10):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 12, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      "İstisnanın bir kısmı asgari matrah hesabında dikkate alınır (çarpan AYARLAR'da). "
      "Bu nokta mevzuatta tartışmalıdır; sonuç yorum B ile yan yana okunmalıdır.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B12:J12")
    h(ws, 13, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2,
      "İstisnanın tamamı matrahtan düşülür. Dosya her iki yorumun asgari KV sonucunu MOTOR'da üretir.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B13:J13")

    h(ws, 15, 1, "Kullanım sırası", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 1,
      "GIRDI → KURALLAR (gerekirse) → MOTOR (inceleme) → SENARYO → VAKALAR → PANO → KANIT_RAPORU → AYARLAR.",
      kaydir=True, yazi="333333")
    ws.merge_cells("A16:J16")
    h(ws, 18, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
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
    ws.conditional_formatting.add("E4", CellIsRule(operator="greaterThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("E4", CellIsRule(operator="equal", formula=["0"], font=yesil))
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
        ws.conditional_formatting.add(f"{col}9:{col}13", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(f"{col}9:{col}13", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("F22:F30", FormulaRule(formula=['F22="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F22:F30", FormulaRule(formula=['F22="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F6:F9", CellIsRule(operator="notEqual", formula=["0"], font=amber))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="greaterThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="equal", formula=["0"], font=yesil))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("FARK",E6))'], font=amber))
    ws.conditional_formatting.add("B16:B18", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add("B5:B11", FormulaRule(formula=['OR(B5="KIRIK",B5="HATALI",B5="NEGATİF",B5="BOŞ",B5="EKSİK",B5="AŞIYOR")'],
                                                        font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B5:B11", FormulaRule(formula=['OR(B5="TEMİZ",B5="DOLU",B5="NORMAL",B5="TUTARLI")'],
                                                        font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B12", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("B12", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("G7:G50", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("F7:F50", FormulaRule(formula=['F7="AKV-003"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F50", FormulaRule(formula=['F7="AKV-004"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "C", "D", "E", "F", "G", "H", "I"):
        ws.conditional_formatting.add(f"{harf}6:{harf}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("I6:I20", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B19", FormulaRule(formula=['ISNUMBER(SEARCH("UYGUN",B19))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B19", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B19))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B18", CellIsRule(operator="greaterThan", formula=["0"], font=kirmizi))


def ad_tanimlari(wb):
    # GIRDI / MOTOR köprüleri
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "akv_matrah", "GIRDI!$B$16")
    ad_ekle(wb, "akv_odenenKv", "GIRDI!$B$13")
    ad_ekle(wb, "akv_asgariKv", "MOTOR!$G$22")
    ad_ekle(wb, "akv_fark", "MOTOR!$G$24")
    ad_ekle(wb, "akv_kararMetni", "MOTOR!$G$27")
    ad_ekle(wb, "akv_kararGerekce", "MOTOR!$G$28")
    ad_ekle(wb, "akv_maddeAtifMetni", "MOTOR!$G$49")
    ad_ekle(wb, "akv_kanitRaporu", "MOTOR!$G$50")
    ad_ekle(wb, "akv_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "akv_parmakIzi", "AYARLAR!$B$19")

    ad_ekle(wb, "akv_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "akv_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "akv_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "akv_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "akv_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "akv_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "akv_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "akv_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "akv_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "akv_vakaFark", "VAKALAR!$B$13")

    # AYARLAR parametre adları (satır 6+)
    ayar_map = {
        "akv_tolerans": 8,
        "akv_farkEsik": 9,
        "akv_senaryoIyi": 10,
        "akv_senaryoKotu": 11,
        "akv_oranArtisPuan": 12,
        "akv_tahminAlt": 13,
        "akv_tahminUst": 14,
        "akv_tornadoMatrah": 15,
        "akv_tornadoOdenen": 16,
        "akv_yorumAIstisnaCarpan": 17,
        "akv_yuzdelikOran": 18,
        "akv_girisBeklenen": 21,
        "akv_riskDusuk": 22,
        "akv_riskOrta": 23,
        "akv_riskYuksek": 24,
        "akv_riskEsikOran": 25,
        "akv_yilTaban": 26,
        "akv_yuvarMax": 27,
        "vaka_matrah_1": 28,
        "vaka_odenen_1": 29,
        "vaka_matrah_2": 30,
        "vaka_odenen_2": 31,
        "vaka_matrah_3": 32,
        "vaka_odenen_3": 33,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    # Modül adları (PANO)
    modul_map = {
        "akv_modulT1": (11, "C"), "akv_modulT1Yorum": (11, "D"),
        "akv_modulT2": (12, "C"), "akv_modulT2Yorum": (12, "D"),
        "akv_modulO1": (13, "C"), "akv_modulO1Yorum": (13, "D"),
        "akv_modulO6": (16, "C"), "akv_modulO6Yorum": (16, "D"),
        "akv_modulO8": (17, "C"), "akv_modulO8Yorum": (17, "D"),
        "akv_modulI2": (20, "C"), "akv_modulI2Yorum": (20, "D"),
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

    # PANO önce oluşturuldu ama formüller MOTOR'a bağlı — sıra OK (Excel lazy)
    # Ancak pano() MOTOR adımlarına satır numarasıyla bağlı — motor() return etmeliydi
    # Yeniden: MOTOR satır numaraları sabit (adım i → satır 6+i)

    ad_tanimlari(wb)
    kosullu_bicimlendirme(wb)
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    # Parmak izini kayıt sonrası güncelleyeceğiz; şimdilik etiket
    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    wb.calculation.fullCalcOnLoad = True

    kok = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    varsayilan = os.path.join(kok, "AsgariKurumlarVergisiSimulasyonMotoru.xlsx")
    hedef = cikti_yolu or varsayilan
    os.makedirs(os.path.dirname(os.path.abspath(hedef)) or ".", exist_ok=True)
    wb.save(hedef)

    # SHA-256 yaz
    hsh = hashlib.sha256()
    with open(hedef, "rb") as f:
        for b in iter(lambda: f.read(65536), b""):
            hsh.update(b)
    dig = hsh.hexdigest()

    # Parmak izini AYARLAR'a yaz (yeniden aç-kaydet)
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
    # cikti/ kopyası
    cikti = os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))),
        "cikti", "AsgariKurumlarVergisiSimulasyonMotoru.xlsx",
    )
    # repo kökü: .../excelarsiv-automation
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
    cikti = os.path.join(repo, "cikti", "AsgariKurumlarVergisiSimulasyonMotoru.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
