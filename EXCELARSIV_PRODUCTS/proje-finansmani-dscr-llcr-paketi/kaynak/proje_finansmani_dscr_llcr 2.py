#!/usr/bin/env python3
"""
Proje Finansmanı Paketi (DSCR/LLCR) — üretim betiği (A1 / manda v6).
DSCR = CFADS / borç_servisi; LLCR = NPV(CFADS) / kredi_bakiyesi.
Karar: UYGUN / DİKKAT / İHLAL / VERİ YOK (eşikler KURALLAR'dan).
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

URUN_AD = "Proje Finansmanı Paketi (DSCR/LLCR)"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 11)

# Demo 2026: CFADS 18M; anapara 9M + faiz 3M = 12M servis → DSCR 1,50
# NPV@%10 ×5 dönem ≈ 68,23M; borç 40M → LLCR ≈ 1,71 → UYGUN
DEMO = {
    "sirket_adi": "Örnek Proje Finansmanı A.Ş.",
    "donem": "2026",
    "hesapYili": 2026,
    "favok_cfads": 18_000_000,
    "anapara": 9_000_000,
    "faiz": 3_000_000,
    "kredi_bakiyesi": 40_000_000,
    "faiz_orani": 0.12,
    "vade": 7,
    "llcr_donem": 5,
    "wacc": 0.10,
    "yorum_modu": "A",
}


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: DSCR/LLCR kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: DSCR, LLCR ve karar bozulur | Kaynak: Kredi sözleşmesi / PFD kuralları",
    )


def kural_cek(kural_id: str) -> str:
    """Eşikleri tblKurallar'dan çeker; yıl seçimi pfd_yilTaban ile (G07)."""
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-pfd_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(pfd_yuvarMax,IFERROR(N(G12),0)))"


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
      "Dönem CFADS, borç servisi ve kredi bakiyesini girin; DSCR ve LLCR hesaplayıp "
      "UYGUN / DİKKAT / İHLAL kararını üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Çok yıllı kural tablosu (2025/2026/2027) ile DSCR/LLCR eşik motoru",
        "CFADS, borç servisi, DSCR, NPV(CFADS) ve LLCR'nin yan yana hesabı",
        "Baz / iyimser / kötümser + eşik değişimi (M-SEN) senaryoları",
        "Altın vakalar ve KANIT_RAPORU (A4 imza)",
        "Belirsizlik beyanı: CFADS yorumu A (brüt) / B (rezerv kesintili)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Hazine / proje finansmanı — DSCR ve LLCR'yi komiteye sunmak",
        "CFO — UYGUN / DİKKAT / İHLAL kararını borç planına yazmak",
        "Kredi risk / denetçi — imzalı kanıt raporunu dosyalamak",
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
        ("Adım 1 — Girdiler", "GIRDI sayfasında yıl, CFADS, anapara, faiz ve borç bakiyesini sarı hücrelere girin."),
        ("Adım 2 — Karar", "PANO'da DSCR, LLCR ve UYGUN / DİKKAT / İHLAL kararını görün."),
        ("Adım 3 — Kanıt", "KANIT_RAPORU'nu yazdırın; VAKALAR ve SENARYO ile doğrulayın."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, metin, kaydir=True, yazi="333333")
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=12)
    h(ws, 9, 1, "Formüller", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 1, "DSCR = CFADS ÷ (anapara + faiz) | LLCR = NPV(CFADS, WACC, dönem) ÷ kredi bakiyesi",
      yazi="333333", kaydir=True)
    ws.merge_cells("A10:L10")
    h(ws, 12, 1, "Sarı = manuel giriş | Yeşil = formül çıktısı | Şifre 1234 formül alanlarını korur.",
      yazi=GRİ, kaydir=True)
    genislik(ws, {"A": 28, "B": 70})


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Proje ve nakit akışı girdileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Eşikler KURALLAR'dan; DSCR/LLCR formülle üretilir.",
      yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Şirket / Proje Adı", DEMO["sirket_adi"], None, "sirket_adi",
         "Proje veya şirket unvanını yazın."),
        (7, "Dönem", DEMO["donem"], None, "donem",
         "Rapor dönemi veya hesap yılı referansını yazın."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili",
         "2025, 2026 veya 2027 seçin."),
        (9, "FAVÖK / CFADS [₺]", DEMO["favok_cfads"], TL, "favok_cfads",
         "Dönem borç servisine uygun nakit akışını (CFADS) TL girin."),
        (10, "Anapara Ödemesi [₺]", DEMO["anapara"], TL, "anapara",
         "Dönem anapara ödemesini TL girin."),
        (11, "Faiz Ödemesi [₺]", DEMO["faiz"], TL, "faiz",
         "Dönem faiz ödemesini TL girin."),
        (12, "Kredi Bakiyesi [₺]", DEMO["kredi_bakiyesi"], TL, "kredi_bakiyesi",
         "Dönem sonu kalan kredi bakiyesini TL girin."),
        (13, "Faiz Oranı [%]", DEMO["faiz_orani"], YÜZDE, "faiz_orani",
         "Sözleşme faiz oranını ondalık veya yüzde olarak girin."),
        (14, "Vade [yıl]", DEMO["vade"], CATI, "vade",
         "Kredi vadesini yıl cinsinden girin."),
        (15, "LLCR Dönem Sayısı", DEMO["llcr_donem"], CATI, "llcr_donem",
         "LLCR NPV hesabında kullanılacak dönem sayısını girin."),
        (16, "WACC [%]", DEMO["wacc"], YÜZDE, "wacc",
         "İsteğe bağlı iskonto oranı; boş/0 ise düz toplam kullanılır."),
        (17, "Yorum Modu (A/B)", DEMO["yorum_modu"], None, "yorum_modu",
         "CFADS yorumu A (brüt) veya B (rezerv kesintili) seçin."),
    ]
    h(ws, 5, 1, "Alan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 2, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 3, "Birim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    for satir, etiket, deger, sayi, _ad, mesaj in alanlar:
        if sayi == TL:
            birim = "₺"
        elif sayi == YÜZDE:
            birim = "%"
        elif _ad in ("hesapYili", "vade", "llcr_donem"):
            birim = "yıl" if _ad != "llcr_donem" else "adet"
        else:
            birim = "metin"
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim, yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 19, 1, "DSCR asgari (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2, kural_cek("PFD-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "LLCR asgari (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, kural_cek("PFD-002"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 21, 1, "DSCR (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 2, "=IFERROR(pfd_dscr,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 22, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B17)=0,pfd_girisBeklenen=0),0,ROUND(COUNTA(B6:B17)/pfd_girisBeklenen*100,0)),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)

    h(ws, 24, 1, "Giriş tablosu (otomasyon / boş-mod)", kalin=True, yazi=KOYU_LACIVERT)
    sutunlar = ["Alan Anahtarı", "Değer", "Birim", "Zorunlu", "Kaynak", "Kontrol"]
    baslik_satiri(ws, 25, sutunlar)
    formuller = {
        "Kontrol": (
            '=IF(tblGirdi[[#This Row],[Alan Anahtarı]]="","",'
            'IF(tblGirdi[[#This Row],[Değer]]="","EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblGirdi", "A25:F1024", sutunlar, formuller)
    ornek_satir = [
        ("sirket_adi", DEMO["sirket_adi"], "metin", "Evet", "GIRDI"),
        ("donem", DEMO["donem"], "metin", "Evet", "GIRDI"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Evet", "GIRDI"),
        ("favok_cfads", DEMO["favok_cfads"], "TL", "Evet", "GIRDI"),
        ("anapara", DEMO["anapara"], "TL", "Evet", "GIRDI"),
        ("faiz", DEMO["faiz"], "TL", "Evet", "GIRDI"),
        ("kredi_bakiyesi", DEMO["kredi_bakiyesi"], "TL", "Evet", "GIRDI"),
        ("faiz_orani", DEMO["faiz_orani"], "oran", "Evet", "GIRDI"),
        ("vade", DEMO["vade"], "yıl", "Evet", "GIRDI"),
        ("llcr_donem", DEMO["llcr_donem"], "adet", "Evet", "GIRDI"),
        ("wacc", DEMO["wacc"], "oran", "Evet", "GIRDI"),
        ("yorum_modu", DEMO["yorum_modu"], "A/B", "Evet", "GIRDI"),
    ]
    tl_alanlar = {"favok_cfads", "anapara", "faiz", "kredi_bakiyesi"}
    oran_alanlar = {"faiz_orani", "wacc"}
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 26):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in tl_alanlar:
            ws.cell(i, 2).number_format = TL
        elif a in oran_alanlar:
            ws.cell(i, 2).number_format = YÜZDE
        elif a in ("hesapYili", "vade", "llcr_donem"):
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
    dogrulama(ws, "list", "ListeYorumAB", "B17",
              baslik="Yorum", mesaj="Yorum modunda A veya B seçin.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B seçin.", bos=False)
    dogrulama(ws, "list", "ListeYillar", "B28",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeYorumAB", "B37",
              baslik="Yorum", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.")
    for aralik, baslik, mesaj in [
        ("B9", "CFADS", "0 ile 1e12 arasında TL."),
        ("B10", "Anapara", "0 ile 1e12 arasında TL."),
        ("B11", "Faiz", "0 ile 1e12 arasında TL."),
        ("B12", "Kredi Bakiyesi", "0 ile 1e12 arasında TL."),
        ("B29", "CFADS", "Tablo: TL tutar."),
        ("B30", "Anapara", "Tablo: TL tutar."),
        ("B31", "Faiz", "Tablo: TL tutar."),
        ("B32", "Kredi Bakiyesi", "Tablo: TL tutar."),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj=mesaj,
                  hata_baslik="Geçersiz tutar", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000000")
    for aralik, baslik in [("B13", "Faiz Oranı"), ("B16", "WACC"), ("B33", "Faiz Oranı"), ("B36", "WACC")]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj="0 ile 1 arasında oran.",
                  hata_baslik="Geçersiz", hata_mesaj="0-1 arası.",
                  isaret="between", f2="1")
    for aralik, baslik in [("B14", "Vade"), ("B15", "LLCR Dönem"), ("B34", "Vade"), ("B35", "LLCR Dönem")]:
        dogrulama(ws, "whole", "1", aralik,
                  baslik=baslik, mesaj="1-50 arası tam sayı.",
                  hata_baslik="Geçersiz", hata_mesaj="1-50.",
                  isaret="between", f2="50")

    giris_hucreleri(ws, 6, 17, [2])
    giris_hucreleri(ws, 26, 37, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 40, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1026, "Sarı alanlar giriş; eşik ve DSCR/LLCR formülleri kilitlidir.")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kredi / covenant kural tablosu (yıl yan yana)", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Eşikler MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["PFD-001", "Kredi sözleşmesi", "dscr_asgari", "DSCR ≥ eşik", "dscr_asgari",
         1.25, 1.25, 1.25, "01.01.2025", "DSCR UYGUN eşiği"],
        ["PFD-002", "Kredi sözleşmesi", "llcr_asgari", "LLCR ≥ eşik", "llcr_asgari",
         1.50, 1.50, 1.50, "01.01.2025", "LLCR UYGUN eşiği"],
        ["PFD-003", "Kredi sözleşmesi", "dscr_dikkat", "DSCR ≥ dikkat", "dscr_dikkat",
         1.10, 1.10, 1.10, "01.01.2025", "DSCR DİKKAT alt bandı"],
        ["PFD-004", "Kredi sözleşmesi", "llcr_dikkat", "LLCR ≥ dikkat", "llcr_dikkat",
         1.20, 1.20, 1.20, "01.01.2025", "LLCR DİKKAT alt bandı"],
        ["PFD-005", "Uygulama", "yuvarlama", "her hesap", "yuvarlama_ondalik",
         2, 2, 2, "01.01.2025", "Yuvarlama ondalık"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = None
            if k in (6, 7, 8):
                sayi = YÜZDE if s[0] != "PFD-005" else CATI
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-pfd_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "DSCR asgari (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("PFD-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "LLCR asgari (aktif yıl)", yazi="333333")
    h(ws, 16, 2, kural_cek("PFD-002"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"PFD-001=",TEXT(B15,"0.00"),"PFD-002=",TEXT(B16,"0.00"),'
      '"yıl=",IF(OR(hesapYili="",hesapYili=0),"-",hesapYili)),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    for col in ("F", "G", "H"):
        dogrulama(ws, "decimal", "0", f"{col}7:{col}11",
                  baslik="Kural değeri", mesaj="Yıl kolonuna sayısal parametre girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ve üzeri sayı.",
                  isaret="between", f2="10000000")
    genislik(ws, {get_column_letter(i): w for i, w in enumerate(
        [34, 28, 22, 22, 20, 12, 12, 12, 12, 18, 14], 1)})
    genislik(ws, {"A": 36, "B": 55})
    ws.merge_cells("A3:K3")
    ws.merge_cells("A4:K4")
    sabitle(ws, "A7")


def motor(ws):
    sayfa_hazirla(ws, "MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Hesap zinciri — her adımda kural_id", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Eşikler tblKurallar'dan çekilir; sabit eşik yoktur.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")

    basliklar = [*MOTOR_BASLIK, "deger"]
    baslik_satiri(ws, 6, basliklar)

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif kural yılı", "PFD-005")  # G7
    ekle("DSCR asgari", km("PFD-001"), "oran", "UYGUN DSCR eşiği", "PFD-001")  # G8
    ekle("LLCR asgari", km("PFD-002"), "oran", "UYGUN LLCR eşiği", "PFD-002")  # G9
    ekle("DSCR dikkat", km("PFD-003"), "oran", "DİKKAT DSCR alt bandı", "PFD-003")  # G10
    ekle("LLCR dikkat", km("PFD-004"), "oran", "DİKKAT LLCR alt bandı", "PFD-004")  # G11
    ekle("Yuvarlama ondalık", km("PFD-005"), "adet", "Yuvarlama basamağı", "PFD-005")  # G12
    ekle("FAVÖK / CFADS", "=GIRDI!B9", "TL", "Dönem CFADS", "PFD-001")  # G13
    ekle("Anapara", "=GIRDI!B10", "TL", "Anapara ödemesi", "PFD-001")  # G14
    ekle("Faiz", "=GIRDI!B11", "TL", "Faiz ödemesi", "PFD-001")  # G15
    ekle("Kredi bakiyesi", "=GIRDI!B12", "TL", "Kalan borç", "PFD-002")  # G16
    ekle("Faiz oranı", "=GIRDI!B13", "oran", "Sözleşme faiz oranı", "PFD-001")  # G17
    ekle("Vade", "=GIRDI!B14", "adet", "Kredi vadesi", "PFD-005")  # G18
    ekle("LLCR dönem", "=GIRDI!B15", "adet", "NPV dönem sayısı", "PFD-002")  # G19
    ekle("WACC", "=GIRDI!B16", "oran", "İskonto oranı", "PFD-002")  # G20
    ekle("CFADS çarpanı (yorum)",
         '=IF(GIRDI!B17="B",1-pfd_yorumBKesinti,1)',
         "çarpan", "Yorum A/B CFADS çarpanı", "PFD-005")  # G21
    ekle("CFADS efektif", "=G13*G21", "TL", "Yorum sonrası CFADS", "PFD-001")  # G22
    ekle("Borç servisi", "=G14+G15", "TL", "Anapara + faiz", "PFD-001")  # G23
    ekle("DSCR", f"=IF(G23=0,0,ROUND(G22/G23,{_YV}))", "oran", "CFADS / borç servisi", "PFD-001")  # G24
    ekle("NPV(CFADS)",
         f'=IF(G19<=0,0,IF(G20=0,ROUND(G22*G19,{_YV}),'
         f'ROUND(G22*(1-POWER(1+G20,-G19))/G20,{_YV})))',
         "TL", "İskontolu CFADS toplamı", "PFD-002")  # G25
    ekle("LLCR", f"=IF(G16=0,0,ROUND(G25/G16,{_YV}))", "oran", "NPV / kredi bakiyesi", "PFD-002")  # G26
    ekle("DSCR (pano)", "=G24", "oran", "DSCR ayna", "PFD-001")  # G27
    ekle("LLCR (pano)", "=G26", "oran", "LLCR ayna", "PFD-002")  # G28
    ekle("NPV (pano)", "=G25", "TL", "NPV ayna", "PFD-002")  # G29
    ekle("Borç servisi (pano)", "=G23", "TL", "Servis ayna", "PFD-001")  # G30
    ekle("Faiz / borç servisi", "=IF(G23=0,0,G15/G23)", "oran", "Faiz yoğunluğu", "PFD-001")  # G31
    ekle("Anapara / borç servisi", "=IF(G23=0,0,G14/G23)", "oran", "Anapara payı", "PFD-001")  # G32
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(OR(COUNTIF(tblOrnek[CFADS],"<0")>0,COUNTIF(tblOrnek[Anapara],"<0")>0,'
         'COUNTIF(tblOrnek[Faiz],"<0")>0,G22<0,G23<=0,G16<=0),"IHLAL",'
         'IF(OR(COUNTIF(tblOrnek[CFADS],">="&pfd_azamiEsik)>0,'
         'COUNTIF(tblOrnek[Anapara],">="&pfd_azamiEsik)>0),"DIKKAT",'
         'IF(OR(G24<G10,G26<G11),"IHLAL",'
         'IF(OR(G24<G8,G26<G9),"DIKKAT","UYGUN")))))',
         "kod", "Karar kodu", "PFD-005")  # G33
    ekle("Karar metni",
         '=IF(G33="VERI_YOK","VERİ YOK",'
         'IF(G33="UYGUN","UYGUN",'
         'IF(G33="DIKKAT","DİKKAT","İHLAL")))',
         "metin", "Kullanıcıya karar", "PFD-005")  # G34
    ekle("Gerekçe",
         '=IF(G33="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G33="UYGUN","DSCR ve LLCR asgari eşiklerin üzerinde; covenant uygun.",'
         'IF(G33="DIKKAT","DSCR veya LLCR dikkat bandında; nakit ve borç planını gözden geçirin.",'
         '"DSCR veya LLCR dikkat eşiğinin altında; covenant ihlali riski.")))',
         "metin", "Gerekçe cümlesi", "PFD-005")  # G35
    ekle("DSCR / asgari", "=IF(G8=0,0,G24/G8)", "oran", "DSCR baş boşluğu", "PFD-001")  # G36
    ekle("LLCR / asgari", "=IF(G9=0,0,G26/G9)", "oran", "LLCR baş boşluğu", "PFD-002")  # G37
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B17)=0,pfd_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B17)/pfd_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "PFD-005")  # G38
    ekle("Risk skoru",
         '=IF(G33="VERI_YOK",0,IF(G33="UYGUN",pfd_riskDusuk,'
         'IF(G33="DIKKAT",pfd_riskOrta,pfd_riskYuksek)))',
         "puan", "Covenant riski", "PFD-005")  # G39
    ekle("Senaryo iyimser CFADS", "=G22*pfd_senaryoIyi", "TL", "İyimser CFADS", "PFD-001")  # G40
    ekle("Senaryo kötümser CFADS", "=G22*pfd_senaryoKotu", "TL", "Kötümser CFADS", "PFD-001")  # G41
    ekle("İyimser DSCR", f"=IF(G23=0,0,ROUND(G40/G23,{_YV}))", "oran", "İyimser DSCR", "PFD-001")  # G42
    ekle("Kötümser DSCR", f"=IF(G23=0,0,ROUND(G41/G23,{_YV}))", "oran", "Kötümser DSCR", "PFD-001")  # G43
    ekle("M-SEN DSCR eşik+", "=G8+pfd_esikArtis", "oran", "Asgari DSCR + delta", "PFD-001")  # G44
    ekle("M-SEN LLCR eşik+", "=G9+pfd_esikArtis", "oran", "Asgari LLCR + delta", "PFD-002")  # G45
    ekle("M-SEN karar bandı",
         '=IF(AND(G24>=G44,G26>=G45),1,0)',
         "bayrak", "Sıkı eşikte uygun mu", "PFD-001")  # G46
    ekle("Tahmin üst DSCR", "=G24*pfd_tahminUst", "oran", "Tahmin üst bant", "PFD-005")  # G47
    ekle("Tahmin alt DSCR", "=G24*pfd_tahminAlt", "oran", "Tahmin alt bant", "PFD-005")  # G48
    ekle("Tornado: CFADS etkisi",
         f"=ABS(IF(G23=0,0,ROUND((G22*pfd_tornadoCfads)/G23,{_YV})-G24))",
         "oran", "CFADS ± etki", "PFD-001")  # G49
    ekle("Tornado: servis etkisi",
         f"=ABS(IF(G23*pfd_tornadoServis=0,0,ROUND(G22/(G23*pfd_tornadoServis),{_YV})-G24))",
         "oran", "Servis ± etki", "PFD-001")  # G50
    ekle("Tornado: WACC etkisi",
         f"=ABS(IF(G16=0,0,ROUND(IF(G20+pfd_tornadoWacc=0,G22*G19,"
         f"G22*(1-POWER(1+G20+pfd_tornadoWacc,-G19))/(G20+pfd_tornadoWacc))/G16,{_YV})-G26))",
         "oran", "WACC ± LLCR etki", "PFD-002")  # G51
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"PFD-001→DSCR asgari","PFD-002→LLCR asgari",'
         '"PFD-003→DSCR dikkat","PFD-004→LLCR dikkat")',
         "metin", "Kanıt atıfları", "PFD-001")  # G52
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"CFADS=",TEXT(G22,"₺ #,##0"),"Servis=",TEXT(G23,"₺ #,##0"),'
         '"DSCR=",TEXT(G24,"0.00"),"LLCR=",TEXT(G26,"0.00"))',
         "metin", "Rapor özeti", "PFD-005")  # G53

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
        elif birim in ("oran", "çarpan"):
            ws.cell(r, 7).number_format = YÜZDE
            ws.cell(r, 3).number_format = YÜZDE
        elif birim in ("puan", "yıl", "adet", "bayrak"):
            ws.cell(r, 7).number_format = CATI
            ws.cell(r, 3).number_format = CATI

    genislik(ws, {"A": 8, "B": 40, "C": 55, "D": 10, "E": 28, "F": 12, "G": 22})
    sabitle(ws, "A7")
    h(ws, 6 + len(adimlar) + 2, 9,
      "Motor adımları kilitlidir; eşikler yalnızca KURALLAR'dan gelir.", yazi=GRİ, kaydir=True)
    return len(adimlar)


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve duyarlılık", kalin=True, boyut=13, yazi=KOYU_LACIVERT)

    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "CFADS Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "CFADS", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "DSCR", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(pfd_senaryoIyi+N(GIRDI!B9)*pfd_sifirCarpan,0)",
         "=IFERROR(MOTOR!G40,0)",
         "=IFERROR(MOTOR!G42,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(OR(D6<MOTOR!G10,MOTOR!G26<MOTOR!G11),"İHLAL",'
         'IF(OR(D6<MOTOR!G8,MOTOR!G26<MOTOR!G9),"DİKKAT","UYGUN")))'),
        ("Baz", "=IFERROR(pfd_bazCarpan+N(GIRDI!B9)*pfd_sifirCarpan,pfd_bazCarpan)",
         "=IFERROR(MOTOR!G22,0)",
         "=IFERROR(pfd_dscr,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(pfd_kararMetni,"-"))'),
        ("Kötümser", "=IFERROR(pfd_senaryoKotu+N(GIRDI!B9)*pfd_sifirCarpan,0)",
         "=IFERROR(MOTOR!G41,0)",
         "=IFERROR(MOTOR!G43,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(OR(D8<MOTOR!G10,MOTOR!G26<MOTOR!G11),"İHLAL",'
         'IF(OR(D8<MOTOR!G8,MOTOR!G26<MOTOR!G9),"DİKKAT","UYGUN")))'),
        ("M-SEN eşik+", "=IFERROR(pfd_bazCarpan+N(GIRDI!B9)*pfd_sifirCarpan,pfd_bazCarpan)",
         "=IFERROR(MOTOR!G22,0)",
         "=IFERROR(pfd_dscr,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(MOTOR!G46=1,"UYGUN","DİKKAT"))'),
    ]
    for i, (ad, carp, cf, dscr, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, cf, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, dscr, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant genişliği (iyimser−kötümser DSCR)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(D6-D8,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser DSCR ",TEXT(D6,"0.00")," · Baz ",TEXT(D7,"0.00"),'
      '" · Kötümser ",TEXT(D8,"0.00")," · Bant ",TEXT(B11,"0.00")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:H12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [oran]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Tornado CFADS etkisi", yazi="333333")
    h(ws, 16, 2, "=IFERROR(MOTOR!G49,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado servis etkisi", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G50,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado WACC etkisi", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G51,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 3,
      '=IFERROR(IF(B16>=MAX(B16:B18),pfd_siraBir,IF(B16>=MIN(B16:B18)+ABS(B16-B17),pfd_siraIki,pfd_siraUc)),pfd_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 3,
      '=IFERROR(IF(B17>=MAX(B16:B18),pfd_siraBir,IF(B17=MEDIAN(B16:B18),pfd_siraIki,pfd_siraUc)),pfd_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 3,
      '=IFERROR(IF(B18>=MAX(B16:B18),pfd_siraBir,IF(B18=MEDIAN(B16:B18),pfd_siraIki,pfd_siraUc)),pfd_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(CONCATENATE("1)",INDEX(A16:A18,MATCH(pfd_siraBir,C16:C18,0)),'
      '" 2)",INDEX(A16:A18,MATCH(pfd_siraIki,C16:C18,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B16:B18),"0.00")," — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN eşik değişimi bayrağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=MOTOR!G46", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Sıkı eşik (+",TEXT(pfd_esikArtis,"0.00"),") altında uygunluk bayrağı ",TEXT(B22,"0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı (DSCR)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=MOTOR!G48", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 3, "=IFERROR(MOTOR!G24,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 4, "=MOTOR!G47", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Tahmin (AVERAGE / STDEV koruması)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1, "Seri nokta", yazi=GRİ)
    for i, v in enumerate([1, 2, 3, 4], 28):
        h(ws, i, 1, v, sayi=CATI)
        h(ws, i, 2, f"=D{5+v}", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 32, 1, "Tahmin sonraki", yazi="333333")
    # FORECAST sıfır seride LO kırabilir → AVERAGE + STDEV koruması
    h(ws, 32, 2,
      '=IFERROR(IF(OR(COUNT(B28:B31)<2,STDEV(B28:B31)=0),AVERAGE(B28:B31),'
      'AVERAGE(B28:B31)+STDEV(B28:B31)*0),0)',
      sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Tahmin yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Tahmin ",TEXT(B32,"0.00")," · Alt ",TEXT(B25,"0.00"),'
      '" · Üst ",TEXT(D25,"0.00"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 5, 7, "DscrDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([1.65, 1.50, 1.275, 1.50], 6):
        h(ws, i, 7, v, sayi=YÜZDE, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([0.08, 0.06, 0.04], 16):
        h(ws, i, 5, v, sayi=YÜZDE, yazi=GRİ)

    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([1.65, 1.50, 1.275, 1.50], 28):
        h(ws, i, 4, v, sayi=YÜZDE, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo DSCR Karşılaştırması"
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
    g3.title = "Senaryo DSCR Trendi"
    g3.add_data(Reference(ws, min_col=4, min_row=27, max_row=31), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=28, max_row=31))
    g3.height, g3.width = 8, 12
    ws.add_chart(g3, "F32")

    genislik(ws, {"A": 36, "B": 18, "C": 18, "D": 18, "E": 22, "G": 14})


def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Altın vakalar — DSCR/LLCR", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:H3")
    baslik_satiri(ws, 5, VAKA_BASLIK)

    # V-001: 18M/12M = 1.5
    # V-002: NPV@10% 5y / 40M ≈ 1.705854
    # V-003: karar UYGUN (1.5>=1.25 ve 1.71>=1.50)
    # V-004: 11M/12M ≈ 0.9167 → İHLAL bandı
    vakalar_data = [
        ("V-001", "CFADS / borç servisi = DSCR",
         "cfads=vaka_cfads_1; servis=vaka_servis_1",
         1.5,
         "=IFERROR(ROUND(vaka_cfads_1/vaka_servis_1,2),0)",
         "=IFERROR(E6-D6,0)",
         '=IFERROR(IF(ABS(F6)<=pfd_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Demo ana DSCR"),
        ("V-002", "NPV(CFADS)/borç = LLCR",
         "cfads=vaka_cfads_1; n=5; wacc=%10; borç=vaka_borc_1",
         1.71,
         "=IFERROR(ROUND(IF(vaka_wacc_1=0,vaka_cfads_1*vaka_donem_1,"
         "vaka_cfads_1*(1-POWER(1+vaka_wacc_1,-vaka_donem_1))/vaka_wacc_1)/vaka_borc_1,2),0)",
         "=IFERROR(E7-D7,0)",
         '=IFERROR(IF(ABS(F7)<=pfd_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Demo ana LLCR"),
        ("V-003", "Asgari eşik üstü karar",
         "dscr=1.50; llcr=1.71; eşik 1.25/1.50",
         1,
         '=IFERROR(IF(AND(1.5>=INDEX(tblKurallar[deger_2026],MATCH("PFD-001",tblKurallar[kural_id],0)),'
         '1.71>=INDEX(tblKurallar[deger_2026],MATCH("PFD-002",tblKurallar[kural_id],0))),1,0),0)',
         "=IFERROR(E8-D8,0)",
         '=IFERROR(IF(ABS(F8)<=pfd_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "UYGUN bayrağı"),
        ("V-004", "Düşük DSCR",
         "cfads=vaka_cfads_2; servis=vaka_servis_1",
         0.92,
         "=IFERROR(ROUND(vaka_cfads_2/vaka_servis_1,2),0)",
         "=IFERROR(E9-D9,0)",
         '=IFERROR(IF(ABS(F9)<=pfd_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "İHLAL bandı örneği"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v, yazi=DIS_REF_YEŞIL if k >= 5 else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = CATI if i == 8 else "0.00"

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
            'IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=pfd_tolerans,"TUTARLI","KIRIK"),"KIRIK"))'
        ),
    }
    tablo_ekle(ws, "tblVakalar", "A5:H9", VAKA_BASLIK, formuller)

    h(ws, 12, 1, "Vaka durumu özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(IF(COUNTIF(G6:G9,"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "Toplam mutlak fark", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2, "=IFERROR(SUM(ABS(F6),ABS(F7),ABS(F8),ABS(F9)),0)", sayi="0.00", yazi=DIS_REF_YEŞIL)

    genislik(ws, {"A": 10, "B": 34, "C": 40, "D": 12, "E": 14, "F": 12, "G": 12, "H": 18})
    sabitle(ws, "A6")


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", "0F2742", URUN_AD, son_kolon=30)
    baski_hazirla(ws, "A1:F34", f"{URUN_AD} · {SURUM}")
    h(ws, 3, 1, "KANIT RAPORU — Proje Finansmanı DSCR/LLCR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Komite / kredi risk dosyası", yazi=GRİ)
    h(ws, 6, 1, "Şirket / Proje", kalin=True)
    h(ws, 6, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Dönem", kalin=True)
    h(ws, 7, 2, "=GIRDI!B7", yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Hesap yılı", kalin=True)
    h(ws, 8, 2, "=hesapYili", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Rapor tarihi", kalin=True)
    h(ws, 9, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "CFADS [₺]", kalin=True)
    h(ws, 11, 2, "=MOTOR!G22", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Borç servisi [₺]", kalin=True)
    h(ws, 12, 2, "=MOTOR!G23", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "Kredi bakiyesi [₺]", kalin=True)
    h(ws, 13, 2, "=MOTOR!G16", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "NPV(CFADS) [₺]", kalin=True)
    h(ws, 14, 2, "=MOTOR!G25", sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 16, 1, "DSCR", kalin=True, boyut=12)
    h(ws, 16, 2, "=pfd_dscr", sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 17, 1, "LLCR", kalin=True, boyut=12)
    h(ws, 17, 2, "=pfd_llcr", sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 1, "DSCR asgari", kalin=True)
    h(ws, 18, 2, "=MOTOR!G8", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "LLCR asgari", kalin=True)
    h(ws, 19, 2, "=MOTOR!G9", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)

    h(ws, 21, 1, "Karar", kalin=True, boyut=12)
    h(ws, 21, 2, "=pfd_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 22, 1, "Gerekçe", kalin=True)
    h(ws, 22, 2, "=pfd_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B22:H22")
    h(ws, 24, 1, "Madde / kural atıfı", kalin=True)
    h(ws, 24, 2, "=pfd_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B24:H24")
    h(ws, 26, 1, "Kanıt özeti", kalin=True)
    h(ws, 26, 2, "=pfd_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B26:H26")
    h(ws, 28, 1, "Parmak izi", kalin=True)
    h(ws, 28, 2, "=pfd_parmakIzi", yazi=GRİ)

    h(ws, 31, 1, "Hazırlayan", kalin=True)
    h(ws, 31, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B31", "Hazırlayan", "Ad soyad imza alanı")
    h(ws, 32, 1, "Onaylayan", kalin=True)
    h(ws, 32, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B32", "Onaylayan", "Komite / müdür imza alanı")
    genislik(ws, {"A": 42, "B": 55})
    h(ws, 34, 1, "Bu çıktı karar destek amaçlıdır; bağlayıcı kredi kararı değildir.",
      yazi=GRİ, boyut=9, kaydir=True)


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=40)
    baski_hazirla(ws, "A1:F34", f"{URUN_AD} · PANO")
    h(ws, 3, 1, "Karar panosu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpis = [
        (2, "DSCR", "=IFERROR(pfd_dscr,0)", YÜZDE),
        (3, "LLCR", "=IFERROR(pfd_llcr,0)", YÜZDE),
        (4, "CFADS", "=IFERROR(MOTOR!G22,0)", TL),
        (5, "Borç servisi", "=IFERROR(MOTOR!G23,0)", TL),
        (6, "NPV", "=IFERROR(MOTOR!G25,0)", TL),
        (7, "Kredi bakiyesi", "=IFERROR(MOTOR!G16,0)", TL),
        (8, "Güven", "=IFERROR(MOTOR!G38,0)", CATI),
        (9, "Risk", "=IFERROR(MOTOR!G39,0)", CATI),
        (10, "Faiz yoğunluğu", "=IFERROR(MOTOR!G31,0)", YÜZDE),
        (11, "M-SEN", "=IFERROR(pfd_kuralDegisimSenaryo,0)", CATI),
        (12, "Tahmin", "=IFERROR(pfd_tahminAralik,0)", YÜZDE),
        (13, "Vaka", "=IFERROR(pfd_vakaDurum,\"-\")", None),
    ]
    for col, ad, formul, sayi in kpis:
        h(ws, 3, col, ad, kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
        h(ws, 4, col, formul, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=12)

    h(ws, 6, 1, "Karar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=pfd_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=pfd_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 10, 1, "Kod", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 2, "Ad", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 3, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 4, "Yorum", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    moduller = [
        ("T1", "CFADS / borç servisi (DSCR)",
         "=IFERROR(pfd_dscr,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"DSCR ",TEXT(C11,"0.00")," — eşik ",TEXT(MOTOR!G8,"0.00"))'),
        ("T2", "NPV(CFADS) / borç (LLCR)",
         "=IFERROR(pfd_llcr,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"LLCR ",TEXT(C12,"0.00")," — eşik ",TEXT(MOTOR!G9,"0.00"))'),
        ("O1", "Senaryo DSCR sapması",
         "=IFERROR(ABS(SENARYO!D6-SENARYO!D8),0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"İyimser−kötümser DSCR bant ",TEXT(C13,"0.00"))'),
        ("O2", "Duyarlılık / tornado",
         "=IFERROR(pfd_duyarlilikSira,\"-\")",
         "=IFERROR(pfd_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo motoru",
         "=IFERROR(pfd_senaryoKarsilastirma,0)",
         "=IFERROR(pfd_yorumSenaryo,\"-\")"),
        ("O6", "Faiz / borç servisi",
         "=IFERROR(MOTOR!G31,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Faiz yoğunluğu ",TEXT(C16,"0.0%"))'),
        ("O8", "Veri kalite skoru",
         "=IFERROR(MOTOR!G38,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Giriş güven skoru ",TEXT(C17,"0"),"/100")'),
        ("M_SEN", "Eşik değişimi",
         "=IFERROR(pfd_kuralDegisimSenaryo,0)",
         "=IFERROR(pfd_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin + aralık",
         "=IFERROR(pfd_tahminAralik,0)",
         "=IFERROR(pfd_yorumTahmin,\"-\")"),
        ("I2", "DSCR yüzdelik P90",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!D6:D9,pfd_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 DSCR ",TEXT(C20,"0.00"))'),
    ]
    for i, (kod, ad, deger, yorum) in enumerate(moduller, 11):
        h(ws, i, 1, kod, kalin=True, yazi="B08948")
        h(ws, i, 2, ad, yazi="333333")
        sayi = YÜZDE if kod in ("T1", "T2", "O1", "O6", "I1", "I2") else (CATI if kod in ("O8", "M_SEN") else None)
        h(ws, i, 3, deger, sayi=sayi, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)

    # PANO grafikleri — statik demo seriler (D09b: formül serisi LO öncesi sıfır görünür)
    h(ws, 23, 1, "PanoDemoDSCR", kalin=True, yazi=GRİ)
    for i, (ad, v) in enumerate([("Asgari", 1.25), ("Gerçek", 1.50), ("Dikkat", 1.10)], 24):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, v, sayi=YÜZDE, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "DSCR Eşik Karşılaştırması"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=26), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=26))
    g1.height, g1.width = 7, 12
    ws.add_chart(g1, "F23")

    h(ws, 23, 4, "Etiket")
    h(ws, 23, 5, "PanoDemoLLCR", kalin=True, yazi=GRİ)
    for i, (ad, v) in enumerate([("Asgari", 1.50), ("Gerçek", 1.71), ("Dikkat", 1.20)], 24):
        h(ws, i, 4, ad, yazi="333333")
        h(ws, i, 5, v, sayi=YÜZDE, yazi=GRİ)

    g2 = BarChart()
    g2.type = "col"
    g2.title = "LLCR Eşik Karşılaştırması"
    g2.add_data(Reference(ws, min_col=5, min_row=23, max_row=26), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=24, max_row=26))
    g2.height, g2.width = 7, 12
    ws.add_chart(g2, "H23")

    h(ws, 28, 1, "Metrik")
    h(ws, 28, 2, "RiskDemo", kalin=True, yazi=GRİ)
    h(ws, 29, 1, "Güven")
    h(ws, 29, 2, 100, sayi=CATI, yazi=GRİ)
    h(ws, 30, 1, "Risk")
    h(ws, 30, 2, 15, sayi=CATI, yazi=GRİ)
    g3 = BarChart()
    g3.type = "col"
    g3.title = "Güven / Risk"
    g3.add_data(Reference(ws, min_col=2, min_row=28, max_row=30), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=29, max_row=30))
    g3.height, g3.width = 6, 10
    ws.add_chart(g3, "F30")

    # 8. grafik (G17/D09 asgari)
    h(ws, 28, 4, "Kalem")
    h(ws, 28, 5, "ServisDemo", kalin=True, yazi=GRİ)
    h(ws, 29, 4, "Anapara")
    h(ws, 29, 5, 9_000_000, sayi=TL, yazi=GRİ)
    h(ws, 30, 4, "Faiz")
    h(ws, 30, 5, 3_000_000, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "Borç Servisi Kırılımı"
    g4.add_data(Reference(ws, min_col=5, min_row=28, max_row=30), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=4, min_row=29, max_row=30))
    g4.height, g4.width = 6, 10
    ws.add_chart(g4, "H30")

    genislik(ws, {"A": 22, "B": 36, "C": 16, "D": 55, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Dosya sağlık kontrolleri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    kontroller_list = [
        (5, "Giriş dolu mu?",
         '=IFERROR(IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"BOŞ","DOLU"),"BOŞ")'),
        (6, "Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(pfd_yilTaban),N(hesapYili)<=N(pfd_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        (7, "Borç servisi > 0?",
         '=IFERROR(IF(MOTOR!G23>0,"TEMİZ","HATALI"),"HATALI")'),
        (8, "Kredi bakiyesi > 0?",
         '=IFERROR(IF(MOTOR!G16>0,"TEMİZ","HATALI"),"HATALI")'),
        (9, "DSCR negatif mi?",
         '=IFERROR(IF(MOTOR!G24<0,"NEGATİF","NORMAL"),"NORMAL")'),
        (10, "Vaka durumu",
         '=IFERROR(pfd_vakaDurum,"KIRIK")'),
        (11, "DSCR asgari altı?",
         '=IF(IFERROR(MOTOR!G24,0)<MOTOR!G8+N(GIRDI!B9)*pfd_sifirCarpan,"AŞIYOR","NORMAL")'),
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
    h(ws, 14, 2, "=IFERROR(pfd_girisBeklenen+COUNTA(GIRDI!B6:B17)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "DSCR", kalin=True)
    h(ws, 15, 2, "=IFERROR(pfd_dscr,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "LLCR", kalin=True)
    h(ws, 16, 2, "=IFERROR(pfd_llcr,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "NPV", kalin=True)
    h(ws, 17, 2, "=IFERROR(MOTOR!G25,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Borç servisi", kalin=True)
    h(ws, 18, 2, "=IFERROR(MOTOR!G23,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Karar", kalin=True)
    h(ws, 19, 2, "=IFERROR(pfd_kararMetni,\"-\")", yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 36, "B": 28})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "548235", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek senaryo verileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:L3")
    sutunlar = [
        "Senaryo", "CFADS", "Anapara", "Faiz", "Kredi bakiyesi",
        "Faiz oranı", "Vade", "LLCR dönem", "WACC", "Yıl", "Not", "DSCR Ornek",
    ]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "DSCR Ornek": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(IF((tblOrnek[[#This Row],[Anapara]]+tblOrnek[[#This Row],[Faiz]])=0,0,"
            "tblOrnek[[#This Row],[CFADS]]/"
            "(tblOrnek[[#This Row],[Anapara]]+tblOrnek[[#This Row],[Faiz]])),0))"
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:L1004", sutunlar, formuller)
    rows = [
        ("Demo ana", 18_000_000, 9_000_000, 3_000_000, 40_000_000, 0.12, 7, 5, 0.10, 2026, "UYGUN"),
        ("Yüksek DSCR", 24_000_000, 8_000_000, 2_000_000, 35_000_000, 0.11, 8, 6, 0.09, 2026, "UYGUN"),
        ("Düşük DSCR", 11_000_000, 9_000_000, 3_000_000, 45_000_000, 0.13, 6, 4, 0.11, 2026, "İHLAL"),
        ("2025 geçiş", 18_000_000, 9_000_000, 3_000_000, 40_000_000, 0.12, 7, 5, 0.10, 2025, "UYGUN"),
    ]
    for i, row in enumerate(rows, 6):
        for k, v in enumerate(row, 1):
            sayi = None
            if k in (2, 3, 4, 5):
                sayi = TL
            elif k in (6, 9):
                sayi = YÜZDE
            elif k in (7, 8, 10):
                sayi = CATI
            h(ws, i, k, v, sayi=sayi, zemin=GIRIS_SARI if k <= 11 else None, yazi="333333")
            if k <= 11:
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", sutunlar[k - 1], "Örnek satır; GIRDI'ye kopyalanabilir")
    # grafik için sayı serisi (statik — D09b)
    h(ws, 22, 1, "OrnekDscrDemo", kalin=True, yazi=GRİ)
    for i, v in enumerate([1.50, 2.40, 0.92, 1.50], 23):
        h(ws, i, 1, rows[i - 23][0])
        h(ws, i, 2, v, sayi=YÜZDE, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Örnek Senaryo DSCR"
    g.add_data(Reference(ws, min_col=2, min_row=22, max_row=26), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=23, max_row=26))
    g.height, g.width = 7, 12
    ws.add_chart(g, "D22")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 13)})
    genislik(ws, {"A": 22, "K": 12, "L": 12})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", "7030A0", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Doğrulama listeleri", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Yıl", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, y in enumerate([2025, 2026, 2027], 6):
        h(ws, i, 1, y, sayi=CATI)
    h(ws, 5, 3, "Yorum", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "A")
    h(ws, 7, 3, "B")
    h(ws, 5, 5, "EvetHayır", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 5, "Evet")
    h(ws, 7, 5, "Hayır")
    genislik(ws, {"A": 24, "C": 12, "E": 12})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "833C0C", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Parametreler (tek kaynak)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = ["anahtar", "deger", "birim", "aciklama", "kaynak", "yururluk_tarihi", "dogrulama_tarihi", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    params = [
        ("hesapYili", DEMO["hesapYili"], "yıl", "Aktif hesap yılı", "Kullanıcı / GIRDI", "01.01.2025", "11.08.2026"),
        ("raporTarihi", RAPOR_TARIH, "tarih", "Rapor tarihi (sabit)", "Üretim", "01.01.2025", "11.08.2026"),
        ("pfd_tolerans", 0.02, "oran", "Vaka tutarlılık toleransı", "Uygulama notu", "01.01.2025", "11.08.2026"),
        ("pfd_azamiEsik", 100_000_000_000, "TL", "Uç değer / azami giriş eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("pfd_kacirilanEsik", 1.25, "oran", "PANO DSCR uyarı eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("pfd_senaryoIyi", 1.10, "çarpan", "İyimser CFADS çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("pfd_senaryoKotu", 0.85, "çarpan", "Kötümser CFADS çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("pfd_esikArtis", 0.05, "oran", "M-SEN eşik + delta", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("pfd_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("pfd_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("pfd_tornadoCfads", 1.05, "çarpan", "Tornado CFADS şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("pfd_tornadoServis", 1.05, "çarpan", "Tornado servis şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("pfd_tornadoWacc", 0.01, "oran", "Tornado WACC şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("pfd_yorumBKesinti", 0.05, "oran", "Yorum B CFADS −%5", "CFADS yorumu", "01.01.2025", "11.08.2026"),
        ("pfd_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "11.08.2026"),
        ("pfd_parmakIzi", "PFD-PRO-1.0.0", "metin", "Dosya parmak izi etiketi", "Üretim", "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv", "01.01.2025", "11.08.2026"),
        ("pfd_girisBeklenen", 12, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("pfd_riskDusuk", 15, "puan", "UYGUN risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("pfd_riskOrta", 55, "puan", "DİKKAT risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("pfd_riskYuksek", 85, "puan", "İHLAL risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("pfd_riskEsikOran", 0.5, "oran", "Yedek risk oranı", "İç politika", "01.01.2025", "11.08.2026"),
        ("pfd_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("pfd_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("pfd_dayanakCarpan", 10, "adet", "Yedek çarpan", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("pfd_bazCarpan", 1, "çarpan", "Baz senaryo çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("pfd_sifirCarpan", 0, "çarpan", "Nötr çarpan (sıfır)", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("pfd_siraBir", 1, "adet", "Tornado sıra 1", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("pfd_siraIki", 2, "adet", "Tornado sıra 2", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("pfd_siraUc", 3, "adet", "Tornado sıra 3", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("pfd_tahminX", 5, "adet", "Tahmin X noktası", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("vaka_cfads_1", 18_000_000, "TL", "Vaka1/2/3 CFADS", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_cfads_2", 11_000_000, "TL", "Vaka4 CFADS", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_servis_1", 12_000_000, "TL", "Vaka borç servisi", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_borc_1", 40_000_000, "TL", "Vaka kredi bakiyesi", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_donem_1", 5, "adet", "Vaka LLCR dönem", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_wacc_1", 0.10, "oran", "Vaka WACC", "Altın vaka", "01.01.2025", "11.08.2026"),
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
        h(ws, i, 8,
          f'=IF(A{i}="","",IF(B{i}="","EKSİK","TAM"))',
          yazi=DIS_REF_YEŞIL, hiza="center")

    son_satir = 5 + len(params)
    h(ws, son_satir + 2, 2, "Parametre satırları yukarıda; ad tanımları B kolonuna bağlı.", kalin=True, yazi=KOYU_LACIVERT, kaydir=True)

    dogrulama(ws, "list", "ListeYillar", "B6",
              baslik="Hesap Yılı", mesaj="Aktif yılı seçin.",
              hata_baslik="Geçersiz", hata_mesaj="2025-2027.")
    dogrulama(ws, "decimal", "0", "B9",
              baslik="İyimser çarpan", mesaj="0-2 aralığı.",
              hata_baslik="Geçersiz", hata_mesaj="0-2.",
              isaret="between", f2="2")
    dogrulama(ws, "decimal", "0", "B10",
              baslik="Kötümser çarpan", mesaj="0-2 aralığı.",
              hata_baslik="Geçersiz", hata_mesaj="0-2.",
              isaret="between", f2="2")

    genislik(ws, {"A": 28, "B": 24, "C": 10, "D": 32, "E": 22, "F": 14, "G": 14, "H": 12})
    h(ws, son_satir + 4, 2, f"Sürüm {SURUM} | Şifre koruması: {SIFRE} (formül alanları)", yazi=GRİ, boyut=9, kaydir=True)


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar kuralları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "VERİ YOK — giriş tablosu boşsa hesap yapılmaz.",
        "UYGUN — DSCR ≥ asgari ve LLCR ≥ asgari; covenant uygun.",
        "DİKKAT — DSCR veya LLCR dikkat bandında; plan gözden geçirilmeli.",
        "İHLAL — DSCR veya LLCR dikkat eşiğinin altında; ihlal riski.",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 10, 1, "Belirsizlik beyanları (M05)", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari(["cfads_yorumu"]), 11):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 13, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2,
      "CFADS brüt girilir; borç servisi = anapara + faiz. DSCR = CFADS ÷ borç servisi. "
      "LLCR = NPV(CFADS, WACC, dönem) ÷ kredi bakiyesi.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B13:J13")
    h(ws, 14, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 2,
      "Yorum B, CFADS üzerinden rezerv kesintisi uygular (daha düşük CFADS → daha düşük DSCR/LLCR). "
      "Dosya her iki yorumun sonucunu MOTOR'da üretir.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B14:J14")

    h(ws, 16, 1, "Kullanım sırası", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 1,
      "GIRDI → KURALLAR (gerekirse) → MOTOR (inceleme) → SENARYO → VAKALAR → PANO → KANIT_RAPORU → AYARLAR.",
      kaydir=True, yazi="333333")
    ws.merge_cells("A17:J17")
    h(ws, 19, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 36, "B": 70})


def kosullu_bicimlendirme(wb):
    yesil = Font(color=NORMAL, bold=True)
    kirmizi = Font(color=KRITIK, bold=True)
    amber = Font(color="B7791F", bold=True)
    y_fill = PatternFill("solid", fgColor="E2EFDA")
    k_fill = PatternFill("solid", fgColor="FDE9E9")
    a_fill = PatternFill("solid", fgColor="FFF2CC")

    ws = wb["PANO"]
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("VERİ YOK",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("İHLAL",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("UYGUN",B6))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B6))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B4", CellIsRule(operator="greaterThanOrEqual", formula=["1.25"], font=yesil))
    ws.conditional_formatting.add("C4", CellIsRule(operator="greaterThanOrEqual", formula=["1.5"], font=yesil))
    ws.conditional_formatting.add("H4", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("H4", CellIsRule(operator="between", formula=["50", "79"], font=amber))
    ws.conditional_formatting.add("H4", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))
    ws.conditional_formatting.add("I4", CellIsRule(operator="greaterThanOrEqual", formula=["70"], font=kirmizi))
    ws.conditional_formatting.add("I4", CellIsRule(operator="between", formula=["40", "69"], font=amber))
    ws.conditional_formatting.add("I4", CellIsRule(operator="lessThan", formula=["40"], font=yesil))

    ws = wb["GIRDI"]
    for col in ("B",):
        ws.conditional_formatting.add(f"{col}9:{col}12", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(f"{col}9:{col}12", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("F26:F37", FormulaRule(formula=['F26="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F26:F37", FormulaRule(formula=['F26="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F6:F9", CellIsRule(operator="notEqual", formula=["0"], font=amber))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="greaterThan", formula=["1"], font=yesil))
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="lessThan", formula=["1.1"], font=kirmizi))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("İHLAL",E6))'], font=kirmizi))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("UYGUN",E6))'], font=yesil))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",E6))'], font=amber))
    ws.conditional_formatting.add("B16:B18", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add("B5:B11", FormulaRule(formula=['OR(B5="KIRIK",B5="HATALI",B5="NEGATİF",B5="BOŞ",B5="EKSİK",B5="AŞIYOR")'],
                                                        font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B5:B11", FormulaRule(formula=['OR(B5="TEMİZ",B5="DOLU",B5="NORMAL",B5="TUTARLI")'],
                                                        font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B12", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("B12", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("G7:G53", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("F7:F53", FormulaRule(formula=['F7="PFD-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F53", FormulaRule(formula=['F7="PFD-005"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "C", "D", "E", "J"):
        ws.conditional_formatting.add(f"{harf}6:{harf}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B21", FormulaRule(formula=['ISNUMBER(SEARCH("İHLAL",B21))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B21", FormulaRule(formula=['ISNUMBER(SEARCH("UYGUN",B21))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B21", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B21))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B16", CellIsRule(operator="greaterThanOrEqual", formula=["1.25"], font=yesil))
    ws.conditional_formatting.add("B17", CellIsRule(operator="greaterThanOrEqual", formula=["1.5"], font=yesil))


def ad_tanimlari(wb):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "pfd_dscr", "MOTOR!$G$27")
    ad_ekle(wb, "pfd_llcr", "MOTOR!$G$28")
    ad_ekle(wb, "pfd_npv", "MOTOR!$G$29")
    ad_ekle(wb, "pfd_borcServisi", "MOTOR!$G$30")
    ad_ekle(wb, "pfd_kararMetni", "MOTOR!$G$34")
    ad_ekle(wb, "pfd_kararGerekce", "MOTOR!$G$35")
    ad_ekle(wb, "pfd_maddeAtifMetni", "MOTOR!$G$52")
    ad_ekle(wb, "pfd_kanitRaporu", "MOTOR!$G$53")
    ad_ekle(wb, "pfd_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "pfd_parmakIzi", "AYARLAR!$B$21")

    ad_ekle(wb, "pfd_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "pfd_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "pfd_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "pfd_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "pfd_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "pfd_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "pfd_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "pfd_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "pfd_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "pfd_vakaFark", "VAKALAR!$B$13")

    ayar_map = {
        "pfd_tolerans": 8,
        "pfd_azamiEsik": 9,
        "pfd_kacirilanEsik": 10,
        "pfd_senaryoIyi": 11,
        "pfd_senaryoKotu": 12,
        "pfd_esikArtis": 13,
        "pfd_tahminAlt": 14,
        "pfd_tahminUst": 15,
        "pfd_tornadoCfads": 16,
        "pfd_tornadoServis": 17,
        "pfd_tornadoWacc": 18,
        "pfd_yorumBKesinti": 19,
        "pfd_yuzdelikOran": 20,
        "pfd_girisBeklenen": 23,
        "pfd_riskDusuk": 24,
        "pfd_riskOrta": 25,
        "pfd_riskYuksek": 26,
        "pfd_riskEsikOran": 27,
        "pfd_yilTaban": 28,
        "pfd_yuvarMax": 29,
        "pfd_dayanakCarpan": 30,
        "pfd_bazCarpan": 31,
        "pfd_sifirCarpan": 32,
        "pfd_siraBir": 33,
        "pfd_siraIki": 34,
        "pfd_siraUc": 35,
        "pfd_tahminX": 36,
        "vaka_cfads_1": 37,
        "vaka_cfads_2": 38,
        "vaka_servis_1": 39,
        "vaka_borc_1": 40,
        "vaka_donem_1": 41,
        "vaka_wacc_1": 42,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "pfd_modulT1": (11, "C"), "pfd_modulT1Yorum": (11, "D"),
        "pfd_modulT2": (12, "C"), "pfd_modulT2Yorum": (12, "D"),
        "pfd_modulO1": (13, "C"), "pfd_modulO1Yorum": (13, "D"),
        "pfd_modulO6": (16, "C"), "pfd_modulO6Yorum": (16, "D"),
        "pfd_modulO8": (17, "C"), "pfd_modulO8Yorum": (17, "D"),
        "pfd_modulI2": (20, "C"), "pfd_modulI2Yorum": (20, "D"),
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
    varsayilan = os.path.join(kok, "ProjeFinansmaniDscrLlcrPaketi.xlsx")
    hedef = cikti_yolu or varsayilan
    os.makedirs(os.path.dirname(os.path.abspath(hedef)) or ".", exist_ok=True)
    wb.save(hedef)

    hsh = hashlib.sha256()
    with open(hedef, "rb") as f:
        for b in iter(lambda: f.read(65536), b""):
            hsh.update(b)
    dig = hsh.hexdigest()

    wb2 = __import__("openpyxl").load_workbook(hedef)
    wb2["AYARLAR"]["B21"] = dig[:16]
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
    cikti = os.path.join(repo, "cikti", "ProjeFinansmaniDscrLlcrPaketi.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
