#!/usr/bin/env python3
"""
Fesih Maliyeti Simülatörü — üretim betiği (A1 / manda v6).
Kıdem + ihbar + işe iade beklenen maliyet; karar FESİH UYGUN / PAHALI / RİSKLİ / VERİ YOK.
Karar: FESİH UYGUN / PAHALI / RİSKLİ / VERİ YOK (eşikler KURALLAR'dan).
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

URUN_AD = "Fesih Maliyeti Simülatörü"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 11)

# Demo 2026: brüt 50k × 5 yıl = 250k kıdem; ihbar 8 hafta ≈ 93.333; beklenen ≈ 389.333 → FESİH UYGUN
DEMO = {
    "sirket_adi": "Örnek İşveren A.Ş.",
    "personel": "Örnek Çalışan",
    "hesapYili": 2026,
    "brut_ucret": 50_000,
    "kidem_yili": 5,
    "ihbar_hafta": 8,
    "ise_iade_olasilik": 0.20,
    "dava_masrafi": 30_000,
    "fesih_turu": "Haksız",
    "yorum_modu": "A",
}


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: fesih maliyeti kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: kıdem, ihbar ve karar bozulur | Kaynak: İş Kanunu / FMS kuralları",
    )


def kural_cek(kural_id: str) -> str:
    """Parametreleri tblKurallar'dan çeker; yıl seçimi fms_yilTaban ile (G07)."""
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-fms_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(fms_yuvarMax,IFERROR(N(G12),0)))"


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
      "Brüt ücret, kıdem yılı, ihbar süresi ve işe iade olasılığını girin; "
      "FESİH UYGUN / PAHALI / RİSKLİ kararını üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Çok yıllı kural tablosu (2025/2026/2027) ile kıdem/ihbar/işe iade motoru",
        "Kıdem, ihbar, doğrudan maliyet ve beklenen işe iade ek maliyetinin yan yana hesabı",
        "Baz / iyimser / kötümser + eşik değişimi (M-SEN) senaryoları",
        "Altın vakalar ve KANIT_RAPORU (A4 imza)",
        "Belirsizlik beyanı: işe iade yorumu A (standart) / B (tavan çarpanlı)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "İK / özlük — fesih maliyetini tek dosyada üretmek",
        "Mali işler / CFO — FESİH UYGUN / PAHALI / RİSKLİ kararını bütçeye yazmak",
        "SMMM / denetçi — imzalı kanıt raporunu dosyalamak",
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
        ("Adım 1 — Girdiler", "GIRDI sayfasında yıl, brüt ücret, kıdem, ihbar ve işe iade olasılığını sarı hücrelere girin."),
        ("Adım 2 — Karar", "PANO'da kıdem, ihbar, beklenen maliyet ve FESİH UYGUN / PAHALI / RİSKLİ kararını görün."),
        ("Adım 3 — Kanıt", "KANIT_RAPORU'nu yazdırın; VAKALAR ve SENARYO ile doğrulayın."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, metin, kaydir=True, yazi="333333")
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=12)
    h(ws, 9, 1, "Formüller", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 1,
      "Kıdem = brüt × yıl × (katsayı/30) | İhbar = brüt/30 × 7 × hafta | "
      "İşe iade ek = olasılık × (brüt × ay + dava)",
      yazi="333333", kaydir=True)
    ws.merge_cells("A10:L10")
    h(ws, 12, 1, "Sarı = manuel giriş | Yeşil = formül çıktısı | Şifre 1234 formül alanlarını korur.",
      yazi=GRİ, kaydir=True)
    genislik(ws, {"A": 28, "B": 70})


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Fesih maliyeti girdileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Eşikler KURALLAR'dan; tutarlar formülle üretilir.",
      yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Şirket / İşveren Adı", DEMO["sirket_adi"], None, "sirket_adi",
         "İşveren unvanını yazın."),
        (7, "Personel / Çalışan", DEMO["personel"], None, "personel",
         "Çalışan adını veya kodunu yazın."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili",
         "2025, 2026 veya 2027 seçin."),
        (9, "Brüt Ücret [₺]", DEMO["brut_ucret"], TL, "brut_ucret",
         "Aylık brüt ücreti TL girin."),
        (10, "Kıdem Yılı", DEMO["kidem_yili"], CATI, "kidem_yili",
         "Tamamlanan kıdem yılını girin."),
        (11, "İhbar Süresi [hafta]", DEMO["ihbar_hafta"], CATI, "ihbar_hafta",
         "İhbar süresini hafta cinsinden girin."),
        (12, "İşe İade Olasılığı [%]", DEMO["ise_iade_olasilik"], YÜZDE, "ise_iade_olasilik",
         "İşe iade davası başarı olasılığını girin."),
        (13, "Dava Masrafı [₺]", DEMO["dava_masrafi"], TL, "dava_masrafi",
         "Tahmini dava / vekalet masrafını TL girin."),
        (14, "Fesih Türü", DEMO["fesih_turu"], None, "fesih_turu",
         "Haklı, Haksız veya Anlaşmalı seçin."),
        (15, "Yorum Modu (A/B)", DEMO["yorum_modu"], None, "yorum_modu",
         "Yorum A (standart) veya B (tavan çarpanlı) seçin."),
    ]
    h(ws, 5, 1, "Alan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 2, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 3, "Birim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    for satir, etiket, deger, sayi, _ad, mesaj in alanlar:
        if sayi == TL:
            birim = "₺"
        elif sayi == YÜZDE:
            birim = "%"
        elif _ad in ("hesapYili", "kidem_yili"):
            birim = "yıl"
        elif _ad == "ihbar_hafta":
            birim = "hafta"
        else:
            birim = "metin"
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim, yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 17, 1, "Kıdem gün katsayısı (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 2, kural_cek("FMS-001"), sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 1, "Pahalı eşik ay (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 2, kural_cek("FMS-003"), sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Beklenen maliyet (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2, "=IFERROR(fms_beklenenToplam,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B15)=0,fms_girisBeklenen=0),0,ROUND(COUNTA(B6:B15)/fms_girisBeklenen*100,0)),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)

    h(ws, 22, 1, "Giriş tablosu (otomasyon / boş-mod)", kalin=True, yazi=KOYU_LACIVERT)
    sutunlar = ["Alan Anahtarı", "Değer", "Birim", "Zorunlu", "Kaynak", "Kontrol"]
    baslik_satiri(ws, 23, sutunlar)
    formuller = {
        "Kontrol": (
            '=IF(tblGirdi[[#This Row],[Alan Anahtarı]]="","",'
            'IF(tblGirdi[[#This Row],[Değer]]="","EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblGirdi", "A23:F1024", sutunlar, formuller)
    ornek_satir = [
        ("sirket_adi", DEMO["sirket_adi"], "metin", "Evet", "GIRDI"),
        ("personel", DEMO["personel"], "metin", "Evet", "GIRDI"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Evet", "GIRDI"),
        ("brut_ucret", DEMO["brut_ucret"], "TL", "Evet", "GIRDI"),
        ("kidem_yili", DEMO["kidem_yili"], "yıl", "Evet", "GIRDI"),
        ("ihbar_hafta", DEMO["ihbar_hafta"], "hafta", "Evet", "GIRDI"),
        ("ise_iade_olasilik", DEMO["ise_iade_olasilik"], "oran", "Evet", "GIRDI"),
        ("dava_masrafi", DEMO["dava_masrafi"], "TL", "Evet", "GIRDI"),
        ("fesih_turu", DEMO["fesih_turu"], "metin", "Evet", "GIRDI"),
        ("yorum_modu", DEMO["yorum_modu"], "A/B", "Evet", "GIRDI"),
    ]
    tl_alanlar = {"brut_ucret", "dava_masrafi"}
    oran_alanlar = {"ise_iade_olasilik"}
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 24):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in tl_alanlar:
            ws.cell(i, 2).number_format = TL
        elif a in oran_alanlar:
            ws.cell(i, 2).number_format = YÜZDE
        elif a in ("hesapYili", "kidem_yili", "ihbar_hafta"):
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
    dogrulama(ws, "list", "ListeFesihTur", "B14",
              baslik="Fesih Türü", mesaj="Haklı, Haksız veya Anlaşmalı seçin.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.", bos=False)
    dogrulama(ws, "list", "ListeYorumAB", "B15",
              baslik="Yorum", mesaj="Yorum modunda A veya B seçin.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B seçin.", bos=False)
    dogrulama(ws, "list", "ListeYillar", "B26",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeFesihTur", "B32",
              baslik="Fesih Türü", mesaj="Tablo satırında tür seçin.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeYorumAB", "B33",
              baslik="Yorum", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.")
    for aralik, baslik, mesaj in [
        ("B9", "Brüt Ücret", "0 ile 1e12 arasında TL."),
        ("B13", "Dava Masrafı", "0 ile 1e12 arasında TL."),
        ("B27", "Brüt Ücret", "Tablo: TL tutar."),
        ("B31", "Dava Masrafı", "Tablo: TL tutar."),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj=mesaj,
                  hata_baslik="Geçersiz tutar", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000000")
    dogrulama(ws, "decimal", "0", "B12",
              baslik="İşe İade Olasılığı", mesaj="0 ile 1 arasında oran.",
              hata_baslik="Geçersiz", hata_mesaj="0-1 arası.",
              isaret="between", f2="1")
    dogrulama(ws, "decimal", "0", "B30",
              baslik="İşe İade Olasılığı", mesaj="0 ile 1 arasında oran.",
              hata_baslik="Geçersiz", hata_mesaj="0-1 arası.",
              isaret="between", f2="1")
    for aralik, baslik, f2 in [
        ("B10", "Kıdem Yılı", "50"),
        ("B11", "İhbar Hafta", "52"),
        ("B28", "Kıdem Yılı", "50"),
        ("B29", "İhbar Hafta", "52"),
    ]:
        dogrulama(ws, "whole", "0", aralik,
                  baslik=baslik, mesaj=f"0-{f2} arası tam sayı.",
                  hata_baslik="Geçersiz", hata_mesaj=f"0-{f2}.",
                  isaret="between", f2=f2)

    giris_hucreleri(ws, 6, 15, [2])
    giris_hucreleri(ws, 24, 33, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 40, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1026, "Sarı alanlar giriş; eşik ve maliyet formülleri kilitlidir.")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Fesih maliyeti kural tablosu (yıl yan yana)", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Parametreler MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["FMS-001", "İş Kanunu m.14", "kidem_gun", "kıdem gün/yıl", "kidem_gun_katsayi",
         30, 30, 30, "01.01.2025", "Kıdem gün katsayısı"],
        ["FMS-002", "İş Kanunu m.17", "ihbar_gun", "ihbar günlük çarpan", "ihbar_gunluk_carpan",
         7, 7, 7, "01.01.2025", "Haftalık ihbar günü"],
        ["FMS-003", "İç politika", "pahali_ay", "doğrudan ≥ brüt×ay", "pahali_esik_ay",
         8, 8, 8, "01.01.2025", "PAHALI eşik (ay)"],
        ["FMS-004", "İç politika", "risk_olasilik", "olasılık ≥ eşik", "riskli_olasilik",
         0.40, 0.40, 0.40, "01.01.2025", "RİSKLİ olasılık eşiği"],
        ["FMS-005", "İş Kanunu işe iade", "iade_ay", "işe iade ay + yuvarlama", "ise_iade_ay",
         4, 4, 4, "01.01.2025", "İşe iade tazminat ayı"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = None
            if k in (6, 7, 8):
                sayi = YÜZDE if s[0] == "FMS-004" else CATI
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-fms_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Kıdem gün (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("FMS-001"), sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Pahalı eşik ay (aktif yıl)", yazi="333333")
    h(ws, 16, 2, kural_cek("FMS-003"), sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"FMS-001=",TEXT(B15,"0"),"FMS-003=",TEXT(B16,"0"),'
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
    h(ws, 4, 1, "Parametreler tblKurallar'dan çekilir; sabit eşik yoktur.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")

    basliklar = [*MOTOR_BASLIK, "deger"]
    baslik_satiri(ws, 6, basliklar)

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif kural yılı", "FMS-005")  # G7
    ekle("Kıdem gün katsayısı", km("FMS-001"), "adet", "Gün/yıl katsayısı", "FMS-001")  # G8
    ekle("İhbar günlük çarpan", km("FMS-002"), "adet", "Haftalık gün", "FMS-002")  # G9
    ekle("Pahalı eşik ay", km("FMS-003"), "adet", "PAHALI brüt×ay", "FMS-003")  # G10
    ekle("Riskli olasılık eşiği", km("FMS-004"), "oran", "RİSKLİ olasılık", "FMS-004")  # G11
    ekle("İşe iade ayı", km("FMS-005"), "adet", "İade tazminat ayı", "FMS-005")  # G12
    ekle("Brüt ücret", "=GIRDI!B9", "TL", "Aylık brüt", "FMS-001")  # G13
    ekle("Kıdem yılı", "=GIRDI!B10", "adet", "Kıdem yılı", "FMS-001")  # G14
    ekle("İhbar hafta", "=GIRDI!B11", "adet", "İhbar süresi", "FMS-002")  # G15
    ekle("İşe iade olasılığı", "=GIRDI!B12", "oran", "Dava olasılığı", "FMS-004")  # G16
    ekle("Dava masrafı", "=GIRDI!B13", "TL", "Dava/vekalet", "FMS-004")  # G17
    ekle("Fesih türü kodu",
         '=IF(GIRDI!B14="Haklı",1,IF(GIRDI!B14="Anlaşmalı",2,3))',
         "adet", "1=Haklı 2=Anlaşmalı 3=Haksız", "FMS-005")  # G18
    ekle("Yorum çarpanı",
         '=IF(GIRDI!B15="B",fms_yorumBCarpan,1)',
         "çarpan", "Yorum A/B kıdem çarpanı", "FMS-005")  # G19
    ekle("Haklı kıdem bayrağı",
         '=IF(G18=1,0,1)',
         "bayrak", "Haklıda kıdem=0", "FMS-001")  # G20
    ekle("Haklı ihbar bayrağı",
         '=IF(G18=1,0,1)',
         "bayrak", "Haklıda ihbar=0", "FMS-002")  # G21
    ekle("Anlaşmalı iade çarpanı",
         '=IF(G18=2,fms_anlasmaliIadeCarpan,1)',
         "çarpan", "Anlaşmalıda iade ↓", "FMS-004")  # G22
    ekle("Kıdem tazminatı",
         f"=IFERROR(ROUND(G13*G14*(G8/30)*G19*G20,{_YV}),0)",
         "TL", "Brüt×yıl×katsayı/30", "FMS-001")  # G23
    ekle("İhbar tazminatı",
         f"=IFERROR(ROUND(IF(G13=0,0,G13/30*G9*G15)*G21,{_YV}),0)",
         "TL", "Brüt/30×gün×hafta", "FMS-002")  # G24
    ekle("Doğrudan maliyet", "=G23+G24", "TL", "Kıdem+ihbar", "FMS-003")  # G25
    ekle("İşe iade brüt ek", "=G13*G12", "TL", "Brüt×iade ayı", "FMS-005")  # G26
    ekle("İşe iade beklenen ek",
         f"=IFERROR(ROUND(G16*(G26+G17)*G22,{_YV}),0)",
         "TL", "Olasılık×(ek+dava)", "FMS-004")  # G27
    ekle("Beklenen toplam", "=G25+G27", "TL", "Doğrudan+beklenen ek", "FMS-003")  # G28
    ekle("Kıdem (pano)", "=G23", "TL", "Kıdem ayna", "FMS-001")  # G29
    ekle("İhbar (pano)", "=G24", "TL", "İhbar ayna", "FMS-002")  # G30
    ekle("Doğrudan (pano)", "=G25", "TL", "Doğrudan ayna", "FMS-003")  # G31
    ekle("Beklenen (pano)", "=G28", "TL", "Beklenen ayna", "FMS-003")  # G32
    ekle("Pahalı eşik TL", "=G13*G10", "TL", "Brüt×pahalı ay", "FMS-003")  # G33
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(OR(G13<=0,G14<0,G15<0,G16<0,G16>1),"RISKLI",'
         'IF(OR(COUNTIF(tblOrnek[Brut],">="&fms_azamiEsik)>0,G16>=G11,G28>G13*12),"RISKLI",'
         'IF(G25>=G33,"PAHALI","UYGUN"))))',
         "kod", "Karar kodu", "FMS-005")  # G34
    ekle("Karar metni",
         '=IF(G34="VERI_YOK","VERİ YOK",'
         'IF(G34="UYGUN","FESİH UYGUN",'
         'IF(G34="PAHALI","PAHALI","RİSKLİ")))',
         "metin", "Kullanıcıya karar", "FMS-005")  # G35
    ekle("Gerekçe",
         '=IF(G34="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G34="UYGUN","Doğrudan maliyet pahalı eşiğin altında ve işe iade riski düşük; fesih bütçe açısından uygun.",'
         'IF(G34="PAHALI","Doğrudan kıdem+ihbar maliyeti pahalı eşiğin üzerinde; alternatif çıkış yolu değerlendirin.",'
         '"İşe iade olasılığı veya beklenen toplam maliyet risk eşiğinin üzerinde; hukuk/İK ile gözden geçirin.")))',
         "metin", "Gerekçe cümlesi", "FMS-005")  # G36
    ekle("Kıdem / brüt", "=IF(G13=0,0,G23/G13)", "oran", "Kıdem ay eşdeğeri", "FMS-001")  # G37
    ekle("İhbar / brüt", "=IF(G13=0,0,G24/G13)", "oran", "İhbar ay eşdeğeri", "FMS-002")  # G38
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B15)=0,fms_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B15)/fms_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "FMS-005")  # G39
    ekle("Risk skoru",
         '=IF(G34="VERI_YOK",0,IF(G34="UYGUN",fms_riskDusuk,'
         'IF(G34="PAHALI",fms_riskOrta,fms_riskYuksek)))',
         "puan", "Fesih riski", "FMS-005")  # G40
    ekle("Senaryo iyimser brüt", "=G13*fms_senaryoIyi", "TL", "İyimser brüt", "FMS-001")  # G41
    ekle("Senaryo kötümser brüt", "=G13*fms_senaryoKotu", "TL", "Kötümser brüt", "FMS-001")  # G42
    ekle("İyimser doğrudan",
         f"=IFERROR(ROUND(G41*G14*(G8/30)*G19*G20+IF(G41=0,0,G41/30*G9*G15)*G21,{_YV}),0)",
         "TL", "İyimser kıdem+ihbar", "FMS-003")  # G43
    ekle("Kötümser doğrudan",
         f"=IFERROR(ROUND(G42*G14*(G8/30)*G19*G20+IF(G42=0,0,G42/30*G9*G15)*G21,{_YV}),0)",
         "TL", "Kötümser kıdem+ihbar", "FMS-003")  # G44
    ekle("M-SEN pahalı eşik+", "=G10+fms_esikArtis", "adet", "Pahalı ay + delta", "FMS-003")  # G45
    ekle("M-SEN risk olasılık+", "=G11+fms_esikArtisOran", "oran", "Risk eşik + delta", "FMS-004")  # G46
    ekle("M-SEN karar bandı",
         '=IF(AND(G25<G13*G45,G16<G46),1,0)',
         "bayrak", "Sıkı eşikte uygun mu", "FMS-003")  # G47
    ekle("Tahmin üst maliyet", "=G28*fms_tahminUst", "TL", "Tahmin üst bant", "FMS-005")  # G48
    ekle("Tahmin alt maliyet", "=G28*fms_tahminAlt", "TL", "Tahmin alt bant", "FMS-005")  # G49
    ekle("Tornado: brüt etkisi",
         f"=ABS(IFERROR(ROUND((G13*fms_tornadoBrut)*G14*(G8/30)*G19*G20"
         f"+IF(G13*fms_tornadoBrut=0,0,G13*fms_tornadoBrut/30*G9*G15)*G21,{_YV})-G25,0))",
         "TL", "Brüt ± etki", "FMS-001")  # G50
    ekle("Tornado: kıdem yılı etkisi",
         f"=ABS(IFERROR(ROUND(G13*(G14+fms_tornadoKidem)*(G8/30)*G19*G20+G24,{_YV})-G25,0))",
         "TL", "Kıdem yılı ± etki", "FMS-001")  # G51
    ekle("Tornado: olasılık etkisi",
         f"=ABS(IFERROR(ROUND((G16+fms_tornadoOlasilik)*(G26+G17)*G22,{_YV})-G27,0))",
         "TL", "Olasılık ± etki", "FMS-004")  # G52
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"FMS-001→kıdem","FMS-002→ihbar",'
         '"FMS-003→pahalı eşik","FMS-004→işe iade risk","FMS-005→iade ay")',
         "metin", "Kanıt atıfları", "FMS-001")  # G53
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Kıdem=",TEXT(G23,"₺ #,##0"),"İhbar=",TEXT(G24,"₺ #,##0"),'
         '"Doğrudan=",TEXT(G25,"₺ #,##0"),"Beklenen=",TEXT(G28,"₺ #,##0"))',
         "metin", "Rapor özeti", "FMS-005")  # G54
    ekle("İşe iade / toplam",
         "=IF(G28=0,0,G27/G28)",
         "oran", "İade yoğunluğu", "FMS-004")  # G55

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
      "Motor adımları kilitlidir; parametreler yalnızca KURALLAR'dan gelir.", yazi=GRİ, kaydir=True)
    return len(adimlar)


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve duyarlılık", kalin=True, boyut=13, yazi=KOYU_LACIVERT)

    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Brüt Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Doğrudan [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Beklenen [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(fms_senaryoIyi+N(GIRDI!B9)*fms_sifirCarpan,0)",
         "=IFERROR(MOTOR!G43,0)",
         "=IFERROR(MOTOR!G43+MOTOR!G27,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(OR(GIRDI!B12>=MOTOR!G11,D6>GIRDI!B9*12),"RİSKLİ",'
         'IF(C6>=GIRDI!B9*MOTOR!G10,"PAHALI","FESİH UYGUN")))'),
        ("Baz", "=IFERROR(fms_bazCarpan+N(GIRDI!B9)*fms_sifirCarpan,fms_bazCarpan)",
         "=IFERROR(MOTOR!G25,0)",
         "=IFERROR(fms_beklenenToplam,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(fms_kararMetni,"-"))'),
        ("Kötümser", "=IFERROR(fms_senaryoKotu+N(GIRDI!B9)*fms_sifirCarpan,0)",
         "=IFERROR(MOTOR!G44,0)",
         "=IFERROR(MOTOR!G44+MOTOR!G27,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(OR(GIRDI!B12>=MOTOR!G11,D8>GIRDI!B9*12),"RİSKLİ",'
         'IF(C8>=GIRDI!B9*MOTOR!G10,"PAHALI","FESİH UYGUN")))'),
        ("M-SEN eşik+", "=IFERROR(fms_bazCarpan+N(GIRDI!B9)*fms_sifirCarpan,fms_bazCarpan)",
         "=IFERROR(MOTOR!G25,0)",
         "=IFERROR(fms_beklenenToplam,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(MOTOR!G47=1,"FESİH UYGUN","PAHALI"))'),
    ]
    for i, (ad, carp, dog, bek, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, dog, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, bek, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant (iyimser−kötümser)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(D6-D8,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser ",TEXT(D6,"₺ #,##0")," · Baz ",TEXT(D7,"₺ #,##0"),'
      '" · Kötümser ",TEXT(D8,"₺ #,##0")," · Bant ",TEXT(B11,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:H12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Tornado brüt etkisi", yazi="333333")
    h(ws, 16, 2, "=IFERROR(MOTOR!G50,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado kıdem yılı etkisi", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G51,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado olasılık etkisi", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G52,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 3,
      '=IFERROR(IF(B16>=MAX(B16:B18),fms_siraBir,IF(B16>=MIN(B16:B18)+ABS(B16-B17),fms_siraIki,fms_siraUc)),fms_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 3,
      '=IFERROR(IF(B17>=MAX(B16:B18),fms_siraBir,IF(B17=MEDIAN(B16:B18),fms_siraIki,fms_siraUc)),fms_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 3,
      '=IFERROR(IF(B18>=MAX(B16:B18),fms_siraBir,IF(B18=MEDIAN(B16:B18),fms_siraIki,fms_siraUc)),fms_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(CONCATENATE("1)",INDEX(A16:A18,MATCH(fms_siraBir,C16:C18,0)),'
      '" 2)",INDEX(A16:A18,MATCH(fms_siraIki,C16:C18,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B16:B18),"₺ #,##0")," — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN eşik değişimi bayrağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=MOTOR!G47", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Sıkı eşik (+",TEXT(fms_esikArtis,"0")," ay / +",TEXT(fms_esikArtisOran,"0%"),'
      '") altında uygunluk bayrağı ",TEXT(B22,"0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı (beklenen maliyet)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=MOTOR!G49", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 3, "=IFERROR(MOTOR!G28,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 4, "=MOTOR!G48", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Tahmin (AVERAGE / STDEV koruması)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1, "Seri nokta", yazi=GRİ)
    for i, v in enumerate([1, 2, 3, 4], 28):
        h(ws, i, 1, v, sayi=CATI)
        h(ws, i, 2, f"=D{5+v}", sayi=TL, yazi=DIS_REF_YEŞIL)
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

    h(ws, 5, 7, "MaliyetDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([360000, 389333, 420000, 389333], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([17500, 50000, 11500], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)

    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([360000, 389333, 420000, 389333], 28):
        h(ws, i, 4, v, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo Beklenen Maliyet"
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
    g3.title = "Senaryo Maliyet Trendi"
    g3.add_data(Reference(ws, min_col=4, min_row=27, max_row=31), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=28, max_row=31))
    g3.height, g3.width = 8, 12
    ws.add_chart(g3, "F32")

    genislik(ws, {"A": 36, "B": 18, "C": 18, "D": 18, "E": 22, "G": 14})


def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Altın vakalar — kıdem / ihbar / işe iade", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:H3")
    baslik_satiri(ws, 5, VAKA_BASLIK)

    # Demo: brut=50000, yıl=5, katsayi=30 → kıdem=250000
    # ihbar: 50000/30*7*8 = 93333.333 → ROUND 93333
    # doğrudan = 343333
    # iade ek: 0.2*(50000*4+30000)=46000
    vakalar_data = [
        ("V-001", "Brüt × yıl × (katsayı/30) = kıdem",
         "brut=vaka_brut_1; yıl=vaka_yil_1; katsayi=30",
         250000,
         "=IFERROR(ROUND(vaka_brut_1*vaka_yil_1*(30/30),0),0)",
         "=IFERROR(E6-D6,0)",
         '=IFERROR(IF(ABS(F6)<=fms_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Demo ana kıdem"),
        ("V-002", "Brüt/30 × 7 × hafta = ihbar",
         "brut=vaka_brut_1; hafta=vaka_hafta_1",
         93333,
         "=IFERROR(ROUND(vaka_brut_1/30*7*vaka_hafta_1,0),0)",
         "=IFERROR(E7-D7,0)",
         '=IFERROR(IF(ABS(F7)<=fms_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Demo ana ihbar"),
        ("V-003", "Kıdem + ihbar = doğrudan",
         "kidem=250000; ihbar=93333",
         343333,
         "=IFERROR(ROUND(vaka_brut_1*vaka_yil_1+vaka_brut_1/30*7*vaka_hafta_1,0),0)",
         "=IFERROR(E8-D8,0)",
         '=IFERROR(IF(ABS(F8)<=fms_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Doğrudan maliyet"),
        ("V-004", "Olasılık × (brüt×ay + dava) = iade ek",
         "p=vaka_olasilik_1; ay=4; dava=vaka_dava_1",
         46000,
         "=IFERROR(ROUND(vaka_olasilik_1*(vaka_brut_1*4+vaka_dava_1),0),0)",
         "=IFERROR(E9-D9,0)",
         '=IFERROR(IF(ABS(F9)<=fms_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "İşe iade beklenen"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v, yazi=DIS_REF_YEŞIL if k >= 5 else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = "#,##0"

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
            'IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=fms_tolerans,"TUTARLI","KIRIK"),"KIRIK"))'
        ),
    }
    tablo_ekle(ws, "tblVakalar", "A5:H9", VAKA_BASLIK, formuller)

    h(ws, 12, 1, "Vaka durumu özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(IF(COUNTIF(G6:G9,"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "Toplam mutlak fark", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2, "=IFERROR(SUM(ABS(F6),ABS(F7),ABS(F8),ABS(F9)),0)", sayi="#,##0", yazi=DIS_REF_YEŞIL)

    genislik(ws, {"A": 10, "B": 40, "C": 40, "D": 12, "E": 14, "F": 12, "G": 12, "H": 18})
    sabitle(ws, "A6")


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", "0F2742", URUN_AD, son_kolon=30)
    baski_hazirla(ws, "A1:F34", f"{URUN_AD} · {SURUM}")
    h(ws, 3, 1, "KANIT RAPORU — Fesih Maliyeti (kıdem+ihbar+işe iade)", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "İK / mali işler dosyası", yazi=GRİ)
    h(ws, 6, 1, "Şirket / İşveren", kalin=True)
    h(ws, 6, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Personel", kalin=True)
    h(ws, 7, 2, "=GIRDI!B7", yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Hesap yılı", kalin=True)
    h(ws, 8, 2, "=hesapYili", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Rapor tarihi", kalin=True)
    h(ws, 9, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Brüt ücret [₺]", kalin=True)
    h(ws, 11, 2, "=MOTOR!G13", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Kıdem tazminatı [₺]", kalin=True)
    h(ws, 12, 2, "=MOTOR!G23", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "İhbar tazminatı [₺]", kalin=True)
    h(ws, 13, 2, "=MOTOR!G24", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "İşe iade beklenen ek [₺]", kalin=True)
    h(ws, 14, 2, "=MOTOR!G27", sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 16, 1, "Doğrudan maliyet", kalin=True, boyut=12)
    h(ws, 16, 2, "=fms_dogrudanMaliyet", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 17, 1, "Beklenen toplam", kalin=True, boyut=12)
    h(ws, 17, 2, "=fms_beklenenToplam", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 1, "Pahalı eşik [₺]", kalin=True)
    h(ws, 18, 2, "=MOTOR!G33", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Riskli olasılık eşiği", kalin=True)
    h(ws, 19, 2, "=MOTOR!G11", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)

    h(ws, 21, 1, "Karar", kalin=True, boyut=12)
    h(ws, 21, 2, "=fms_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 22, 1, "Gerekçe", kalin=True)
    h(ws, 22, 2, "=fms_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B22:H22")
    h(ws, 24, 1, "Madde / kural atıfı", kalin=True)
    h(ws, 24, 2, "=fms_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B24:H24")
    h(ws, 26, 1, "Kanıt özeti", kalin=True)
    h(ws, 26, 2, "=fms_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B26:H26")
    h(ws, 28, 1, "Parmak izi", kalin=True)
    h(ws, 28, 2, "=fms_parmakIzi", yazi=GRİ)

    h(ws, 31, 1, "Hazırlayan", kalin=True)
    h(ws, 31, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B31", "Hazırlayan", "Ad soyad imza alanı")
    h(ws, 32, 1, "Onaylayan", kalin=True)
    h(ws, 32, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B32", "Onaylayan", "İK / mali işler imza alanı")
    genislik(ws, {"A": 42, "B": 55})
    h(ws, 34, 1, "Bu çıktı karar destek amaçlıdır; bağlayıcı hukuki tavsiye değildir.",
      yazi=GRİ, boyut=9, kaydir=True)


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=40)
    baski_hazirla(ws, "A1:F34", f"{URUN_AD} · PANO")
    h(ws, 3, 1, "Karar panosu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpis = [
        (2, "Kıdem", "=IFERROR(fms_kidem,0)", TL),
        (3, "İhbar", "=IFERROR(fms_ihbar,0)", TL),
        (4, "Doğrudan", "=IFERROR(fms_dogrudanMaliyet,0)", TL),
        (5, "İade ek", "=IFERROR(MOTOR!G27,0)", TL),
        (6, "Beklenen", "=IFERROR(fms_beklenenToplam,0)", TL),
        (7, "Brüt", "=IFERROR(MOTOR!G13,0)", TL),
        (8, "Güven", "=IFERROR(MOTOR!G39,0)", CATI),
        (9, "Risk", "=IFERROR(MOTOR!G40,0)", CATI),
        (10, "İade yoğun.", "=IFERROR(MOTOR!G55,0)", YÜZDE),
        (11, "M-SEN", "=IFERROR(fms_kuralDegisimSenaryo,0)", CATI),
        (12, "Tahmin", "=IFERROR(fms_tahminAralik,0)", TL),
        (13, "Vaka", "=IFERROR(fms_vakaDurum,\"-\")", None),
    ]
    for col, ad, formul, sayi in kpis:
        h(ws, 3, col, ad, kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
        h(ws, 4, col, formul, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=12)

    h(ws, 6, 1, "Karar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=fms_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=fms_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 10, 1, "Kod", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 2, "Ad", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 3, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 4, "Yorum", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    moduller = [
        ("T1", "Kıdem / brüt ücret oranı",
         "=IFERROR(MOTOR!G37,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kıdem/brüt ",TEXT(C11,"0.00")," — yıl ",TEXT(MOTOR!G14,"0"))'),
        ("T2", "İhbar / brüt ücret oranı",
         "=IFERROR(MOTOR!G38,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"İhbar/brüt ",TEXT(C12,"0.00")," — hafta ",TEXT(MOTOR!G15,"0"))'),
        ("O1", "Senaryo maliyet sapması",
         "=IFERROR(ABS(SENARYO!D6-SENARYO!D8),0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"İyimser−kötümser bant ",TEXT(C13,"₺ #,##0"))'),
        ("O2", "Duyarlılık / tornado",
         "=IFERROR(fms_duyarlilikSira,\"-\")",
         "=IFERROR(fms_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo motoru",
         "=IFERROR(fms_senaryoKarsilastirma,0)",
         "=IFERROR(fms_yorumSenaryo,\"-\")"),
        ("O6", "İşe iade / toplam",
         "=IFERROR(MOTOR!G55,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"İade yoğunluğu ",TEXT(C16,"0.0%"))'),
        ("O8", "Veri kalite skoru",
         "=IFERROR(MOTOR!G39,0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Giriş güven skoru ",TEXT(C17,"0"),"/100")'),
        ("M_SEN", "Eşik değişimi",
         "=IFERROR(fms_kuralDegisimSenaryo,0)",
         "=IFERROR(fms_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin + aralık",
         "=IFERROR(fms_tahminAralik,0)",
         "=IFERROR(fms_yorumTahmin,\"-\")"),
        ("I2", "Beklenen maliyet P90",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!D6:D9,fms_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 beklenen ",TEXT(C20,"₺ #,##0"))'),
    ]
    for i, (kod, ad, deger, yorum) in enumerate(moduller, 11):
        h(ws, i, 1, kod, kalin=True, yazi="B08948")
        h(ws, i, 2, ad, yazi="333333")
        sayi = YÜZDE if kod in ("T1", "T2", "O6") else (CATI if kod in ("O8", "M_SEN") else (TL if kod in ("O1", "O3", "I1", "I2") else None))
        h(ws, i, 3, deger, sayi=sayi, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 23, 1, "PanoDemoMaliyet", kalin=True, yazi=GRİ)
    for i, (ad, v) in enumerate([("Kıdem", 250000), ("İhbar", 93333), ("İade ek", 46000)], 24):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, v, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Maliyet Kırılımı"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=26), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=26))
    g1.height, g1.width = 7, 12
    ws.add_chart(g1, "F23")

    h(ws, 23, 4, "Etiket")
    h(ws, 23, 5, "PanoDemoEsik", kalin=True, yazi=GRİ)
    for i, (ad, v) in enumerate([("Doğrudan", 343333), ("Pahalı eşik", 400000), ("Beklenen", 389333)], 24):
        h(ws, i, 4, ad, yazi="333333")
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)

    g2 = BarChart()
    g2.type = "col"
    g2.title = "Eşik Karşılaştırması"
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

    h(ws, 28, 4, "Kalem")
    h(ws, 28, 5, "SenaryoDemo", kalin=True, yazi=GRİ)
    h(ws, 29, 4, "İyimser")
    h(ws, 29, 5, 360000, sayi=TL, yazi=GRİ)
    h(ws, 30, 4, "Kötümser")
    h(ws, 30, 5, 420000, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "Senaryo Bant"
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
         '=IFERROR(IF(AND(N(hesapYili)>N(fms_yilTaban),N(hesapYili)<=N(fms_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        (7, "Brüt ücret > 0?",
         '=IFERROR(IF(MOTOR!G13>0,"TEMİZ","HATALI"),"HATALI")'),
        (8, "Kıdem yılı ≥ 0?",
         '=IFERROR(IF(MOTOR!G14>=0,"TEMİZ","HATALI"),"HATALI")'),
        (9, "Doğrudan negatif mi?",
         '=IFERROR(IF(MOTOR!G25<0,"NEGATİF","NORMAL"),"NORMAL")'),
        (10, "Vaka durumu",
         '=IFERROR(fms_vakaDurum,"KIRIK")'),
        (11, "Pahalı eşik aşıldı mı?",
         '=IF(IFERROR(MOTOR!G25,0)>=MOTOR!G33+N(GIRDI!B9)*fms_sifirCarpan,"AŞIYOR","NORMAL")'),
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
    h(ws, 14, 2, "=IFERROR(fms_girisBeklenen+COUNTA(GIRDI!B6:B15)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Kıdem", kalin=True)
    h(ws, 15, 2, "=IFERROR(fms_kidem,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "İhbar", kalin=True)
    h(ws, 16, 2, "=IFERROR(fms_ihbar,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Doğrudan", kalin=True)
    h(ws, 17, 2, "=IFERROR(fms_dogrudanMaliyet,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Beklenen toplam", kalin=True)
    h(ws, 18, 2, "=IFERROR(fms_beklenenToplam,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Karar", kalin=True)
    h(ws, 19, 2, "=IFERROR(fms_kararMetni,\"-\")", yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 36, "B": 28})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "548235", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek senaryo verileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:L3")
    sutunlar = [
        "Senaryo", "Brut", "KidemYil", "IhbarHafta", "IadeOlasilik",
        "Dava", "FesihTur", "Yil", "Not", "DogruOrnek", "BeklenenOrnek", "Kontrol",
    ]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "DogruOrnek": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(tblOrnek[[#This Row],[Brut]]*tblOrnek[[#This Row],[KidemYil]]"
            "+tblOrnek[[#This Row],[Brut]]/30*7*tblOrnek[[#This Row],[IhbarHafta]],0))"
        ),
        "BeklenenOrnek": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(tblOrnek[[#This Row],[DogruOrnek]]"
            "+tblOrnek[[#This Row],[IadeOlasilik]]*"
            "(tblOrnek[[#This Row],[Brut]]*4+tblOrnek[[#This Row],[Dava]]),0))"
        ),
        "Kontrol": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IF(tblOrnek[[#This Row],[Brut]]=\"\",\"EKSİK\",\"TAM\"))"
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:L1004", sutunlar, formuller)
    rows = [
        ("Demo ana", 50_000, 5, 8, 0.20, 30_000, "Haksız", 2026, "FESİH UYGUN"),
        ("Yüksek kıdem", 80_000, 12, 8, 0.15, 40_000, "Haksız", 2026, "PAHALI"),
        ("Riskli iade", 45_000, 4, 6, 0.55, 35_000, "Haksız", 2026, "RİSKLİ"),
        ("2025 geçiş", 50_000, 5, 8, 0.20, 30_000, "Haksız", 2025, "FESİH UYGUN"),
    ]
    for i, row in enumerate(rows, 6):
        for k, v in enumerate(row, 1):
            sayi = None
            if k in (2, 6):
                sayi = TL
            elif k == 5:
                sayi = YÜZDE
            elif k in (3, 4, 8):
                sayi = CATI
            h(ws, i, k, v, sayi=sayi, zemin=GIRIS_SARI if k <= 9 else None, yazi="333333")
            if k <= 9:
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", sutunlar[k - 1], "Örnek satır; GIRDI'ye kopyalanabilir")
    h(ws, 22, 1, "OrnekMaliyetDemo", kalin=True, yazi=GRİ)
    for i, v in enumerate([389333, 1125333, 314500, 389333], 23):
        h(ws, i, 1, rows[i - 23][0])
        h(ws, i, 2, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Örnek Senaryo Beklenen"
    g.add_data(Reference(ws, min_col=2, min_row=22, max_row=26), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=23, max_row=26))
    g.height, g.width = 7, 12
    ws.add_chart(g, "D22")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 13)})
    genislik(ws, {"A": 22, "G": 12, "I": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", "7030A0", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Doğrulama listeleri", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Yıl", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, y in enumerate([2025, 2026, 2027], 6):
        h(ws, i, 1, y, sayi=CATI)
    h(ws, 5, 3, "Yorum", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "A")
    h(ws, 7, 3, "B")
    h(ws, 5, 5, "FesihTür", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 5, "Haklı")
    h(ws, 7, 5, "Haksız")
    h(ws, 8, 5, "Anlaşmalı")
    h(ws, 5, 7, "EvetHayır", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 7, "Evet")
    h(ws, 7, 7, "Hayır")
    genislik(ws, {"A": 24, "C": 12, "E": 14, "G": 12})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "833C0C", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Parametreler (tek kaynak)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = ["anahtar", "deger", "birim", "aciklama", "kaynak", "yururluk_tarihi", "dogrulama_tarihi", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    params = [
        ("hesapYili", DEMO["hesapYili"], "yıl", "Aktif hesap yılı", "Kullanıcı / GIRDI", "01.01.2025", "11.08.2026"),
        ("raporTarihi", RAPOR_TARIH, "tarih", "Rapor tarihi (sabit)", "Üretim", "01.01.2025", "11.08.2026"),
        ("fms_tolerans", 1, "TL", "Vaka tutarlılık toleransı (TL)", "Uygulama notu", "01.01.2025", "11.08.2026"),
        ("fms_azamiEsik", 100_000_000_000, "TL", "Uç değer / azami giriş eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("fms_kacirilanEsik", 400000, "TL", "PANO doğrudan uyarı eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("fms_senaryoIyi", 0.95, "çarpan", "İyimser brüt çarpanı (düşük maliyet)", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("fms_senaryoKotu", 1.10, "çarpan", "Kötümser brüt çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("fms_esikArtis", 1, "adet", "M-SEN pahalı ay + delta", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("fms_esikArtisOran", 0.05, "oran", "M-SEN risk olasılık + delta", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("fms_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("fms_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("fms_tornadoBrut", 1.05, "çarpan", "Tornado brüt şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("fms_tornadoKidem", 1, "adet", "Tornado kıdem yılı şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("fms_tornadoOlasilik", 0.05, "oran", "Tornado olasılık şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("fms_yorumBCarpan", 0.95, "çarpan", "Yorum B kıdem çarpanı", "İşe iade yorumu", "01.01.2025", "11.08.2026"),
        ("fms_anlasmaliIadeCarpan", 0.25, "çarpan", "Anlaşmalı işe iade çarpanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("fms_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "11.08.2026"),
        ("fms_parmakIzi", "FMS-PRO-1.0.0", "metin", "Dosya parmak izi etiketi", "Üretim", "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv", "01.01.2025", "11.08.2026"),
        ("fms_girisBeklenen", 10, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("fms_riskDusuk", 15, "puan", "FESİH UYGUN risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("fms_riskOrta", 55, "puan", "PAHALI risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("fms_riskYuksek", 85, "puan", "RİSKLİ risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("fms_riskEsikOran", 0.5, "oran", "Yedek risk oranı", "İç politika", "01.01.2025", "11.08.2026"),
        ("fms_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("fms_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("fms_dayanakCarpan", 10, "adet", "Yedek çarpan", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("fms_bazCarpan", 1, "çarpan", "Baz senaryo çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("fms_sifirCarpan", 0, "çarpan", "Nötr çarpan (sıfır)", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("fms_siraBir", 1, "adet", "Tornado sıra 1", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("fms_siraIki", 2, "adet", "Tornado sıra 2", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("fms_siraUc", 3, "adet", "Tornado sıra 3", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("fms_tahminX", 5, "adet", "Tahmin X noktası", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("vaka_brut_1", 50_000, "TL", "Vaka brüt ücret", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_yil_1", 5, "adet", "Vaka kıdem yılı", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_hafta_1", 8, "adet", "Vaka ihbar hafta", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_olasilik_1", 0.20, "oran", "Vaka işe iade olasılık", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_dava_1", 30_000, "TL", "Vaka dava masrafı", "Altın vaka", "01.01.2025", "11.08.2026"),
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
    dogrulama(ws, "decimal", "0", "B11",
              baslik="İyimser çarpan", mesaj="0-2 aralığı.",
              hata_baslik="Geçersiz", hata_mesaj="0-2.",
              isaret="between", f2="2")
    dogrulama(ws, "decimal", "0", "B12",
              baslik="Kötümser çarpan", mesaj="0-2 aralığı.",
              hata_baslik="Geçersiz", hata_mesaj="0-2.",
              isaret="between", f2="2")

    genislik(ws, {"A": 28, "B": 24, "C": 10, "D": 36, "E": 22, "F": 14, "G": 14, "H": 12})
    h(ws, son_satir + 4, 2, f"Sürüm {SURUM} | Şifre koruması: {SIFRE} (formül alanları)", yazi=GRİ, boyut=9, kaydir=True)


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar kuralları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "VERİ YOK — giriş tablosu boşsa hesap yapılmaz.",
        "FESİH UYGUN — doğrudan maliyet pahalı eşiğin altında ve işe iade riski düşük.",
        "PAHALI — kıdem+ihbar doğrudan maliyeti pahalı eşiğin üzerinde.",
        "RİSKLİ — işe iade olasılığı veya beklenen toplam maliyet risk eşiğinin üzerinde.",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 10, 1, "Belirsizlik beyanları (M05)", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari(["ise_iade_yorumu"]), 11):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 13, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2,
      "Standart hesap: kıdem = brüt × yıl × (katsayı/30); ihbar = brüt/30 × günlük çarpan × hafta. "
      "İşe iade beklenen ek = olasılık × (brüt × iade ayı + dava). Haklı fesihte kıdem/ihbar sıfırlanır.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B13:J13")
    h(ws, 14, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 2,
      "Yorum B, kıdem tutarına tavan çarpanı uygular (daha düşük kıdem → daha düşük doğrudan maliyet). "
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
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("RİSKLİ",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("FESİH UYGUN",B6))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("PAHALI",B6))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("H4", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("H4", CellIsRule(operator="between", formula=["50", "79"], font=amber))
    ws.conditional_formatting.add("H4", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))
    ws.conditional_formatting.add("I4", CellIsRule(operator="greaterThanOrEqual", formula=["70"], font=kirmizi))
    ws.conditional_formatting.add("I4", CellIsRule(operator="between", formula=["40", "69"], font=amber))
    ws.conditional_formatting.add("I4", CellIsRule(operator="lessThan", formula=["40"], font=yesil))

    ws = wb["GIRDI"]
    for col in ("B",):
        ws.conditional_formatting.add(f"{col}9:{col}9", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(f"{col}9:{col}9", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("F24:F33", FormulaRule(formula=['F24="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F24:F33", FormulaRule(formula=['F24="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F6:F9", CellIsRule(operator="notEqual", formula=["0"], font=amber))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("RİSKLİ",E6))'], font=kirmizi))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("FESİH UYGUN",E6))'], font=yesil))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("PAHALI",E6))'], font=amber))
    ws.conditional_formatting.add("B16:B18", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add("B5:B11", FormulaRule(formula=['OR(B5="KIRIK",B5="HATALI",B5="NEGATİF",B5="BOŞ",B5="EKSİK",B5="AŞIYOR")'],
                                                        font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B5:B11", FormulaRule(formula=['OR(B5="TEMİZ",B5="DOLU",B5="NORMAL",B5="TUTARLI")'],
                                                        font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B12", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("B12", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("G7:G55", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("F7:F55", FormulaRule(formula=['F7="FMS-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F55", FormulaRule(formula=['F7="FMS-005"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "C", "D", "F", "H"):
        ws.conditional_formatting.add(f"{harf}6:{harf}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B21", FormulaRule(formula=['ISNUMBER(SEARCH("RİSKLİ",B21))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B21", FormulaRule(formula=['ISNUMBER(SEARCH("FESİH UYGUN",B21))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B21", FormulaRule(formula=['ISNUMBER(SEARCH("PAHALI",B21))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B21", FormulaRule(formula=['ISNUMBER(SEARCH("VERİ YOK",B21))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B16", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("B17", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))

    ws = wb["KURALLAR"]
    ws.conditional_formatting.add("F7:H11", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("K7:K11", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))

    ws = wb["AYARLAR"]
    ws.conditional_formatting.add("H6:H50", FormulaRule(formula=['H6="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("H6:H50", FormulaRule(formula=['H6="TAM"'], font=yesil, fill=y_fill))


def ad_tanimlari(wb):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "fms_kidem", "MOTOR!$G$29")
    ad_ekle(wb, "fms_ihbar", "MOTOR!$G$30")
    ad_ekle(wb, "fms_dogrudanMaliyet", "MOTOR!$G$31")
    ad_ekle(wb, "fms_beklenenToplam", "MOTOR!$G$32")
    ad_ekle(wb, "fms_kararMetni", "MOTOR!$G$35")
    ad_ekle(wb, "fms_kararGerekce", "MOTOR!$G$36")
    ad_ekle(wb, "fms_maddeAtifMetni", "MOTOR!$G$53")
    ad_ekle(wb, "fms_kanitRaporu", "MOTOR!$G$54")
    ad_ekle(wb, "fms_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "fms_parmakIzi", "AYARLAR!$B$23")

    ad_ekle(wb, "fms_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "fms_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "fms_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "fms_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "fms_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "fms_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "fms_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "fms_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "fms_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "fms_vakaFark", "VAKALAR!$B$13")

    ayar_map = {
        "fms_tolerans": 8,
        "fms_azamiEsik": 9,
        "fms_kacirilanEsik": 10,
        "fms_senaryoIyi": 11,
        "fms_senaryoKotu": 12,
        "fms_esikArtis": 13,
        "fms_esikArtisOran": 14,
        "fms_tahminAlt": 15,
        "fms_tahminUst": 16,
        "fms_tornadoBrut": 17,
        "fms_tornadoKidem": 18,
        "fms_tornadoOlasilik": 19,
        "fms_yorumBCarpan": 20,
        "fms_anlasmaliIadeCarpan": 21,
        "fms_yuzdelikOran": 22,
        "fms_girisBeklenen": 25,
        "fms_riskDusuk": 26,
        "fms_riskOrta": 27,
        "fms_riskYuksek": 28,
        "fms_riskEsikOran": 29,
        "fms_yilTaban": 30,
        "fms_yuvarMax": 31,
        "fms_dayanakCarpan": 32,
        "fms_bazCarpan": 33,
        "fms_sifirCarpan": 34,
        "fms_siraBir": 35,
        "fms_siraIki": 36,
        "fms_siraUc": 37,
        "fms_tahminX": 38,
        "vaka_brut_1": 39,
        "vaka_yil_1": 40,
        "vaka_hafta_1": 41,
        "vaka_olasilik_1": 42,
        "vaka_dava_1": 43,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "fms_modulT1": (11, "C"), "fms_modulT1Yorum": (11, "D"),
        "fms_modulT2": (12, "C"), "fms_modulT2Yorum": (12, "D"),
        "fms_modulO1": (13, "C"), "fms_modulO1Yorum": (13, "D"),
        "fms_modulO6": (16, "C"), "fms_modulO6Yorum": (16, "D"),
        "fms_modulO8": (17, "C"), "fms_modulO8Yorum": (17, "D"),
        "fms_modulI2": (20, "C"), "fms_modulI2Yorum": (20, "D"),
    }
    for ad, (satir, kol) in modul_map.items():
        ad_ekle(wb, ad, f"PANO!${kol}${satir}")

    ad_ekle(wb, "ListeYillar", "LISTELER!$A$6:$A$8")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$C$6:$C$7")
    ad_ekle(wb, "ListeFesihTur", "LISTELER!$E$6:$E$8")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$G$6:$G$7")


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
    varsayilan = os.path.join(kok, "FesihMaliyetiSimulatoru.xlsx")
    hedef = cikti_yolu or varsayilan
    os.makedirs(os.path.dirname(os.path.abspath(hedef)) or ".", exist_ok=True)
    wb.save(hedef)

    hsh = hashlib.sha256()
    with open(hedef, "rb") as f:
        for b in iter(lambda: f.read(65536), b""):
            hsh.update(b)
    dig = hsh.hexdigest()

    wb2 = __import__("openpyxl").load_workbook(hedef)
    wb2["AYARLAR"]["B23"] = dig[:16]
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
    cikti = os.path.join(repo, "cikti", "FesihMaliyetiSimulatoru.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
