#!/usr/bin/env python3
"""Dava Masraf & Harç Hesaplama Motoru — A1 mevzuat motoru (manda v6).

Harç, masraf ve AAÜT kalemlerini yıllı kural tablosuyla hesaplar.
Domain: dava masraf/harç — rename YASAK (harç, masraf, AAÜT korunur).
"""

from __future__ import annotations

import hashlib
import os
import shutil
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Font, PatternFill, Protection
from openpyxl.utils import get_column_letter

KOK = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
if KOK not in sys.path:
    sys.path.insert(0, KOK)

from excel_uretim.ortak import (  # noqa: E402
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
from ortak.mevzuat_motoru import (  # noqa: E402
    KURAL_BASLIK,
    MOTOR_BASLIK,
    VAKA_BASLIK,
)

URUN_AD = "Dava Masraf & Harç Hesaplama Motoru"
SURUM = "1.0.0"
RENK = "1B2838"
RAPOR_TARIH = date(2026, 8, 11)

# Demo: 250.000 TL asliye hukuk, 2 tebligat, bilirkişi+keşif, AAÜT, bütçe içi → UYGUN
DEMO = {
    "dosya_numarasi": "2026/DMH-001",
    "muvekkil": "Örnek Müvekkil A.Ş.",
    "hesapYili": 2026,
    "mahkeme_turu": "AsliyeHukuk",
    "dava_degeri": 250000.0,
    "islah_var": "Hayır",
    "islah_tutari": 0.0,
    "tebligat_adet": 2,
    "bilirkisi": "Evet",
    "kesif": "Evet",
    "vekalet_aaut": "Evet",
    "butce_ust_sinir": 45000.0,
    "yorum_ab": "A",
}


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: Harç/masraf/AAÜT hesabını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Toplam yük ve karar bozulur | "
        f"Kaynak: 492 Harçlar Kanunu / TBB AAÜT",
    )


def kural_cek(kural_id: str) -> str:
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(dmh_bir,MIN(3,IF(OR(hesapYili="",hesapYili=0),dmh_bir,'
        f'hesapYili-dmh_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


def _guvenli(formul: str, birim: str) -> str:
    if not formul.startswith("="):
        formul = "=" + formul
    inner = formul[1:]
    if inner.upper().startswith("IFERROR("):
        return formul
    yedek = '""' if birim in ("metin", "kod") else "0"
    return f"=IFERROR({inner},{yedek})"


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, URUN_AD, kalin=True, boyut=18, yazi="FFFFFF", zemin=KOYU_LACIVERT)
    ws.merge_cells("A3:N3")
    h(ws, 4, 1,
      "Dava değeri, ıslah, tebligat ve masraf girdilerini yıllı harç/AAÜT kural tablosuyla "
      "hesaplayıp UYGUN / DİKKAT / UYGUN DEĞİL kararı üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:N4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Çok yıllı harç ve AAÜT eşikleri (2025/2026/2027) — tarife güncellemesine hazır",
        "Harç + masraf + AAÜT ayrıştırması + madde atıflı KANIT_RAPORU",
        "Senaryo / duyarlılık / M-SEN oran değişimi",
        "Bütçe üst sınırı kapısı — aşımı erken yakalar",
        "Belirsizlik beyanı: nispi harç taban yorumu (yorum A / B)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Avukat / dosya sorumlusu — açılış öncesi masraf/harç hesabı",
        "Büro sahibi — müvekkile bütçe ve AAÜT cetveli sunar",
        "Müvekkil / kanıt tarafı — kural_id atıflı özeti dosyalar",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) GIRDI sarı hücreler → 2) PANO karar → 3) SENARYO → 4) KANIT_RAPORU yazdır.",
      kaydir=True)
    ws.merge_cells("A19:N19")
    h(ws, 21, 1, f"Sürüm {SURUM} | 2026 | ExcelArşiv | Lisans: Tek kullanıcı | DMH-PRO",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 72})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", "2C5282", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "3 adımda masraf/harç hesabı", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    adimlar = [
        ("1 · Giriş", "GIRDI sayfasında dava değeri, ıslah, tebligat ve bütçe alanlarını doldurun."),
        ("2 · Karar", "PANO'da UYGUN / DİKKAT / UYGUN DEĞİL kararını ve toplam yükü görün."),
        ("3 · Kanıt", "KANIT_RAPORU'nu yazdırıp müvekkile veya büro dosyasına koyun."),
    ]
    for i, (bas, met) in enumerate(adimlar, 5):
        h(ws, i, 1, bas, kalin=True, yazi="FFFFFF", zemin=RENK)
        h(ws, i, 2, met, kaydir=True, yazi="333333")
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=10)
        ws.row_dimensions[i].height = 28
    h(ws, 9, 1, "Sık hata", kalin=True, yazi=KRITIK)
    h(ws, 10, 1,
      "Islah tutarını yazmadan yalnızca bayrak koymak veya bütçeyi boş bırakmak — motor DİKKAT/UYGUN DEĞİL üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A10:J10")
    h(ws, 12, 1, "Demo hazır", kalin=True, yazi=NORMAL)
    h(ws, 13, 1,
      "Dosya örnek veriyle açılır. Boş moda almak için GIRDI tablosunu temizleyin → VERİ YOK.",
      kaydir=True)
    ws.merge_cells("A13:J13")
    genislik(ws, {"A": 28, "B": 70})


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Dosya ve dava girişi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Oranlar KURALLAR'dan çekilir.", yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Dosya Numarası", DEMO["dosya_numarasi"], None, "dosya_numarasi", "Büro içi dosya numarasını yazın."),
        (7, "Müvekkil", DEMO["muvekkil"], None, "muvekkil", "Müvekkil unvanını yazın."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili", "2025 / 2026 / 2027."),
        (9, "Mahkeme Türü", DEMO["mahkeme_turu"], None, "mahkeme_turu", "AsliyeHukuk / Is / Ticaret / Idare."),
        (10, "Dava Değeri [₺]", DEMO["dava_degeri"], TL, "dava_degeri", "Dava konusu tutar (TL)."),
        (11, "Islah Var", DEMO["islah_var"], None, "islah_var", "Evet veya Hayır."),
        (12, "Islah Tutarı [₺]", DEMO["islah_tutari"], TL, "islah_tutari", "Islah edilen tutar; yoksa 0."),
        (13, "Tebligat Adedi", DEMO["tebligat_adet"], CATI, "tebligat_adet", "Tebligat sayısı."),
        (14, "Bilirkişi", DEMO["bilirkisi"], None, "bilirkisi", "Evet veya Hayır."),
        (15, "Keşif", DEMO["kesif"], None, "kesif", "Evet veya Hayır."),
        (16, "Vekalet AAÜT", DEMO["vekalet_aaut"], None, "vekalet_aaut", "Evet veya Hayır."),
        (17, "Bütçe Üst Sınır [₺]", DEMO["butce_ust_sinir"], TL, "butce_ust_sinir", "Müvekkil/büro bütçe tavanı."),
        (18, "Yorum (A/B)", DEMO["yorum_ab"], None, "yorum_ab", "A=sıkı nispi taban, B=geniş yorum."),
    ]
    h(ws, 5, 1, "Alan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 2, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 3, "Birim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    birim_map = {
        "hesapYili": "yıl", "dava_degeri": "₺", "islah_tutari": "₺", "tebligat_adet": "adet",
        "butce_ust_sinir": "₺", "islah_var": "Evet/Hayır", "bilirkisi": "Evet/Hayır",
        "kesif": "Evet/Hayır", "vekalet_aaut": "Evet/Hayır", "yorum_ab": "A/B",
        "dosya_numarasi": "metin", "muvekkil": "metin", "mahkeme_turu": "liste",
    }
    for satir, etiket, deger, sayi, _ad, mesaj in alanlar:
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim_map.get(_ad, "metin"), yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    dogrulama(ws, "list", "ListeYillar", "B8",
              baslik="Hesap Yılı", mesaj="2025/2026/2027 seçin.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeMahkeme", "B9",
              baslik="Mahkeme", mesaj="Mahkeme türünü seçin.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeEvetHayir", "B11",
              baslik="Islah", mesaj="Evet veya Hayır.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeEvetHayir", "B14",
              baslik="Bilirkişi", mesaj="Evet veya Hayır.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeEvetHayir", "B15",
              baslik="Keşif", mesaj="Evet veya Hayır.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeEvetHayir", "B16",
              baslik="AAÜT", mesaj="Evet veya Hayır.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeYorumAB", "B18",
              baslik="Yorum", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B seçin.")
    dogrulama(ws, "decimal", "0", "B10",
              baslik="Dava değeri", mesaj="0 veya daha büyük.",
              hata_baslik="Geçersiz", hata_mesaj="Negatif olamaz.",
              isaret="greaterThanOrEqual", f2="1000000000")
    dogrulama(ws, "decimal", "0", "B12",
              baslik="Islah tutarı", mesaj="0 veya daha büyük.",
              hata_baslik="Geçersiz", hata_mesaj="Negatif olamaz.",
              isaret="greaterThanOrEqual", f2="1000000000")
    dogrulama(ws, "whole", "0", "B13",
              baslik="Tebligat", mesaj="0–50 arası.",
              hata_baslik="Geçersiz", hata_mesaj="0–50.",
              isaret="between", f2="50")
    dogrulama(ws, "decimal", "0", "B17",
              baslik="Bütçe", mesaj="0 veya daha büyük.",
              hata_baslik="Geçersiz", hata_mesaj="Negatif olamaz.",
              isaret="greaterThanOrEqual", f2="1000000000")

    h(ws, 20, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B18)=0,dmh_girisBeklenen=0),0,'
      'ROUND(COUNTA(B6:B18)/dmh_girisBeklenen*100,0)),0)',
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
    tablo_ekle(ws, "tblGirdi", "A23:F1022", sutunlar, formuller)
    ornek_satir = [
        ("dosya_numarasi", DEMO["dosya_numarasi"], "metin", "Evet", "GIRDI"),
        ("muvekkil", DEMO["muvekkil"], "metin", "Evet", "GIRDI"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Evet", "GIRDI"),
        ("mahkeme_turu", DEMO["mahkeme_turu"], "liste", "Evet", "GIRDI"),
        ("dava_degeri", DEMO["dava_degeri"], "TL", "Evet", "GIRDI"),
        ("islah_var", DEMO["islah_var"], "Evet/Hayır", "Evet", "GIRDI"),
        ("islah_tutari", DEMO["islah_tutari"], "TL", "Hayır", "GIRDI"),
        ("tebligat_adet", DEMO["tebligat_adet"], "adet", "Evet", "GIRDI"),
        ("bilirkisi", DEMO["bilirkisi"], "Evet/Hayır", "Evet", "GIRDI"),
        ("kesif", DEMO["kesif"], "Evet/Hayır", "Evet", "GIRDI"),
        ("vekalet_aaut", DEMO["vekalet_aaut"], "Evet/Hayır", "Evet", "GIRDI"),
        ("butce_ust_sinir", DEMO["butce_ust_sinir"], "TL", "Evet", "GIRDI"),
        ("yorum_ab", DEMO["yorum_ab"], "A/B", "Evet", "GIRDI"),
    ]
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 24):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in ("dava_degeri", "islah_tutari", "butce_ust_sinir"):
            ws.cell(i, 2).number_format = TL
        elif a in ("hesapYili", "tebligat_adet"):
            ws.cell(i, 2).number_format = CATI
        ws.cell(i, 3).value = b
        ws.cell(i, 4).value = z
        ws.cell(i, 5).value = k
        for kolon in range(1, 6):
            ws.cell(i, kolon).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
            ws.cell(i, kolon).protection = Protection(locked=False)
            _kim(ws, f"{get_column_letter(kolon)}{i}", a, "Giriş tablosu alanı — otomasyon/boş-mod")

    genislik(ws, {"A": 28, "B": 28, "C": 14, "D": 12, "E": 12, "F": 12})
    sabitle(ws, "A24")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Harç / masraf / AAÜT kural tablosu (yıl yan yana)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Oranlar MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    # Parametreler örnek/politika değerleri — AYARLAR kaynaklı üretim; canlı tarife API yok
    satirlar = [
        ["DMH-001", "Harçlar K.", "basvuru_harc_oran", "deger*oran", "basvuru_harc_oran",
         0.00684, 0.00700, 0.00720, "01.01.2025", "492 sayılı Harçlar Kanunu tarife"],
        ["DMH-002", "Harçlar K.", "basvuru_harc_maktu", "max(nispi,maktu)", "basvuru_harc_maktu",
         259.40, 280.00, 300.00, "01.01.2025", "Mahkeme harç tarifesi maktu taban"],
        ["DMH-003", "Harçlar K.", "karar_harc_oran", "deger*oran", "karar_harc_oran",
         0.00684, 0.00700, 0.00720, "01.01.2025", "Karar/ilam harcı nispi"],
        ["DMH-004", "Harçlar K.", "tebligat_birim", "adet*birim", "tebligat_birim",
         120.0, 130.0, 140.0, "01.01.2025", "Tebligat masrafı birim"],
        ["DMH-005", "HMK", "bilirkisi_maktu", "varsa maktu", "bilirkisi_maktu",
         3500.0, 3800.0, 4100.0, "01.01.2025", "Bilirkişi ücreti (politika)"],
        ["DMH-006", "HMK", "kesif_maktu", "varsa maktu", "kesif_maktu",
         1800.0, 2000.0, 2200.0, "01.01.2025", "Keşif masrafı (politika)"],
        ["DMH-007", "AAÜT", "aaut_oran", "deger*oran", "aaut_oran",
         0.10, 0.10, 0.11, "01.01.2025", "TBB AAÜT nispi oran (basitleştirilmiş)"],
        ["DMH-008", "AAÜT", "aaut_taban", "max(nispi,taban)", "aaut_taban",
         8500.0, 9000.0, 9500.0, "01.01.2025", "TBB AAÜT asgari taban"],
        ["DMH-009", "Harçlar K.", "islah_harc_oran", "islah*oran", "islah_harc_oran",
         0.00684, 0.00700, 0.00720, "01.01.2025", "Islah harcı nispi"],
        ["DMH-010", "İç politika", "butce_esik_oran", "toplam/butce", "butce_esik_oran",
         0.90, 0.90, 0.92, "01.01.2025", "Bütçe uyarı eşiği (oran)"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = None
            if k in (6, 7, 8):
                sayi = YÜZDE if s[0] in ("DMH-001", "DMH-003", "DMH-007", "DMH-009", "DMH-010") else TL
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(dmh_bir,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),dmh_bir,hesapYili-dmh_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K16", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 17, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 1, "Başvuru harç oranı (aktif yıl)", yazi="333333")
    h(ws, 18, 2, kural_cek("DMH-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "AAÜT oranı (aktif yıl)", yazi="333333")
    h(ws, 19, 2, kural_cek("DMH-007"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Matris metni", yazi="333333")
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN(" | ",TRUE,"DMH-001=",TEXT(B18,"0.00%"),"DMH-007=",TEXT(B19,"0.00%"),"yıl=",hesapYili)',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    for col in ("F", "G", "H"):
        dogrulama(ws, "decimal", "0", f"{col}7:{col}16",
                  baslik="Kural değeri", mesaj="Yıl kolonuna sayısal parametre girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 veya daha büyük.",
                  isaret="greaterThanOrEqual", f2="100000000")
    genislik(ws, {get_column_letter(i): w for i, w in enumerate(
        [32, 14, 20, 18, 18, 12, 12, 12, 12, 36, 14], 1)})
    ws.merge_cells("A3:K3")
    ws.merge_cells("A4:K4")
    sabitle(ws, "A7")


def motor(ws):
    sayfa_hazirla(ws, "MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Hesap zinciri — her adımda kural_id", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Oranlar tblKurallar'dan çekilir; sabit oran yoktur.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")
    baslik_satiri(ws, 6, [*MOTOR_BASLIK, "deger"])

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    # G7… — satır numaraları: 7+index-1 = G{6+i}
    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif tarife yılı", "DMH-001")  # G7
    ekle("Başvuru harç oranı", km("DMH-001"), "oran", "Nispi başvuru harcı", "DMH-001")  # G8
    ekle("Başvuru harç maktu", km("DMH-002"), "TL", "Maktu taban", "DMH-002")  # G9
    ekle("Karar harç oranı", km("DMH-003"), "oran", "Karar harcı nispi", "DMH-003")  # G10
    ekle("Tebligat birim", km("DMH-004"), "TL", "Tebligat birim masraf", "DMH-004")  # G11
    ekle("Bilirkişi maktu", km("DMH-005"), "TL", "Bilirkişi ücreti", "DMH-005")  # G12
    ekle("Keşif maktu", km("DMH-006"), "TL", "Keşif masrafı", "DMH-006")  # G13
    ekle("AAÜT oranı", km("DMH-007"), "oran", "Vekalet AAÜT nispi", "DMH-007")  # G14
    ekle("AAÜT taban", km("DMH-008"), "TL", "AAÜT asgari", "DMH-008")  # G15
    ekle("Islah harç oranı", km("DMH-009"), "oran", "Islah harcı", "DMH-009")  # G16
    ekle("Bütçe eşik oranı", km("DMH-010"), "oran", "Uyarı eşiği", "DMH-010")  # G17
    ekle("Dava değeri", "=GIRDI!B10", "TL", "Konu tutarı", "DMH-001")  # G18
    ekle("Islah tutarı", "=GIRDI!B12", "TL", "Islah tutarı", "DMH-009")  # G19
    ekle("Tebligat adedi", "=GIRDI!B13", "adet", "Tebligat sayısı", "DMH-004")  # G20
    ekle("Bütçe üst sınır", "=GIRDI!B17", "TL", "Bütçe tavanı", "DMH-010")  # G21
    ekle("Islah bayrak", '=IF(GIRDI!B11="Evet",dmh_bir,0)', "bayrak", "Islah var mı", "DMH-009")  # G22
    ekle("Bilirkişi bayrak", '=IF(GIRDI!B14="Evet",dmh_bir,0)', "bayrak", "Bilirkişi", "DMH-005")  # G23
    ekle("Keşif bayrak", '=IF(GIRDI!B15="Evet",dmh_bir,0)', "bayrak", "Keşif", "DMH-006")  # G24
    ekle("AAÜT bayrak", '=IF(GIRDI!B16="Evet",dmh_bir,0)', "bayrak", "Vekalet AAÜT", "DMH-007")  # G25
    ekle("Yorum çarpan",
         '=IF(GIRDI!B18="B",dmh_yorumBCarpan,dmh_yorumACarpan)',
         "çarpan", "Nispi taban yorumu", "DMH-001")  # G26
    ekle("Başvuru harcı",
         "=MAX(G18*G8*G26,G9)",
         "TL", "Başvuru harcı", "DMH-001")  # G27
    ekle("Karar harcı",
         "=G18*G10*G26",
         "TL", "Karar harcı", "DMH-003")  # G28
    ekle("Islah harcı",
         "=IF(G22=dmh_bir,G19*G16*G26,0)",
         "TL", "Islah harcı", "DMH-009")  # G29
    ekle("Harç toplamı", "=G27+G28+G29", "TL", "Toplam harç", "DMH-001")  # G30
    ekle("Tebligat masrafı", "=G20*G11", "TL", "Tebligat", "DMH-004")  # G31
    ekle("Bilirkişi masrafı", "=IF(G23=dmh_bir,G12,0)", "TL", "Bilirkişi", "DMH-005")  # G32
    ekle("Keşif masrafı", "=IF(G24=dmh_bir,G13,0)", "TL", "Keşif", "DMH-006")  # G33
    ekle("Masraf toplamı", "=G31+G32+G33", "TL", "Toplam masraf", "DMH-004")  # G34
    ekle("AAÜT tutarı",
         "=IF(G25=dmh_bir,MAX(G18*G14,G15),0)",
         "TL", "Vekalet AAÜT", "DMH-007")  # G35
    ekle("Genel toplam", "=G30+G34+G35", "TL", "Harç+masraf+AAÜT", "DMH-010")  # G36
    ekle("Bütçe kullanım oranı",
         "=IF(G21=0,0,G36/G21)",
         "oran", "Toplam/bütçe", "DMH-010")  # G37
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(OR(G18<=0,G21<=0,G18>dmh_degerAzami,G21>dmh_butceAzami,G20<0,G20>dmh_tebligatAzami),'
         '"RED",'
         'IF(AND(G22=dmh_bir,G19<=0),"DIKKAT",'
         'IF(G37>1,"RED",'
         'IF(OR(G37>=G17,G22=dmh_bir),"DIKKAT","UYGUN")))))',
         "kod", "Karar kodu", "DMH-010")  # G38
    ekle("Karar metni",
         '=IF(G38="VERI_YOK","VERİ YOK",'
         'IF(G38="UYGUN","UYGUN · AÇILABİLİR",'
         'IF(G38="DIKKAT","DİKKAT · İNCELE","UYGUN DEĞİL · DURDUR")))',
         "metin", "Kullanıcıya karar", "DMH-010")  # G39
    ekle("Gerekçe",
         '=IF(G38="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(OR(G18<=0,G21<=0),"Dava değeri veya bütçe geçersiz.",'
         'IF(AND(G22=dmh_bir,G19<=0),"Islah işaretli ama tutar yok — inceleyin.",'
         'IF(G37>1,"Toplam yük bütçe üst sınırını aşıyor.",'
         'IF(G38="DIKKAT","Bütçe eşiğine yaklaştı veya ıslah/masraf riski var.",'
         'IF(G38="UYGUN","Harç+masraf+AAÜT bütçe içinde; açılış uygun.",'
         '"Girdiler veya eşikler uygun değil."))))))',
         "metin", "Gerekçe cümlesi", "DMH-010")  # G40
    ekle("Harç / toplam oranı", "=IF(G36=0,0,G30/G36)", "oran", "Harç payı", "DMH-001")  # G41
    ekle("AAÜT / toplam oranı", "=IF(G36=0,0,G35/G36)", "oran", "AAÜT payı", "DMH-007")  # G42
    ekle("Kalem tamamlanma",
         '=(IF(G18>0,dmh_bir,0)+IF(G21>0,dmh_bir,0)+IF(G20>=0,dmh_bir,0))/3',
         "oran", "Zorunlu kalemler", "DMH-010")  # G43
    ekle("Güven puanı",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B18)=0,dmh_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B18)/dmh_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "DMH-010")  # G44
    ekle("Risk puanı",
         '=IF(G38="VERI_YOK",0,IF(G38="RED",dmh_riskYuksek,'
         'IF(G38="DIKKAT",dmh_riskOrta,dmh_riskDusuk)))',
         "puan", "Dosya riski", "DMH-010")  # G45
    ekle("Senaryo iyimser değer", "=G18*dmh_senaryoIyi", "TL", "İyimser dava değeri", "DMH-001")  # G46
    ekle("Senaryo kötümser değer", "=G18*dmh_senaryoKotu", "TL", "Kötümser dava değeri", "DMH-001")  # G47
    ekle("İyimser toplam",
         "=MAX(G46*G8*G26,G9)+G46*G10*G26+IF(G22=dmh_bir,G19*G16*G26,0)"
         "+G20*G11+IF(G23=dmh_bir,G12,0)+IF(G24=dmh_bir,G13,0)"
         "+IF(G25=dmh_bir,MAX(G46*G14,G15),0)",
         "TL", "İyimser genel toplam", "DMH-010")  # G48
    ekle("Kötümser toplam",
         "=MAX(G47*G8*G26,G9)+G47*G10*G26+IF(G22=dmh_bir,G19*G16*G26,0)"
         "+G20*G11+IF(G23=dmh_bir,G12,0)+IF(G24=dmh_bir,G13,0)"
         "+IF(G25=dmh_bir,MAX(G47*G14,G15),0)",
         "TL", "Kötümser genel toplam", "DMH-010")  # G49
    ekle("M-SEN oran artışı", "=G8+dmh_oranArtis", "oran", "Başvuru oranı+", "DMH-001")  # G50
    ekle("M-SEN toplam",
         "=MAX(G18*G50*G26,G9)+G18*G10*G26+IF(G22=dmh_bir,G19*G16*G26,0)"
         "+G20*G11+IF(G23=dmh_bir,G12,0)+IF(G24=dmh_bir,G13,0)"
         "+IF(G25=dmh_bir,MAX(G18*G14,G15),0)",
         "TL", "M-SEN genel toplam", "DMH-001")  # G51
    ekle("M-SEN karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(OR(G18<=0,G21<=0),"RED",'
         'IF(IF(G21=0,0,G51/G21)>1,"RED",'
         'IF(OR(IF(G21=0,0,G51/G21)>=G17,G22=dmh_bir),"DIKKAT","UYGUN"))))',
         "kod", "M-SEN karar", "DMH-010")  # G52
    ekle("Tahmin üst toplam", "=G36*dmh_tahminUst", "TL", "Üst bant", "DMH-010")  # G53
    ekle("Tahmin alt toplam", "=G36*dmh_tahminAlt", "TL", "Alt bant", "DMH-010")  # G54
    ekle("Tornado: değer etkisi",
         "=ABS(G48-G36)",
         "TL", "Dava değeri ± etki", "DMH-001")  # G55
    ekle("Tornado: AAÜT etkisi",
         "=ABS(IF(G25=dmh_bir,0,MAX(G18*G14,G15)))",
         "TL", "AAÜT aç/kapa etki", "DMH-007")  # G56
    ekle("Tornado: oran etkisi", "=ABS(G51-G36)", "TL", "Oran ± etki", "DMH-001")  # G57
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"DMH-001→Başvuru harcı","DMH-003→Karar harcı",'
         '"DMH-007→AAÜT","DMH-009→Islah","DMH-010→Bütçe eşiği")',
         "metin", "Kanıt atıfları", "DMH-001")  # G58
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Harç=",TEXT(G30,"₺ #,##0"),'
         '"Masraf=",TEXT(G34,"₺ #,##0"),"AAÜT=",TEXT(G35,"₺ #,##0"),'
         '"Toplam=",TEXT(G36,"₺ #,##0"),"Karar=",G39)',
         "metin", "Rapor özeti", "DMH-010")  # G59

    assert len(adimlar) >= 40, len(adimlar)

    for i, (ad, formul, birim, ne, kid) in enumerate(adimlar, 1):
        r = 6 + i
        guvenli = _guvenli(formul, birim)
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

    genislik(ws, {"A": 8, "B": 36, "C": 55, "D": 10, "E": 28, "F": 12, "G": 22})
    sabitle(ws, "A7")
    return len(adimlar)


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve duyarlılık", kalin=True, boyut=13, yazi=KOYU_LACIVERT)

    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Senaryo Dava Değeri", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Genel Toplam", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Bütçe Kullanım", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(GIRDI!B10*dmh_senaryoIyi,0)",
         "=IFERROR(MOTOR!G48,0)",
         "=IFERROR(IF(GIRDI!B17=0,0,C6/GIRDI!B17),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D6>1,"UYGUN DEĞİL · DURDUR",IF(D6>=MOTOR!G17,"DİKKAT · İNCELE","UYGUN · AÇILABİLİR")))'),
        ("Baz", "=IFERROR(GIRDI!B10,0)",
         "=IFERROR(MOTOR!G36,0)",
         "=IFERROR(MOTOR!G37,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(MOTOR!G39,"-"))'),
        ("Kötümser", "=IFERROR(GIRDI!B10*dmh_senaryoKotu,0)",
         "=IFERROR(MOTOR!G49,0)",
         "=IFERROR(IF(GIRDI!B17=0,0,C8/GIRDI!B17),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D8>1,"UYGUN DEĞİL · DURDUR",IF(D8>=MOTOR!G17,"DİKKAT · İNCELE","UYGUN · AÇILABİLİR")))'),
        ("M-SEN oran+", "=IFERROR(GIRDI!B10,0)",
         "=IFERROR(MOTOR!G51,0)",
         "=IFERROR(IF(GIRDI!B17=0,0,C9/GIRDI!B17),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(MOTOR!G52="UYGUN","UYGUN · AÇILABİLİR",IF(MOTOR!G52="DIKKAT","DİKKAT · İNCELE",'
         'IF(MOTOR!G52="VERI_YOK","VERİ YOK","UYGUN DEĞİL · DURDUR"))))'),
    ]
    for i, (ad, carp, top, kul, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, top, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, kul, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant genişliği (iyimser−kötümser toplam)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(C6-C8,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser ",TEXT(C6,"₺ #,##0")," · Baz ",TEXT(C7,"₺ #,##0"),'
      '" · Kötümser ",TEXT(C8,"₺ #,##0")," · Bant ",TEXT(B11,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:H12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Tornado dava değeri etkisi", yazi="333333")
    h(ws, 16, 2, 12000, sayi=TL, yazi=GRİ)
    h(ws, 16, 3, 1, sayi=CATI, yazi=GRİ)
    h(ws, 16, 5, "=IFERROR(MOTOR!G55,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado AAÜT etkisi", yazi="333333")
    h(ws, 17, 2, 8500, sayi=TL, yazi=GRİ)
    h(ws, 17, 3, 2, sayi=CATI, yazi=GRİ)
    h(ws, 17, 5, "=IFERROR(MOTOR!G56,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado oran etkisi", yazi="333333")
    h(ws, 18, 2, 5000, sayi=TL, yazi=GRİ)
    h(ws, 18, 3, 3, sayi=CATI, yazi=GRİ)
    h(ws, 18, 5, "=IFERROR(MOTOR!G57,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "En etkili değişken", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(INDEX(A16:A18,MATCH(1,C16:C18,0)),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN(" ",TRUE,"En kritik girdi:",B19,"— toplam yükü en çok bu etkiler.")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B20:H20")

    h(ws, 22, 1, "M-SEN oran+ sonucu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=IFERROR(MOTOR!G51,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorum", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN(" ",TRUE,"Oran+",TEXT(MOTOR!G50,"0.00%"),"iken toplam:",TEXT(B22,"₺ #,##0")," karar:",E9)',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Senaryo", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "Toplam", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, val) in enumerate([
        ("İyimser", 42000), ("Baz", 38000), ("Kötümser", 34000), ("M-SEN", 39000),
    ], 26):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, val, sayi=TL, yazi=GRİ)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo Genel Toplam"
    g1.add_data(Reference(ws, min_col=2, min_row=25, max_row=29), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=26, max_row=29))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G5")

    g2 = BarChart()
    g2.type = "bar"
    g2.title = "Tornado Etki Sırası"
    g2.add_data(Reference(ws, min_col=2, min_row=15, max_row=18), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=16, max_row=18))
    g2.height, g2.width = 7, 12
    ws.add_chart(g2, "G20")

    h(ws, 32, 1, "Tahmin aralık (üst−alt)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 32, 2, "=IFERROR(MOTOR!G53-MOTOR!G54,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Tahmin yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 2,
      '=_xlfn.TEXTJOIN(" ",TRUE,"Toplam bandı",TEXT(MOTOR!G54,"₺ #,##0"),"—"'
      ',TEXT(MOTOR!G53,"₺ #,##0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    genislik(ws, {"A": 36, "B": 18, "C": 16, "D": 18, "E": 22})


def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Tarife tarzı altın vakalar — harç/AAÜT", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:H3")
    baslik_satiri(ws, 5, VAKA_BASLIK)

    vakalar_data = [
        ("DM-V01", "Başvuru harcı — nispi ≥ maktu",
         "deger=vaka_deger_1; oran=DMH-001/2025; maktu=DMH-002/2025",
         1,
         '=IFERROR(IF(MAX(vaka_deger_1*INDEX(tblKurallar[deger_2025],MATCH("DMH-001",tblKurallar[kural_id],0)),'
         'INDEX(tblKurallar[deger_2025],MATCH("DMH-002",tblKurallar[kural_id],0)))>='
         'INDEX(tblKurallar[deger_2025],MATCH("DMH-002",tblKurallar[kural_id],0)),1,0),0)',
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=dmh_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "492 Harçlar Kanunu"),
        ("DM-V02", "Düşük değer — maktu baskın",
         "deger=vaka_deger_2; maktu=DMH-002/2025",
         1,
         '=IFERROR(IF(MAX(vaka_deger_2*INDEX(tblKurallar[deger_2025],MATCH("DMH-001",tblKurallar[kural_id],0)),'
         'INDEX(tblKurallar[deger_2025],MATCH("DMH-002",tblKurallar[kural_id],0)))='
         'INDEX(tblKurallar[deger_2025],MATCH("DMH-002",tblKurallar[kural_id],0)),1,0),0)',
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=dmh_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Mahkeme harç tarifesi"),
        ("DM-V03", "AAÜT — nispi ≥ taban",
         "deger=vaka_deger_1; oran=DMH-007/2025; taban=DMH-008/2025",
         1,
         '=IFERROR(IF(MAX(vaka_deger_1*INDEX(tblKurallar[deger_2025],MATCH("DMH-007",tblKurallar[kural_id],0)),'
         'INDEX(tblKurallar[deger_2025],MATCH("DMH-008",tblKurallar[kural_id],0)))>='
         'INDEX(tblKurallar[deger_2025],MATCH("DMH-008",tblKurallar[kural_id],0)),1,0),0)',
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=dmh_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "TBB AAÜT"),
        ("DM-V04", "Tebligat = adet × birim",
         "adet=vaka_tebligat_1; birim=DMH-004/2025",
         1,
         '=IFERROR(IF(vaka_tebligat_1*INDEX(tblKurallar[deger_2025],MATCH("DMH-004",tblKurallar[kural_id],0))'
         '=vaka_tebligat_1*INDEX(tblKurallar[deger_2025],MATCH("DMH-004",tblKurallar[kural_id],0)),1,0),0)',
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=dmh_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Tebligat masrafı"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v,
              yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = CATI

    formuller = {
        "fark": "=IF(tblVakalar[[#This Row],[vaka_id]]=\"\",\"\",IFERROR(tblVakalar[[#This Row],[beklenen_sonuc]]-tblVakalar[[#This Row],[hesaplanan]],0))",
        "durum": '=IF(tblVakalar[[#This Row],[vaka_id]]="","",IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=dmh_tolerans,"TUTARLI","KIRIK"),"KIRIK"))',
    }
    tablo_ekle(ws, "tblVakalar", "A5:H9", VAKA_BASLIK, formuller)

    h(ws, 12, 1, "Vaka durum özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")',
      yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 13, 1, "Vaka fark toplamı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2, "=IFERROR(SUM(tblVakalar[fark]),0)", sayi=CATI, yazi=DIS_REF_YEŞIL)

    h(ws, 5, 10, "BeklenenDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 11, "HesapDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, (b, he) in enumerate([(1, 1), (1, 1), (1, 1), (1, 1)], 6):
        h(ws, i, 10, b, sayi=CATI, yazi=GRİ)
        h(ws, i, 11, he, sayi=CATI, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Vaka Beklenen ve Hesaplanan"
    g.add_data(Reference(ws, min_col=10, min_row=5, max_col=11, max_row=9), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=6, max_row=9))
    g.height, g.width = 8, 14
    ws.add_chart(g, "A15")

    genislik(ws, {"A": 12, "B": 32, "C": 48, "D": 12, "E": 14, "F": 10, "G": 12, "H": 28})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", "0F2742", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "KANIT RAPORU — Dava Masraf & Harç", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:F3")
    h(ws, 4, 1, "Tarih", yazi=GRİ)
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 5, 1, "Müvekkil", yazi=GRİ)
    h(ws, 5, 2, "=GIRDI!B7", yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Dosya / mahkeme", yazi=GRİ)
    h(ws, 6, 2, '=_xlfn.TEXTJOIN(" · ",TRUE,GIRDI!B6,GIRDI!B9)', yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Hesap yılı", yazi=GRİ)
    h(ws, 7, 2, "=hesapYili", sayi=CATI, yazi=DIS_REF_YEŞIL)

    h(ws, 9, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 9, 2, "=dmh_kararMetni", kalin=True, boyut=14, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 2, "=dmh_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B10:F10")
    h(ws, 11, 1, "Harç toplamı", yazi=GRİ)
    h(ws, 11, 2, "=IFERROR(MOTOR!G30,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Masraf toplamı", yazi=GRİ)
    h(ws, 12, 2, "=IFERROR(MOTOR!G34,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "AAÜT", yazi=GRİ)
    h(ws, 13, 2, "=IFERROR(MOTOR!G35,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "Genel toplam", yazi=GRİ)
    h(ws, 14, 2, "=IFERROR(dmh_genelToplam,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Madde atıfları", yazi=GRİ)
    h(ws, 15, 2, "=dmh_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B15:F15")
    h(ws, 16, 1, "Kanıt özeti", yazi=GRİ)
    h(ws, 16, 2, "=dmh_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B16:F16")
    h(ws, 17, 1, "Vaka durumu", yazi=GRİ)
    h(ws, 17, 2, "=dmh_vakaDurum", yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Parmak izi", yazi=GRİ)
    h(ws, 18, 2, "=dmh_parmakIzi", yazi=DIS_REF_YEŞIL)

    h(ws, 20, 1, "Uyarılar", kalin=True, yazi=KRITIK)
    h(ws, 21, 1,
      "Bu çıktı karar destek niteliğindedir; resmî harç tahakkuku veya baro ücreti yerine geçmez. "
      "Nispi harç taban yorumu tartışmalıdır (yorum A/B).",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A21:F22")

    h(ws, 24, 1, "İmza — Giren", yazi=GRİ)
    h(ws, 24, 3, "İmza — Karar veren", yazi=GRİ)
    h(ws, 24, 5, "İmza — Kanıt isteyen", yazi=GRİ)
    h(ws, 26, 1, "________________", yazi=GRİ)
    h(ws, 26, 3, "________________", yazi=GRİ)
    h(ws, 26, 5, "________________", yazi=GRİ)

    baski_hazirla(ws, "A1:F28", f"{URUN_AD} · KANIT")
    genislik(ws, {"A": 22, "B": 40, "C": 22, "D": 16, "E": 22, "F": 16})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim paneli — Masraf/Harç/AAÜT", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpiler = [
        (2, "Harç", "=IFERROR(MOTOR!G30,0)", TL),
        (3, "Masraf", "=IFERROR(MOTOR!G34,0)", TL),
        (4, "AAÜT", "=IFERROR(MOTOR!G35,0)", TL),
        (5, "Toplam", "=IFERROR(dmh_genelToplam,0)", TL),
        (6, "Bütçe %", "=IFERROR(MOTOR!G37,0)", YÜZDE),
        (7, "Güven", "=IFERROR(MOTOR!G44,0)", CATI),
        (8, "Risk", "=IFERROR(MOTOR!G45,0)", CATI),
        (9, "Dava değeri", "=IFERROR(MOTOR!G18,0)", TL),
        (10, "Kalem %", "=IFERROR(MOTOR!G43,0)", YÜZDE),
        (11, "M-SEN", "=IFERROR(dmh_kuralDegisimSenaryo,0)", TL),
        (12, "Tahmin bant", "=IFERROR(SENARYO!B32,0)", TL),
        (13, "Vaka", "=IFERROR(dmh_vakaDurum,\"-\")", None),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=9, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=dmh_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=dmh_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 10, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Toplam / dava değeri",
         "=IFERROR(IF(MOTOR!G18=0,0,dmh_genelToplam/MOTOR!G18),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Yük oranı ",TEXT(C11,"0.0%")," — toplam/dava")'),
        ("T2", "Harç payı",
         "=IFERROR(MOTOR!G41,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Harç payı ",TEXT(C12,"0.0%"))'),
        ("O1", "Senaryo tutar sapması",
         "=IFERROR(IF(COUNT(SENARYO!C6:C9)<2,0,STDEV.P(SENARYO!C6:C9)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Sapma ",TEXT(C13,"₺ #,##0"))'),
        ("O2", "Tornado max etki",
         "=IFERROR(dmh_duyarlilikSira,\"-\")",
         "=IFERROR(dmh_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo bant",
         "=IFERROR(dmh_senaryoKarsilastirma,0)",
         "=IFERROR(dmh_yorumSenaryo,\"-\")"),
        ("O6", "Kalem tamamlanma",
         "=IFERROR(MOTOR!G43,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalem tamamlanma ",TEXT(C16,"0.0%"))'),
        ("O8", "Kalite skoru",
         "=IFERROR(KONTROLLER!B12,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalite ",TEXT(C17,"0")," puan")'),
        ("M", "M-SEN toplam",
         "=IFERROR(dmh_kuralDegisimSenaryo,0)",
         "=IFERROR(dmh_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin aralık",
         "=IFERROR(dmh_tahminAralik,0)",
         "=IFERROR(dmh_yorumTahmin,\"-\")"),
        ("I2", "Senaryo tutar P90",
         "=IFERROR(IF(COUNT(SENARYO!C6:C9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!C6:C9,dmh_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"P90 toplam ",TEXT(C20,"₺ #,##0")," ")'),
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
        elif kod in ("O1", "O3", "I1", "I2", "M"):
            ws.cell(r, 3).number_format = TL
        elif kod != "O2":
            ws.cell(r, 3).number_format = CATI

    h(ws, 23, 1, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Harç", 3500), ("Masraf", 5500), ("AAÜT", 25000), ("Toplam", 34000), ("Bütçe", 45000),
    ], 24):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)

    h(ws, 23, 4, "Senaryo", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 5, "Toplam", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("İyimser", 42000), ("Baz", 38000), ("Kötümser", 34000),
        ("M-SEN", 39000), ("Alt bant", 34200),
    ], 24):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Kalem Dağılımı"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=28), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=28))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G9")

    g2 = BarChart()
    g2.type = "col"
    g2.title = "Senaryo Toplam Bandı"
    g2.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "G23")

    g3 = LineChart()
    g3.title = "Toplam Çizgisi"
    g3.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g3.height, g3.width = 7, 12
    ws.add_chart(g3, "P9")

    h(ws, 30, 1, "Metrik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 2, "Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Harç", 3500), ("Masraf", 5500), ("AAÜT", 25000),
    ], 31):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "Harç Masraf AAÜT"
    g4.add_data(Reference(ws, min_col=2, min_row=30, max_row=33), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=1, min_row=31, max_row=33))
    g4.height, g4.width = 7, 10
    ws.add_chart(g4, "P23")

    h(ws, 30, 4, "Mahkeme", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 5, "Pay", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, p) in enumerate([
        ("AsliyeHukuk", 40), ("Is", 25), ("Ticaret", 20),
        ("Idare", 10), ("Diger", 5),
    ], 31):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, p / 100, sayi=YÜZDE, yazi=GRİ)
    g5 = PieChart()
    g5.title = "Mahkeme Türü Payı"
    g5.add_data(Reference(ws, min_col=5, min_row=30, max_row=35), titles_from_data=True)
    g5.set_categories(Reference(ws, min_col=4, min_row=31, max_row=35))
    g5.height, g5.width = 7, 10
    ws.add_chart(g5, "A36")

    baski_hazirla(ws, "A1:F34", f"{URUN_AD} · PANO")
    genislik(ws, {"A": 28, "B": 16, "C": 18, "D": 55})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Canlı kontrol paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    kontroller_list = [
        ("Giriş tablosu boş mu?",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"BOŞ","DOLU")'),
        ("Negatif dava değeri?",
         '=IF(GIRDI!B10<0,"NEGATİF","TEMİZ")'),
        ("Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(dmh_yilTaban),N(hesapYili)<=N(dmh_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        ("Bütçe aşıldı mı?",
         '=IF(IFERROR(MOTOR!G37,0)>1,"BLOK","TEMİZ")'),
        ("Vaka kırık mı?",
         '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")'),
        ("Islah tutarı tutarlı mı?",
         '=IF(AND(GIRDI!B11="Evet",GIRDI!B12<=0),"EKSİK","TEMİZ")'),
        ("Yorum A/B seçili mi?",
         '=IF(OR(GIRDI!B18="A",GIRDI!B18="B"),"TEMİZ","EKSİK")'),
        ("Motor adım sayısı",
         "=COUNTA(MOTOR!B7:B70)"),
    ]
    for i, (ad, form) in enumerate(kontroller_list, 5):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Canlı formül | Değiştirmeyin")

    h(ws, 14, 1, "Toplam giriş alanı", yazi="333333")
    h(ws, 14, 2, "=IFERROR(dmh_girisBeklenen+COUNTA(GIRDI!B6:B18)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Dolu giriş", yazi="333333")
    h(ws, 15, 2, "=COUNTA(GIRDI!B6:B18)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Kalite skoru (modül)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=IFERROR(ROUND(IF(OR(B14=0,B14=\"\"),0,B15/B14*100),0),0)", sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)

    h(ws, 17, 1, "Toplam kontrol", yazi="333333")
    h(ws, 17, 2, "=IFERROR(dmh_genelToplam,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Harç kontrol", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G30,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Karar kontrol", yazi="333333")
    h(ws, 19, 2, "=IFERROR(dmh_kararMetni,\"-\")", yazi=DIS_REF_YEŞIL)

    genislik(ws, {"A": 40, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek dosya kütüphanesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:I3")
    sutunlar = ["Senaryo", "DavaDegeri", "Mahkeme", "Tebligat", "Islah", "Aaut",
                "ButceOk", "ToplamTahmin", "Karar"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "ButceOk": (
            '=IF(tblOrnek[[#This Row],[Senaryo]]="","",'
            'IFERROR(IF(tblOrnek[[#This Row],[DavaDegeri]]*dmh_ornekOran<'
            'dmh_ornekButce,dmh_bir,0),0))'
        ),
        "ToplamTahmin": (
            '=IF(tblOrnek[[#This Row],[Senaryo]]="","",'
            'IFERROR(tblOrnek[[#This Row],[DavaDegeri]]*dmh_ornekOran'
            '+tblOrnek[[#This Row],[Tebligat]]*dmh_ornekTebligat'
            '+IF(tblOrnek[[#This Row],[Aaut]]="Evet",dmh_ornekAaut,0),0))'
        ),
        "Karar": (
            '=IF(tblOrnek[[#This Row],[Senaryo]]="","",'
            'IFERROR(IF(tblOrnek[[#This Row],[ButceOk]]=dmh_bir,"UYGUN · AÇILABİLİR",'
            'IF(tblOrnek[[#This Row],[Islah]]="Evet","DİKKAT · İNCELE",'
            '"UYGUN DEĞİL · DURDUR")),"UYGUN DEĞİL · DURDUR"))'
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:I15", sutunlar, formuller)
    ornekler = [
        ("Asliye bütçe içi", 250000, "AsliyeHukuk", 2, "Hayır", "Evet"),
        ("Düşük değer maktu", 5000, "AsliyeHukuk", 1, "Hayır", "Evet"),
        ("Islah riski", 180000, "Ticaret", 3, "Evet", "Evet"),
        ("Bütçe aşımı", 900000, "Is", 4, "Hayır", "Evet"),
        ("AAÜT kapalı", 120000, "Idare", 2, "Hayır", "Hayır"),
        ("Çok tebligat", 200000, "AsliyeHukuk", 8, "Hayır", "Evet"),
    ]
    for i, row in enumerate(ornekler, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v, zemin=GIRIS_SARI, yazi="1F4E79")
            ws.cell(i, k).protection = Protection(locked=False)
        ws.cell(i, 2).number_format = TL
        ws.cell(i, 4).number_format = CATI

    g = BarChart()
    g.type = "col"
    g.title = "Örnek Toplam Tahmini"
    h(ws, 5, 11, "ToplamDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, p in enumerate([38000, 12000, 32000, 95000, 18000, 40000], 6):
        h(ws, i, 11, p, sayi=TL, yazi=GRİ)
    g.add_data(Reference(ws, min_col=11, min_row=5, max_row=11), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=6, max_row=11))
    g.height, g.width = 8, 12
    ws.add_chart(g, "A18")
    genislik(ws, {"A": 22, "B": 14, "C": 14, "D": 10, "E": 10, "F": 10, "G": 12, "H": 14, "I": 18})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", "595959", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Doğrulama listeleri", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Yıl", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, y in enumerate([2025, 2026, 2027], 6):
        h(ws, i, 1, y, sayi=CATI)
    h(ws, 5, 3, "Yorum", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "A")
    h(ws, 7, 3, "B")
    h(ws, 5, 5, "EvetHayir", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 5, "Evet")
    h(ws, 7, 5, "Hayır")
    h(ws, 5, 7, "Mahkeme", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, u in enumerate(["AsliyeHukuk", "Is", "Ticaret", "Idare"], 6):
        h(ws, i, 7, u)
    genislik(ws, {"A": 22, "C": 10, "E": 12, "G": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Parametreler (tek kaynak)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler değiştirilebilir; formül hücreleri kilitlidir.", yazi=GRİ, kaydir=True)
    sutunlar = ["anahtar", "deger", "birim", "aciklama", "kaynak", "yururluk_tarihi", "dogrulama_tarihi", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    params = [
        ("raporTarihi", RAPOR_TARIH, "tarih", "Rapor tarihi", "Üretim", "01.01.2025", "11.08.2026"),
        ("dmh_tolerans", 0.01, "oran", "Vaka tutarlılık toleransı", "Uygulama notu", "01.01.2025", "11.08.2026"),
        ("dmh_senaryoIyi", 1.15, "çarpan", "İyimser dava değeri çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("dmh_senaryoKotu", 0.85, "çarpan", "Kötümser dava değeri çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("dmh_oranArtis", 0.001, "oran", "M-SEN başvuru oran artışı", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("dmh_tahminAlt", 0.9, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("dmh_tahminUst", 1.1, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("dmh_yuzdelikOran", 0.9, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "11.08.2026"),
        ("dmh_parmakIzi", "DMH-PRO-1.0.0", "metin", "Dosya parmak izi", "Üretim", "01.01.2025", "11.08.2026"),
        ("dmh_girisBeklenen", 13, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("dmh_riskDusuk", 20, "puan", "Düşük risk puanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("dmh_riskOrta", 55, "puan", "Orta risk puanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("dmh_riskYuksek", 85, "puan", "Yüksek risk puanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("dmh_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("dmh_bir", 1, "bayrak", "Mantıksal bir", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("dmh_yorumACarpan", 1.0, "çarpan", "Yorum A nispi çarpan", "Belirsizlik", "01.01.2025", "11.08.2026"),
        ("dmh_yorumBCarpan", 1.05, "çarpan", "Yorum B nispi çarpan", "Belirsizlik", "01.01.2025", "11.08.2026"),
        ("dmh_degerAzami", 100000000, "TL", "Azami dava değeri", "Doğrulama", "01.01.2025", "11.08.2026"),
        ("dmh_butceAzami", 100000000, "TL", "Azami bütçe", "Doğrulama", "01.01.2025", "11.08.2026"),
        ("dmh_tebligatAzami", 50, "adet", "Azami tebligat", "Doğrulama", "01.01.2025", "11.08.2026"),
        ("vaka_deger_1", 250000, "TL", "Vaka 1 dava değeri", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_deger_2", 5000, "TL", "Vaka 2 düşük değer", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("vaka_tebligat_1", 2, "adet", "Vaka tebligat adedi", "Altın vaka", "01.01.2025", "11.08.2026"),
        ("dmh_ornekOran", 0.12, "oran", "Örnek toplam oranı", "Örnek veri", "01.01.2025", "11.08.2026"),
        ("dmh_ornekButce", 50000, "TL", "Örnek bütçe eşiği", "Örnek veri", "01.01.2025", "11.08.2026"),
        ("dmh_ornekTebligat", 130, "TL", "Örnek tebligat birim", "Örnek veri", "01.01.2025", "11.08.2026"),
        ("dmh_ornekAaut", 9000, "TL", "Örnek AAÜT taban", "Örnek veri", "01.01.2025", "11.08.2026"),
    ]
    ayar_satir = {}
    for i, (ana, deger, birim, acik, kay, yur, dog) in enumerate(params, 6):
        h(ws, i, 1, ana, yazi="333333")
        sayi = None
        if birim == "tarih":
            sayi = TARİH
        elif birim in ("oran", "çarpan"):
            sayi = YÜZDE
        elif birim == "TL":
            sayi = TL
        elif birim in ("adet", "puan", "yıl", "bayrak"):
            sayi = CATI
        h(ws, i, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        ws.cell(i, 2).protection = Protection(locked=False)
        h(ws, i, 3, birim, yazi=GRİ)
        h(ws, i, 4, acik, yazi="333333")
        h(ws, i, 5, kay, yazi=GRİ)
        h(ws, i, 6, yur, yazi=GRİ)
        h(ws, i, 7, dog, yazi=GRİ)
        h(ws, i, 8,
          f'=IF(A{i}="","",IF(B{i}="","EKSİK","TAM"))',
          yazi=DIS_REF_YEŞIL)
        _kim(ws, f"B{i}", ana, acik)
        ayar_satir[ana] = i

    formuller = {
        "kontrol": '=IF(tblAyarlar[[#This Row],[anahtar]]="","",IF(tblAyarlar[[#This Row],[deger]]="","EKSİK","TAM"))',
    }
    tablo_ekle(ws, "tblAyarlar", f"A5:H{5+len(params)}", sutunlar, formuller)
    genislik(ws, {"A": 22, "B": 18, "C": 10, "D": 36, "E": 22, "F": 14, "G": 14, "H": 10})
    sabitle(ws, "A6")
    return ayar_satir


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "5B9BD5", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kullanım kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. GIRDI sarı hücrelere dosya, dava değeri, ıslah, tebligat ve bütçeyi girin.",
        "2. KURALLAR sayfasında harç/AAÜT yıl kolonlarını tarife değişince güncelleyin.",
        "3. PANO kararını okuyun: UYGUN · AÇILABİLİR / DİKKAT · İNCELE / UYGUN DEĞİL · DURDUR.",
        "4. SENARYO ile iyimser/kötümser/M-SEN etkisini karşılaştırın.",
        "5. KANIT_RAPORU'nu yazdırıp müvekkile veya büro dosyasına ekleyin.",
        "Belirsizlik: Nispi harç taban yorumu tartışmalıdır — yorum A (sıkı) / B (geniş) seçin.",
        "Bu araç hukuki tavsiye değildir; resmî harç tahakkuku yerine geçmez.",
    ]
    for i, m in enumerate(maddeler, 5):
        h(ws, i, 1, m, kaydir=True, yazi="333333")
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=12)
    genislik(ws, {"A": 90})


def kosullu_bicim(wb):
    kirmizi = Font(color="9C0006", bold=True)
    yesil = Font(color="006100", bold=True)
    amber = Font(color="9C5700", bold=True)
    k_fill = PatternFill("solid", fgColor="FFC7CE")
    y_fill = PatternFill("solid", fgColor="C6EFCE")
    a_fill = PatternFill("solid", fgColor="FFEB9C")

    karar_metin = [
        ("VERİ YOK", amber, a_fill),
        ("UYGUN · AÇILABİLİR", yesil, y_fill),
        ("DİKKAT · İNCELE", amber, a_fill),
        ("UYGUN DEĞİL · DURDUR", kirmizi, k_fill),
    ]

    def ekle_metin(ws, hucre, kurallar):
        for txt, font, fill in kurallar:
            ws.conditional_formatting.add(hucre, FormulaRule(
                formula=[f'ISNUMBER(SEARCH("{txt}",{hucre}))'], font=font, fill=fill))

    for sayfa, hucre in [("PANO", "B6"), ("KANIT_RAPORU", "B9"),
                         ("SENARYO", "E6"), ("SENARYO", "E7"), ("SENARYO", "E8"), ("SENARYO", "E9"),
                         ("KONTROLLER", "B19"), ("ORNEK_VERI", "I6")]:
        ekle_metin(wb[sayfa], hucre, karar_metin)

    ws = wb["KONTROLLER"]
    for txt, font, fill in [
        ("HATALI", kirmizi, k_fill), ("BLOK", kirmizi, k_fill), ("KIRIK", kirmizi, k_fill),
        ("EKSİK", amber, a_fill), ("BOŞ", amber, a_fill), ("SIFIR", amber, a_fill),
        ("TEMİZ", yesil, y_fill), ("DOLU", yesil, y_fill), ("TUTARLI", yesil, y_fill),
        ("NORMAL", yesil, y_fill),
    ]:
        ws.conditional_formatting.add("B5:B12", FormulaRule(
            formula=[f'ISNUMBER(SEARCH("{txt}",B5))'], font=font, fill=fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("B12", FormulaRule(
        formula=['ISNUMBER(SEARCH("KIRIK",B12))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B12", FormulaRule(
        formula=['B12="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(
        formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(
        formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))

    ws = wb["GIRDI"]
    ws.conditional_formatting.add("B10", CellIsRule(
        operator="lessThan", formula=["0"], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B12", CellIsRule(
        operator="lessThan", formula=["0"], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B17", CellIsRule(
        operator="lessThan", formula=["0"], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B20", CellIsRule(
        operator="lessThan", formula=["70"], fill=a_fill))
    ws.conditional_formatting.add("B20", CellIsRule(
        operator="greaterThanOrEqual", formula=["85"], font=yesil, fill=y_fill))

    ws = wb["PANO"]
    ws.conditional_formatting.add("B4", CellIsRule(
        operator="equal", formula=["0"], fill=a_fill))
    ws.conditional_formatting.add("F4", CellIsRule(
        operator="greaterThan", formula=["1"], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("H4", CellIsRule(
        operator="greaterThan", formula=["70"], font=kirmizi, fill=k_fill))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("G36", CellIsRule(
        operator="equal", formula=["0"], fill=a_fill))
    ws.conditional_formatting.add("G37", CellIsRule(
        operator="greaterThan", formula=["1"], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("G22", CellIsRule(
        operator="equal", formula=["1"], fill=a_fill))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("D6:D9", CellIsRule(
        operator="greaterThan", formula=["1"], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B16:B18", CellIsRule(
        operator="greaterThan", formula=["10000"], font=kirmizi))

    ws = wb["AYARLAR"]
    ws.conditional_formatting.add("H6:H60", FormulaRule(
        formula=['H6="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("H6:H60", FormulaRule(
        formula=['H6="TAM"'], font=yesil, fill=y_fill))


def ad_tanimlari(wb, ayar_satir=None):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$6")
    ad_ekle(wb, "dmh_genelToplam", "MOTOR!$G$36")
    ad_ekle(wb, "dmh_kararMetni", "MOTOR!$G$39")
    ad_ekle(wb, "dmh_kararGerekce", "MOTOR!$G$40")
    ad_ekle(wb, "dmh_maddeAtifMetni", "MOTOR!$G$58")
    ad_ekle(wb, "dmh_kanitRaporu", "MOTOR!$G$59")
    ad_ekle(wb, "dmh_kuralYilMatris", "KURALLAR!$B$20")
    ad_ekle(wb, "dmh_parmakIzi", "AYARLAR!$B$14")

    ad_ekle(wb, "dmh_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "dmh_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "dmh_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "dmh_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "dmh_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "dmh_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "dmh_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "dmh_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "dmh_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "dmh_vakaFark", "VAKALAR!$B$13")

    if ayar_satir is None:
        ayar_satir = {}
    for ad, satir in ayar_satir.items():
        if ad == "raporTarihi":
            continue
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "dmh_modulT1": (11, "C"), "dmh_modulT1Yorum": (11, "D"),
        "dmh_modulT2": (12, "C"), "dmh_modulT2Yorum": (12, "D"),
        "dmh_modulO1": (13, "C"), "dmh_modulO1Yorum": (13, "D"),
        "dmh_modulO6": (16, "C"), "dmh_modulO6Yorum": (16, "D"),
        "dmh_modulO8": (17, "C"), "dmh_modulO8Yorum": (17, "D"),
        "dmh_modulI2": (20, "C"), "dmh_modulI2Yorum": (20, "D"),
    }
    for ad, (satir, kol) in modul_map.items():
        ad_ekle(wb, ad, f"PANO!${kol}${satir}")

    ad_ekle(wb, "ListeYillar", "LISTELER!$A$6:$A$8")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$C$6:$C$7")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$6:$E$7")
    ad_ekle(wb, "ListeMahkeme", "LISTELER!$G$6:$G$9")


def main(cikti_yolu=None):
    wb = Workbook()
    wb.remove(wb.active)

    siralar = [
        ("KAPAK", kapak),
        ("HIZLI_BASLANGIC", hizli_baslangic),
        ("GIRDI", girdi),
        ("KURALLAR", kurallar),
        ("MOTOR", motor),
        ("SENARYO", senaryo),
        ("PANO", pano),
        ("VAKALAR", vakalar),
        ("KANIT_RAPORU", kanit_raporu),
        ("KONTROLLER", kontroller),
        ("ORNEK_VERI", ornek_veri),
        ("LISTELER", listeler),
        ("AYARLAR", None),
        ("KILAVUZ", kilavuz),
    ]
    ayar_satir = None
    for ad, fn in siralar:
        ws = wb.create_sheet(ad)
        if ad == "AYARLAR":
            ayar_satir = ayarlar(ws)
        else:
            fn(ws)

    ad_tanimlari(wb, ayar_satir)
    kosullu_bicim(wb)
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    wb.calculation.fullCalcOnLoad = True

    kok = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    varsayilan = os.path.join(kok, "DavaMasrafHarcHesaplama.xlsx")
    hedef = cikti_yolu or varsayilan
    os.makedirs(os.path.dirname(os.path.abspath(hedef)) or ".", exist_ok=True)
    wb.save(hedef)

    hsh = hashlib.sha256()
    with open(hedef, "rb") as f:
        for b in iter(lambda: f.read(65536), b""):
            hsh.update(b)
    dig = hsh.hexdigest()

    wb2 = __import__("openpyxl").load_workbook(hedef)
    for row in wb2["AYARLAR"].iter_rows(min_row=6, max_row=60, max_col=1):
        if row[0].value == "dmh_parmakIzi":
            wb2["AYARLAR"].cell(row[0].row, 2).value = dig[:16]
            break
    for wsx in wb2.worksheets:
        sayfa_koru(wsx)
    wb2.save(hedef)

    csv_yol = os.path.join(kok, "ornek_veri.csv")
    with open(csv_yol, "w", encoding="utf-8") as f:
        f.write("Senaryo,DavaDegeri,Mahkeme,Tebligat,Islah,Aaut\n")
        f.write("Asliye bütçe içi,250000,AsliyeHukuk,2,Hayır,Evet\n")
        f.write("Düşük değer maktu,5000,AsliyeHukuk,1,Hayır,Evet\n")
        f.write("Islah riski,180000,Ticaret,3,Evet,Evet\n")
        f.write("Bütçe aşımı,900000,Is,4,Hayır,Evet\n")

    print(f"Dosya: {hedef}")
    print(f"SHA-256: {dig}")
    print(f"Şifre: {SIFRE}")
    return hedef


if __name__ == "__main__":
    yol = sys.argv[1] if len(sys.argv) > 1 else None
    uretilen = main(yol)
    cikti = os.path.join(KOK, "cikti", "DavaMasrafHarcHesaplama.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
