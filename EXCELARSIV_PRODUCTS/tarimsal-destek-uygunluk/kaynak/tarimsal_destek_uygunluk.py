#!/usr/bin/env python3
"""Tarımsal Destek/Teşvik Uygunluk Testi — A1 mevzuat motoru (manda v6).

Hibe/teşvik başvuru öncesi ÇKS, alan, çağrı programı ve borç şartlarını
yıllı kural tablosuyla test eder. Domain: tarımsal destek — SGK/e-belge kopyası değil.
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

URUN_AD = "Tarımsal Destek/Teşvik Uygunluk Testi"
SURUM = "1.0.0"
RENK = "1B4332"
RAPOR_TARIH = date(2026, 8, 11)

# Demo: ÇKS'li, 45 dekar, mazot-gübre çağrısı, borçsuz → UYGUN bandı
DEMO = {
    "isletme_adi": "Örnek Tarım İşletmesi",
    "donem": "2026 Çağrı",
    "hesapYili": 2026,
    "cks_kayit": "Evet",
    "toplam_dekar": 45.0,
    "urun_grubu": "Bitkisel",
    "cagri_programi": "MazotGubre",
    "yas": 38,
    "organik_sertifika": "Hayır",
    "hayvan_adet": 0,
    "yatirim_tutari": 0,
    "vergi_sgk_borcu": "Hayır",
    "basvuru_yorumu": "A",
}


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: Destek başvuru uygunluğunu etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Karar ve tahmini destek bozulur | "
        f"Kaynak: Tarım ve Orman Bakanlığı destek tebliğleri",
    )


def kural_cek(kural_id: str) -> str:
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(tdu_bir,MIN(3,IF(OR(hesapYili="",hesapYili=0),tdu_bir,'
        f'hesapYili-tdu_yilTaban))),'
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
      "ÇKS, alan, çağrı programı ve borç durumunu yıllı kural tablosuyla "
      "karşılaştırıp hibe/teşvik başvurusuna UYGUN / ŞARTLI / UYGUN DEĞİL kararı üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:N4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Çok yıllı destek eşikleri (2025/2026/2027) — tebliğ güncellemesine hazır",
        "Çağrı programına göre uygunluk puanı + madde atıflı KANIT_RAPORU",
        "Senaryo / duyarlılık / M-SEN eşik değişimi",
        "Borç ve ÇKS zorunluluk kapısı — hatalı başvuruyu önler",
        "Belirsizlik beyanı: ÇKS güncelleme penceresi (yorum A / B)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Çiftçi / işletme — başvuru öncesi kendi uygunluğunu test eder",
        "Kooperatif / danışman — çağrı seçimi ve karar gerekçesini sunar",
        "Hibe / bakanlık kanıt tarafı — kural_id atıflı özeti dosyalar",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) GIRDI sarı hücreler → 2) PANO karar → 3) SENARYO → 4) KANIT_RAPORU yazdır.",
      kaydir=True)
    ws.merge_cells("A19:N19")
    h(ws, 21, 1, f"Sürüm {SURUM} | 2026 | ExcelArşiv | Lisans: Tek kullanıcı | TDU-PRO",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 72})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", "2D6A4F", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "3 adımda başvuru uygunluğu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    adimlar = [
        ("1 · Giriş", "GIRDI sayfasında ÇKS, dekar, çağrı programı ve borç alanlarını doldurun."),
        ("2 · Karar", "PANO'da UYGUN / ŞARTLI / UYGUN DEĞİL kararını ve puanı görün."),
        ("3 · Kanıt", "KANIT_RAPORU'nu yazdırıp kooperatif veya danışmana sunun."),
    ]
    for i, (bas, met) in enumerate(adimlar, 5):
        h(ws, i, 1, bas, kalin=True, yazi="FFFFFF", zemin=RENK)
        h(ws, i, 2, met, kaydir=True, yazi="333333")
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=10)
        ws.row_dimensions[i].height = 28
    h(ws, 9, 1, "Sık hata", kalin=True, yazi=KRITIK)
    h(ws, 10, 1,
      "ÇKS kaydı yokken veya vergi/SGK borcu varken başvurmak — motor UYGUN DEĞİL üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A10:J10")
    h(ws, 12, 1, "Demo hazır", kalin=True, yazi=NORMAL)
    h(ws, 13, 1,
      "Dosya örnek veriyle açılır. Boş moda almak için GIRDI tablosunu temizleyin → VERİ YOK.",
      kaydir=True)
    ws.merge_cells("A13:J13")
    genislik(ws, {"A": 18, "B": 70})


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "İşletme ve başvuru girişi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Eşikler KURALLAR'dan çekilir.", yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "İşletme Adı", DEMO["isletme_adi"], None, "isletme_adi", "Resmi işletme unvanını yazın."),
        (7, "Dönem / Çağrı Etiketi", DEMO["donem"], None, "donem", "Çağrı dönemini yazın."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili", "2025 / 2026 / 2027."),
        (9, "ÇKS Kayıt", DEMO["cks_kayit"], None, "cks_kayit", "Evet veya Hayır."),
        (10, "Toplam Alan [dekar]", DEMO["toplam_dekar"], "0.0", "toplam_dekar", "Dekar cinsinden alan."),
        (11, "Ürün Grubu", DEMO["urun_grubu"], None, "urun_grubu", "Bitkisel / Hayvansal / Karma."),
        (12, "Çağrı Programı", DEMO["cagri_programi"], None, "cagri_programi",
         "MazotGubre / AlanBazli / GencCiftci / Organik / Hayvancilik / IPARD."),
        (13, "Başvuran Yaşı", DEMO["yas"], CATI, "yas", "Tam yaş (yıl)."),
        (14, "Organik Sertifika", DEMO["organik_sertifika"], None, "organik_sertifika", "Evet veya Hayır."),
        (15, "Hayvan Adedi", DEMO["hayvan_adet"], CATI, "hayvan_adet", "Küpe kayıtlı hayvan sayısı."),
        (16, "Yatırım Tutarı [₺]", DEMO["yatirim_tutari"], TL, "yatirim_tutari", "IPARD için yatırım tutarı."),
        (17, "Vergi/SGK Borcu", DEMO["vergi_sgk_borcu"], None, "vergi_sgk_borcu", "Evet veya Hayır."),
        (18, "Başvuru Yorumu (A/B)", DEMO["basvuru_yorumu"], None, "basvuru_yorumu",
         "A=sıkı ÇKS penceresi, B=geniş yorum."),
    ]
    h(ws, 5, 1, "Alan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 2, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 3, "Birim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    birim_map = {
        "hesapYili": "yıl", "toplam_dekar": "dekar", "yas": "yıl", "hayvan_adet": "adet",
        "yatirim_tutari": "₺", "cks_kayit": "Evet/Hayır", "organik_sertifika": "Evet/Hayır",
        "vergi_sgk_borcu": "Evet/Hayır", "basvuru_yorumu": "A/B",
        "isletme_adi": "metin", "donem": "metin", "urun_grubu": "liste", "cagri_programi": "liste",
    }
    for satir, etiket, deger, sayi, _ad, mesaj in alanlar:
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim_map.get(_ad, "metin"), yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    dogrulama(ws, "list", "ListeYillar", "B8",
              baslik="Hesap Yılı", mesaj="2025/2026/2027 seçin.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeEvetHayir", "B9",
              baslik="ÇKS", mesaj="Evet veya Hayır.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeUrun", "B11",
              baslik="Ürün", mesaj="Ürün grubunu seçin.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeCagri", "B12",
              baslik="Çağrı", mesaj="Programı seçin.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeEvetHayir", "B14",
              baslik="Organik", mesaj="Evet veya Hayır.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeEvetHayir", "B17",
              baslik="Borç", mesaj="Evet veya Hayır.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeYorumAB", "B18",
              baslik="Yorum", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B seçin.")
    dogrulama(ws, "decimal", "0", "B10",
              baslik="Dekar", mesaj="0 veya daha büyük.",
              hata_baslik="Geçersiz", hata_mesaj="Negatif olamaz.",
              isaret="greaterThanOrEqual", f2="100000")
    dogrulama(ws, "whole", "18", "B13",
              baslik="Yaş", mesaj="18–80 arası.",
              hata_baslik="Geçersiz", hata_mesaj="18–80.",
              isaret="between", f2="80")

    h(ws, 20, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B18)=0,tdu_girisBeklenen=0),0,'
      'ROUND(COUNTA(B6:B18)/tdu_girisBeklenen*100,0)),0)',
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
        ("isletme_adi", DEMO["isletme_adi"], "metin", "Evet", "GIRDI"),
        ("donem", DEMO["donem"], "metin", "Evet", "GIRDI"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Evet", "GIRDI"),
        ("cks_kayit", DEMO["cks_kayit"], "Evet/Hayır", "Evet", "GIRDI"),
        ("toplam_dekar", DEMO["toplam_dekar"], "dekar", "Evet", "GIRDI"),
        ("urun_grubu", DEMO["urun_grubu"], "liste", "Evet", "GIRDI"),
        ("cagri_programi", DEMO["cagri_programi"], "liste", "Evet", "GIRDI"),
        ("yas", DEMO["yas"], "yıl", "Evet", "GIRDI"),
        ("organik_sertifika", DEMO["organik_sertifika"], "Evet/Hayır", "Hayır", "GIRDI"),
        ("hayvan_adet", DEMO["hayvan_adet"], "adet", "Hayır", "GIRDI"),
        ("yatirim_tutari", DEMO["yatirim_tutari"], "TL", "Hayır", "GIRDI"),
        ("vergi_sgk_borcu", DEMO["vergi_sgk_borcu"], "Evet/Hayır", "Evet", "GIRDI"),
        ("basvuru_yorumu", DEMO["basvuru_yorumu"], "A/B", "Evet", "GIRDI"),
    ]
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 24):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a == "yatirim_tutari":
            ws.cell(i, 2).number_format = TL
        elif a == "toplam_dekar":
            ws.cell(i, 2).number_format = "0.0"
        elif a in ("hesapYili", "yas", "hayvan_adet"):
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
    h(ws, 3, 1, "Destek mevzuat kural tablosu (yıl yan yana)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Eşikler MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["TDU-001", "ÇKS Tebliği", "cks_zorunlu", "cks=Evet", "cks_zorunlu",
         1, 1, 1, "01.01.2025", "ÇKS kayıt zorunluluğu"],
        ["TDU-002", "Mazot-Gübre", "min_dekar_mazot", "dekar>=min", "min_dekar_mazot",
         10, 10, 12, "01.01.2025", "Mazot-gübre destek tebliği"],
        ["TDU-003", "Alan Bazlı", "min_dekar_alan", "dekar>=min", "min_dekar_alan",
         1, 1, 1, "01.01.2025", "Alan bazlı destek tebliği"],
        ["TDU-004", "Genç Çiftçi", "yas_ust_sinir", "yas<=ust", "yas_ust_sinir",
         40, 40, 41, "01.01.2025", "Genç çiftçi destek tebliği"],
        ["TDU-005", "Organik", "organik_min_dekar", "sertifika+dekar", "organik_min_dekar",
         5, 5, 5, "01.01.2025", "Organik tarım destek tebliği"],
        ["TDU-006", "Hayvancılık", "hayvan_min_adet", "adet>=min", "hayvan_min_adet",
         10, 10, 10, "01.01.2025", "Hayvancılık destek tebliği"],
        ["TDU-007", "IPARD", "yatirim_min_tl", "yatirim>=min", "yatirim_min_tl",
         500000, 550000, 600000, "01.01.2025", "IPARD başvuru rehberi"],
        ["TDU-008", "Uygunluk", "puan_esik", "puan>=esik", "puan_esik",
         70, 70, 75, "01.01.2025", "İç uygunluk eşik politikası"],
        ["TDU-009", "IPARD", "ipard_hibe_oran", "yatirim*oran", "ipard_hibe_oran",
         0.50, 0.50, 0.55, "01.01.2025", "IPARD hibe oranı"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = None
            if k in (6, 7, 8):
                if s[0] == "TDU-007":
                    sayi = TL
                elif s[0] == "TDU-009":
                    sayi = YÜZDE
                else:
                    sayi = CATI
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(tdu_bir,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),tdu_bir,hesapYili-tdu_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K15", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 17, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 1, "Mazot min dekar (aktif yıl)", yazi="333333")
    h(ws, 18, 2, kural_cek("TDU-002"), sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Uygunluk puan eşiği", yazi="333333")
    h(ws, 19, 2, kural_cek("TDU-008"), sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Matris metni", yazi="333333")
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN(" | ",TRUE,"TDU-002=",TEXT(B18,"0"),"TDU-008=",TEXT(B19,"0"),"yıl=",hesapYili)',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    for col in ("F", "G", "H"):
        dogrulama(ws, "decimal", "0", f"{col}7:{col}15",
                  baslik="Kural değeri", mesaj="Yıl kolonuna sayısal parametre girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 veya daha büyük.",
                  isaret="greaterThanOrEqual", f2="100000000")
    genislik(ws, {get_column_letter(i): w for i, w in enumerate(
        [12, 14, 20, 18, 18, 12, 12, 12, 12, 28, 14], 1)})
    ws.merge_cells("A3:K3")
    ws.merge_cells("A4:K4")
    sabitle(ws, "A7")


def motor(ws):
    sayfa_hazirla(ws, "MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Hesap zinciri — her adımda kural_id", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Eşikler tblKurallar'dan çekilir; sabit eşik yoktur.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")
    baslik_satiri(ws, 6, [*MOTOR_BASLIK, "deger"])

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    # G7…
    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif mevzuat yılı", "TDU-008")
    ekle("ÇKS zorunlu bayrak", km("TDU-001"), "bayrak", "ÇKS zorunluluğu", "TDU-001")
    ekle("Mazot min dekar", km("TDU-002"), "dekar", "Mazot-gübre alt sınır", "TDU-002")
    ekle("Alan bazlı min dekar", km("TDU-003"), "dekar", "Alan bazlı alt sınır", "TDU-003")
    ekle("Genç çiftçi yaş üst", km("TDU-004"), "yıl", "Yaş üst sınırı", "TDU-004")
    ekle("Organik min dekar", km("TDU-005"), "dekar", "Organik alt sınır", "TDU-005")
    ekle("Hayvan min adet", km("TDU-006"), "adet", "Hayvancılık alt sınır", "TDU-006")
    ekle("IPARD yatırım min", km("TDU-007"), "TL", "IPARD yatırım eşiği", "TDU-007")
    ekle("Uygunluk puan eşiği", km("TDU-008"), "puan", "Karar eşiği", "TDU-008")
    ekle("ÇKS kayıt (1/0)", '=IF(GIRDI!B9="Evet",tdu_bir,0)', "bayrak", "ÇKS girişi", "TDU-001")
    ekle("Toplam dekar", "=GIRDI!B10", "dekar", "İşletme alanı", "TDU-002")
    ekle("Yaş", "=GIRDI!B13", "yıl", "Başvuran yaşı", "TDU-004")
    ekle("Organik (1/0)", '=IF(GIRDI!B14="Evet",tdu_bir,0)', "bayrak", "Sertifika", "TDU-005")
    ekle("Hayvan adedi", "=GIRDI!B15", "adet", "Küpe kayıt", "TDU-006")
    ekle("Yatırım tutarı", "=GIRDI!B16", "TL", "IPARD yatırım", "TDU-007")
    ekle("Borç var (1/0)", '=IF(GIRDI!B17="Evet",tdu_bir,0)', "bayrak", "Vergi/SGK borcu", "TDU-001")
    ekle("ÇKS şart sonucu",
         '=IF(G8=0,tdu_bir,IF(G16=tdu_bir,tdu_bir,0))',
         "bayrak", "ÇKS geçti mi", "TDU-001")
    ekle("Borç blok", "=G22", "bayrak", "Borç varsa blok", "TDU-001")
    ekle("Mazot şart", '=IF(G17>=G9,tdu_bir,0)', "bayrak", "Dekar ≥ mazot min", "TDU-002")
    ekle("Alan bazlı şart", '=IF(G17>=G10,tdu_bir,0)', "bayrak", "Dekar ≥ alan min", "TDU-003")
    ekle("Genç çiftçi şart", '=IF(AND(G18>0,G18<=G11),tdu_bir,0)', "bayrak", "Yaş uygun mu", "TDU-004")
    ekle("Organik şart", '=IF(AND(G19=tdu_bir,G17>=G12),tdu_bir,0)', "bayrak", "Organik+dekar", "TDU-005")
    ekle("Hayvan şart", '=IF(G20>=G13,tdu_bir,0)', "bayrak", "Hayvan ≥ min", "TDU-006")
    ekle("IPARD şart", '=IF(G21>=G14,tdu_bir,0)', "bayrak", "Yatırım ≥ min", "TDU-007")
    ekle("Program şart bayrağı",
         '=IF(GIRDI!B12="MazotGubre",G25,'
         'IF(GIRDI!B12="AlanBazli",G26,'
         'IF(GIRDI!B12="GencCiftci",G27,'
         'IF(GIRDI!B12="Organik",G28,'
         'IF(GIRDI!B12="Hayvancilik",G29,'
         'IF(GIRDI!B12="IPARD",G30,0))))))',
         "bayrak", "Seçili çağrı şartı", "TDU-008")
    ekle("Zorunlu şart puanı",
         '=IF(G23=0,0,tdu_puanCks)+IF(G24=0,tdu_puanBorc,0)+IF(G31=tdu_bir,tdu_puanProgram,0)',
         "puan", "ÇKS+borçsuz+program", "TDU-001")
    ekle("Alan puanı",
         '=MIN(tdu_puanAlanMax,IF(G17<=0,0,ROUND(G17/MAX(G9,tdu_bir)*tdu_puanAlanCarpan,0)))',
         "puan", "Dekar yoğunluğu", "TDU-002")
    ekle("Belge puanı",
         '=IF(GIRDI!B18="B",tdu_puanYorumB,tdu_puanYorumA)'
         '+IF(G19=tdu_bir,tdu_puanOrganikEk,0)+IF(G20>0,tdu_puanHayvanEk,0)',
         "puan", "Yorum+belge", "TDU-005")
    ekle("Uygunluk puanı", "=MIN(100,G32+G33+G34)", "puan", "Toplam puan", "TDU-008")
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(OR(G17<tdu_dekarMinGecerli,G17>tdu_dekarAzami,G18<tdu_yasMinGecerli,G18>tdu_yasAzami,'
         'G20<0,G20>tdu_hayvanAzami,G21<0,G21>tdu_yatirimAzami),"RED",'
         'IF(OR(G23=0,G24=tdu_bir),"RED",'
         'IF(AND(G31=tdu_bir,G35>=G15),"UYGUN",'
         'IF(AND(G31=tdu_bir,G35>=G15-tdu_sartliBand),"SARTLI","RED")))))',
         "kod", "Karar kodu", "TDU-008")
    ekle("Karar metni",
         '=IF(G36="VERI_YOK","VERİ YOK",'
         'IF(G36="UYGUN","UYGUN · BAŞVUR",'
         'IF(G36="SARTLI","ŞARTLI · BELGE EKLE","UYGUN DEĞİL · DURDUR")))',
         "metin", "Kullanıcıya karar", "TDU-008")
    ekle("Gerekçe",
         '=IF(G36="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(OR(G17<tdu_dekarMinGecerli,G17>tdu_dekarAzami,G18<tdu_yasMinGecerli,G18>tdu_yasAzami,'
         'G20<0,G20>tdu_hayvanAzami,G21<0,G21>tdu_yatirimAzami),'
         '"Alan/yaş/hayvan/yatırım uç veya geçersiz — başvuru uygun değildir.",'
         'IF(G24=tdu_bir,"Vergi/SGK borcu varken başvuru uygun değildir.",'
         'IF(G23=0,"ÇKS kaydı zorunlu; kayıt yoksa başvuru reddedilir.",'
         'IF(G31=0,"Seçilen çağrı programının asgari şartı sağlanmadı.",'
         'IF(G36="UYGUN","Puan eşik üstünde ve zorunlu şartlar tamam.",'
         'IF(G36="SARTLI","Puan sınır bandında; eksik belge tamamlanmalı.",'
         '"Puan eşik altında veya şartlar eksik.")))))))',
         "metin", "Gerekçe cümlesi", "TDU-008")
    ekle("Tahmini destek birim",
         '=IF(GIRDI!B12="MazotGubre",tdu_birimMazot,'
         'IF(GIRDI!B12="AlanBazli",tdu_birimAlan,'
         'IF(GIRDI!B12="GencCiftci",tdu_birimGenc,'
         'IF(GIRDI!B12="Organik",tdu_birimOrganik,'
         'IF(GIRDI!B12="Hayvancilik",tdu_birimHayvan,'
         'IF(GIRDI!B12="IPARD",tdu_birimIpard,0))))))',
         "TL", "Program birim tutarı", "TDU-002")
    ekle("Tahmini destek tutarı",
         '=IF(OR(G36="VERI_YOK",G36="RED"),0,'
         'IF(GIRDI!B12="IPARD",ROUND(G21*INDEX(tblKurallar[aktif_deger],'
         'MATCH("TDU-009",tblKurallar[kural_id],0)),0),'
         'IF(GIRDI!B12="Hayvancilik",ROUND(G20*G39,0),ROUND(G17*G39,0))))',
         "TL", "Tahmini hibe/teşvik", "TDU-007")
    ekle("Puan / eşik oranı", "=IF(G15=0,0,G35/G15)", "oran", "Eşik yoğunluğu", "TDU-008")
    ekle("Zorunlu şart oranı",
         '=(IF(G23=tdu_bir,tdu_bir,0)+IF(G24=0,tdu_bir,0)+IF(G31=tdu_bir,tdu_bir,0))/3',
         "oran", "Zorunlu tamamlanma", "TDU-001")
    ekle("Güven puanı",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B18)=0,tdu_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B18)/tdu_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "TDU-008")
    ekle("Risk puanı",
         '=IF(G36="VERI_YOK",0,IF(G36="RED",tdu_riskYuksek,'
         'IF(G36="SARTLI",tdu_riskOrta,tdu_riskDusuk)))',
         "puan", "Başvuru riski", "TDU-008")
    ekle("Senaryo iyimser dekar", "=G17*tdu_senaryoIyi", "dekar", "İyimser alan", "TDU-002")
    ekle("Senaryo kötümser dekar", "=G17*tdu_senaryoKotu", "dekar", "Kötümser alan", "TDU-002")
    ekle("İyimser puan",
         "=MIN(100,IF(G45>=G9,tdu_puanCks,0)+IF(G24=0,tdu_puanBorc,0)+IF(G23=tdu_bir,tdu_puanCks,0)"
         "+MIN(tdu_puanAlanMax,ROUND(G45/MAX(G9,tdu_bir)*tdu_puanAlanCarpan,0))+G34)",
         "puan", "İyimser uygunluk", "TDU-008")
    ekle("Kötümser puan",
         "=MIN(100,IF(G46>=G9,tdu_puanCks,0)+IF(G24=0,tdu_puanBorc,0)+IF(G23=tdu_bir,tdu_puanCks,0)"
         "+MIN(tdu_puanAlanMax,ROUND(G46/MAX(G9,tdu_bir)*tdu_puanAlanCarpan,0))"
         "+MAX(0,G34-tdu_puanYorumA))",
         "puan", "Kötümser uygunluk", "TDU-008")
    ekle("M-SEN eşik+5 puan", "=G15+tdu_oranArtisPuan", "puan", "Eşik yükselince", "TDU-008")
    ekle("M-SEN karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(OR(G17<tdu_dekarMinGecerli,G17>tdu_dekarAzami,G18<tdu_yasMinGecerli,G18>tdu_yasAzami,'
         'G20<0,G20>tdu_hayvanAzami,G21<0,G21>tdu_yatirimAzami),"RED",'
         'IF(OR(G23=0,G24=tdu_bir),"RED",'
         'IF(AND(G31=tdu_bir,G35>=G49),"UYGUN",'
         'IF(AND(G31=tdu_bir,G35>=G49-tdu_sartliBand),"SARTLI","RED")))))',
         "kod", "M-SEN karar", "TDU-008")
    ekle("Tahmin üst destek", "=G40*tdu_tahminUst", "TL", "Üst bant", "TDU-007")
    ekle("Tahmin alt destek", "=G40*tdu_tahminAlt", "TL", "Alt bant", "TDU-007")
    ekle("Tornado: dekar etkisi",
         "=ABS(MIN(100,G32+MIN(tdu_puanAlanMax,ROUND((G17*tdu_tornadoDekar)/MAX(G9,tdu_bir)*tdu_puanAlanCarpan,0))+G34)-G35)",
         "puan", "Dekar ± etki", "TDU-002")
    ekle("Tornado: yaş etkisi",
         "=ABS(IF(AND(G18*tdu_tornadoYas>0,G18*tdu_tornadoYas<=G11),G35,"
         "MAX(0,G35-tdu_sartliBand))-G35)",
         "puan", "Yaş ± etki", "TDU-004")
    ekle("Tornado: eşik etkisi", "=ABS(G49-G15)", "puan", "Eşik ± etki", "TDU-008")
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"TDU-001→ÇKS","TDU-002→Mazot-Gübre",'
         '"TDU-004→Genç Çiftçi","TDU-007→IPARD","TDU-008→Puan eşiği")',
         "metin", "Kanıt atıfları", "TDU-001")
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Puan=",TEXT(G35,"0"),'
         '"Destek=",TEXT(G40,"₺ #,##0"),"Karar=",G37)',
         "metin", "Rapor özeti", "TDU-008")

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
        elif birim == "oran":
            ws.cell(r, 7).number_format = YÜZDE
            ws.cell(r, 3).number_format = YÜZDE
        elif birim == "dekar":
            ws.cell(r, 7).number_format = "0.0"
            ws.cell(r, 3).number_format = "0.0"
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
    h(ws, 5, 2, "Senaryo Dekar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Uygunluk Puanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Tahmini Destek", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(GIRDI!B10*tdu_senaryoIyi,0)",
         "=IFERROR(MOTOR!G47,0)",
         "=IFERROR(MOTOR!G40*tdu_senaryoIyi,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(C6>=MOTOR!G15,"UYGUN · BAŞVUR",IF(C6>=MOTOR!G15-tdu_sartliBand,"ŞARTLI · BELGE EKLE","UYGUN DEĞİL · DURDUR")))'),
        ("Baz", "=IFERROR(GIRDI!B10,0)",
         "=IFERROR(MOTOR!G35,0)",
         "=IFERROR(MOTOR!G40,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(MOTOR!G37,"-"))'),
        ("Kötümser", "=IFERROR(GIRDI!B10*tdu_senaryoKotu,0)",
         "=IFERROR(MOTOR!G48,0)",
         "=IFERROR(MOTOR!G40*tdu_senaryoKotu,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(C8>=MOTOR!G15,"UYGUN · BAŞVUR",IF(C8>=MOTOR!G15-tdu_sartliBand,"ŞARTLI · BELGE EKLE","UYGUN DEĞİL · DURDUR")))'),
        ("M-SEN eşik+5", "=IFERROR(GIRDI!B10,0)",
         "=IFERROR(MOTOR!G35,0)",
         "=IFERROR(MOTOR!G40,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(MOTOR!G50="UYGUN","UYGUN · BAŞVUR",IF(MOTOR!G50="SARTLI","ŞARTLI · BELGE EKLE",'
         'IF(MOTOR!G50="VERI_YOK","VERİ YOK","UYGUN DEĞİL · DURDUR"))))'),
    ]
    for i, (ad, carp, puan, dest, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi="0.0", yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, puan, sayi=CATI, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, dest, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant genişliği (iyimser−kötümser puan)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(C6-C8,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser ",TEXT(C6,"0")," · Baz ",TEXT(C7,"0"),'
      '" · Kötümser ",TEXT(C8,"0")," · Bant ",TEXT(B11,"0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:H12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [puan]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Tornado dekar etkisi", yazi="333333")
    h(ws, 16, 2, 12, sayi=CATI, yazi=GRİ)
    h(ws, 16, 3, 1, sayi=CATI, yazi=GRİ)
    h(ws, 16, 5, "=IFERROR(MOTOR!G53,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado yaş etkisi", yazi="333333")
    h(ws, 17, 2, 8, sayi=CATI, yazi=GRİ)
    h(ws, 17, 3, 2, sayi=CATI, yazi=GRİ)
    h(ws, 17, 5, "=IFERROR(MOTOR!G54,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado eşik etkisi", yazi="333333")
    h(ws, 18, 2, 5, sayi=CATI, yazi=GRİ)
    h(ws, 18, 3, 3, sayi=CATI, yazi=GRİ)
    h(ws, 18, 5, "=IFERROR(MOTOR!G55,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "En etkili değişken", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(INDEX(A16:A18,MATCH(1,C16:C18,0)),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN(" ",TRUE,"En kritik girdi:",B19,"— başvuru puanını en çok bu etkiler.")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B20:H20")

    h(ws, 22, 1, "M-SEN eşik+5 sonucu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=IFERROR(MOTOR!G49,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorum", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN(" ",TRUE,"Eşik",TEXT(MOTOR!G49,"0"),"iken karar:",E9)',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Senaryo", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "Puan", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, val) in enumerate([
        ("İyimser", 85), ("Baz", 78), ("Kötümser", 62), ("M-SEN", 78),
    ], 26):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, val, sayi=CATI, yazi=GRİ)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo Uygunluk Puanı"
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
    h(ws, 32, 2, "=IFERROR(MOTOR!G51-MOTOR!G52,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Tahmin yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 2,
      '=_xlfn.TEXTJOIN(" ",TRUE,"Destek bandı",TEXT(MOTOR!G52,"₺ #,##0"),"—"'
      ',TEXT(MOTOR!G51,"₺ #,##0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    genislik(ws, {"A": 36, "B": 18, "C": 16, "D": 18, "E": 18})


def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Tebliğ tarzı altın vakalar — destek uygunluk", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:H3")
    baslik_satiri(ws, 5, VAKA_BASLIK)

    # V1: 45 dekar ≥ 10 → 1
    # V2: 8 dekar < 10 → 0
    # V3: yaş 38 ≤ 40 → 1
    # V4: yatırım 600k ≥ 500k (2025) → 1
    vakalar_data = [
        ("TD-V01", "Mazot min dekar — uygun",
         "dekar=vaka_dekar_1; min=TDU-002/2025",
         1,
         '=IFERROR(IF(vaka_dekar_1>=INDEX(tblKurallar[deger_2025],MATCH("TDU-002",tblKurallar[kural_id],0)),1,0),0)',
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=tdu_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Mazot-gübre destek tebliği"),
        ("TD-V02", "Mazot min dekar — yetersiz",
         "dekar=vaka_dekar_2; min=TDU-002/2025",
         0,
         '=IFERROR(IF(vaka_dekar_2>=INDEX(tblKurallar[deger_2025],MATCH("TDU-002",tblKurallar[kural_id],0)),1,0),0)',
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=tdu_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Mazot-gübre destek tebliği"),
        ("TD-V03", "Genç çiftçi yaş uygun",
         "yas=vaka_yas_1; ust=TDU-004/2025",
         1,
         '=IFERROR(IF(AND(vaka_yas_1>0,vaka_yas_1<=INDEX(tblKurallar[deger_2025],MATCH("TDU-004",tblKurallar[kural_id],0))),1,0),0)',
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=tdu_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Genç çiftçi destek tebliği"),
        ("TD-V04", "IPARD yatırım eşiği",
         "yatirim=vaka_yatirim_1; min=TDU-007/2025",
         1,
         '=IFERROR(IF(vaka_yatirim_1>=INDEX(tblKurallar[deger_2025],MATCH("TDU-007",tblKurallar[kural_id],0)),1,0),0)',
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=tdu_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "IPARD başvuru rehberi"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v,
              yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = CATI

    formuller = {
        "fark": "=IF(tblVakalar[[#This Row],[vaka_id]]=\"\",\"\",IFERROR(tblVakalar[[#This Row],[beklenen_sonuc]]-tblVakalar[[#This Row],[hesaplanan]],0))",
        "durum": '=IF(tblVakalar[[#This Row],[vaka_id]]="","",IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=tdu_tolerans,"TUTARLI","KIRIK"),"KIRIK"))',
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
    for i, (b, he) in enumerate([(1, 1), (0, 0), (1, 1), (1, 1)], 6):
        h(ws, i, 10, b, sayi=CATI, yazi=GRİ)
        h(ws, i, 11, he, sayi=CATI, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Vaka Beklenen ve Hesaplanan"
    g.add_data(Reference(ws, min_col=10, min_row=5, max_col=11, max_row=9), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=6, max_row=9))
    g.height, g.width = 8, 14
    ws.add_chart(g, "A15")

    genislik(ws, {"A": 12, "B": 32, "C": 40, "D": 12, "E": 14, "F": 10, "G": 12, "H": 28})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", "0F2742", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "KANIT RAPORU — Tarımsal Destek/Teşvik Uygunluk", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:F3")
    h(ws, 4, 1, "Tarih", yazi=GRİ)
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 5, 1, "İşletme", yazi=GRİ)
    h(ws, 5, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Çağrı programı", yazi=GRİ)
    h(ws, 6, 2, "=GIRDI!B12", yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Hesap yılı", yazi=GRİ)
    h(ws, 7, 2, "=hesapYili", sayi=CATI, yazi=DIS_REF_YEŞIL)

    h(ws, 9, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 9, 2, "=tdu_kararMetni", kalin=True, boyut=14, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 2, "=tdu_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B10:F10")
    h(ws, 11, 1, "Uygunluk puanı", yazi=GRİ)
    h(ws, 11, 2, "=IFERROR(MOTOR!G35,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Tahmini destek", yazi=GRİ)
    h(ws, 12, 2, "=IFERROR(tdu_tahminiDestek,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "Madde atıfları", yazi=GRİ)
    h(ws, 13, 2, "=tdu_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B13:F13")
    h(ws, 14, 1, "Kanıt özeti", yazi=GRİ)
    h(ws, 14, 2, "=tdu_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B14:F14")
    h(ws, 15, 1, "Vaka durumu", yazi=GRİ)
    h(ws, 15, 2, "=tdu_vakaDurum", yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Parmak izi", yazi=GRİ)
    h(ws, 16, 2, "=tdu_parmakIzi", yazi=DIS_REF_YEŞIL)

    h(ws, 18, 1, "Uyarılar", kalin=True, yazi=KRITIK)
    h(ws, 19, 1,
      "Bu çıktı karar destek niteliğindedir; resmî başvuru sonucu yerine geçmez. "
      "ÇKS güncelleme penceresi tartışmalıdır (yorum A/B).",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A19:F20")

    h(ws, 22, 1, "İmza — Giren", yazi=GRİ)
    h(ws, 22, 3, "İmza — Karar veren", yazi=GRİ)
    h(ws, 22, 5, "İmza — Kanıt isteyen", yazi=GRİ)
    h(ws, 24, 1, "________________", yazi=GRİ)
    h(ws, 24, 3, "________________", yazi=GRİ)
    h(ws, 24, 5, "________________", yazi=GRİ)

    baski_hazirla(ws, "A1:F26", f"{URUN_AD} · KANIT")
    genislik(ws, {"A": 22, "B": 40, "C": 22, "D": 16, "E": 22, "F": 16})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim paneli — Destek Uygunluk", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpiler = [
        (2, "Puan", "=IFERROR(MOTOR!G35,0)", CATI),
        (3, "Eşik", "=IFERROR(MOTOR!G15,0)", CATI),
        (4, "Destek", "=IFERROR(tdu_tahminiDestek,0)", TL),
        (5, "Dekar", "=IFERROR(MOTOR!G17,0)", "0.0"),
        (6, "Güven", "=IFERROR(MOTOR!G43,0)", CATI),
        (7, "Risk", "=IFERROR(MOTOR!G44,0)", CATI),
        (8, "ÇKS", "=IFERROR(MOTOR!G23,0)", CATI),
        (9, "Program", "=IFERROR(MOTOR!G31,0)", CATI),
        (10, "Zorunlu %", "=IFERROR(MOTOR!G42,0)", YÜZDE),
        (11, "M-SEN eşik", "=IFERROR(tdu_kuralDegisimSenaryo,0)", CATI),
        (12, "Tahmin bant", "=IFERROR(SENARYO!B32,0)", TL),
        (13, "Vaka", "=IFERROR(tdu_vakaDurum,\"-\")", None),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=9, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=tdu_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=tdu_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 10, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Uygunluk puanı",
         "=IFERROR(MOTOR!G35,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Puan ",TEXT(C11,"0")," — eşik ",TEXT(MOTOR!G15,"0"))'),
        ("T2", "Destek / dekar",
         "=IFERROR(IF(MOTOR!G17=0,0,tdu_tahminiDestek/MOTOR!G17),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Birim destek ",TEXT(C12,"₺ #,##0")," / dekar")'),
        ("O1", "Senaryo puan sapması",
         "=IFERROR(IF(COUNT(SENARYO!C6:C9)<2,0,STDEV.P(SENARYO!C6:C9)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Sapma ",TEXT(C13,"0.0")," puan")'),
        ("O2", "Tornado max etki",
         "=IFERROR(tdu_duyarlilikSira,\"-\")",
         "=IFERROR(tdu_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo bant",
         "=IFERROR(tdu_senaryoKarsilastirma,0)",
         "=IFERROR(tdu_yorumSenaryo,\"-\")"),
        ("O6", "Zorunlu şart oranı",
         "=IFERROR(MOTOR!G42,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Zorunlu tamamlanma ",TEXT(C16,"0.0%"))'),
        ("O8", "Kalite skoru",
         "=IFERROR(KONTROLLER!B12,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalite ",TEXT(C17,"0")," puan")'),
        ("M", "M-SEN eşik",
         "=IFERROR(tdu_kuralDegisimSenaryo,0)",
         "=IFERROR(tdu_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin aralık",
         "=IFERROR(tdu_tahminAralik,0)",
         "=IFERROR(tdu_yorumTahmin,\"-\")"),
        ("I2", "Senaryo puan P90",
         "=IFERROR(IF(COUNT(SENARYO!C6:C9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!C6:C9,tdu_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"P90 puan ",TEXT(C20,"0")," ")'),
    ]
    for i, (kod, ad, form, yorum) in enumerate(moduller):
        r = 11 + i
        h(ws, r, 1, kod, kalin=True, hiza="center", yazi="B08948")
        h(ws, r, 2, ad, kaydir=True)
        h(ws, r, 3, form, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[r].height = 28
        if kod in ("O6",):
            ws.cell(r, 3).number_format = YÜZDE
        elif kod == "T2" or kod == "I1":
            ws.cell(r, 3).number_format = TL
        elif kod != "O2":
            ws.cell(r, 3).number_format = CATI

    h(ws, 23, 1, "Şart", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "Sonuç", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("ÇKS", 1), ("Borçsuz", 1), ("Program", 1), ("Alan", 1), ("Puan", 78),
    ], 24):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=CATI, yazi=GRİ)

    h(ws, 23, 4, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 5, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Tahmini destek", 45000), ("Üst bant", 49500), ("Alt bant", 40500),
        ("Birim mazot", 1000), ("Birim alan", 800),
    ], 24):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Şart Sonuçları"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=28), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=28))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G9")

    g2 = BarChart()
    g2.type = "col"
    g2.title = "Destek Tutar Bandı"
    g2.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "G23")

    g3 = LineChart()
    g3.title = "Destek Çizgisi"
    g3.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g3.height, g3.width = 7, 12
    ws.add_chart(g3, "P9")

    h(ws, 30, 1, "Metrik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 2, "Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Puan", 78), ("Eşik", 70), ("Destek", 45000),
    ], 31):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=CATI if ad != "Destek" else TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "Puan Eşik Destek"
    g4.add_data(Reference(ws, min_col=2, min_row=30, max_row=33), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=1, min_row=31, max_row=33))
    g4.height, g4.width = 7, 10
    ws.add_chart(g4, "P23")

    # pasta — çağrı dağılımı demo
    h(ws, 30, 4, "Program", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 5, "Pay", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, p) in enumerate([
        ("MazotGubre", 40), ("AlanBazli", 25), ("GencCiftci", 15),
        ("Organik", 10), ("Diger", 10),
    ], 31):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, p / 100, sayi=YÜZDE, yazi=GRİ)
    g5 = PieChart()
    g5.title = "Çağrı Program Payı"
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
        ("Negatif dekar?",
         '=IF(GIRDI!B10<0,"NEGATİF","TEMİZ")'),
        ("Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(tdu_yilTaban),N(hesapYili)<=N(tdu_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        ("Borç blok aktif mi?",
         '=IF(IFERROR(MOTOR!G24,0)=1,"BLOK","TEMİZ")'),
        ("Vaka kırık mı?",
         '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")'),
        ("ÇKS şartı?",
         '=IF(IFERROR(MOTOR!G23,0)=1,"TEMİZ","EKSİK")'),
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
    h(ws, 14, 2, "=IFERROR(tdu_girisBeklenen+COUNTA(GIRDI!B6:B18)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Dolu giriş", yazi="333333")
    h(ws, 15, 2, "=COUNTA(GIRDI!B6:B18)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Kalite skoru (modül)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=IFERROR(ROUND(IF(OR(B14=0,B14=\"\"),0,B15/B14*100),0),0)", sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)

    h(ws, 17, 1, "Puan kontrol", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G35,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Destek kontrol", yazi="333333")
    h(ws, 18, 2, "=IFERROR(tdu_tahminiDestek,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Karar kontrol", yazi="333333")
    h(ws, 19, 2, "=IFERROR(tdu_kararMetni,\"-\")", yazi=DIS_REF_YEŞIL)

    genislik(ws, {"A": 40, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek başvuru kütüphanesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:I3")
    sutunlar = ["Senaryo", "Dekar", "Cagri", "Yas", "Cks", "Borc",
                "ProgramSart", "PuanTahmin", "Karar"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "ProgramSart": (
            '=IF(tblOrnek[[#This Row],[Senaryo]]="","",'
            'IFERROR(IF(AND(tblOrnek[[#This Row],[Cks]]="Evet",'
            'tblOrnek[[#This Row],[Borc]]="Hayır",'
            'tblOrnek[[#This Row],[Dekar]]>=tdu_ornekMinDekar),tdu_bir,0),0))'
        ),
        "PuanTahmin": (
            '=IF(tblOrnek[[#This Row],[Senaryo]]="","",'
            'IFERROR(tdu_ornekTaban+IF(tblOrnek[[#This Row],[ProgramSart]]=tdu_bir,tdu_puanProgram,0)'
            '+MIN(tdu_puanAlanMax,ROUND(tblOrnek[[#This Row],[Dekar]]/tdu_puanAlanCarpan,0)),0))'
        ),
        "Karar": (
            '=IF(tblOrnek[[#This Row],[Senaryo]]="","",'
            'IFERROR(IF(tblOrnek[[#This Row],[PuanTahmin]]>=tdu_ornekEsik,"UYGUN · BAŞVUR",'
            'IF(tblOrnek[[#This Row],[PuanTahmin]]>=tdu_ornekEsik-tdu_sartliBand,'
            '"ŞARTLI · BELGE EKLE","UYGUN DEĞİL · DURDUR")),"UYGUN DEĞİL · DURDUR"))'
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:I15", sutunlar, formuller)
    ornekler = [
        ("Mazot uygun", 45, "MazotGubre", 38, "Evet", "Hayır"),
        ("Mazot yetersiz alan", 8, "MazotGubre", 42, "Evet", "Hayır"),
        ("Genç çiftçi", 20, "GencCiftci", 32, "Evet", "Hayır"),
        ("Borçlu blok", 50, "AlanBazli", 40, "Evet", "Evet"),
        ("ÇKS yok", 30, "MazotGubre", 35, "Hayır", "Hayır"),
        ("Organik", 12, "Organik", 45, "Evet", "Hayır"),
    ]
    for i, row in enumerate(ornekler, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v, zemin=GIRIS_SARI, yazi="1F4E79")
            ws.cell(i, k).protection = Protection(locked=False)
        if isinstance(row[1], (int, float)):
            ws.cell(i, 2).number_format = "0.0"
        ws.cell(i, 4).number_format = CATI

    g = BarChart()
    g.type = "col"
    g.title = "Örnek Puan Tahmini"
    h(ws, 5, 11, "PuanDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, p in enumerate([85, 40, 78, 20, 20, 70], 6):
        h(ws, i, 11, p, sayi=CATI, yazi=GRİ)
    g.add_data(Reference(ws, min_col=11, min_row=5, max_row=11), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=6, max_row=11))
    g.height, g.width = 8, 12
    ws.add_chart(g, "A18")
    genislik(ws, {"A": 22, "B": 10, "C": 14, "D": 8, "E": 10, "F": 10, "G": 12, "H": 12, "I": 14})


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
    h(ws, 5, 7, "Urun", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, u in enumerate(["Bitkisel", "Hayvansal", "Karma"], 6):
        h(ws, i, 7, u)
    h(ws, 5, 9, "Cagri", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, c in enumerate(
        ["MazotGubre", "AlanBazli", "GencCiftci", "Organik", "Hayvancilik", "IPARD"], 6
    ):
        h(ws, i, 9, c)
    genislik(ws, {"A": 10, "C": 10, "E": 12, "G": 12, "I": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Parametreler (tek kaynak)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler değiştirilebilir; formül hücreleri kilitlidir.", yazi=GRİ, kaydir=True)
    sutunlar = ["anahtar", "deger", "birim", "aciklama", "kaynak", "yururluk_tarihi", "dogrulama_tarihi", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    params = [
        ("raporTarihi", RAPOR_TARIH, "tarih", "Rapor tarihi", "Üretim", "01.01.2025", "11.08.2026"),
        ("tdu_tolerans", 0.01, "oran", "Vaka tutarlılık toleransı", "Uygulama notu", "01.01.2025", "11.08.2026"),
        ("tdu_sartliBand", 15, "puan", "Şartlı puan bandı", "İç politika", "01.01.2025", "11.08.2026"),
        ("tdu_farkEsik", 50, "puan", "PANO uyarı eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("tdu_senaryoIyi", 1.15, "çarpan", "İyimser dekar çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("tdu_senaryoKotu", 0.85, "çarpan", "Kötümser dekar çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("tdu_oranArtisPuan", 5, "puan", "M-SEN eşik artışı", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("tdu_tahminAlt", 0.9, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("tdu_tahminUst", 1.1, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("tdu_tornadoDekar", 1.2, "çarpan", "Tornado dekar şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("tdu_tornadoYas", 1.0, "çarpan", "Tornado yaş şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("tdu_yuzdelikOran", 0.9, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "11.08.2026"),
        ("tdu_parmakIzi", "TDU-PRO-1.0.0", "metin", "Dosya parmak izi", "Üretim", "01.01.2025", "11.08.2026"),
        ("tdu_girisBeklenen", 13, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("tdu_riskDusuk", 20, "puan", "Düşük risk puanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("tdu_riskOrta", 55, "puan", "Orta risk puanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("tdu_riskYuksek", 85, "puan", "Yüksek risk puanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("tdu_kararEsik", 5, "puan", "Yedek karar eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("tdu_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("tdu_bir", 1, "bayrak", "Mantıksal bir", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("tdu_puanCks", 20, "puan", "ÇKS şart puanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("tdu_puanBorc", 20, "puan", "Borçsuzluk puanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("tdu_puanProgram", 30, "puan", "Program şart puanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("tdu_puanAlanMax", 25, "puan", "Alan puan tavanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("tdu_puanAlanCarpan", 10, "çarpan", "Alan puan çarpanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("tdu_puanYorumA", 5, "puan", "Yorum A belge puanı", "Belirsizlik", "01.01.2025", "11.08.2026"),
        ("tdu_puanYorumB", 10, "puan", "Yorum B belge puanı", "Belirsizlik", "01.01.2025", "11.08.2026"),
        ("tdu_puanOrganikEk", 5, "puan", "Organik ek puan", "İç politika", "01.01.2025", "11.08.2026"),
        ("tdu_puanHayvanEk", 5, "puan", "Hayvan ek puan", "İç politika", "01.01.2025", "11.08.2026"),
            ("tdu_dekarMinGecerli", 0, "dekar", "Geçerli min dekar", "Doğrulama", "01.01.2025", "11.08.2026"),
        ("tdu_dekarAzami", 100000, "dekar", "Azami dekar tavanı", "Doğrulama", "01.01.2025", "11.08.2026"),
        ("tdu_yasMinGecerli", 0, "yıl", "Geçerli min yaş", "Doğrulama", "01.01.2025", "11.08.2026"),
        ("tdu_yasAzami", 120, "yıl", "Azami yaş tavanı", "Doğrulama", "01.01.2025", "11.08.2026"),
        ("tdu_hayvanAzami", 1000000, "adet", "Azami hayvan", "Doğrulama", "01.01.2025", "11.08.2026"),
        ("tdu_yatirimAzami", 1000000000, "TL", "Azami yatırım", "Doğrulama", "01.01.2025", "11.08.2026"),
        ("vaka_dekar_1", 45, "dekar", "Vaka1 dekar", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_dekar_2", 8, "dekar", "Vaka2 dekar", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_yas_1", 38, "yıl", "Vaka3 yaş", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_yatirim_1", 600000, "TL", "Vaka4 yatırım", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("tdu_birimMazot", 1000, "TL", "Mazot birim tutar", "Destek tebliği", "01.01.2025", "11.08.2026"),
        ("tdu_birimAlan", 800, "TL", "Alan bazlı birim", "Destek tebliği", "01.01.2025", "11.08.2026"),
        ("tdu_birimGenc", 1200, "TL", "Genç çiftçi birim", "Destek tebliği", "01.01.2025", "11.08.2026"),
        ("tdu_birimOrganik", 1500, "TL", "Organik birim", "Destek tebliği", "01.01.2025", "11.08.2026"),
        ("tdu_birimHayvan", 2500, "TL", "Hayvan birim", "Destek tebliği", "01.01.2025", "11.08.2026"),
        ("tdu_birimIpard", 0, "TL", "IPARD gösterge birim", "IPARD rehberi", "01.01.2025", "11.08.2026"),
        ("tdu_ornekTaban", 40, "puan", "Örnek puan tabanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("tdu_ornekMinDekar", 10, "dekar", "Örnek min dekar", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("tdu_ornekEsik", 70, "puan", "Örnek karar eşiği", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv", "01.01.2025", "11.08.2026"),
    ]
    for i, row in enumerate(params, 6):
        for k, v in enumerate(row, 1):
            sayi = None
            if k == 2:
                if row[2] == "tarih":
                    sayi = TARİH
                elif row[2] in ("oran",):
                    sayi = YÜZDE
                elif row[2] == "çarpan":
                    sayi = "0.00"
                elif row[2] == "TL":
                    sayi = TL
                elif row[2] == "dekar":
                    sayi = "0.0"
                elif row[2] in ("yıl", "adet", "puan", "bayrak"):
                    sayi = CATI
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k == 2 and sayi is None and isinstance(v, (int, float)) and not isinstance(v, bool):
                ws.cell(i, k).number_format = CATI
            if k == 2:
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"B{i}", row[0], row[3])
        h(ws, i, 8,
          f'=IF(A{i}="","",IF(B{i}="","EKSİK","TAM"))',
          yazi=DIS_REF_YEŞIL, hiza="center")

    son = 5 + len(params)
    genislik(ws, {"A": 24, "B": 18, "C": 10, "D": 28, "E": 18, "F": 14, "G": 14, "H": 10})
    h(ws, son + 2, 4, f"Sürüm {SURUM} | Şifre: {SIFRE}", yazi=GRİ, boyut=9)
    return {p[0]: 6 + i for i, p in enumerate(params)}


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kullanım kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    metinler = [
        "GIRDI sarı hücrelere işletme ve çağrı bilgilerini girin.",
        "KURALLAR sayfasında yıllı eşikleri güncelleyin; MOTOR sabit eşik kullanmaz.",
        "PANO kararını okuyun; ŞARTLI ise eksik belge listesini tamamlayın.",
        "KANIT_RAPORU'nu PDF/yazıcıya gönderin.",
        "Şifre 1234 — formül sayfaları korumalıdır; giriş hücreleri açıktır.",
    ]
    for i, m in enumerate(metinler, 5):
        h(ws, i, 1, f"{i-4}. {m}", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=12)

    h(ws, 11, 1, "Belirsizlik beyanları", kalin=True, boyut=12, yazi=KRITIK)
    h(ws, 12, 1,
      "ÇKS güncelleme penceresi tartışmalıdır. Yorum A: başvuru tarihinde güncel "
      "ÇKS kaydı zorunlu (sıkı). Yorum B: çağrı ilan tarihindeki kayıt kabul "
      "(geniş). Belirsiz durumlarda her iki yorumu SENARYO ile karşılaştırın.",
      kaydir=True, yazi="333333")
    ws.merge_cells("A12:L14")
    h(ws, 16, 1, "Sorumluluk reddi", kalin=True, yazi=GRİ)
    h(ws, 17, 1,
      "Ürün karar destek aracıdır; Tarım ve Orman Bakanlığı veya IPARD resmî "
      "sonucu yerine geçmez. Parametreleri güncel tebliğle doğrulayın.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A17:L18")
    genislik(ws, {"A": 80})


def kosullu_bicim(wb):
    yesil = Font(color="006100", bold=True)
    kirmizi = Font(color="9C0006", bold=True)
    amber = Font(color="9C5700", bold=True)
    y_fill = PatternFill("solid", fgColor="C6EFCE")
    k_fill = PatternFill("solid", fgColor="FFC7CE")
    a_fill = PatternFill("solid", fgColor="FFEB9C")

    def ekle_metin(ws, hucre, metinler):
        for metin, font, fill in metinler:
            ws.conditional_formatting.add(hucre, FormulaRule(
                formula=[f'ISNUMBER(SEARCH({metin},{hucre}))'], font=font, fill=fill))

    karar_metin = [
        ('"UYGUN"', yesil, y_fill),
        ('"ŞARTLI"', amber, a_fill),
        ('"DURDUR"', kirmizi, k_fill),
        ('"VERİ YOK"', kirmizi, k_fill),
    ]
    for sayfa, hucre in [("PANO", "B6"), ("KANIT_RAPORU", "B9"), ("SENARYO", "E6"),
                         ("SENARYO", "E7"), ("SENARYO", "E8"), ("SENARYO", "E9"),
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
    ws.conditional_formatting.add("B13", CellIsRule(
        operator="lessThan", formula=["18"], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B16", CellIsRule(
        operator="lessThan", formula=["0"], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B20", CellIsRule(
        operator="lessThan", formula=["70"], fill=a_fill))
    ws.conditional_formatting.add("B20", CellIsRule(
        operator="greaterThanOrEqual", formula=["85"], font=yesil, fill=y_fill))

    ws = wb["PANO"]
    ws.conditional_formatting.add("B4", CellIsRule(
        operator="lessThan", formula=["70"], fill=a_fill))
    ws.conditional_formatting.add("C4", CellIsRule(
        operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("D4", CellIsRule(
        operator="equal", formula=["0"], fill=a_fill))
    ws.conditional_formatting.add("G4", CellIsRule(
        operator="greaterThan", formula=["70"], font=kirmizi, fill=k_fill))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("G35", CellIsRule(
        operator="lessThan", formula=["70"], fill=a_fill))
    ws.conditional_formatting.add("G35", CellIsRule(
        operator="greaterThanOrEqual", formula=["70"], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("G40", CellIsRule(
        operator="equal", formula=["0"], fill=a_fill))
    ws.conditional_formatting.add("G24", CellIsRule(
        operator="equal", formula=["1"], font=kirmizi, fill=k_fill))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("C6:C9", CellIsRule(
        operator="lessThan", formula=["70"], fill=a_fill))
    ws.conditional_formatting.add("C6:C9", CellIsRule(
        operator="greaterThanOrEqual", formula=["70"], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B16:B18", CellIsRule(
        operator="greaterThan", formula=["10"], font=kirmizi))

    ws = wb["AYARLAR"]
    ws.conditional_formatting.add("H6:H60", FormulaRule(
        formula=['H6="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("H6:H60", FormulaRule(
        formula=['H6="TAM"'], font=yesil, fill=y_fill))


def ad_tanimlari(wb, ayar_satir=None):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$6")
    ad_ekle(wb, "tdu_tahminiDestek", "MOTOR!$G$40")
    ad_ekle(wb, "tdu_kararMetni", "MOTOR!$G$37")
    ad_ekle(wb, "tdu_kararGerekce", "MOTOR!$G$38")
    ad_ekle(wb, "tdu_maddeAtifMetni", "MOTOR!$G$56")
    ad_ekle(wb, "tdu_kanitRaporu", "MOTOR!$G$57")
    ad_ekle(wb, "tdu_kuralYilMatris", "KURALLAR!$B$20")
    ad_ekle(wb, "tdu_parmakIzi", "AYARLAR!$B$18")

    ad_ekle(wb, "tdu_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "tdu_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "tdu_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "tdu_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "tdu_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "tdu_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "tdu_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "tdu_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "tdu_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "tdu_vakaFark", "VAKALAR!$B$13")

    # AYARLAR satırları: params sırası (raporTarihi=6 ... )
    # 6 raporTarihi, 7 tolerans ... 18 parmakIzi, 19 girisBeklenen ...
    if ayar_satir is None:
        ayar_satir = {
            "tdu_tolerans": 7, "tdu_sartliBand": 8, "tdu_farkEsik": 9,
            "tdu_senaryoIyi": 10, "tdu_senaryoKotu": 11, "tdu_oranArtisPuan": 12,
            "tdu_tahminAlt": 13, "tdu_tahminUst": 14, "tdu_tornadoDekar": 15,
            "tdu_tornadoYas": 16, "tdu_yuzdelikOran": 17, "tdu_parmakIzi": 18,
            "tdu_girisBeklenen": 19, "tdu_riskDusuk": 20, "tdu_riskOrta": 21,
            "tdu_riskYuksek": 22, "tdu_kararEsik": 23, "tdu_yilTaban": 24,
            "tdu_bir": 25, "tdu_puanCks": 26, "tdu_puanBorc": 27,
            "tdu_puanProgram": 28, "tdu_puanAlanMax": 29, "tdu_puanAlanCarpan": 30,
            "tdu_puanYorumA": 31, "tdu_puanYorumB": 32, "tdu_puanOrganikEk": 33,
            "tdu_puanHayvanEk": 34,             "tdu_dekarMinGecerli": 35, "tdu_dekarAzami": 36,
            "tdu_yasMinGecerli": 37, "tdu_yasAzami": 38,
            "tdu_hayvanAzami": 39, "tdu_yatirimAzami": 40,
            "vaka_dekar_1": 41, "vaka_dekar_2": 42, "vaka_yas_1": 43,
            "vaka_yatirim_1": 44, "tdu_birimMazot": 45, "tdu_birimAlan": 46,
            "tdu_birimGenc": 47, "tdu_birimOrganik": 48, "tdu_birimHayvan": 49,
            "tdu_birimIpard": 50, "tdu_ornekTaban": 51,
        }
    for ad, satir in ayar_satir.items():
        if ad == "raporTarihi":
            continue
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "tdu_modulT1": (11, "C"), "tdu_modulT1Yorum": (11, "D"),
        "tdu_modulT2": (12, "C"), "tdu_modulT2Yorum": (12, "D"),
        "tdu_modulO1": (13, "C"), "tdu_modulO1Yorum": (13, "D"),
        "tdu_modulO6": (16, "C"), "tdu_modulO6Yorum": (16, "D"),
        "tdu_modulO8": (17, "C"), "tdu_modulO8Yorum": (17, "D"),
        "tdu_modulI2": (20, "C"), "tdu_modulI2Yorum": (20, "D"),
    }
    for ad, (satir, kol) in modul_map.items():
        ad_ekle(wb, ad, f"PANO!${kol}${satir}")

    ad_ekle(wb, "ListeYillar", "LISTELER!$A$6:$A$8")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$C$6:$C$7")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$6:$E$7")
    ad_ekle(wb, "ListeUrun", "LISTELER!$G$6:$G$8")
    ad_ekle(wb, "ListeCagri", "LISTELER!$I$6:$I$11")


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
    varsayilan = os.path.join(kok, "TarimsalDestekUygunluk.xlsx")
    hedef = cikti_yolu or varsayilan
    os.makedirs(os.path.dirname(os.path.abspath(hedef)) or ".", exist_ok=True)
    wb.save(hedef)

    hsh = hashlib.sha256()
    with open(hedef, "rb") as f:
        for b in iter(lambda: f.read(65536), b""):
            hsh.update(b)
    dig = hsh.hexdigest()

    wb2 = __import__("openpyxl").load_workbook(hedef)
    # parmak izi satırı
    for row in wb2["AYARLAR"].iter_rows(min_row=6, max_row=60, max_col=1):
        if row[0].value == "tdu_parmakIzi":
            wb2["AYARLAR"].cell(row[0].row, 2).value = dig[:16]
            break
    for wsx in wb2.worksheets:
        sayfa_koru(wsx)
    wb2.save(hedef)

    csv_yol = os.path.join(kok, "ornek_veri.csv")
    with open(csv_yol, "w", encoding="utf-8") as f:
        f.write("Senaryo,Dekar,Cagri,Yas,Cks,Borc\n")
        f.write("Mazot uygun,45,MazotGubre,38,Evet,Hayır\n")
        f.write("Mazot yetersiz alan,8,MazotGubre,42,Evet,Hayır\n")
        f.write("Genç çiftçi,20,GencCiftci,32,Evet,Hayır\n")
        f.write("Borçlu blok,50,AlanBazli,40,Evet,Evet\n")

    print(f"Dosya: {hedef}")
    print(f"SHA-256: {dig}")
    print(f"Şifre: {SIFRE}")
    return hedef


if __name__ == "__main__":
    yol = sys.argv[1] if len(sys.argv) > 1 else None
    uretilen = main(yol)
    cikti = os.path.join(KOK, "cikti", "TarimsalDestekUygunluk.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
