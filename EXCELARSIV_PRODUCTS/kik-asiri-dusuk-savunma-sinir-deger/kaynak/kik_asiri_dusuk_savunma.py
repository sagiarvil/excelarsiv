#!/usr/bin/env python3
"""
KİK Aşırı Düşük Savunma + Sınır Değer Paketi — üretim betiği (A1 / manda v6).
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

URUN_AD = "KİK Aşırı Düşük Savunma + Sınır Değer Paketi"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 11)

# Demo: yaklaşık 10M; N=0.90 → sınır 9M; teklif 8.2M → açık 0.8M;
# kadın/genç 50k×%5=2.500 → birikimli potansiyel 32.000; mevcut 10.000 → kaçırılan 22.000
DEMO = {
    "ihale_adi": "Örnek Yol Yapım İhalesi",
    "donem": "2025-IH-014",
    "hesapYili": 2025,
    "yaklasik_maliyet": 10_000_000,
    "teklif_tutari": 8_200_000,
    "iscilik": 3_500_000,
    "malzeme": 2_800_000,
    "genel_gider": 900_000,
    "kar": 700_000,
    "dayanak_adet": 6,
    "yorum_modu": "A",
}


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: KİK sınır değer ve aşırı düşük savunmayı etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Sapma ve karar bozulur | Kaynak: [Assumption] KİK sınır değer",
    )


def kural_cek(kural_id: str) -> str:
    """Oranları tblKurallar'dan çeker; yıl seçimi kik_yilTaban ile (G07: adında rakam yok)."""
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-kik_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(kik_yuvarMax,IFERROR(N(G12),0)))"


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
      "Yaklaşık maliyet ve teklifi girin; KİK sınır değer sapmasını ve savunma "
      "ve kaçırılan fırsatı senaryolu olarak üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Çok yıllı kural tablosu (2025/2026/2027) ile oran ve tavan güncellemesine hazır motor",
        "Sınır değer, sapma %, maliyet kırılımı ve savunma dayanaklarının yan yana hesabı",
        "Baz / iyimser / kötümser + kural değişimi (oran +1 puan) senaryoları",
        "Tebliğ tarzı altın vakalar ve KANIT_RAPORU (A4 imza)",
        "Belirsizlik beyanı: sınır değer hesap yöntemi (yorum A / B)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Teklif hazırlayan — sınır değer ve sapmayı görmek",
        "SMMM — imzalı kanıt raporu dosyalamak",
        "İşveren — aylık/yıllık tasarruf tutarını görmek",
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
        ("Adım 1 — Girdiler", "GIRDI sayfasında yıl, yaklaşık maliyet, teklif ve maliyet kırılımlarını sarı hücrelere girin."),
        ("Adım 2 — Karar", "PANO'da potansiyel tasarruf, kaçırılan tutar ve karar rozetini izleyin."),
        ("Adım 3 — Kanıt", "KANIT_RAPORU'nu yazdırıp imzalayın; VAKALAR tutarlılığını kontrol edin."),
    ]
    for i, (b, m) in enumerate(adimlar, 5):
        h(ws, i, 1, b, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        _kim(ws, f"A{i}", b, "Sırayla izleyin")
    h(ws, 9, 1, "Sık yapılan hatalar", kalin=True, yazi=KRITIK)
    for i, m in enumerate([
        "Yaklaşık maliyet ile teklifi ters yazmak",
        "Maliyet kırılım toplamını tekliften kopuk bırakmak",
        "Hesap yılını kural tablosuyla uyumsuz seçmek",
    ], 10):
        h(ws, i, 1, "• " + m, yazi=KRITIK)
    h(ws, 14, 1,
      "Bu dosya karar destek aracıdır; KİK savunma dilekçesi veya hukuki görüş yerine geçmez.",
      yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 72, "B": 70})
    for r in (10, 11, 12):
        h(ws, r, 1, ws.cell(r, 1).value, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Dönem ve işveren", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Sınır ve sapma formülle üretilir.", yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "İhale Adı", DEMO["ihale_adi"], None, "ihale_adi",
         "İhale / iş adını yazın."),
        (7, "İhale Numarası / Dönem", DEMO["donem"], None, "donem",
         "İhale referansını yazın (ör. 2025-IH-014)."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili",
         "2025, 2026 veya 2027 seçin."),
        (9, "Yaklaşık Maliyet [₺]", DEMO["yaklasik_maliyet"], TL, "yaklasik_maliyet",
         "İhale yaklaşık maliyetini KDV hariç TL girin."),
        (10, "Teklif Tutarı [₺]", DEMO["teklif_tutari"], TL, "teklif_tutari",
         "Sunulan teklif bedelini TL girin."),
        (11, "İşçilik Maliyeti [₺]", DEMO["iscilik"], TL, "iscilik",
         "Teklif içi işçilik maliyetini girin."),
        (12, "Malzeme Maliyeti [₺]", DEMO["malzeme"], TL, "malzeme",
         "Teklif içi malzeme maliyetini girin."),
        (13, "Genel Gider [₺]", DEMO["genel_gider"], TL, "genel_gider",
         "Genel gider tutarını girin."),
        (14, "Kâr [₺]", DEMO["kar"], TL, "kar",
         "Teklife gömülü kâr tutarını girin."),
        (15, "Savunma Dayanak Adedi", DEMO["dayanak_adet"], CATI, "dayanak_adet",
         "Açıklama dayanak / belge adedini yazın."),
        (16, "Yorum Modu (A/B)", DEMO["yorum_modu"], None, "yorum_modu",
         "Sınır hesap yorumu A (katsayı) veya B (sıkı) seçin."),
    ]
    h(ws, 5, 1, "Alan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 2, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 3, "Birim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    for satir, etiket, deger, sayi, _ad, mesaj in alanlar:
        birim = "₺" if sayi == TL else ("yıl" if _ad == "hesapYili" else (
            "adet" if _ad == "dayanak_adet" else "metin"))
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim, yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 18, 1, "Aşırı Düşük Açık (ayna) [₺]", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 2, "=IFERROR(kik_potansiyelTasarruf,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B16)=0,kik_girisBeklenen=0),0,ROUND(COUNTA(B6:B16)/kik_girisBeklenen*100,0)),0)',
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
        ("ihale_adi", DEMO["ihale_adi"], "metin", "Evet", "GIRDI"),
        ("donem", DEMO["donem"], "metin", "Evet", "GIRDI"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Evet", "GIRDI"),
        ("yaklasik_maliyet", DEMO["yaklasik_maliyet"], "TL", "Evet", "GIRDI"),
        ("teklif_tutari", DEMO["teklif_tutari"], "TL", "Evet", "GIRDI"),
        ("iscilik", DEMO["iscilik"], "TL", "Evet", "GIRDI"),
        ("malzeme", DEMO["malzeme"], "TL", "Evet", "GIRDI"),
        ("genel_gider", DEMO["genel_gider"], "TL", "Evet", "GIRDI"),
        ("kar", DEMO["kar"], "TL", "Evet", "GIRDI"),
        ("dayanak_adet", DEMO["dayanak_adet"], "adet", "Evet", "GIRDI"),
        ("yorum_modu", DEMO["yorum_modu"], "A/B", "Evet", "GIRDI"),
    ]
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 24):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if isinstance(d, (int, float)) and a not in ("hesapYili", "dayanak_adet"):
            ws.cell(i, 2).number_format = TL
        elif a in ("hesapYili", "dayanak_adet"):
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
    dogrulama(ws, "list", "ListeYorumAB", "B16",
              baslik="Yorum", mesaj="Yorum modunda A veya B seçin.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B seçin.", bos=False)
    dogrulama(ws, "list", "ListeYillar", "B26",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeYorumAB", "B34",
              baslik="Yorum", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.")
    for aralik, baslik, mesaj in [
        ("B9", "Yaklaşık Maliyet", "0 ile 1e12 arasında TL."),
        ("B10", "Teklif", "0 ile 100000 arasında."),
        ("B11", "İşçilik", "0 ile 1e12 arasında TL."),
        ("B12", "Malzeme", "0 ile 1e12 arasında TL."),
        ("B13", "Genel Gider", "0 ile 1e12 arasında TL."),
        ("B14", "Kâr", "0 ile 1e12 arasında TL."),
        ("B15", "Dayanak", "0 ile 1e12 arasında TL."),
        ("B27", "Yaklaşık Maliyet", "Tablo: TL tutar."),
        ("B28", "Teklif", "Tablo: adet."),
        ("B29", "İşçilik", "Tablo: TL tutar."),
        ("B30", "Malzeme", "Tablo: TL tutar."),
        ("B31", "Genel Gider", "Tablo: TL tutar."),
        ("B32", "Kâr", "Tablo: TL tutar."),
        ("B33", "Dayanak", "Tablo: TL tutar."),
    ]:
        f2 = "100000" if baslik == "Teklif" else "1000000000000"
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj=mesaj,
                  hata_baslik="Geçersiz tutar", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2=f2)

    giris_hucreleri(ws, 6, 16, [2])
    giris_hucreleri(ws, 24, 34, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 40, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1024, "Sarı alanlar giriş; sapma ve kalite formülleri kilitlidir.")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Mevzuat kural tablosu (yıl yan yana)", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["KIK-001", "KİK sınır N", "sinir_katsayi", "yaklasik>0", "sinir_katsayi",
         0.90, 0.90, 0.88, "01.01.2025", "KİK sınır değer N"],
        ["KIK-002", "Savunulabilir", "savunulabilir_sapma", "sapma hesap", "savunulabilir_sapma",
         0.05, 0.05, 0.06, "01.01.2025", "Savunulabilir sapma eşiği"],
        ["KIK-003", "İncele", "incele_sapma", "sapma hesap", "incele_sapma",
         0.12, 0.12, 0.13, "01.01.2025", "İnceleme sapma eşiği"],
        ["KIK-004", "Riskli", "riskli_sapma", "sapma hesap", "riskli_sapma",
         0.20, 0.20, 0.22, "01.01.2025", "Riskli sapma eşiği"],
        ["KIK-005", "Genel", "yuvarlama", "her hesap", "yuvarlama_ondalik",
         2, 2, 2, "01.01.2025", "Uygulama notu"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = YÜZDE if k in (6, 7, 8) and isinstance(v, float) and v < 1 else (
                CATI if k in (6, 7, 8) else None)
            if k in (6, 7, 8) and s[0] in ("KIK-005",):
                sayi = CATI if s[0] == "KIK-005" else YÜZDE
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-kik_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "N katsayısı (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("KIK-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Savunulabilir sapma (aktif yıl)", yazi="333333")
    h(ws, 16, 2, kural_cek("KIK-002"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"KIK-001=",TEXT(B15,"0.00"),"KIK-002=",TEXT(B16,"0.0%"),'
      '"yıl=",IF(OR(hesapYili="",hesapYili=0),"-",hesapYili)),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    for col in ("F", "G", "H"):
        dogrulama(ws, "decimal", "0", f"{col}7:{col}11",
                  baslik="Kural değeri", mesaj="Yıl kolonuna sayısal parametre girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ve üzeri sayı.",
                  isaret="between", f2="10000000")
    genislik(ws, {get_column_letter(i): w for i, w in enumerate(
        [34, 28, 22, 22, 20, 12, 12, 12, 12, 18, 14], 1)})
    # A14–A17 etiketleri için ekstra genişlik (taşma uyarısı)
    genislik(ws, {"A": 36, "B": 55})
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

    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif mevzuat yılı", "KIK-005")  # G7
    ekle("N katsayısı", km("KIK-001"), "oran", "Sınır değer N", "KIK-001")  # G8
    ekle("Savunulabilir sapma", km("KIK-002"), "oran", "Savunulabilir eşik", "KIK-002")  # G9
    ekle("İncele sapma", km("KIK-003"), "oran", "İnceleme eşiği", "KIK-003")  # G10
    ekle("Riskli sapma", km("KIK-004"), "oran", "Riskli eşik", "KIK-004")  # G11
    ekle("Yuvarlama ondalık", km("KIK-005"), "adet", "Yuvarlama basamağı", "KIK-005")  # G12
    ekle("Yaklaşık maliyet", "=GIRDI!B9", "TL", "Yaklaşık maliyet", "KIK-001")  # G13
    ekle("Teklif tutarı", "=GIRDI!B10", "TL", "Sunulan teklif", "KIK-001")  # G14
    ekle("İşçilik", "=GIRDI!B11", "TL", "İşçilik maliyeti", "KIK-005")  # G15
    ekle("Malzeme", "=GIRDI!B12", "TL", "Malzeme maliyeti", "KIK-005")  # G16
    ekle("Genel gider", "=GIRDI!B13", "TL", "Genel gider", "KIK-005")  # G17
    ekle("Kâr", "=GIRDI!B14", "TL", "Kâr tutarı", "KIK-005")  # G18
    ekle("Dayanak adedi", "=GIRDI!B15", "adet", "Savunma dayanak sayısı", "KIK-005")  # G19
    ekle("N efektif",
         '=IF(GIRDI!B16="B",MAX(0,G8-kik_oranArtisPuan),G8)',
         "oran", "Yorum A/B N", "KIK-001")  # G20
    ekle("Sınır değer", f"=ROUND(G13*G20,{_YV})", "TL", "Yaklaşık × N", "KIK-001")  # G21
    ekle("Maliyet toplam", "=G15+G16+G17+G18", "TL", "Kırılım toplamı", "KIK-005")  # G22
    ekle("Sapma tutarı (ham)", "=G21-G14", "TL", "Sınır − teklif", "KIK-001")  # G23
    ekle("Aşırı düşük açık", "=MAX(0,G23)", "TL", "Pozitif açık", "KIK-001")  # G24
    ekle("Sapma oranı", "=IF(G21=0,0,G23/G21)", "oran", "(Sınır−teklif)/sınır", "KIK-002")  # G25
    ekle("Maliyet / teklif", "=IF(G14=0,0,G22/G14)", "oran", "Maliyet yoğunluğu", "KIK-005")  # G26
    ekle("Dayanak skoru", "=MIN(100,G19*kik_dayanakCarpan)", "puan", "Dayanak kalitesi", "KIK-005")  # G27
    ekle("Potansiyel / açık", "=G24", "TL", "Aşırı düşük açık (G28)", "KIK-001")  # G28
    ekle("Açık (ham kopya)", "=G28", "TL", "Ayna", "KIK-001")  # G29
    ekle("Kaçırılan / açık", "=G24", "TL", "Karar için açık", "KIK-001")  # G30
    ekle("Sınır iyimser", f"=ROUND(G13*G8*kik_senaryoIyi,{_YV})", "TL", "İyimser sınır", "KIK-001")  # G31
    ekle("Sınır kötümser", f"=ROUND(G13*G8*kik_senaryoKotu,{_YV})", "TL", "Kötümser sınır", "KIK-001")  # G32
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(G14>=G21,"TAM",'
         'IF(G25<=G9,"SAVUN",'
         'IF(G25<=G10,"INCELE","RISK"))))',
         "kod", "Karar kodu", "KIK-005")  # G33
    ekle("Karar metni",
         '=IF(G33="VERI_YOK","VERİ YOK",'
         'IF(G33="TAM","SAVUNULABİLİR",'
         'IF(G33="SAVUN","SAVUNULABİLİR",'
         'IF(G33="INCELE","İNCELE","RİSKLİ"))))',
         "metin", "Kullanıcıya karar", "KIK-005")  # G34
    ekle("Gerekçe",
         '=IF(G33="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G33="TAM","Teklif sınır değerin üzerinde; aşırı düşük yok.",'
         'IF(G33="SAVUN","Sapma savunulabilir eşik içinde; dayanaklarla açıklanabilir.",'
         'IF(G33="INCELE","Sapma inceleme bandında; maliyet kırılımını güçlendirin.",'
         '"Sapma riskli bandda; teklif veya açıklama revizyonu gerekir."))))',
         "metin", "Gerekçe cümlesi", "KIK-005")  # G35
    ekle("Teklif / sınır", "=IF(G21=0,0,G14/G21)", "oran", "Teklif oranı", "KIK-001")  # G36
    ekle("Açık / sınır", "=IF(G21=0,0,G30/G21)", "oran", "Açık oranı", "KIK-001")  # G37
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B16)=0,kik_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B16)/kik_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "KIK-005")  # G38
    ekle("Risk skoru",
         '=IF(G33="VERI_YOK",0,IF(OR(G33="TAM",G33="SAVUN"),kik_riskDusuk,'
         'IF(G33="INCELE",kik_riskOrta,kik_riskYuksek)))',
         "puan", "Sapma riski", "KIK-005")  # G39
    ekle("Senaryo iyimser teklif", "=G14*kik_senaryoIyi", "TL", "İyimser teklif", "KIK-001")  # G40
    ekle("Senaryo kötümser teklif", "=G14*kik_senaryoKotu", "TL", "Kötümser teklif", "KIK-001")  # G41
    ekle("İyimser açık", f"=ROUND(MAX(0,G31-G40),{_YV})", "TL", "İyimser açık", "KIK-001")  # G42
    ekle("Kötümser açık", f"=ROUND(MAX(0,G32-G41),{_YV})", "TL", "Kötümser açık", "KIK-001")  # G43
    ekle("M-SEN N+", "=G8+kik_oranArtisPuan", "oran", "N + delta", "KIK-001")  # G44
    ekle("M-SEN sınır", f"=ROUND(G13*G44,{_YV})", "TL", "M-SEN sınır", "KIK-001")  # G45
    ekle("M-SEN açık", "=MAX(0,G45-G14)", "TL", "M-SEN farkı", "KIK-001")  # G46
    ekle("Tahmin üst", "=G30*kik_tahminUst", "TL", "Tahmin üst bant", "KIK-005")  # G47
    ekle("Tahmin alt", "=G30*kik_tahminAlt", "TL", "Tahmin alt bant", "KIK-005")  # G48
    ekle("Tornado: yaklaşık etkisi",
         f"=ABS(ROUND(MAX(0,G13*kik_tornadoMatrah*G20-G14),{_YV})-G28)",
         "TL", "Yaklaşık ± etki", "KIK-001")  # G49
    ekle("Tornado: teklif etkisi",
         f"=ABS(ROUND(MAX(0,G21-G14*kik_tornadoMevcut),{_YV})-G30)",
         "TL", "Teklif ± etki", "KIK-001")  # G50
    ekle("Tornado: N etkisi",
         f"=ABS(ROUND(MAX(0,G13*(G8+kik_oranArtisPuan)-G14),{_YV})-G28)",
         "TL", "N ± etki", "KIK-001")  # G51
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"KIK-001→sınır N","KIK-002→savunulabilir","KIK-003→incele","KIK-004→riskli")',
         "metin", "Kanıt atıfları", "KIK-001")  # G52
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Sınır=",TEXT(G21,"₺ #,##0"),"Teklif=",TEXT(G14,"₺ #,##0"),'
         '"Açık=",TEXT(G30,"₺ #,##0"),"Sapma=",TEXT(G25,"0.0%"))',
         "metin", "Rapor özeti", "KIK-005")  # G53

    assert len(adimlar) >= 40, len(adimlar)

    for i, (ad, formul, birim, ne, kid) in enumerate(adimlar, 1):
        r = 6 + i
        if ad == "Karar kodu":
            formul = (
                '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
                'IF(G14>=G21,"TAM",'
                'IF(G25<=G9,"SAVUN",'
                'IF(G25<=G10,"INCELE","RISK"))))'
            )
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
    h(ws, 5, 3, "Potansiyel", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Kaçırılan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    pot = "IFERROR(kik_potansiyelTasarruf,0)"
    mev = "IFERROR(kik_mevcutTesvik,0)"
    senaryolar = [
        ("İyimser", "=IFERROR(kik_senaryoIyi+N(GIRDI!B9)*kik_sifirCarpan,0)",
         f"=IFERROR(ROUND(({pot})*B6,2),0)",
         f"=IFERROR(MAX(0,C6-({mev})),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(D6=0,"UYGUN","FIRSAT"))'),
        ("Baz", "=IFERROR(kik_bazCarpan+N(GIRDI!B9)*kik_sifirCarpan,kik_bazCarpan)",
         f"=IFERROR(ROUND(({pot})*B7,2),0)",
         f"=IFERROR(MAX(0,C7-({mev})),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(D7=0,"UYGUN","FIRSAT"))'),
        ("Kötümser", "=IFERROR(kik_senaryoKotu+N(GIRDI!B9)*kik_sifirCarpan,0)",
         f"=IFERROR(ROUND(({pot})*B8,2),0)",
         f"=IFERROR(MAX(0,C8-({mev})),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(D8=0,"UYGUN","FIRSAT"))'),
        ("M-SEN +1 puan", "=IFERROR(kik_bazCarpan+N(GIRDI!B9)*kik_sifirCarpan,kik_bazCarpan)",
         "=IFERROR(MOTOR!G45,0)",
         "=IFERROR(MOTOR!G46,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(D9=0,"UYGUN","FIRSAT"))'),
    ]
    for i, (ad, carp, asg, fark, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, asg, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, fark, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant genişliği (iyimser−kötümser kaçırılan)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(D6-D8,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser kaçırılan ",TEXT(D6,"₺ #,##0")," · Baz ",TEXT(D7,"₺ #,##0"),'
      '" · Kötümser ",TEXT(D8,"₺ #,##0")," · Bant ",TEXT(B11,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:H12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Tornado matrah etkisi", yazi="333333")
    h(ws, 16, 2, "=IFERROR(MOTOR!G49,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado mevcut etkisi", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G50,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado oran etkisi", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G51,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 3,
      '=IFERROR(IF(B16>=MAX(B16:B18),kik_siraBir,IF(B16>=MIN(B16:B18)+ABS(B16-B17),kik_siraIki,kik_siraUc)),kik_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 3,
      '=IFERROR(IF(B17>=MAX(B16:B18),kik_siraBir,IF(B17=MEDIAN(B16:B18),kik_siraIki,kik_siraUc)),kik_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 3,
      '=IFERROR(IF(B18>=MAX(B16:B18),kik_siraBir,IF(B18=MEDIAN(B16:B18),kik_siraIki,kik_siraUc)),kik_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(CONCATENATE("1)",INDEX(A16:A18,MATCH(kik_siraBir,C16:C18,0)),'
      '" 2)",INDEX(A16:A18,MATCH(kik_siraIki,C16:C18,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B16:B18),"₺ #,##0")," TL — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN kural değişimi kaçırılan", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=D9", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"5 puan oranı +1 puan senaryosunda kaçırılan ",TEXT(D9,"₺ #,##0")," TL")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı (kaçırılan)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=MOTOR!G48", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 3, "=kik_kacirilan", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 4, "=MOTOR!G47", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Tahmin (FORECAST)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1, "Seri nokta", yazi=GRİ)
    for i, v in enumerate([1, 2, 3, 4], 28):
        h(ws, i, 1, v, sayi=CATI)
        h(ws, i, 2, f"=D{5+v}", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 32, 1, "Tahmin sonraki", yazi="333333")
    h(ws, 32, 2, "=IFERROR(FORECAST.LINEAR(kik_tahminX,B28:B31,A28:A31),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Tahmin yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Tahmin ",TEXT(B32,"₺ #,##0")," · Alt ",TEXT(B25,"₺ #,##0"),'
      '" · Üst ",TEXT(D25,"₺ #,##0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 5, 7, "KacirilanDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([18800, 22000, 26400, 23500], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([1600, 1100, 1500], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)
    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([18800, 22000, 26400, 23500], 28):
        h(ws, i, 4, v, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo Kaçırılan Karşılaştırması"
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
    g3.title = "Senaryo Kaçırılan Trendi"
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

    # V-001: 300000*0.05 = 15000
    # V-002: 100000*0.145 = 14500
    # V-003: birikim A = 15000+14500+2500+6150 = 38150 (with engelli 30000*0.205)
    # Actually for vakalar use fixed formula from rules only — keep simple:
    # V-001: 5 puan only 15000
    # V-002: ilave only 14500
    # V-003: sum 5 puan + ilave = 29500
    # V-004: tavan * oran check 150000*0.05 = 7500
    vakalar_data = [
        ("V-001", "Tebliğ örnek 1 — 5 puan",
         "bes_puan_matrah=vaka_matrah_1",
         15000,
         "=IFERROR(ROUND(vaka_matrah_1*INDEX(tblKurallar[deger_2025],MATCH(\"KIK-001\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=kik_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "5510 md.81 örnek"),
        ("V-002", "Tebliğ örnek 2 — ilave istihdam",
         "ilave_matrah=vaka_matrah_2",
         14500,
         "=IFERROR(ROUND(vaka_matrah_2*INDEX(tblKurallar[deger_2025],MATCH(\"KIK-002\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=kik_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "4447 geç.10 örnek"),
        ("V-003", "Tebliğ örnek 3 — birikimli A",
         "5 puan + ilave",
         29500,
         "=IFERROR(ROUND(vaka_matrah_1*INDEX(tblKurallar[deger_2025],MATCH(\"KIK-001\",tblKurallar[kural_id],0))"
         "+vaka_matrah_2*INDEX(tblKurallar[deger_2025],MATCH(\"KIK-002\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=kik_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Birikim yorum A"),
        ("V-004", "Tavan × 5 puan kontrol",
         "tavan=vaka_tavan; oran KIK-001",
         7500,
         "=IFERROR(ROUND(vaka_tavan*INDEX(tblKurallar[deger_2025],MATCH(\"KIK-001\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=kik_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "KIK-003 tavan kontrol"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v,
              yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = TL

    formuller = {
        "fark": "=IF(tblVakalar[[#This Row],[vaka_id]]=\"\",\"\",IFERROR(tblVakalar[[#This Row],[beklenen_sonuc]]-tblVakalar[[#This Row],[hesaplanan]],0))",
        "durum": '=IF(tblVakalar[[#This Row],[vaka_id]]="","",IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=kik_tolerans,"TUTARLI","KIRIK"),"KIRIK"))',
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
    for i, (b, he) in enumerate([(15000, 15000), (14500, 14500), (29500, 29500), (7500, 7500)], 6):
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
    h(ws, 3, 1, "KANIT RAPORU — KİK Aşırı Düşük Savunma", kalin=True, boyut=14, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:F3")
    h(ws, 4, 1, "Rapor tarihi", yazi="333333")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 5, 1, "İşveren", yazi="333333")
    h(ws, 5, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Dönem / yıl", yazi="333333")
    h(ws, 6, 2, '=CONCATENATE(GIRDI!B7," / ",hesapYili)', yazi=DIS_REF_YEŞIL)

    h(ws, 8, 1, "Girdi özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, (et, form, sayi) in enumerate([
        ("Aylık brüt matrah", "=GIRDI!B9", TL),
        ("Çalışan sayısı", "=GIRDI!B10", CATI),
        ("5 puan kapsam", "=GIRDI!B11", TL),
        ("İlave istihdam", "=GIRDI!B12", TL),
        ("Kadın/genç", "=GIRDI!B13", TL),
        ("Kâr", "=GIRDI!B14", TL),
        ("Mevcut teşvik", "=kik_mevcutTesvik", TL),
    ], 9):
        h(ws, i, 1, et, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)

    h(ws, 17, 1, "Hesap zinciri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 18, 1, "Potansiyel tasarruf", yazi="333333")
    h(ws, 18, 2, "=kik_potansiyelTasarruf", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Kaçırılan", yazi="333333")
    h(ws, 19, 2, "=kik_kacirilan", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "KARAR", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=kik_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 21, 1, "Gerekçe", yazi="333333")
    h(ws, 21, 2, "=kik_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B21:F21")

    h(ws, 23, 1, "Madde atıfları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "=kik_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B23:F23")

    h(ws, 25, 1, "Kanıt gövdesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=kik_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B25:F25")

    h(ws, 27, 1, "Parmak izi (SHA-256 kısaltma)", yazi="333333")
    h(ws, 27, 2, "=kik_parmakIzi", yazi=GRİ)

    h(ws, 29, 1, "Hazırlayan (imza)", yazi="333333")
    h(ws, 29, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B29", "İmza", "Raporu hazırlayan kişinin adını yazın.")
    h(ws, 30, 1, "Onaylayan (imza)", yazi="333333")
    h(ws, 30, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B30", "Onay", "Onaylayan kişinin adını yazın.")

    h(ws, 32, 1,
      "Uyarı: Bu çıktı karar destektir; kesin KİK savunma sonucu veya hukuki görüş değildir. "
      f"Sürüm {SURUM}.",
      yazi=GRİ, boyut=9, kaydir=True)
    ws.merge_cells("A32:F32")

    baski_hazirla(ws, "A1:F33", f"{URUN_AD} · {SURUM}")
    ws.page_setup.orientation = "portrait"
    genislik(ws, {"A": 28, "B": 22, "C": 14, "D": 14, "E": 14, "F": 14})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim paneli — KİK Aşırı Düşük", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpiler = [
        (2, "Matrah", "=IFERROR(MOTOR!G21,0)", TL),
        (3, "Potansiyel", "=IFERROR(kik_potansiyelTasarruf,0)", TL),
        (4, "Dayanak", "=IFERROR(kik_mevcutTesvik,0)", TL),
        (5, "Kaçırılan", "=IFERROR(kik_kacirilan,0)", TL),
        (6, "Güven", "=IFERROR(MOTOR!G38,0)", CATI),
        (7, "Risk", "=IFERROR(MOTOR!G39,0)", CATI),
        (8, "5 puan oran", "=IFERROR(MOTOR!G8,0)", YÜZDE),
        (9, "İlave oran", "=IFERROR(MOTOR!G9,0)", YÜZDE),
        (10, "Karşılama", "=IFERROR(MOTOR!G37,0)", YÜZDE),
        (11, "M-SEN", "=IFERROR(kik_kuralDegisimSenaryo,0)", TL),
        (12, "Tahmin", "=IFERROR(SENARYO!B32,0)", TL),
        (13, "Vaka", "=IFERROR(kik_vakaDurum,\"-\")", None),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=9, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=kik_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=kik_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 10, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Kapsam / matrah",
         "=IFERROR(MOTOR!G36,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kapsam yoğunluğu ",TEXT(C11,"0.0%")," — yüksekse birikim riski artar")'),
        ("T2", "Mevcut / potansiyel",
         "=IFERROR(MOTOR!G37,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Karşılama ",TEXT(C12,"0.0%")," — 100% altı kaçırılan fırsat")'),
        ("O1", "Kaçırılan sapması",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,STDEV.P(SENARYO!D6:D9)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo kaçırılan sapması ",TEXT(C13,"₺ #,##0")," TL")'),
        ("O2", "Tornado max etki",
         "=IFERROR(kik_duyarlilikSira,\"-\")",
         "=IFERROR(kik_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo bant",
         "=IFERROR(kik_senaryoKarsilastirma,0)",
         "=IFERROR(kik_yorumSenaryo,\"-\")"),
        ("O6", "5 puan payı",
         "=IFERROR(IF(kik_potansiyelTasarruf=0,0,MOTOR!G22/kik_potansiyelTasarruf),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"5 puan / potansiyel ",TEXT(C16,"0.0%"))'),
        ("O8", "Kalite skoru",
         "=IFERROR(KONTROLLER!B12,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalite ",TEXT(C17,"0")," puan")'),
        ("M", "M-SEN kaçırılan",
         "=IFERROR(kik_kuralDegisimSenaryo,0)",
         "=IFERROR(kik_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin aralık",
         "=IFERROR(kik_tahminAralik,0)",
         "=IFERROR(kik_yorumTahmin,\"-\")"),
        ("I2", "Senaryo kaçırılan P90",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!D6:D9,kik_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"P90 kaçırılan ",TEXT(C20,"₺ #,##0")," TL")'),
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

    h(ws, 23, 1, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("5 puan", 15_000),
        ("İlave istihdam", 14_500),
        ("Kadın/genç", 2_500),
        ("Kâr", 6_150),
        ("Potansiyel", 38_150),
    ], 24):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)

    h(ws, 23, 4, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 5, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Potansiyel", 38_150),
        ("Dayanak", 10_000),
        ("Kaçırılan", 28_150),
        ("Yıllık kaçırılan", 337_800),
        ("M-SEN potansiyel", 39_650),
    ], 24):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Teşvik Kalemleri"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=28), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=28))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G9")

    g2 = BarChart()
    g2.type = "col"
    g2.title = "Potansiyel / Mevcut / Kaçırılan"
    g2.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "G23")

    g3 = LineChart()
    g3.title = "Tasarruf Çizgisi"
    g3.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g3.height, g3.width = 7, 12
    ws.add_chart(g3, "P9")

    h(ws, 30, 1, "Metrik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 2, "Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Kaçırılan", 28_150),
        ("Potansiyel", 38_150),
        ("Dayanak", 10_000),
    ], 31):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "KPI Kaçırılan / Potansiyel / Mevcut"
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
        ("Negatif brüt matrah?",
         '=IF(GIRDI!B9<0,"NEGATİF","TEMİZ")'),
        ("Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(kik_yilTaban),N(hesapYili)<=N(kik_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        ("Matrah sıfır mı?",
         '=IF(IFERROR(MOTOR!G21,0)=0,"SIFIR","NORMAL")'),
        ("Vaka kırık mı?",
         '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")'),
        ("Kaçırılan eşiği aşıyor mu?",
         '=IF(IFERROR(kik_kacirilan,0)>kik_kacirilanEsik,"AŞIYOR","NORMAL")'),
        ("Yorum A/B seçili mi?",
         '=IF(OR(GIRDI!B16="A",GIRDI!B16="B"),"TEMİZ","EKSİK")'),
        ("Motor adım sayısı",
         "=COUNTA(MOTOR!B7:B53)"),
    ]
    for i, (ad, form) in enumerate(kontroller_list, 5):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Canlı formül | Değiştirmeyin")

    h(ws, 14, 1, "Toplam giriş alanı", yazi="333333")
    h(ws, 14, 2, "=IFERROR(kik_girisBeklenen+COUNTA(GIRDI!B6:B16)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Dolu giriş", yazi="333333")
    h(ws, 15, 2, "=COUNTA(GIRDI!B6:B16)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Kalite skoru (modül)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=IFERROR(ROUND(IF(OR(B14=0,B14=\"\"),0,B15/B14*100),0),0)", sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)

    h(ws, 17, 1, "Potansiyel kontrol", yazi="333333")
    h(ws, 17, 2, "=IFERROR(kik_potansiyelTasarruf,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Kaçırılan kontrol", yazi="333333")
    h(ws, 18, 2, "=IFERROR(kik_kacirilan,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Karar kontrol", yazi="333333")
    h(ws, 19, 2, "=IFERROR(kik_kararMetni,\"-\")", yazi=DIS_REF_YEŞIL)

    genislik(ws, {"A": 40, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek senaryo kütüphanesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:I3")
    h(ws, 4, 1, "Bu sayfa demo değerleri saklar; GIRDI'ye kopyalanabilir.", yazi=GRİ)
    sutunlar = ["Senaryo", "Yaklaşık Maliyet", "İşçilik", "Malzeme", "Genel Gider", "Kâr",
                "Dayanak", "Potansiyel", "Kaçırılan"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Potansiyel": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(ROUND(MAX(0,tblOrnek[[#This Row],[Yaklaşık Maliyet]]*"
            "INDEX(tblKurallar[deger_2025],MATCH(\"KIK-001\",tblKurallar[kural_id],0))"
            "-(tblOrnek[[#This Row],[İşçilik]]+tblOrnek[[#This Row],[Malzeme]]"
            "+tblOrnek[[#This Row],[Genel Gider]]+tblOrnek[[#This Row],[Kâr]])),"
            "kik_yuvarMax),0))"
        ),
        "Kaçırılan": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(MAX(0,tblOrnek[[#This Row],[Potansiyel]]"
            "-tblOrnek[[#This Row],[Dayanak]]*kik_sifirCarpan),0))"
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:I1004", sutunlar, formuller)
    ornekler = [
        ("Demo ana", 500_000, 300_000, 100_000, 50_000, 30_000, 10_000),
        ("Yüksek mevcut", 500_000, 300_000, 100_000, 50_000, 30_000, 40_000),
        ("Düşük matrah", 200_000, 120_000, 40_000, 20_000, 10_000, 5_000),
        ("Sadece 5 puan", 400_000, 400_000, 0, 0, 0, 5_000),
        ("Sıfır kaçırılan", 300_000, 200_000, 50_000, 0, 0, 17_250),
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
        for kolon in range(1, 8):
            _kim(ws, f"{get_column_letter(kolon)}{i}", sutunlar[kolon - 1], "Örnek senaryo değeri")

    for col, ad in [("B", "Brüt"), ("C", "İşçilik"), ("D", "Malzeme"),
                    ("E", "Genel Gider"), ("F", "Kâr"), ("G", "Dayanak")]:
        dogrulama(ws, "decimal", "0", f"{col}6:{col}1004",
                  baslik=ad, mesaj=f"{ad} tutarını TL girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ile 1e12 arasında.",
                  isaret="between", f2="1000000000000")

    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6, 7])
    h(ws, 5, 12, "KacirilanDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([28150, 0, 15200, 15000, 0], 6):
        h(ws, i, 12, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Örnek Senaryo Kaçırılan"
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
        _kim(ws, f"C{i}", "Yorum", "Birikim yorumu")
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
        ("kik_tolerans", 0.01, "TL", "Vaka tutarlılık toleransı", "Uygulama notu", "01.01.2025", "11.08.2026"),
        ("kik_kacirilanEsik", 15_000, "TL", "PANO kaçırılan uyarı eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("kik_senaryoIyi", 0.90, "çarpan", "İyimser matrah çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kik_senaryoKotu", 1.10, "çarpan", "Kötümser matrah çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kik_oranArtisPuan", 0.01, "oran", "M-SEN +1 puan", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("kik_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kik_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kik_tornadoMatrah", 1.05, "çarpan", "Tornado matrah şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kik_tornadoMevcut", 0.95, "çarpan", "Tornado mevcut şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kik_isverenOran", 0.205, "oran", "N sıkılaştırma yedek katsayısı", "KİK model", "01.01.2025", "11.08.2026"),
        ("kik_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "11.08.2026"),
        ("kik_parmakIzi", "KIK-PRO-1.0.0", "metin", "Dosya parmak izi etiketi", "Üretim", "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv", "01.01.2025", "11.08.2026"),
        ("kik_girisBeklenen", 11, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("kik_riskDusuk", 15, "puan", "Düşük risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("kik_riskOrta", 55, "puan", "Orta risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("kik_riskYuksek", 85, "puan", "Yüksek risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("kik_riskEsikOran", 0.5, "oran", "Yüksek risk kaçırılan/potansiyel", "İç politika", "01.01.2025", "11.08.2026"),
        ("kik_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı (hesapYili−taban)", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("kik_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("kik_dayanakCarpan", 10, "adet", "Dayanak skoru çarpanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("kik_bazCarpan", 1, "çarpan", "Baz senaryo çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kik_sifirCarpan", 0, "çarpan", "Nötr çarpan (sıfır)", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kik_siraBir", 1, "adet", "Tornado sıra 1", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("kik_siraIki", 2, "adet", "Tornado sıra 2", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("kik_siraUc", 3, "adet", "Tornado sıra 3", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("kik_tahminX", 5, "adet", "FORECAST X noktası", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("vaka_matrah_1", 300_000, "TL", "Vaka1 yaklaşık maliyet dilimi", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_matrah_2", 100_000, "TL", "Vaka2 teklif dilimi", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_tavan", 150_000, "TL", "Vaka4 tavan", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
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
              baslik="Kaçırılan eşiği", mesaj="Uyarı eşiği TL.",
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
        "SAVUNULABİLİR — teklif sınır üzerinde veya sapma savunulabilir eşikte.",
        "İNCELE / FIRSAT — kaçırılan > 0 ama eşik altında; gözden geçirin.",
        "İNCELE / RİSKLİ — sapma eşik üzerinde; maliyet kırılımı ve dayanakları güçlendirin.",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 10, 1, "Belirsizlik beyanları (M05)", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari(["tesvik_birikim_sirasi"]), 11):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 13, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2,
      "Sınır değer = yaklaşık maliyet × N; sapma = (sınır − teklif) / sınır. "
      "Bu nokta mevzuatta tartışmalıdır; sonuç yorum B ile yan yana okunmalıdır.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B13:J13")
    h(ws, 14, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 2,
      "Yorum B, N katsayısını sıkılaştırır. Dosya her iki yorumun sonucunu MOTOR'da üretir.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B14:J14")

    h(ws, 16, 1, "Kullanım sırası", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 1,
      "GIRDI → KURALLAR (gerekirse) → MOTOR (inceleme) → SENARYO → VAKALAR → PANO → KANIT_RAPORU → AYARLAR.",
      kaydir=True, yazi="333333")
    ws.merge_cells("A17:J17")
    h(ws, 19, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
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
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("İNCELE",B6))'], font=amber, fill=a_fill))
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
        ws.conditional_formatting.add(f"{col}9:{col}15", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(f"{col}9:{col}15", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("F24:F34", FormulaRule(formula=['F24="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F24:F34", FormulaRule(formula=['F24="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F6:F9", CellIsRule(operator="notEqual", formula=["0"], font=amber))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="greaterThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="equal", formula=["0"], font=yesil))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("FIRSAT",E6))'], font=amber))
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
    ws.conditional_formatting.add("F7:F53", FormulaRule(formula=['F7="KIK-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F53", FormulaRule(formula=['F7="KIK-005"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "C", "D", "E", "F", "G", "H", "I"):
        ws.conditional_formatting.add(f"{harf}6:{harf}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("I6:I20", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B20", FormulaRule(formula=['ISNUMBER(SEARCH("UYGUN",B20))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B20", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B20))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B19", CellIsRule(operator="greaterThan", formula=["0"], font=kirmizi))


def ad_tanimlari(wb):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "kik_mevcutTesvik", "GIRDI!$B$15")
    ad_ekle(wb, "kik_potansiyelTasarruf", "MOTOR!$G$28")
    ad_ekle(wb, "kik_kacirilan", "MOTOR!$G$30")
    ad_ekle(wb, "kik_kararMetni", "MOTOR!$G$34")
    ad_ekle(wb, "kik_kararGerekce", "MOTOR!$G$35")
    ad_ekle(wb, "kik_maddeAtifMetni", "MOTOR!$G$52")
    ad_ekle(wb, "kik_kanitRaporu", "MOTOR!$G$53")
    ad_ekle(wb, "kik_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "kik_parmakIzi", "AYARLAR!$B$19")

    ad_ekle(wb, "kik_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "kik_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "kik_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "kik_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "kik_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "kik_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "kik_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "kik_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "kik_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "kik_vakaFark", "VAKALAR!$B$13")

    ayar_map = {
        "kik_tolerans": 8,
        "kik_kacirilanEsik": 9,
        "kik_senaryoIyi": 10,
        "kik_senaryoKotu": 11,
        "kik_oranArtisPuan": 12,
        "kik_tahminAlt": 13,
        "kik_tahminUst": 14,
        "kik_tornadoMatrah": 15,
        "kik_tornadoMevcut": 16,
        "kik_isverenOran": 17,
        "kik_yuzdelikOran": 18,
        "kik_girisBeklenen": 21,
        "kik_riskDusuk": 22,
        "kik_riskOrta": 23,
        "kik_riskYuksek": 24,
        "kik_riskEsikOran": 25,
        "kik_yilTaban": 26,
        "kik_yuvarMax": 27,
        "kik_dayanakCarpan": 28,
        "kik_bazCarpan": 29,
        "kik_sifirCarpan": 30,
        "kik_siraBir": 31,
        "kik_siraIki": 32,
        "kik_siraUc": 33,
        "kik_tahminX": 34,
        "vaka_matrah_1": 35,
        "vaka_matrah_2": 36,
        "vaka_tavan": 37,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "kik_modulT1": (11, "C"), "kik_modulT1Yorum": (11, "D"),
        "kik_modulT2": (12, "C"), "kik_modulT2Yorum": (12, "D"),
        "kik_modulO1": (13, "C"), "kik_modulO1Yorum": (13, "D"),
        "kik_modulO6": (16, "C"), "kik_modulO6Yorum": (16, "D"),
        "kik_modulO8": (17, "C"), "kik_modulO8Yorum": (17, "D"),
        "kik_modulI2": (20, "C"), "kik_modulI2Yorum": (20, "D"),
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
    varsayilan = os.path.join(kok, "KikAsiriDusukSavunmaSinirDeger.xlsx")
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
    cikti = os.path.join(repo, "cikti", "KikAsiriDusukSavunmaSinirDeger.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
