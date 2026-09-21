#!/usr/bin/env python3
"""
İhale Fiyat Farkı (Eskalasyon) Pro — üretim betiği (A1 / manda v6).
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

URUN_AD = "İhale Fiyat Farkı (Eskalasyon) Pro"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 11)

# Demo: baz 100 → güncel 118; sabit %10; değişken %90 × 1.18
# Pn = 0.10 + 0.90×1.18 = 1.162; hakediş 2.5M → fiyat farkı 405.000 TL → TALEP ET
DEMO = {
    "ihale_adi": "Örnek Yol Yapım İhalesi",
    "donem": "2025-HK-03",
    "hesapYili": 2025,
    "sozlesme_bedeli": 10_000_000,
    "baz_endeks": 100.0,
    "guncel_endeks": 118.0,
    "agirlik_iscilik": 0.40,
    "agirlik_malzeme": 0.45,
    "agirlik_akaryakit": 0.15,
    "hakedis_tutari": 2_500_000,
    "yorum_modu": "A",
}


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: Fiyat farkı / eskalasyon kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Fiyat farkı ve karar bozulur | Kaynak: [Assumption] KİK fiyat farkı cetveli",
    )


def kural_cek(kural_id: str) -> str:
    """Oranları tblKurallar'dan çeker; yıl seçimi iff_yilTaban ile (G07: adında rakam yok)."""
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-iff_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(iff_yuvarMax,IFERROR(N(G12),0)))"


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
      "Sözleşme bedeli, endeks ve kalem ağırlıklarını girin; hakediş fiyat farkını "
      "hesaplayıp TALEP ET / BEKLE / VAZGEÇ kararını üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Çok yıllı kural tablosu (2025/2026/2027) ile eşik ve sabit pay güncellemesine hazır motor",
        "Endeks oranı, kalem ağırlıkları ve hakediş üzerinden fiyat farkı tutarı",
        "Baz / iyimser / kötümser + kural değişimi (eşik +1 puan) senaryoları",
        "Tebliğ tarzı altın vakalar ve KANIT_RAPORU (A4 imza)",
        "Belirsizlik beyanı: endeks seçimi (yorum A / B)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Maliyet kontrol — eskalasyon tutarını görmek",
        "Müteahhit — TALEP ET / BEKLE / VAZGEÇ kararını almak",
        "Danışman — imzalı kanıt raporunu dosyalamak",
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
        ("Adım 1 — Girdiler", "GIRDI sayfasında yıl, sözleşme, endeks, ağırlık ve hakedişi sarı hücrelere girin."),
        ("Adım 2 — Karar", "PANO'da fiyat farkı tutarını ve TALEP ET / BEKLE / VAZGEÇ rozetini izleyin."),
        ("Adım 3 — Kanıt", "KANIT_RAPORU'nu yazdırıp imzalayın; VAKALAR tutarlılığını kontrol edin."),
    ]
    for i, (b, m) in enumerate(adimlar, 5):
        h(ws, i, 1, b, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        _kim(ws, f"A{i}", b, "Sırayla izleyin")
    h(ws, 9, 1, "Sık yapılan hatalar", kalin=True, yazi=KRITIK)
    for i, m in enumerate([
        "Baz ve güncel endeksi ters yazmak",
        "Ağırlık toplamını 1'den sapık bırakmak (işçilik+malzeme+akaryakıt)",
        "Hesap yılını kural tablosuyla uyumsuz seçmek",
    ], 10):
        h(ws, i, 1, "• " + m, yazi=KRITIK)
    h(ws, 14, 1,
      "Bu dosya karar destek aracıdır; fiyat farkı talebi dilekçesi veya hukuki görüş yerine geçmez.",
      yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 72, "B": 70})
    for r in (10, 11, 12):
        h(ws, r, 1, ws.cell(r, 1).value, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Dönem ve sözleşme", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Fiyat farkı formülle üretilir.", yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "İhale / İş Adı", DEMO["ihale_adi"], None, "ihale_adi",
         "İhale veya iş adını yazın."),
        (7, "Hakediş Dönemi", DEMO["donem"], None, "donem",
         "Hakediş dönem referansını yazın (ör. 2025-HK-03)."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili",
         "2025, 2026 veya 2027 seçin."),
        (9, "Sözleşme Bedeli [₺]", DEMO["sozlesme_bedeli"], TL, "sozlesme_bedeli",
         "Sözleşme bedelini KDV hariç TL girin."),
        (10, "Baz Endeks", DEMO["baz_endeks"], CATI, "baz_endeks",
         "Sözleşme / ihale tarihindeki baz endeksi girin."),
        (11, "Güncel Endeks", DEMO["guncel_endeks"], CATI, "guncel_endeks",
         "Hakediş dönemindeki güncel endeksi girin."),
        (12, "İşçilik Ağırlığı [%]", DEMO["agirlik_iscilik"], YÜZDE, "agirlik_iscilik",
         "İşçilik kalem ağırlığını 0-1 veya yüzde olarak girin."),
        (13, "Malzeme Ağırlığı [%]", DEMO["agirlik_malzeme"], YÜZDE, "agirlik_malzeme",
         "Malzeme kalem ağırlığını girin."),
        (14, "Akaryakıt Ağırlığı [%]", DEMO["agirlik_akaryakit"], YÜZDE, "agirlik_akaryakit",
         "Akaryakıt kalem ağırlığını girin."),
        (15, "Hakediş Tutarı [₺]", DEMO["hakedis_tutari"], TL, "hakedis_tutari",
         "Dönem hakediş tutarını TL girin."),
        (16, "Yorum Modu (A/B)", DEMO["yorum_modu"], None, "yorum_modu",
         "Endeks yorumu A (standart) veya B (sıkı sabit pay) seçin."),
    ]
    h(ws, 5, 1, "Alan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 2, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 3, "Birim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    for satir, etiket, deger, sayi, _ad, mesaj in alanlar:
        birim = "₺" if sayi == TL else ("yıl" if _ad == "hesapYili" else (
            "oran" if sayi == YÜZDE else ("endeks" if "endeks" in _ad else "metin")))
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim, yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 18, 1, "Fiyat Farkı (ayna) [₺]", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 2, "=IFERROR(iff_fiyatFarki,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B16)=0,iff_girisBeklenen=0),0,ROUND(COUNTA(B6:B16)/iff_girisBeklenen*100,0)),0)',
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
        ("sozlesme_bedeli", DEMO["sozlesme_bedeli"], "TL", "Evet", "GIRDI"),
        ("baz_endeks", DEMO["baz_endeks"], "endeks", "Evet", "GIRDI"),
        ("guncel_endeks", DEMO["guncel_endeks"], "endeks", "Evet", "GIRDI"),
        ("agirlik_iscilik", DEMO["agirlik_iscilik"], "oran", "Evet", "GIRDI"),
        ("agirlik_malzeme", DEMO["agirlik_malzeme"], "oran", "Evet", "GIRDI"),
        ("agirlik_akaryakit", DEMO["agirlik_akaryakit"], "oran", "Evet", "GIRDI"),
        ("hakedis_tutari", DEMO["hakedis_tutari"], "TL", "Evet", "GIRDI"),
        ("yorum_modu", DEMO["yorum_modu"], "A/B", "Evet", "GIRDI"),
    ]
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 24):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in ("sozlesme_bedeli", "hakedis_tutari"):
            ws.cell(i, 2).number_format = TL
        elif a in ("agirlik_iscilik", "agirlik_malzeme", "agirlik_akaryakit"):
            ws.cell(i, 2).number_format = YÜZDE
        elif a in ("hesapYili", "baz_endeks", "guncel_endeks"):
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
    for aralik, baslik, mesaj, f2 in [
        ("B9", "Sözleşme Bedeli", "0 ile 1e12 arasında TL.", "1000000000000"),
        ("B10", "Baz Endeks", "0 ile 100000 arasında.", "100000"),
        ("B11", "Güncel Endeks", "0 ile 100000 arasında.", "100000"),
        ("B12", "İşçilik Ağırlığı", "0 ile 1 arasında.", "1"),
        ("B13", "Malzeme Ağırlığı", "0 ile 1 arasında.", "1"),
        ("B14", "Akaryakıt Ağırlığı", "0 ile 1 arasında.", "1"),
        ("B15", "Hakediş", "0 ile 1e12 arasında TL.", "1000000000000"),
        ("B27", "Sözleşme Bedeli", "Tablo: TL tutar.", "1000000000000"),
        ("B28", "Baz Endeks", "Tablo: endeks.", "100000"),
        ("B29", "Güncel Endeks", "Tablo: endeks.", "100000"),
        ("B30", "İşçilik", "Tablo: oran.", "1"),
        ("B31", "Malzeme", "Tablo: oran.", "1"),
        ("B32", "Akaryakıt", "Tablo: oran.", "1"),
        ("B33", "Hakediş", "Tablo: TL tutar.", "1000000000000"),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj=mesaj,
                  hata_baslik="Geçersiz tutar", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2=f2)

    giris_hucreleri(ws, 6, 16, [2])
    giris_hucreleri(ws, 24, 34, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 40, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1024, "Sarı alanlar giriş; fiyat farkı ve kalite formülleri kilitlidir.")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Mevzuat kural tablosu (yıl yan yana)", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["IFF-001", "Sabit pay", "sabit_pay", "Pn hesabı", "sabit_pay",
         0.10, 0.10, 0.12, "01.01.2025", "Eskalasyona tabi olmayan sabit pay"],
        ["IFF-002", "Talep eşiği", "talep_esik", "karar", "talep_esik",
         0.05, 0.05, 0.06, "01.01.2025", "TALEP ET için min fiyat farkı oranı"],
        ["IFF-003", "Bekle eşiği", "bekle_esik", "karar", "bekle_esik",
         0.02, 0.02, 0.025, "01.01.2025", "BEKLE bandı alt eşiği"],
        ["IFF-004", "Ağırlık tolerans", "agirlik_tolerans", "ağırlık kontrol", "agirlik_tolerans",
         0.02, 0.02, 0.02, "01.01.2025", "Ağırlık toplamı 1 sapma toleransı"],
        ["IFF-005", "Yuvarlama", "yuvarlama", "her hesap", "yuvarlama_ondalik",
         2, 2, 2, "01.01.2025", "Uygulama notu"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = YÜZDE if k in (6, 7, 8) and isinstance(v, float) and v < 1 else (
                CATI if k in (6, 7, 8) else None)
            if k in (6, 7, 8) and s[0] in ("IFF-005",):
                sayi = CATI
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-iff_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Sabit pay (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("IFF-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Talep eşiği (aktif yıl)", yazi="333333")
    h(ws, 16, 2, kural_cek("IFF-002"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"IFF-001=",TEXT(B15,"0.00"),"IFF-002=",TEXT(B16,"0.0%"),'
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
    h(ws, 4, 1, "Oranlar tblKurallar'dan çekilir; sabit oran yoktur.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")

    basliklar = [*MOTOR_BASLIK, "deger"]
    baslik_satiri(ws, 6, basliklar)

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    adimlar.clear()
    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif mevzuat yılı", "IFF-005")  # G7
    ekle("Sabit pay", km("IFF-001"), "oran", "Sabit (eskalasyonsuz) pay", "IFF-001")  # G8
    ekle("Talep eşiği", km("IFF-002"), "oran", "TALEP ET eşiği", "IFF-002")  # G9
    ekle("Bekle eşiği", km("IFF-003"), "oran", "BEKLE eşiği", "IFF-003")  # G10
    ekle("Ağırlık tolerans", km("IFF-004"), "oran", "Ağırlık sapma toleransı", "IFF-004")  # G11
    ekle("Yuvarlama ondalık", km("IFF-005"), "adet", "Yuvarlama basamağı", "IFF-005")  # G12
    ekle("Sözleşme bedeli", "=GIRDI!B9", "TL", "Sözleşme bedeli", "IFF-001")  # G13
    ekle("Baz endeks", "=GIRDI!B10", "adet", "Baz endeks", "IFF-001")  # G14
    ekle("Güncel endeks", "=GIRDI!B11", "adet", "Güncel endeks", "IFF-001")  # G15
    ekle("İşçilik ağırlığı", "=GIRDI!B12", "oran", "İşçilik payı", "IFF-004")  # G16
    ekle("Malzeme ağırlığı", "=GIRDI!B13", "oran", "Malzeme payı", "IFF-004")  # G17
    ekle("Akaryakıt ağırlığı", "=GIRDI!B14", "oran", "Akaryakıt payı", "IFF-004")  # G18
    ekle("Hakediş tutarı", "=GIRDI!B15", "TL", "Dönem hakedişi", "IFF-001")  # G19
    ekle("Sabit efektif",
         '=IF(GIRDI!B16="B",MIN(1,G8+iff_oranArtisPuan),G8)',
         "oran", "Yorum A/B sabit pay", "IFF-001")  # G20
    ekle("Endeks oranı", "=IF(G14=0,0,G15/G14)", "oran", "Güncel/baz", "IFF-001")  # G21
    ekle("Ağırlık toplam", "=G16+G17+G18", "oran", "Kalem ağırlık toplamı", "IFF-004")  # G22
    ekle("Ağırlık sapması", "=ABS(G22-1)", "oran", "|toplam−1|", "IFF-004")  # G23
    ekle("Pn katsayısı",
         "=G20+(1-G20)*G21",
         "oran", "Sabit + değişken×endeks", "IFF-001")  # G24
    ekle("Fiyat farkı ham", "=G19*(G24-1)", "TL", "Hakediş×(Pn−1)", "IFF-001")  # G25
    ekle("Fiyat farkı", f"=ROUND(G25,{_YV})", "TL", "Yuvarlanmış fiyat farkı", "IFF-005")  # G26
    ekle("Fiyat farkı oranı", "=IF(G19=0,0,G26/G19)", "oran", "Fark/hakediş", "IFF-002")  # G27
    ekle("Potansiyel / fark", "=G26", "TL", "Fiyat farkı (G28)", "IFF-001")  # G28
    ekle("Açık (ham kopya)", "=G28", "TL", "Ayna", "IFF-001")  # G29
    ekle("Kaçırılan / fark", "=G26", "TL", "Karar için fark", "IFF-001")  # G30
    ekle("Pn iyimser",
         "=G20+(1-G20)*G21*iff_senaryoIyi",
         "oran", "İyimser Pn", "IFF-001")  # G31
    ekle("Pn kötümser",
         "=G20+(1-G20)*G21*iff_senaryoKotu",
         "oran", "Kötümser Pn", "IFF-001")  # G32
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(OR(G19<=0,G14<=0),"VAZGEC",'
         'IF(G27>=G9,"TALEP",'
         'IF(G27>=G10,"BEKLE","VAZGEC"))))',
         "kod", "Karar kodu", "IFF-005")  # G33
    ekle("Karar metni",
         '=IF(G33="VERI_YOK","VERİ YOK",'
         'IF(G33="TALEP","TALEP ET",'
         'IF(G33="BEKLE","BEKLE","VAZGEÇ")))',
         "metin", "Kullanıcıya karar", "IFF-005")  # G34
    ekle("Gerekçe",
         '=IF(G33="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G33="TALEP","Fiyat farkı oranı talep eşiğinin üzerinde; talep dosyası hazırlayın.",'
         'IF(G33="BEKLE","Fark bekle bandında; endeks ve hakediş dönemini izleyin.",'
         '"Fiyat farkı düşük, endeks/hakediş eksik veya olumsuz; talep maliyetine değmeyebilir.")))',
         "metin", "Gerekçe cümlesi", "IFF-005")  # G35
    ekle("Hakediş / sözleşme", "=IF(G13=0,0,G19/G13)", "oran", "Hakediş yoğunluğu", "IFF-001")  # G36
    ekle("Fark / sözleşme", "=IF(G13=0,0,G30/G13)", "oran", "Fark oranı (sözleşme)", "IFF-001")  # G37
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B16)=0,iff_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B16)/iff_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "IFF-005")  # G38
    ekle("Risk skoru",
         '=IF(G33="VERI_YOK",0,IF(G33="TALEP",iff_riskDusuk,'
         'IF(G33="BEKLE",iff_riskOrta,iff_riskYuksek)))',
         "puan", "Talep riski (ters)", "IFF-005")  # G39
    ekle("Senaryo iyimser fark", f"=ROUND(G19*(G31-1),{_YV})", "TL", "İyimser fark", "IFF-001")  # G40
    ekle("Senaryo kötümser fark", f"=ROUND(G19*(G32-1),{_YV})", "TL", "Kötümser fark", "IFF-001")  # G41
    ekle("İyimser açık", "=G40", "TL", "İyimser fark ayna", "IFF-001")  # G42
    ekle("Kötümser açık", "=G41", "TL", "Kötümser fark ayna", "IFF-001")  # G43
    ekle("M-SEN sabit+", "=G8+iff_oranArtisPuan", "oran", "Sabit + delta", "IFF-001")  # G44
    ekle("M-SEN Pn", "=G44+(1-G44)*G21", "oran", "M-SEN Pn", "IFF-001")  # G45
    ekle("M-SEN fark", f"=ROUND(G19*(G45-1),{_YV})", "TL", "M-SEN farkı", "IFF-001")  # G46
    ekle("Tahmin üst", "=G30*iff_tahminUst", "TL", "Tahmin üst bant", "IFF-005")  # G47
    ekle("Tahmin alt", "=G30*iff_tahminAlt", "TL", "Tahmin alt bant", "IFF-005")  # G48
    ekle("Tornado: endeks etkisi",
         f"=ABS(ROUND(G19*((G20+(1-G20)*G21*iff_tornadoMatrah)-1),{_YV})-G28)",
         "TL", "Endeks ± etki", "IFF-001")  # G49
    ekle("Tornado: hakediş etkisi",
         f"=ABS(ROUND(G19*iff_tornadoMevcut*(G24-1),{_YV})-G30)",
         "TL", "Hakediş ± etki", "IFF-001")  # G50
    ekle("Tornado: sabit etkisi",
         f"=ABS(ROUND(G19*(((G8+iff_oranArtisPuan)+(1-(G8+iff_oranArtisPuan))*G21)-1),{_YV})-G28)",
         "TL", "Sabit ± etki", "IFF-001")  # G51
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"IFF-001→sabit pay","IFF-002→talep","IFF-003→bekle","IFF-004→ağırlık")',
         "metin", "Kanıt atıfları", "IFF-001")  # G52
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Pn=",TEXT(G24,"0.000"),"Hakediş=",TEXT(G19,"₺ #,##0"),'
         '"Fark=",TEXT(G30,"₺ #,##0"),"Oran=",TEXT(G27,"0.0%"))',
         "metin", "Rapor özeti", "IFF-005")  # G53

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

    genislik(ws, {"A": 8, "B": 40, "C": 55, "D": 10, "E": 28, "F": 12, "G": 22})
    sabitle(ws, "A7")
    h(ws, 6 + len(adimlar) + 2, 9,
      "Motor adımları kilitlidir; oranlar yalnızca KURALLAR'dan gelir.", yazi=GRİ, kaydir=True)
    return len(adimlar)


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve duyarlılık", kalin=True, boyut=13, yazi=KOYU_LACIVERT)

    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Endeks Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Fiyat Farkı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Net Fark", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(iff_senaryoIyi+N(GIRDI!B9)*iff_sifirCarpan,0)",
         "=IFERROR(MOTOR!G40,0)",
         "=IFERROR(MOTOR!G40,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(D6>=iff_fiyatFarki*iff_tahminAlt,"TALEP ET","BEKLE"))'),
        ("Baz", "=IFERROR(iff_bazCarpan+N(GIRDI!B9)*iff_sifirCarpan,iff_bazCarpan)",
         "=IFERROR(iff_fiyatFarki,0)",
         "=IFERROR(iff_fiyatFarki,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(iff_kararMetni,"-"))'),
        ("Kötümser", "=IFERROR(iff_senaryoKotu+N(GIRDI!B9)*iff_sifirCarpan,0)",
         "=IFERROR(MOTOR!G41,0)",
         "=IFERROR(MOTOR!G41,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(D8<=0,"VAZGEÇ","BEKLE"))'),
        ("M-SEN +1 puan", "=IFERROR(iff_bazCarpan+N(GIRDI!B9)*iff_sifirCarpan,iff_bazCarpan)",
         "=IFERROR(MOTOR!G46,0)",
         "=IFERROR(MOTOR!G46,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(D9=0,"VAZGEÇ","TALEP ET"))'),
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
    h(ws, 16, 1, "Tornado endeks etkisi", yazi="333333")
    h(ws, 16, 2, "=IFERROR(MOTOR!G49,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado hakediş etkisi", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G50,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado sabit etkisi", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G51,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 3,
      '=IFERROR(IF(B16>=MAX(B16:B18),iff_siraBir,IF(B16>=MIN(B16:B18)+ABS(B16-B17),iff_siraIki,iff_siraUc)),iff_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 3,
      '=IFERROR(IF(B17>=MAX(B16:B18),iff_siraBir,IF(B17=MEDIAN(B16:B18),iff_siraIki,iff_siraUc)),iff_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 3,
      '=IFERROR(IF(B18>=MAX(B16:B18),iff_siraBir,IF(B18=MEDIAN(B16:B18),iff_siraIki,iff_siraUc)),iff_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(CONCATENATE("1)",INDEX(A16:A18,MATCH(iff_siraBir,C16:C18,0)),'
      '" 2)",INDEX(A16:A18,MATCH(iff_siraIki,C16:C18,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B16:B18),"₺ #,##0")," TL — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN kural değişimi fark", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=D9", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Sabit pay +1 puan senaryosunda fiyat farkı ",TEXT(D9,"₺ #,##0")," TL")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı (fark)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=MOTOR!G48", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 3, "=iff_kacirilan", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 4, "=MOTOR!G47", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Tahmin (FORECAST)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1, "Seri nokta", yazi=GRİ)
    for i, v in enumerate([1, 2, 3, 4], 28):
        h(ws, i, 1, v, sayi=CATI)
        h(ws, i, 2, f"=D{5+v}", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 32, 1, "Tahmin sonraki", yazi="333333")
    h(ws, 32, 2, "=IFERROR(AVERAGE(B28:B31),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Tahmin yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Tahmin ",TEXT(B32,"₺ #,##0")," · Alt ",TEXT(B25,"₺ #,##0"),'
      '" · Üst ",TEXT(D25,"₺ #,##0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 5, 7, "FarkDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([364500, 405000, 445500, 382500], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([20250, 20250, 22500], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)
    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([364500, 405000, 445500, 382500], 28):
        h(ws, i, 4, v, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo Fiyat Farkı Karşılaştırması"
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
    g3.title = "Senaryo Fiyat Farkı Trendi"
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

    # V-001: hakediş 100000, endeks oranı 1.10, sabit 0.10 → Pn=1.09 → fark=9000
    # V-002: hakediş 200000 * talep_esik 0.05 = 10000
    # V-003: V-001 + V-002 = 19000
    # V-004: tavan 150000 * sabit 0.10 = 15000
    vakalar_data = [
        ("V-001", "Endeks %10 — fiyat farkı",
         "hakedis=vaka_matrah_1; oran=1.10; sabit=IFF-001",
         9000,
         "=IFERROR(ROUND(vaka_matrah_1*((INDEX(tblKurallar[deger_2025],MATCH(\"IFF-001\",tblKurallar[kural_id],0))"
         "+(1-INDEX(tblKurallar[deger_2025],MATCH(\"IFF-001\",tblKurallar[kural_id],0)))*vaka_endeks_orani)-1),2),0)",
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=iff_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Fiyat farkı cetveli örnek"),
        ("V-002", "Talep eşiği × hakediş",
         "hakedis=vaka_matrah_2; esik=IFF-002",
         10000,
         "=IFERROR(ROUND(vaka_matrah_2*INDEX(tblKurallar[deger_2025],MATCH(\"IFF-002\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=iff_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Talep eşiği kontrol"),
        ("V-003", "Birikimli fark + eşik",
         "V-001 + V-002",
         19000,
         "=IFERROR(ROUND(vaka_matrah_1*((INDEX(tblKurallar[deger_2025],MATCH(\"IFF-001\",tblKurallar[kural_id],0))"
         "+(1-INDEX(tblKurallar[deger_2025],MATCH(\"IFF-001\",tblKurallar[kural_id],0)))*vaka_endeks_orani)-1)"
         "+vaka_matrah_2*INDEX(tblKurallar[deger_2025],MATCH(\"IFF-002\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=iff_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Birikim yorum A"),
        ("V-004", "Tavan × sabit pay",
         "tavan=vaka_tavan; sabit=IFF-001",
         15000,
         "=IFERROR(ROUND(vaka_tavan*INDEX(tblKurallar[deger_2025],MATCH(\"IFF-001\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=iff_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "IFF-001 tavan kontrol"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v,
              yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = TL

    formuller = {
        "fark": "=IF(tblVakalar[[#This Row],[vaka_id]]=\"\",\"\",IFERROR(tblVakalar[[#This Row],[beklenen_sonuc]]-tblVakalar[[#This Row],[hesaplanan]],0))",
        "durum": '=IF(tblVakalar[[#This Row],[vaka_id]]="","",IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=iff_tolerans,"TUTARLI","KIRIK"),"KIRIK"))',
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
    for i, (b, he) in enumerate([(9000, 9000), (10000, 10000), (19000, 19000), (15000, 15000)], 6):
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
    h(ws, 3, 1, "KANIT RAPORU — İhale Fiyat Farkı Eskalasyon", kalin=True, boyut=14, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:F3")
    h(ws, 4, 1, "Rapor tarihi", yazi="333333")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 5, 1, "İhale / iş", yazi="333333")
    h(ws, 5, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Dönem / yıl", yazi="333333")
    h(ws, 6, 2, '=CONCATENATE(GIRDI!B7," / ",hesapYili)', yazi=DIS_REF_YEŞIL)

    h(ws, 8, 1, "Girdi özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, (et, form, sayi) in enumerate([
        ("Sözleşme bedeli", "=GIRDI!B9", TL),
        ("Baz endeks", "=GIRDI!B10", CATI),
        ("Güncel endeks", "=GIRDI!B11", CATI),
        ("İşçilik ağırlığı", "=GIRDI!B12", YÜZDE),
        ("Malzeme ağırlığı", "=GIRDI!B13", YÜZDE),
        ("Akaryakıt ağırlığı", "=GIRDI!B14", YÜZDE),
        ("Hakediş tutarı", "=iff_hakedisTutari", TL),
    ], 9):
        h(ws, i, 1, et, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)

    h(ws, 17, 1, "Hesap zinciri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 18, 1, "Fiyat farkı", yazi="333333")
    h(ws, 18, 2, "=iff_fiyatFarki", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Fiyat farkı (karar)", yazi="333333")
    h(ws, 19, 2, "=iff_kacirilan", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "KARAR", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=iff_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 21, 1, "Gerekçe", yazi="333333")
    h(ws, 21, 2, "=iff_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B21:F21")

    h(ws, 23, 1, "Madde atıfları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "=iff_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B23:F23")

    h(ws, 25, 1, "Kanıt gövdesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=iff_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B25:F25")

    h(ws, 27, 1, "Parmak izi (SHA-256 kısaltma)", yazi="333333")
    h(ws, 27, 2, "=iff_parmakIzi", yazi=GRİ)

    h(ws, 29, 1, "Hazırlayan (imza)", yazi="333333")
    h(ws, 29, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B29", "İmza", "Raporu hazırlayan kişinin adını yazın.")
    h(ws, 30, 1, "Onaylayan (imza)", yazi="333333")
    h(ws, 30, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B30", "Onay", "Onaylayan kişinin adını yazın.")

    h(ws, 32, 1,
      "Uyarı: Bu çıktı karar destektir; kesin fiyat farkı talebi sonucu veya hukuki görüş değildir. "
      f"Sürüm {SURUM}.",
      yazi=GRİ, boyut=9, kaydir=True)
    ws.merge_cells("A32:F32")

    baski_hazirla(ws, "A1:F33", f"{URUN_AD} · {SURUM}")
    ws.page_setup.orientation = "portrait"
    genislik(ws, {"A": 28, "B": 22, "C": 14, "D": 14, "E": 14, "F": 14})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim paneli — İhale Fiyat Farkı", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpiler = [
        (2, "Pn", "=IFERROR(MOTOR!G24,0)", YÜZDE),
        (3, "Fiyat farkı", "=IFERROR(iff_fiyatFarki,0)", TL),
        (4, "Hakediş", "=IFERROR(iff_hakedisTutari,0)", TL),
        (5, "Fark (karar)", "=IFERROR(iff_kacirilan,0)", TL),
        (6, "Güven", "=IFERROR(MOTOR!G38,0)", CATI),
        (7, "Risk", "=IFERROR(MOTOR!G39,0)", CATI),
        (8, "Sabit pay", "=IFERROR(MOTOR!G8,0)", YÜZDE),
        (9, "Talep eşiği", "=IFERROR(MOTOR!G9,0)", YÜZDE),
        (10, "Fark oranı", "=IFERROR(MOTOR!G27,0)", YÜZDE),
        (11, "M-SEN", "=IFERROR(iff_kuralDegisimSenaryo,0)", TL),
        (12, "Tahmin", "=IFERROR(SENARYO!B32,0)", TL),
        (13, "Vaka", "=IFERROR(iff_vakaDurum,\"-\")", None),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=9, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=iff_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=iff_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 10, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Sözleşme / güncel endeks",
         "=IFERROR(MOTOR!G21,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Endeks oranı ",TEXT(C11,"0.0%")," — yükseldikçe fark artar")'),
        ("T2", "Fiyat farkı tutarı",
         "=IFERROR(iff_fiyatFarki,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Fiyat farkı ",TEXT(C12,"₺ #,##0")," TL")'),
        ("O1", "Senaryo fark sapması",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,STDEV.P(SENARYO!D6:D9)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo fark sapması ",TEXT(C13,"₺ #,##0")," TL")'),
        ("O2", "Tornado max etki",
         "=IFERROR(iff_duyarlilikSira,\"-\")",
         "=IFERROR(iff_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo bant",
         "=IFERROR(iff_senaryoKarsilastirma,0)",
         "=IFERROR(iff_yorumSenaryo,\"-\")"),
        ("O6", "İşçilik payı",
         "=IFERROR(IF(MOTOR!G22=0,0,MOTOR!G16/MOTOR!G22),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İşçilik / ağırlık toplam ",TEXT(C16,"0.0%"))'),
        ("O8", "Kalite skoru",
         "=IFERROR(KONTROLLER!B12,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalite ",TEXT(C17,"0")," puan")'),
        ("M", "M-SEN fark",
         "=IFERROR(iff_kuralDegisimSenaryo,0)",
         "=IFERROR(iff_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin aralık",
         "=IFERROR(iff_tahminAralik,0)",
         "=IFERROR(iff_yorumTahmin,\"-\")"),
        ("I2", "Senaryo fark P90",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!D6:D9,iff_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"P90 fark ",TEXT(C20,"₺ #,##0")," TL")'),
    ]
    for i, (kod, ad, form, yorum) in enumerate(moduller):
        r = 11 + i
        h(ws, r, 1, kod, kalin=True, hiza="center", yazi="B08948")
        h(ws, r, 2, ad, kaydir=True)
        h(ws, r, 3, form, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[r].height = 28
        if kod in ("T1", "O6"):
            ws.cell(r, 3).number_format = YÜZDE
        elif kod != "O2":
            ws.cell(r, 3).number_format = TL

    h(ws, 23, 1, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("İşçilik payı farkı", 162_000),
        ("Malzeme payı farkı", 182_250),
        ("Akaryakıt payı farkı", 60_750),
        ("Toplam fiyat farkı", 405_000),
        ("Sözleşme bedeli", 10_000_000),
    ], 24):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)

    h(ws, 23, 4, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 5, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Fiyat farkı", 405_000),
        ("Hakediş", 2_500_000),
        ("İyimser fark", 364_500),
        ("Kötümser fark", 445_500),
        ("M-SEN fark", 382_500),
    ], 24):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Kalem Payı Farkları"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=28), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=28))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G9")

    g2 = BarChart()
    g2.type = "col"
    g2.title = "Fiyat Farkı / Hakediş / Senaryo"
    g2.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "G23")

    g3 = LineChart()
    g3.title = "Fark Çizgisi"
    g3.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g3.height, g3.width = 7, 12
    ws.add_chart(g3, "P9")

    h(ws, 30, 1, "Metrik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 2, "Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Fiyat farkı", 405_000),
        ("Hakediş", 2_500_000),
        ("Sözleşme", 10_000_000),
    ], 31):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "KPI Fark / Hakediş / Sözleşme"
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
        ("Negatif sözleşme bedeli?",
         '=IF(GIRDI!B9<0,"NEGATİF","TEMİZ")'),
        ("Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(iff_yilTaban),N(hesapYili)<=N(iff_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        ("Baz endeks sıfır mı?",
         '=IF(IFERROR(MOTOR!G14,0)=0,"SIFIR","NORMAL")'),
        ("Vaka kırık mı?",
         '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")'),
        ("Fark eşiği aşıyor mu?",
         '=IF(IFERROR(iff_kacirilan,0)>iff_kacirilanEsik,"AŞIYOR","NORMAL")'),
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
    h(ws, 14, 2, "=IFERROR(iff_girisBeklenen+COUNTA(GIRDI!B6:B16)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Dolu giriş", yazi="333333")
    h(ws, 15, 2, "=COUNTA(GIRDI!B6:B16)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Kalite skoru (modül)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=IFERROR(ROUND(IF(OR(B14=0,B14=\"\"),0,B15/B14*100),0),0)", sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)

    h(ws, 17, 1, "Fiyat farkı kontrol", yazi="333333")
    h(ws, 17, 2, "=IFERROR(iff_fiyatFarki,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Fark (karar) kontrol", yazi="333333")
    h(ws, 18, 2, "=IFERROR(iff_kacirilan,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Karar kontrol", yazi="333333")
    h(ws, 19, 2, "=IFERROR(iff_kararMetni,\"-\")", yazi=DIS_REF_YEŞIL)

    genislik(ws, {"A": 40, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek senaryo kütüphanesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:I3")
    h(ws, 4, 1, "Bu sayfa demo değerleri saklar; GIRDI'ye kopyalanabilir.", yazi=GRİ)
    sutunlar = ["Senaryo", "Sözleşme", "Baz", "Güncel", "İşçilik", "Malzeme",
                "Akaryakıt", "Hakediş", "Fiyat Farkı"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Fiyat Farkı": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(ROUND(tblOrnek[[#This Row],[Hakediş]]*"
            "((INDEX(tblKurallar[deger_2025],MATCH(\"IFF-001\",tblKurallar[kural_id],0))"
            "+(1-INDEX(tblKurallar[deger_2025],MATCH(\"IFF-001\",tblKurallar[kural_id],0)))"
            "*IF(tblOrnek[[#This Row],[Baz]]=0,0,tblOrnek[[#This Row],[Güncel]]/tblOrnek[[#This Row],[Baz]]))"
            "-1),iff_yuvarMax),0))"
        ),
    }
    # Need 9 columns - Fiyat Farkı is last; add dummy for table formula on another col
    # Table has Senaryo..Hakediş as inputs, Fiyat Farkı calculated
    tablo_ekle(ws, "tblOrnek", "A5:I1004", sutunlar, formuller)
    ornekler = [
        ("Demo ana", 10_000_000, 100, 118, 0.40, 0.45, 0.15, 2_500_000),
        ("Yüksek endeks", 10_000_000, 100, 130, 0.40, 0.45, 0.15, 2_500_000),
        ("Düşük hakediş", 8_000_000, 100, 110, 0.35, 0.50, 0.15, 800_000),
        ("Sadece işçilik ağır", 5_000_000, 100, 115, 0.70, 0.20, 0.10, 1_200_000),
        ("Sıfır fark", 3_000_000, 100, 100, 0.40, 0.40, 0.20, 500_000),
    ]
    for i, row in enumerate(ornekler, 6):
        for k, v in enumerate(row, 1):
            ws.cell(i, k).value = v
            if k == 2 or k == 8:
                ws.cell(i, k).number_format = TL
            elif k in (5, 6, 7):
                ws.cell(i, k).number_format = YÜZDE
            elif k in (3, 4):
                ws.cell(i, k).number_format = CATI
            if k >= 2:
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
        ws.cell(i, 1).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
        ws.cell(i, 1).protection = Protection(locked=False)
        for kolon in range(1, 9):
            _kim(ws, f"{get_column_letter(kolon)}{i}", sutunlar[kolon - 1], "Örnek senaryo değeri")

    for col, ad in [("B", "Sözleşme"), ("C", "Baz"), ("D", "Güncel"),
                    ("E", "İşçilik"), ("F", "Malzeme"), ("G", "Akaryakıt"), ("H", "Hakediş")]:
        dogrulama(ws, "decimal", "0", f"{col}6:{col}1004",
                  baslik=ad, mesaj=f"{ad} değerini girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ile 1e12 arasında.",
                  isaret="between", f2="1000000000000")

    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6, 7, 8])
    h(ws, 5, 12, "FarkDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([405000, 675000, 72000, 162000, 0], 6):
        h(ws, i, 12, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Örnek Senaryo Fiyat Farkı"
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
        _kim(ws, f"C{i}", "Yorum", "Endeks yorumu")
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
        ("iff_tolerans", 0.01, "TL", "Vaka tutarlılık toleransı", "Uygulama notu", "01.01.2025", "11.08.2026"),
        ("iff_kacirilanEsik", 50_000, "TL", "PANO fark uyarı eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("iff_senaryoIyi", 0.90, "çarpan", "İyimser endeks çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("iff_senaryoKotu", 1.10, "çarpan", "Kötümser endeks çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("iff_oranArtisPuan", 0.01, "oran", "M-SEN sabit +1 puan", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("iff_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("iff_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("iff_tornadoMatrah", 1.05, "çarpan", "Tornado endeks şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("iff_tornadoMevcut", 0.95, "çarpan", "Tornado hakediş şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("iff_isverenOran", 0.10, "oran", "Yorum B yedek sabit artışı", "Fiyat farkı model", "01.01.2025", "11.08.2026"),
        ("iff_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "11.08.2026"),
        ("iff_parmakIzi", "IFF-PRO-1.0.0", "metin", "Dosya parmak izi etiketi", "Üretim", "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv", "01.01.2025", "11.08.2026"),
        ("iff_girisBeklenen", 11, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("iff_riskDusuk", 15, "puan", "TALEP ET risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("iff_riskOrta", 55, "puan", "BEKLE risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("iff_riskYuksek", 85, "puan", "VAZGEÇ risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("iff_riskEsikOran", 0.5, "oran", "Yedek risk oranı", "İç politika", "01.01.2025", "11.08.2026"),
        ("iff_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı (hesapYili−taban)", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("iff_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("iff_dayanakCarpan", 10, "adet", "Yedek çarpan", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("iff_bazCarpan", 1, "çarpan", "Baz senaryo çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("iff_sifirCarpan", 0, "çarpan", "Nötr çarpan (sıfır)", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("iff_siraBir", 1, "adet", "Tornado sıra 1", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("iff_siraIki", 2, "adet", "Tornado sıra 2", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("iff_siraUc", 3, "adet", "Tornado sıra 3", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("iff_tahminX", 5, "adet", "FORECAST X noktası", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("vaka_matrah_1", 100_000, "TL", "Vaka1 hakediş", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_matrah_2", 200_000, "TL", "Vaka2 hakediş", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_tavan", 150_000, "TL", "Vaka4 tavan", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_endeks_orani", 1.10, "çarpan", "Vaka1/3 endeks oranı", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
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
        "TALEP ET — fiyat farkı oranı talep eşiğinin üzerinde; talep dosyası hazırlayın.",
        "BEKLE — fark bekle bandında; endeks ve hakediş dönemini izleyin.",
        "VAZGEÇ — fark düşük veya olumsuz; talep maliyetine değmeyebilir.",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 10, 1, "Belirsizlik beyanları (M05)", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari(["endeks_secimi"]), 11):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 13, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2,
      "Pn = sabit pay + (1 − sabit pay) × (güncel endeks / baz endeks); "
      "fiyat farkı = hakediş × (Pn − 1). Endeks seçimi tartışmalıdır; yorum B ile yan yana okunmalıdır.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B13:J13")
    h(ws, 14, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 2,
      "Yorum B, sabit payı +1 puan sıkılaştırır. Dosya her iki yorumun sonucunu MOTOR'da üretir.",
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
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("TALEP ET",B6))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("BEKLE",B6))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("VAZGEÇ",B6))'], font=kirmizi, fill=k_fill))
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
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("TALEP",E6))'], font=yesil))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("BEKLE",E6))'], font=amber))
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
    ws.conditional_formatting.add("F7:F53", FormulaRule(formula=['F7="IFF-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F53", FormulaRule(formula=['F7="IFF-005"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "C", "D", "E", "F", "G", "H", "I"):
        ws.conditional_formatting.add(f"{harf}6:{harf}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("I6:I20", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B20", FormulaRule(formula=['ISNUMBER(SEARCH("TALEP ET",B20))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B20", FormulaRule(formula=['ISNUMBER(SEARCH("BEKLE",B20))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B20", FormulaRule(formula=['ISNUMBER(SEARCH("VAZGEÇ",B20))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B19", CellIsRule(operator="greaterThan", formula=["0"], font=kirmizi))


def ad_tanimlari(wb):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "iff_hakedisTutari", "GIRDI!$B$15")
    ad_ekle(wb, "iff_fiyatFarki", "MOTOR!$G$28")
    ad_ekle(wb, "iff_kacirilan", "MOTOR!$G$30")
    ad_ekle(wb, "iff_kararMetni", "MOTOR!$G$34")
    ad_ekle(wb, "iff_kararGerekce", "MOTOR!$G$35")
    ad_ekle(wb, "iff_maddeAtifMetni", "MOTOR!$G$52")
    ad_ekle(wb, "iff_kanitRaporu", "MOTOR!$G$53")
    ad_ekle(wb, "iff_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "iff_parmakIzi", "AYARLAR!$B$19")

    ad_ekle(wb, "iff_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "iff_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "iff_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "iff_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "iff_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "iff_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "iff_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "iff_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "iff_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "iff_vakaFark", "VAKALAR!$B$13")

    ayar_map = {
        "iff_tolerans": 8,
        "iff_kacirilanEsik": 9,
        "iff_senaryoIyi": 10,
        "iff_senaryoKotu": 11,
        "iff_oranArtisPuan": 12,
        "iff_tahminAlt": 13,
        "iff_tahminUst": 14,
        "iff_tornadoMatrah": 15,
        "iff_tornadoMevcut": 16,
        "iff_isverenOran": 17,
        "iff_yuzdelikOran": 18,
        "iff_girisBeklenen": 21,
        "iff_riskDusuk": 22,
        "iff_riskOrta": 23,
        "iff_riskYuksek": 24,
        "iff_riskEsikOran": 25,
        "iff_yilTaban": 26,
        "iff_yuvarMax": 27,
        "iff_dayanakCarpan": 28,
        "iff_bazCarpan": 29,
        "iff_sifirCarpan": 30,
        "iff_siraBir": 31,
        "iff_siraIki": 32,
        "iff_siraUc": 33,
        "iff_tahminX": 34,
        "vaka_matrah_1": 35,
        "vaka_matrah_2": 36,
        "vaka_tavan": 37,
        "vaka_endeks_orani": 38,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "iff_modulT1": (11, "C"), "iff_modulT1Yorum": (11, "D"),
        "iff_modulT2": (12, "C"), "iff_modulT2Yorum": (12, "D"),
        "iff_modulO1": (13, "C"), "iff_modulO1Yorum": (13, "D"),
        "iff_modulO6": (16, "C"), "iff_modulO6Yorum": (16, "D"),
        "iff_modulO8": (17, "C"), "iff_modulO8Yorum": (17, "D"),
        "iff_modulI2": (20, "C"), "iff_modulI2Yorum": (20, "D"),
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
    varsayilan = os.path.join(kok, "IhaleFiyatFarkiEskalasyonPro.xlsx")
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
    cikti = os.path.join(repo, "cikti", "IhaleFiyatFarkiEskalasyonPro.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
