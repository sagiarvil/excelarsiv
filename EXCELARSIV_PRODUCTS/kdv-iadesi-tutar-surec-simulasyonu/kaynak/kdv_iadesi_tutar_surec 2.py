#!/usr/bin/env python3
"""
KDV İadesi Tutar ve Süreç Simülasyonu — üretim betiği (A1+A3 / manda v6).
Azami iade tutarı + teminat + süreç ayı + finansman maliyeti + durum makinesi.
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
from ortak.akis_motoru import (
    DURUM_BASLIK,
    durum_haritasi_dogrula,
    gecersiz_gecis_formul,
    yas_kova_formul,
    yaslandirma_gun,
    yetim_kayit_formul,
    zincir_kirik_formul,
)
from ortak.mevzuat_motoru import (
    KURAL_BASLIK,
    MOTOR_BASLIK,
    VAKA_BASLIK,
    belirsizlik_beyanlari,
)

URUN_AD = "KDV İadesi Tutar ve Süreç Simülasyonu"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 11)

# Demo 2026: yüklenilen 5M − indirilecek 2M = 3M devreden; ihraç 20M×%20=4M
# azami=MIN(3M,4M)=3M; teminat %20=600k; inceleme 6 ay; finansman %30 → 450k
# 450k < 600k ve azami ≥ eşik → BAŞVUR
DEMO = {
    "sirket_adi": "Örnek İhracat A.Ş.",
    "donem": "2026-Q2",
    "hesapYili": 2026,
    "iade_turu": "TAM",
    "yuklenilen_kdv": 5_000_000,
    "indirilecek_kdv": 2_000_000,
    "ihrac_hasilati": 20_000_000,
    "teminat_orani": 0.20,
    "inceleme_ayi": 6,
    "finansman_yillik": 0.30,
    "yorum_modu": "A",
}

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "BASVURU", "durum_adi": "Başvuru", "sira": 1,
     "izinli_sonraki_durumlar": "INCELME|IPTAL", "kategori": "acik"},
    {"durum_id": "INCELME", "durum_adi": "İnceleme", "sira": 2,
     "izinli_sonraki_durumlar": "EK_BELGE|ODEME|IPTAL", "kategori": "acik"},
    {"durum_id": "EK_BELGE", "durum_adi": "Ek Belge", "sira": 3,
     "izinli_sonraki_durumlar": "INCELME|IPTAL", "kategori": "acik"},
    {"durum_id": "ODEME", "durum_adi": "Ödeme", "sira": 4,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "IPTAL", "durum_adi": "İptal", "sira": 5,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS: list[str] = []
for _d in DURUM_HARITASI:
    IZINLI_GECIS.extend(
        f'{_d["durum_id"]}|{_s}'
        for _s in (_d["izinli_sonraki_durumlar"] or "").split("|")
        if _s
    )

AKIS_HDR, AKIS_ILK, AKIS_SON = 5, 6, 1005
KART_HDR, KART_ILK, KART_SON = 5, 6, 1005


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    h(ws, r, c, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
    ws.cell(r, c).protection = Protection(locked=False)
    if baslik:
        _kim(ws, f"{get_column_letter(c)}{r}", baslik, mesaj or baslik)


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: KDV iadesi tutar/süreç kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Azami iade, teminat ve karar bozulur | Kaynak: KDVK / iade tebliği",
    )


def kural_cek(kural_id: str) -> str:
    """Oranları tblKurallar'dan çeker; yıl seçimi kis_yilTaban ile (G07: adında rakam yok)."""
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-kis_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(kis_yuvarMax,IFERROR(N(G12),0)))"


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
      "Dönem, iade türü, yüklenilen/indirilecek KDV ve süreç girdilerini girin; "
      "azami iade, teminat, net nakit ve BAŞVUR / BEKLE / TEMİNATLI BAŞVUR kararını üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Çok yıllı kural tablosu (2025/2026/2027) ile KDV oranı ve eşikler",
        "Azami iade, teminat, finansman maliyeti ve net nakit etkisi yan yana",
        "Başvuru→inceleme→ek belge→ödeme durum makinesi (A3)",
        "Baz / iyimser / kötümser + M-SEN (teminat oranı +5 puan) senaryoları",
        "Altın vakalar ve KANIT_RAPORU (A4 imza) + belirsizlik: inceleme süre yorumu A/B",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "SMMM — azami iade ve madde atfını müşteriye vermek",
        "CFO — BAŞVUR / BEKLE / TEMİNATLI BAŞVUR kararını nakit planına yazmak",
        "Ortak / denetçi — imzalı kanıt ve süreç zincirini dosyalamak",
    ], 14):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) HIZLI_BASLANGIC → 2) GIRDI sarı hücreler → 3) PANO karar → 4) AKIS süreç → 5) KANIT_RAPORU.",
      yazi="333333", kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 21, 1, f"Sürüm {SURUM} | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 72})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Üç adımda sonuç", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Girdiler", "GIRDI'de dönem, iade türü, KDV tutarları, teminat, inceleme ayı ve finansman oranını girin."),
        ("Adım 2 — Karar", "PANO'da azami iade, net nakit ve BAŞVUR / BEKLE / TEMİNATLI BAŞVUR rozetini izleyin."),
        ("Adım 3 — Süreç", "AKIS'te başvuru→inceleme→ödeme adımlarını işleyin; KANIT_RAPORU yazdırın."),
    ]
    for i, (bas, acik) in enumerate(adimlar, 5):
        h(ws, i * 2, 1, bas, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i * 2 + 1, 1, acik, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i * 2 + 1, start_column=1, end_row=i * 2 + 1, end_column=20)
    h(ws, 14, 1, "Sık yapılan hatalar", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate([
        "İhraç hasılatını KDV dahil yazmak (hasılat KDV hariç olmalı)",
        "Teminat oranını yüzde yerine ondalık yanlış girmek (0,20 = %20)",
        "İnceleme ayını boş bırakıp finansman maliyetini yok saymak",
        "AKIS'te kaynak kart kimliğini boş bırakmak (YETİM riski)",
    ], 15):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
    genislik(ws, {"A": 80})


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Şirket ve iade girdileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Oranlar KURALLAR'dan; iade ve nakit formülle üretilir.",
      yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Şirket Adı", DEMO["sirket_adi"], None, "sirket_adi",
         "Şirket unvanını yazın."),
        (7, "Dönem", DEMO["donem"], None, "donem",
         "İade dönemini yazın (ör. 2026-Q2)."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili",
         "2025, 2026 veya 2027 seçin."),
        (9, "İade Türü", DEMO["iade_turu"], None, "iade_turu",
         "TAM veya KISMİ seçin."),
        (10, "Yüklenilen KDV [₺]", DEMO["yuklenilen_kdv"], TL, "yuklenilen_kdv",
         "Dönem yüklenilen KDV tutarını TL girin."),
        (11, "İndirilecek KDV [₺]", DEMO["indirilecek_kdv"], TL, "indirilecek_kdv",
         "İndirilecek KDV tutarını TL girin."),
        (12, "İhraç / Teslim Hasılatı [₺]", DEMO["ihrac_hasilati"], TL, "ihrac_hasilati",
         "KDV hariç ihraç/teslim hasılatını TL girin."),
        (13, "Teminat Oranı", DEMO["teminat_orani"], YÜZDE, "teminat_orani",
         "Teminat oranını ondalık girin (0,20 = %20)."),
        (14, "İnceleme Süresi [ay]", DEMO["inceleme_ayi"], CATI, "inceleme_ayi",
         "Beklenen inceleme süresini ay olarak girin."),
        (15, "Finansman Maliyeti Yıllık", DEMO["finansman_yillik"], YÜZDE, "finansman_yillik",
         "Yıllık finansman maliyet oranını girin."),
        (16, "Yorum Modu (A/B)", DEMO["yorum_modu"], None, "yorum_modu",
         "İnceleme süre yorumu A (standart) veya B (+%20 süre) seçin."),
    ]
    h(ws, 5, 1, "Alan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 2, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 3, "Birim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    for satir, etiket, deger, sayi, _ad, mesaj in alanlar:
        if sayi == TL:
            birim = "₺"
        elif sayi == YÜZDE:
            birim = "%"
        elif _ad == "hesapYili":
            birim = "yıl"
        elif _ad == "inceleme_ayi":
            birim = "ay"
        else:
            birim = "metin"
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim, yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 18, 1, "KDV oranı (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 2, kural_cek("KIS-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Kısmi oran (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2, kural_cek("KIS-002"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Azami iade (ayna) [₺]", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=IFERROR(kis_azamiIade,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 21, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B16)=0,kis_girisBeklenen=0),0,ROUND(COUNTA(B6:B16)/kis_girisBeklenen*100,0)),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)

    h(ws, 23, 1, "Giriş tablosu (otomasyon / boş-mod)", kalin=True, yazi=KOYU_LACIVERT)
    sutunlar = ["Alan Anahtarı", "Değer", "Birim", "Zorunlu", "Kaynak", "Kontrol"]
    baslik_satiri(ws, 24, sutunlar)
    formuller = {
        "Kontrol": (
            '=IF(tblGirdi[[#This Row],[Alan Anahtarı]]="","",'
            'IF(tblGirdi[[#This Row],[Değer]]="","EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblGirdi", "A24:F1023", sutunlar, formuller)
    ornek_satir = [
        ("sirket_adi", DEMO["sirket_adi"], "metin", "Evet", "GIRDI"),
        ("donem", DEMO["donem"], "metin", "Evet", "GIRDI"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Evet", "GIRDI"),
        ("iade_turu", DEMO["iade_turu"], "metin", "Evet", "GIRDI"),
        ("yuklenilen_kdv", DEMO["yuklenilen_kdv"], "TL", "Evet", "GIRDI"),
        ("indirilecek_kdv", DEMO["indirilecek_kdv"], "TL", "Evet", "GIRDI"),
        ("ihrac_hasilati", DEMO["ihrac_hasilati"], "TL", "Evet", "GIRDI"),
        ("teminat_orani", DEMO["teminat_orani"], "oran", "Evet", "GIRDI"),
        ("inceleme_ayi", DEMO["inceleme_ayi"], "ay", "Evet", "GIRDI"),
        ("finansman_yillik", DEMO["finansman_yillik"], "oran", "Evet", "GIRDI"),
        ("yorum_modu", DEMO["yorum_modu"], "A/B", "Evet", "GIRDI"),
    ]
    tl_alanlar = {"yuklenilen_kdv", "indirilecek_kdv", "ihrac_hasilati"}
    yuzde_alanlar = {"teminat_orani", "finansman_yillik"}
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 25):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in tl_alanlar:
            ws.cell(i, 2).number_format = TL
        elif a in yuzde_alanlar:
            ws.cell(i, 2).number_format = YÜZDE
        elif a in ("hesapYili", "inceleme_ayi"):
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
    dogrulama(ws, "list", "ListeIadeTur", "B9",
              baslik="İade Türü", mesaj="TAM veya KISMİ seçin.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.", bos=False)
    dogrulama(ws, "list", "ListeYorumAB", "B16",
              baslik="Yorum", mesaj="Yorum modunda A veya B seçin.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B seçin.", bos=False)
    dogrulama(ws, "list", "ListeYillar", "B27",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeIadeTur", "B28",
              baslik="İade Türü", mesaj="Tablo: TAM/KISMİ.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeYorumAB", "B35",
              baslik="Yorum", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.")
    for aralik, baslik, mesaj in [
        ("B10", "Yüklenilen KDV", "0 ile 1e12 arasında TL."),
        ("B11", "İndirilecek KDV", "0 ile 1e12 arasında TL."),
        ("B12", "İhraç Hasılatı", "0 ile 1e12 arasında TL."),
        ("B29", "Yüklenilen KDV", "Tablo: TL tutar."),
        ("B30", "İndirilecek KDV", "Tablo: TL tutar."),
        ("B31", "İhraç Hasılatı", "Tablo: TL tutar."),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj=mesaj,
                  hata_baslik="Geçersiz tutar", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000000")
    for aralik, baslik in [("B13", "Teminat"), ("B15", "Finansman"), ("B32", "Teminat tablo"), ("B34", "Finansman tablo")]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj="0-1 arası oran.",
                  hata_baslik="Geçersiz oran", hata_mesaj="0 ile 1 arası.",
                  isaret="between", f2="1")
    for aralik in ("B14", "B33"):
        dogrulama(ws, "whole", "0", aralik,
                  baslik="İnceleme Ayı", mesaj="0-60 ay.",
                  hata_baslik="Geçersiz", hata_mesaj="0-60.",
                  isaret="between", f2="60")

    giris_hucreleri(ws, 6, 16, [2])
    giris_hucreleri(ws, 25, 35, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 40, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1024, "Sarı alanlar giriş; oran ve iade formülleri kilitlidir.")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Mevzuat kural tablosu (yıl yan yana)", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["KIS-001", "KDVK md.28/29", "kdv_orani", "genel KDV", "kdv_genel_oran",
         0.20, 0.20, 0.20, "01.01.2025", "Genel KDV oranı"],
        ["KIS-002", "İade tebliği", "kismi_oran", "kısmi iade", "kismi_iade_orani",
         0.50, 0.50, 0.50, "01.01.2025", "Kısmi iade çarpanı"],
        ["KIS-003", "İç politika", "basvur_esik", "azami ≥ eşik", "basvur_esik_tl",
         500000, 500000, 500000, "01.01.2025", "BAŞVUR karar eşiği (TL)"],
        ["KIS-004", "İç politika", "bekle_finansman", "finansman/azami", "bekle_finansman_oran",
         0.25, 0.25, 0.25, "01.01.2025", "BEKLE finansman oranı eşiği"],
        ["KIS-005", "Genel", "yuvarlama", "her hesap", "yuvarlama_ondalik",
         2, 2, 2, "01.01.2025", "Uygulama notu"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = None
            if k in (6, 7, 8):
                sayi = YÜZDE if s[0] in ("KIS-001", "KIS-002", "KIS-004") else CATI
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-kis_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "KDV oranı (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("KIS-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Kısmi oran (aktif yıl)", yazi="333333")
    h(ws, 16, 2, kural_cek("KIS-002"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"KIS-001=",TEXT(B15,"0.0%"),"KIS-002=",TEXT(B16,"0.0%"),'
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

    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif mevzuat yılı", "KIS-005")  # G7
    ekle("KDV oranı", km("KIS-001"), "oran", "Genel KDV oranı", "KIS-001")  # G8
    ekle("Kısmi oran", km("KIS-002"), "oran", "Kısmi iade çarpanı", "KIS-002")  # G9
    ekle("Başvur eşiği", km("KIS-003"), "TL", "BAŞVUR azami eşiği", "KIS-003")  # G10
    ekle("Bekle finansman oranı", km("KIS-004"), "oran", "BEKLE finansman eşiği", "KIS-004")  # G11
    ekle("Yuvarlama ondalık", km("KIS-005"), "adet", "Yuvarlama basamağı", "KIS-005")  # G12
    ekle("Yüklenilen KDV", "=GIRDI!B10", "TL", "Yüklenilen KDV", "KIS-001")  # G13
    ekle("İndirilecek KDV", "=GIRDI!B11", "TL", "İndirilecek KDV", "KIS-001")  # G14
    ekle("İhraç hasılatı", "=GIRDI!B12", "TL", "İhraç/teslim hasılatı", "KIS-001")  # G15
    ekle("Teminat oranı girdi", "=GIRDI!B13", "oran", "Teminat oranı", "KIS-005")  # G16
    ekle("İnceleme ayı girdi", "=GIRDI!B14", "adet", "İnceleme süresi ay", "KIS-005")  # G17
    ekle("Finansman yıllık", "=GIRDI!B15", "oran", "Yıllık finansman", "KIS-005")  # G18
    ekle("Süre çarpanı (yorum)",
         '=IF(GIRDI!B16="B",1+kis_sureArtisOran,1)',
         "çarpan", "Yorum A/B süre çarpanı", "KIS-005")  # G19
    ekle("İnceleme ayı efektif", "=G17*G19", "adet", "Yorum sonrası süre", "KIS-005")  # G20
    ekle("Devreden KDV", "=MAX(0,G13-G14)", "TL", "Yüklenilen−indirilecek", "KIS-001")  # G21
    ekle("İhraç KDV sınırı", f"=ROUND(G15*G8,{_YV})", "TL", "Hasılat × KDV oranı", "KIS-001")  # G22
    ekle("Azami iade brüt", f"=ROUND(MIN(G21,G22),{_YV})", "TL", "MIN(devreden,ihraç sınırı)", "KIS-001")  # G23
    ekle("Azami iade (tür)",
         f'=ROUND(IF(GIRDI!B9="KISMİ",G23*G9,G23),{_YV})',
         "TL", "TAM veya KISMİ azami", "KIS-002")  # G24
    ekle("Teminat tutarı", f"=ROUND(G24*G16,{_YV})", "TL", "Azami × teminat oranı", "KIS-005")  # G25
    ekle("Finansman maliyeti",
         f"=ROUND(G24*G18*(G20/12),{_YV})",
         "TL", "Azami×yıllık×(ay/12)", "KIS-005")  # G26
    ekle("Net nakit (bekle)", "=G24-G26", "TL", "Azami − finansman", "KIS-005")  # G27
    ekle("Net nakit (pano)", "=G27", "TL", "Net nakit ayna", "KIS-005")  # G28
    ekle("Azami iade (pano)", "=G24", "TL", "Azami ayna", "KIS-001")  # G29
    ekle("Teminat (pano)", "=G25", "TL", "Teminat ayna", "KIS-005")  # G30
    ekle("Azami / ihraç KDV", "=IF(G22=0,0,G24/G22)", "oran", "İade / ihraç sınırı", "KIS-001")  # G31
    ekle("Teminat / azami", "=IF(G24=0,0,G25/G24)", "oran", "Teminat yoğunluğu", "KIS-005")  # G32
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(G24<G10,"BEKLE",'
         'IF(G26>G25,"TEMINATLI",'
         'IF(IF(G24=0,1,G26/G24)>G11,"BEKLE","BASVUR"))))',
         "kod", "Karar kodu", "KIS-005")  # G33
    ekle("Karar metni",
         '=IF(G33="VERI_YOK","VERİ YOK",'
         'IF(G33="BASVUR","BAŞVUR",'
         'IF(G33="TEMINATLI","TEMİNATLI BAŞVUR","BEKLE")))',
         "metin", "Kullanıcıya karar", "KIS-005")  # G34
    ekle("Gerekçe",
         '=IF(G33="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G33="BASVUR","Azami iade eşik üstünde ve finansman maliyeti teminattan düşük; teminatsız başvuru uygun.",'
         'IF(G33="TEMINATLI","Finansman maliyeti teminat tutarını aşıyor; teminatlı başvuru nakit etkisini iyileştirir.",'
         '"Azami iade düşük veya finansman/azami oranı yüksek; başvuruyu erteleyin veya girdileri gözden geçirin.")))',
         "metin", "Gerekçe cümlesi", "KIS-005")  # G35
    ekle("Finansman / azami", "=IF(G24=0,0,G26/G24)", "oran", "Finansman oranı", "KIS-004")  # G36
    ekle("Süreç ayı", "=G20", "adet", "Efektif inceleme ayı", "KIS-005")  # G37
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B16)=0,kis_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B16)/kis_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "KIS-005")  # G38
    ekle("Risk skoru",
         '=IF(G33="VERI_YOK",0,IF(G33="BASVUR",kis_riskDusuk,'
         'IF(G33="TEMINATLI",kis_riskOrta,kis_riskYuksek)))',
         "puan", "Süreç/nakit riski", "KIS-005")  # G39
    ekle("Senaryo iyimser azami", "=G24*kis_senaryoIyi", "TL", "İyimser azami", "KIS-001")  # G40
    ekle("Senaryo kötümser azami", "=G24*kis_senaryoKotu", "TL", "Kötümser azami", "KIS-001")  # G41
    ekle("İyimser net nakit", f"=ROUND(G40-G40*G18*(G20/12),{_YV})", "TL", "İyimser net", "KIS-005")  # G42
    ekle("Kötümser net nakit", f"=ROUND(G41-G41*G18*(G20/12),{_YV})", "TL", "Kötümser net", "KIS-005")  # G43
    ekle("M-SEN teminat+", "=MIN(1,G16+kis_teminatArtisPuan)", "oran", "Teminat + delta", "KIS-005")  # G44
    ekle("M-SEN teminat tutarı", f"=ROUND(G24*G44,{_YV})", "TL", "M-SEN teminat", "KIS-005")  # G45
    ekle("M-SEN net fark", "=G25-G45", "TL", "Teminat farkı (baz−MSEN)", "KIS-005")  # G46
    ekle("Tahmin üst", "=G27*kis_tahminUst", "TL", "Tahmin üst bant", "KIS-005")  # G47
    ekle("Tahmin alt", "=G27*kis_tahminAlt", "TL", "Tahmin alt bant", "KIS-005")  # G48
    ekle("Tornado: yüklenilen etkisi",
         f"=ABS(ROUND(MIN(MAX(0,(G13*kis_tornadoYuklenilen)-G14),G22)*IF(GIRDI!B9=\"KISMİ\",G9,1),{_YV})-G24)",
         "TL", "Yüklenilen ± etki", "KIS-001")  # G49
    ekle("Tornado: hasılat etkisi",
         f"=ABS(ROUND(MIN(G21,ROUND(G15*kis_tornadoHasilat*G8,{_YV}))*IF(GIRDI!B9=\"KISMİ\",G9,1),{_YV})-G24)",
         "TL", "Hasılat ± etki", "KIS-001")  # G50
    ekle("Tornado: süre etkisi",
         f"=ABS(ROUND(G24*G18*((G20*kis_tornadoSure)/12),{_YV})-G26)",
         "TL", "Süre ± finansman etki", "KIS-005")  # G51
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"KIS-001→KDV oranı","KIS-002→kısmi oran",'
         '"KIS-003→başvur eşiği","KIS-004→bekle finansman")',
         "metin", "Kanıt atıfları", "KIS-001")  # G52
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Azami=",TEXT(G24,"₺ #,##0"),"Teminat=",TEXT(G25,"₺ #,##0"),'
         '"Finansman=",TEXT(G26,"₺ #,##0"),"Net=",TEXT(G27,"₺ #,##0"))',
         "metin", "Rapor özeti", "KIS-005")  # G53

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
      "Motor adımları kilitlidir; oranlar yalnızca KURALLAR'dan gelir.", yazi=GRİ, kaydir=True)
    return len(adimlar)


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve duyarlılık", kalin=True, boyut=13, yazi=KOYU_LACIVERT)

    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Azami Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Azami İade", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Net Nakit", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(kis_senaryoIyi+N(GIRDI!B10)*kis_sifirCarpan,0)",
         "=IFERROR(MOTOR!G40,0)",
         "=IFERROR(MOTOR!G42,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(C6<MOTOR!G10,"BEKLE",IF(MOTOR!G26>MOTOR!G25,"TEMİNATLI BAŞVUR","BAŞVUR")))'),
        ("Baz", "=IFERROR(kis_bazCarpan+N(GIRDI!B10)*kis_sifirCarpan,kis_bazCarpan)",
         "=IFERROR(kis_azamiIade,0)",
         "=IFERROR(kis_netNakit,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(kis_kararMetni,"-"))'),
        ("Kötümser", "=IFERROR(kis_senaryoKotu+N(GIRDI!B10)*kis_sifirCarpan,0)",
         "=IFERROR(MOTOR!G41,0)",
         "=IFERROR(MOTOR!G43,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(C8<MOTOR!G10,"BEKLE",IF(MOTOR!G26>MOTOR!G25,"TEMİNATLI BAŞVUR","BAŞVUR")))'),
        ("M-SEN teminat+", "=IFERROR(kis_bazCarpan+N(GIRDI!B10)*kis_sifirCarpan,kis_bazCarpan)",
         "=IFERROR(kis_azamiIade,0)",
         "=IFERROR(kis_azamiIade-MOTOR!G45,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(C9<MOTOR!G10,"BEKLE","TEMİNATLI BAŞVUR"))'),
    ]
    for i, (ad, carp, az, net, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, az, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, net, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant genişliği (iyimser−kötümser net)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(D6-D8,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorum", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser net ",TEXT(D6,"₺ #,##0")," · Baz ",TEXT(D7,"₺ #,##0"),'
      '" · Kötümser ",TEXT(D8,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:E12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    tornado = [
        ("Yüklenilen KDV ±", "=IFERROR(MOTOR!G49,0)"),
        ("İhraç hasılatı ±", "=IFERROR(MOTOR!G50,0)"),
        ("İnceleme süresi ±", "=IFERROR(MOTOR!G51,0)"),
    ]
    for i, (ad, form) in enumerate(tornado, 16):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, i - 15, sayi=CATI, yazi=GRİ)
    h(ws, 19, 1, "Tornado zirve sırası", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(INDEX(A16:A18,MATCH(MAX(B16:B18),B16:B18,0)),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Tornado yorum", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En etkili değişken: ",B19," etki ",TEXT(MAX(B16:B18),"₺ #,##0")," TL")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN kural değişimi (teminat +5 puan)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=IFERROR(MOTOR!G46,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorum", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Teminat oranı +5 puan senaryosunda teminat farkı ",TEXT(B22,"₺ #,##0")," TL")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı (net nakit)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 1, "Alt", yazi="333333")
    h(ws, 26, 2, "=IFERROR(MOTOR!G48,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 27, 1, "Üst", yazi="333333")
    h(ws, 27, 2, "=IFERROR(MOTOR!G47,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 28, 1, "Orta (ortalama)", yazi="333333")
    h(ws, 28, 2,
      "=IFERROR(IF(COUNT(D6:D8)=0,0,AVERAGE(D6:D8)),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 32, 1, "Tahmin aralık genişliği", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 32, 2, "=IFERROR(B27-B26,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Tahmin yorum", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Net nakit tahmin bandı ",TEXT(B26,"₺ #,##0")," – ",TEXT(B27,"₺ #,##0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 5, 7, "NetDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([2_700_000, 2_550_000, 2_167_500, 2_400_000], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([150000, 120000, 90000], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)
    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([2_700_000, 2_550_000, 2_167_500, 2_400_000], 28):
        h(ws, i, 4, v, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo Net Nakit Karşılaştırması"
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
    g3.title = "Senaryo Net Nakit Trendi"
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

    # V-001: MIN(3M, 20M*0.20)=MIN(3M,4M)=3M with vaka_devreden_1=3M, vaka_hasilat_1=20M
    # Actually use simpler: vaka_hasilat_1 * KIS-001 = expected
    vakalar_data = [
        ("V-001", "Hasılat × KDV oranı",
         "hasilat=vaka_hasilat_1; oran=KIS-001",
         4_000_000,
         "=IFERROR(ROUND(vaka_hasilat_1*INDEX(tblKurallar[deger_2026],MATCH(\"KIS-001\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=kis_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "KDVK ihraç sınırı"),
        ("V-002", "Devreden vs ihraç sınırı MIN",
         "devreden=vaka_devreden_1; sınır=V-001",
         3_000_000,
         "=IFERROR(ROUND(MIN(vaka_devreden_1,vaka_hasilat_1*INDEX(tblKurallar[deger_2026],MATCH(\"KIS-001\",tblKurallar[kural_id],0))),2),0)",
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=kis_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Azami iade brüt"),
        ("V-003", "Kısmi iade = brüt × kısmi oran",
         "brüt=V-002; oran=KIS-002",
         1_500_000,
         "=IFERROR(ROUND(MIN(vaka_devreden_1,vaka_hasilat_1*INDEX(tblKurallar[deger_2026],MATCH(\"KIS-001\",tblKurallar[kural_id],0)))"
         "*INDEX(tblKurallar[deger_2026],MATCH(\"KIS-002\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=kis_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Kısmi iade"),
        ("V-004", "Küçük hasılat × KDV",
         "hasilat=vaka_hasilat_2; oran=KIS-001",
         200_000,
         "=IFERROR(ROUND(vaka_hasilat_2*INDEX(tblKurallar[deger_2026],MATCH(\"KIS-001\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=kis_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "KIS-001 küçük hasılat"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v,
              yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = TL

    formuller = {
        "fark": "=IF(tblVakalar[[#This Row],[vaka_id]]=\"\",\"\",IFERROR(tblVakalar[[#This Row],[beklenen_sonuc]]-tblVakalar[[#This Row],[hesaplanan]],0))",
        "durum": '=IF(tblVakalar[[#This Row],[vaka_id]]="","",IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=kis_tolerans,"TUTARLI","KIRIK"),"KIRIK"))',
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
    for i, (b, he) in enumerate([(4000000, 4000000), (3000000, 3000000), (1500000, 1500000), (200000, 200000)], 6):
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
    h(ws, 3, 1, "KANIT RAPORU — KDV İadesi Tutar ve Süreç", kalin=True, boyut=14, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:F3")
    h(ws, 4, 1, "Rapor tarihi", yazi="333333")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 5, 1, "Şirket", yazi="333333")
    h(ws, 5, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Dönem / yıl", yazi="333333")
    h(ws, 6, 2, '=CONCATENATE(GIRDI!B7," / ",hesapYili)', yazi=DIS_REF_YEŞIL)

    h(ws, 8, 1, "Girdi özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, (et, form, sayi) in enumerate([
        ("İade türü", "=GIRDI!B9", None),
        ("Yüklenilen KDV", "=GIRDI!B10", TL),
        ("İndirilecek KDV", "=GIRDI!B11", TL),
        ("İhraç hasılatı", "=GIRDI!B12", TL),
        ("Teminat oranı", "=GIRDI!B13", YÜZDE),
        ("İnceleme ayı", "=GIRDI!B14", CATI),
    ], 9):
        h(ws, i, 1, et, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)

    h(ws, 16, 1, "Hesap zinciri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 17, 1, "Azami iade", yazi="333333")
    h(ws, 17, 2, "=kis_azamiIade", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 1, "Teminat", yazi="333333")
    h(ws, 18, 2, "=MOTOR!G25", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Finansman maliyeti", yazi="333333")
    h(ws, 19, 2, "=MOTOR!G26", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Net nakit", yazi="333333")
    h(ws, 20, 2, "=kis_netNakit", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 21, 1, "KARAR", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 2, "=kis_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 22, 1, "Gerekçe", yazi="333333")
    h(ws, 22, 2, "=kis_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B22:F22")

    h(ws, 24, 1, "Madde atıfları", yazi="333333")
    h(ws, 24, 2, "=kis_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B24:F24")
    h(ws, 25, 1, "Süreç özeti", yazi="333333")
    h(ws, 25, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"Süreç ayı=",TEXT(MOTOR!G37,"0"),'
      '"Açık kuyruk=",TEXT(COUNTIFS(tblAkis[Durum],"BASVURU")+COUNTIFS(tblAkis[Durum],"INCELME")+COUNTIFS(tblAkis[Durum],"EK_BELGE"),"0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B25:F25")

    h(ws, 27, 1, "Hazırlayan", yazi="333333")
    h(ws, 27, 2, "", zemin=GIRIS_SARI)
    ws.cell(27, 2).protection = Protection(locked=False)
    _kim(ws, "B27", "Hazırlayan", "İmza için ad soyad yazın.")
    h(ws, 28, 1, "Onaylayan", yazi="333333")
    h(ws, 28, 2, "", zemin=GIRIS_SARI)
    ws.cell(28, 2).protection = Protection(locked=False)
    _kim(ws, "B28", "Onaylayan", "Onay imzası için ad yazın.")
    h(ws, 30, 1, "Uyarı: Bu çıktı karar destek niteliğindedir; kesin iade tutarı taahhüt etmez.",
      yazi=KRITIK, boyut=9, kaydir=True)
    ws.merge_cells("A30:F30")
    h(ws, 31, 1, f"Sürüm {SURUM} | Parmak izi: ", yazi=GRİ, boyut=9)
    h(ws, 31, 2, "=kis_parmakIzi", yazi=GRİ, boyut=9)
    baski_hazirla(ws, "A1:F32", f"{URUN_AD} · KANIT")
    genislik(ws, {"A": 28, "B": 55})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim paneli — KDV İadesi Tutar/Süreç", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpiler = [
        (2, "Azamiİade", "=IFERROR(kis_azamiIade,0)", TL),
        (3, "Teminat", "=IFERROR(MOTOR!G25,0)", TL),
        (4, "Finansman", "=IFERROR(MOTOR!G26,0)", TL),
        (5, "NetNakit", "=IFERROR(kis_netNakit,0)", TL),
        (6, "Güven", "=IFERROR(MOTOR!G38,0)", CATI),
        (7, "Risk", "=IFERROR(MOTOR!G39,0)", CATI),
        (8, "KDV.Oran", "=IFERROR(MOTOR!G8,0)", YÜZDE),
        (9, "SüreçAy", "=IFERROR(MOTOR!G37,0)", CATI),
        (10, "İade/İhraç", "=IFERROR(MOTOR!G31,0)", YÜZDE),
        (11, "M-SEN", "=IFERROR(kis_kuralDegisimSenaryo,0)", TL),
        (12, "Tahmin", "=IFERROR(SENARYO!B32,0)", TL),
        (13, "Vaka", "=IFERROR(kis_vakaDurum,\"-\")", None),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=9, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=kis_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=kis_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    # A05 kuyruk izi
    h(ws, 8, 1, "Açık süreç adedi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 8, 2,
      '=COUNTIFS(tblAkis[Durum],"BASVURU")+COUNTIFS(tblAkis[Durum],"INCELME")+COUNTIFS(tblAkis[Durum],"EK_BELGE")',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 8, 3, "Açık tutar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 8, 4,
      '=SUMIF(tblAkis[Durum],"BASVURU",tblAkis[Tutar])+SUMIF(tblAkis[Durum],"INCELME",tblAkis[Tutar])+SUMIF(tblAkis[Durum],"EK_BELGE",tblAkis[Tutar])',
      sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 10, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Azami iade / ihraç KDV",
         "=IFERROR(MOTOR!G31,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İade/ihraç oranı ",TEXT(C11,"0.0%"))'),
        ("T2", "Finansman / azami iade",
         "=IFERROR(MOTOR!G36,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Finansman oranı ",TEXT(C12,"0.0%"))'),
        ("O1", "Senaryo nakit sapması",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,STDEV.P(SENARYO!D6:D9)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo sapması ",TEXT(C13,"₺ #,##0")," TL")'),
        ("O2", "Tornado max etki",
         "=IFERROR(kis_duyarlilikSira,\"-\")",
         "=IFERROR(kis_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo bant",
         "=IFERROR(kis_senaryoKarsilastirma,0)",
         "=IFERROR(kis_yorumSenaryo,\"-\")"),
        ("O6", "Teminat / azami",
         "=IFERROR(MOTOR!G32,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Teminat yoğunluğu ",TEXT(C16,"0.0%"))'),
        ("O8", "Kalite skoru",
         "=IFERROR(KONTROLLER!B12,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalite ",TEXT(C17,"0")," puan")'),
        ("M", "M-SEN teminat farkı",
         "=IFERROR(kis_kuralDegisimSenaryo,0)",
         "=IFERROR(kis_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin aralık",
         "=IFERROR(kis_tahminAralik,0)",
         "=IFERROR(kis_yorumTahmin,\"-\")"),
        ("I2", "Senaryo net P90",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!D6:D9,kis_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"P90 net nakit ",TEXT(C20,"₺ #,##0")," TL")'),
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
        ("Yüklenilen KDV", 5_000_000),
        ("İndirilecek KDV", 2_000_000),
        ("Devreden KDV", 3_000_000),
        ("Azami iade", 3_000_000),
        ("Net nakit", 2_550_000),
    ], 24):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)

    h(ws, 23, 4, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 5, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Teminat", 600_000),
        ("Finansman", 450_000),
        ("İyimser net", 2_805_000),
        ("Kötümser net", 2_167_500),
        ("M-SEN fark", -150_000),
    ], 24):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "İade Bileşenleri"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=28), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=28))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G9")

    g2 = BarChart()
    g2.type = "col"
    g2.title = "Teminat / Finansman / Senaryo"
    g2.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "G23")

    g3 = LineChart()
    g3.title = "Net Nakit Çizgisi"
    g3.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g3.height, g3.width = 7, 12
    ws.add_chart(g3, "P9")

    h(ws, 30, 1, "Metrik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 2, "Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Azami iade", 3_000_000),
        ("Net nakit", 2_550_000),
        ("Teminat", 600_000),
    ], 31):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "KPI Azami / Net / Teminat"
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
        ("Negatif yüklenilen?",
         '=IF(N(GIRDI!B10)<0,"NEGATİF","NORMAL")'),
        ("İndirilecek > yüklenilen?",
         '=IF(N(GIRDI!B11)>N(GIRDI!B10),"AŞIYOR","NORMAL")'),
        ("Teminat oranı > 1?",
         '=IF(N(GIRDI!B13)>1,"AŞIYOR","NORMAL")'),
        ("Vaka tutarlı mı?",
         '=IFERROR(IF(kis_vakaDurum="TUTARLI","TUTARLI","KIRIK"),"KIRIK")'),
        ("Yetim kayıt var mı?",
         '=IF(COUNTIF(tblAkis[YetimBayrak],"YETİM")>0,"HATALI","TEMİZ")'),
        ("Geçersiz geçiş var mı?",
         '=IF(COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")>0,"HATALI","TEMİZ")'),
        ("Zincir kırık var mı?",
         '=IF(COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")>0,"HATALI","TEMİZ")'),
        ("Karar üretildi mi?",
         '=IF(OR(kis_kararMetni="",kis_kararMetni="VERİ YOK"),"EKSİK","TAM")'),
    ]
    h(ws, 4, 1, "Kontrol", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 4, 2, "Sonuç", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, (ad, form) in enumerate(kontroller_list, 5):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Kalite skoru", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(ROUND((COUNTIF(B5:B11,"DOLU")+COUNTIF(B5:B11,"NORMAL")+COUNTIF(B5:B11,"TUTARLI")'
      '+COUNTIF(B5:B11,"TEMİZ")+COUNTIF(B5:B11,"TAM"))/9*100,0),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "Özet tutarlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Azami iade", yazi="333333")
    h(ws, 15, 2, "=IFERROR(kis_azamiIade,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Net nakit", yazi="333333")
    h(ws, 16, 2, "=IFERROR(kis_netNakit,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Azami (ayna)", yazi="333333")
    h(ws, 17, 2, "=IFERROR(kis_azamiIade,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Net (ayna)", yazi="333333")
    h(ws, 18, 2, "=IFERROR(kis_netNakit,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 36, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "548235", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Örnek senaryo satırları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    basliklar = ["Senaryo", "Yuklenilen", "Indirilecek", "IhracHasilat", "TeminatOran",
                 "IncelemeAy", "FinansmanYillik", "IadeTur", "Yil", "Kontrol"]
    baslik_satiri(ws, 5, basliklar)
    satirlar = [
        ("Demo ana", 5_000_000, 2_000_000, 20_000_000, 0.20, 6, 0.30, "TAM", 2026),
        ("Yuksek iade", 8_000_000, 1_000_000, 40_000_000, 0.15, 4, 0.25, "TAM", 2026),
        ("Dusuk tutar", 400_000, 200_000, 1_000_000, 0.20, 8, 0.40, "TAM", 2026),
        ("Kismi iade", 5_000_000, 2_000_000, 20_000_000, 0.20, 6, 0.30, "KISMİ", 2026),
        ("2025 oran", 5_000_000, 2_000_000, 20_000_000, 0.20, 6, 0.30, "TAM", 2025),
    ]
    for i, row in enumerate(satirlar, 6):
        for k, v in enumerate(row, 1):
            sayi = None
            if k in (2, 3, 4):
                sayi = TL
            elif k in (5, 7):
                sayi = YÜZDE
            elif k in (6, 9):
                sayi = CATI
            h(ws, i, k, v, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
            ws.cell(i, k).protection = Protection(locked=False)
            _kim(ws, f"{get_column_letter(k)}{i}", basliklar[k - 1], "Örnek veri satırı")
    for r in range(11, 21):
        for c in range(1, 10):
            h(ws, r, c, None, zemin=GIRIS_SARI)
            ws.cell(r, c).protection = Protection(locked=False)
    formuller = {
        "Kontrol": (
            '=IF(tblOrnek[[#This Row],[Senaryo]]="","",'
            'IF(OR(tblOrnek[[#This Row],[Yuklenilen]]="",tblOrnek[[#This Row],[Yil]]=""),"EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:J20", basliklar, formuller)
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 11)})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", "7030A0", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Yıllar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Yil", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, y in enumerate([2025, 2026, 2027], 6):
        h(ws, i, 1, y, zemin=GIRIS_SARI, sayi=CATI)
        ws.cell(i, 1).protection = Protection(locked=False)
        _kim(ws, f"A{i}", "Yıl", "Hesap yılı listesi")

    h(ws, 3, 3, "Yorum A/B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 3, "Yorum", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate(["A", "B"], 6):
        h(ws, i, 3, v, zemin=GIRIS_SARI)
        ws.cell(i, 3).protection = Protection(locked=False)
        _kim(ws, f"C{i}", "Yorum", "A veya B")

    h(ws, 3, 5, "Evet/Hayır", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 5, "Secim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate(["EVET", "HAYIR"], 6):
        h(ws, i, 5, v, zemin=GIRIS_SARI)
        ws.cell(i, 5).protection = Protection(locked=False)

    h(ws, 3, 7, "İade Türü", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 7, "Tur", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate(["TAM", "KISMİ"], 6):
        h(ws, i, 7, v, zemin=GIRIS_SARI)
        ws.cell(i, 7).protection = Protection(locked=False)
        _kim(ws, f"G{i}", "İade Türü", "TAM veya KISMİ")

    h(ws, 3, 9, "Durumlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 9, "DurumId", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, d in enumerate(DURUM_HARITASI, 6):
        h(ws, i, 9, d["durum_id"], zemin=GIRIS_SARI)
        ws.cell(i, 9).protection = Protection(locked=False)
        _kim(ws, f"I{i}", "Durum", "Durum kimliği")

    h(ws, 3, 11, "İzinli Geçişler", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 11, "Gecis", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, g in enumerate(IZINLI_GECIS, 6):
        h(ws, i, 11, g, zemin=GIRIS_SARI)
        ws.cell(i, 11).protection = Protection(locked=False)
        _kim(ws, f"K{i}", "Geçiş", "eski|yeni")

    genislik(ws, {"A": 12, "C": 12, "E": 12, "G": 12, "I": 14, "K": 22})


def kartlar(ws):
    sayfa_hazirla(ws, "KARTLAR", "2E75B6", URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "İade dosya kartları — her kaydın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, '=CONCATENATE("KIS-",TEXT(COUNTA(tblKartlar[KartId])+1,"00000"))', kalin=True)
    _kim(ws, "B4", "Sonraki Kimlik", "Yeni kartta kullanın")

    sutunlar = ["KartId", "Unvan", "Donem", "IadeTur", "LimitTl",
                "RiskNotu", "Aktif", "ToplamHacim", "OrtYas"]
    baslik_satiri(ws, KART_HDR, sutunlar)
    form = {
        "ToplamHacim": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'SUMIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],tblAkis[Tutar]))'
        ),
        "OrtYas": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'IFERROR(AVERAGEIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],'
            'tblAkis[YasGun]),0))'
        ),
    }
    demo_kartlar = [
        ("KIS-00001", "Örnek İhracat A.Ş.", "2026-Q2", "TAM"),
        ("KIS-00002", "Demo Metal Dış Ticaret", "2026-Q1", "KISMİ"),
        ("KIS-00003", "Anadolu Tekstil İhracat", "2025-Q4", "TAM"),
    ]
    for i, (kid, unvan, donem, tur) in enumerate(demo_kartlar):
        r = KART_ILK + i
        _sari(ws, r, 1, kid, baslik="Kart Kimliği", mesaj="KIS-#####")
        _sari(ws, r, 2, unvan, baslik="Unvan", mesaj="Firma unvanı")
        _sari(ws, r, 3, donem, baslik="Dönem", mesaj="İade dönemi")
        _sari(ws, r, 4, tur, baslik="İade Türü", mesaj="TAM/KISMİ")
        _sari(ws, r, 5, 5_000_000 + i * 500_000, sayi=TL, baslik="Limit", mesaj="Limit TL")
        _sari(ws, r, 6, "Normal", baslik="Risk", mesaj="Risk notu")
        _sari(ws, r, 7, "EVET", baslik="Aktif", mesaj="EVET/HAYIR")
    for r in range(KART_ILK + 3, KART_SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None, sayi=TL if c == 5 else None,
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblKartlar", f"A{KART_HDR}:I{KART_SON}", sutunlar, formuller=form)
    dogrulama(ws, "list", "ListeIadeTur", f"D{KART_ILK}:D{KART_SON}",
              baslik="İade Türü", mesaj="TAM veya KISMİ",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeEvetHayir", f"G{KART_ILK}:G{KART_SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "decimal", "0", f"E{KART_ILK}:E{KART_SON}",
              baslik="Limit", mesaj="Limit ≥ 0",
              hata_baslik="Limit", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    sabitle(ws, f"A{KART_ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 10)})
    ws.column_dimensions["B"].width = 28


def akis(ws):
    sayfa_hazirla(ws, "AKIS", "C65911", URUN_AD, son_kolon=22)
    alt_bant(ws, 3,
             "Süreç satırları — başvuru→inceleme→ödeme + yetim + geçersiz geçiş + zincir",
             son_kolon=18)
    sutunlar = [
        "IslemId", "KaynakKartId", "Tarih", "OncekiDurum", "Durum",
        "BasvuruTutar", "IncelemeTutar", "OdemeTutar", "Tutar", "Aciklama",
        "YasGun", "YasKova", "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu",
    ]
    baslik_satiri(ws, AKIS_HDR, sutunlar)
    yas_f = yaslandirma_gun("tblAkis[[#This Row],[Tarih]]", "raporTarihi").lstrip("=")
    kova_f = yas_kova_formul("tblAkis[[#This Row],[YasGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblAkis[[#This Row],[KaynakKartId]]", "tblKartlar[KartId]").lstrip("=")
    gecis_birlesik = 'tblAkis[[#This Row],[OncekiDurum]]&"|"&tblAkis[[#This Row],[Durum]]'
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    zincir_f = zincir_kirik_formul(
        "tblAkis[[#This Row],[BasvuruTutar]]",
        "tblAkis[[#This Row],[IncelemeTutar]]",
        "zincirTolerans",
    ).lstrip("=")
    form = {
        "YasGun": f'IF(tblAkis[[#This Row],[IslemId]]="","",{yas_f})',
        "YasKova": f'IF(tblAkis[[#This Row],[IslemId]]="","",{kova_f})',
        "YetimBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{yetim_f})',
        "GecisUyarisi": (
            f'IF(OR(tblAkis[[#This Row],[IslemId]]="",'
            f'tblAkis[[#This Row],[OncekiDurum]]=""),"",{gecis_f})'
        ),
        "ZincirBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{zincir_f})',
        "KayitDolu": 'IF(tblAkis[[#This Row],[IslemId]]="","",1)',
    }
    ornek = [
        {"id": "IAD-00001", "kart": "KIS-00001", "tarih": date(2026, 4, 10),
         "onceki": "", "durum": "BASVURU", "bas": 3_000_000, "inc": 3_000_000,
         "odeme": 0, "tutar": 3_000_000, "acik": "Q2 tam iade başvurusu"},
        {"id": "IAD-00002", "kart": "KIS-00001", "tarih": date(2026, 5, 1),
         "onceki": "BASVURU", "durum": "INCELME", "bas": 3_000_000, "inc": 2_950_000,
         "odeme": 0, "tutar": 2_950_000, "acik": "İnceleme aşaması"},
        {"id": "IAD-00003", "kart": "KIS-00002", "tarih": date(2026, 3, 15),
         "onceki": "INCELME", "durum": "EK_BELGE", "bas": 1_500_000, "inc": 1_500_000,
         "odeme": 0, "tutar": 1_500_000, "acik": "Ek belge talebi"},
        {"id": "IAD-00004", "kart": "KIS-00003", "tarih": date(2026, 2, 20),
         "onceki": "INCELME", "durum": "ODEME", "bas": 2_000_000, "inc": 2_000_000,
         "odeme": 2_000_000, "tutar": 2_000_000, "acik": "Ödeme tamam"},
        {"id": "IAD-00005", "kart": "", "tarih": date(2026, 6, 1),
         "onceki": "BASVURU", "durum": "INCELME", "bas": 500_000, "inc": 480_000,
         "odeme": 0, "tutar": 480_000, "acik": "Yetim örnek"},
        {"id": "IAD-00006", "kart": "KIS-00002", "tarih": date(2026, 6, 5),
         "onceki": "BASVURU", "durum": "ODEME", "bas": 800_000, "inc": 800_000,
         "odeme": 800_000, "tutar": 800_000, "acik": "Geçersiz geçiş örnek"},
        {"id": "IAD-00007", "kart": "KIS-00001", "tarih": date(2026, 5, 20),
         "onceki": "INCELME", "durum": "ODEME", "bas": 1_000_000, "inc": 700_000,
         "odeme": 700_000, "tutar": 700_000, "acik": "Zincir sapma örnek"},
    ]
    for i, s in enumerate(ornek):
        r = AKIS_ILK + i
        _sari(ws, r, 1, s["id"], baslik="İşlem", mesaj="IAD-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Kart", mesaj="KartId")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Olay tarihi")
        _sari(ws, r, 4, s["onceki"] or None, baslik="Önceki", mesaj="Önceki durum")
        _sari(ws, r, 5, s["durum"], baslik="Durum", mesaj="Durum haritası")
        _sari(ws, r, 6, s["bas"], sayi=TL, baslik="Başvuru", mesaj="Başvuru tutarı")
        _sari(ws, r, 7, s["inc"], sayi=TL, baslik="İnceleme", mesaj="İnceleme tutarı")
        _sari(ws, r, 8, s["odeme"], sayi=TL, baslik="Ödeme", mesaj="Ödeme tutarı")
        _sari(ws, r, 9, s["tutar"], sayi=TL, baslik="Tutar", mesaj="Kuyruk tutarı")
        _sari(ws, r, 10, s["acik"], baslik="Açıklama", mesaj="Kısa açıklama")
    for r in range(AKIS_ILK + len(ornek), AKIS_SON + 1):
        for c in range(1, 11):
            fmt = TARİH if c == 3 else (TL if c in (6, 7, 8, 9) else None)
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblAkis", f"A{AKIS_HDR}:P{AKIS_SON}", sutunlar, formuller=form)
    dogrulama(ws, "list", "ListeDurumlar", f"D{AKIS_ILK}:D{AKIS_SON}",
              baslik="Önceki Durum", mesaj="Durum seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"E{AKIS_ILK}:E{AKIS_SON}",
              baslik="Durum", mesaj="Durum seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    for c, ad in [(6, "Başvuru"), (7, "İnceleme"), (8, "Ödeme"), (9, "Tutar")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{AKIS_ILK}:{get_column_letter(c)}{AKIS_SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    sabitle(ws, f"A{AKIS_ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 17)})
    ws.column_dimensions["J"].width = 22
    ws.column_dimensions["N"].width = 16


def kuyruklar(ws):
    sayfa_hazirla(ws, "KUYRUKLAR", "1F4E79", URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Durum bazlı canlı kuyruk — kapalı kayıtlar düşer", son_kolon=10)
    h(ws, 5, 1, "Kuyruk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Tutar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for r, (ad, kod) in enumerate([
        ("Başvuru", "BASVURU"),
        ("İnceleme", "INCELME"),
        ("Ek Belge", "EK_BELGE"),
    ], 6):
        h(ws, r, 1, ad)
        h(ws, r, 2, f'=COUNTIFS(tblAkis[Durum],"{kod}",tblAkis[Tutar],">0")', sayi=CATI)
        h(ws, r, 3, f'=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"{kod}")', sayi=TL)
    h(ws, 10, 1, "Kapalı (Ödeme+İptal) — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 1, "Ödeme")
    h(ws, 11, 2, '=COUNTIFS(tblAkis[Durum],"ODEME")', sayi=CATI)
    h(ws, 11, 3, '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"ODEME")', sayi=TL)
    h(ws, 12, 1, "İptal")
    h(ws, 12, 2, '=COUNTIFS(tblAkis[Durum],"IPTAL")', sayi=CATI)
    h(ws, 12, 3, '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"IPTAL")', sayi=TL)
    h(ws, 14, 1, "Yetim adet")
    h(ws, 14, 2, '=COUNTIF(tblAkis[YetimBayrak],"YETİM")', sayi=CATI)
    h(ws, 15, 1, "Geçersiz geçiş")
    h(ws, 15, 2, '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")', sayi=CATI)
    h(ws, 16, 1, "Zincir kırık")
    h(ws, 16, 2, '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")', sayi=CATI)
    h(ws, 18, 1, "Açık toplam adet", kalin=True)
    h(ws, 18, 2, "=B6+B7+B8", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Açık toplam tutar", kalin=True)
    h(ws, 19, 2, "=C6+C7+C8", sayi=TL, yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 40, "B": 12, "C": 16})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Parametreler (tek kaynak) + durum haritası", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = ["anahtar", "deger", "birim", "aciklama", "kaynak", "yururluk_tarihi", "dogrulama_tarihi", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    params = [
        ("hesapYili", DEMO["hesapYili"], "yıl", "Aktif hesap yılı", "Kullanıcı / GIRDI", "01.01.2025", "11.08.2026"),
        ("raporTarihi", RAPOR_TARIH, "tarih", "Rapor tarihi (sabit)", "Üretim", "01.01.2025", "11.08.2026"),
        ("kis_tolerans", 0.01, "TL", "Vaka tutarlılık toleransı", "Uygulama notu", "01.01.2025", "11.08.2026"),
        ("kis_kacirilanEsik", 500_000, "TL", "PANO azami uyarı eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("kis_senaryoIyi", 1.10, "çarpan", "İyimser azami çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kis_senaryoKotu", 0.85, "çarpan", "Kötümser azami çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kis_teminatArtisPuan", 0.05, "oran", "M-SEN teminat +5 puan", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("kis_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kis_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kis_tornadoYuklenilen", 1.05, "çarpan", "Tornado yüklenilen şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kis_tornadoHasilat", 1.05, "çarpan", "Tornado hasılat şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kis_tornadoSure", 1.20, "çarpan", "Tornado süre şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kis_sureArtisOran", 0.20, "oran", "Yorum B süre +%20", "İnceleme süre modeli", "01.01.2025", "11.08.2026"),
        ("kis_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "11.08.2026"),
        ("kis_parmakIzi", "KIS-PRO-1.0.0", "metin", "Dosya parmak izi etiketi", "Üretim", "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv", "01.01.2025", "11.08.2026"),
        ("kis_girisBeklenen", 11, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("kis_riskDusuk", 15, "puan", "BAŞVUR risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("kis_riskOrta", 55, "puan", "TEMİNATLI risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("kis_riskYuksek", 85, "puan", "BEKLE risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("kis_riskEsikOran", 0.5, "oran", "Yedek risk oranı", "İç politika", "01.01.2025", "11.08.2026"),
        ("kis_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("kis_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("kis_dayanakCarpan", 10, "adet", "Yedek çarpan", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("kis_bazCarpan", 1, "çarpan", "Baz senaryo çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kis_sifirCarpan", 0, "çarpan", "Nötr çarpan", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("kis_siraBir", 1, "adet", "Tornado sıra 1", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("kis_siraIki", 2, "adet", "Tornado sıra 2", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("kis_siraUc", 3, "adet", "Tornado sıra 3", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("kis_tahminX", 5, "adet", "FORECAST X noktası", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("vaka_hasilat_1", 20_000_000, "TL", "Vaka1/2/3 hasılat", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_devreden_1", 3_000_000, "TL", "Vaka2/3 devreden", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_hasilat_2", 1_000_000, "TL", "Vaka4 hasılat", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("zincirTolerans", 50_000, "TL", "ZİNCİR KIRIK eşiği", "A04", "01.01.2025", "11.08.2026"),
        ("kis_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi", "01.01.2025", "11.08.2026"),
        ("kis_kapaliDusum", "kapali kategori kuyruktan düşer", "metin", "A06 kapali", "01.01.2025", "11.08.2026"),
        ("kis_olcekHedef", 20000, "adet", "Ölçek hedefi", "Manda A3", "01.01.2025", "11.08.2026"),
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
        h(ws, i, 8, f'=IF(A{i}="","",IF(B{i}="","EKSİK","TAM"))',
          yazi=DIS_REF_YEŞIL, hiza="center")

    # Durum haritası (A01) — izinli / sonraki / kapali
    h(ws, 6, 10, "Durum Haritası", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 7, DURUM_BASLIK, basla=10)
    for i, d in enumerate(DURUM_HARITASI):
        r = 8 + i
        h(ws, r, 10, d["durum_id"])
        h(ws, r, 11, d["durum_adi"])
        h(ws, r, 12, d["sira"], sayi=CATI)
        h(ws, r, 13, d["izinli_sonraki_durumlar"])
        h(ws, r, 14, d["kategori"])
    tablo_ekle(ws, "tblDurumHaritasi", "J7:N12", DURUM_BASLIK)

    dogrulama(ws, "list", "ListeYillar", "B6",
              baslik="Hesap Yılı", mesaj="Aktif yılı seçin.",
              hata_baslik="Geçersiz", hata_mesaj="2025-2027.")
    genislik(ws, {"A": 28, "B": 28, "C": 10, "D": 32, "E": 22, "F": 14, "G": 14, "H": 12,
                  "J": 14, "K": 14, "L": 8, "M": 28, "N": 12})
    h(ws, 45, 2, f"Sürüm {SURUM} | Şifre koruması: {SIFRE} (formül alanları)", yazi=GRİ, boyut=9, kaydir=True)


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar kuralları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "VERİ YOK — giriş tablosu boşsa hesap yapılmaz.",
        "BAŞVUR — azami iade eşik üstünde ve finansman maliyeti teminattan düşük.",
        "TEMİNATLI BAŞVUR — finansman maliyeti teminat tutarını aşıyor; teminatlı yol tercih edilir.",
        "BEKLE — azami iade düşük veya finansman/azami oranı yüksek.",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 10, 1, "Belirsizlik beyanları (M05)", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari(["inceleme_sure_yorumu"]), 11):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 13, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2,
      "İnceleme süresi girdideki ay değeriyle aynen kullanılır. "
      "Finansman = azami × yıllık oran × (ay/12).",
      kaydir=True, yazi="333333")
    ws.merge_cells("B13:J13")
    h(ws, 14, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 2,
      "Yorum B, inceleme süresini +%20 uzatır (daha yüksek finansman maliyeti). "
      "Dosya her iki yorumun sonucunu MOTOR'da üretir.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B14:J14")

    h(ws, 16, 1, "Kullanım sırası", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 1,
      "GIRDI → KURALLAR → MOTOR → KARTLAR/AKIS/KUYRUKLAR → SENARYO → VAKALAR → PANO → KANIT_RAPORU → AYARLAR.",
      kaydir=True, yazi="333333")
    ws.merge_cells("A17:J17")
    h(ws, 19, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv | Koruma: {SIFRE}", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 18, "B": 70})


def kosullu_bicimlendirme(wb):
    yesil = Font(color=NORMAL, bold=True)
    kirmizi = Font(color=KRITIK, bold=True)
    amber = Font(color="B7791F", bold=True)
    y_fill = PatternFill("solid", fgColor="E2EFDA")
    k_fill = PatternFill("solid", fgColor="FDE9E9")
    a_fill = PatternFill("solid", fgColor="FFF2CC")

    ws = wb["PANO"]
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("VERİ YOK",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("BEKLE",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("TEMİNATLI",B6))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['AND(ISNUMBER(SEARCH("BAŞVUR",B6)),NOT(ISNUMBER(SEARCH("TEMİNATLI",B6))))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("E4", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("E4", CellIsRule(operator="equal", formula=["0"], font=amber))
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
        ws.conditional_formatting.add(f"{col}10:{col}12", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(f"{col}10:{col}12", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("F25:F35", FormulaRule(formula=['F25="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F25:F35", FormulaRule(formula=['F25="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F6:F9", CellIsRule(operator="notEqual", formula=["0"], font=amber))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("BEKLE",E6))'], font=kirmizi))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("TEMİNATLI",E6))'], font=amber))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['AND(ISNUMBER(SEARCH("BAŞVUR",E6)),NOT(ISNUMBER(SEARCH("TEMİNATLI",E6))))'], font=yesil))
    ws.conditional_formatting.add("B16:B18", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add("B5:B11", FormulaRule(formula=['OR(B5="KIRIK",B5="HATALI",B5="NEGATİF",B5="BOŞ",B5="EKSİK",B5="AŞIYOR")'],
                                                        font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B5:B11", FormulaRule(formula=['OR(B5="TEMİZ",B5="DOLU",B5="NORMAL",B5="TUTARLI",B5="TAM")'],
                                                        font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B12", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("B12", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("G7:G53", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("F7:F53", FormulaRule(formula=['F7="KIS-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F53", FormulaRule(formula=['F7="KIS-005"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "C", "D"):
        ws.conditional_formatting.add(f"{harf}6:{harf}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("J6:J20", FormulaRule(formula=['J6="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("J6:J20", FormulaRule(formula=['J6="TAM"'], font=yesil, fill=y_fill))

    ws = wb["AKIS"]
    for col, val in [("M", "YETİM"), ("N", "GEÇERSİZ GEÇİŞ"), ("O", "ZİNCİR KIRIK")]:
        ws.conditional_formatting.add(
            f"{col}{AKIS_ILK}:{col}{AKIS_SON}",
            CellIsRule(operator="equal", formula=[f'"{val}"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add(
        f"O{AKIS_ILK}:O{AKIS_SON}",
        CellIsRule(operator="equal", formula=['"OK"'], font=yesil, fill=y_fill))
    for durum, fill in [("BASVURU", a_fill), ("INCELME", a_fill), ("EK_BELGE", a_fill),
                        ("ODEME", y_fill), ("IPTAL", k_fill)]:
        ws.conditional_formatting.add(
            f"E{AKIS_ILK}:E{AKIS_SON}",
            CellIsRule(operator="equal", formula=[f'"{durum}"'], fill=fill))

    ws = wb["KUYRUKLAR"]
    ws.conditional_formatting.add("B6:B8", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("C6:C8", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("B14:B16", CellIsRule(operator="greaterThan", formula=["0"], font=kirmizi))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B21", FormulaRule(formula=['ISNUMBER(SEARCH("BEKLE",B21))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B21", FormulaRule(formula=['ISNUMBER(SEARCH("TEMİNATLI",B21))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B21", FormulaRule(formula=['AND(ISNUMBER(SEARCH("BAŞVUR",B21)),NOT(ISNUMBER(SEARCH("TEMİNATLI",B21))))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B20", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))


def ad_tanimlari(wb):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "zincirTolerans", "AYARLAR!$B$39")
    ad_ekle(wb, "kis_azamiIade", "MOTOR!$G$29")
    ad_ekle(wb, "kis_netNakit", "MOTOR!$G$28")
    ad_ekle(wb, "kis_teminat", "MOTOR!$G$30")
    ad_ekle(wb, "kis_kacirilan", "MOTOR!$G$28")
    ad_ekle(wb, "kis_kararMetni", "MOTOR!$G$34")
    ad_ekle(wb, "kis_kararGerekce", "MOTOR!$G$35")
    ad_ekle(wb, "kis_maddeAtifMetni", "MOTOR!$G$52")
    ad_ekle(wb, "kis_kanitRaporu", "MOTOR!$G$53")
    ad_ekle(wb, "kis_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "kis_parmakIzi", "AYARLAR!$B$20")

    ad_ekle(wb, "kis_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "kis_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "kis_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "kis_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "kis_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "kis_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "kis_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "kis_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "kis_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "kis_vakaFark", "VAKALAR!$B$13")
    ad_ekle(wb, "kis_sonrakiKimlik", "KARTLAR!$B$4")

    ayar_map = {
        "kis_tolerans": 8,
        "kis_kacirilanEsik": 9,
        "kis_senaryoIyi": 10,
        "kis_senaryoKotu": 11,
        "kis_teminatArtisPuan": 12,
        "kis_tahminAlt": 13,
        "kis_tahminUst": 14,
        "kis_tornadoYuklenilen": 15,
        "kis_tornadoHasilat": 16,
        "kis_tornadoSure": 17,
        "kis_sureArtisOran": 18,
        "kis_yuzdelikOran": 19,
        "kis_girisBeklenen": 22,
        "kis_riskDusuk": 23,
        "kis_riskOrta": 24,
        "kis_riskYuksek": 25,
        "kis_riskEsikOran": 26,
        "kis_yilTaban": 27,
        "kis_yuvarMax": 28,
        "kis_dayanakCarpan": 29,
        "kis_bazCarpan": 30,
        "kis_sifirCarpan": 31,
        "kis_siraBir": 32,
        "kis_siraIki": 33,
        "kis_siraUc": 34,
        "kis_tahminX": 35,
        "vaka_hasilat_1": 36,
        "vaka_devreden_1": 37,
        "vaka_hasilat_2": 38,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "kis_modulT1": (11, "C"), "kis_modulT1Yorum": (11, "D"),
        "kis_modulT2": (12, "C"), "kis_modulT2Yorum": (12, "D"),
        "kis_modulO1": (13, "C"), "kis_modulO1Yorum": (13, "D"),
        "kis_modulO6": (16, "C"), "kis_modulO6Yorum": (16, "D"),
        "kis_modulO8": (17, "C"), "kis_modulO8Yorum": (17, "D"),
        "kis_modulI2": (20, "C"), "kis_modulI2Yorum": (20, "D"),
    }
    for ad, (satir, kol) in modul_map.items():
        ad_ekle(wb, ad, f"PANO!${kol}${satir}")

    ad_ekle(wb, "ListeYillar", "LISTELER!$A$6:$A$8")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$C$6:$C$7")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$6:$E$7")
    ad_ekle(wb, "ListeIadeTur", "LISTELER!$G$6:$G$7")
    ad_ekle(wb, "ListeDurumlar", "LISTELER!$I$6:$I$10")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$K$6:$K$14")


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
        ("KARTLAR", kartlar),
        ("AKIS", akis),
        ("KUYRUKLAR", kuyruklar),
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
    varsayilan = os.path.join(kok, "KdvIadesiTutarSurecSimulasyonu.xlsx")
    hedef = cikti_yolu or varsayilan
    os.makedirs(os.path.dirname(os.path.abspath(hedef)) or ".", exist_ok=True)
    wb.save(hedef)

    hsh = hashlib.sha256()
    with open(hedef, "rb") as f:
        for b in iter(lambda: f.read(65536), b""):
            hsh.update(b)
    dig = hsh.hexdigest()

    wb2 = __import__("openpyxl").load_workbook(hedef)
    wb2["AYARLAR"]["B20"] = dig[:16]
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
    cikti = os.path.join(repo, "cikti", "KdvIadesiTutarSurecSimulasyonu.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
