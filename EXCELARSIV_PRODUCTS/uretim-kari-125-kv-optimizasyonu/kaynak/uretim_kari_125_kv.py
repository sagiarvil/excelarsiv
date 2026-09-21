#!/usr/bin/env python3
"""
Üretim Kârı %12,5 KV Optimizasyonu — üretim betiği (A1 / manda v6).
KVK md.32 indirimli oran senaryosu: üretim kârı matrahı × %12,5 vs standart KV.
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

URUN_AD = "Üretim Kârı %12,5 KV Optimizasyonu"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 11)

# Demo 2026: hasılat 40M − maliyet 28M − GY 1.5M − FG 0.5M = 10M üretim kârı
# + diğer 0.2M − istisna 0 → matrah 10.2M
# standart %25 = 2.55M; indirimli %12,5 = 1.275M; tasarruf 1.275M → UYGULA
DEMO = {
    "sirket_adi": "Örnek İmalat A.Ş.",
    "donem": "2026",
    "hesapYili": 2026,
    "uretim_hasilati": 40_000_000,
    "uretim_maliyeti": 28_000_000,
    "genel_yonetim_payi": 1_500_000,
    "finansman_gider_payi": 500_000,
    "diger_faaliyet_kar": 200_000,
    "istisna_tutari": 0,
    "yorum_modu": "A",
}


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: Üretim KV tasarrufu ve kararı etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Matrah, tasarruf ve karar bozulur | Kaynak: KVK md.32",
    )


def kural_cek(kural_id: str) -> str:
    """Oranları tblKurallar'dan çeker; yıl seçimi ukv_yilTaban ile (G07: adında rakam yok)."""
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-ukv_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(ukv_yuvarMax,IFERROR(N(G12),0)))"


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
      "Üretim hasılatı, maliyeti ve payları girin; standart KV ile indirimli %12,5 "
      "oranı karşılaştırıp UYGULA / İNCELE / UYGULAMA kararını üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Çok yıllı kural tablosu (2025/2026/2027) ile indirimli üretim KV oranına hazır motor",
        "Üretim kârı, matrah, standart KV, indirimli KV ve tasarrufun yan yana hesabı",
        "Baz / iyimser / kötümser + kural değişimi (indirimli oran +1 puan) senaryoları",
        "Tebliğ tarzı altın vakalar ve KANIT_RAPORU (A4 imza)",
        "Belirsizlik beyanı: genel yönetim / finansman pay dağıtım yorumu (A / B)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "SMMM — indirimli KV tasarrufunu ve madde atfını müşteriye vermek",
        "CFO — UYGULA / İNCELE / UYGULAMA kararını bütçeye yazmak",
        "Ortak / denetçi — imzalı kanıt raporunu dosyalamak",
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
        ("Adım 1 — Girdiler", "GIRDI sayfasında yıl, hasılat, maliyet ve payları sarı hücrelere girin."),
        ("Adım 2 — Karar", "PANO'da tasarruf tutarını ve UYGULA / İNCELE / UYGULAMA rozetini izleyin."),
        ("Adım 3 — Kanıt", "KANIT_RAPORU'nu yazdırıp imzalayın; VAKALAR tutarlılığını kontrol edin."),
    ]
    for i, (b, m) in enumerate(adimlar, 5):
        h(ws, i, 1, b, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        _kim(ws, f"A{i}", b, "Sırayla izleyin")
    h(ws, 9, 1, "Sık yapılan hatalar", kalin=True, yazi=KRITIK)
    for i, m in enumerate([
        "Genel yönetim / finansman payını üretim dışı bırakmak",
        "İstisna tutarını matrahtan düşmeden standart oranla karşılaştırmak",
        "Hesap yılını kural tablosuyla uyumsuz seçmek (2025 ≠ 2026 oranı)",
    ], 10):
        h(ws, i, 1, "• " + m, yazi=KRITIK)
    h(ws, 14, 1,
      "Bu dosya karar destek aracıdır; kesin vergi görüşü veya beyanname yerine geçmez.",
      yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 72, "B": 70})
    for r in (10, 11, 12):
        h(ws, r, 1, ws.cell(r, 1).value, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Şirket ve üretim girdileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Oranlar KURALLAR'dan; tasarruf formülle üretilir.",
      yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Şirket Adı", DEMO["sirket_adi"], None, "sirket_adi",
         "Şirket unvanını yazın."),
        (7, "Dönem", DEMO["donem"], None, "donem",
         "Beyan dönemi veya hesap yılı referansını yazın."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili",
         "2025, 2026 veya 2027 seçin."),
        (9, "Üretim Hasılatı [₺]", DEMO["uretim_hasilati"], TL, "uretim_hasilati",
         "Üretim/imalat faaliyetinden doğan hasılatı TL girin."),
        (10, "Üretim Maliyeti [₺]", DEMO["uretim_maliyeti"], TL, "uretim_maliyeti",
         "Satılan mamul / üretim maliyetini TL girin."),
        (11, "Genel Yönetim Payı [₺]", DEMO["genel_yonetim_payi"], TL, "genel_yonetim_payi",
         "Üretime düşen genel yönetim gider payını TL girin."),
        (12, "Finansman Gider Payı [₺]", DEMO["finansman_gider_payi"], TL, "finansman_gider_payi",
         "Üretime düşen finansman gider payını TL girin."),
        (13, "Diğer Faaliyet Kârı [₺]", DEMO["diger_faaliyet_kar"], TL, "diger_faaliyet_kar",
         "Üretim dışı faaliyet kârını (matraha ek) TL girin; zarar negatif."),
        (14, "İstisna Tutarı [₺]", DEMO["istisna_tutari"], TL, "istisna_tutari",
         "Matrahtan düşülecek istisna tutarını TL girin."),
        (15, "Yorum Modu (A/B)", DEMO["yorum_modu"], None, "yorum_modu",
         "Pay dağıtım yorumu A (standart) veya B (sıkı +%10 pay) seçin."),
    ]
    h(ws, 5, 1, "Alan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 2, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 3, "Birim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    for satir, etiket, deger, sayi, _ad, mesaj in alanlar:
        birim = "₺" if sayi == TL else ("yıl" if _ad == "hesapYili" else "metin")
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim, yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 17, 1, "İndirimli oran (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 2, kural_cek("UKV-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 1, "Standart oran (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 2, kural_cek("UKV-002"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Tasarruf (ayna) [₺]", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2, "=IFERROR(ukv_tasarruf,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B15)=0,ukv_girisBeklenen=0),0,ROUND(COUNTA(B6:B15)/ukv_girisBeklenen*100,0)),0)',
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
        ("sirket_adi", DEMO["sirket_adi"], "metin", "Evet", "GIRDI"),
        ("donem", DEMO["donem"], "metin", "Evet", "GIRDI"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Evet", "GIRDI"),
        ("uretim_hasilati", DEMO["uretim_hasilati"], "TL", "Evet", "GIRDI"),
        ("uretim_maliyeti", DEMO["uretim_maliyeti"], "TL", "Evet", "GIRDI"),
        ("genel_yonetim_payi", DEMO["genel_yonetim_payi"], "TL", "Evet", "GIRDI"),
        ("finansman_gider_payi", DEMO["finansman_gider_payi"], "TL", "Evet", "GIRDI"),
        ("diger_faaliyet_kar", DEMO["diger_faaliyet_kar"], "TL", "Evet", "GIRDI"),
        ("istisna_tutari", DEMO["istisna_tutari"], "TL", "Evet", "GIRDI"),
        ("yorum_modu", DEMO["yorum_modu"], "A/B", "Evet", "GIRDI"),
    ]
    tl_alanlar = {
        "uretim_hasilati", "uretim_maliyeti", "genel_yonetim_payi",
        "finansman_gider_payi", "diger_faaliyet_kar", "istisna_tutari",
    }
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 24):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in tl_alanlar:
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
    dogrulama(ws, "list", "ListeYorumAB", "B15",
              baslik="Yorum", mesaj="Yorum modunda A veya B seçin.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B seçin.", bos=False)
    dogrulama(ws, "list", "ListeYillar", "B26",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeYorumAB", "B33",
              baslik="Yorum", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.")
    for aralik, baslik, mesaj in [
        ("B9", "Üretim Hasılatı", "0 ile 1e12 arasında TL."),
        ("B10", "Üretim Maliyeti", "0 ile 1e12 arasında TL."),
        ("B11", "Genel Yönetim Payı", "0 ile 1e12 arasında TL."),
        ("B12", "Finansman Gider Payı", "0 ile 1e12 arasında TL."),
        ("B14", "İstisna Tutarı", "0 ile 1e12 arasında TL."),
        ("B27", "Üretim Hasılatı", "Tablo: TL tutar."),
        ("B28", "Üretim Maliyeti", "Tablo: TL tutar."),
        ("B29", "Genel Yönetim", "Tablo: TL tutar."),
        ("B30", "Finansman Payı", "Tablo: TL tutar."),
        ("B32", "İstisna", "Tablo: TL tutar."),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj=mesaj,
                  hata_baslik="Geçersiz tutar", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000000")
    dogrulama(ws, "decimal", "-1000000000000", "B13",
              baslik="Diğer Faaliyet", mesaj="-1e12 ile 1e12 arasında TL.",
              hata_baslik="Geçersiz", hata_mesaj="Sınır dışı.",
              isaret="between", f2="1000000000000")
    dogrulama(ws, "decimal", "-1000000000000", "B31",
              baslik="Diğer Faaliyet", mesaj="Tablo: diğer faaliyet.",
              hata_baslik="Geçersiz", hata_mesaj="Sınır dışı.",
              isaret="between", f2="1000000000000")

    giris_hucreleri(ws, 6, 15, [2])
    giris_hucreleri(ws, 24, 33, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 40, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1024, "Sarı alanlar giriş; oran ve tasarruf formülleri kilitlidir.")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Mevzuat kural tablosu (yıl yan yana)", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    # UKV-001: 2025=%20 (geçiş), 2026/2027=%12,5 — ortak/parametreler.yaml ile uyumlu
    satirlar = [
        ["UKV-001", "KVK md.32", "indirimli_oran", "üretim kazancı", "uretim_kv_orani",
         0.20, 0.125, 0.125, "01.01.2025", "Üretim faaliyetleri indirimli KV"],
        ["UKV-002", "KVK md.32", "standart_oran", "genel kurum kazancı", "kv_genel_oran",
         0.25, 0.25, 0.25, "01.01.2025", "Genel kurumlar vergisi oranı"],
        ["UKV-003", "İç politika", "uygula_esik", "tasarruf ≥ eşik", "uygula_esik_tl",
         100000, 100000, 100000, "01.01.2025", "UYGULA karar eşiği (TL)"],
        ["UKV-004", "İç politika", "incele_esik", "tasarruf ≥ eşik", "incele_esik_tl",
         25000, 25000, 25000, "01.01.2025", "İNCELE karar eşiği (TL)"],
        ["UKV-005", "Genel", "yuvarlama", "her hesap", "yuvarlama_ondalik",
         2, 2, 2, "01.01.2025", "Uygulama notu"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = None
            if k in (6, 7, 8):
                sayi = YÜZDE if s[0] in ("UKV-001", "UKV-002") else CATI
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-ukv_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "İndirimli oran (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("UKV-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Standart oran (aktif yıl)", yazi="333333")
    h(ws, 16, 2, kural_cek("UKV-002"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"UKV-001=",TEXT(B15,"0.0%"),"UKV-002=",TEXT(B16,"0.0%"),'
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

    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif mevzuat yılı", "UKV-005")  # G7
    ekle("İndirimli oran", km("UKV-001"), "oran", "Üretim KV indirimli oran", "UKV-001")  # G8
    ekle("Standart oran", km("UKV-002"), "oran", "Genel KV oranı", "UKV-002")  # G9
    ekle("Uygula eşiği", km("UKV-003"), "TL", "UYGULA tasarruf eşiği", "UKV-003")  # G10
    ekle("İncele eşiği", km("UKV-004"), "TL", "İNCELE tasarruf eşiği", "UKV-004")  # G11
    ekle("Yuvarlama ondalık", km("UKV-005"), "adet", "Yuvarlama basamağı", "UKV-005")  # G12
    ekle("Üretim hasılatı", "=GIRDI!B9", "TL", "Üretim hasılatı", "UKV-001")  # G13
    ekle("Üretim maliyeti", "=GIRDI!B10", "TL", "Üretim maliyeti", "UKV-001")  # G14
    ekle("Genel yönetim payı", "=GIRDI!B11", "TL", "GY payı", "UKV-001")  # G15
    ekle("Finansman gider payı", "=GIRDI!B12", "TL", "FG payı", "UKV-001")  # G16
    ekle("Diğer faaliyet kârı", "=GIRDI!B13", "TL", "Diğer kâr/zarar", "UKV-001")  # G17
    ekle("İstisna tutarı", "=GIRDI!B14", "TL", "İstisna", "UKV-001")  # G18
    ekle("Pay çarpanı (yorum)",
         '=IF(GIRDI!B15="B",1+ukv_payArtisOran,1)',
         "çarpan", "Yorum A/B pay çarpanı", "UKV-005")  # G19
    ekle("GY efektif", "=G15*G19", "TL", "Yorum sonrası GY", "UKV-001")  # G20
    ekle("FG efektif", "=G16*G19", "TL", "Yorum sonrası FG", "UKV-001")  # G21
    ekle("Üretim kârı", "=G13-G14-G20-G21", "TL", "Hasılat−maliyet−paylar", "UKV-001")  # G22
    ekle("Matrah", f"=ROUND(MAX(0,G22+G17-G18),{_YV})", "TL", "Vergiye tabi matrah", "UKV-001")  # G23
    ekle("Standart KV", f"=ROUND(G23*G9,{_YV})", "TL", "Matrah × standart oran", "UKV-002")  # G24
    ekle("İndirimli KV", f"=ROUND(G23*G8,{_YV})", "TL", "Matrah × indirimli oran", "UKV-001")  # G25
    ekle("Tasarruf", "=G24-G25", "TL", "Standart − indirimli", "UKV-001")  # G26
    ekle("Oran farkı", "=G9-G8", "oran", "Standart − indirimli puan", "UKV-001")  # G27
    ekle("Tasarruf (pano)", "=G26", "TL", "Tasarruf ayna", "UKV-001")  # G28
    ekle("Matrah (pano)", "=G23", "TL", "Matrah ayna", "UKV-001")  # G29
    ekle("İndirimli KV (pano)", "=G25", "TL", "İndirimli ayna", "UKV-001")  # G30
    ekle("Üretim kâr / hasılat", "=IF(G13=0,0,G22/G13)", "oran", "Marj göstergesi", "UKV-001")  # G31
    ekle("İstisna / matrah brüt", "=IF(G22+G17=0,0,G18/MAX(1,G22+G17))", "oran", "İstisna yoğunluğu", "UKV-001")  # G32
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(G26>=G10,"UYGULA",'
         'IF(G26>=G11,"INCELE","UYGULAMA")))',
         "kod", "Karar kodu", "UKV-005")  # G33
    ekle("Karar metni",
         '=IF(G33="VERI_YOK","VERİ YOK",'
         'IF(G33="UYGULA","UYGULA",'
         'IF(G33="INCELE","İNCELE","UYGULAMA")))',
         "metin", "Kullanıcıya karar", "UKV-005")  # G34
    ekle("Gerekçe",
         '=IF(G33="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G33="UYGULA","Tasarruf uygula eşiğinin üzerinde; indirimli üretim KV oranını uygulayın ve kanıt dosyalayın.",'
         'IF(G33="INCELE","Tasarruf incele bandında; pay dağıtımı ve istisna dayanağını gözden geçirin.",'
         '"Tasarruf eşik altında veya matrah yok; indirimli oran uygulaması öncelikli değil.")))',
         "metin", "Gerekçe cümlesi", "UKV-005")  # G35
    ekle("Tasarruf / standart KV", "=IF(G24=0,0,G26/G24)", "oran", "Tasarruf oranı", "UKV-002")  # G36
    ekle("Efektif oran", "=IF(G23=0,0,G25/G23)", "oran", "Ödenen efektif", "UKV-001")  # G37
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B15)=0,ukv_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B15)/ukv_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "UKV-005")  # G38
    ekle("Risk skoru",
         '=IF(G33="VERI_YOK",0,IF(G33="UYGULA",ukv_riskDusuk,'
         'IF(G33="INCELE",ukv_riskOrta,ukv_riskYuksek)))',
         "puan", "Kaçırma riski (ters)", "UKV-005")  # G39
    ekle("Senaryo iyimser matrah", "=G23*ukv_senaryoIyi", "TL", "İyimser matrah", "UKV-001")  # G40
    ekle("Senaryo kötümser matrah", "=G23*ukv_senaryoKotu", "TL", "Kötümser matrah", "UKV-001")  # G41
    ekle("İyimser tasarruf", f"=ROUND(G40*(G9-G8),{_YV})", "TL", "İyimser tasarruf", "UKV-001")  # G42
    ekle("Kötümser tasarruf", f"=ROUND(G41*(G9-G8),{_YV})", "TL", "Kötümser tasarruf", "UKV-001")  # G43
    ekle("M-SEN indirimli+", "=MIN(G9,G8+ukv_oranArtisPuan)", "oran", "İndirimli + delta", "UKV-001")  # G44
    ekle("M-SEN indirimli KV", f"=ROUND(G23*G44,{_YV})", "TL", "M-SEN indirimli KV", "UKV-001")  # G45
    ekle("M-SEN tasarruf", "=G24-G45", "TL", "M-SEN tasarruf farkı", "UKV-001")  # G46
    ekle("Tahmin üst", "=G26*ukv_tahminUst", "TL", "Tahmin üst bant", "UKV-005")  # G47
    ekle("Tahmin alt", "=G26*ukv_tahminAlt", "TL", "Tahmin alt bant", "UKV-005")  # G48
    ekle("Tornado: hasılat etkisi",
         f"=ABS(ROUND(MAX(0,(G13*ukv_tornadoHasilat)-G14-G20-G21+G17-G18)*(G9-G8),{_YV})-G26)",
         "TL", "Hasılat ± etki", "UKV-001")  # G49
    ekle("Tornado: maliyet etkisi",
         f"=ABS(ROUND(MAX(0,G13-(G14*ukv_tornadoMaliyet)-G20-G21+G17-G18)*(G9-G8),{_YV})-G26)",
         "TL", "Maliyet ± etki", "UKV-001")  # G50
    ekle("Tornado: oran etkisi",
         f"=ABS(ROUND(G23*((G9-G8)+ukv_oranArtisPuan),{_YV})-G26)",
         "TL", "Oran ± etki", "UKV-001")  # G51
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"UKV-001→indirimli oran","UKV-002→standart oran",'
         '"UKV-003→uygula eşiği","UKV-004→incele eşiği")',
         "metin", "Kanıt atıfları", "UKV-001")  # G52
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Matrah=",TEXT(G23,"₺ #,##0"),"StandartKV=",TEXT(G24,"₺ #,##0"),'
         '"İndirimliKV=",TEXT(G25,"₺ #,##0"),"Tasarruf=",TEXT(G26,"₺ #,##0"))',
         "metin", "Rapor özeti", "UKV-005")  # G53

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
    h(ws, 5, 2, "Matrah Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Matrah", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Tasarruf", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(ukv_senaryoIyi+N(GIRDI!B9)*ukv_sifirCarpan,0)",
         "=IFERROR(MOTOR!G40,0)",
         "=IFERROR(MOTOR!G42,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D6>=MOTOR!G10,"UYGULA",IF(D6>=MOTOR!G11,"İNCELE","UYGULAMA")))'),
        ("Baz", "=IFERROR(ukv_bazCarpan+N(GIRDI!B9)*ukv_sifirCarpan,ukv_bazCarpan)",
         "=IFERROR(ukv_matrah,0)",
         "=IFERROR(ukv_tasarruf,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(ukv_kararMetni,"-"))'),
        ("Kötümser", "=IFERROR(ukv_senaryoKotu+N(GIRDI!B9)*ukv_sifirCarpan,0)",
         "=IFERROR(MOTOR!G41,0)",
         "=IFERROR(MOTOR!G43,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D8>=MOTOR!G10,"UYGULA",IF(D8>=MOTOR!G11,"İNCELE","UYGULAMA")))'),
        ("M-SEN oran+1", "=IFERROR(ukv_bazCarpan+N(GIRDI!B9)*ukv_sifirCarpan,ukv_bazCarpan)",
         "=IFERROR(ukv_matrah,0)",
         "=IFERROR(MOTOR!G46,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D9>=MOTOR!G10,"UYGULA",IF(D9>=MOTOR!G11,"İNCELE","UYGULAMA")))'),
    ]
    for i, (ad, carp, mat, tas, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, mat, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, tas, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant genişliği (iyimser−kötümser tasarruf)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(D6-D8,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser tasarruf ",TEXT(D6,"₺ #,##0")," · Baz ",TEXT(D7,"₺ #,##0"),'
      '" · Kötümser ",TEXT(D8,"₺ #,##0")," · Bant ",TEXT(B11,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:H12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Tornado hasılat etkisi", yazi="333333")
    h(ws, 16, 2, "=IFERROR(MOTOR!G49,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado maliyet etkisi", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G50,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado oran etkisi", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G51,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 3,
      '=IFERROR(IF(B16>=MAX(B16:B18),ukv_siraBir,IF(B16>=MIN(B16:B18)+ABS(B16-B17),ukv_siraIki,ukv_siraUc)),ukv_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 3,
      '=IFERROR(IF(B17>=MAX(B16:B18),ukv_siraBir,IF(B17=MEDIAN(B16:B18),ukv_siraIki,ukv_siraUc)),ukv_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 3,
      '=IFERROR(IF(B18>=MAX(B16:B18),ukv_siraBir,IF(B18=MEDIAN(B16:B18),ukv_siraIki,ukv_siraUc)),ukv_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(CONCATENATE("1)",INDEX(A16:A18,MATCH(ukv_siraBir,C16:C18,0)),'
      '" 2)",INDEX(A16:A18,MATCH(ukv_siraIki,C16:C18,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B16:B18),"₺ #,##0")," TL — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN kural değişimi tasarruf", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=D9", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"İndirimli oran +1 puan senaryosunda tasarruf ",TEXT(D9,"₺ #,##0")," TL")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı (tasarruf)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=MOTOR!G48", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 3, "=IFERROR(MOTOR!G26,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 4, "=MOTOR!G47", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Tahmin (AVERAGE)", kalin=True, yazi=KOYU_LACIVERT)
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

    h(ws, 5, 7, "TasarrufDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([1_402_500, 1_275_000, 1_083_750, 1_147_500], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([250000, 180000, 127500], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)

    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([1_402_500, 1_275_000, 1_083_750, 1_147_500], 28):
        h(ws, i, 4, v, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo Tasarruf Karşılaştırması"
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
    g3.title = "Senaryo Tasarruf Trendi"
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

    # V-001: 1_000_000 * 0.125 = 125_000 (2026 indirimli)
    # V-002: 1_000_000 * 0.25 = 250_000
    # V-003: 250_000 - 125_000 = 125_000
    # V-004: 200_000 * 0.125 = 25_000
    vakalar_data = [
        ("V-001", "Matrah × indirimli oran",
         "matrah=vaka_matrah_1; oran=UKV-001(2026)",
         125000,
         "=IFERROR(ROUND(vaka_matrah_1*INDEX(tblKurallar[deger_2026],MATCH(\"UKV-001\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=ukv_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "KVK md.32 indirimli"),
        ("V-002", "Matrah × standart oran",
         "matrah=vaka_matrah_1; oran=UKV-002",
         250000,
         "=IFERROR(ROUND(vaka_matrah_1*INDEX(tblKurallar[deger_2025],MATCH(\"UKV-002\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=ukv_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "KVK md.32 standart"),
        ("V-003", "Tasarruf = standart − indirimli",
         "V-002 − V-001",
         125000,
         "=IFERROR(ROUND(vaka_matrah_1*INDEX(tblKurallar[deger_2025],MATCH(\"UKV-002\",tblKurallar[kural_id],0))"
         "-vaka_matrah_1*INDEX(tblKurallar[deger_2026],MATCH(\"UKV-001\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=ukv_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Tasarruf farkı"),
        ("V-004", "Küçük matrah × indirimli",
         "matrah=vaka_matrah_2; oran=UKV-001(2026)",
         25000,
         "=IFERROR(ROUND(vaka_matrah_2*INDEX(tblKurallar[deger_2026],MATCH(\"UKV-001\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=ukv_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "UKV-001 küçük matrah"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v,
              yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = TL

    formuller = {
        "fark": "=IF(tblVakalar[[#This Row],[vaka_id]]=\"\",\"\",IFERROR(tblVakalar[[#This Row],[beklenen_sonuc]]-tblVakalar[[#This Row],[hesaplanan]],0))",
        "durum": '=IF(tblVakalar[[#This Row],[vaka_id]]="","",IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=ukv_tolerans,"TUTARLI","KIRIK"),"KIRIK"))',
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
    for i, (b, he) in enumerate([(125000, 125000), (250000, 250000), (125000, 125000), (25000, 25000)], 6):
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
    h(ws, 3, 1, "KANIT RAPORU — Üretim Kârı %12,5 KV", kalin=True, boyut=14, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:F3")
    h(ws, 4, 1, "Rapor tarihi", yazi="333333")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 5, 1, "Şirket", yazi="333333")
    h(ws, 5, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Dönem / yıl", yazi="333333")
    h(ws, 6, 2, '=CONCATENATE(GIRDI!B7," / ",hesapYili)', yazi=DIS_REF_YEŞIL)

    h(ws, 8, 1, "Girdi özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, (et, form, sayi) in enumerate([
        ("Üretim hasılatı", "=GIRDI!B9", TL),
        ("Üretim maliyeti", "=GIRDI!B10", TL),
        ("Genel yönetim payı", "=GIRDI!B11", TL),
        ("Finansman gider payı", "=GIRDI!B12", TL),
        ("Diğer faaliyet kârı", "=GIRDI!B13", TL),
        ("İstisna tutarı", "=GIRDI!B14", TL),
    ], 9):
        h(ws, i, 1, et, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)

    h(ws, 16, 1, "Hesap zinciri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 17, 1, "Matrah", yazi="333333")
    h(ws, 17, 2, "=ukv_matrah", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 1, "Standart KV", yazi="333333")
    h(ws, 18, 2, "=MOTOR!G24", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "İndirimli KV", yazi="333333")
    h(ws, 19, 2, "=MOTOR!G25", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Tasarruf", yazi="333333")
    h(ws, 20, 2, "=ukv_tasarruf", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 21, 1, "KARAR", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 2, "=ukv_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 22, 1, "Gerekçe", yazi="333333")
    h(ws, 22, 2, "=ukv_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B22:F22")

    h(ws, 24, 1, "Madde atıfları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, "=ukv_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B24:F24")

    h(ws, 26, 1, "Kanıt gövdesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 2, "=ukv_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B26:F26")

    h(ws, 28, 1, "Parmak izi (SHA-256 kısaltma)", yazi="333333")
    h(ws, 28, 2, "=ukv_parmakIzi", yazi=GRİ)

    h(ws, 30, 1, "Hazırlayan (imza)", yazi="333333")
    h(ws, 30, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B30", "İmza", "Raporu hazırlayan kişinin adını yazın.")
    h(ws, 31, 1, "Onaylayan (imza)", yazi="333333")
    h(ws, 31, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B31", "Onay", "Onaylayan kişinin adını yazın.")

    h(ws, 33, 1,
      "Uyarı: Bu çıktı karar destektir; kesin vergi sonucu veya hukuki görüş değildir. "
      f"Sürüm {SURUM}.",
      yazi=GRİ, boyut=9, kaydir=True)
    ws.merge_cells("A33:F33")

    baski_hazirla(ws, "A1:F34", f"{URUN_AD} · {SURUM}")
    ws.page_setup.orientation = "portrait"
    genislik(ws, {"A": 28, "B": 22, "C": 14, "D": 14, "E": 14, "F": 14})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim paneli — Üretim KV %12,5", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpiler = [
        (2, "Matrah", "=IFERROR(ukv_matrah,0)", TL),
        (3, "StandartKV", "=IFERROR(MOTOR!G24,0)", TL),
        (4, "İndirimliKV", "=IFERROR(MOTOR!G25,0)", TL),
        (5, "Tasarruf", "=IFERROR(ukv_tasarruf,0)", TL),
        (6, "Güven", "=IFERROR(MOTOR!G38,0)", CATI),
        (7, "Risk", "=IFERROR(MOTOR!G39,0)", CATI),
        (8, "İnd.Oran", "=IFERROR(MOTOR!G8,0)", YÜZDE),
        (9, "Std.Oran", "=IFERROR(MOTOR!G9,0)", YÜZDE),
        (10, "Marj", "=IFERROR(MOTOR!G31,0)", YÜZDE),
        (11, "M-SEN", "=IFERROR(ukv_kuralDegisimSenaryo,0)", TL),
        (12, "Tahmin", "=IFERROR(SENARYO!B32,0)", TL),
        (13, "Vaka", "=IFERROR(ukv_vakaDurum,\"-\")", None),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=9, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=ukv_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=ukv_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 10, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Üretim kârı / hasılat",
         "=IFERROR(MOTOR!G31,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Marj ",TEXT(C11,"0.0%")," — üretim kârı / hasılat")'),
        ("T2", "Tasarruf / standart KV",
         "=IFERROR(MOTOR!G36,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Tasarruf oranı ",TEXT(C12,"0.0%"))'),
        ("O1", "Senaryo tasarruf sapması",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,STDEV.P(SENARYO!D6:D9)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo sapması ",TEXT(C13,"₺ #,##0")," TL")'),
        ("O2", "Tornado max etki",
         "=IFERROR(ukv_duyarlilikSira,\"-\")",
         "=IFERROR(ukv_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo bant",
         "=IFERROR(ukv_senaryoKarsilastirma,0)",
         "=IFERROR(ukv_yorumSenaryo,\"-\")"),
        ("O6", "İstisna / matrah brüt",
         "=IFERROR(MOTOR!G32,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İstisna yoğunluğu ",TEXT(C16,"0.0%"))'),
        ("O8", "Kalite skoru",
         "=IFERROR(KONTROLLER!B12,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalite ",TEXT(C17,"0")," puan")'),
        ("M", "M-SEN tasarruf",
         "=IFERROR(ukv_kuralDegisimSenaryo,0)",
         "=IFERROR(ukv_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin aralık",
         "=IFERROR(ukv_tahminAralik,0)",
         "=IFERROR(ukv_yorumTahmin,\"-\")"),
        ("I2", "Senaryo tasarruf P90",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!D6:D9,ukv_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"P90 tasarruf ",TEXT(C20,"₺ #,##0")," TL")'),
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
        ("Üretim hasılatı", 40_000_000),
        ("Üretim maliyeti", 28_000_000),
        ("Üretim kârı", 10_000_000),
        ("Matrah", 10_200_000),
        ("Tasarruf", 1_275_000),
    ], 24):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)

    h(ws, 23, 4, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 5, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Standart KV", 2_550_000),
        ("İndirimli KV", 1_275_000),
        ("İyimser tasarruf", 1_402_500),
        ("Kötümser tasarruf", 1_083_750),
        ("M-SEN tasarruf", 1_147_500),
    ], 24):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Üretim Bileşenleri"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=28), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=28))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G9")

    g2 = BarChart()
    g2.type = "col"
    g2.title = "KV / Senaryo Tasarruf"
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
        ("Matrah", 10_200_000),
        ("Tasarruf", 1_275_000),
        ("Standart KV", 2_550_000),
    ], 31):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "KPI Matrah / Tasarruf / Standart"
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
        ("Negatif hasılat?",
         '=IF(GIRDI!B9<0,"NEGATİF","TEMİZ")'),
        ("Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(ukv_yilTaban),N(hesapYili)<=N(ukv_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        ("Matrah sıfır mı?",
         '=IF(IFERROR(MOTOR!G23,0)=0,"SIFIR","NORMAL")'),
        ("Vaka kırık mı?",
         '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")'),
        ("Tasarruf eşiği aşıyor mu?",
         '=IF(IFERROR(MOTOR!G26,0)>ukv_kacirilanEsik+N(GIRDI!B9)*ukv_sifirCarpan,"AŞIYOR","NORMAL")'),
        ("Yorum A/B seçili mi?",
         '=IF(OR(GIRDI!B15="A",GIRDI!B15="B"),"TEMİZ","EKSİK")'),
        ("Motor adım sayısı",
         "=COUNTA(MOTOR!B7:B53)"),
    ]
    for i, (ad, form) in enumerate(kontroller_list, 5):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Canlı formül | Değiştirmeyin")

    h(ws, 14, 1, "Toplam giriş alanı", yazi="333333")
    h(ws, 14, 2, "=IFERROR(ukv_girisBeklenen+COUNTA(GIRDI!B6:B15)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Dolu giriş", yazi="333333")
    h(ws, 15, 2, "=COUNTA(GIRDI!B6:B15)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Kalite skoru (modül)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=IFERROR(ROUND(IF(OR(B14=0,B14=\"\"),0,B15/B14*100),0),0)", sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)

    h(ws, 17, 1, "Matrah kontrol", yazi="333333")
    h(ws, 17, 2, "=IFERROR(ukv_matrah,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tasarruf kontrol", yazi="333333")
    h(ws, 18, 2, "=IFERROR(ukv_tasarruf,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Karar kontrol", yazi="333333")
    h(ws, 19, 2, "=IFERROR(ukv_kararMetni,\"-\")", yazi=DIS_REF_YEŞIL)

    genislik(ws, {"A": 40, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek senaryo kütüphanesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:I3")
    h(ws, 4, 1, "Bu sayfa demo değerleri saklar; GIRDI'ye kopyalanabilir.", yazi=GRİ)
    sutunlar = ["Senaryo", "Hasılat", "Maliyet", "GY Payı", "FG Payı", "Diğer",
                "İstisna", "Yıl", "Üretim Kârı"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Üretim Kârı": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(tblOrnek[[#This Row],[Hasılat]]-tblOrnek[[#This Row],[Maliyet]]"
            "-tblOrnek[[#This Row],[GY Payı]]-tblOrnek[[#This Row],[FG Payı]],0))"
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:I1004", sutunlar, formuller)
    ornekler = [
        ("Demo ana", 40_000_000, 28_000_000, 1_500_000, 500_000, 200_000, 0, 2026),
        ("Yüksek marj", 50_000_000, 30_000_000, 1_000_000, 400_000, 0, 0, 2026),
        ("Düşük tasarruf", 12_000_000, 10_500_000, 800_000, 300_000, 0, 0, 2026),
        ("2025 geçiş oranı", 40_000_000, 28_000_000, 1_500_000, 500_000, 200_000, 0, 2025),
        ("İstisna yoğun", 40_000_000, 28_000_000, 1_500_000, 500_000, 200_000, 5_000_000, 2026),
    ]
    for i, row in enumerate(ornekler, 6):
        for k, v in enumerate(row, 1):
            ws.cell(i, k).value = v
            if k in range(2, 8):
                ws.cell(i, k).number_format = TL
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
            elif k == 8:
                ws.cell(i, k).number_format = CATI
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
        ws.cell(i, 1).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
        ws.cell(i, 1).protection = Protection(locked=False)
        for kolon in range(1, 9):
            _kim(ws, f"{get_column_letter(kolon)}{i}", sutunlar[kolon - 1], "Örnek senaryo değeri")

    for col, ad in [("B", "Hasılat"), ("C", "Maliyet"), ("D", "GY Payı"),
                    ("E", "FG Payı"), ("F", "Diğer"), ("G", "İstisna")]:
        lo = "-1000000000000" if col == "F" else "0"
        dogrulama(ws, "decimal", lo, f"{col}6:{col}1004",
                  baslik=ad, mesaj=f"{ad} değerini girin.",
                  hata_baslik="Geçersiz", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000000")

    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6, 7, 8])
    h(ws, 5, 12, "KarDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([10_000_000, 18_600_000, 400_000, 10_000_000, 10_000_000], 6):
        h(ws, i, 12, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Örnek Senaryo Üretim Kârı"
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
        _kim(ws, f"C{i}", "Yorum", "Pay dağıtım yorumu")
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
        ("ukv_tolerans", 0.01, "TL", "Vaka tutarlılık toleransı", "Uygulama notu", "01.01.2025", "11.08.2026"),
        ("ukv_kacirilanEsik", 50_000, "TL", "PANO tasarruf uyarı eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("ukv_senaryoIyi", 1.10, "çarpan", "İyimser matrah çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ukv_senaryoKotu", 0.85, "çarpan", "Kötümser matrah çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ukv_oranArtisPuan", 0.01, "oran", "M-SEN sabit +1 puan", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("ukv_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ukv_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ukv_tornadoHasilat", 1.05, "çarpan", "Tornado hasılat şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ukv_tornadoMaliyet", 1.05, "çarpan", "Tornado maliyet şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ukv_payArtisOran", 0.10, "oran", "Yorum B pay +%10", "Pay dağıtım modeli", "01.01.2025", "11.08.2026"),
        ("ukv_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "11.08.2026"),
        ("ukv_parmakIzi", "UKV-PRO-1.0.0", "metin", "Dosya parmak izi etiketi", "Üretim", "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv", "01.01.2025", "11.08.2026"),
        ("ukv_girisBeklenen", 10, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ukv_riskDusuk", 15, "puan", "UYGULA risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("ukv_riskOrta", 55, "puan", "İNCELE risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("ukv_riskYuksek", 85, "puan", "UYGULAMA risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("ukv_riskEsikOran", 0.5, "oran", "Yedek risk oranı", "İç politika", "01.01.2025", "11.08.2026"),
        ("ukv_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı (hesapYili−taban)", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ukv_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ukv_dayanakCarpan", 10, "adet", "Yedek çarpan", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ukv_bazCarpan", 1, "çarpan", "Baz senaryo çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ukv_sifirCarpan", 0, "çarpan", "Nötr çarpan (sıfır)", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ukv_siraBir", 1, "adet", "Tornado sıra 1", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ukv_siraIki", 2, "adet", "Tornado sıra 2", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ukv_siraUc", 3, "adet", "Tornado sıra 3", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ukv_tahminX", 5, "adet", "FORECAST X noktası", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("vaka_matrah_1", 1_000_000, "TL", "Vaka1/2/3 matrah", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_matrah_2", 200_000, "TL", "Vaka4 matrah", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_tavan", 150_000, "TL", "Yedek vaka tavanı", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_endeks_orani", 1.10, "çarpan", "Yedek endeks oranı", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
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
        "UYGULA — tasarruf uygula eşiğinin üzerinde; indirimli oranı uygulayın.",
        "İNCELE — tasarruf incele bandında; pay ve istisna dayanağını gözden geçirin.",
        "UYGULAMA — tasarruf eşik altında veya matrah yok; öncelik düşük.",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 10, 1, "Belirsizlik beyanları (M05)", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari(["pay_dagitim_yorumu"]), 11):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 13, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2,
      "Genel yönetim ve finansman payları girdideki tutarlarla aynen üretim kârından düşülür. "
      "Matrah = MAX(0, üretim kârı + diğer faaliyet − istisna).",
      kaydir=True, yazi="333333")
    ws.merge_cells("B13:J13")
    h(ws, 14, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 2,
      "Yorum B, payları +%10 sıkılaştırır (daha yüksek maliyet payı → daha düşük matrah). "
      "Dosya her iki yorumun sonucunu MOTOR'da üretir.",
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
    # UYGULAMA önce (UYGULA alt dizgesi); sonra tam UYGULA
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("VERİ YOK",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("UYGULAMA",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['AND(ISNUMBER(SEARCH("UYGULA",B6)),NOT(ISNUMBER(SEARCH("UYGULAMA",B6))))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("İNCELE",B6))'], font=amber, fill=a_fill))
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
        ws.conditional_formatting.add(f"{col}9:{col}14", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(f"{col}9:{col}14", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("F24:F33", FormulaRule(formula=['F24="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F24:F33", FormulaRule(formula=['F24="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F6:F9", CellIsRule(operator="notEqual", formula=["0"], font=amber))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("UYGULAMA",E6))'], font=kirmizi))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['AND(ISNUMBER(SEARCH("UYGULA",E6)),NOT(ISNUMBER(SEARCH("UYGULAMA",E6))))'], font=yesil))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("İNCELE",E6))'], font=amber))
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
    ws.conditional_formatting.add("F7:F53", FormulaRule(formula=['F7="UKV-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F53", FormulaRule(formula=['F7="UKV-005"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "C", "D", "E", "F", "G", "I"):
        ws.conditional_formatting.add(f"{harf}6:{harf}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("I6:I20", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B21", FormulaRule(formula=['ISNUMBER(SEARCH("UYGULAMA",B21))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B21", FormulaRule(formula=['AND(ISNUMBER(SEARCH("UYGULA",B21)),NOT(ISNUMBER(SEARCH("UYGULAMA",B21))))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B21", FormulaRule(formula=['ISNUMBER(SEARCH("İNCELE",B21))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B20", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))


def ad_tanimlari(wb):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "ukv_matrah", "MOTOR!$G$29")
    ad_ekle(wb, "ukv_tasarruf", "MOTOR!$G$28")
    ad_ekle(wb, "ukv_indirimliKv", "MOTOR!$G$30")
    ad_ekle(wb, "ukv_kacirilan", "MOTOR!$G$28")
    ad_ekle(wb, "ukv_kararMetni", "MOTOR!$G$34")
    ad_ekle(wb, "ukv_kararGerekce", "MOTOR!$G$35")
    ad_ekle(wb, "ukv_maddeAtifMetni", "MOTOR!$G$52")
    ad_ekle(wb, "ukv_kanitRaporu", "MOTOR!$G$53")
    ad_ekle(wb, "ukv_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "ukv_parmakIzi", "AYARLAR!$B$19")

    ad_ekle(wb, "ukv_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "ukv_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "ukv_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "ukv_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "ukv_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "ukv_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "ukv_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "ukv_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "ukv_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "ukv_vakaFark", "VAKALAR!$B$13")

    ayar_map = {
        "ukv_tolerans": 8,
        "ukv_kacirilanEsik": 9,
        "ukv_senaryoIyi": 10,
        "ukv_senaryoKotu": 11,
        "ukv_oranArtisPuan": 12,
        "ukv_tahminAlt": 13,
        "ukv_tahminUst": 14,
        "ukv_tornadoHasilat": 15,
        "ukv_tornadoMaliyet": 16,
        "ukv_payArtisOran": 17,
        "ukv_yuzdelikOran": 18,
        "ukv_girisBeklenen": 21,
        "ukv_riskDusuk": 22,
        "ukv_riskOrta": 23,
        "ukv_riskYuksek": 24,
        "ukv_riskEsikOran": 25,
        "ukv_yilTaban": 26,
        "ukv_yuvarMax": 27,
        "ukv_dayanakCarpan": 28,
        "ukv_bazCarpan": 29,
        "ukv_sifirCarpan": 30,
        "ukv_siraBir": 31,
        "ukv_siraIki": 32,
        "ukv_siraUc": 33,
        "ukv_tahminX": 34,
        "vaka_matrah_1": 35,
        "vaka_matrah_2": 36,
        "vaka_tavan": 37,
        "vaka_endeks_orani": 38,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "ukv_modulT1": (11, "C"), "ukv_modulT1Yorum": (11, "D"),
        "ukv_modulT2": (12, "C"), "ukv_modulT2Yorum": (12, "D"),
        "ukv_modulO1": (13, "C"), "ukv_modulO1Yorum": (13, "D"),
        "ukv_modulO6": (16, "C"), "ukv_modulO6Yorum": (16, "D"),
        "ukv_modulO8": (17, "C"), "ukv_modulO8Yorum": (17, "D"),
        "ukv_modulI2": (20, "C"), "ukv_modulI2Yorum": (20, "D"),
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
    varsayilan = os.path.join(kok, "UretimKari125KvOptimizasyonu.xlsx")
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
    cikti = os.path.join(repo, "cikti", "UretimKari125KvOptimizasyonu.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
