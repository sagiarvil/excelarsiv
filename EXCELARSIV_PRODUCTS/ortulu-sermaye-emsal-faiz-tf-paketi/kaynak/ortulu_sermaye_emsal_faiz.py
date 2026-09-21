#!/usr/bin/env python3
"""
Örtülü Sermaye + Emsal Faiz + TF Paketi — üretim betiği (A1 / manda v6).
KVK md.12 örtülü sermaye + emsal faiz farkı + ilişkili işlem TF farkı → KKEG riski.
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

URUN_AD = "Örtülü Sermaye + Emsal Faiz + TF Paketi"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 11)

# Demo 2026: özkaynak 10M; ilişkili borç 40M → oran 4,0 > 3x
# faiz 4M; uygulanan %12 / emsal %10 → faiz farkı 800k
# aşan borç 10M → kabul edilmeyen faiz 1M; TF farkı 500k
# toplam KKEG riski 2,3M → KKEG RİSKİ
DEMO = {
    "sirket_adi": "Örnek Holding A.Ş.",
    "donem": "2026",
    "hesapYili": 2026,
    "ozkaynak": 10_000_000,
    "iliskili_borc": 40_000_000,
    "yillik_faiz_gideri": 4_000_000,
    "uygulanan_faiz_orani": 0.12,
    "emsal_faiz_orani": 0.10,
    "iliskili_islem_tutari": 5_000_000,
    "emsal_bedel": 4_500_000,
    "yorum_modu": "A",
}


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: Örtülü sermaye / emsal faiz / TF KKEG riskini etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Oran, faiz farkı, TF farkı ve karar bozulur | Kaynak: KVK md.12",
    )


def kural_cek(kural_id: str) -> str:
    """Oranları tblKurallar'dan çeker; yıl seçimi ose_yilTaban ile (G07: adında rakam yok)."""
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-ose_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(ose_yuvarMax,IFERROR(N(G12),0)))"


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
      "Özkaynak, ilişkili borç, faiz ve TF girdilerini girin; örtülü sermaye eşiği (3x), "
      "emsal faiz farkı ve TF farkını toplayıp TEMİZ / DİKKAT / KKEG RİSKİ kararını üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Çok yıllı kural tablosu (2025/2026/2027) ile 3x örtülü sermaye eşiği ve emsal faiz",
        "Borç/özkaynak oranı, kabul edilmeyen faiz, emsal faiz farkı ve TF farkının yan yana hesabı",
        "Baz / iyimser / kötümser + kural değişimi (eşik sıkılaştırma) senaryoları",
        "Tebliğ tarzı altın vakalar ve KANIT_RAPORU (A4 imza)",
        "Belirsizlik beyanı: ilişkili kişi kapsam yorumu (A / B)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "SMMM — örtülü sermaye / emsal faiz / TF KKEG riskini ve madde atfını müşteriye vermek",
        "CFO — TEMİZ / DİKKAT / KKEG RİSKİ kararını savunma dosyasına yazmak",
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
        ("Adım 1 — Girdiler", "GIRDI sayfasında yıl, özkaynak, ilişkili borç, faiz ve TF tutarlarını sarı hücrelere girin."),
        ("Adım 2 — Karar", "PANO'da toplam KKEG riskini ve TEMİZ / DİKKAT / KKEG RİSKİ rozetini izleyin."),
        ("Adım 3 — Kanıt", "KANIT_RAPORU'nu yazdırıp imzalayın; VAKALAR tutarlılığını kontrol edin."),
    ]
    for i, (b, m) in enumerate(adimlar, 5):
        h(ws, i, 1, b, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        _kim(ws, f"A{i}", b, "Sırayla izleyin")
    h(ws, 9, 1, "Sık yapılan hatalar", kalin=True, yazi=KRITIK)
    for i, m in enumerate([
        "İlişkili borç toplamına yalnızca faizsiz borçları yazmak",
        "Emsal faiz oranını TCMB/AVG yerine uygulanan oranla aynı bırakmak",
        "Hesap yılını kural tablosuyla uyumsuz seçmek (2025 ≠ 2026 eşiği)",
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
    h(ws, 3, 1, "Şirket ve örtülü sermaye / TF girdileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Eşikler KURALLAR'dan; KKEG riski formülle üretilir.",
      yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Şirket Adı", DEMO["sirket_adi"], None, "sirket_adi",
         "Şirket unvanını yazın."),
        (7, "Dönem", DEMO["donem"], None, "donem",
         "Beyan dönemi veya hesap yılı referansını yazın."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili",
         "2025, 2026 veya 2027 seçin."),
        (9, "Özkaynak [₺]", DEMO["ozkaynak"], TL, "ozkaynak",
         "Dönem sonu özkaynak tutarını TL girin."),
        (10, "İlişkili Borç Toplamı [₺]", DEMO["iliskili_borc"], TL, "iliskili_borc",
         "Ortak/ilişkili kişiden borç toplamını TL girin."),
        (11, "Yıllık Faiz Gideri [₺]", DEMO["yillik_faiz_gideri"], TL, "yillik_faiz_gideri",
         "İlişkili borçlara isabet eden yıllık faiz giderini TL girin."),
        (12, "Uygulanan Faiz Oranı [%]", DEMO["uygulanan_faiz_orani"], YÜZDE, "uygulanan_faiz_orani",
         "Sözleşmede uygulanan faiz oranını yüzde girin."),
        (13, "Emsal Faiz Oranı (TCMB/AVG) [%]", DEMO["emsal_faiz_orani"], YÜZDE, "emsal_faiz_orani",
         "Emsal faiz oranını (TCMB/AVG) yüzde girin."),
        (14, "İlişkili İşlem Tutarı [₺]", DEMO["iliskili_islem_tutari"], TL, "iliskili_islem_tutari",
         "İlişkili kişi işlem tutarını TL girin."),
        (15, "Emsal Bedel [₺]", DEMO["emsal_bedel"], TL, "emsal_bedel",
         "İşlemin emsal bedelini TL girin."),
        (16, "Yorum Modu (A/B)", DEMO["yorum_modu"], None, "yorum_modu",
         "İlişkili kişi yorumu A (standart) veya B (sıkı +%10 aşan borç) seçin."),
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
        else:
            birim = "metin"
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim, yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 18, 1, "Örtülü eşik (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 2, kural_cek("OSE-001"), sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Dikkat eşiği (ayna) [₺]", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2, kural_cek("OSE-003"), sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Toplam KKEG riski (ayna) [₺]", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=IFERROR(ose_toplamKkeg,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 21, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B16)=0,ose_girisBeklenen=0),0,ROUND(COUNTA(B6:B16)/ose_girisBeklenen*100,0)),0)',
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
        ("ozkaynak", DEMO["ozkaynak"], "TL", "Evet", "GIRDI"),
        ("iliskili_borc", DEMO["iliskili_borc"], "TL", "Evet", "GIRDI"),
        ("yillik_faiz_gideri", DEMO["yillik_faiz_gideri"], "TL", "Evet", "GIRDI"),
        ("uygulanan_faiz_orani", DEMO["uygulanan_faiz_orani"], "oran", "Evet", "GIRDI"),
        ("emsal_faiz_orani", DEMO["emsal_faiz_orani"], "oran", "Evet", "GIRDI"),
        ("iliskili_islem_tutari", DEMO["iliskili_islem_tutari"], "TL", "Evet", "GIRDI"),
        ("emsal_bedel", DEMO["emsal_bedel"], "TL", "Evet", "GIRDI"),
        ("yorum_modu", DEMO["yorum_modu"], "A/B", "Evet", "GIRDI"),
    ]
    tl_alanlar = {
        "ozkaynak", "iliskili_borc", "yillik_faiz_gideri",
        "iliskili_islem_tutari", "emsal_bedel",
    }
    oran_alanlar = {"uygulanan_faiz_orani", "emsal_faiz_orani"}
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 25):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in tl_alanlar:
            ws.cell(i, 2).number_format = TL
        elif a in oran_alanlar:
            ws.cell(i, 2).number_format = YÜZDE
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
    dogrulama(ws, "list", "ListeYorumAB", "B16",
              baslik="Yorum", mesaj="Yorum modunda A veya B seçin.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B seçin.", bos=False)
    dogrulama(ws, "list", "ListeYillar", "B27",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeYorumAB", "B35",
              baslik="Yorum", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.")
    for aralik, baslik, mesaj in [
        ("B9", "Özkaynak", "0 ile 1e12 arasında TL."),
        ("B10", "İlişkili Borç", "0 ile 1e12 arasında TL."),
        ("B11", "Yıllık Faiz", "0 ile 1e12 arasında TL."),
        ("B14", "İlişkili İşlem", "0 ile 1e12 arasında TL."),
        ("B15", "Emsal Bedel", "0 ile 1e12 arasında TL."),
        ("B28", "Özkaynak", "Tablo: TL tutar."),
        ("B29", "İlişkili Borç", "Tablo: TL tutar."),
        ("B30", "Yıllık Faiz", "Tablo: TL tutar."),
        ("B33", "İlişkili İşlem", "Tablo: TL tutar."),
        ("B34", "Emsal Bedel", "Tablo: TL tutar."),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj=mesaj,
                  hata_baslik="Geçersiz tutar", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000000")
    for aralik, baslik in [("B12", "Uygulanan Faiz"), ("B13", "Emsal Faiz"),
                           ("B31", "Uygulanan Faiz"), ("B32", "Emsal Faiz")]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj="0 ile 1 arasında oran.",
                  hata_baslik="Geçersiz oran", hata_mesaj="0-1 aralığı.",
                  isaret="between", f2="1")

    giris_hucreleri(ws, 6, 16, [2])
    giris_hucreleri(ws, 25, 35, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 40, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1025, "Sarı alanlar giriş; eşik ve KKEG riski formülleri kilitlidir.")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Mevzuat kural tablosu (yıl yan yana)", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Eşikler MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    # OSE-001: 3x örtülü sermaye eşiği; OSE-002: referans emsal faiz (bilgi)
    # OSE-003: dikkat eşiği; OSE-004: KKEG risk eşiği; OSE-005: yuvarlama
    satirlar = [
        ["OSE-001", "KVK md.12", "ortulu_oran_esigi", "ilişkili borç / özkaynak", "ortulu_oran_esigi",
         3.0, 3.0, 3.0, "01.01.2025", "Örtülü sermaye borç/özkaynak eşiği (kat)"],
        ["OSE-002", "KVK md.12 / TCMB", "emsal_faiz_referans", "faiz karşılaştırması", "emsal_faiz_referans",
         0.45, 0.40, 0.35, "01.01.2025", "Referans emsal faiz (bilgi; GIRDI öncelikli)"],
        ["OSE-003", "İç politika", "dikkat_esik", "KKEG risk ≥ eşik", "dikkat_esik_tl",
         100000, 100000, 100000, "01.01.2025", "DİKKAT karar eşiği (TL)"],
        ["OSE-004", "İç politika", "kkeg_esik", "KKEG risk ≥ eşik", "kkeg_esik_tl",
         500000, 500000, 500000, "01.01.2025", "KKEG RİSKİ karar eşiği (TL)"],
        ["OSE-005", "Genel", "yuvarlama", "her hesap", "yuvarlama_ondalik",
         2, 2, 2, "01.01.2025", "Uygulama notu"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = None
            if k in (6, 7, 8):
                sayi = YÜZDE if s[0] == "OSE-002" else CATI
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-ose_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Örtülü eşik (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("OSE-001"), sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "KKEG eşiği (aktif yıl)", yazi="333333")
    h(ws, 16, 2, kural_cek("OSE-004"), sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"OSE-001=",TEXT(B15,"0.0"),"OSE-004=",TEXT(B16,"₺ #,##0"),'
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

    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif mevzuat yılı", "OSE-005")  # G7
    ekle("Örtülü oran eşiği", km("OSE-001"), "kat", "Borç/özkaynak eşiği", "OSE-001")  # G8
    ekle("Emsal faiz referans", km("OSE-002"), "oran", "Referans emsal faiz", "OSE-002")  # G9
    ekle("Dikkat eşiği", km("OSE-003"), "TL", "DİKKAT KKEG eşiği", "OSE-003")  # G10
    ekle("KKEG risk eşiği", km("OSE-004"), "TL", "KKEG RİSKİ eşiği", "OSE-004")  # G11
    ekle("Yuvarlama ondalık", km("OSE-005"), "adet", "Yuvarlama basamağı", "OSE-005")  # G12
    ekle("Özkaynak", "=GIRDI!B9", "TL", "Özkaynak", "OSE-001")  # G13
    ekle("İlişkili borç", "=GIRDI!B10", "TL", "İlişkili borç toplamı", "OSE-001")  # G14
    ekle("Yıllık faiz gideri", "=GIRDI!B11", "TL", "Faiz gideri", "OSE-001")  # G15
    ekle("Uygulanan faiz oranı", "=GIRDI!B12", "oran", "Uygulanan faiz", "OSE-002")  # G16
    ekle("Emsal faiz oranı (giriş)", "=GIRDI!B13", "oran", "Emsal faiz GIRDI", "OSE-002")  # G17
    ekle("İlişkili işlem tutarı", "=GIRDI!B14", "TL", "TF işlem tutarı", "OSE-005")  # G18
    ekle("Emsal bedel", "=GIRDI!B15", "TL", "TF emsal bedel", "OSE-005")  # G19
    ekle("Yorum çarpanı",
         '=IF(GIRDI!B16="B",1+ose_payArtisOran,1)',
         "çarpan", "Yorum A/B çarpanı", "OSE-005")  # G20
    ekle("Emsal faiz efektif",
         "=IF(G17>0,G17,G9)",
         "oran", "GIRDI yoksa referans", "OSE-002")  # G21
    ekle("Borç / özkaynak oranı",
         "=IF(G13=0,0,G14/G13)",
         "oran", "Örtülü sermaye oranı", "OSE-001")  # G22
    ekle("Aşan borç",
         f"=ROUND(MAX(0,G14-G13*G8)*G20,{_YV})",
         "TL", "Eşik üstü borç (yorumlu)", "OSE-001")  # G23
    ekle("Kabul edilmeyen faiz",
         f"=ROUND(IF(G14=0,0,G15*(G23/G14)),{_YV})",
         "TL", "Örtülü sermaye faiz KKEG", "OSE-001")  # G24
    ekle("Emsal faiz farkı",
         f"=ROUND(MAX(0,G16-G21)*G14,{_YV})",
         "TL", "(Uygulanan−emsal)×borç", "OSE-002")  # G25
    ekle("TF farkı",
         f"=ROUND(ABS(G18-G19),{_YV})",
         "TL", "|İşlem−emsal bedel|", "OSE-005")  # G26
    ekle("Toplam KKEG riski",
         "=G24+G25+G26",
         "TL", "Üç kalem toplamı", "OSE-004")  # G27
    ekle("Toplam KKEG (pano)", "=G27", "TL", "KKEG ayna", "OSE-004")  # G28
    ekle("Borç oranı (pano)", "=G22", "oran", "Oran ayna", "OSE-001")  # G29
    ekle("Kabul edilmeyen (pano)", "=G24", "TL", "Faiz KKEG ayna", "OSE-001")  # G30
    ekle("Kabul edilmeyen / faiz",
         "=IF(G15=0,0,G24/G15)",
         "oran", "Faiz kısıt oranı", "OSE-001")  # G31
    ekle("TF / işlem yoğunluğu",
         "=IF(G18=0,0,G26/G18)",
         "oran", "TF fark yoğunluğu", "OSE-005")  # G32
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(OR(G27>=G11,AND(G22>G8,G24>0,G27>=G10)),"KKEG_RISKI",'
         'IF(OR(G27>=G10,G22>G8),"DIKKAT","TEMIZ")))',
         "kod", "Karar kodu", "OSE-005")  # G33
    ekle("Karar metni",
         '=IF(G33="VERI_YOK","VERİ YOK",'
         'IF(G33="KKEG_RISKI","KKEG RİSKİ",'
         'IF(G33="DIKKAT","DİKKAT","TEMİZ")))',
         "metin", "Kullanıcıya karar", "OSE-005")  # G34
    ekle("Gerekçe",
         '=IF(G33="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G33="KKEG_RISKI","Toplam KKEG riski eşiğin üzerinde veya örtülü sermaye + faiz kısıtı oluştu; savunma dosyası zorunlu.",'
         'IF(G33="DIKKAT","Oran veya KKEG riski dikkat bandında; ilişkili borç ve emsal faiz dayanağını gözden geçirin.",'
         '"Borç/özkaynak eşiğin altında ve KKEG riski düşük; dosya TEMİZ görünüyor.")))',
         "metin", "Gerekçe cümlesi", "OSE-005")  # G35
    ekle("Emsal faiz fark / toplam",
         "=IF(G27=0,0,G25/G27)",
         "oran", "Faiz fark payı", "OSE-002")  # G36
    ekle("Örtülü bayrak",
         '=IF(G22>G8,1,0)',
         "bayrak", "Eşik aşıldı mı", "OSE-001")  # G37
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B16)=0,ose_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B16)/ose_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "OSE-005")  # G38
    ekle("Risk skoru",
         '=IF(G33="VERI_YOK",0,IF(G33="TEMIZ",ose_riskDusuk,'
         'IF(G33="DIKKAT",ose_riskOrta,ose_riskYuksek)))',
         "puan", "KKEG risk skoru", "OSE-005")  # G39
    ekle("Senaryo iyimser KKEG", "=G27*ose_senaryoIyi", "TL", "İyimser (düşük) risk", "OSE-004")  # G40
    # iyimser = daha düşük risk: senaryoIyi aslında 0.85 for risk? Keep UKV pattern: iyi=1.10 higher
    # For risk product: iyimser = lower KKEG. Use senaryoKotu for optimistic low risk.
    ekle("Senaryo kötümser KKEG", "=G27*ose_senaryoKotu", "TL", "Kötümser (yüksek) risk", "OSE-004")  # G41
    ekle("İyimser KKEG (düşük)", f"=ROUND(G27*ose_senaryoKotu,{_YV})", "TL", "İyimser düşük risk", "OSE-004")  # G42
    ekle("Kötümser KKEG (yüksek)", f"=ROUND(G27*ose_senaryoIyi,{_YV})", "TL", "Kötümser yüksek risk", "OSE-004")  # G43
    ekle("M-SEN eşik sıkı", "=MAX(1,G8-ose_esikSikilastirma)", "kat", "Eşik sıkılaştırma", "OSE-001")  # G44
    # oranArtisPuan=0.01 → *100 = 1; eşik 3→2. Better: ose_esikSikilastirma named as 0.5
    ekle("M-SEN aşan borç",
         f"=ROUND(MAX(0,G14-G13*G44)*G20,{_YV})",
         "TL", "Sıkı eşikte aşan borç", "OSE-001")  # G45
    ekle("M-SEN KKEG",
         f"=ROUND(IF(G14=0,0,G15*(G45/G14))+G25+G26,{_YV})",
         "TL", "M-SEN toplam KKEG", "OSE-001")  # G46
    ekle("Tahmin üst", "=G27*ose_tahminUst", "TL", "Tahmin üst bant", "OSE-005")  # G47
    ekle("Tahmin alt", "=G27*ose_tahminAlt", "TL", "Tahmin alt bant", "OSE-005")  # G48
    ekle("Tornado: borç etkisi",
         f"=ABS(ROUND(IF(G14*ose_tornadoHasilat=0,0,G15*(MAX(0,G14*ose_tornadoHasilat-G13*G8)*G20)"
         f"/(G14*ose_tornadoHasilat))+MAX(0,G16-G21)*G14*ose_tornadoHasilat+G26,{_YV})-G27)",
         "TL", "Borç ± etki", "OSE-001")  # G49
    ekle("Tornado: faiz oranı etkisi",
         f"=ABS(ROUND(G24+MAX(0,G16*ose_tornadoMaliyet-G21)*G14+G26,{_YV})-G27)",
         "TL", "Uygulanan faiz ± etki", "OSE-002")  # G50
    ekle("Tornado: TF etkisi",
         f"=ABS(ROUND(G24+G25+ABS(G18*ose_tornadoHasilat-G19),{_YV})-G27)",
         "TL", "TF ± etki", "OSE-005")  # G51
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"OSE-001→örtülü eşik","OSE-002→emsal faiz",'
         '"OSE-003→dikkat eşiği","OSE-004→KKEG eşiği")',
         "metin", "Kanıt atıfları", "OSE-001")  # G52
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Oran=",TEXT(G22,"0.00"),"KabulEdilmeyen=",TEXT(G24,"₺ #,##0"),'
         '"EmsalFark=",TEXT(G25,"₺ #,##0"),"TF=",TEXT(G26,"₺ #,##0"),"Toplam=",TEXT(G27,"₺ #,##0"))',
         "metin", "Rapor özeti", "OSE-005")  # G53

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
        elif birim in ("oran", "çarpan", "kat"):
            ws.cell(r, 7).number_format = YÜZDE if birim != "kat" else CATI
            ws.cell(r, 3).number_format = YÜZDE if birim != "kat" else CATI
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
    h(ws, 5, 2, "Risk Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "KKEG Riski", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "KKEG (karar)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    # İyimser = düşük risk (ose_senaryoKotu=0.85); Kötümser = yüksek (ose_senaryoIyi=1.10)
    senaryolar = [
        ("İyimser", "=IFERROR(ose_senaryoKotu+N(GIRDI!B9)*ose_sifirCarpan,0)",
         "=IFERROR(MOTOR!G42,0)",
         "=IFERROR(MOTOR!G42,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D6>=MOTOR!G11,"KKEG RİSKİ",IF(D6>=MOTOR!G10,"DİKKAT","TEMİZ")))'),
        ("Baz", "=IFERROR(ose_bazCarpan+N(GIRDI!B9)*ose_sifirCarpan,ose_bazCarpan)",
         "=IFERROR(ose_toplamKkeg,0)",
         "=IFERROR(ose_toplamKkeg,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(ose_kararMetni,"-"))'),
        ("Kötümser", "=IFERROR(ose_senaryoIyi+N(GIRDI!B9)*ose_sifirCarpan,0)",
         "=IFERROR(MOTOR!G43,0)",
         "=IFERROR(MOTOR!G43,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D8>=MOTOR!G11,"KKEG RİSKİ",IF(D8>=MOTOR!G10,"DİKKAT","TEMİZ")))'),
        ("M-SEN eşik sıkı", "=IFERROR(ose_bazCarpan+N(GIRDI!B9)*ose_sifirCarpan,ose_bazCarpan)",
         "=IFERROR(ose_toplamKkeg,0)",
         "=IFERROR(MOTOR!G46,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D9>=MOTOR!G11,"KKEG RİSKİ",IF(D9>=MOTOR!G10,"DİKKAT","TEMİZ")))'),
    ]
    for i, (ad, carp, mat, tas, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, mat, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, tas, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant genişliği (kötümser−iyimser KKEG)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(D8-D6,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser KKEG ",TEXT(D6,"₺ #,##0")," · Baz ",TEXT(D7,"₺ #,##0"),'
      '" · Kötümser ",TEXT(D8,"₺ #,##0")," · Bant ",TEXT(B11,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:H12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Tornado borç etkisi", yazi="333333")
    h(ws, 16, 2, "=IFERROR(MOTOR!G49,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado faiz oranı etkisi", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G50,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado TF etkisi", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G51,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 3,
      '=IFERROR(IF(B16>=MAX(B16:B18),ose_siraBir,IF(B16>=MIN(B16:B18)+ABS(B16-B17),ose_siraIki,ose_siraUc)),ose_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 3,
      '=IFERROR(IF(B17>=MAX(B16:B18),ose_siraBir,IF(B17=MEDIAN(B16:B18),ose_siraIki,ose_siraUc)),ose_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 3,
      '=IFERROR(IF(B18>=MAX(B16:B18),ose_siraBir,IF(B18=MEDIAN(B16:B18),ose_siraIki,ose_siraUc)),ose_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(CONCATENATE("1)",INDEX(A16:A18,MATCH(ose_siraBir,C16:C18,0)),'
      '" 2)",INDEX(A16:A18,MATCH(ose_siraIki,C16:C18,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B16:B18),"₺ #,##0")," TL — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN kural değişimi KKEG", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=D9", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Eşik sıkılaştırma senaryosunda KKEG ",TEXT(D9,"₺ #,##0")," TL")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı (KKEG)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=MOTOR!G48", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 3, "=IFERROR(MOTOR!G27,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
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

    h(ws, 5, 7, "KkegDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([1_955_000, 2_300_000, 2_530_000, 2_800_000], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([400000, 320000, 250000], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)

    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([1_955_000, 2_300_000, 2_530_000, 2_800_000], 28):
        h(ws, i, 4, v, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo KKEG Karşılaştırması"
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
    g3.title = "Senaryo KKEG Trendi"
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

    vakalar_data = [
        ("V-001", "Aşan borç × faiz / borç = kabul edilmeyen",
         "borc=vaka_borc; ozk=vaka_ozkaynak; faiz=vaka_faiz; esik=OSE-001",
         1000000,
         "=IFERROR(ROUND(IF(vaka_borc=0,0,vaka_faiz*(MAX(0,vaka_borc-vaka_ozkaynak*"
         "INDEX(tblKurallar[deger_2026],MATCH(\"OSE-001\",tblKurallar[kural_id],0)))/vaka_borc)),2),0)",
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=ose_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "KVK md.12 örtülü faiz"),
        ("V-002", "Emsal faiz farkı = (uygulanan−emsal)×borç",
         "uygulanan=vaka_uygulanan; emsal=vaka_emsal; borc=vaka_borc",
         800000,
         "=IFERROR(ROUND(MAX(0,vaka_uygulanan-vaka_emsal)*vaka_borc*"
         "IF(INDEX(tblKurallar[deger_2026],MATCH(\"OSE-001\",tblKurallar[kural_id],0))>0,1,0),2),0)",
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=ose_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Emsal faiz farkı"),
        ("V-003", "TF farkı = |işlem−emsal bedel|",
         "islem=vaka_islem; bedel=vaka_bedel",
         500000,
         "=IFERROR(ROUND(ABS(vaka_islem-vaka_bedel)*"
         "IF(INDEX(tblKurallar[deger_2026],MATCH(\"OSE-001\",tblKurallar[kural_id],0))>0,1,0),2),0)",
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=ose_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "TF farkı"),
        ("V-004", "Toplam KKEG = V-001+V-002+V-003",
         "üç kalem toplamı",
         2300000,
         "=IFERROR(ROUND("
         "IF(vaka_borc=0,0,vaka_faiz*(MAX(0,vaka_borc-vaka_ozkaynak*"
         "INDEX(tblKurallar[deger_2026],MATCH(\"OSE-001\",tblKurallar[kural_id],0)))/vaka_borc))"
         "+MAX(0,vaka_uygulanan-vaka_emsal)*vaka_borc+ABS(vaka_islem-vaka_bedel),2),0)",
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=ose_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Toplam KKEG riski"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v,
              yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = TL

    formuller = {
        "fark": "=IF(tblVakalar[[#This Row],[vaka_id]]=\"\",\"\",IFERROR(tblVakalar[[#This Row],[beklenen_sonuc]]-tblVakalar[[#This Row],[hesaplanan]],0))",
        "durum": '=IF(tblVakalar[[#This Row],[vaka_id]]="","",IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=ose_tolerans,"TUTARLI","KIRIK"),"KIRIK"))',
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
    for i, (b, he) in enumerate([(1000000, 1000000), (800000, 800000), (500000, 500000), (2300000, 2300000)], 6):
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
    h(ws, 3, 1, "KANIT RAPORU — Örtülü Sermaye + Emsal Faiz + TF", kalin=True, boyut=14, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:F3")
    h(ws, 4, 1, "Rapor tarihi", yazi="333333")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 5, 1, "Şirket", yazi="333333")
    h(ws, 5, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Dönem / yıl", yazi="333333")
    h(ws, 6, 2, '=CONCATENATE(GIRDI!B7," / ",hesapYili)', yazi=DIS_REF_YEŞIL)

    h(ws, 8, 1, "Girdi özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, (et, form, sayi) in enumerate([
        ("Özkaynak", "=GIRDI!B9", TL),
        ("İlişkili borç", "=GIRDI!B10", TL),
        ("Yıllık faiz gideri", "=GIRDI!B11", TL),
        ("Uygulanan faiz oranı", "=GIRDI!B12", YÜZDE),
        ("Emsal faiz oranı", "=GIRDI!B13", YÜZDE),
        ("İlişkili işlem tutarı", "=GIRDI!B14", TL),
        ("Emsal bedel", "=GIRDI!B15", TL),
    ], 9):
        h(ws, i, 1, et, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)

    h(ws, 17, 1, "Hesap zinciri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 18, 1, "Borç / özkaynak oranı", yazi="333333")
    h(ws, 18, 2, "=ose_borcOrani", sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Kabul edilmeyen faiz", yazi="333333")
    h(ws, 19, 2, "=MOTOR!G24", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Emsal faiz farkı", yazi="333333")
    h(ws, 20, 2, "=MOTOR!G25", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 21, 1, "TF farkı", yazi="333333")
    h(ws, 21, 2, "=MOTOR!G26", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 22, 1, "Toplam KKEG riski", yazi="333333")
    h(ws, 22, 2, "=ose_toplamKkeg", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 23, 1, "KARAR", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "=ose_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 24, 1, "Gerekçe", yazi="333333")
    h(ws, 24, 2, "=ose_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B24:F24")

    h(ws, 26, 1, "Madde atıfları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 2, "=ose_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B26:F26")

    h(ws, 28, 1, "Kanıt gövdesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 28, 2, "=ose_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B28:F28")

    h(ws, 30, 1, "Parmak izi (SHA-256 kısaltma)", yazi="333333")
    h(ws, 30, 2, "=ose_parmakIzi", yazi=GRİ)

    h(ws, 32, 1, "Hazırlayan (imza)", yazi="333333")
    h(ws, 32, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B32", "İmza", "Raporu hazırlayan kişinin adını yazın.")
    h(ws, 33, 1, "Onaylayan (imza)", yazi="333333")
    h(ws, 33, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B33", "Onay", "Onaylayan kişinin adını yazın.")

    h(ws, 35, 1,
      "Uyarı: Bu çıktı karar destektir; kesin vergi sonucu veya hukuki görüş değildir. "
      f"Sürüm {SURUM}.",
      yazi=GRİ, boyut=9, kaydir=True)
    ws.merge_cells("A35:F35")

    baski_hazirla(ws, "A1:F36", f"{URUN_AD} · {SURUM}")
    ws.page_setup.orientation = "portrait"
    genislik(ws, {"A": 28, "B": 22, "C": 14, "D": 14, "E": 14, "F": 14})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim paneli — Örtülü Sermaye / Emsal Faiz / TF", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpiler = [
        (2, "Oran", "=IFERROR(ose_borcOrani,0)", YÜZDE),
        (3, "KabulEdilmeyen", "=IFERROR(MOTOR!G24,0)", TL),
        (4, "EmsalFark", "=IFERROR(MOTOR!G25,0)", TL),
        (5, "TFFark", "=IFERROR(MOTOR!G26,0)", TL),
        (6, "ToplamKKEG", "=IFERROR(ose_toplamKkeg,0)", TL),
        (7, "Güven", "=IFERROR(MOTOR!G38,0)", CATI),
        (8, "Risk", "=IFERROR(MOTOR!G39,0)", CATI),
        (9, "Eşik", "=IFERROR(MOTOR!G8,0)", CATI),
        (10, "DikkatEsik", "=IFERROR(MOTOR!G10,0)", TL),
        (11, "M-SEN", "=IFERROR(ose_kuralDegisimSenaryo,0)", TL),
        (12, "Tahmin", "=IFERROR(SENARYO!B32,0)", TL),
        (13, "Vaka", "=IFERROR(ose_vakaDurum,\"-\")", None),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=9, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=ose_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=ose_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 10, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Borç / özkaynak oranı",
         "=IFERROR(MOTOR!G22,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Oran ",TEXT(C11,"0.00")," — ilişkili borç / özkaynak")'),
        ("T2", "Kabul edilmeyen / yıllık faiz",
         "=IFERROR(MOTOR!G31,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Faiz kısıt oranı ",TEXT(C12,"0.0%"))'),
        ("O1", "Senaryo KKEG sapması",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,STDEV.P(SENARYO!D6:D9)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo sapması ",TEXT(C13,"₺ #,##0")," TL")'),
        ("O2", "Tornado max etki",
         "=IFERROR(ose_duyarlilikSira,\"-\")",
         "=IFERROR(ose_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo bant",
         "=IFERROR(ose_senaryoKarsilastirma,0)",
         "=IFERROR(ose_yorumSenaryo,\"-\")"),
        ("O6", "TF / işlem yoğunluğu",
         "=IFERROR(MOTOR!G32,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"TF yoğunluğu ",TEXT(C16,"0.0%"))'),
        ("O8", "Kalite skoru",
         "=IFERROR(KONTROLLER!B12,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalite ",TEXT(C17,"0")," puan")'),
        ("M", "M-SEN KKEG",
         "=IFERROR(ose_kuralDegisimSenaryo,0)",
         "=IFERROR(ose_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin aralık",
         "=IFERROR(ose_tahminAralik,0)",
         "=IFERROR(ose_yorumTahmin,\"-\")"),
        ("I2", "Senaryo KKEG P90",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!D6:D9,ose_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"P90 KKEG ",TEXT(C20,"₺ #,##0")," TL")'),
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
        ("Özkaynak", 10_000_000),
        ("İlişkili borç", 40_000_000),
        ("Yıllık faiz", 4_000_000),
        ("Kabul edilmeyen", 1_000_000),
        ("Toplam KKEG", 2_300_000),
    ], 24):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)

    h(ws, 23, 4, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 5, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Emsal faiz farkı", 800_000),
        ("TF farkı", 500_000),
        ("İyimser KKEG", 1_955_000),
        ("Kötümser KKEG", 2_530_000),
        ("M-SEN KKEG", 2_800_000),
    ], 24):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Örtülü Sermaye Bileşenleri"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=28), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=28))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G9")

    g2 = BarChart()
    g2.type = "col"
    g2.title = "KKEG / Senaryo"
    g2.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "G23")

    g3 = LineChart()
    g3.title = "KKEG Çizgisi"
    g3.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g3.height, g3.width = 7, 12
    ws.add_chart(g3, "P9")

    h(ws, 30, 1, "Metrik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 2, "Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Toplam KKEG", 2_300_000),
        ("Kabul edilmeyen", 1_000_000),
        ("Emsal fark", 800_000),
    ], 31):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "KPI KKEG / Faiz / Emsal"
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
        ("Negatif özkaynak?",
         '=IF(GIRDI!B9<0,"NEGATİF","TEMİZ")'),
        ("Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(ose_yilTaban),N(hesapYili)<=N(ose_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        ("Toplam KKEG sıfır mı?",
         '=IF(IFERROR(MOTOR!G27,0)=0,"SIFIR","NORMAL")'),
        ("Vaka kırık mı?",
         '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")'),
        ("KKEG eşiği aşıyor mu?",
         '=IF(IFERROR(MOTOR!G27,0)>ose_kacirilanEsik+N(GIRDI!B9)*ose_sifirCarpan,"AŞIYOR","NORMAL")'),
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
    h(ws, 14, 2, "=IFERROR(ose_girisBeklenen+COUNTA(GIRDI!B6:B16)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Dolu giriş", yazi="333333")
    h(ws, 15, 2, "=COUNTA(GIRDI!B6:B16)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Kalite skoru (modül)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=IFERROR(ROUND(IF(OR(B14=0,B14=\"\"),0,B15/B14*100),0),0)", sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)

    h(ws, 17, 1, "Toplam KKEG kontrol", yazi="333333")
    h(ws, 17, 2, "=IFERROR(ose_toplamKkeg,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Borç oranı kontrol", yazi="333333")
    h(ws, 18, 2, "=IFERROR(ose_borcOrani,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Karar kontrol", yazi="333333")
    h(ws, 19, 2, "=IFERROR(ose_kararMetni,\"-\")", yazi=DIS_REF_YEŞIL)

    genislik(ws, {"A": 40, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek senaryo kütüphanesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:I3")
    h(ws, 4, 1, "Bu sayfa demo değerleri saklar; GIRDI'ye kopyalanabilir.", yazi=GRİ)
    sutunlar = ["Senaryo", "Özkaynak", "İlişkili Borç", "Faiz Gideri", "Uygulanan", "Emsal",
                "İşlem", "Emsal Bedel", "Yıl", "Oran"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Oran": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(IF(tblOrnek[[#This Row],[Özkaynak]]=0,0,"
            "tblOrnek[[#This Row],[İlişkili Borç]]/tblOrnek[[#This Row],[Özkaynak]]),0))"
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:J1004", sutunlar, formuller)
    ornekler = [
        ("Demo ana", 10_000_000, 40_000_000, 4_000_000, 0.12, 0.10, 5_000_000, 4_500_000, 2026),
        ("Temiz oran", 20_000_000, 40_000_000, 2_000_000, 0.10, 0.10, 1_000_000, 1_000_000, 2026),
        ("Yüksek TF", 10_000_000, 25_000_000, 2_000_000, 0.11, 0.10, 8_000_000, 5_000_000, 2026),
        ("2025 referans", 10_000_000, 40_000_000, 4_000_000, 0.12, 0.10, 5_000_000, 4_500_000, 2025),
        ("Dikkat bandı", 15_000_000, 50_000_000, 1_500_000, 0.105, 0.10, 2_000_000, 1_900_000, 2026),
    ]
    for i, row in enumerate(ornekler, 6):
        for k, v in enumerate(row, 1):
            ws.cell(i, k).value = v
            if k in (2, 3, 4, 7, 8):
                ws.cell(i, k).number_format = TL
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
            elif k in (5, 6):
                ws.cell(i, k).number_format = YÜZDE
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
            elif k == 9:
                ws.cell(i, k).number_format = CATI
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
        ws.cell(i, 1).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
        ws.cell(i, 1).protection = Protection(locked=False)
        for kolon in range(1, 10):
            _kim(ws, f"{get_column_letter(kolon)}{i}", sutunlar[kolon - 1], "Örnek senaryo değeri")

    for col, ad in [("B", "Özkaynak"), ("C", "İlişkili Borç"), ("D", "Faiz"),
                    ("G", "İşlem"), ("H", "Emsal Bedel")]:
        dogrulama(ws, "decimal", "0", f"{col}6:{col}1004",
                  baslik=ad, mesaj=f"{ad} değerini girin.",
                  hata_baslik="Geçersiz", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000000")
    for col, ad in [("E", "Uygulanan"), ("F", "Emsal")]:
        dogrulama(ws, "decimal", "0", f"{col}6:{col}1004",
                  baslik=ad, mesaj="0-1 oran.",
                  hata_baslik="Geçersiz", hata_mesaj="0-1.",
                  isaret="between", f2="1")

    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6, 7, 8, 9])
    h(ws, 5, 12, "OranDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([4.0, 2.0, 2.5, 4.0, 3.333], 6):
        h(ws, i, 12, v, sayi=CATI, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Örnek Senaryo Borç/Özkaynak"
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
        _kim(ws, f"C{i}", "Yorum", "İlişkili kişi yorumu")
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
        ("ose_tolerans", 0.01, "TL", "Vaka tutarlılık toleransı", "Uygulama notu", "01.01.2025", "11.08.2026"),
        ("ose_kacirilanEsik", 50_000, "TL", "PANO KKEG uyarı eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("ose_senaryoIyi", 1.10, "çarpan", "Kötümser KKEG çarpanı (yüksek)", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ose_senaryoKotu", 0.85, "çarpan", "İyimser KKEG çarpanı (düşük)", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ose_oranArtisPuan", 0.01, "oran", "Yedek oran artış puanı", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("ose_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ose_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ose_tornadoHasilat", 1.05, "çarpan", "Tornado borç şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ose_tornadoMaliyet", 1.05, "çarpan", "Tornado faiz oranı şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ose_payArtisOran", 0.10, "oran", "Yorum B aşan borç +%10", "İlişkili kişi modeli", "01.01.2025", "11.08.2026"),
        ("ose_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "11.08.2026"),
        ("ose_parmakIzi", "OSE-PRO-1.0.0", "metin", "Dosya parmak izi etiketi", "Üretim", "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv", "01.01.2025", "11.08.2026"),
        ("ose_girisBeklenen", 11, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ose_riskDusuk", 15, "puan", "TEMİZ risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("ose_riskOrta", 55, "puan", "DİKKAT risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("ose_riskYuksek", 85, "puan", "KKEG RİSKİ skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("ose_riskEsikOran", 0.5, "oran", "Yedek risk oranı", "İç politika", "01.01.2025", "11.08.2026"),
        ("ose_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı (hesapYili−taban)", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ose_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ose_dayanakCarpan", 10, "adet", "Yedek çarpan", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ose_bazCarpan", 1, "çarpan", "Baz senaryo çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ose_sifirCarpan", 0, "çarpan", "Nötr çarpan (sıfır)", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ose_siraBir", 1, "adet", "Tornado sıra 1", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ose_siraIki", 2, "adet", "Tornado sıra 2", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ose_siraUc", 3, "adet", "Tornado sıra 3", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ose_tahminX", 5, "adet", "FORECAST X noktası", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ose_esikSikilastirma", 0.5, "kat", "M-SEN eşik azaltma (kat)", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("vaka_ozkaynak", 10_000_000, "TL", "Vaka özkaynak", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_borc", 40_000_000, "TL", "Vaka ilişkili borç", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_faiz", 4_000_000, "TL", "Vaka yıllık faiz", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_uygulanan", 0.12, "oran", "Vaka uygulanan faiz", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_emsal", 0.10, "oran", "Vaka emsal faiz", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_islem", 5_000_000, "TL", "Vaka ilişkili işlem", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
        ("vaka_bedel", 4_500_000, "TL", "Vaka emsal bedel", "Tebliğ örneği", "01.01.2025", "11.08.2026"),
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
                elif row[2] in ("yıl", "adet", "puan", "kat"):
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
        "KKEG RİSKİ — toplam KKEG eşiğin üzerinde veya örtülü sermaye + faiz kısıtı oluştu.",
        "DİKKAT — oran veya KKEG riski dikkat bandında; dayanak gözden geçirilmeli.",
        "TEMİZ — borç/özkaynak eşiğin altında ve KKEG riski düşük.",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 10, 1, "Belirsizlik beyanları (M05)", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari(["iliskili_kisi_yorumu"]), 11):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 13, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2,
      "Aşan borç = MAX(0, ilişkili borç − özkaynak × 3x eşik). "
      "Kabul edilmeyen faiz = yıllık faiz × (aşan borç / ilişkili borç).",
      kaydir=True, yazi="333333")
    ws.merge_cells("B13:J13")
    h(ws, 14, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 2,
      "Yorum B, aşan borcu +%10 sıkılaştırır (daha yüksek örtülü faiz KKEG). "
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
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("VERİ YOK",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("KKEG RİSKİ",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("TEMİZ",B6))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B6))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("F4", CellIsRule(operator="greaterThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("F4", CellIsRule(operator="equal", formula=["0"], font=yesil))
    ws.conditional_formatting.add("G4", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("G4", CellIsRule(operator="between", formula=["50", "79"], font=amber))
    ws.conditional_formatting.add("G4", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))
    ws.conditional_formatting.add("H4", CellIsRule(operator="greaterThanOrEqual", formula=["70"], font=kirmizi))
    ws.conditional_formatting.add("H4", CellIsRule(operator="between", formula=["40", "69"], font=amber))
    ws.conditional_formatting.add("H4", CellIsRule(operator="lessThan", formula=["40"], font=yesil))
    for col in range(2, 14):
        harf = get_column_letter(col)
        ws.conditional_formatting.add(f"{harf}4", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))

    ws = wb["GIRDI"]
    for col in ("B",):
        ws.conditional_formatting.add(f"{col}9:{col}15", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(f"{col}9:{col}11", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("F25:F35", FormulaRule(formula=['F25="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F25:F35", FormulaRule(formula=['F25="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F6:F9", CellIsRule(operator="notEqual", formula=["0"], font=amber))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="greaterThan", formula=["0"], font=amber))
    ws.conditional_formatting.add("D6:D9", CellIsRule(operator="equal", formula=["0"], font=yesil))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("KKEG RİSKİ",E6))'], font=kirmizi))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("TEMİZ",E6))'], font=yesil))
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
    ws.conditional_formatting.add("F7:F53", FormulaRule(formula=['F7="OSE-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F53", FormulaRule(formula=['F7="OSE-005"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "C", "D", "G", "H", "J"):
        ws.conditional_formatting.add(f"{harf}6:{harf}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("J6:J20", CellIsRule(operator="greaterThan", formula=["3"], font=kirmizi))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B23", FormulaRule(formula=['ISNUMBER(SEARCH("KKEG RİSKİ",B23))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B23", FormulaRule(formula=['ISNUMBER(SEARCH("TEMİZ",B23))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B23", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B23))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B22", CellIsRule(operator="greaterThan", formula=["0"], font=kirmizi))


def ad_tanimlari(wb):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "ose_toplamKkeg", "MOTOR!$G$28")
    ad_ekle(wb, "ose_borcOrani", "MOTOR!$G$29")
    ad_ekle(wb, "ose_kabulEdilmeyen", "MOTOR!$G$30")
    ad_ekle(wb, "ose_kacirilan", "MOTOR!$G$28")
    ad_ekle(wb, "ose_kararMetni", "MOTOR!$G$34")
    ad_ekle(wb, "ose_kararGerekce", "MOTOR!$G$35")
    ad_ekle(wb, "ose_maddeAtifMetni", "MOTOR!$G$52")
    ad_ekle(wb, "ose_kanitRaporu", "MOTOR!$G$53")
    ad_ekle(wb, "ose_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "ose_parmakIzi", "AYARLAR!$B$19")

    ad_ekle(wb, "ose_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "ose_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "ose_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "ose_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "ose_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "ose_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "ose_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "ose_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "ose_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "ose_vakaFark", "VAKALAR!$B$13")

    ayar_map = {
        "ose_tolerans": 8,
        "ose_kacirilanEsik": 9,
        "ose_senaryoIyi": 10,
        "ose_senaryoKotu": 11,
        "ose_oranArtisPuan": 12,
        "ose_tahminAlt": 13,
        "ose_tahminUst": 14,
        "ose_tornadoHasilat": 15,
        "ose_tornadoMaliyet": 16,
        "ose_payArtisOran": 17,
        "ose_yuzdelikOran": 18,
        "ose_girisBeklenen": 21,
        "ose_riskDusuk": 22,
        "ose_riskOrta": 23,
        "ose_riskYuksek": 24,
        "ose_riskEsikOran": 25,
        "ose_yilTaban": 26,
        "ose_yuvarMax": 27,
        "ose_dayanakCarpan": 28,
        "ose_bazCarpan": 29,
        "ose_sifirCarpan": 30,
        "ose_siraBir": 31,
        "ose_siraIki": 32,
        "ose_siraUc": 33,
        "ose_tahminX": 34,
        "ose_esikSikilastirma": 35,
        "vaka_ozkaynak": 36,
        "vaka_borc": 37,
        "vaka_faiz": 38,
        "vaka_uygulanan": 39,
        "vaka_emsal": 40,
        "vaka_islem": 41,
        "vaka_bedel": 42,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "ose_modulT1": (11, "C"), "ose_modulT1Yorum": (11, "D"),
        "ose_modulT2": (12, "C"), "ose_modulT2Yorum": (12, "D"),
        "ose_modulO1": (13, "C"), "ose_modulO1Yorum": (13, "D"),
        "ose_modulO6": (16, "C"), "ose_modulO6Yorum": (16, "D"),
        "ose_modulO8": (17, "C"), "ose_modulO8Yorum": (17, "D"),
        "ose_modulI2": (20, "C"), "ose_modulI2Yorum": (20, "D"),
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
    varsayilan = os.path.join(kok, "OrtuluSermayeEmsalFaizTfPaketi.xlsx")
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
    cikti = os.path.join(repo, "cikti", "OrtuluSermayeEmsalFaizTfPaketi.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
