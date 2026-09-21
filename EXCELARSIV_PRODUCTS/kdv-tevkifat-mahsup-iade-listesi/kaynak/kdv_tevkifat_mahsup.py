#!/usr/bin/env python3
"""
KDV Tevkifat Mahsup ve KDV İade Listesi — üretim betiği (A1 / manda v6).
Ayırt edici: 2 Numaralı beyanname tevkifat mahsup listesi + iade satır cetveli.
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

URUN_AD = "KDV Tevkifat Mahsup ve KDV İade Listesi"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 14)

# Demo 2026: matrah×KDV×tevkifat oranları → toplam tevkifat 168.000
# KDV1 ödenecek 100.000 → mahsup 100.000; iade uygun 88.000 → İADE
DEMO = {
    "mukellef_adi": "Örnek Ticaret A.Ş.",
    "donem": "2026-06",
    "hesapYili": 2026,
    "matrah_2_10": 1_000_000,
    "matrah_5_10": 500_000,
    "matrah_7_10": 300_000,
    "matrah_9_10": 200_000,
    "kdv1_odenecek": 100_000,
    "onceki_mahsup_bakiyesi": 20_000,
    "iade_talep": 80_000,
    "mahsup_yorumu": "A",
}


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: Tevkifat mahsup/iade kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Mahsup listesi ve iade cetveli bozulur | Kaynak: KDVK / KDV2 tebliğ",
    )


def kural_cek(kural_id: str) -> str:
    """Oranları tblKurallar'dan çeker; yıl seçimi ktm_yilTaban ile."""
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-ktm_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(ktm_yuvarMax,IFERROR(N(G13),0)))"


def _guvenli_formul(formul: str, birim: str) -> str:
    if not formul.startswith("="):
        formul = "=" + formul
    inner = formul[1:]
    if inner.upper().startswith("IFERROR("):
        return formul
    yedek = '""' if birim in ("metin", "kod") else "0"
    return f"=IFERROR({inner},{yedek})"


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Dönem ve tevkifat girdileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Mahsup/iade tutarları formülle üretilir.",
      yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Mükellef Adı", DEMO["mukellef_adi"], None, "mukellef_adi",
         "Mükellefin resmi unvanını yazın."),
        (7, "Dönem", DEMO["donem"], None, "donem",
         "KDV beyan dönemi etiketini yazın (ör. 2026-06)."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili",
         "2025, 2026 veya 2027 seçin."),
        (9, "2/10 Matrah [₺]", DEMO["matrah_2_10"], TL, "matrah_2_10",
         "2/10 tevkifata tabi işlem matrahını TL girin."),
        (10, "5/10 Matrah [₺]", DEMO["matrah_5_10"], TL, "matrah_5_10",
         "5/10 tevkifata tabi işlem matrahını TL girin."),
        (11, "7/10 Matrah [₺]", DEMO["matrah_7_10"], TL, "matrah_7_10",
         "7/10 tevkifata tabi işlem matrahını TL girin."),
        (12, "9/10 Matrah [₺]", DEMO["matrah_9_10"], TL, "matrah_9_10",
         "9/10 tevkifata tabi işlem matrahını TL girin."),
        (13, "KDV1 Ödenecek [₺]", DEMO["kdv1_odenecek"], TL, "kdv1_odenecek",
         "1 Numaralı beyannamede ödenecek KDV tutarını girin."),
        (14, "Önceki Mahsup Bakiyesi [₺]", DEMO["onceki_mahsup_bakiyesi"], TL,
         "onceki_mahsup_bakiyesi", "Devreden tevkifat mahsup bakiyesini girin."),
        (15, "İade Talep [₺]", DEMO["iade_talep"], TL, "iade_talep",
         "Talep edilen iade tutarını girin (0 olabilir)."),
        (16, "Mahsup Yorumu (A/B)", DEMO["mahsup_yorumu"], None, "mahsup_yorumu",
         "Mahsup sırası yorumunda A veya B seçin."),
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

    h(ws, 18, 1, "Toplam Tevkifat (ayna) [₺]", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 2, "=IFERROR(ktm_toplamTevkifat,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B16)=0,ktm_girisBeklenen=0),0,'
      'ROUND(COUNTA(B6:B16)/ktm_girisBeklenen*100,0)),0)',
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
    tablo_ekle(ws, "tblGirdi", "A23:F122", sutunlar, formuller)
    ornek_satir = [
        ("mukellef_adi", DEMO["mukellef_adi"], "metin", "Evet", "GIRDI"),
        ("donem", DEMO["donem"], "metin", "Evet", "GIRDI"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Evet", "GIRDI"),
        ("matrah_2_10", DEMO["matrah_2_10"], "TL", "Evet", "GIRDI"),
        ("matrah_5_10", DEMO["matrah_5_10"], "TL", "Evet", "GIRDI"),
        ("matrah_7_10", DEMO["matrah_7_10"], "TL", "Evet", "GIRDI"),
        ("matrah_9_10", DEMO["matrah_9_10"], "TL", "Evet", "GIRDI"),
        ("kdv1_odenecek", DEMO["kdv1_odenecek"], "TL", "Evet", "GIRDI"),
        ("onceki_mahsup_bakiyesi", DEMO["onceki_mahsup_bakiyesi"], "TL", "Evet", "GIRDI"),
        ("iade_talep", DEMO["iade_talep"], "TL", "Evet", "GIRDI"),
        ("mahsup_yorumu", DEMO["mahsup_yorumu"], "A/B", "Evet", "GIRDI"),
    ]
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 24):
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
    dogrulama(ws, "list", "ListeYorumAB", "B16",
              baslik="Yorum", mesaj="Mahsup sırası yorumunda A veya B seçin.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B seçin.", bos=False)
    dogrulama(ws, "list", "ListeYillar", "B26",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeYorumAB", "B34",
              baslik="Yorum", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.")
    for aralik, baslik in [
        ("B9", "2/10 Matrah"), ("B10", "5/10 Matrah"), ("B11", "7/10 Matrah"),
        ("B12", "9/10 Matrah"), ("B13", "KDV1 Ödenecek"), ("B14", "Önceki Bakiye"),
        ("B15", "İade Talep"),
        ("B27", "2/10"), ("B28", "5/10"), ("B29", "7/10"), ("B30", "9/10"),
        ("B31", "KDV1"), ("B32", "Bakiye"), ("B33", "İade"),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj="0 ile 1e12 arasında TL.",
                  hata_baslik="Geçersiz tutar", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000000")

    giris_hucreleri(ws, 6, 16, [2])
    giris_hucreleri(ws, 24, 34, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 40, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1024, "Sarı alanlar giriş; mahsup/iade formülleri kilitlidir.")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Tevkifat kural tablosu (yıl yan yana)", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.",
      yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["KTM-001", "KDVK / tebliğ", "tevkifat_2_10", "matrah_2_10>0", "tevkifat_2_10",
         0.20, 0.20, 0.20, "01.01.2025", "2/10 tevkifat oranı"],
        ["KTM-002", "KDVK / tebliğ", "tevkifat_5_10", "matrah_5_10>0", "tevkifat_5_10",
         0.50, 0.50, 0.50, "01.01.2025", "5/10 tevkifat oranı"],
        ["KTM-003", "KDVK / tebliğ", "tevkifat_7_10", "matrah_7_10>0", "tevkifat_7_10",
         0.70, 0.70, 0.70, "01.01.2025", "7/10 tevkifat oranı"],
        ["KTM-004", "KDVK / tebliğ", "tevkifat_9_10", "matrah_9_10>0", "tevkifat_9_10",
         0.90, 0.90, 0.90, "01.01.2025", "9/10 tevkifat oranı"],
        ["KTM-005", "KDVK genel", "kdv_genel_oran", "her hesap", "kdv_genel_oran",
         0.20, 0.20, 0.20, "01.01.2025", "Genel KDV oranı + yuvarlama"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = YÜZDE if k in (6, 7, 8) and isinstance(v, float) else None
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-ktm_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "2/10 oranı (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("KTM-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "5/10 oranı (aktif yıl)", yazi="333333")
    h(ws, 16, 2, kural_cek("KTM-002"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"KTM-001=",TEXT(B15,"0.0%"),"KTM-002=",TEXT(B16,"0.0%"),'
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
    h(ws, 3, 1, "Hesap zinciri — her adımda kural_id", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Tevkifat oranları tblKurallar'dan çekilir; sabit oran yoktur.",
      yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")

    basliklar = [*MOTOR_BASLIK, "deger"]
    baslik_satiri(ws, 6, basliklar)

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif mevzuat yılı", "KTM-005")
    ekle("2/10 tevkifat oranı", km("KTM-001"), "oran", "2/10 oranı", "KTM-001")
    ekle("5/10 tevkifat oranı", km("KTM-002"), "oran", "5/10 oranı", "KTM-002")
    ekle("7/10 tevkifat oranı", km("KTM-003"), "oran", "7/10 oranı", "KTM-003")
    ekle("9/10 tevkifat oranı", km("KTM-004"), "oran", "9/10 oranı", "KTM-004")
    ekle("Genel KDV oranı", km("KTM-005"), "oran", "KDV genel oranı", "KTM-005")
    ekle("Yuvarlama ondalık", "=IFERROR(ktm_yuvarMax+N(GIRDI!B9)*0,0)", "adet", "Yuvarlama basamağı", "KTM-005")
    ekle("2/10 matrah", "=GIRDI!B9", "TL", "2/10 ham matrah", "KTM-001")
    ekle("5/10 matrah", "=GIRDI!B10", "TL", "5/10 ham matrah", "KTM-002")
    ekle("7/10 matrah", "=GIRDI!B11", "TL", "7/10 ham matrah", "KTM-003")
    ekle("9/10 matrah", "=GIRDI!B12", "TL", "9/10 ham matrah", "KTM-004")
    ekle("KDV1 ödenecek", "=GIRDI!B13", "TL", "1 Numaralı ödenecek", "KTM-005")
    ekle("Önceki mahsup bakiyesi", "=GIRDI!B14", "TL", "Devreden bakiye", "KTM-005")
    ekle("İade talep", "=GIRDI!B15", "TL", "Talep edilen iade", "KTM-005")
    ekle("2/10 KDV", f"=ROUND(G14*G12,{_YV})", "TL", "2/10 hesaplanan KDV", "KTM-001")
    ekle("5/10 KDV", f"=ROUND(G15*G12,{_YV})", "TL", "5/10 hesaplanan KDV", "KTM-002")
    ekle("7/10 KDV", f"=ROUND(G16*G12,{_YV})", "TL", "7/10 hesaplanan KDV", "KTM-003")
    ekle("9/10 KDV", f"=ROUND(G17*G12,{_YV})", "TL", "9/10 hesaplanan KDV", "KTM-004")
    ekle("2/10 tevkifat", f"=ROUND(G21*G8,{_YV})", "TL", "2/10 tevkifat tutarı", "KTM-001")
    ekle("5/10 tevkifat", f"=ROUND(G22*G9,{_YV})", "TL", "5/10 tevkifat tutarı", "KTM-002")
    ekle("7/10 tevkifat", f"=ROUND(G23*G10,{_YV})", "TL", "7/10 tevkifat tutarı", "KTM-003")
    ekle("9/10 tevkifat", f"=ROUND(G24*G11,{_YV})", "TL", "9/10 tevkifat tutarı", "KTM-004")
    ekle("Dönem tevkifat toplamı", "=G25+G26+G27+G28", "TL", "Dönem tevkifatı", "KTM-005")
    ekle("Mahsup havuzu", "=G29+G19", "TL", "Dönem + önceki bakiye", "KTM-005")
    ekle("Mahsup (yorum A)", "=MIN(G30,G18)", "TL", "Önce mahsup sonra iade", "KTM-005")
    ekle("İade uygun (yorum A)", "=MAX(0,G30-G31)", "TL", "Mahsup sonrası iade", "KTM-005")
    ekle("İade (yorum B)", "=MIN(G30,G20)", "TL", "Önce iade sonra mahsup", "KTM-005")
    ekle("Mahsup (yorum B)", "=MAX(0,G30-G33)", "TL", "İade sonrası mahsup", "KTM-005")
    ekle("Seçili mahsup",
         '=IF(GIRDI!B16="B",G34,G31)',
         "TL", "Yoruma göre mahsup", "KTM-005")
    ekle("Seçili iade uygun",
         '=IF(GIRDI!B16="B",G33,G32)',
         "TL", "Yoruma göre iade", "KTM-005")
    ekle("İade sapması", "=G36-G20", "TL", "Uygun − talep", "KTM-005")
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(AND(G36>=ktm_iadeEsik,G36>0),"IADE",'
         'IF(G35>0,"MAHSUP","BEKLE")))',
         "kod", "Karar kodu", "KTM-005")
    ekle("Karar metni",
         '=IF(G38="VERI_YOK","VERİ YOK",'
         'IF(G38="IADE","İADE",'
         'IF(G38="MAHSUP","MAHSUP","BEKLE")))',
         "metin", "Kullanıcıya karar", "KTM-005")
    ekle("Gerekçe",
         '=IF(G38="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G38="IADE","İade uygun tutar eşik üzerinde; iade listesi öncelikli.",'
         'IF(G38="MAHSUP","Mahsup edilebilir tutar var; 2 Numaralı mahsup listesi öncelikli.",'
         '"Mahsup ve iade net değil; belge/oran belirsizliği — bekleyin.")))',
         "metin", "Gerekçe cümlesi", "KTM-005")
    ekle("Tevkifat / matrah oranı",
         "=IF((G14+G15+G16+G17)=0,0,G29/(G14+G15+G16+G17))",
         "oran", "Ortalama tevkifat yoğunluğu", "KTM-005")
    ekle("Mahsup / havuz", "=IF(G30=0,0,G35/G30)", "oran", "Mahsup karşılama", "KTM-005")
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B16)=0,ktm_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B16)/ktm_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "KTM-005")
    ekle("Risk skoru",
         '=IF(G38="VERI_YOK",0,IF(G38="BEKLE",ktm_riskYuksek,'
         'IF(G38="IADE",ktm_riskOrta,ktm_riskDusuk)))',
         "puan", "Karar riski", "KTM-005")
    ekle("Senaryo iyimser matrah çarpanı", "=IFERROR(ktm_senaryoIyi+N(GIRDI!B9)*0,0)", "oran", "İyimser çarpan", "KTM-005")
    ekle("Senaryo kötümser matrah çarpanı", "=IFERROR(ktm_senaryoKotu+N(GIRDI!B9)*0,0)", "oran", "Kötümser çarpan", "KTM-005")
    ekle("İyimser tevkifat",
         f"=ROUND((G14*G45*G12*G8+G15*G45*G12*G9+G16*G45*G12*G10+G17*G45*G12*G11)+G19,{_YV})",
         "TL", "İyimser havuz", "KTM-001")
    ekle("Kötümser tevkifat",
         f"=ROUND((G14*G46*G12*G8+G15*G46*G12*G9+G16*G46*G12*G10+G17*G46*G12*G11)+G19,{_YV})",
         "TL", "Kötümser havuz", "KTM-001")
    ekle("M-SEN 2/10 +1 puan",
         f"=ROUND(G21*(G8+ktm_oranArtisPuan),{_YV})",
         "TL", "M-SEN 2/10", "KTM-001")
    ekle("M-SEN toplam tevkifat", "=G49+G26+G27+G28", "TL", "Kural değişimi toplam", "KTM-001")
    ekle("M-SEN mahsup", "=MIN(G50+G19,G18)", "TL", "M-SEN mahsup", "KTM-005")
    ekle("M-SEN iade uygun", "=MAX(0,G50+G19-G51)", "TL", "M-SEN iade", "KTM-005")
    ekle("Tahmin üst iade", "=G36*ktm_tahminUst", "TL", "Tahmin üst bant", "KTM-005")
    ekle("Tahmin alt iade", "=G36*ktm_tahminAlt", "TL", "Tahmin alt bant", "KTM-005")
    ekle("Tornado: matrah etkisi",
         f"=ABS(ROUND(G14*ktm_tornadoMatrah*G12*G8+G15*ktm_tornadoMatrah*G12*G9"
         f"+G16*ktm_tornadoMatrah*G12*G10+G17*ktm_tornadoMatrah*G12*G11,{_YV})-G29)",
         "TL", "Matrah ± etki", "KTM-005")
    ekle("Tornado: KDV1 etkisi",
         "=ABS(MIN(G30,G18*ktm_tornadoKdv1)-G35)",
         "TL", "KDV1 ± etki", "KTM-005")
    ekle("Tornado: oran etkisi",
         f"=ABS(ROUND(G21*(G8+ktm_oranArtisPuan)+G26+G27+G28,{_YV})-G29)",
         "TL", "Oran ± etki", "KTM-001")
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"KTM-001→2/10","KTM-002→5/10","KTM-003→7/10","KTM-004→9/10")',
         "metin", "Kanıt atıfları", "KTM-001")
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Tevkifat=",TEXT(G29,"₺ #,##0"),"Mahsup=",TEXT(G35,"₺ #,##0"),'
         '"İade=",TEXT(G36,"₺ #,##0"))',
         "metin", "Rapor özeti", "KTM-005")
    ekle("Kalite skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B16)=0,ktm_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B16)/ktm_girisBeklenen*100),0),0)",
         "puan", "Giriş kalitesi", "KTM-005")

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


def mahsup_liste(ws):
    sayfa_hazirla(ws, "MAHSUP_LISTE", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "2 Numaralı beyanname — tevkifat mahsup listesi", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Satır bazlı mahsup cetveli; oranlar kural_id ile MOTOR'dan gelir.",
      yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:H3")
    ws.merge_cells("A4:H4")

    sutunlar = ["Satır Numarası", "Oran Kodu", "Matrah", "KDV", "Tevkifat Oranı",
                "Tevkifat Tutarı", "Mahsup Edilebilir", "Kural", "Kontrol"]
    baslik_satiri(ws, 6, sutunlar)
    # Ayırt edici liste: Excel Tablosu yok (ölçek şişmesi); canlı özet MOTOR'dan

    satirlar = [
        (1, "2/10", 1_000_000, 200_000, 0.20, 40_000, 23_810, "KTM-001"),
        (2, "5/10", 500_000, 100_000, 0.50, 50_000, 29_762, "KTM-002"),
        (3, "7/10", 300_000, 60_000, 0.70, 42_000, 25_000, "KTM-003"),
        (4, "9/10", 200_000, 40_000, 0.90, 36_000, 21_429, "KTM-004"),
    ]
    for i, (no, kod, mat, kdv, oran, tevk, mah, kid) in enumerate(satirlar, 7):
        h(ws, i, 1, no, sayi=CATI, zemin=GIRIS_SARI)
        h(ws, i, 2, kod, zemin=GIRIS_SARI)
        h(ws, i, 3, mat, sayi=TL, zemin=GIRIS_SARI)
        h(ws, i, 4, kdv, sayi=TL, zemin=GIRIS_SARI)
        h(ws, i, 5, oran, sayi=YÜZDE, zemin=GIRIS_SARI)
        h(ws, i, 6, tevk, sayi=TL, zemin=GIRIS_SARI)
        h(ws, i, 7, mah, sayi=TL, zemin=GIRIS_SARI)
        h(ws, i, 8, kid, yazi="B08948", kalin=True, zemin=GIRIS_SARI)
        h(ws, i, 9, "TAM", zemin=GIRIS_SARI)
        for kolon in range(1, 10):
            ws.cell(i, kolon).protection = Protection(locked=False)
            _kim(ws, f"{get_column_letter(kolon)}{i}", sutunlar[kolon - 1], "Mahsup satırı — MOTOR ile hizalayın")

    for i in range(11, 56):
        for kolon in range(1, 10):
            ws.cell(i, kolon).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
            ws.cell(i, kolon).protection = Protection(locked=False)
            if ws.cell(i, kolon).comment is None:
                _kim(ws, f"{get_column_letter(kolon)}{i}", sutunlar[kolon - 1], "Mahsup satırı")

    h(ws, 12, 10, "Canlı mahsup (MOTOR)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 11, "=IFERROR(ktm_mahsupToplam,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 13, 10, "Liste metni", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 11,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"2/10=",TEXT(F7,"₺ #,##0"),"5/10=",TEXT(F8,"₺ #,##0"),'
      '"7/10=",TEXT(F9,"₺ #,##0"),"9/10=",TEXT(F10,"₺ #,##0"),"Mahsup=",TEXT(K12,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 6, 13, "TevkifatDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([40000, 50000, 42000, 36000], 7):
        h(ws, i, 13, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Mahsup Listesi Tevkifat Dağılımı"
    g.add_data(Reference(ws, min_col=13, min_row=6, max_row=10), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=2, min_row=7, max_row=10))
    g.height, g.width = 8, 12
    ws.add_chart(g, "A15")

    genislik(ws, {"A": 14, "B": 12, "C": 16, "D": 14, "E": 14, "F": 16, "G": 16, "H": 12, "J": 22, "K": 55})
    sabitle(ws, "A7")
    giris_hucreleri(ws, 7, 55, [1, 2, 3, 4, 5, 6, 7, 8, 9])
    # Kontrol (I) tablo formülü — kilitli kalsın


def iade_liste(ws):
    sayfa_hazirla(ws, "IADE_LISTE", "C45911", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "KDV iade satır cetveli", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Mahsup sonrası iade uygun tutar ve talep sapması.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")

    sutunlar = ["Satır Numarası", "Kalem", "Tutar", "Kaynak", "Durum", "Not", "Kontrol"]
    baslik_satiri(ws, 6, sutunlar)
    # Ayırt edici iade cetveli: tablo değil; canlı özet MOTOR'dan

    satirlar = [
        (1, "İade uygun (seçili)", 88_000, "MOTOR", "İADE ADAYI", "Mahsup sonrası", "TAM"),
        (2, "İade talep", 80_000, "GIRDI", "TALEP VAR", "Kullanıcı girişi", "TAM"),
        (3, "Sapma (uygun−talep)", 8_000, "MOTOR", "FAZLA UYGUN", "Sapma", "TAM"),
        (4, "Mahsup edilen", 100_000, "MOTOR", "MAHSUP VAR", "Öncelik A/B", "TAM"),
        (5, "Havuz bakiyesi", 188_000, "MOTOR", "HAVUZ DOLU", "Dönem+önceki", "TAM"),
    ]
    for i, (no, kalem, tutar, kaynak, durum, not_, kontrol) in enumerate(satirlar, 7):
        h(ws, i, 1, no, sayi=CATI, zemin=GIRIS_SARI)
        h(ws, i, 2, kalem, zemin=GIRIS_SARI)
        h(ws, i, 3, tutar, sayi=TL, zemin=GIRIS_SARI)
        h(ws, i, 4, kaynak, zemin=GIRIS_SARI)
        h(ws, i, 5, durum, zemin=GIRIS_SARI)
        h(ws, i, 6, not_, zemin=GIRIS_SARI)
        h(ws, i, 7, kontrol, zemin=GIRIS_SARI)
        for kolon in range(1, 8):
            ws.cell(i, kolon).protection = Protection(locked=False)
            _kim(ws, f"{get_column_letter(kolon)}{i}", sutunlar[kolon - 1], "İade cetveli — MOTOR ile hizalayın")

    for i in range(12, 56):
        for kolon in range(1, 8):
            ws.cell(i, kolon).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
            ws.cell(i, kolon).protection = Protection(locked=False)
            if ws.cell(i, kolon).comment is None:
                _kim(ws, f"{get_column_letter(kolon)}{i}", sutunlar[kolon - 1], "İade satırı")

    h(ws, 13, 9, "Canlı iade (MOTOR)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 10, "=IFERROR(ktm_iadeUygun,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 14, 9, "Liste metni", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 10,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"Uygun=",TEXT(J13,"₺ #,##0"),"Talep=",TEXT(C8,"₺ #,##0"),'
      '"Sapma=",TEXT(C9,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 6, 12, "IadeDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([88000, 80000, 8000, 100000, 188000], 7):
        h(ws, i, 12, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "İade Cetveli Tutarları"
    g.add_data(Reference(ws, min_col=12, min_row=6, max_row=11), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=2, min_row=7, max_row=11))
    g.height, g.width = 8, 12
    ws.add_chart(g, "A16")

    genislik(ws, {"A": 14, "B": 28, "C": 16, "D": 12, "E": 16, "F": 18, "G": 12, "I": 20, "J": 55})
    sabitle(ws, "A7")
    giris_hucreleri(ws, 7, 55, [1, 2, 3, 4, 5, 6, 7])


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve duyarlılık", kalin=True, boyut=13, yazi=KOYU_LACIVERT)

    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Matrah Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Mahsup", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "İade Uygun", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    mah = "IFERROR(ktm_mahsupToplam,0)"
    iad = "IFERROR(ktm_iadeUygun,0)"
    senaryolar = [
        ("İyimser", "=IFERROR(ktm_senaryoIyi+N(GIRDI!B9)*0,0)",
         f"=IFERROR(ROUND(({mah})*B6,2),0)",
         f"=IFERROR(ROUND(({iad})*B6,2),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D6>=ktm_iadeEsik,"İADE",IF(C6>0,"MAHSUP","BEKLE")))'),
        ("Baz", "=IFERROR(1+N(GIRDI!B9)*0,1)",
         f"=IFERROR(ROUND(({mah})*B7,2),0)",
         f"=IFERROR(ROUND(({iad})*B7,2),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D7>=ktm_iadeEsik,"İADE",IF(C7>0,"MAHSUP","BEKLE")))'),
        ("Kötümser", "=IFERROR(ktm_senaryoKotu+N(GIRDI!B9)*0,0)",
         f"=IFERROR(ROUND(({mah})*B8,2),0)",
         f"=IFERROR(ROUND(({iad})*B8,2),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D8>=ktm_iadeEsik,"İADE",IF(C8>0,"MAHSUP","BEKLE")))'),
        ("M-SEN +1 puan", "=IFERROR(1+N(GIRDI!B9)*0,1)",
         "=IFERROR(MOTOR!G51,0)",
         "=IFERROR(MOTOR!G52,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D9>=ktm_iadeEsik,"İADE",IF(C9>0,"MAHSUP","BEKLE")))'),
    ]
    for i, (ad, carp, asg, fark, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, asg, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, fark, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant genişliği (iyimser−kötümser iade)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(D6-D8,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser iade ",TEXT(D6,"₺ #,##0")," · Baz ",TEXT(D7,"₺ #,##0"),'
      '" · Kötümser ",TEXT(D8,"₺ #,##0")," · Bant ",TEXT(B11,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:H12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Tornado matrah etkisi", yazi="333333")
    h(ws, 16, 2, "=IFERROR(MOTOR!G55,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado KDV1 etkisi", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G56,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado oran etkisi", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G57,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 3,
      '=IFERROR(IF(B16>=MAX(B16:B18),1,IF(B16>=MIN(B16:B18)+ABS(B16-B17),2,3)),3)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 3,
      '=IFERROR(IF(B17>=MAX(B16:B18),1,IF(B17=MEDIAN(B16:B18),2,3)),3)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 3,
      '=IFERROR(IF(B18>=MAX(B16:B18),1,IF(B18=MEDIAN(B16:B18),2,3)),3)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(CONCATENATE("1)",INDEX(A16:A18,MATCH(1,C16:C18,0))," 2)",'
      'INDEX(A16:A18,MATCH(2,C16:C18,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B16:B18),"₺ #,##0"),'
      '" TL — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN kural değişimi iade", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=D9", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"2/10 oranı +1 puan senaryosunda iade uygun ",'
      'TEXT(D9,"₺ #,##0")," TL")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı (iade)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=MOTOR!G54", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 3, "=ktm_iadeUygun", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 4, "=MOTOR!G53", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Tahmin (FORECAST)", kalin=True, yazi=KOYU_LACIVERT)
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

    h(ws, 5, 7, "IadeDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([79200, 88000, 96800, 91000], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([8400, 5000, 4000], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)
    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([79200, 88000, 96800, 91000], 28):
        h(ws, i, 4, v, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo İade Karşılaştırması"
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
    g3.title = "Senaryo İade Trendi"
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

    # V-001: 1_000_000 * 0.20 * 0.20 = 40_000
    # V-002: 500_000 * 0.20 * 0.50 = 50_000
    # V-003: 40_000 + 50_000 = 90_000
    # V-004: 200_000 * 0.20 * 0.90 = 36_000
    vakalar_data = [
        ("V-001", "Tebliğ örnek 1 — 2/10 tevkifat",
         "matrah=vaka_matrah_1; oran KTM-001×KTM-005",
         40000,
         "=IFERROR(ROUND(vaka_matrah_1*INDEX(tblKurallar[deger_2025],MATCH(\"KTM-005\",tblKurallar[kural_id],0))"
         "*INDEX(tblKurallar[deger_2025],MATCH(\"KTM-001\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=ktm_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "2/10 tevkifat örneği"),
        ("V-002", "Tebliğ örnek 2 — 5/10 tevkifat",
         "matrah=vaka_matrah_2; oran KTM-002×KTM-005",
         50000,
         "=IFERROR(ROUND(vaka_matrah_2*INDEX(tblKurallar[deger_2025],MATCH(\"KTM-005\",tblKurallar[kural_id],0))"
         "*INDEX(tblKurallar[deger_2025],MATCH(\"KTM-002\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=ktm_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "5/10 tevkifat örneği"),
        ("V-003", "Tebliğ örnek 3 — 2/10+5/10 birikim",
         "vaka_matrah_1 + vaka_matrah_2",
         90000,
         "=IFERROR(ROUND(vaka_matrah_1*INDEX(tblKurallar[deger_2025],MATCH(\"KTM-005\",tblKurallar[kural_id],0))"
         "*INDEX(tblKurallar[deger_2025],MATCH(\"KTM-001\",tblKurallar[kural_id],0))"
         "+vaka_matrah_2*INDEX(tblKurallar[deger_2025],MATCH(\"KTM-005\",tblKurallar[kural_id],0))"
         "*INDEX(tblKurallar[deger_2025],MATCH(\"KTM-002\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=ktm_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Birikimli tevkifat"),
        ("V-004", "9/10 tevkifat kontrol",
         "matrah=vaka_matrah_9; oran KTM-004×KTM-005",
         36000,
         "=IFERROR(ROUND(vaka_matrah_9*INDEX(tblKurallar[deger_2025],MATCH(\"KTM-005\",tblKurallar[kural_id],0))"
         "*INDEX(tblKurallar[deger_2025],MATCH(\"KTM-004\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=ktm_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "9/10 tevkifat kontrol"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v,
              yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = TL

    formuller = {
        "fark": (
            "=IF(tblVakalar[[#This Row],[vaka_id]]=\"\",\"\","
            "IFERROR(tblVakalar[[#This Row],[beklenen_sonuc]]-"
            "tblVakalar[[#This Row],[hesaplanan]],0))"
        ),
        "durum": (
            '=IF(tblVakalar[[#This Row],[vaka_id]]="","",'
            'IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=ktm_tolerans,"TUTARLI","KIRIK"),"KIRIK"))'
        ),
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
    for i, (b, he) in enumerate([(40000, 40000), (50000, 50000), (90000, 90000), (36000, 36000)], 6):
        h(ws, i, 10, b, sayi=TL, yazi=GRİ)
        h(ws, i, 11, he, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Vaka Beklenen vs Hesaplanan"
    g.add_data(Reference(ws, min_col=10, min_row=5, max_col=11, max_row=9), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=6, max_row=9))
    g.height, g.width = 8, 14
    ws.add_chart(g, "A15")

    genislik(ws, {"A": 12, "B": 36, "C": 42, "D": 14, "E": 14, "F": 12, "G": 12, "H": 28})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", "0F2742", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "KANIT RAPORU — KDV Tevkifat Mahsup / İade", kalin=True, boyut=14,
      yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:F3")
    h(ws, 4, 1, "Rapor tarihi", yazi="333333")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 5, 1, "Mükellef", yazi="333333")
    h(ws, 5, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Dönem / yıl", yazi="333333")
    h(ws, 6, 2, '=CONCATENATE(GIRDI!B7," / ",hesapYili)', yazi=DIS_REF_YEŞIL)

    h(ws, 8, 1, "Girdi özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, (et, form, sayi) in enumerate([
        ("2/10 matrah", "=GIRDI!B9", TL),
        ("5/10 matrah", "=GIRDI!B10", TL),
        ("7/10 matrah", "=GIRDI!B11", TL),
        ("9/10 matrah", "=GIRDI!B12", TL),
        ("KDV1 ödenecek", "=GIRDI!B13", TL),
        ("Önceki mahsup bakiyesi", "=GIRDI!B14", TL),
        ("İade talep", "=GIRDI!B15", TL),
    ], 9):
        h(ws, i, 1, et, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)

    h(ws, 17, 1, "Hesap zinciri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 18, 1, "Toplam tevkifat", yazi="333333")
    h(ws, 18, 2, "=ktm_toplamTevkifat", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Mahsup toplam", yazi="333333")
    h(ws, 19, 2, "=ktm_mahsupToplam", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "İade uygun", yazi="333333")
    h(ws, 20, 2, "=ktm_iadeUygun", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 21, 1, "KARAR", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 2, "=ktm_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 22, 1, "Gerekçe", yazi="333333")
    h(ws, 22, 2, "=ktm_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B22:F22")

    h(ws, 24, 1, "Madde atıfları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, "=ktm_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B24:F24")

    h(ws, 26, 1, "Kanıt gövdesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 2, "=ktm_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B26:F26")

    h(ws, 28, 1, "Parmak izi (SHA-256 kısaltma)", yazi="333333")
    h(ws, 28, 2, "=ktm_parmakIzi", yazi=GRİ)

    h(ws, 30, 1, "Hazırlayan (imza)", yazi="333333")
    h(ws, 30, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B30", "İmza", "Raporu hazırlayan kişinin adını yazın.")
    h(ws, 31, 1, "Onaylayan (imza)", yazi="333333")
    h(ws, 31, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B31", "Onay", "Onaylayan kişinin adını yazın.")

    h(ws, 33, 1,
      "Uyarı: Bu çıktı karar destektir; kesin KDV tahakkuku veya hukuki görüş değildir. "
      f"Sürüm {SURUM}.",
      yazi=GRİ, boyut=9, kaydir=True)
    ws.merge_cells("A33:F33")

    baski_hazirla(ws, "A1:F34", f"{URUN_AD} · {SURUM}")
    ws.page_setup.orientation = "portrait"
    genislik(ws, {"A": 28, "B": 22, "C": 14, "D": 14, "E": 14, "F": 14})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim paneli — KDV Tevkifat Mahsup / İade", kalin=True, boyut=14,
      yazi=KOYU_LACIVERT)
    genislik(ws, {"A": 44})

    kpiler = [
        (2, "Tevkifat", "=IFERROR(ktm_toplamTevkifat,0)", TL),
        (3, "Mahsup", "=IFERROR(ktm_mahsupToplam,0)", TL),
        (4, "İade", "=IFERROR(ktm_iadeUygun,0)", TL),
        (5, "Talep", "=IFERROR(GIRDI!B15,0)", TL),
        (6, "Güven", "=IFERROR(MOTOR!G43,0)", CATI),
        (7, "Risk", "=IFERROR(MOTOR!G44,0)", CATI),
        (8, "2/10 oran", "=IFERROR(MOTOR!G8,0)", YÜZDE),
        (9, "5/10 oran", "=IFERROR(MOTOR!G9,0)", YÜZDE),
        (10, "Mahsup/Havuz", "=IFERROR(MOTOR!G42,0)", YÜZDE),
        (11, "M-SEN iade", "=IFERROR(ktm_kuralDegisimSenaryo,0)", TL),
        (12, "Tahmin", "=IFERROR(SENARYO!B32,0)", TL),
        (13, "Vaka", "=IFERROR(ktm_vakaDurum,\"-\")", None),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=9, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=ktm_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=ktm_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 10, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Mahsup toplam",
         "=IFERROR(ktm_mahsupToplam,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Mahsup ",TEXT(C11,"₺ #,##0")," TL — 2 Numaralı liste")'),
        ("T2", "İade toplam",
         "=IFERROR(ktm_iadeUygun,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İade uygun ",TEXT(C12,"₺ #,##0")," TL")'),
        ("O1", "İade sapması",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,STDEV.P(SENARYO!D6:D9)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo iade sapması ",TEXT(C13,"₺ #,##0")," TL")'),
        ("O2", "Tornado max etki",
         "=IFERROR(ktm_duyarlilikSira,\"-\")",
         "=IFERROR(ktm_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo bant",
         "=IFERROR(ktm_senaryoKarsilastirma,0)",
         "=IFERROR(ktm_yorumSenaryo,\"-\")"),
        ("O6", "2/10 tevkifat payı",
         "=IFERROR(IF(ktm_toplamTevkifat=0,0,MOTOR!G25/ktm_toplamTevkifat),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"2/10 / toplam tevkifat ",TEXT(C16,"0.0%"))'),
        ("O8", "Kalite skoru",
         "=IFERROR(MOTOR!G60,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalite ",TEXT(C17,"0")," puan")'),
        ("M", "M-SEN iade",
         "=IFERROR(ktm_kuralDegisimSenaryo,0)",
         "=IFERROR(ktm_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin aralık",
         "=IFERROR(ktm_tahminAralik,0)",
         "=IFERROR(ktm_yorumTahmin,\"-\")"),
        ("I2", "Senaryo iade P90",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,"
         "_xlfn.PERCENTILE.INC(SENARYO!D6:D9,ktm_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"P90 iade ",TEXT(C20,"₺ #,##0")," TL")'),
    ]
    for i, (kod, ad, form, yorum) in enumerate(moduller):
        r = 11 + i
        h(ws, r, 1, kod, kalin=True, hiza="center", yazi="B08948")
        h(ws, r, 2, ad, kaydir=True)
        h(ws, r, 3, form, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[r].height = 28
        if kod == "O6":
            ws.cell(r, 3).number_format = YÜZDE
        elif kod != "O2":
            ws.cell(r, 3).number_format = TL

    h(ws, 23, 1, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("2/10 tevkifat", 40_000),
        ("5/10 tevkifat", 50_000),
        ("7/10 tevkifat", 42_000),
        ("9/10 tevkifat", 36_000),
        ("Toplam tevkifat", 168_000),
    ], 24):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)

    h(ws, 23, 4, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 5, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Mahsup havuzu", 188_000),
        ("Mahsup", 100_000),
        ("İade uygun", 88_000),
        ("İade talep", 80_000),
        ("M-SEN iade", 91_000),
    ], 24):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Tevkifat Oran Kalemleri"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=28), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=28))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G9")

    g2 = BarChart()
    g2.type = "col"
    g2.title = "Mahsup / İade / Talep"
    g2.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "G23")

    g3 = LineChart()
    g3.title = "Mahsup-İade Çizgisi"
    g3.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g3.height, g3.width = 7, 12
    ws.add_chart(g3, "P9")

    h(ws, 30, 1, "Metrik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 2, "Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("İade uygun", 88_000),
        ("Mahsup", 100_000),
        ("Tevkifat", 168_000),
    ], 31):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "KPI İade / Mahsup / Tevkifat"
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
        ("Negatif matrah var mı?",
         '=IF(OR(GIRDI!B9<0,GIRDI!B10<0,GIRDI!B11<0,GIRDI!B12<0),"NEGATİF","TEMİZ")'),
        ("Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(ktm_yilTaban),N(hesapYili)<=N(ktm_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        ("Tevkifat sıfır mı?",
         '=IF(IFERROR(ktm_toplamTevkifat,0)=0,"SIFIR","NORMAL")'),
        ("Vaka kırık mı?",
         '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")'),
        ("İade eşiği aşıyor mu?",
         '=IF(IFERROR(ktm_iadeUygun,0)>=ktm_iadeEsik,"AŞIYOR","NORMAL")'),
        ("Yorum A/B seçili mi?",
         '=IF(OR(GIRDI!B16="A",GIRDI!B16="B"),"TEMİZ","EKSİK")'),
        ("Motor adım sayısı",
         "=COUNTA(MOTOR!B7:B60)"),
    ]
    for i, (ad, form) in enumerate(kontroller_list, 5):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Canlı formül | Değiştirmeyin")

    h(ws, 14, 1, "Toplam giriş alanı", yazi="333333")
    h(ws, 14, 2, "=IFERROR(ktm_girisBeklenen+COUNTA(GIRDI!B6:B16)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Dolu giriş", yazi="333333")
    h(ws, 15, 2, "=COUNTA(GIRDI!B6:B16)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Kalite skoru (modül)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=IFERROR(ROUND(IF(OR(B14=0,B14=\"\"),0,B15/B14*100),0),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)

    h(ws, 17, 1, "Tevkifat kontrol", yazi="333333")
    h(ws, 17, 2, "=IFERROR(ktm_toplamTevkifat,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Mahsup kontrol", yazi="333333")
    h(ws, 18, 2, "=IFERROR(ktm_mahsupToplam,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Karar kontrol", yazi="333333")
    h(ws, 19, 2, "=IFERROR(ktm_kararMetni,\"-\")", yazi=DIS_REF_YEŞIL)

    genislik(ws, {"A": 40, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek senaryo kütüphanesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:I3")
    h(ws, 4, 1, "Bu sayfa demo değerleri saklar; GIRDI'ye kopyalanabilir.", yazi=GRİ)
    sutunlar = ["Senaryo", "2/10 Matrah", "5/10 Matrah", "7/10 Matrah", "9/10 Matrah",
                "KDV1", "Önceki", "İade Talep", "Tevkifat", "Kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Kontrol": (
            '=IF(tblOrnek[[#This Row],[Senaryo]]="","",'
            'IF(tblOrnek[[#This Row],[Tevkifat]]="","EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:J104", sutunlar, formuller)
    ornekler = [
        ("Demo ana", 1_000_000, 500_000, 300_000, 200_000, 100_000, 20_000, 80_000, 168_000),
        ("Yüksek mahsup", 1_000_000, 500_000, 300_000, 200_000, 200_000, 20_000, 10_000, 168_000),
        ("Düşük matrah", 200_000, 100_000, 50_000, 25_000, 30_000, 5_000, 5_000, 33_500),
        ("Sadece 2/10", 800_000, 0, 0, 0, 50_000, 0, 20_000, 32_000),
        ("Bekle senaryo", 100_000, 50_000, 0, 0, 80_000, 0, 0, 15_000),
    ]
    for i, row in enumerate(ornekler, 6):
        for k, v in enumerate(row, 1):
            ws.cell(i, k).value = v
            if k > 1:
                ws.cell(i, k).number_format = TL
            ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
            ws.cell(i, k).protection = Protection(locked=False)
            _kim(ws, f"{get_column_letter(k)}{i}", sutunlar[k - 1], "Örnek senaryo değeri")

    for col, ad in [("B", "2/10"), ("C", "5/10"), ("D", "7/10"), ("E", "9/10"),
                    ("F", "KDV1"), ("G", "Önceki"), ("H", "İade")]:
        dogrulama(ws, "decimal", "0", f"{col}6:{col}104",
                  baslik=ad, mesaj=f"{ad} tutarını TL girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ile 1e12 arasında.",
                  isaret="between", f2="1000000000000")

    giris_hucreleri(ws, 6, 104, [1, 2, 3, 4, 5, 6, 7, 8, 9])
    h(ws, 5, 12, "TevkifatDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([168000, 168000, 33500, 32000, 15000], 6):
        h(ws, i, 12, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Örnek Senaryo Tevkifat"
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
        _kim(ws, f"C{i}", "Yorum", "Mahsup sırası yorumu")
    h(ws, 5, 5, "Evet/Hayır", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate(["Evet", "Hayır"], 6):
        h(ws, i, 5, v, zemin=GIRIS_SARI)
        ws.cell(i, 5).protection = Protection(locked=False)
        _kim(ws, f"E{i}", "Seçim", "Evet veya Hayır")
    genislik(ws, {"A": 28, "C": 14, "E": 14})
    genislik(ws, {"A": 22})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Parametreler (tek kaynak)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = ["anahtar", "deger", "birim", "aciklama", "kaynak",
                "yururluk_tarihi", "dogrulama_tarihi", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    params = [
        ("hesapYili", DEMO["hesapYili"], "yıl", "Aktif hesap yılı", "Kullanıcı / GIRDI",
         "01.01.2025", "14.08.2026"),
        ("raporTarihi", RAPOR_TARIH, "tarih", "Rapor tarihi (sabit)", "Üretim",
         "01.01.2025", "14.08.2026"),
        ("ktm_tolerans", 0.01, "TL", "Vaka tutarlılık toleransı", "Uygulama notu",
         "01.01.2025", "14.08.2026"),
        ("ktm_iadeEsik", 50_000, "TL", "İADE karar eşiği", "İç politika",
         "01.01.2025", "14.08.2026"),
        ("ktm_senaryoIyi", 0.90, "çarpan", "İyimser matrah çarpanı", "Model varsayımı",
         "01.01.2025", "14.08.2026"),
        ("ktm_senaryoKotu", 1.10, "çarpan", "Kötümser matrah çarpanı", "Model varsayımı",
         "01.01.2025", "14.08.2026"),
        ("ktm_oranArtisPuan", 0.01, "oran", "M-SEN +1 puan", "Senaryo motoru",
         "01.01.2025", "14.08.2026"),
        ("ktm_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı",
         "01.01.2025", "14.08.2026"),
        ("ktm_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı",
         "01.01.2025", "14.08.2026"),
        ("ktm_tornadoMatrah", 1.05, "çarpan", "Tornado matrah şoku", "Model varsayımı",
         "01.01.2025", "14.08.2026"),
        ("ktm_tornadoKdv1", 0.95, "çarpan", "Tornado KDV1 şoku", "Model varsayımı",
         "01.01.2025", "14.08.2026"),
        ("ktm_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik",
         "01.01.2025", "14.08.2026"),
        ("ktm_parmakIzi", "KTM-PRO-1.0.0", "metin", "Dosya parmak izi etiketi", "Üretim",
         "01.01.2025", "14.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv",
         "01.01.2025", "14.08.2026"),
        ("ktm_girisBeklenen", 11, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi",
         "01.01.2025", "14.08.2026"),
        ("ktm_riskDusuk", 15, "puan", "Düşük risk skoru", "İç politika",
         "01.01.2025", "14.08.2026"),
        ("ktm_riskOrta", 55, "puan", "Orta risk skoru", "İç politika",
         "01.01.2025", "14.08.2026"),
        ("ktm_riskYuksek", 85, "puan", "Yüksek risk skoru", "İç politika",
         "01.01.2025", "14.08.2026"),
        ("ktm_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı (hesapYili−taban)", "Ürün mimarisi",
         "01.01.2025", "14.08.2026"),
        ("ktm_yuvarMax", 2, "adet", "ROUND basamak tavanı", "Ürün mimarisi",
         "01.01.2025", "14.08.2026"),
        ("vaka_matrah_1", 1_000_000, "TL", "Vaka1 2/10 matrah", "Tebliğ örneği",
         "01.01.2025", "14.08.2026"),
        ("vaka_matrah_2", 500_000, "TL", "Vaka2 5/10 matrah", "Tebliğ örneği",
         "01.01.2025", "14.08.2026"),
        ("vaka_matrah_9", 200_000, "TL", "Vaka4 9/10 matrah", "Tebliğ örneği",
         "01.01.2025", "14.08.2026"),
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
    h(ws, son_satir + 2, 2,
      "Parametre satırları yukarıda; ad tanımları B kolonuna bağlı.",
      kalin=True, yazi=KOYU_LACIVERT, kaydir=True)

    dogrulama(ws, "list", "ListeYillar", "B6",
              baslik="Hesap Yılı", mesaj="Aktif yılı seçin.",
              hata_baslik="Geçersiz", hata_mesaj="2025-2027.")
    dogrulama(ws, "decimal", "0", "B9",
              baslik="İade eşiği", mesaj="Uyarı eşiği TL.",
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
    h(ws, son_satir + 4, 2,
      f"Sürüm {SURUM} | Şifre koruması: {SIFRE} (formül alanları)",
      yazi=GRİ, boyut=9, kaydir=True)


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar kuralları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "VERİ YOK — giriş tablosu boşsa hesap yapılmaz.",
        "İADE — iade uygun tutar eşik üzerinde; iade cetveli öncelikli.",
        "MAHSUP — mahsup edilebilir tutar var; 2 Numaralı mahsup listesi öncelikli.",
        "BEKLE — mahsup/iade net değil veya belge belirsizliği var.",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 10, 1, "Belirsizlik beyanları (M05)", kalin=True, boyut=12, yazi=KRITIK)
    ws.merge_cells("A10:J10")
    for i, m in enumerate(belirsizlik_beyanlari(["mahsup_sirasi"]), 11):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 13, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2,
      "Önce mahsup (KDV1 ödenecek kadar), kalan iadeye gider. "
      "Bu nokta mevzuatta tartışmalıdır; sonuç yorum B ile yan yana okunmalıdır.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B13:J13")
    h(ws, 14, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 2,
      "Önce iade talep kadar iade, kalan mahsup edilir. Dosya her iki yorumun sonucunu MOTOR'da üretir.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B14:J14")

    h(ws, 16, 1, "Kullanım sırası", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 1,
      "GIRDI → KURALLAR → MOTOR → MAHSUP_LISTE → IADE_LISTE → SENARYO → VAKALAR → "
      "PANO → KANIT_RAPORU → AYARLAR → KILAVUZ.",
      kaydir=True, yazi="333333")
    ws.merge_cells("A17:J17")
    genislik(ws, {"A": 36, "B": 70})
    h(ws, 19, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)


def kosullu_bicimlendirme(wb):
    yesil = Font(color=NORMAL, bold=True)
    kirmizi = Font(color=KRITIK, bold=True)
    amber = Font(color="B7791F", bold=True)
    y_fill = PatternFill("solid", fgColor="E2EFDA")
    k_fill = PatternFill("solid", fgColor="FDE9E9")
    a_fill = PatternFill("solid", fgColor="FFF2CC")

    ws = wb["PANO"]
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("MAHSUP",B6))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("İADE",B6))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("BEKLE",B6))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("VERİ YOK",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("D4", CellIsRule(operator="greaterThan", formula=["0"], font=amber))
    ws.conditional_formatting.add("C4", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
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
        ws.conditional_formatting.add(
            f"{col}9:{col}15", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(
            f"{col}9:{col}15", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add(
        "F24:F34", FormulaRule(formula=['F24="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add(
        "F24:F34", FormulaRule(formula=['F24="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add(
        "G6:G9", FormulaRule(formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add(
        "G6:G9", FormulaRule(formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add(
        "F6:F9", CellIsRule(operator="notEqual", formula=["0"], font=amber))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add(
        "D6:D9", CellIsRule(operator="greaterThan", formula=["0"], font=amber))
    ws.conditional_formatting.add(
        "D6:D9", CellIsRule(operator="equal", formula=["0"], font=yesil))
    ws.conditional_formatting.add(
        "E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("İADE",E6))'], font=amber))
    ws.conditional_formatting.add(
        "E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("MAHSUP",E6))'], font=yesil))
    ws.conditional_formatting.add(
        "B16:B18", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add(
        "G7:G60", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add(
        "F7:F60", FormulaRule(formula=['F7="KTM-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add(
        "F7:F60", FormulaRule(formula=['F7="KTM-005"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["MAHSUP_LISTE"]
    ws.conditional_formatting.add(
        "F7:F10", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add(
        "C7:C10", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add(
        "G7:G10", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["IADE_LISTE"]
    ws.conditional_formatting.add(
        "C7:C11", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add(
        "E7:E11", FormulaRule(formula=['ISNUMBER(SEARCH("İADE",E7))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add(
        "E7:E11", FormulaRule(formula=['ISNUMBER(SEARCH("YOK",E7))'], font=kirmizi))
    ws.conditional_formatting.add(
        "G7:G11", FormulaRule(formula=['G7="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add(
        "G7:G11", FormulaRule(formula=['G7="TAM"'], font=yesil, fill=y_fill))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add(
        "B21", FormulaRule(formula=['ISNUMBER(SEARCH("MAHSUP",B21))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add(
        "B21", FormulaRule(formula=['ISNUMBER(SEARCH("İADE",B21))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add(
        "B21", FormulaRule(formula=['ISNUMBER(SEARCH("VERİ YOK",B21))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add(
        "B20", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add(
        "B5:B11", FormulaRule(
            formula=['OR(B5="KIRIK",B5="HATALI",B5="NEGATİF",B5="BOŞ",B5="EKSİK",B5="AŞIYOR")'],
            font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add(
        "B5:B11", FormulaRule(
            formula=['OR(B5="TEMİZ",B5="DOLU",B5="NORMAL",B5="TUTARLI")'],
            font=yesil, fill=y_fill))
    ws.conditional_formatting.add(
        "B12", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add(
        "B12", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "C", "D", "E", "F", "G", "H", "I"):
        ws.conditional_formatting.add(
            f"{harf}6:{harf}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add(
        "I6:I20", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KURALLAR"]
    for col in ("F", "G", "H"):
        ws.conditional_formatting.add(
            f"{col}7:{col}11", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(
            f"{col}7:{col}11", CellIsRule(operator="equal", formula=["0"], font=amber))


def ad_tanimlari(wb):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "ktm_toplamTevkifat", "MOTOR!$G$29")
    ad_ekle(wb, "ktm_mahsupToplam", "MOTOR!$G$35")
    ad_ekle(wb, "ktm_iadeUygun", "MOTOR!$G$36")
    ad_ekle(wb, "ktm_kararMetni", "MOTOR!$G$39")
    ad_ekle(wb, "ktm_kararGerekce", "MOTOR!$G$40")
    ad_ekle(wb, "ktm_kararKapi", "MOTOR!$G$39")
    ad_ekle(wb, "ktm_maddeAtifMetni", "MOTOR!$G$58")
    ad_ekle(wb, "ktm_kanitRaporu", "MOTOR!$G$59")
    ad_ekle(wb, "ktm_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "ktm_mahsupListe", "MAHSUP_LISTE!$K$13")
    ad_ekle(wb, "ktm_parmakIzi", "AYARLAR!$B$18")

    ad_ekle(wb, "ktm_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "ktm_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "ktm_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "ktm_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "ktm_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "ktm_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "ktm_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "ktm_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "ktm_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "ktm_vakaFark", "VAKALAR!$B$13")

    ayar_map = {
        "ktm_tolerans": 8,
        "ktm_iadeEsik": 9,
        "ktm_senaryoIyi": 10,
        "ktm_senaryoKotu": 11,
        "ktm_oranArtisPuan": 12,
        "ktm_tahminAlt": 13,
        "ktm_tahminUst": 14,
        "ktm_tornadoMatrah": 15,
        "ktm_tornadoKdv1": 16,
        "ktm_yuzdelikOran": 17,
        "ktm_girisBeklenen": 20,
        "ktm_riskDusuk": 21,
        "ktm_riskOrta": 22,
        "ktm_riskYuksek": 23,
        "ktm_yilTaban": 24,
        "ktm_yuvarMax": 25,
        "vaka_matrah_1": 26,
        "vaka_matrah_2": 27,
        "vaka_matrah_9": 28,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "ktm_modulT1": (11, "C"), "ktm_modulT1Yorum": (11, "D"),
        "ktm_modulT2": (12, "C"), "ktm_modulT2Yorum": (12, "D"),
        "ktm_modulO1": (13, "C"), "ktm_modulO1Yorum": (13, "D"),
        "ktm_modulO6": (16, "C"), "ktm_modulO6Yorum": (16, "D"),
        "ktm_modulO8": (17, "C"), "ktm_modulO8Yorum": (17, "D"),
        "ktm_modulI2": (20, "C"), "ktm_modulI2Yorum": (20, "D"),
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
        ("PANO", pano),
        ("GIRDI", girdi),
        ("KURALLAR", kurallar),
        ("MOTOR", motor),
        ("MAHSUP_LISTE", mahsup_liste),
        ("IADE_LISTE", iade_liste),
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
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=120)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    wb.calculation.fullCalcOnLoad = True

    kok = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    varsayilan = os.path.join(kok, "KdvTevkifatMahsupIadeListesi.xlsx")
    hedef = cikti_yolu or varsayilan
    os.makedirs(os.path.dirname(os.path.abspath(hedef)) or ".", exist_ok=True)
    wb.save(hedef)

    hsh = hashlib.sha256()
    with open(hedef, "rb") as f:
        for b in iter(lambda: f.read(65536), b""):
            hsh.update(b)
    dig = hsh.hexdigest()

    wb2 = __import__("openpyxl").load_workbook(hedef)
    wb2["AYARLAR"]["B18"] = dig[:16]
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
    cikti = os.path.join(repo, "cikti", "KdvTevkifatMahsupIadeListesi.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
