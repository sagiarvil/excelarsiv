#!/usr/bin/env python3
"""
KDV İade Listesi Robotu (GİB 7 Liste) — üretim betiği (A1 / Hat D D-01 / manda v6).
Ayırt edici: GİB'in istediği 7 liste (İndirilecek/Yüklenilen/GÇB+); tevkifat mahsup değil.
"""

from __future__ import annotations

import hashlib
import os
import shutil
import sys
from datetime import date

KOK = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
if KOK not in sys.path:
    sys.path.insert(0, KOK)

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

URUN_AD = "KDV İade Listesi Robotu (GİB 7 Liste)"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 14)

LISTE_ADLARI = [
    ("LISTE_1", "İndirilecek KDV Listesi"),
    ("LISTE_2", "Yüklenilen KDV Listesi"),
    ("LISTE_3", "GÇB Listesi"),
    ("LISTE_4", "Satış Faturaları Listesi"),
    ("LISTE_5", "Alış Faturaları Listesi"),
    ("LISTE_6", "İade Hesap Cetveli"),
    ("LISTE_7", "Ek Belge / Tamamlayıcı Liste"),
]

# Demo 2026: net=1.2M; A=MIN(1.2M,5M*0.2)=1M → İADE; B=0.9M
DEMO = {
    "mukellef_adi": "Örnek İhracat A.Ş.",
    "donem": "2026-06",
    "hesapYili": 2026,
    "indirilecek_kdv": 800_000,
    "yuklenilen_kdv": 2_000_000,
    "gcb_tutari": 4_500_000,
    "ihrac_hasilati": 5_000_000,
    "belge_eksik_adet": 0,
    "iade_talep": 900_000,
    "liste_yorumu": "A",
}


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: GİB 7 liste iade kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Liste tutarlılığı ve karar kapısı bozulur | Kaynak: KDVK / GİB iade listeleri",
    )


def kural_cek(kural_id: str) -> str:
    """Oranları tblKurallar'dan çeker; yıl seçimi kil_yilTaban ile."""
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-kil_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(kil_yuvarMax,IFERROR(N(GIRDI!B9)*0,0)))"


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
    h(ws, 3, 1, "Dönem ve GİB 7 liste girdileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. İade tutarları ve karar MOTOR'da üretilir.",
      yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Mükellef Adı", DEMO["mukellef_adi"], None, "mukellef_adi",
         "Mükellefin resmi unvanını yazın."),
        (7, "Dönem", DEMO["donem"], None, "donem",
         "KDV iade dönemi etiketini yazın (ör. 2026-06)."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili",
         "2025, 2026 veya 2027 seçin."),
        (9, "İndirilecek KDV [₺]", DEMO["indirilecek_kdv"], TL, "indirilecek_kdv",
         "Dönem indirilecek KDV toplamını TL girin (LISTE_1)."),
        (10, "Yüklenilen KDV [₺]", DEMO["yuklenilen_kdv"], TL, "yuklenilen_kdv",
         "Yüklenilen KDV toplamını TL girin (LISTE_2)."),
        (11, "GÇB Tutarı [₺]", DEMO["gcb_tutari"], TL, "gcb_tutari",
         "Gümrük çıkış beyannamesi hasılatını KDV hariç TL girin (LISTE_3)."),
        (12, "İhraç Hasılatı [₺]", DEMO["ihrac_hasilati"], TL, "ihrac_hasilati",
         "İhraç hasılatını KDV hariç TL girin."),
        (13, "Belge Eksik Adet", DEMO["belge_eksik_adet"], CATI, "belge_eksik_adet",
         "Eksik belge sayısını girin (0 = eksik yok)."),
        (14, "İade Talep [₺]", DEMO["iade_talep"], TL, "iade_talep",
         "Talep edilen iade tutarını TL girin."),
        (15, "Liste Yorumu (A/B)", DEMO["liste_yorumu"], None, "liste_yorumu",
         "Liste sırası belirsizliğinde A veya B seçin."),
    ]
    h(ws, 5, 1, "Alan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 2, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 3, "Birim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    for satir, etiket, deger, sayi, _ad, mesaj in alanlar:
        if _ad == "hesapYili":
            birim = "yıl"
        elif _ad == "belge_eksik_adet":
            birim = "adet"
        elif sayi == TL:
            birim = "₺"
        else:
            birim = "metin"
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim, yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 17, 1, "Seçili iade (ayna) [₺]", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 2, "=IFERROR(kil_iadeSecili,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B15)=0,kil_girisBeklenen=0),0,'
      'ROUND(COUNTA(B6:B15)/kil_girisBeklenen*100,0)),0)',
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
    tablo_ekle(ws, "tblGirdi", "A23:F1023", sutunlar, formuller)
    ornek_satir = [
        ("mukellef_adi", DEMO["mukellef_adi"], "metin", "Evet", "GIRDI"),
        ("donem", DEMO["donem"], "metin", "Evet", "GIRDI"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Evet", "GIRDI"),
        ("indirilecek_kdv", DEMO["indirilecek_kdv"], "TL", "Evet", "GIRDI"),
        ("yuklenilen_kdv", DEMO["yuklenilen_kdv"], "TL", "Evet", "GIRDI"),
        ("gcb_tutari", DEMO["gcb_tutari"], "TL", "Evet", "GIRDI"),
        ("ihrac_hasilati", DEMO["ihrac_hasilati"], "TL", "Evet", "GIRDI"),
        ("belge_eksik_adet", DEMO["belge_eksik_adet"], "adet", "Evet", "GIRDI"),
        ("iade_talep", DEMO["iade_talep"], "TL", "Evet", "GIRDI"),
        ("liste_yorumu", DEMO["liste_yorumu"], "A/B", "Evet", "GIRDI"),
    ]
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 24):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if isinstance(d, (int, float)) and a not in ("hesapYili", "belge_eksik_adet"):
            ws.cell(i, 2).number_format = TL
        elif a in ("hesapYili", "belge_eksik_adet"):
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
              baslik="Yorum", mesaj="Liste sırası yorumunda A veya B seçin.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B seçin.", bos=False)
    dogrulama(ws, "list", "ListeYillar", "B26",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeYorumAB", "B33",
              baslik="Yorum", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.")
    for aralik, baslik in [
        ("B9", "İndirilecek KDV"), ("B10", "Yüklenilen KDV"), ("B11", "GÇB"),
        ("B12", "İhraç Hasılat"), ("B14", "İade Talep"),
        ("B27", "İndirilecek"), ("B28", "Yüklenilen"), ("B29", "GÇB"),
        ("B30", "İhraç"), ("B32", "İade"),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj="0 ile 1e12 arasında TL.",
                  hata_baslik="Geçersiz tutar", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000000")
    dogrulama(ws, "whole", "0", "B13",
              baslik="Belge eksik", mesaj="0 ve üzeri tam sayı.",
              hata_baslik="Geçersiz", hata_mesaj="Negatif olamaz.",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "whole", "0", "B31",
              baslik="Belge eksik", mesaj="0 ve üzeri.",
              hata_baslik="Geçersiz", hata_mesaj="Negatif olamaz.",
              isaret="greaterThanOrEqual")

    giris_hucreleri(ws, 6, 15, [2])
    giris_hucreleri(ws, 24, 33, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 40, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1024, "Sarı alanlar giriş; iade/liste formülleri kilitlidir.")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "GİB iade kural tablosu (yıl yan yana)", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir. "
      "Eşikler AYARLAR'da (kil_iadeEsik, kil_belgeLimit, kil_tolerans).",
      yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["KIL-001", "KDVK genel", "kdv_genel_oran", "her hesap", "kdv_genel_oran",
         0.20, 0.20, 0.20, "01.01.2025", "Genel KDV oranı"],
        ["KIL-002", "KDVK / iade", "iade_esik", "iade_secili>0", "iade_esik",
         50000, 50000, 50000, "01.01.2025", "İade eşiği (bilgi; operasyon AYARLAR)"],
        ["KIL-003", "Uygulama", "sapma_tolerans", "vaka kontrol", "sapma_tolerans",
         0.01, 0.01, 0.01, "01.01.2025", "Vaka tutarlılık toleransı"],
        ["KIL-004", "GİB belge", "belge_limit", "belge_eksik kontrol", "belge_limit",
         0, 0, 0, "01.01.2025", "İzin verilen eksik belge üst sınırı"],
        ["KIL-005", "KDVK genel", "yuvarlama", "her hesap", "yuvarlama",
         2, 2, 2, "01.01.2025", "Yuvarlama basamağı"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = None
            if k in (6, 7, 8):
                if isinstance(v, float) and v < 1:
                    sayi = YÜZDE
                elif isinstance(v, (int, float)):
                    sayi = CATI if s[2] in ("yuvarlama", "belge_limit") else TL
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-kil_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "KDV oranı (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("KIL-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Yuvarlama (aktif yıl)", yazi="333333")
    h(ws, 16, 2, kural_cek("KIL-005"), sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"KIL-001=",TEXT(B15,"0.0%"),"KIL-005=",TEXT(B16,"0"),'
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
    h(ws, 4, 1, "KDV oranı tblKurallar'dan; eşikler AYARLAR'dan. Sabit oran yoktur.",
      yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")

    basliklar = [*MOTOR_BASLIK, "deger"]
    baslik_satiri(ws, 6, basliklar)

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif mevzuat yılı", "KIL-001")
    ekle("KDV genel oranı", km("KIL-001"), "oran", "Genel KDV oranı", "KIL-001")
    ekle("Yuvarlama ondalık", f"=IFERROR(kil_yuvarMax+N(GIRDI!B9)*0,0)", "adet", "Yuvarlama basamağı", "KIL-005")
    ekle("İade eşiği", f"=IFERROR(kil_iadeEsik+N(GIRDI!B9)*0,0)", "TL", "İADE eşiği", "KIL-002")
    ekle("Belge limiti", f"=IFERROR(kil_belgeLimit+N(GIRDI!B9)*0,0)", "adet", "Eksik belge üst sınırı", "KIL-004")
    ekle("Tolerans", f"=IFERROR(kil_tolerans+N(GIRDI!B9)*0,0)", "TL", "Vaka toleransı", "KIL-003")
    ekle("İndirilecek KDV", "=GIRDI!B9", "TL", "LISTE_1 toplam", "KIL-001")
    ekle("Yüklenilen KDV", "=GIRDI!B10", "TL", "LISTE_2 toplam", "KIL-001")
    ekle("GÇB tutarı", "=GIRDI!B11", "TL", "LISTE_3 hasılat", "KIL-001")
    ekle("İhraç hasılatı", "=GIRDI!B12", "TL", "İhraç hasılat KDV hariç", "KIL-001")
    ekle("Belge eksik adet", "=GIRDI!B13", "adet", "Eksik belge sayısı", "KIL-004")
    ekle("İade talep", "=GIRDI!B14", "TL", "Talep edilen iade", "KIL-002")
    ekle("Liste yorumu", '=IF(GIRDI!B15="B","B","A")', "metin", "Yorum A/B", "KIL-001")
    ekle("Net yüklenen", "=MAX(0,G14-G13)", "TL", "Yüklenen−indirilecek", "KIL-001")
    ekle("İhraç KDV tavanı", f"=ROUND(G16*G8,{_YV})", "TL", "Hasılat×KDV oranı", "KIL-001")
    ekle("GÇB KDV tavanı", f"=ROUND(G15*G8,{_YV})", "TL", "GÇB×KDV oranı", "KIL-001")
    ekle("İade A (hasılat)", "=MIN(G20,G21)", "TL", "Yorum A: hasılat tavanlı", "KIL-001")
    ekle("İade B (GÇB)", "=MIN(G20,G22)", "TL", "Yorum B: GÇB tavanlı", "KIL-001")
    ekle("Seçili iade", '=IF(G19="B",G24,G23)', "TL", "Yoruma göre iade", "KIL-001")
    ekle("İade sapması", "=ABS(G25-G18)", "TL", "|Seçili−talep|", "KIL-003")
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(G17>G11,"EKSIK_BELGE",'
         'IF(G25>=G10,"IADE","BEKLE")))',
         "kod", "Karar kodu", "KIL-002")
    ekle("Karar metni",
         '=IF(G27="VERI_YOK","VERİ YOK",'
         'IF(G27="EKSIK_BELGE","EKSİK BELGE",'
         'IF(G27="IADE","İADE","BEKLE")))',
         "metin", "Kullanıcıya karar", "KIL-002")
    ekle("Gerekçe",
         '=IF(G27="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G27="EKSIK_BELGE","Eksik belge adedi limitin üzerinde; tamamlayıcı belge bekleniyor.",'
         'IF(G27="IADE","Seçili iade tutarı eşik üzerinde; GİB 7 liste dosyası hazırlanabilir.",'
         '"Seçili iade eşik altında veya sapma belirsiz — bekleyin.")))',
         "metin", "Gerekçe cümlesi", "KIL-002")
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B15)=0,kil_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B15)/kil_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "KIL-005")
    ekle("Risk skoru",
         '=IF(G27="VERI_YOK",0,IF(OR(G27="BEKLE",G27="EKSIK_BELGE"),kil_riskYuksek,'
         'IF(G27="IADE",kil_riskOrta,kil_riskDusuk)))',
         "puan", "Karar riski", "KIL-004")
    ekle("Senaryo iyimser çarpan", "=IFERROR(kil_senaryoIyi+N(GIRDI!B9)*0,0)", "oran", "İyimser çarpan", "KIL-001")
    ekle("Senaryo kötümser çarpan", "=IFERROR(kil_senaryoKotu+N(GIRDI!B9)*0,0)", "oran", "Kötümser çarpan", "KIL-001")
    ekle("İyimser iade", f"=ROUND(G25*G32,{_YV})", "TL", "İyimser seçili iade", "KIL-001")
    ekle("Kötümser iade", f"=ROUND(G25*G33,{_YV})", "TL", "Kötümser seçili iade", "KIL-001")
    ekle("M-SEN oran (+1 puan)", f"=ROUND(G8+kil_oranArtisPuan,{_YV})", "oran", "M-SEN KDV oranı", "KIL-001")
    ekle("M-SEN iade",
         f"=ROUND(MIN(G20,IF(G19=\"B\",G15*G36,G16*G36)),{_YV})",
         "TL", "Kural değişimi iade", "KIL-001")
    ekle("Tahmin üst iade", "=G25*kil_tahminUst", "TL", "Tahmin üst bant", "KIL-002")
    ekle("Tahmin alt iade", "=G25*kil_tahminAlt", "TL", "Tahmin alt bant", "KIL-002")
    ekle("Tornado: yüklenen etkisi",
         f"=ABS(ROUND(MAX(0,G14*kil_tornadoYuklenen-G13)-G20,{_YV}))",
         "TL", "Yüklenen ± etki", "KIL-001")
    ekle("Tornado: ihraç etkisi",
         f"=ABS(ROUND(MIN(G20,G16*kil_tornadoIhrac*G8)-G23,{_YV}))",
         "TL", "İhraç ± etki", "KIL-001")
    ekle("Tornado: GÇB etkisi",
         f"=ABS(ROUND(MIN(G20,G15*kil_tornadoGcb*G8)-G24,{_YV}))",
         "TL", "GÇB ± etki", "KIL-001")
    ekle("LISTE_1 özet", "=G13", "TL", "İndirilecek KDV Listesi", "KIL-001")
    ekle("LISTE_2 özet", "=G14", "TL", "Yüklenilen KDV Listesi", "KIL-001")
    ekle("LISTE_3 özet", "=G15", "TL", "GÇB Listesi", "KIL-001")
    ekle("LISTE_4 özet", "=G16", "TL", "Satış Faturaları (hasılat)", "KIL-001")
    ekle("LISTE_5 özet", "=G13", "TL", "Alış Faturaları (indirilecek ayna)", "KIL-001")
    ekle("LISTE_6 özet", "=G25", "TL", "İade Hesap Cetveli", "KIL-002")
    ekle("LISTE_7 özet", "=G17", "adet", "Ek Belge / Tamamlayıcı", "KIL-004")
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"KIL-001→KDV oranı","KIL-002→iade eşiği",'
         '"KIL-003→tolerans","KIL-004→belge","KIL-005→yuvarlama")',
         "metin", "Kanıt atıfları", "KIL-001")
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"İade=",TEXT(G25,"₺ #,##0"),"Talep=",TEXT(G18,"₺ #,##0"),'
         '"Sapma=",TEXT(G26,"₺ #,##0"),"Karar=",G28)',
         "metin", "Rapor özeti", "KIL-002")
    ekle("Kalite skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B15)=0,kil_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B15)/kil_girisBeklenen*100),0),0)",
         "puan", "Giriş kalitesi", "KIL-005")
    ekle("Yedi liste özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,'
         '"LISTE_1 İndirilecek=",TEXT(G43,"₺ #,##0"),'
         '"LISTE_2 Yüklenilen=",TEXT(G44,"₺ #,##0"),'
         '"LISTE_3 GÇB=",TEXT(G45,"₺ #,##0"),'
         '"LISTE_4 Satış=",TEXT(G46,"₺ #,##0"),'
         '"LISTE_5 Alış=",TEXT(G47,"₺ #,##0"),'
         '"LISTE_6 İade Cetveli=",TEXT(G48,"₺ #,##0"),'
         '"LISTE_7 Ek Belge=",TEXT(G49,"0"))',
         "metin", "kil_yediListe gövdesi", "KIL-001")
    ekle("Liste sayfaları adları",
         '=_xlfn.TEXTJOIN(" | ",TRUE,'
         '"LISTE_1: İndirilecek KDV Listesi",'
         '"LISTE_2: Yüklenilen KDV Listesi",'
         '"LISTE_3: GÇB Listesi",'
         '"LISTE_4: Satış Faturaları Listesi",'
         '"LISTE_5: Alış Faturaları Listesi",'
         '"LISTE_6: İade Hesap Cetveli",'
         '"LISTE_7: Ek Belge / Tamamlayıcı Liste")',
         "metin", "kil_listeSayfalari gövdesi", "KIL-001")
    ekle("Liste dağılım oranı",
         "=IF(G14=0,0,G13/G14)",
         "oran", "İndirilecek/yüklenen", "KIL-001")
    ekle("Net / yüklenen oranı",
         "=IF(G14=0,0,G20/G14)",
         "oran", "Net yüklenen payı", "KIL-001")
    ekle("İade / talep oranı",
         "=IF(G18=0,0,G25/G18)",
         "oran", "Karşılama oranı", "KIL-002")
    ekle("Belge durum bayrağı",
         '=IF(G17>G11,1,0)',
         "bayrak", "Eksik belge bayrağı", "KIL-004")

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


def _liste_sayfasi(ws, kod, turkce_ad, renk, demo_satirlar, ozet_ad, ozet_formul,
                  grafik=False, grafik_baslik=""):
    sayfa_hazirla(ws, kod, renk, URUN_AD, son_kolon=30)
    h(ws, 3, 1, f"{kod}: {turkce_ad}", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Sarı satırlar manuel satır girişi; canlı özet MOTOR'dan gelir.",
      yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:H3")
    ws.merge_cells("A4:H4")

    sutunlar = ["Satır Numarası", "Belge/Kalem", "Tutar", "KDV", "Kaynak", "Durum", "Not", "Kontrol"]
    baslik_satiri(ws, 6, sutunlar)

    for i, row in enumerate(demo_satirlar, 7):
        for kolon, v in enumerate(row, 1):
            sayi = TL if kolon in (3, 4) and isinstance(v, (int, float)) else (CATI if kolon == 1 else None)
            h(ws, i, kolon, v, sayi=sayi, zemin=GIRIS_SARI)
            ws.cell(i, kolon).protection = Protection(locked=False)
            _kim(ws, f"{get_column_letter(kolon)}{i}", sutunlar[kolon - 1], f"{turkce_ad} satırı")

    bas = 7 + len(demo_satirlar)
    for i in range(bas, bas + 18):
        for kolon in range(1, 9):
            ws.cell(i, kolon).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
            ws.cell(i, kolon).protection = Protection(locked=False)
            if ws.cell(i, kolon).comment is None:
                _kim(ws, f"{get_column_letter(kolon)}{i}", sutunlar[kolon - 1], f"{turkce_ad} satırı")

    h(ws, 7, 10, ozet_ad, kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 11, ozet_formul, sayi=TL if "adet" not in ozet_ad.lower() else CATI,
      yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 8, 10, f"{kod} özet metni", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 8, 11,
      f'=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"{kod}","{turkce_ad}",'
      f'"Özet=",TEXT(K7,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    if grafik:
        h(ws, 6, 13, "Demo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
        for i, row in enumerate(demo_satirlar[:5], 7):
            h(ws, i, 13, row[2] if isinstance(row[2], (int, float)) else 0, sayi=TL, yazi=GRİ)
        g = BarChart()
        g.type = "col"
        g.title = grafik_baslik or turkce_ad
        son = 6 + min(5, len(demo_satirlar))
        g.add_data(Reference(ws, min_col=13, min_row=6, max_row=son), titles_from_data=True)
        g.set_categories(Reference(ws, min_col=2, min_row=7, max_row=son))
        g.height, g.width = 8, 12
        ws.add_chart(g, "A28")

    genislik(ws, {"A": 14, "B": 28, "C": 16, "D": 14, "E": 12, "F": 14, "G": 18, "H": 12,
                  "J": 22, "K": 55})
    sabitle(ws, "A7")
    giris_hucreleri(ws, 7, bas + 17, [1, 2, 3, 4, 5, 6, 7, 8])


def liste_1(ws):
    _liste_sayfasi(
        ws, "LISTE_1", "İndirilecek KDV Listesi", "1F4E79",
        [
            (1, "Alış faturası A1", 400_000, 80_000, "GIRDI", "TAM", "Demo", "TAM"),
            (2, "Alış faturası A2", 300_000, 60_000, "GIRDI", "TAM", "Demo", "TAM"),
            (3, "Alış faturası A3", 100_000, 20_000, "GIRDI", "TAM", "Demo", "TAM"),
        ],
        "Canlı indirilecek (MOTOR)", "=IFERROR(kil_liste1,0)",
        grafik=True, grafik_baslik="İndirilecek KDV Satırları",
    )


def liste_2(ws):
    _liste_sayfasi(
        ws, "LISTE_2", "Yüklenilen KDV Listesi", "2E75B6",
        [
            (1, "Yüklenen kalem Y1", 1_000_000, 200_000, "GIRDI", "TAM", "Demo", "TAM"),
            (2, "Yüklenen kalem Y2", 700_000, 140_000, "GIRDI", "TAM", "Demo", "TAM"),
            (3, "Yüklenen kalem Y3", 300_000, 60_000, "GIRDI", "TAM", "Demo", "TAM"),
        ],
        "Canlı yüklenilen (MOTOR)", "=IFERROR(kil_liste2,0)",
        grafik=True, grafik_baslik="Yüklenilen KDV Satırları",
    )


def liste_3(ws):
    _liste_sayfasi(
        ws, "LISTE_3", "GÇB Listesi", "548235",
        [
            (1, "GÇB-2026-001", 2_000_000, 0, "GIRDI", "TAM", "Hasılat", "TAM"),
            (2, "GÇB-2026-002", 1_500_000, 0, "GIRDI", "TAM", "Hasılat", "TAM"),
            (3, "GÇB-2026-003", 1_000_000, 0, "GIRDI", "TAM", "Hasılat", "TAM"),
        ],
        "Canlı GÇB (MOTOR)", "=IFERROR(kil_liste3,0)",
        grafik=True, grafik_baslik="GÇB Hasılat Dağılımı",
    )


def liste_4(ws):
    _liste_sayfasi(
        ws, "LISTE_4", "Satış Faturaları Listesi", "C45911",
        [
            (1, "Satış F-1001", 2_500_000, 0, "GIRDI", "TAM", "İhraç", "TAM"),
            (2, "Satış F-1002", 1_500_000, 0, "GIRDI", "TAM", "İhraç", "TAM"),
            (3, "Satış F-1003", 1_000_000, 0, "GIRDI", "TAM", "İhraç", "TAM"),
        ],
        "Canlı satış/hasılat (MOTOR)", "=IFERROR(kil_liste4,0)",
        grafik=False,
    )


def liste_5(ws):
    _liste_sayfasi(
        ws, "LISTE_5", "Alış Faturaları Listesi", "7030A0",
        [
            (1, "Alış AF-2001", 400_000, 80_000, "GIRDI", "TAM", "Demo", "TAM"),
            (2, "Alış AF-2002", 250_000, 50_000, "GIRDI", "TAM", "Demo", "TAM"),
            (3, "Alış AF-2003", 150_000, 30_000, "GIRDI", "TAM", "Demo", "TAM"),
        ],
        "Canlı alış (MOTOR)", "=IFERROR(kil_liste5,0)",
        grafik=False,
    )


def liste_6(ws):
    _liste_sayfasi(
        ws, "LISTE_6", "İade Hesap Cetveli", "B08948",
        [
            (1, "Seçili iade", 1_000_000, 0, "MOTOR", "İADE", "Yorum A", "TAM"),
            (2, "İade talep", 900_000, 0, "GIRDI", "TALEP", "Kullanıcı", "TAM"),
            (3, "Sapma", 100_000, 0, "MOTOR", "SAPMA", "Mutlak", "TAM"),
            (4, "İade A", 1_000_000, 0, "MOTOR", "YORUM A", "Hasılat", "TAM"),
            (5, "İade B", 900_000, 0, "MOTOR", "YORUM B", "GÇB", "TAM"),
        ],
        "Canlı iade cetveli (MOTOR)", "=IFERROR(kil_liste6,0)",
        grafik=True, grafik_baslik="İade Hesap Cetveli",
    )


def liste_7(ws):
    _liste_sayfasi(
        ws, "LISTE_7", "Ek Belge / Tamamlayıcı Liste", "C00000",
        [
            (1, "GÇB aslı", 0, 0, "Kullanıcı", "VAR", "Belge", "TAM"),
            (2, "Navlun faturası", 0, 0, "Kullanıcı", "VAR", "Belge", "TAM"),
            (3, "Sigorta poliçesi", 0, 0, "Kullanıcı", "VAR", "Belge", "TAM"),
        ],
        "Canlı ek belge adedi (MOTOR)", "=IFERROR(kil_liste7,0)",
        grafik=False,
    )


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve duyarlılık", kalin=True, boyut=13, yazi=KOYU_LACIVERT)

    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Çarpan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Seçili İade", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    iad = "IFERROR(kil_iadeSecili,0)"
    senaryolar = [
        ("İyimser", "=IFERROR(kil_senaryoIyi+N(GIRDI!B9)*0,0)",
         f"=IFERROR(ROUND(({iad})*B6,2),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(GIRDI!B13>kil_belgeLimit,"EKSİK BELGE",'
         'IF(C6>=kil_iadeEsik,"İADE","BEKLE")))'),
        ("Baz", "=IFERROR(1+N(GIRDI!B9)*0,1)",
         f"=IFERROR(ROUND(({iad})*B7,2),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(GIRDI!B13>kil_belgeLimit,"EKSİK BELGE",'
         'IF(C7>=kil_iadeEsik,"İADE","BEKLE")))'),
        ("Kötümser", "=IFERROR(kil_senaryoKotu+N(GIRDI!B9)*0,0)",
         f"=IFERROR(ROUND(({iad})*B8,2),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(GIRDI!B13>kil_belgeLimit,"EKSİK BELGE",'
         'IF(C8>=kil_iadeEsik,"İADE","BEKLE")))'),
        ("M-SEN +1 puan", "=IFERROR(1+N(GIRDI!B9)*0,1)",
         "=IFERROR(MOTOR!G37,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(GIRDI!B13>kil_belgeLimit,"EKSİK BELGE",'
         'IF(C9>=kil_iadeEsik,"İADE","BEKLE")))'),
    ]
    for i, (ad, carp, asg, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, asg, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant genişliği (iyimser−kötümser iade)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(C6-C8,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser iade ",TEXT(C6,"₺ #,##0")," · Baz ",TEXT(C7,"₺ #,##0"),'
      '" · Kötümser ",TEXT(C8,"₺ #,##0")," · Bant ",TEXT(B11,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:H12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Tornado yüklenen etkisi", yazi="333333")
    h(ws, 16, 2, "=IFERROR(MOTOR!G40,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado ihraç etkisi", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G41,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado GÇB etkisi", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G42,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
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
    h(ws, 22, 2, "=C9", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"KDV oranı +1 puan senaryosunda iade ",'
      'TEXT(C9,"₺ #,##0")," TL")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı (iade)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=MOTOR!G39", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 3, "=kil_iadeSecili", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 4, "=MOTOR!G38", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Tahmin (FORECAST)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1, "Seri nokta", yazi=GRİ)
    for i, v in enumerate([1, 2, 3, 4], 28):
        h(ws, i, 1, v, sayi=CATI)
        h(ws, i, 2, f"=C{5+v}", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 32, 1, "Tahmin sonraki", yazi="333333")
    h(ws, 32, 2, "=IFERROR(FORECAST.LINEAR(5,B28:B31,A28:A31),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Tahmin yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Tahmin ",TEXT(B32,"₺ #,##0")," · Alt ",TEXT(B25,"₺ #,##0"),'
      '" · Üst ",TEXT(D25,"₺ #,##0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 5, 7, "IadeDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([900_000, 1_000_000, 1_100_000, 1_050_000], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([100_000, 50_000, 45_000], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)
    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([900_000, 1_000_000, 1_100_000, 1_050_000], 28):
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

    genislik(ws, {"A": 36, "B": 18, "C": 18, "D": 22, "E": 14, "G": 14})


def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "GİB iade altın vakaları", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:H3")
    baslik_satiri(ws, 5, VAKA_BASLIK)

    # V-001: net = 2M-0.8M = 1.2M
    # V-002: A = MIN(1.2M, 5M*0.2) = 1M
    # V-003: B = MIN(1.2M, 4.5M*0.2) = 0.9M
    # V-004: oran kontrol 0.20
    vakalar_data = [
        ("V-001", "Net yüklenen = yüklenen − indirilecek",
         "yuk=vaka_yuklenen; ind=vaka_indirilecek",
         1_200_000,
         "=IFERROR(ROUND(vaka_yuklenen-vaka_indirilecek+N(GIRDI!B10)*0,2),0)",
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=kil_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Net yüklenen kontrol"),
        ("V-002", "İade A — hasılat tavanlı",
         "MIN(net, ihrac×KIL-001)",
         1_000_000,
         "=IFERROR(ROUND(MIN(vaka_yuklenen-vaka_indirilecek,"
         "vaka_ihrac*INDEX(tblKurallar[deger_2026],MATCH(\"KIL-001\",tblKurallar[kural_id],0)))"
         "+N(GIRDI!B12)*0,2),0)",
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=kil_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Yorum A iade"),
        ("V-003", "İade B — GÇB tavanlı",
         "MIN(net, gcb×KIL-001)",
         900_000,
         "=IFERROR(ROUND(MIN(vaka_yuklenen-vaka_indirilecek,"
         "vaka_gcb*INDEX(tblKurallar[deger_2026],MATCH(\"KIL-001\",tblKurallar[kural_id],0)))"
         "+N(GIRDI!B11)*0,2),0)",
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=kil_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Yorum B iade"),
        ("V-004", "KDV oranı kontrol",
         "KIL-001 deger_2026",
         0.20,
         "=IFERROR(INDEX(tblKurallar[deger_2026],MATCH(\"KIL-001\",tblKurallar[kural_id],0))"
         "+N(GIRDI!B8)*0,0)",
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=kil_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Oran tutarlılığı"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v,
              yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if k in (4, 5, 6) and i < 9:
                ws.cell(i, k).number_format = TL
            if i == 9 and k in (4, 5, 6):
                ws.cell(i, k).number_format = YÜZDE

    formuller = {
        "fark": (
            "=IF(tblVakalar[[#This Row],[vaka_id]]=\"\",\"\","
            "IFERROR(tblVakalar[[#This Row],[beklenen_sonuc]]-"
            "tblVakalar[[#This Row],[hesaplanan]],0))"
        ),
        "durum": (
            '=IF(tblVakalar[[#This Row],[vaka_id]]="","",'
            'IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=kil_tolerans,"TUTARLI","KIRIK"),"KIRIK"))'
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
    for i, (b, he) in enumerate([(1_200_000, 1_200_000), (1_000_000, 1_000_000),
                                  (900_000, 900_000), (0.20, 0.20)], 6):
        h(ws, i, 10, b, sayi=(YÜZDE if i == 9 else TL), yazi=GRİ)
        h(ws, i, 11, he, sayi=(YÜZDE if i == 9 else TL), yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Vaka Beklenen vs Hesaplanan"
    g.add_data(Reference(ws, min_col=10, min_row=5, max_col=11, max_row=8), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=6, max_row=8))
    g.height, g.width = 8, 14
    ws.add_chart(g, "A15")

    genislik(ws, {"A": 12, "B": 40, "C": 36, "D": 14, "E": 14, "F": 12, "G": 12, "H": 28})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", "0F2742", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "KANIT RAPORU — KDV İade Listesi (GİB 7 Liste)", kalin=True, boyut=14,
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
        ("İndirilecek KDV", "=GIRDI!B9", TL),
        ("Yüklenilen KDV", "=GIRDI!B10", TL),
        ("GÇB tutarı", "=GIRDI!B11", TL),
        ("İhraç hasılatı", "=GIRDI!B12", TL),
        ("Belge eksik adet", "=GIRDI!B13", CATI),
        ("İade talep", "=GIRDI!B14", TL),
        ("Liste yorumu", "=GIRDI!B15", None),
    ], 9):
        h(ws, i, 1, et, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)

    h(ws, 17, 1, "Hesap zinciri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 18, 1, "Seçili iade", yazi="333333")
    h(ws, 18, 2, "=kil_iadeSecili", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Sapma", yazi="333333")
    h(ws, 19, 2, "=kil_modulO1", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Yedi liste", yazi="333333")
    h(ws, 20, 2, "=kil_yediListe", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B20:F20")
    h(ws, 21, 1, "KARAR", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 2, "=kil_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 22, 1, "Gerekçe", yazi="333333")
    h(ws, 22, 2, "=kil_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B22:F22")

    h(ws, 24, 1, "Madde atıfları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, "=kil_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B24:F24")

    h(ws, 26, 1, "Kanıt gövdesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 2, "=kil_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B26:F26")

    h(ws, 27, 1, "Liste sayfaları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 2, "=kil_listeSayfalari", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B27:F27")

    h(ws, 29, 1, "Parmak izi (SHA-256 kısaltma)", yazi="333333")
    h(ws, 29, 2, "=kil_parmakIzi", yazi=GRİ)

    h(ws, 31, 1, "Hazırlayan (imza)", yazi="333333")
    h(ws, 31, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B31", "İmza", "Raporu hazırlayan kişinin adını yazın.")
    h(ws, 32, 1, "Onaylayan (imza)", yazi="333333")
    h(ws, 32, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B32", "Onay", "Onaylayan kişinin adını yazın.")

    h(ws, 34, 1,
      "Uyarı: Bu çıktı karar destektir; kesin KDV tahakkuku veya hukuki görüş değildir. "
      f"Sürüm {SURUM}.",
      yazi=GRİ, boyut=9, kaydir=True)
    ws.merge_cells("A34:F34")

    baski_hazirla(ws, "A1:F35", f"{URUN_AD} · {SURUM}")
    ws.page_setup.orientation = "portrait"
    genislik(ws, {"A": 28, "B": 22, "C": 14, "D": 14, "E": 14, "F": 14})
    giris_hucreleri(ws, 31, 32, [2])


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim paneli — KDV İade Listesi (GİB 7 Liste)", kalin=True, boyut=14,
      yazi=KOYU_LACIVERT)
    genislik(ws, {"A": 44})

    kpiler = [
        (2, "İndirilecek", "=IFERROR(kil_modulT1,0)", TL),
        (3, "Yüklenilen", "=IFERROR(kil_modulT2,0)", TL),
        (4, "İade", "=IFERROR(kil_iadeSecili,0)", TL),
        (5, "Talep", "=IFERROR(GIRDI!B14,0)", TL),
        (6, "Sapma", "=IFERROR(kil_modulO1,0)", TL),
        (7, "Güven", "=IFERROR(MOTOR!G30,0)", CATI),
        (8, "Risk", "=IFERROR(MOTOR!G31,0)", CATI),
        (9, "GÇB", "=IFERROR(kil_liste3,0)", TL),
        (10, "İhraç", "=IFERROR(GIRDI!B12,0)", TL),
        (11, "Dağılım", "=IFERROR(kil_modulO6,0)", YÜZDE),
        (12, "M-SEN", "=IFERROR(kil_kuralDegisimSenaryo,0)", TL),
        (13, "Tahmin", "=IFERROR(SENARYO!B32,0)", TL),
        (14, "Vaka", "=IFERROR(kil_vakaDurum,\"-\")", None),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=9, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=kil_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=kil_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 8, 1, "GİB 7 liste adları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 8, 2, "=kil_listeSayfalari", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B8:N8")

    h(ws, 10, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 11, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "İndirilecek toplam",
         "=IFERROR(kil_liste1,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İndirilecek KDV ",TEXT(C12,"₺ #,##0")," TL — LISTE_1")'),
        ("T2", "Yüklenilen toplam",
         "=IFERROR(kil_liste2,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Yüklenilen KDV ",TEXT(C13,"₺ #,##0")," TL — LISTE_2")'),
        ("O1", "İade sapması",
         "=IFERROR(MOTOR!G26,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"|Seçili−talep| ",TEXT(C14,"₺ #,##0")," TL")'),
        ("O2", "Tornado max etki",
         "=IFERROR(kil_duyarlilikSira,\"-\")",
         "=IFERROR(kil_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo bant",
         "=IFERROR(kil_senaryoKarsilastirma,0)",
         "=IFERROR(kil_yorumSenaryo,\"-\")"),
        ("O6", "Liste dağılımı",
         "=IFERROR(MOTOR!G55,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İndirilecek/yüklenen ",TEXT(C17,"0.0%"))'),
        ("O8", "Kalite skoru",
         "=IFERROR(MOTOR!G52,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalite ",TEXT(C18,"0")," puan")'),
        ("M_SEN", "M-SEN iade",
         "=IFERROR(kil_kuralDegisimSenaryo,0)",
         "=IFERROR(kil_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin aralık",
         "=IFERROR(kil_tahminAralik,0)",
         "=IFERROR(kil_yorumTahmin,\"-\")"),
        ("I2", "Senaryo iade P90",
         "=IFERROR(IF(COUNT(SENARYO!C6:C9)<2,0,"
         "_xlfn.PERCENTILE.INC(SENARYO!C6:C9,kil_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"P90 iade ",TEXT(C21,"₺ #,##0")," TL")'),
    ]
    for i, (kod, ad, form, yorum) in enumerate(moduller):
        r = 12 + i
        h(ws, r, 1, kod, kalin=True, hiza="center", yazi="B08948")
        h(ws, r, 2, ad, kaydir=True)
        h(ws, r, 3, form, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[r].height = 28
        if kod == "O6":
            ws.cell(r, 3).number_format = YÜZDE
        elif kod != "O2":
            ws.cell(r, 3).number_format = TL

    h(ws, 24, 1, "Liste", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("LISTE_1 İndirilecek", 800_000),
        ("LISTE_2 Yüklenilen", 2_000_000),
        ("LISTE_3 GÇB", 4_500_000),
        ("LISTE_4 Satış", 5_000_000),
        ("LISTE_6 İade", 1_000_000),
    ], 25):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)

    h(ws, 24, 4, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 5, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Seçili iade", 1_000_000),
        ("İade talep", 900_000),
        ("İade A", 1_000_000),
        ("İade B", 900_000),
        ("M-SEN iade", 1_050_000),
    ], 25):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "GİB 7 Liste Özet Tutarları"
    g1.add_data(Reference(ws, min_col=2, min_row=24, max_row=29), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=25, max_row=29))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G10")

    g2 = BarChart()
    g2.type = "col"
    g2.title = "İade / Talep / Yorum A-B"
    g2.add_data(Reference(ws, min_col=5, min_row=24, max_row=29), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=25, max_row=29))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "G24")

    g3 = LineChart()
    g3.title = "İade Karşılaştırma Çizgisi"
    g3.add_data(Reference(ws, min_col=5, min_row=24, max_row=29), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=4, min_row=25, max_row=29))
    g3.height, g3.width = 7, 12
    ws.add_chart(g3, "P10")

    baski_hazirla(ws, "A1:F34", f"{URUN_AD} · PANO")
    genislik(ws, {"A": 28, "B": 16, "C": 18, "D": 55})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Canlı kontrol paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    kontroller_list = [
        ("Giriş tablosu boş mu?",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"BOŞ","DOLU")'),
        ("Negatif KDV tutarı var mı?",
         '=IF(OR(GIRDI!B9<0,GIRDI!B10<0,GIRDI!B11<0,GIRDI!B12<0),"NEGATİF","TEMİZ")'),
        ("Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(kil_yilTaban),N(hesapYili)<=N(kil_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        ("İade seçili sıfır mı?",
         '=IF(IFERROR(kil_iadeSecili,0)=0,"SIFIR","NORMAL")'),
        ("Vaka kırık mı?",
         '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")'),
        ("İade eşiği aşıyor mu?",
         '=IF(IFERROR(kil_iadeSecili,0)>=kil_iadeEsik,"AŞIYOR","NORMAL")'),
        ("Belge eksik mi?",
         '=IF(IFERROR(GIRDI!B13,0)>kil_belgeLimit,"EKSİK","TEMİZ")'),
        ("Yorum A/B seçili mi?",
         '=IF(OR(GIRDI!B15="A",GIRDI!B15="B"),"TEMİZ","EKSİK")'),
        ("Motor adım sayısı",
         "=COUNTA(MOTOR!B7:B70)"),
        ("Yedi liste metni dolu mu?",
         '=IF(IFERROR(kil_yediListe,"")="","EKSİK","TAM")'),
    ]
    for i, (ad, form) in enumerate(kontroller_list, 5):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Canlı formül | Değiştirmeyin")

    h(ws, 16, 1, "Toplam giriş alanı", yazi="333333")
    h(ws, 16, 2, "=IFERROR(kil_girisBeklenen+COUNTA(GIRDI!B6:B15)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Dolu giriş", yazi="333333")
    h(ws, 17, 2, "=COUNTA(GIRDI!B6:B15)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Kalite skoru (modül)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 2, "=IFERROR(ROUND(IF(OR(B16=0,B16=\"\"),0,B17/B16*100),0),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)

    h(ws, 19, 1, "İade seçili kontrol", yazi="333333")
    h(ws, 19, 2, "=IFERROR(kil_iadeSecili,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Karar kontrol", yazi="333333")
    h(ws, 20, 2, "=IFERROR(kil_kararMetni,\"-\")", yazi=DIS_REF_YEŞIL)
    h(ws, 21, 1, "Liste sayfaları kontrol", yazi="333333")
    h(ws, 21, 2, "=IFERROR(kil_listeSayfalari,\"-\")", yazi=DIS_REF_YEŞIL)

    genislik(ws, {"A": 40, "B": 55})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek senaryo kütüphanesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:L3")
    h(ws, 4, 1, "Bu sayfa demo değerleri saklar; GIRDI'ye kopyalanabilir.", yazi=GRİ)
    sutunlar = ["Senaryo", "İndirilecek", "Yüklenilen", "GÇB", "İhraç",
                "Belge Eksik", "İade Talep", "Yorum", "İade Seçili", "Kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Kontrol": (
            '=IF(tblOrnek[[#This Row],[Senaryo]]="","",'
            'IF(OR(tblOrnek[[#This Row],[İndirilecek]]="",'
            'tblOrnek[[#This Row],[Yüklenilen]]=""),"EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:J25", sutunlar, formuller)
    satirlar = [
        ("Demo ana", 800_000, 2_000_000, 4_500_000, 5_000_000, 0, 900_000, "A", 1_000_000),
        ("Yorum B", 800_000, 2_000_000, 4_500_000, 5_000_000, 0, 900_000, "B", 900_000),
        ("Eksik belge", 800_000, 2_000_000, 4_500_000, 5_000_000, 2, 900_000, "A", 1_000_000),
        ("Düşük iade", 100_000, 120_000, 200_000, 250_000, 0, 10_000, "A", 20_000),
        ("Yüksek hasılat", 500_000, 3_000_000, 8_000_000, 9_000_000, 0, 1_500_000, "A", 1_800_000),
    ]
    for i, row in enumerate(satirlar, 6):
        for k, v in enumerate(row, 1):
            sayi = TL if k in (2, 3, 4, 5, 7, 9) else (CATI if k == 6 else None)
            h(ws, i, k, v, sayi=sayi, zemin=GIRIS_SARI, yazi="1F4E79")
            ws.cell(i, k).protection = Protection(locked=False)
            _kim(ws, f"{get_column_letter(k)}{i}", sutunlar[k - 1], "Örnek senaryo satırı")
    for i in range(11, 26):
        for k in range(1, 10):
            ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
            ws.cell(i, k).protection = Protection(locked=False)
            if ws.cell(i, k).comment is None:
                _kim(ws, f"{get_column_letter(k)}{i}", sutunlar[k - 1], "Örnek satır")
    giris_hucreleri(ws, 6, 25, [1, 2, 3, 4, 5, 6, 7, 8, 9])
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 11)})
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
        _kim(ws, f"C{i}", "Yorum", "Liste sırası yorumu")
    h(ws, 5, 5, "Evet/Hayır", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate(["Evet", "Hayır"], 6):
        h(ws, i, 5, v, zemin=GIRIS_SARI)
        ws.cell(i, 5).protection = Protection(locked=False)
        _kim(ws, f"E{i}", "Seçim", "Evet veya Hayır")
    h(ws, 10, 1, "GİB 7 liste adları (referans)", kalin=True, yazi=KOYU_LACIVERT)
    for i, (kod, ad) in enumerate(LISTE_ADLARI, 11):
        h(ws, i, 1, kod, yazi="333333")
        h(ws, i, 2, ad, yazi="333333")
    genislik(ws, {"A": 28, "B": 40, "C": 14, "E": 14})


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
        ("kil_tolerans", 0.01, "TL", "Vaka tutarlılık toleransı", "Uygulama notu",
         "01.01.2025", "14.08.2026"),
        ("kil_iadeEsik", 50_000, "TL", "İADE karar eşiği", "İç politika",
         "01.01.2025", "14.08.2026"),
        ("kil_belgeLimit", 0, "adet", "İzin verilen eksik belge üst sınırı", "GİB belge",
         "01.01.2025", "14.08.2026"),
        ("kil_senaryoIyi", 0.90, "çarpan", "İyimser iade çarpanı", "Model varsayımı",
         "01.01.2025", "14.08.2026"),
        ("kil_senaryoKotu", 1.10, "çarpan", "Kötümser iade çarpanı", "Model varsayımı",
         "01.01.2025", "14.08.2026"),
        ("kil_oranArtisPuan", 0.01, "oran", "M-SEN +1 puan", "Senaryo motoru",
         "01.01.2025", "14.08.2026"),
        ("kil_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı",
         "01.01.2025", "14.08.2026"),
        ("kil_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı",
         "01.01.2025", "14.08.2026"),
        ("kil_tornadoYuklenen", 1.05, "çarpan", "Tornado yüklenen şoku", "Model varsayımı",
         "01.01.2025", "14.08.2026"),
        ("kil_tornadoIhrac", 1.05, "çarpan", "Tornado ihraç şoku", "Model varsayımı",
         "01.01.2025", "14.08.2026"),
        ("kil_tornadoGcb", 0.95, "çarpan", "Tornado GÇB şoku", "Model varsayımı",
         "01.01.2025", "14.08.2026"),
        ("kil_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik",
         "01.01.2025", "14.08.2026"),
        ("kil_parmakIzi", "KIL-PRO-1.0.0", "metin", "Dosya parmak izi etiketi", "Üretim",
         "01.01.2025", "14.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv",
         "01.01.2025", "14.08.2026"),
        ("kil_girisBeklenen", 10, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi",
         "01.01.2025", "14.08.2026"),
        ("kil_riskDusuk", 15, "puan", "Düşük risk skoru", "İç politika",
         "01.01.2025", "14.08.2026"),
        ("kil_riskOrta", 55, "puan", "Orta risk skoru", "İç politika",
         "01.01.2025", "14.08.2026"),
        ("kil_riskYuksek", 85, "puan", "Yüksek risk skoru", "İç politika",
         "01.01.2025", "14.08.2026"),
        ("kil_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı (hesapYili−taban)", "Ürün mimarisi",
         "01.01.2025", "14.08.2026"),
        ("kil_yuvarMax", 2, "adet", "ROUND basamak tavanı", "Ürün mimarisi",
         "01.01.2025", "14.08.2026"),
        ("vaka_indirilecek", 800_000, "TL", "Vaka indirilecek KDV", "Demo vaka",
         "01.01.2025", "14.08.2026"),
        ("vaka_yuklenen", 2_000_000, "TL", "Vaka yüklenen KDV", "Demo vaka",
         "01.01.2025", "14.08.2026"),
        ("vaka_ihrac", 5_000_000, "TL", "Vaka ihraç hasılat", "Demo vaka",
         "01.01.2025", "14.08.2026"),
        ("vaka_gcb", 4_500_000, "TL", "Vaka GÇB tutarı", "Demo vaka",
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
    dogrulama(ws, "whole", "0", "B10",
              baslik="Belge limiti", mesaj="0 ve üzeri.",
              hata_baslik="Geçersiz", hata_mesaj="Negatif olamaz.",
              isaret="greaterThanOrEqual")

    genislik(ws, {"A": 28, "B": 24, "C": 10, "D": 36, "E": 22, "F": 14, "G": 14, "H": 12})
    h(ws, son_satir + 4, 2,
      f"Sürüm {SURUM} | Şifre koruması: {SIFRE} (formül alanları)",
      yazi=GRİ, boyut=9, kaydir=True)


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar kuralları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "VERİ YOK — giriş tablosu boşsa hesap yapılmaz.",
        "İADE — seçili iade tutarı eşik üzerinde ve belge limiti aşılmamış.",
        "EKSİK BELGE — belge eksik adedi kil_belgeLimit üstünde.",
        "BEKLE — seçili iade eşik altında veya sapma belirsiz.",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 10, 1, "GİB 7 liste adları", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, (kod, ad) in enumerate(LISTE_ADLARI, 11):
        h(ws, i, 1, f"{kod}: {ad}", yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 19, 1, "Belirsizlik beyanları (M05)", kalin=True, boyut=12, yazi=KRITIK)
    ws.merge_cells("A19:J19")
    for i, m in enumerate(belirsizlik_beyanlari(["liste_sirasi"]), 20):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 22, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2,
      "İade tavanı ihraç hasılatı × KDV oranı ile sınırlanır (LISTE_4 / hasılat). "
      "Bu nokta mevzuatta tartışmalıdır; sonuç yorum B ile yan yana okunmalıdır.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B22:J22")
    h(ws, 23, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      "İade tavanı GÇB tutarı × KDV oranı ile sınırlanır (LISTE_3). "
      "Dosya her iki yorumun sonucunu MOTOR'da üretir.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B23:J23")

    h(ws, 25, 1, "Kullanım sırası", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 1,
      "GIRDI → KURALLAR → MOTOR → LISTE_1..LISTE_7 → SENARYO → VAKALAR → "
      "PANO → KANIT_RAPORU → AYARLAR → KILAVUZ.",
      kaydir=True, yazi="333333")
    ws.merge_cells("A26:J26")
    genislik(ws, {"A": 36, "B": 70})
    h(ws, 28, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)


def kosullu_bicimlendirme(wb):
    yesil = Font(color=NORMAL, bold=True)
    kirmizi = Font(color=KRITIK, bold=True)
    amber = Font(color="B7791F", bold=True)
    y_fill = PatternFill("solid", fgColor="E2EFDA")
    k_fill = PatternFill("solid", fgColor="FDE9E9")
    a_fill = PatternFill("solid", fgColor="FFF2CC")

    ws = wb["PANO"]
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("İADE",B6))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("EKSİK",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("BEKLE",B6))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("VERİ YOK",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("D4", CellIsRule(operator="greaterThan", formula=["0"], font=amber))
    ws.conditional_formatting.add("G4", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("G4", CellIsRule(operator="between", formula=["50", "79"], font=amber))
    ws.conditional_formatting.add("G4", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))
    ws.conditional_formatting.add("H4", CellIsRule(operator="greaterThanOrEqual", formula=["70"], font=kirmizi))
    ws.conditional_formatting.add("H4", CellIsRule(operator="between", formula=["40", "69"], font=amber))
    ws.conditional_formatting.add("H4", CellIsRule(operator="lessThan", formula=["40"], font=yesil))
    for col in range(2, 15):
        harf = get_column_letter(col)
        ws.conditional_formatting.add(f"{harf}4", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))

    ws = wb["GIRDI"]
    for col in ("B",):
        ws.conditional_formatting.add(
            f"{col}9:{col}14", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(
            f"{col}9:{col}12", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add(
        "F24:F33", FormulaRule(formula=['F24="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add(
        "F24:F33", FormulaRule(formula=['F24="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add(
        "G6:G9", FormulaRule(formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add(
        "G6:G9", FormulaRule(formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add(
        "F6:F9", CellIsRule(operator="notEqual", formula=["0"], font=amber))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add(
        "C6:C9", CellIsRule(operator="greaterThan", formula=["0"], font=amber))
    ws.conditional_formatting.add(
        "D6:D9", FormulaRule(formula=['ISNUMBER(SEARCH("İADE",D6))'], font=yesil))
    ws.conditional_formatting.add(
        "D6:D9", FormulaRule(formula=['ISNUMBER(SEARCH("EKSİK",D6))'], font=kirmizi))
    ws.conditional_formatting.add(
        "B16:B18", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add(
        "G7:G70", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add(
        "F7:F70", FormulaRule(formula=['F7="KIL-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add(
        "F7:F70", FormulaRule(formula=['F7="KIL-004"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add(
        "B21", FormulaRule(formula=['ISNUMBER(SEARCH("İADE",B21))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add(
        "B21", FormulaRule(formula=['ISNUMBER(SEARCH("EKSİK",B21))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add(
        "B21", FormulaRule(formula=['ISNUMBER(SEARCH("VERİ YOK",B21))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add(
        "B18", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KURALLAR"]
    for col in ("F", "G", "H"):
        ws.conditional_formatting.add(
            f"{col}7:{col}11", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(
            f"{col}7:{col}11", CellIsRule(operator="equal", formula=["0"], font=amber))

    for kod in ("LISTE_1", "LISTE_2", "LISTE_3", "LISTE_4", "LISTE_5", "LISTE_6", "LISTE_7"):
        ws = wb[kod]
        ws.conditional_formatting.add(
            "C7:C20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(
            "H7:H20", FormulaRule(formula=['H7="EKSİK"'], font=kirmizi, fill=k_fill))
        ws.conditional_formatting.add(
            "H7:H20", FormulaRule(formula=['H7="TAM"'], font=yesil, fill=y_fill))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add(
        "B5:B14", FormulaRule(
            formula=['OR(B5="KIRIK",B5="HATALI",B5="NEGATİF",B5="BOŞ",B5="EKSİK",B5="AŞIYOR",B5="SIFIR")'],
            font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add(
        "B5:B14", FormulaRule(
            formula=['OR(B5="TEMİZ",B5="DOLU",B5="NORMAL",B5="TUTARLI",B5="TAM")'],
            font=yesil, fill=y_fill))
    ws.conditional_formatting.add(
        "B15", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add(
        "B15", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "C", "D", "E", "F", "G", "I"):
        ws.conditional_formatting.add(
            f"{harf}6:{harf}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add(
        "J6:J20", FormulaRule(formula=['J6="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add(
        "J6:J20", FormulaRule(formula=['J6="TAM"'], font=yesil, fill=y_fill))


def ad_tanimlari(wb):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")

    ad_ekle(wb, "kil_iadeSecili", "MOTOR!$G$25")
    ad_ekle(wb, "kil_kararMetni", "MOTOR!$G$28")
    ad_ekle(wb, "kil_kararGerekce", "MOTOR!$G$29")
    ad_ekle(wb, "kil_kararKapi", "MOTOR!$G$28")
    ad_ekle(wb, "kil_maddeAtifMetni", "MOTOR!$G$50")
    ad_ekle(wb, "kil_kanitRaporu", "MOTOR!$G$51")
    ad_ekle(wb, "kil_yediListe", "MOTOR!$G$53")
    ad_ekle(wb, "kil_listeSayfalari", "MOTOR!$G$54")
    ad_ekle(wb, "kil_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "kil_parmakIzi", "AYARLAR!$B$20")

    ad_ekle(wb, "kil_liste1", "MOTOR!$G$43")
    ad_ekle(wb, "kil_liste2", "MOTOR!$G$44")
    ad_ekle(wb, "kil_liste3", "MOTOR!$G$45")
    ad_ekle(wb, "kil_liste4", "MOTOR!$G$46")
    ad_ekle(wb, "kil_liste5", "MOTOR!$G$47")
    ad_ekle(wb, "kil_liste6", "MOTOR!$G$48")
    ad_ekle(wb, "kil_liste7", "MOTOR!$G$49")

    ad_ekle(wb, "kil_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "kil_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "kil_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "kil_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "kil_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "kil_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "kil_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "kil_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "kil_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "kil_vakaFark", "VAKALAR!$B$13")

    # AYARLAR satır haritası (params sırası, B6=hesapYili ...):
    # 6 hesapYili, 7 raporTarihi, 8 kil_tolerans, 9 kil_iadeEsik, 10 kil_belgeLimit,
    # 11 kil_senaryoIyi, 12 kil_senaryoKotu, 13 kil_oranArtisPuan, 14 kil_tahminAlt,
    # 15 kil_tahminUst, 16 kil_tornadoYuklenen, 17 kil_tornadoIhrac, 18 kil_tornadoGcb,
    # 19 kil_yuzdelikOran, 20 kil_parmakIzi, 21 dosya_surumu, 22 kil_girisBeklenen,
    # 23 kil_riskDusuk, 24 kil_riskOrta, 25 kil_riskYuksek, 26 kil_yilTaban, 27 kil_yuvarMax,
    # 28 vaka_indirilecek, 29 vaka_yuklenen, 30 vaka_ihrac, 31 vaka_gcb
    ayar_map = {
        "kil_tolerans": 8,
        "kil_iadeEsik": 9,
        "kil_belgeLimit": 10,
        "kil_senaryoIyi": 11,
        "kil_senaryoKotu": 12,
        "kil_oranArtisPuan": 13,
        "kil_tahminAlt": 14,
        "kil_tahminUst": 15,
        "kil_tornadoYuklenen": 16,
        "kil_tornadoIhrac": 17,
        "kil_tornadoGcb": 18,
        "kil_yuzdelikOran": 19,
        "kil_girisBeklenen": 22,
        "kil_riskDusuk": 23,
        "kil_riskOrta": 24,
        "kil_riskYuksek": 25,
        "kil_yilTaban": 26,
        "kil_yuvarMax": 27,
        "vaka_indirilecek": 28,
        "vaka_yuklenen": 29,
        "vaka_ihrac": 30,
        "vaka_gcb": 31,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    # PANO modül satırları: başlık 11, veri 12..21
    modul_map = {
        "kil_modulT1": (12, "C"), "kil_modulT1Yorum": (12, "D"),
        "kil_modulT2": (13, "C"), "kil_modulT2Yorum": (13, "D"),
        "kil_modulO1": (14, "C"), "kil_modulO1Yorum": (14, "D"),
        "kil_modulO6": (17, "C"), "kil_modulO6Yorum": (17, "D"),
        "kil_modulO8": (18, "C"), "kil_modulO8Yorum": (18, "D"),
        "kil_modulI2": (21, "C"), "kil_modulI2Yorum": (21, "D"),
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
        ("LISTE_1", liste_1),
        ("LISTE_2", liste_2),
        ("LISTE_3", liste_3),
        ("LISTE_4", liste_4),
        ("LISTE_5", liste_5),
        ("LISTE_6", liste_6),
        ("LISTE_7", liste_7),
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
    varsayilan = os.path.join(kok, "KdvIadeListesiRobotuGib7.xlsx")
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
    print(f"Sayfalar: {wb2.sheetnames}")
    return hedef


if __name__ == "__main__":
    yol = sys.argv[1] if len(sys.argv) > 1 else None
    uretilen = main(yol)
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
    cikti = os.path.join(repo, "cikti", "KdvIadeListesiRobotuGib7.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
