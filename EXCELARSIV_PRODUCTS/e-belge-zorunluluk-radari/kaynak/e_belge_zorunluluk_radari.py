#!/usr/bin/env python3
"""e-Belge Zorunluluk Radarı — A1+A4 üretim betiği (manda v6)."""

from __future__ import annotations

import hashlib
import os
import shutil
import sys
from datetime import date, timedelta

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
    FONT,
    GIRIS_SARI,
    GIRIS_YAZI,
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
from ortak.akis_motoru import (  # noqa: E402
    durum_haritasi_dogrula,
    gecersiz_gecis_formul,
    kuyruk_sayim,
    kuyruk_tutar,
    yas_kova_formul,
    yaslandirma_gun,
    yetim_kayit_formul,
    zincir_kirik_formul,
)
from ortak.kanit_uretici import imza_alani, rapor_baslik_satirlari  # noqa: E402
from ortak.mevzuat_motoru import (  # noqa: E402
    KURAL_BASLIK,
    MOTOR_BASLIK,
    VAKA_BASLIK,
    belirsizlik_beyanlari,
)

URUN_AD = "e-Belge Zorunluluk Radarı"
SURUM = "1.0.0"
RENK = "0F2742"
KAPASITE = 200
DEMO_ENV = 8
DEMO_EYLEM = 16
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

DEMO = {
    "firma": "Örnek Ticaret A.Ş.",
    "donem": "2025",
    "hesapYili": 2025,
    "brut_satis": 4_500_000,
    "ciro_donem_yorumu": "A",
    "sektor": "Ticaret",
}

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "PLANLANDI", "durum_adi": "Planlandı", "sira": 1,
     "izinli_sonraki_durumlar": "DEVAM|IPTAL", "kategori": "acik"},
    {"durum_id": "DEVAM", "durum_adi": "Devam", "sira": 2,
     "izinli_sonraki_durumlar": "BEKLIYOR|TAMAMLANDI|IPTAL", "kategori": "acik"},
    {"durum_id": "BEKLIYOR", "durum_adi": "Bekliyor", "sira": 3,
     "izinli_sonraki_durumlar": "DEVAM|TAMAMLANDI|IPTAL", "kategori": "acik"},
    {"durum_id": "TAMAMLANDI", "durum_adi": "Tamamlandı", "sira": 4,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "IPTAL", "durum_adi": "İptal", "sira": 5,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "PLANLANDI|DEVAM", "PLANLANDI|IPTAL",
    "DEVAM|BEKLIYOR", "DEVAM|TAMAMLANDI", "DEVAM|IPTAL",
    "BEKLIYOR|DEVAM", "BEKLIYOR|TAMAMLANDI", "BEKLIYOR|IPTAL",
]

# Envanter: belge türü kartları
ENVANTER_DEMO = [
    ("EBV-00001", "e-Fatura", "EBR-001", "ZORUNLU", "Hayır", 3_000_000),
    ("EBV-00002", "e-Arşiv", "EBR-002", "ZORUNLU", "Hayır", 500_000),
    ("EBV-00003", "e-Defter", "EBR-003", "ZORUNLU", "Hayır", 3_000_000),
    ("EBV-00004", "e-İrsaliye", "EBR-004", "ESIK_ALTI", "Hayır", 10_000_000),
    ("EBV-00005", "e-Fatura (şube)", "EBR-001", "HAZIRLIK", "Evet", 3_000_000),
    ("EBV-00006", "e-Arşiv (perakende)", "EBR-002", "ZORUNLU", "Evet", 500_000),
    ("EBV-00007", "e-Defter (yıl geçişi)", "EBR-003", "HAZIRLIK", "Hayır", 3_000_000),
    ("EBV-00008", "e-İrsaliye (lojistik)", "EBR-004", "ESIK_ALTI", "Hayır", 10_000_000),
]


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: e-belge zorunluluk kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Karar ve eşik sapması bozulur | Kaynak: GİB e-belge tebliğleri",
    )


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        _kim(ws, f"{get_column_letter(c)}{r}", baslik, mesaj or "Uygun türde değer girin")
    return hucre


def kural_cek(kural_id: str) -> str:
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-ebr_yilTaban))),'
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


def _demo_eylem():
    satirlar = []
    durumlar = ["PLANLANDI", "DEVAM", "BEKLIYOR", "TAMAMLANDI", "IPTAL",
                "DEVAM", "BEKLIYOR", "PLANLANDI"]
    oncekiler = ["", "PLANLANDI", "DEVAM", "BEKLIYOR", "DEVAM",
                 "PLANLANDI", "DEVAM", "PLANLANDI"]
    for i in range(1, DEMO_EYLEM + 1):
        kay = ENVANTER_DEMO[(i - 1) % DEMO_ENV][0]
        if i in (12, 15):
            kay = "" if i == 12 else "EBV-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 10:
            o, d = "PLANLANDI", "TAMAMLANDI"
        if i == 14:
            o, d = "TAMAMLANDI", "DEVAM"
        ilk = 40 + (i % 5) * 10
        hedef = max(5, ilk - (i % 4) * 8)
        if i in (8, 11):
            hedef = ilk + 50
        tar = RAPOR_TARIHI - timedelta(days=(i * 5) % 90)
        termin = RAPOR_TARIHI + timedelta(days=(i % 20) - 8)
        satirlar.append({
            "id": f"EYL-{i:05d}",
            "kaynak": kay,
            "tarih": tar,
            "termin": termin,
            "onceki": o,
            "durum": d,
            "sorumlu": f"Sorumlu {(i % 5) + 1}",
            "aciklama": f"e-Belge geçiş eylemi {i}",
            "ilk": ilk,
            "hedef": hedef,
            "tutar": max(2500, ilk * 180),
        })
    return satirlar


ORNEK_EYLEM = _demo_eylem()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=30)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=30)
    h(ws, 4, 1,
      "Brüt satış hasılatını girin; e-Fatura / e-Arşiv / e-Defter / e-İrsaliye "
      "eşiklerini yıllı kural tablosuyla karşılaştırıp geçiş kararını ve eylem planını üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:N4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Çok yıllı eşik tablosu (2025/2026/2027) — tebliğ güncellemesine hazır",
        "Zorunluluk kararı + madde atıflı KANIT_RAPORU",
        "Senaryo / duyarlılık / M-SEN eşik değişimi",
        "Envanter + eylem durum makinesi + yaşlandırma",
        "Belirsizlik beyanı: ciro hesap dönemi (yorum A / B)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "SMMM — müşteriye zorunluluk tarihini bildirir",
        "İşletme — geçiş planını yapar",
        "Denetçi — kanıt raporunu inceler",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) GIRDI sarı hücreler → 2) PANO karar → 3) EYLEM_PLANI → 4) KANIT_RAPORU yazdır.",
      kaydir=True)
    ws.merge_cells("A19:N19")
    h(ws, 21, 1, f"Sürüm {SURUM} | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 72})


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Firma ve ciro girişi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Eşikler KURALLAR'dan çekilir.", yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Firma Unvanı", DEMO["firma"], None, "firma", "Resmi unvanı yazın."),
        (7, "Dönem", DEMO["donem"], None, "donem", "Dönem etiketi (ör. 2025)."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili", "2025 / 2026 / 2027."),
        (9, "Brüt Satış Hasılatı [₺]", DEMO["brut_satis"], TL, "brut_satis",
         "KDV hariç brüt satış hasılatını TL girin."),
        (10, "Ciro Dönem Yorumu (A/B)", DEMO["ciro_donem_yorumu"], None, "ciro_yorum",
         "Belirsizlikte yorum A veya B seçin."),
        (11, "Sektör", DEMO["sektor"], None, "sektor", "Faaliyet sektörünü yazın."),
    ]
    h(ws, 5, 1, "Alan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 2, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 3, "Birim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    for satir, etiket, deger, sayi, _ad, mesaj in alanlar:
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3,
          "₺" if sayi == TL else ("yıl" if _ad == "hesapYili" else "metin"), yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 13, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B11)=0,ebr_girisBeklenen=0),0,'
      'ROUND(COUNTA(B6:B11)/ebr_girisBeklenen*100,0)),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)

    h(ws, 16, 1, "Giriş tablosu (otomasyon / boş-mod)", kalin=True, yazi=KOYU_LACIVERT)
    sutunlar = ["Alan Anahtarı", "Değer", "Birim", "Zorunlu", "Kaynak", "Kontrol"]
    baslik_satiri(ws, 17, sutunlar)
    formuller = {
        "Kontrol": (
            '=IF(tblGirdi[[#This Row],[Alan Anahtarı]]="","",'
            'IF(tblGirdi[[#This Row],[Değer]]="","EKSİK","TAM"))'
        ),
    }
    tablo_ekle(ws, "tblGirdi", "A17:F1016", sutunlar, formuller)
    ornek = [
        ("firma", DEMO["firma"], "metin", "Evet", "GIRDI"),
        ("donem", DEMO["donem"], "metin", "Evet", "GIRDI"),
        ("hesapYili", DEMO["hesapYili"], "yıl", "Evet", "GIRDI"),
        ("brut_satis", DEMO["brut_satis"], "TL", "Evet", "GIRDI"),
        ("ciro_yorum", DEMO["ciro_donem_yorumu"], "A/B", "Evet", "GIRDI"),
        ("sektor", DEMO["sektor"], "metin", "Evet", "GIRDI"),
    ]
    for i, (a, d, b, z, k) in enumerate(ornek, 18):
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
            _kim(ws, f"{get_column_letter(kolon)}{i}", a, "Tablo girişi")

    dogrulama(ws, "list", "ListeYillar", "B8",
              baslik="Hesap Yılı", mesaj="2025, 2026 veya 2027.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.", bos=False)
    dogrulama(ws, "list", "ListeYorumAB", "B10",
              baslik="Yorum", mesaj="A veya B seçin.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.", bos=False)
    dogrulama(ws, "decimal", "0", "B9",
              baslik="Brüt Satış", mesaj="0 ile 1e12 arasında TL.",
              hata_baslik="Geçersiz", hata_mesaj="0–1.000.000.000.000",
              isaret="between", f2="1000000000000")
    giris_hucreleri(ws, 6, 11, [2])
    giris_hucreleri(ws, 18, 23, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 36, "B": 28, "C": 12})
    alt_bant(ws, 1018, "Sarı alanlar giriş; kalite formülü kilitlidir.")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "e-Belge eşik kural tablosu (yıl yan yana)", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Eşikler MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.",
      yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    # Eşikler TL — GİB tarzı örnek parametreler (AYARLAR/yıl kolonundan güncellenir)
    satirlar = [
        ["EBR-001", "VUK e-Fatura", "efatura_esik", "brut_satis>=esik", "efatura_esik",
         3_000_000, 3_500_000, 4_000_000, "01.01.2025", "GİB e-Fatura tebliği"],
        ["EBR-002", "VUK e-Arşiv", "earsiv_esik", "brut_satis>=esik", "earsiv_esik",
         500_000, 600_000, 700_000, "01.01.2025", "GİB e-Arşiv tebliği"],
        ["EBR-003", "VUK e-Defter", "edefter_esik", "brut_satis>=esik", "edefter_esik",
         3_000_000, 3_500_000, 4_000_000, "01.01.2025", "GİB e-Defter tebliği"],
        ["EBR-004", "VUK e-İrsaliye", "eirsaliye_esik", "brut_satis>=esik", "eirsaliye_esik",
         10_000_000, 12_000_000, 14_000_000, "01.01.2025", "GİB e-İrsaliye tebliği"],
        ["EBR-005", "Genel", "yuvarlama", "her hesap", "yuvarlama_ondalik",
         2, 2, 2, "01.01.2025", "Uygulama notu"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = TL if k in (6, 7, 8) and s[0] != "EBR-005" else (CATI if k in (6, 7, 8) else None)
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı eşik; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-ebr_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "e-Fatura eşiği", yazi="333333")
    h(ws, 15, 2, kural_cek("EBR-001"), sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "e-Arşiv eşiği", yazi="333333")
    h(ws, 16, 2, kural_cek("EBR-002"), sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=_xlfn.TEXTJOIN(" | ",TRUE,"EBR-001=",TEXT(B15,"₺ #,##0"),'
      '"EBR-002=",TEXT(B16,"₺ #,##0"),"yıl=",hesapYili)',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    for col in ("F", "G", "H"):
        dogrulama(ws, "decimal", "0", f"{col}7:{col}11",
                  baslik="Kural değeri", mesaj="Yıl kolonuna sayısal parametre.",
                  hata_baslik="Geçersiz", hata_mesaj="0 veya üzeri.",
                  isaret="between", f2="1000000000000")
    genislik(ws, {get_column_letter(i): w for i, w in enumerate(
        [28, 18, 20, 22, 18, 14, 14, 14, 12, 24, 14], 1)})
    sabitle(ws, "A7")


def motor(ws):
    sayfa_hazirla(ws, "MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Hesap zinciri — her adımda kural_id", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Eşikler tblKurallar'dan çekilir; sabit eşik yoktur.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*MOTOR_BASLIK, "deger"])

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    # G7..G50 = adım 1..44 (satır = 6 + adım)
    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif mevzuat yılı", "EBR-005")  # G7
    ekle("e-Fatura eşiği", km("EBR-001"), "TL", "e-Fatura ciro eşiği", "EBR-001")  # G8
    ekle("e-Arşiv eşiği", km("EBR-002"), "TL", "e-Arşiv ciro eşiği", "EBR-002")  # G9
    ekle("e-Defter eşiği", km("EBR-003"), "TL", "e-Defter ciro eşiği", "EBR-003")  # G10
    ekle("e-İrsaliye eşiği", km("EBR-004"), "TL", "e-İrsaliye ciro eşiği", "EBR-004")  # G11
    ekle("Yuvarlama", km("EBR-005"), "adet", "Yuvarlama basamağı", "EBR-005")  # G12
    ekle("Brüt satış ham", "=GIRDI!B9", "TL", "Girdi ciro", "EBR-001")  # G13
    ekle("Yorum A ciro", "=G13*ebr_yorumACarpan", "TL", "Dönem yorumu A", "EBR-001")  # G14
    ekle("Yorum B ciro", "=G13", "TL", "Dönem yorumu B (tam)", "EBR-001")  # G15
    ekle("Seçili ciro", '=IF(GIRDI!B10="B",G15,G14)', "TL", "A/B seçimi", "EBR-001")  # G16
    ekle("Ciro (yuvarlanmış)",
         "=ROUND(G16,MAX(0,MIN(ebr_yuvarMax,IFERROR(N(G12),0))))", "TL",
         "Yuvarlanmış ciro", "EBR-005")  # G17
    ekle("e-Fatura zorunlu mu", '=IF(G17>=G8,1,0)', "bayrak", "Eşik aşımı", "EBR-001")  # G18
    ekle("e-Arşiv zorunlu mu", '=IF(G17>=G9,1,0)', "bayrak", "Eşik aşımı", "EBR-002")  # G19
    ekle("e-Defter zorunlu mu", '=IF(G17>=G10,1,0)', "bayrak", "Eşik aşımı", "EBR-003")  # G20
    ekle("e-İrsaliye zorunlu mu", '=IF(G17>=G11,1,0)', "bayrak", "Eşik aşımı", "EBR-004")  # G21
    ekle("Zorunlu belge adedi", "=G18+G19+G20+G21", "adet", "Kaç belge zorunlu", "EBR-001")  # G22
    ekle("e-Fatura mesafe", "=G17-G8", "TL", "Ciro−eşik", "EBR-001")  # G23
    ekle("e-Arşiv mesafe", "=G17-G9", "TL", "Ciro−eşik", "EBR-002")  # G24
    ekle("e-Defter mesafe", "=G17-G10", "TL", "Ciro−eşik", "EBR-003")  # G25
    ekle("e-İrsaliye mesafe", "=G17-G11", "TL", "Ciro−eşik", "EBR-004")  # G26
    ekle("En yakın eşik mesafesi",
         "=MIN(ABS(G23),ABS(G24),ABS(G25),ABS(G26))", "TL", "Radar mesafe", "EBR-001")  # G27
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(G22=0,"ESIK_ALTI",IF(G22>=ebr_kararEsik,"ZORUNLU","HAZIRLIK")))',
         "kod", "Karar kodu", "EBR-001")  # G28
    ekle("Karar metni",
         '=IF(G28="VERI_YOK","VERİ YOK",'
         'IF(G28="ESIK_ALTI","EŞİK ALTI / UYGUN",'
         'IF(G28="HAZIRLIK","HAZIRLIK / DİKKAT","ZORUNLU / GEÇİŞ")))',
         "metin", "Kullanıcıya karar", "EBR-001")  # G29
    ekle("Gerekçe",
         '=IF(G28="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G28="ESIK_ALTI","Ciro tüm eşiklerin altında; zorunluluk yok.",'
         'IF(G28="HAZIRLIK","Kısmi eşik aşımı; geçiş hazırlığı önerilir.",'
         '"Bir veya daha fazla e-belge zorunluluğu tetiklendi.")))',
         "metin", "Gerekçe cümlesi", "EBR-001")  # G30
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B11)=0,ebr_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B11)/ebr_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "EBR-005")  # G31
    ekle("Risk skoru",
         '=IF(G28="VERI_YOK",0,IF(G22=0,ebr_riskDusuk,'
         'IF(G22>=ebr_kararEsik,ebr_riskYuksek,ebr_riskOrta)))',
         "puan", "Zorunluluk riski", "EBR-001")  # G32
    ekle("Senaryo iyimser ciro", "=G17*ebr_senaryoIyi", "TL", "İyimser ciro", "EBR-001")  # G33
    ekle("Senaryo kötümser ciro", "=G17*ebr_senaryoKotu", "TL", "Kötümser ciro", "EBR-001")  # G34
    ekle("İyimser zorunlu adet",
         '=IF(G33>=G8,1,0)+IF(G33>=G9,1,0)+IF(G33>=G10,1,0)+IF(G33>=G11,1,0)',
         "adet", "İyimser adet", "EBR-001")  # G35
    ekle("Kötümser zorunlu adet",
         '=IF(G34>=G8,1,0)+IF(G34>=G9,1,0)+IF(G34>=G10,1,0)+IF(G34>=G11,1,0)',
         "adet", "Kötümser adet", "EBR-001")  # G36
    ekle("M-SEN e-Fatura eşiği", "=G8*ebr_esikArtisCarpan", "TL", "Eşik artışı", "EBR-001")  # G37
    ekle("M-SEN zorunlu adet",
         '=IF(G17>=G37,1,0)+IF(G17>=G9,1,0)+IF(G17>=G10,1,0)+IF(G17>=G11,1,0)',
         "adet", "Kural değişimi adet", "EBR-001")  # G38
    ekle("M-SEN fark (adet)", "=G38-G22", "adet", "M-SEN delta", "EBR-001")  # G39
    ekle("Tahmin üst", "=G22*ebr_tahminUst", "adet", "Tahmin üst", "EBR-005")  # G40
    ekle("Tahmin alt", "=G22*ebr_tahminAlt", "adet", "Tahmin alt", "EBR-005")  # G41
    ekle("Tornado: ciro etkisi",
         "=ABS((IF(G17*ebr_tornadoCiro>=G8,1,0)+IF(G17*ebr_tornadoCiro>=G9,1,0)"
         "+IF(G17*ebr_tornadoCiro>=G10,1,0)+IF(G17*ebr_tornadoCiro>=G11,1,0))-G22)",
         "adet", "Ciro ± etki", "EBR-001")  # G42
    ekle("Tornado: eşik etkisi", "=ABS(G38-G22)", "adet", "Eşik ± etki", "EBR-001")  # G43
    ekle("Tornado: yorum etkisi",
         "=ABS(IF(G14>=G8,1,0)+IF(G14>=G9,1,0)+IF(G14>=G10,1,0)+IF(G14>=G11,1,0)"
         "-(IF(G15>=G8,1,0)+IF(G15>=G9,1,0)+IF(G15>=G10,1,0)+IF(G15>=G11,1,0)))",
         "adet", "Yorum A/B etki", "EBR-001")  # G44
    ekle("Yetim eylem adedi",
         '=COUNTIF(tblEylem[YetimBayrak],"YETİM")', "adet", "Yetim kayıt", "EBR-005")  # G45
    ekle("Geçersiz geçiş adedi",
         '=COUNTIF(tblEylem[GecisUyarisi],"GEÇERSİZ GEÇİŞ")', "adet", "Geçiş uyarısı", "EBR-005")  # G46
    ekle("Zincir kırık adedi",
         '=COUNTIF(tblEylem[ZincirBayrak],"ZİNCİR KIRIK")', "adet", "Zincir bayrağı", "EBR-005")  # G47
    ekle("Eylem yaş özeti",
         '=_xlfn.TEXTJOIN("/",TRUE,COUNTIF(tblEylem[YasKova],"0-7"),'
         'COUNTIF(tblEylem[YasKova],"8-30"),COUNTIF(tblEylem[YasKova],"31-60"),'
         'COUNTIF(tblEylem[YasKova],"60+"))',
         "metin", "Yaş kovaları", "EBR-005")  # G48
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"EBR-001→e-Fatura","EBR-002→e-Arşiv",'
         '"EBR-003→e-Defter","EBR-004→e-İrsaliye")',
         "metin", "Kanıt atıfları", "EBR-001")  # G49
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Ciro=",TEXT(G17,"₺ #,##0"),'
         '"Zorunlu=",TEXT(G22,"0"),"Karar=",G29)',
         "metin", "Rapor özeti", "EBR-005")  # G50

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
        elif birim in ("puan", "yıl", "adet", "bayrak"):
            ws.cell(r, 7).number_format = CATI
            ws.cell(r, 3).number_format = CATI

    genislik(ws, {"A": 8, "B": 36, "C": 55, "D": 10, "E": 26, "F": 12, "G": 22})
    sabitle(ws, "A7")
    return len(adimlar)


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    sabitle(ws, "A3")
    h(ws, 3, 1, "Karar Destek Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 3, 3, "Rapor Tarihi", yazi=GRİ)
    h(ws, 3, 4, "=raporTarihi", sayi=TARİH)

    h(ws, 4, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2, "=ebr_kararMetni", kalin=True, boyut=14, yazi=NORMAL)
    _kim(ws, "B4", "Karar", "Otomatik; boş girdide VERİ YOK")
    h(ws, 5, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 2, "=ebr_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B5:H5")

    kpi = [
        (7, "Brüt Satış", "=IFERROR(ebr_ciro,0)", TL),
        (8, "Zorunlu Belge", "=IFERROR(ebr_zorunluAdet,0)", CATI),
        (9, "e-Fatura Eşik", "=IFERROR(MOTOR!G8,0)", TL),
        (10, "e-Arşiv Eşik", "=IFERROR(MOTOR!G9,0)", TL),
        (11, "e-Defter Eşik", "=IFERROR(MOTOR!G10,0)", TL),
        (12, "e-İrsaliye Eşik", "=IFERROR(MOTOR!G11,0)", TL),
        (13, "Güven", "=IFERROR(MOTOR!G31,0)", CATI),
        (14, "Risk", "=IFERROR(MOTOR!G32,0)", CATI),
        (15, "Yetim", "=IFERROR(MOTOR!G45,0)", CATI),
        (16, "Geçersiz Geçiş", "=IFERROR(MOTOR!G46,0)", CATI),
        (17, "Zincir Kırık", "=IFERROR(MOTOR!G47,0)", CATI),
        (18, "Eylem Yaş", "=IFERROR(ebr_eylemYas,\"-\")", None),
        (19, "Senaryo Bant", "=IFERROR(ebr_senaryoKarsilastirma,0)", CATI),
        (20, "M-SEN Fark", "=IFERROR(ebr_kuralDegisimSenaryo,0)", CATI),
        (21, "Tahmin", "=IFERROR(ebr_tahminAralik,0)", CATI),
        (22, "Tornado Max", "=IFERROR(MAX(MOTOR!G42:G44),0)", CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 7, 4, "Analitik modüller", kalin=True, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 8, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Eylem yaş", "=ebr_eylemYas", "=ebr_yorumEylemYas"),
        ("T2", "Zorunlu pay", "=ebr_modulT2", "=ebr_modulT2Yorum"),
        ("O2", "Duyarlılık", "=ebr_duyarlilikSira", "=ebr_yorumDuyarlilik"),
        ("O3", "Senaryo bant", "=ebr_senaryoKarsilastirma", "=ebr_yorumSenaryo"),
        ("M", "M-SEN fark", "=ebr_kuralDegisimSenaryo", "=ebr_yorumKuralDegisim"),
        ("I1", "Tahmin aralık", "=ebr_tahminAralik", "=ebr_yorumTahmin"),
        ("I2", "Ciro P90", "=ebr_modulI2", "=ebr_modulI2Yorum"),
    ]
    for i, (kod, ad, form, yorum) in enumerate(moduller):
        r = 9 + i
        h(ws, r, 1, kod, kalin=True, hiza="center", yazi="B08948")
        h(ws, r, 2, ad)
        h(ws, r, 3, form, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    # A05 kuyruk
    h(ws, 18, 4, "Açık Kuyruk", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 4, "Planlandı")
    h(ws, 19, 5, '=COUNTIFS(tblEylem[Durum],"PLANLANDI",tblEylem[Tutar],">0")', sayi=CATI)
    h(ws, 19, 6, '=SUMIFS(tblEylem[Tutar],tblEylem[Durum],"PLANLANDI")', sayi=TL)
    h(ws, 20, 4, "Devam")
    h(ws, 20, 5, '=COUNTIFS(tblEylem[Durum],"DEVAM",tblEylem[Tutar],">0")', sayi=CATI)
    h(ws, 20, 6, '=SUMIFS(tblEylem[Tutar],tblEylem[Durum],"DEVAM")', sayi=TL)
    h(ws, 21, 4, "Bekliyor")
    h(ws, 21, 5, '=COUNTIFS(tblEylem[Durum],"BEKLIYOR",tblEylem[Tutar],">0")', sayi=CATI)
    h(ws, 21, 6, '=SUMIFS(tblEylem[Tutar],tblEylem[Durum],"BEKLIYOR")', sayi=TL)

    # Grafik demo kaynakları
    h(ws, 25, 1, "Belge")
    h(ws, 25, 2, "Zorunlu")
    for i, (ad, v) in enumerate([("e-Fatura", 1), ("e-Arşiv", 1), ("e-Defter", 1), ("e-İrsaliye", 0)], 26):
        h(ws, i, 1, ad)
        h(ws, i, 2, v, sayi=CATI)
    h(ws, 25, 4, "Yaş Kova")
    h(ws, 25, 5, "Adet")
    for i, (ad, v) in enumerate([("0-7", 4), ("8-30", 5), ("31-60", 4), ("60+", 3)], 26):
        h(ws, i, 4, ad)
        h(ws, i, 5, v, sayi=CATI)
    h(ws, 25, 7, "Senaryo")
    h(ws, 25, 8, "Adet")
    for i, (ad, v) in enumerate([("İyimser", 2), ("Baz", 3), ("Kötümser", 4), ("M-SEN", 2)], 26):
        h(ws, i, 7, ad)
        h(ws, i, 8, v, sayi=CATI)
    h(ws, 25, 10, "Uyarı")
    h(ws, 25, 11, "Adet")
    for i, (ad, v) in enumerate([("Yetim", 2), ("Geçersiz", 2), ("Zincir", 2), ("Geciken", 3)], 26):
        h(ws, i, 10, ad)
        h(ws, i, 11, v, sayi=CATI)
    h(ws, 25, 13, "Hafta")
    h(ws, 25, 14, "Yeni")
    for i in range(8):
        h(ws, 26 + i, 13, i + 1, sayi=CATI)
        h(ws, 26 + i, 14, 1 + (i % 3), sayi=CATI)

    charts = [
        (PieChart, "Belge Zorunluluk", 2, 25, 29, 1, 26, 29, "A35"),
        (BarChart, "Eylem Yaşlandırma", 5, 25, 29, 4, 26, 29, "F35"),
        (BarChart, "Senaryo Karşılaştırma", 8, 25, 29, 7, 26, 29, "A50"),
        (BarChart, "Uyarı Türleri", 11, 25, 29, 10, 26, 29, "F50"),
        (LineChart, "Haftalık Yeni", 14, 25, 33, 13, 26, 33, "A65"),
        (PieChart, "Yaş Payı", 5, 25, 29, 4, 26, 29, "F65"),
        (BarChart, "Belge Adet", 2, 25, 29, 1, 26, 29, "A80"),
        (BarChart, "Senaryo Pay", 8, 25, 29, 7, 26, 29, "F80"),
    ]
    for cls, title, dc, dr0, dr1, cc, cr0, cr1, anchor in charts:
        c = cls()
        c.title = title
        c.add_data(Reference(ws, min_col=dc, min_row=dr0, max_row=dr1), titles_from_data=True)
        c.set_categories(Reference(ws, min_col=cc, min_row=cr0, max_row=cr1))
        c.width, c.height = 10, 7
        ws.add_chart(c, anchor)

    baski_hazirla(ws, "A1:N24", f"{URUN_AD} | Pano | {SURUM}")
    genislik(ws, {"A": 42, "B": 18, "C": 14, "D": 28, "E": 14, "F": 14})


def envanter(ws):
    sayfa_hazirla(ws, "ENVANTER", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "e-Belge envanteri — her kaydın tek kimliği vardır", son_kolon=12)
    sutunlar = [
        "BelgeId", "BelgeAdi", "KuralId", "DurumEtiket", "Aktif", "EsikReferans",
        "KayitDolu", "ZorunluBayrak",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "KayitDolu": 'IF(tblEnvanter[[#This Row],[BelgeId]]="","",1)',
        "ZorunluBayrak": (
            'IF(tblEnvanter[[#This Row],[BelgeId]]="","",'
            'IF(OR(tblEnvanter[[#This Row],[DurumEtiket]]="ZORUNLU",'
            'tblEnvanter[[#This Row],[DurumEtiket]]="HAZIRLIK"),1,0))'
        ),
    }
    for i, (bid, ad, kid, durum, aktif, esik) in enumerate(ENVANTER_DEMO):
        r = ILK + i
        _sari(ws, r, 1, bid, baslik="Belge Kimliği", mesaj="EBV-#####")
        _sari(ws, r, 2, ad, baslik="Belge Adı", mesaj="e-belge türü")
        _sari(ws, r, 3, kid, baslik="Kural", mesaj="EBR-001..004")
        _sari(ws, r, 4, durum, baslik="Durum", mesaj="ZORUNLU/HAZIRLIK/ESIK_ALTI")
        _sari(ws, r, 5, aktif, baslik="Aktif", mesaj="Evet veya Hayır")
        _sari(ws, r, 6, esik, sayi=TL, baslik="Eşik Ref", mesaj="Referans eşik TL")
    for r in range(ILK + DEMO_ENV, SON + 1):
        for c in range(1, 7):
            fmt = TL if c == 6 else None
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblEnvanter", f"A{HDR}:H{SON}", sutunlar, formuller=form)
    dogrulama(ws, "list", "ListeEvetHayir", f"E{ILK}:E{SON}",
              baslik="Aktif", mesaj="Evet veya Hayır",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeKural", f"C{ILK}:C{SON}",
              baslik="Kural", mesaj="EBR kodu",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 9)})
    ws.column_dimensions["B"].width = 22


def eylem_plani(ws):
    sayfa_hazirla(ws, "EYLEM_PLANI", RENK, URUN_AD, son_kolon=22)
    alt_bant(ws, 3, "Geçiş eylemleri — durum makinesi + yetim + geçiş + zincir", son_kolon=16)
    sutunlar = [
        "EylemId", "KaynakBelgeId", "Tarih", "Termin", "OncekiDurum", "Durum",
        "Sorumlu", "Aciklama", "IlkSkor", "HedefSkor", "Tutar",
        "YasGun", "YasKova", "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    yas_f = yaslandirma_gun("tblEylem[[#This Row],[Tarih]]", "raporTarihi").lstrip("=")
    kova_f = yas_kova_formul("tblEylem[[#This Row],[YasGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblEylem[[#This Row],[KaynakBelgeId]]", "tblEnvanter[BelgeId]").lstrip("=")
    gecis_birlesik = 'tblEylem[[#This Row],[OncekiDurum]]&"|"&tblEylem[[#This Row],[Durum]]'
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    zincir_f = zincir_kirik_formul(
        "tblEylem[[#This Row],[IlkSkor]]", "tblEylem[[#This Row],[HedefSkor]]",
        "zincirTolerans").lstrip("=")
    form = {
        "YasGun": f'IF(tblEylem[[#This Row],[EylemId]]="","",{yas_f})',
        "YasKova": f'IF(tblEylem[[#This Row],[EylemId]]="","",{kova_f})',
        "YetimBayrak": f'IF(tblEylem[[#This Row],[EylemId]]="","",{yetim_f})',
        "GecisUyarisi": (
            f'IF(OR(tblEylem[[#This Row],[EylemId]]="",'
            f'tblEylem[[#This Row],[OncekiDurum]]=""),"",{gecis_f})'
        ),
        "ZincirBayrak": f'IF(tblEylem[[#This Row],[EylemId]]="","",{zincir_f})',
        "KayitDolu": 'IF(tblEylem[[#This Row],[EylemId]]="","",1)',
    }
    for i, s in enumerate(ORNEK_EYLEM):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="Eylem Kimliği", mesaj="EYL-#####")
        _sari(ws, r, 2, s["kaynak"] or None, baslik="Kaynak Belge", mesaj="ENVANTER BelgeId")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Açılış tarihi")
        _sari(ws, r, 4, s["termin"], sayi=TARİH, baslik="Termin", mesaj="Hedef bitiş")
        _sari(ws, r, 5, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 6, s["durum"], baslik="Durum", mesaj="Durum haritasından")
        _sari(ws, r, 7, s["sorumlu"], baslik="Sorumlu", mesaj="Sorumlu kişi")
        _sari(ws, r, 8, s["aciklama"], baslik="Açıklama", mesaj="Kısa açıklama")
        _sari(ws, r, 9, s["ilk"], sayi=CATI, baslik="İlk Skor", mesaj="Başlangıç")
        _sari(ws, r, 10, s["hedef"], sayi=CATI, baslik="Hedef Skor", mesaj="Hedef")
        _sari(ws, r, 11, s["tutar"], sayi=TL, baslik="Tutar", mesaj="Maliyet TL")
    for r in range(ILK + DEMO_EYLEM, SON + 1):
        for c in range(1, 12):
            fmt = TARİH if c in (3, 4) else (TL if c == 11 else (CATI if c in (9, 10) else None))
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblEylem", f"A{HDR}:Q{SON}", sutunlar, formuller=form)
    dogrulama(ws, "list", "ListeDurumlar", f"E{ILK}:E{SON}",
              baslik="Önceki Durum", mesaj="Durum seçin",
              hata_baslik="Liste", hata_mesaj="Listeden")
    dogrulama(ws, "list", "ListeDurumlar", f"F{ILK}:F{SON}",
              baslik="Durum", mesaj="Durum seçin",
              hata_baslik="Liste", hata_mesaj="Listeden")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 18)})
    ws.column_dimensions["H"].width = 28


def akis(ws):
    sayfa_hazirla(ws, "AKIS", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Eylem akış özeti — KAYNAK_KART_ID / yetim / kuyruk", son_kolon=10)
    h(ws, 5, 1, "KAYNAK_KART_ID kontrolü", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "Yetim (YETİM) adet")
    h(ws, 6, 2, '=COUNTIF(tblEylem[YetimBayrak],"YETİM")', sayi=CATI)
    h(ws, 7, 1, "Geçersiz geçiş")
    h(ws, 7, 2, '=COUNTIF(tblEylem[GecisUyarisi],"GEÇERSİZ GEÇİŞ")', sayi=CATI)
    h(ws, 8, 1, "Zincir kırık")
    h(ws, 8, 2, '=COUNTIF(tblEylem[ZincirBayrak],"ZİNCİR KIRIK")', sayi=CATI)
    h(ws, 10, 1, "Kuyruk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 10, 2, "Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 10, 3, "Tutar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, (ad, d) in enumerate([
        ("Planlandı", "PLANLANDI"), ("Devam", "DEVAM"), ("Bekliyor", "BEKLIYOR"),
    ], 11):
        h(ws, i, 1, ad)
        h(ws, i, 2, kuyruk_sayim("tblEylem", d), sayi=CATI)
        h(ws, i, 3, kuyruk_tutar("tblEylem", d), sayi=TL)
    h(ws, 15, 1, "Kapalı (Tamamlandı+İptal) kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 1, "Tamamlandı")
    h(ws, 16, 2, kuyruk_sayim("tblEylem", "TAMAMLANDI"), sayi=CATI)
    h(ws, 17, 1, "İptal")
    h(ws, 17, 2, kuyruk_sayim("tblEylem", "IPTAL"), sayi=CATI)
    genislik(ws, {"A": 42, "B": 14, "C": 16})


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve duyarlılık", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Ciro Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Zorunlu Adet", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    rows = [
        ("İyimser", "=IFERROR(ebr_senaryoIyi+N(GIRDI!B9)*0,0.9)", "=IFERROR(MOTOR!G35,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(C6=0,"EŞİK ALTI","ZORUNLU"))'),
        ("Baz", "=IFERROR(1+N(GIRDI!B9)*0,1)", "=IFERROR(ebr_zorunluAdet,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(C7=0,"EŞİK ALTI","ZORUNLU"))'),
        ("Kötümser", "=IFERROR(ebr_senaryoKotu+N(GIRDI!B9)*0,1.1)", "=IFERROR(MOTOR!G36,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(C8=0,"EŞİK ALTI","ZORUNLU"))'),
        ("M-SEN eşik artışı", "=IFERROR(1+N(GIRDI!B9)*0,1)", "=IFERROR(MOTOR!G38,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(C9=0,"EŞİK ALTI","ZORUNLU"))'),
    ]
    for i, (ad, carp, adet, kar) in enumerate(rows, 6):
        h(ws, i, 1, ad)
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, adet, sayi=CATI, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant (kötümser−iyimser)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(C8-C6,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"İyimser ",TEXT(C6,"0")," · Baz ",TEXT(C7,"0"),'
      '" · Kötümser ",TEXT(C8,"0")," · Bant ",TEXT(B11,"0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Tornado ciro etkisi")
    h(ws, 16, 2, "=IFERROR(MOTOR!G42,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado eşik etkisi")
    h(ws, 17, 2, "=IFERROR(MOTOR!G43,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado yorum etkisi")
    h(ws, 18, 2, "=IFERROR(MOTOR!G44,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 3, '=IFERROR(IF(B16>=MAX(B16:B18),1,IF(B16=MEDIAN(B16:B18),2,3)),3)', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 3, '=IFERROR(IF(B17>=MAX(B16:B18),1,IF(B17=MEDIAN(B16:B18),2,3)),3)', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 3, '=IFERROR(IF(B18>=MAX(B16:B18),1,IF(B18=MEDIAN(B16:B18),2,3)),3)', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(CONCATENATE("1)",INDEX(A16:A18,MATCH(1,C16:C18,0)),'
      '" 2)",INDEX(A16:A18,MATCH(2,C16:C18,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B16:B18),"0")," — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN kural değişimi farkı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=IFERROR(MOTOR!G39,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Eşik artışı senaryosunda zorunlu adet farkı ",TEXT(B22,"0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=MOTOR!G41", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 3, "=ebr_zorunluAdet", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 4, "=MOTOR!G40", sayi=CATI, yazi=DIS_REF_YEŞIL)
    for i, v in enumerate([1, 2, 3, 4], 28):
        h(ws, i, 1, v, sayi=CATI)
        h(ws, i, 2, f"=C{5+v}", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 32, 1, "Tahmin sonraki")
    h(ws, 32, 2, "=IFERROR(FORECAST.LINEAR(5,B28:B31,A28:A31),0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Tahmin yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Tahmin ",TEXT(B32,"0")," · Alt ",TEXT(B25,"0"),'
      '" · Üst ",TEXT(D25,"0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 5, 7, "AdetDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([2, 3, 4, 2], 6):
        h(ws, i, 7, v, sayi=CATI)
    g1 = BarChart()
    g1.title = "Senaryo Zorunlu Adet"
    g1.add_data(Reference(ws, min_col=7, min_row=5, max_row=9), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=6, max_row=9))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "F5")
    genislik(ws, {"A": 36, "B": 16, "C": 14, "D": 18})


def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Tebliğ tarzı altın vakalar", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 5, VAKA_BASLIK)
    # V1: 4.5M ≥ 3M e-Fatura → 1 (sadece EBR-001 sayımı basitleştirilmiş: zorunlu adet beklenen)
    # Basit: e-Fatura zorunlu bayrağı = 1
    data = [
        ("V-001", "Ciro 4.5M — e-Fatura zorunlu",
         "ciro=vaka_ciro_1; eşik=EBR-001/2025", 1,
         '=IFERROR(IF(vaka_ciro_1>=INDEX(tblKurallar[deger_2025],MATCH("EBR-001",tblKurallar[kural_id],0)),1,0),0)',
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=ebr_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "GİB e-Fatura örnek"),
        ("V-002", "Ciro 400K — e-Arşiv altı",
         "ciro=vaka_ciro_2; eşik=EBR-002/2025", 0,
         '=IFERROR(IF(vaka_ciro_2>=INDEX(tblKurallar[deger_2025],MATCH("EBR-002",tblKurallar[kural_id],0)),1,0),0)',
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=ebr_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "GİB e-Arşiv örnek"),
        ("V-003", "Ciro 600K — e-Arşiv zorunlu",
         "ciro=vaka_ciro_3; eşik=EBR-002/2025", 1,
         '=IFERROR(IF(vaka_ciro_3>=INDEX(tblKurallar[deger_2025],MATCH("EBR-002",tblKurallar[kural_id],0)),1,0),0)',
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=ebr_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "GİB e-Arşiv örnek"),
        ("V-004", "Ciro 11M — e-İrsaliye zorunlu",
         "ciro=vaka_ciro_4; eşik=EBR-004/2025", 1,
         '=IFERROR(IF(vaka_ciro_4>=INDEX(tblKurallar[deger_2025],MATCH("EBR-004",tblKurallar[kural_id],0)),1,0),0)',
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=ebr_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "GİB e-İrsaliye örnek"),
    ]
    for i, row in enumerate(data, 6):
        for k, v in enumerate(row, 1):
            sayi = CATI if k in (4, 5, 6) and not (isinstance(v, str) and v.startswith("=")) else None
            h(ws, i, k, v, sayi=sayi, yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if isinstance(v, str) and v.startswith("=") and k in (5, 6):
                ws.cell(i, k).number_format = CATI

    h(ws, 12, 1, "Vaka durum özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=_xlfn.TEXTJOIN(" | ",TRUE,IF(G6="TUTARLI","V1 OK",G6),IF(G7="TUTARLI","V2 OK",G7),'
      'IF(G8="TUTARLI","V3 OK",G8),IF(G9="TUTARLI","V4 OK",G9))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    h(ws, 13, 1, "Vaka fark toplamı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2, "=IFERROR(SUM(F6:F9),0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    genislik(ws, {get_column_letter(i): w for i, w in enumerate(
        [10, 36, 36, 12, 14, 10, 12, 22], 1)})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Denetim Kanıt Raporu — e-Belge Zorunluluk", boyut=16, son_kolon=10)
    for i, (etiket, deger) in enumerate(rapor_baslik_satirlari(URUN_AD, SURUM), 4):
        h(ws, i, 1, etiket, yazi=GRİ)
        h(ws, i, 2, deger)
    h(ws, 8, 1, "Firma", yazi=GRİ)
    h(ws, 8, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Karar", kalin=True)
    h(ws, 9, 2, "=ebr_kararMetni", kalin=True, boyut=14)
    h(ws, 10, 1, "Gerekçe")
    h(ws, 10, 2, "=ebr_kararGerekce", kaydir=True)
    ws.merge_cells("B10:H10")
    h(ws, 12, 1, "Brüt Satış")
    h(ws, 12, 2, "=ebr_ciro", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "Zorunlu Adet")
    h(ws, 13, 2, "=ebr_zorunluAdet", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "Madde Atıf")
    h(ws, 14, 2, "=ebr_maddeAtifMetni", kaydir=True, yazi=DIS_REF_YEŞIL)
    ws.merge_cells("B14:H14")
    h(ws, 15, 1, "Kanıt Özeti")
    h(ws, 15, 2, "=ebr_kanitRaporu", kaydir=True, yazi=DIS_REF_YEŞIL)
    ws.merge_cells("B15:H15")
    h(ws, 16, 1, "Vaka Durum")
    h(ws, 16, 2, "=ebr_vakaDurum", kaydir=True, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Uyarılar", kalin=True, yazi=KRITIK)
    h(ws, 19, 1, "Yetim / Geçiş / Zincir")
    h(ws, 19, 2,
      '=_xlfn.TEXTJOIN(" | ",TRUE,"Yetim=",TEXT(MOTOR!G45,"0"),'
      '"Geçiş=",TEXT(MOTOR!G46,"0"),"Zincir=",TEXT(MOTOR!G47,"0"))',
      yazi=DIS_REF_YEŞIL)
    for i, (etiket, deger) in enumerate(imza_alani(), 22):
        h(ws, i, 1, etiket, yazi=GRİ)
        h(ws, i, 2, deger if deger != "" else None, zemin=GIRIS_SARI if deger == "" else None)
        if deger == "":
            ws.cell(i, 2).protection = Protection(locked=False)
            _kim(ws, f"B{i}", etiket, "İmza / onay alanı")
        elif isinstance(deger, str) and deger.startswith("="):
            ws.cell(i, 2).value = deger
    h(ws, 30, 1,
      "Bu rapor karar destek çıktısıdır; GİB başvurusu veya mali müşavir görüşü yerine geçmez.",
      yazi=GRİ, boyut=9, kaydir=True)
    baski_hazirla(ws, "A1:H32", f"{URUN_AD} | Kanıt | {SURUM}")
    genislik(ws, {"A": 28, "B": 50})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", "5B9BD5", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Doğrulama listeleri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Yıl", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, y in enumerate([2025, 2026, 2027], 6):
        h(ws, i, 1, y, sayi=CATI)
    h(ws, 5, 3, "Yorum", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 6, 3, "A")
    h(ws, 7, 3, "B")
    h(ws, 5, 5, "EvetHayır", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 6, 5, "Evet")
    h(ws, 7, 5, "Hayır")
    h(ws, 5, 7, "Durum", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, d in enumerate(["PLANLANDI", "DEVAM", "BEKLIYOR", "TAMAMLANDI", "IPTAL"], 6):
        h(ws, i, 7, d)
    h(ws, 5, 9, "İzinli Geçiş", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, g in enumerate(IZINLI_GECIS, 6):
        h(ws, i, 9, g)
    h(ws, 5, 11, "Kural", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, k in enumerate(["EBR-001", "EBR-002", "EBR-003", "EBR-004", "EBR-005"], 6):
        h(ws, i, 11, k)
    genislik(ws, {"A": 10, "C": 10, "E": 12, "G": 14, "I": 28, "K": 12})



def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Canlı kontrol paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    kontroller_list = [
        ("Giriş tablosu boş mu?",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"BOŞ","DOLU")'),
        ("Negatif brüt satış?",
         '=IF(GIRDI!B9<0,"NEGATİF","TEMİZ")'),
        ("Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(ebr_yilTaban),N(hesapYili)<=N(ebr_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        ("Ciro sıfır mı?",
         '=IF(IFERROR(ebr_ciro,0)=0,"SIFIR","NORMAL")'),
        ("Vaka fark toplamı",
         "=IFERROR(ebr_vakaFark,0)"),
        ("Yetim kayıt var mı?",
         '=IF(IFERROR(MOTOR!G45,0)>0,"YETİM VAR","TEMİZ")'),
        ("Geçersiz geçiş var mı?",
         '=IF(IFERROR(MOTOR!G46,0)>0,"GEÇERSİZ VAR","TEMİZ")'),
        ("Motor adım sayısı",
         "=COUNTA(MOTOR!B7:B50)"),
    ]
    for i, (ad, form) in enumerate(kontroller_list, 5):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Canlı formül | Değiştirmeyin")
    h(ws, 14, 1, "Beklenen giriş", yazi="333333")
    h(ws, 14, 2, "=IFERROR(ebr_girisBeklenen+N(GIRDI!B9)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Dolu giriş", yazi="333333")
    h(ws, 15, 2, "=COUNTA(GIRDI!B6:B11)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Kalite skoru", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=IFERROR(ROUND(IF(OR(B14=0,B14=\"\"),0,B15/B14*100),0),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 17, 1, "Zorunlu adet kontrol", yazi="333333")
    h(ws, 17, 2, "=IFERROR(ebr_zorunluAdet,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Karar kontrol", yazi="333333")
    h(ws, 18, 2, '=IFERROR(ebr_kararMetni,"-")', yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Zincir kırık kontrol", yazi="333333")
    h(ws, 19, 2, "=IFERROR(MOTOR!G47,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 40, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Örnek senaryo kütüphanesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Demo değerler; GIRDI'ye kopyalanabilir.", yazi=GRİ)
    sutunlar = ["Senaryo", "BrutSatis", "HesapYili", "CiroYorum", "EfaturaZorunlu", "EarsivZorunlu"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "EfaturaZorunlu": (
            '=IF(tblOrnek[[#This Row],[Senaryo]]="","",'
            'IFERROR(IF(tblOrnek[[#This Row],[BrutSatis]]>=INDEX(tblKurallar[deger_2025],'
            'MATCH("EBR-001",tblKurallar[kural_id],0)),1,0),0))'
        ),
        "EarsivZorunlu": (
            '=IF(tblOrnek[[#This Row],[Senaryo]]="","",'
            'IFERROR(IF(tblOrnek[[#This Row],[BrutSatis]]>=INDEX(tblKurallar[deger_2025],'
            'MATCH("EBR-002",tblKurallar[kural_id],0)),1,0),0))'
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:F1004", sutunlar, formuller)
    ornekler = [
        ("Demo ana", 4_500_000, 2025, "A"),
        ("Eşik altı", 400_000, 2025, "A"),
        ("e-Arşiv zorunlu", 600_000, 2025, "B"),
        ("e-İrsaliye", 11_000_000, 2025, "A"),
        ("Yıl 2026", 4_500_000, 2026, "A"),
    ]
    for i, (ad, ciro, yil, yorum) in enumerate(ornekler, 6):
        h(ws, i, 1, ad, zemin=GIRIS_SARI, yazi=GIRIS_YAZI)
        h(ws, i, 2, ciro, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=TL)
        h(ws, i, 3, yil, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=CATI)
        h(ws, i, 4, yorum, zemin=GIRIS_SARI, yazi=GIRIS_YAZI)
        for c in range(1, 5):
            ws.cell(i, c).protection = Protection(locked=False)
            _kim(ws, f"{get_column_letter(c)}{i}", sutunlar[c-1], "Örnek satır")
    for r in range(11, 1005):
        for c in range(1, 5):
            ws.cell(r, c).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
            ws.cell(r, c).protection = Protection(locked=False)
    genislik(ws, {"A": 20, "B": 16, "C": 12, "D": 12, "E": 16, "F": 16})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "7030A0", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Parametreler (tek kaynak) — GEÇERSİZ GEÇİŞ uyarısı EYLEM_PLANI ile bağlı",
      kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:F3")
    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Üretim / dönem", "01.01.2025", "Rapor tarihi (TODAY yasak)"),
        ("zincirTolerans", 30, "skor", "İş kuralı", "01.01.2025", "ZİNCİR KIRIK eşiği"),
        ("ebr_tolerans", 0.01, "adet", "Uygulama notu", "01.01.2025", "Vaka fark toleransı"),
        ("ebr_kararEsik", 2, "adet", "İç politika", "01.01.2025", "ZORUNLU min belge"),
        ("ebr_senaryoIyi", 0.9, "çarpan", "Senaryo motoru", "01.01.2025", "İyimser ciro çarpanı"),
        ("ebr_senaryoKotu", 1.15, "çarpan", "Senaryo motoru", "01.01.2025", "Kötümser ciro çarpanı"),
        ("ebr_esikArtisCarpan", 1.2, "çarpan", "M-SEN senaryo", "01.01.2025", "Eşik artışı çarpanı"),
        ("ebr_tahminAlt", 0.85, "çarpan", "Model varsayımı", "01.01.2025", "Tahmin alt"),
        ("ebr_tahminUst", 1.2, "çarpan", "Model varsayımı", "01.01.2025", "Tahmin üst"),
        ("ebr_tornadoCiro", 1.1, "çarpan", "Duyarlılık", "01.01.2025", "Tornado ciro"),
        ("ebr_yorumACarpan", 0.95, "çarpan", "Belirsizlik A", "01.01.2025", "Yorum A çarpanı"),
        ("ebr_yuzdelikOran", 0.9, "oran", "İstatistik", "01.01.2025", "P90 oranı"),
        ("ebr_parmakIzi", "bekleniyor", "metin", "Üretim", "01.01.2025", "SHA parmak izi"),
        ("ebr_girisBeklenen", 6, "adet", "Ürün mimarisi", "01.01.2025", "Beklenen giriş"),
        ("ebr_riskDusuk", 20, "puan", "İç politika", "01.01.2025", "Düşük risk"),
        ("ebr_riskOrta", 55, "puan", "İç politika", "01.01.2025", "Orta risk"),
        ("ebr_riskYuksek", 85, "puan", "İç politika", "01.01.2025", "Yüksek risk"),
        ("ebr_yilTaban", 2024, "yıl", "Ürün mimarisi", "01.01.2025", "CHOOSE yıl tabanı"),
        ("ebr_yuvarMax", 6, "adet", "Ürün mimarisi", "01.01.2025", "Yuvarlama tavanı"),
        ("vaka_ciro_1", 4_500_000, "TL", "Tebliğ örneği", "01.01.2025", "Vaka 1 ciro"),
        ("vaka_ciro_2", 400_000, "TL", "Tebliğ örneği", "01.01.2025", "Vaka 2 ciro"),
        ("vaka_ciro_3", 600_000, "TL", "Tebliğ örneği", "01.01.2025", "Vaka 3 ciro"),
        ("vaka_ciro_4", 11_000_000, "TL", "Tebliğ örneği", "01.01.2025", "Vaka 4 ciro"),
        ("ebr_esikGecikmeGun", 30, "gün", "İç politika", "01.01.2025", "Gecikme eşiği gün"),
    ]
    for i, (ana, deg, bir, kay, yur, acik) in enumerate(params):
        r = 6 + i
        h(ws, r, 1, ana)
        if isinstance(deg, date):
            h(ws, r, 2, deg, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=TARİH)
        elif isinstance(deg, float) and deg <= 2:
            h(ws, r, 2, deg, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=YÜZDE)
        elif isinstance(deg, (int, float)) and bir == "TL":
            h(ws, r, 2, deg, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=TL)
        elif isinstance(deg, (int, float)):
            h(ws, r, 2, deg, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=CATI)
        else:
            h(ws, r, 2, deg, zemin=GIRIS_SARI, yazi=GIRIS_YAZI)
        ws.cell(r, 2).protection = Protection(locked=False)
        _kim(ws, f"B{r}", ana, acik)
        h(ws, r, 3, bir)
        h(ws, r, 4, kay)
        h(ws, r, 5, yur)
        h(ws, r, 6, acik, kaydir=True)

    # Durum haritası — kolon H+ (anahtar kolonunu kirletme)
    h(ws, 32, 8, "Durum haritası (durum_id / izinli_sonraki / kategori kapalı)",
      kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 8, "durum_id", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 33, 9, "durum_adi", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 33, 10, "izinli_sonraki_durumlar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 33, 11, "kategori", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, d in enumerate(DURUM_HARITASI, 34):
        h(ws, i, 8, d["durum_id"])
        h(ws, i, 9, d["durum_adi"])
        h(ws, i, 10, d["izinli_sonraki_durumlar"])
        h(ws, i, 11, d["kategori"])
    h(ws, 40, 8, "GEÇERSİZ GEÇİŞ uyarısı EYLEM_PLANI GecisUyarisi kolonunda üretilir",
      yazi=KRITIK, kaydir=True)

    h(ws, 42, 8, "Modül köprü", kalin=True, yazi=KOYU_LACIVERT)
    bridges = [
        (43, "ebr_eylemYas", "=MOTOR!G48"),
        (44, "ebr_yorumEylemYas",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Yaş kovaları ",MOTOR!G48)'),
        (45, "ebr_modulT2", "=IFERROR(ebr_zorunluAdet/4,0)"),
        (46, "ebr_modulT2Yorum",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Zorunlu belge payı ",TEXT(I45,"0.0%"))'),
        (47, "ebr_duyarlilikSira", "=SENARYO!B19"),
        (48, "ebr_yorumDuyarlilik", "=SENARYO!B20"),
        (49, "ebr_senaryoKarsilastirma", "=SENARYO!B11"),
        (50, "ebr_yorumSenaryo", "=SENARYO!B12"),
        (51, "ebr_kuralDegisimSenaryo", "=SENARYO!B22"),
        (52, "ebr_yorumKuralDegisim", "=SENARYO!B23"),
        (53, "ebr_tahminAralik", "=SENARYO!B32"),
        (54, "ebr_yorumTahmin", "=SENARYO!B33"),
        (55, "ebr_modulI2",
         "=IFERROR(IF(COUNT(SENARYO!C6:C9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!C6:C9,ebr_yuzdelikOran)),0)"),
        (56, "ebr_modulI2Yorum",
         '=_xlfn.TEXTJOIN("; ",TRUE,"P90 zorunlu adet ",TEXT(I55,"0"))'),
        (57, "ebr_kararMetni", "=MOTOR!G29"),
        (58, "ebr_kararGerekce", "=MOTOR!G30"),
        (59, "ebr_maddeAtifMetni", "=MOTOR!G49"),
        (60, "ebr_kanitRaporu", "=MOTOR!G50"),
        (61, "ebr_vakaDurum", "=VAKALAR!B12"),
        (62, "ebr_vakaFark", "=VAKALAR!B13"),
        (63, "ebr_ciro", "=MOTOR!G17"),
        (64, "ebr_zorunluAdet", "=MOTOR!G22"),
        (65, "ebr_kuralYilMatris", "=KURALLAR!B17"),
    ]
    for r, ad, form in bridges:
        h(ws, r, 8, ad, yazi=GRİ)
        h(ws, r, 9, form, yazi=DIS_REF_YEŞIL, kaydir=True)

    genislik(ws, {"A": 24, "B": 18, "C": 12, "D": 18, "E": 14, "F": 36, "H": 28, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "548235", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Kullanım kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "1) GIRDI'de firma, yıl ve brüt satış hasılatını sarı hücrelere girin.",
        "2) PANO'da karar, zorunlu belge adedi ve riski izleyin.",
        "3) ENVANTER ve EYLEM_PLANI'da geçiş kartlarını güncelleyin.",
        "4) SENARYO ile eşik değişimini (M-SEN) ve duyarlılığı görün.",
        "5) VAKALAR tutarlılığını kontrol edin; KANIT_RAPORU'nu yazdırın.",
    ], 5):
        h(ws, i, 1, m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=12)
    h(ws, 12, 1, "Belirsizlik beyanı", kalin=True, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari([
        "ciro hesap dönemi (takvim yılı / özel hesap dönemi) — yorum A ve yorum B yan yana",
    ]), 13):
        h(ws, i, 1, m, kaydir=True, yazi="333333")
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=12)
    h(ws, 16, 1,
      "Bu dosya karar destek aracıdır; GİB başvurusu veya mali müşavir görüşü yerine geçmez.",
      yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 90})


def kosullu_bicim(wb):
    yesil = Font(color="006100", name=FONT, bold=True)
    kirmizi = Font(color="9C0006", name=FONT, bold=True)
    amber = Font(color="9C5700", name=FONT, bold=True)
    y_fill = PatternFill("solid", fgColor="C6EFCE")
    k_fill = PatternFill("solid", fgColor="FFC7CE")
    a_fill = PatternFill("solid", fgColor="FFEB9C")

    ws = wb["PANO"]
    for txt, fill, font in [
        ('"VERİ YOK"', k_fill, kirmizi),
        ('"EŞİK ALTI / UYGUN"', y_fill, yesil),
        ('"HAZIRLIK / DİKKAT"', a_fill, amber),
        ('"ZORUNLU / GEÇİŞ"', a_fill, amber),
    ]:
        ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=[txt], fill=fill, font=font))
    for r in range(7, 23):
        ws.conditional_formatting.add(f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=y_fill))
        ws.conditional_formatting.add(f"B{r}", CellIsRule(operator="lessThan", formula=["0"], fill=k_fill, font=kirmizi))

    ws = wb["EYLEM_PLANI"]
    for col, val, fill, font in [
        ("N", '"YETİM"', k_fill, kirmizi),
        ("O", '"GEÇERSİZ GEÇİŞ"', k_fill, kirmizi),
        ("P", '"ZİNCİR KIRIK"', k_fill, kirmizi),
        ("P", '"OK"', y_fill, yesil),
    ]:
        ws.conditional_formatting.add(
            f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=[val], fill=fill, font=font))
    for durum, fill in [("BEKLIYOR", a_fill), ("TAMAMLANDI", y_fill), ("IPTAL", k_fill)]:
        ws.conditional_formatting.add(
            f"F{ILK}:F{SON}", CellIsRule(operator="equal", formula=[f'"{durum}"'], fill=fill))

    ws = wb["ENVANTER"]
    for seviye, fill in [("ZORUNLU", k_fill), ("HAZIRLIK", a_fill), ("ESIK_ALTI", y_fill)]:
        ws.conditional_formatting.add(
            f"D{ILK}:D{SON}", CellIsRule(operator="equal", formula=[f'"{seviye}"'], fill=fill))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("G7:G50", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("F7:F50", FormulaRule(formula=['F7="EBR-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F50", FormulaRule(formula=['F7="EBR-004"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["VAKALAR"]
    for r in range(6, 10):
        ws.conditional_formatting.add(f"G{r}", CellIsRule(operator="equal", formula=['"TUTARLI"'], fill=y_fill, font=yesil))
        ws.conditional_formatting.add(f"G{r}", CellIsRule(operator="equal", formula=['"KIRIK"'], fill=k_fill, font=kirmizi))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B9", FormulaRule(formula=['ISNUMBER(SEARCH("UYGUN",B9))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B9", FormulaRule(formula=['ISNUMBER(SEARCH("DİKKAT",B9))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B9", FormulaRule(formula=['ISNUMBER(SEARCH("VERİ YOK",B9))'], font=kirmizi, fill=k_fill))
    for r in range(12, 17):
        ws.conditional_formatting.add(f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=y_fill))

    ws = wb["GIRDI"]
    ws.conditional_formatting.add("B13", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("B13", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))
    ws.conditional_formatting.add("B9", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))

    ws = wb["AKIS"]
    for r in range(6, 9):
        ws.conditional_formatting.add(f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=k_fill, font=kirmizi))
        ws.conditional_formatting.add(f"B{r}", CellIsRule(operator="equal", formula=["0"], fill=y_fill, font=yesil))
    for r in range(11, 14):
        ws.conditional_formatting.add(f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=a_fill))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("C6:C9", CellIsRule(operator="greaterThan", formula=["0"], fill=a_fill))
    ws.conditional_formatting.add("B16:B18", CellIsRule(operator="greaterThan", formula=["0"], fill=y_fill))


def ad_tanimlari(wb):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$6")
    ad_ekle(wb, "zincirTolerans", "AYARLAR!$B$7")

    ayar_map = {
        "ebr_tolerans": 8, "ebr_kararEsik": 9, "ebr_senaryoIyi": 10, "ebr_senaryoKotu": 11,
        "ebr_esikArtisCarpan": 12, "ebr_tahminAlt": 13, "ebr_tahminUst": 14,
        "ebr_tornadoCiro": 15, "ebr_yorumACarpan": 16, "ebr_yuzdelikOran": 17,
        "ebr_parmakIzi": 18, "ebr_girisBeklenen": 19, "ebr_riskDusuk": 20,
        "ebr_riskOrta": 21, "ebr_riskYuksek": 22, "ebr_yilTaban": 23, "ebr_yuvarMax": 24,
        "vaka_ciro_1": 25, "vaka_ciro_2": 26, "vaka_ciro_3": 27, "vaka_ciro_4": 28,
        "ebr_esikGecikmeGun": 29,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    bridge = {
        "ebr_eylemYas": 43, "ebr_yorumEylemYas": 44, "ebr_modulT2": 45, "ebr_modulT2Yorum": 46,
        "ebr_duyarlilikSira": 47, "ebr_yorumDuyarlilik": 48, "ebr_senaryoKarsilastirma": 49,
        "ebr_yorumSenaryo": 50, "ebr_kuralDegisimSenaryo": 51, "ebr_yorumKuralDegisim": 52,
        "ebr_tahminAralik": 53, "ebr_yorumTahmin": 54, "ebr_modulI2": 55, "ebr_modulI2Yorum": 56,
        "ebr_kararMetni": 57, "ebr_kararGerekce": 58, "ebr_maddeAtifMetni": 59,
        "ebr_kanitRaporu": 60, "ebr_vakaDurum": 61, "ebr_vakaFark": 62,
        "ebr_ciro": 63, "ebr_zorunluAdet": 64, "ebr_kuralYilMatris": 65,
    }
    for ad, satir in bridge.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${satir}")

    ad_ekle(wb, "ListeYillar", "LISTELER!$A$6:$A$8")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$C$6:$C$7")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$6:$E$7")
    ad_ekle(wb, "ListeDurumlar", "LISTELER!$G$6:$G$10")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$I$6:$I$13")
    ad_ekle(wb, "ListeKural", "LISTELER!$K$6:$K$10")


def main(cikti_yolu=None):
    wb = Workbook()
    wb.remove(wb.active)
    siralar = [
        ("KAPAK", kapak),
        ("PANO", pano),
        ("GIRDI", girdi),
        ("KURALLAR", kurallar),
        ("MOTOR", motor),
        ("ENVANTER", envanter),
        ("EYLEM_PLANI", eylem_plani),
        ("AKIS", akis),
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
    kosullu_bicim(wb)
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=ILK, satir_sonu=SON)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)
    wb.calculation.fullCalcOnLoad = True

    urun_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    dosya_ad = "EBelgeZorunlulukRadari.xlsx"
    hedef = cikti_yolu or os.path.join(urun_dir, dosya_ad)
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

    cikti = os.path.join(KOK, "cikti", dosya_ad)
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(hedef) != os.path.abspath(cikti):
        shutil.copy2(hedef, cikti)

    print(f"Dosya: {hedef}")
    print(f"Kopya: {cikti}")
    print(f"SHA-256: {dig}")
    print(f"Şifre: {SIFRE}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
