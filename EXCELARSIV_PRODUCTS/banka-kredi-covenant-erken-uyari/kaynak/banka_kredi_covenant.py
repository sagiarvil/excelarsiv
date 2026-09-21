#!/usr/bin/env python3
"""
Banka Kredi + Covenant Erken Uyarı — üretim betiği (A1+A3 / manda v6).
Covenant sapma + erken uyarı skoru + İZLEME→UYARI→İHLAL durum makinesi.
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

URUN_AD = "Banka Kredi + Covenant Erken Uyarı"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 11)

# Demo 2026: DSCR 1.45 ≥ 1.20; borç/FAVÖK 1.875 ≤ 3.5; cari 1.25 ≥ 1.0
# faiz yükü 10M×0.28/8M=0.35 ≤ 0.40; erken uyarı bandında değil → UYGUN
DEMO = {
    "sirket_adi": "Örnek Sanayi A.Ş.",
    "donem": "2026-Q2",
    "hesapYili": 2026,
    "kredi_bakiyesi": 10_000_000,
    "faiz_orani": 0.28,
    "vade_ay": 36,
    "favok": 8_000_000,
    "net_borc": 15_000_000,
    "cari_oran": 1.25,
    "dscr_gerceklesen": 1.45,
    "yorum_modu": "A",
}

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "IZLEME", "durum_adi": "İzleme", "sira": 1,
     "izinli_sonraki_durumlar": "UYARI|KAPALI", "kategori": "acik"},
    {"durum_id": "UYARI", "durum_adi": "Uyarı", "sira": 2,
     "izinli_sonraki_durumlar": "IHLAL|IYILESTIRME|IZLEME", "kategori": "acik"},
    {"durum_id": "IHLAL", "durum_adi": "İhlal", "sira": 3,
     "izinli_sonraki_durumlar": "IYILESTIRME|KAPALI", "kategori": "acik"},
    {"durum_id": "IYILESTIRME", "durum_adi": "İyileştirme", "sira": 4,
     "izinli_sonraki_durumlar": "IZLEME|UYARI|KAPALI", "kategori": "acik"},
    {"durum_id": "KAPALI", "durum_adi": "Kapalı", "sira": 5,
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
        f"Tanım: {baslik} | Neden önemli: Covenant erken uyarı kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Covenant sapma ve karar bozulur | Kaynak: Kredi sözleşmesi / TTK",
    )


def kural_cek(kural_id: str) -> str:
    """Oranları tblKurallar'dan çeker; yıl seçimi bke_yilTaban ile (G07: adında rakam yok)."""
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-bke_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(bke_yuvarMax,IFERROR(N(G12),0)))"


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
      "Kredi bakiyesi, faiz, FAVÖK, net borç, cari oran ve DSCR girdilerini girin; "
      "covenant sapması, erken uyarı skoru ve UYGUN / UYARI / İHLAL kararını üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Çok yıllı kural tablosu (2025/2026/2027) ile DSCR, borç/FAVÖK, cari eşikleri",
        "Covenant sapma, erken uyarı skoru ve faiz yükü yan yana",
        "İZLEME→UYARI→İHLAL→İYİLEŞTİRME/KAPALI durum makinesi (A3)",
        "Baz / iyimser / kötümser + M-SEN (DSCR eşiği +0,05) senaryoları",
        "Altın vakalar ve KANIT_RAPORU (A4 imza) + belirsizlik: covenant eşik yorumu A/B",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Hazine — DSCR ve borç/FAVÖK sapmasını bankaya raporlamak",
        "CFO — UYGUN / UYARI / İHLAL kararını kredi komitesine yazmak",
        "Banka / denetçi — imzalı kanıt ve durum zincirini dosyalamak",
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
        ("Adım 1 — Girdiler", "GIRDI'de kredi bakiyesi, faiz, FAVÖK, net borç, cari, DSCR ve yorum A/B girin."),
        ("Adım 2 — Karar", "PANO'da sapma, erken uyarı skoru ve UYGUN / UYARI / İHLAL rozetini izleyin."),
        ("Adım 3 — Süreç", "AKIS'te İZLEME→UYARI→İHLAL adımlarını işleyin; KANIT_RAPORU yazdırın."),
    ]
    for i, (bas, acik) in enumerate(adimlar, 5):
        h(ws, i * 2, 1, bas, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i * 2 + 1, 1, acik, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i * 2 + 1, start_column=1, end_row=i * 2 + 1, end_column=20)
    h(ws, 14, 1, "Sık yapılan hatalar", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate([
        "FAVÖK'ü sıfır veya negatif bırakmak (borç/FAVÖK ve faiz yükü bozulur)",
        "DSCR'yi yüzde yerine oran yazmamak (1,45 = 1,45x)",
        "Cari oranı yüzde olarak 125 yazmak (doğrusu 1,25)",
        "AKIS'te kaynak kart kimliğini boş bırakmak (YETİM riski)",
    ], 15):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
    genislik(ws, {"A": 80})

def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Şirket ve covenant girdileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Eşikler KURALLAR'dan; sapma ve karar formülle üretilir.",
      yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Şirket Adı", DEMO["sirket_adi"], None, "sirket_adi",
         "Şirket unvanını yazın."),
        (7, "Dönem", DEMO["donem"], None, "donem",
         "Covenant dönemini yazın (ör. 2026-Q2)."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili",
         "2025, 2026 veya 2027 seçin."),
        (9, "Kredi Bakiyesi [₺]", DEMO["kredi_bakiyesi"], TL, "kredi_bakiyesi",
         "Toplam kredi bakiyesini TL girin."),
        (10, "Faiz Oranı", DEMO["faiz_orani"], YÜZDE, "faiz_orani",
         "Yıllık faiz oranını ondalık girin (0,28 = %28)."),
        (11, "Vade [ay]", DEMO["vade_ay"], CATI, "vade_ay",
         "Kalan vade ayını girin."),
        (12, "FAVÖK [₺]", DEMO["favok"], TL, "favok",
         "Dönem FAVÖK tutarını TL girin."),
        (13, "Net Borç [₺]", DEMO["net_borc"], TL, "net_borc",
         "Net borç tutarını TL girin."),
        (14, "Cari Oran", DEMO["cari_oran"], CATI, "cari_oran",
         "Cari oranı sayı olarak girin (ör. 1,25)."),
        (15, "DSCR Gerçekleşen", DEMO["dscr_gerceklesen"], CATI, "dscr_gerceklesen",
         "Gerçekleşen DSCR değerini girin (ör. 1,45)."),
        (16, "Yorum Modu (A/B)", DEMO["yorum_modu"], None, "yorum_modu",
         "Covenant eşik yorumu A (standart) veya B (sıkı +0,05 DSCR) seçin."),
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
        elif _ad == "vade_ay":
            birim = "ay"
        elif _ad in ("cari_oran", "dscr_gerceklesen"):
            birim = "oran"
        else:
            birim = "metin"
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim, yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 18, 1, "DSCR min (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 2, kural_cek("BKE-001"), sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Borç/FAVÖK max (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2, kural_cek("BKE-002"), sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Erken uyarı skoru (ayna)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=IFERROR(bke_erkenUyariSkor,0)", sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 21, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B16)=0,bke_girisBeklenen=0),0,ROUND(COUNTA(B6:B16)/bke_girisBeklenen*100,0)),0)',
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
        ("kredi_bakiyesi", DEMO["kredi_bakiyesi"], "TL", "Evet", "GIRDI"),
        ("faiz_orani", DEMO["faiz_orani"], "oran", "Evet", "GIRDI"),
        ("vade_ay", DEMO["vade_ay"], "ay", "Evet", "GIRDI"),
        ("favok", DEMO["favok"], "TL", "Evet", "GIRDI"),
        ("net_borc", DEMO["net_borc"], "TL", "Evet", "GIRDI"),
        ("cari_oran", DEMO["cari_oran"], "oran", "Evet", "GIRDI"),
        ("dscr_gerceklesen", DEMO["dscr_gerceklesen"], "oran", "Evet", "GIRDI"),
        ("yorum_modu", DEMO["yorum_modu"], "A/B", "Evet", "GIRDI"),
    ]
    tl_alanlar = {"kredi_bakiyesi", "favok", "net_borc"}
    yuzde_alanlar = {"faiz_orani"}
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 25):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in tl_alanlar:
            ws.cell(i, 2).number_format = TL
        elif a in yuzde_alanlar:
            ws.cell(i, 2).number_format = YÜZDE
        elif a in ("hesapYili", "vade_ay", "cari_oran", "dscr_gerceklesen"):
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
        ("B9", "Kredi Bakiyesi", "0 ile 1e12 arasında TL."),
        ("B12", "FAVÖK", "0 ile 1e12 arasında TL."),
        ("B13", "Net Borç", "0 ile 1e12 arasında TL."),
        ("B28", "Kredi Bakiyesi", "Tablo: TL tutar."),
        ("B31", "FAVÖK", "Tablo: TL tutar."),
        ("B32", "Net Borç", "Tablo: TL tutar."),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj=mesaj,
                  hata_baslik="Geçersiz tutar", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000000")
    for aralik, baslik in [("B10", "Faiz"), ("B29", "Faiz tablo")]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj="0-1 arası oran.",
                  hata_baslik="Geçersiz oran", hata_mesaj="0 ile 1 arası.",
                  isaret="between", f2="1")
    for aralik in ("B11", "B30"):
        dogrulama(ws, "whole", "0", aralik,
                  baslik="Vade Ay", mesaj="0-360 ay.",
                  hata_baslik="Geçersiz", hata_mesaj="0-360.",
                  isaret="between", f2="360")
    for aralik, baslik in [("B14", "Cari"), ("B15", "DSCR"), ("B33", "Cari tablo"), ("B34", "DSCR tablo")]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj="0-20 arası oran.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ile 20 arası.",
                  isaret="between", f2="20")

    giris_hucreleri(ws, 6, 16, [2])
    giris_hucreleri(ws, 25, 35, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 40, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1024, "Sarı alanlar giriş; eşik ve sapma formülleri kilitlidir.")

def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Covenant kural tablosu (yıl yan yana)", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Eşikler MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["BKE-001", "Kredi sözleşmesi / DSCR", "dscr_min", "DSCR ≥ eşik", "dscr_min",
         1.20, 1.20, 1.20, "01.01.2025", "Minimum DSCR eşiği"],
        ["BKE-002", "Kredi sözleşmesi / kaldıraç", "borc_favok_max", "net borç/FAVÖK ≤ max", "borc_favok_max",
         3.50, 3.50, 3.50, "01.01.2025", "Maksimum borç/FAVÖK"],
        ["BKE-003", "Kredi sözleşmesi / likidite", "cari_min", "cari oran ≥ eşik", "cari_min",
         1.00, 1.00, 1.00, "01.01.2025", "Minimum cari oran"],
        ["BKE-004", "İç politika", "erken_uyari_band", "eşik × (1±band)", "erken_uyari_band",
         0.10, 0.10, 0.10, "01.01.2025", "Erken uyarı sapma bandı"],
        ["BKE-005", "İç politika", "faiz_yuku_max", "faiz×bakiye/FAVÖK ≤ max", "faiz_yuku_max",
         0.40, 0.40, 0.40, "01.01.2025", "Maksimum faiz yükü oranı"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = None
            if k in (6, 7, 8):
                sayi = YÜZDE if s[0] in ("BKE-004", "BKE-005") else CATI
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-bke_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "DSCR min (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("BKE-001"), sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Borç/FAVÖK max (aktif yıl)", yazi="333333")
    h(ws, 16, 2, kural_cek("BKE-002"), sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"BKE-001=",TEXT(B15,"0.00"),"BKE-002=",TEXT(B16,"0.00"),'
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
    h(ws, 4, 1, "Eşikler tblKurallar'dan çekilir; sabit oran yoktur.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")

    basliklar = [*MOTOR_BASLIK, "deger"]
    baslik_satiri(ws, 6, basliklar)

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif mevzuat yılı", "BKE-001")
    ekle("DSCR min eşiği", km("BKE-001"), "oran", "Minimum DSCR", "BKE-001")
    ekle("Borç/FAVÖK max", km("BKE-002"), "oran", "Maksimum kaldıraç", "BKE-002")
    ekle("Cari min eşiği", km("BKE-003"), "oran", "Minimum cari oran", "BKE-003")
    ekle("Erken uyarı bandı", km("BKE-004"), "oran", "Uyarı sapma bandı", "BKE-004")
    ekle("Faiz yükü max", km("BKE-005"), "oran", "Maksimum faiz yükü", "BKE-005")
    ekle("Kredi bakiyesi", "=GIRDI!B9", "TL", "Kredi bakiyesi", "BKE-005")
    ekle("Faiz oranı", "=GIRDI!B10", "oran", "Yıllık faiz", "BKE-005")
    ekle("Vade ay", "=GIRDI!B11", "adet", "Kalan vade", "BKE-001")
    ekle("FAVÖK", "=GIRDI!B12", "TL", "Dönem FAVÖK", "BKE-002")
    ekle("Net borç", "=GIRDI!B13", "TL", "Net borç", "BKE-002")
    ekle("Cari oran girdi", "=GIRDI!B14", "oran", "Cari oran", "BKE-003")
    ekle("DSCR gerçekleşen", "=GIRDI!B15", "oran", "Gerçekleşen DSCR", "BKE-001")
    ekle("DSCR eşik efektif",
         '=IF(GIRDI!B16="B",G8+bke_esikArtis,G8)',
         "oran", "Yorum A/B DSCR eşiği", "BKE-001")
    ekle("Borç/FAVÖK gerçekleşen", "=IF(G16=0,0,G17/G16)", "oran", "Net borç / FAVÖK", "BKE-002")
    ekle("Faiz yükü", "=IF(G16=0,0,G13*G14/G16)", "oran", "Faiz×bakiye/FAVÖK", "BKE-005")
    ekle("DSCR sapma", "=G19-G20", "oran", "Gerçekleşen − eşik", "BKE-001")
    ekle("Borç/FAVÖK sapma", "=G21-G9", "oran", "Gerçekleşen − max (negatif iyi)", "BKE-002")
    ekle("Cari sapma", "=G18-G10", "oran", "Gerçekleşen − min", "BKE-003")
    ekle("Faiz yükü sapma", "=G22-G12", "oran", "Gerçekleşen − max", "BKE-005")
    ekle("DSCR uyarı eşiği", "=G20*(1+G11)", "oran", "Erken uyarı alt sınır", "BKE-004")
    ekle("Borç/FAVÖK uyarı", "=G9*(1-G11)", "oran", "Erken uyarı üst sınır", "BKE-004")
    ekle("Cari uyarı eşiği", "=G10*(1+G11)", "oran", "Erken uyarı alt sınır", "BKE-004")
    ekle("İhlal bayrağı",
         '=IF(OR(G19<G20,G21>G9,G18<G10,G22>G12),1,0)',
         "bayrak", "Eşik aşımı", "BKE-001")
    ekle("Uyarı bayrağı",
         '=IF(G30=1,0,IF(OR(G19<G27,G21>G28,G18<G29),1,0))',
         "bayrak", "Erken uyarı bandı", "BKE-004")
    ekle("Erken uyarı skoru",
         '=IFERROR(ROUND(MAX(0,MIN(bke_skorTavan,'
         'IF(G30=1,bke_skorIhlal,IF(G31=1,bke_skorUyari,'
         'bke_skorTaban+MAX(0,(G27-G19)/MAX(G11,bke_paydaMin)*bke_agirlikDscr)'
         '+MAX(0,(G21-G28)/MAX(G9,bke_paydaMin)*bke_agirlikKaldirac)'
         '+MAX(0,(G29-G18)/MAX(G11,bke_paydaMin)*bke_agirlikCari))))),0),0)',
         "puan", "0-100 erken uyarı", "BKE-004")
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(G30=1,"IHLAL",IF(G31=1,"UYARI","UYGUN")))',
         "kod", "Karar kodu", "BKE-001")
    ekle("Karar metni",
         '=IF(G33="VERI_YOK","VERİ YOK",'
         'IF(G33="IHLAL","İHLAL",'
         'IF(G33="UYARI","UYARI","UYGUN")))',
         "metin", "Kullanıcıya karar", "BKE-001")
    ekle("Gerekçe",
         '=IF(G33="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G33="IHLAL","En az bir covenant eşiği aşıldı; banka bildirimi ve iyileştirme planı gerekir.",'
         'IF(G33="UYARI","Eşikler içinde ama erken uyarı bandında; izlemeyi sıkılaştırın.",'
         '"Tüm covenant metrikleri eşik üzerinde; izlemeye devam.")))',
         "metin", "Gerekçe cümlesi", "BKE-001")
    ekle("DSCR / eşik oranı", "=IF(G20=0,0,G19/G20)", "oran", "DSCR kapsama", "BKE-001")
    ekle("Borç-FAVÖK / max", "=IF(G9=0,0,G21/G9)", "oran", "Kaldıraç yoğunluğu", "BKE-002")
    ekle("Faiz yükü / max", "=IF(G12=0,0,G22/G12)", "oran", "Faiz yoğunluğu", "BKE-005")
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B16)=0,bke_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B16)/bke_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "BKE-001")
    ekle("Risk skoru",
         '=IF(G33="VERI_YOK",0,IF(G33="UYGUN",bke_riskDusuk,'
         'IF(G33="UYARI",bke_riskOrta,bke_riskYuksek)))',
         "puan", "Covenant riski", "BKE-004")
    ekle("Senaryo iyimser skor", "=MAX(0,G32-bke_senaryoIyi)", "puan", "İyimser skor", "BKE-004")
    ekle("Senaryo kötümser skor", "=MIN(100,G32+bke_senaryoKotu)", "puan", "Kötümser skor", "BKE-004")
    ekle("İyimser DSCR", "=G19*bke_dscrIyi", "oran", "İyimser DSCR", "BKE-001")
    ekle("Kötümser DSCR", "=G19*bke_dscrKotu", "oran", "Kötümser DSCR", "BKE-001")
    ekle("M-SEN DSCR eşik+", "=G20+bke_esikArtis", "oran", "DSCR eşik + delta", "BKE-001")
    ekle("M-SEN ihlal",
         '=IF(OR(G19<G45,G21>G9,G18<G10,G22>G12),1,0)',
         "bayrak", "M-SEN ihlal", "BKE-001")
    ekle("M-SEN skor farkı", "=IF(G46=1,G32+15,G32)-G32", "puan", "M-SEN skor etkisi", "BKE-001")
    ekle("Tahmin üst skor", "=MIN(100,G32*bke_tahminUst)", "puan", "Tahmin üst", "BKE-004")
    ekle("Tahmin alt skor", "=MAX(0,G32*bke_tahminAlt)", "puan", "Tahmin alt", "BKE-004")
    ekle("Tornado: DSCR etkisi",
         "=ABS((G19*bke_tornadoDscr)-G19)/MAX(G20,0.01)*100",
         "puan", "DSCR ± etki", "BKE-001")
    ekle("Tornado: FAVÖK etkisi",
         "=ABS(IF(G16*bke_tornadoFavok=0,0,G17/(G16*bke_tornadoFavok))-G21)/MAX(G9,0.01)*100",
         "puan", "FAVÖK ± etki", "BKE-002")
    ekle("Tornado: faiz etkisi",
         "=ABS(IF(G16=0,0,G13*G14*bke_tornadoFaiz/G16)-G22)/MAX(G12,0.01)*100",
         "puan", "Faiz ± etki", "BKE-005")
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"BKE-001→DSCR min","BKE-002→borç/FAVÖK max",'
         '"BKE-003→cari min","BKE-004→erken uyarı","BKE-005→faiz yükü")',
         "metin", "Kanıt atıfları", "BKE-001")
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"DSCR=",TEXT(G19,"0.00"),"B/F=",TEXT(G21,"0.00"),'
         '"Cari=",TEXT(G18,"0.00"),"Skor=",TEXT(G32,"0"),"Karar=",G34)',
         "metin", "Rapor özeti", "BKE-001")
    ekle("Erken uyarı ayna", "=G32", "puan", "Skor ayna", "BKE-004")
    ekle("Karar ayna", "=G34", "metin", "Karar ayna", "BKE-001")

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
      "Motor adımları kilitlidir; eşikler yalnızca KURALLAR'dan gelir.", yazi=GRİ, kaydir=True)
    return len(adimlar)

def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve duyarlılık", kalin=True, boyut=13, yazi=KOYU_LACIVERT)

    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "DSCR Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Erken Uyarı Skoru", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "DSCR", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(bke_dscrIyi+N(GIRDI!B9)*bke_sifirCarpan,0)",
         "=IFERROR(MOTOR!G41,0)",
         "=IFERROR(MOTOR!G43,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D6<MOTOR!G20,"İHLAL",IF(D6<MOTOR!G27,"UYARI","UYGUN")))'),
        ("Baz", "=IFERROR(bke_bazCarpan+N(GIRDI!B9)*bke_sifirCarpan,bke_bazCarpan)",
         "=IFERROR(bke_erkenUyariSkor,0)",
         "=IFERROR(MOTOR!G19,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(bke_kararMetni,"-"))'),
        ("Kötümser", "=IFERROR(bke_dscrKotu+N(GIRDI!B9)*bke_sifirCarpan,0)",
         "=IFERROR(MOTOR!G42,0)",
         "=IFERROR(MOTOR!G44,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(D8<MOTOR!G20,"İHLAL",IF(D8<MOTOR!G27,"UYARI","UYGUN")))'),
        ("M-SEN eşik+", "=IFERROR(bke_bazCarpan+N(GIRDI!B9)*bke_sifirCarpan,bke_bazCarpan)",
         "=IFERROR(bke_erkenUyariSkor+MOTOR!G47,0)",
         "=IFERROR(MOTOR!G19,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",'
         'IF(MOTOR!G46=1,"İHLAL",IFERROR(bke_kararMetni,"-")))'),
    ]
    for i, (ad, carp, skor, dscr, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, skor, sayi=CATI, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, dscr, sayi=CATI, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo skor bant genişliği (kötümser−iyimser)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(C8-C6,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorum", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser skor ",TEXT(C6,"0")," · Baz ",TEXT(C7,"0"),'
      '" · Kötümser ",TEXT(C8,"0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:E12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [puan]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    tornado = [
        ("DSCR ±", "=IFERROR(MOTOR!G50,0)"),
        ("FAVÖK ±", "=IFERROR(MOTOR!G51,0)"),
        ("Faiz oranı ±", "=IFERROR(MOTOR!G52,0)"),
    ]
    for i, (ad, form) in enumerate(tornado, 16):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, sayi=CATI, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, i - 15, sayi=CATI, yazi=GRİ)
    h(ws, 19, 1, "Tornado zirve sırası", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(INDEX(A16:A18,MATCH(MAX(B16:B18),B16:B18,0)),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Tornado yorum", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En etkili değişken: ",B19," etki ",TEXT(MAX(B16:B18),"0")," puan")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN kural değişimi (DSCR eşik +0,05)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=IFERROR(MOTOR!G47,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorum", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"DSCR eşiği +0,05 senaryosunda skor farkı ",TEXT(B22,"0")," puan")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı (erken uyarı skoru)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 1, "Alt", yazi="333333")
    h(ws, 26, 2, "=IFERROR(MOTOR!G49,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 27, 1, "Üst", yazi="333333")
    h(ws, 27, 2, "=IFERROR(MOTOR!G48,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 28, 1, "Orta (ortalama)", yazi="333333")
    h(ws, 28, 2,
      "=IFERROR(IF(COUNT(C6:C8)=0,0,AVERAGE(C6:C8)),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 32, 1, "Tahmin aralık genişliği", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 32, 2, "=IFERROR(B27-B26,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Tahmin yorum", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Erken uyarı skor bandı ",TEXT(B26,"0")," – ",TEXT(B27,"0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 5, 7, "SkorDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([10, 20, 45, 35], 6):
        h(ws, i, 7, v, sayi=CATI, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([25, 18, 12], 16):
        h(ws, i, 5, v, sayi=CATI, yazi=GRİ)
    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([10, 20, 45, 35], 28):
        h(ws, i, 4, v, sayi=CATI, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo Erken Uyarı Skoru"
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
    g3.title = "Senaryo Skor Trendi"
    g3.add_data(Reference(ws, min_col=4, min_row=27, max_row=31), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=28, max_row=31))
    g3.height, g3.width = 8, 12
    ws.add_chart(g3, "F32")

    genislik(ws, {"A": 36, "B": 18, "C": 18, "D": 18, "E": 22, "G": 14})

def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", "2E75B6", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Covenant altın vakaları", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:H3")
    baslik_satiri(ws, 5, VAKA_BASLIK)

    vakalar_data = [
        ("V-001", "DSCR min eşiği okuma",
         "eşik=BKE-001 (2026)",
         1.20,
         "=IFERROR(INDEX(tblKurallar[deger_2026],MATCH(\"BKE-001\",tblKurallar[kural_id],0)),0)",
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=bke_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "DSCR min"),
        ("V-002", "Borç/FAVÖK gerçekleşen",
         "net_borç=GIRDI!B13; FAVÖK=GIRDI!B12",
         1.875,
         "=IFERROR(ROUND(GIRDI!B13/GIRDI!B12,3),0)",
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=bke_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Kaldıraç"),
        ("V-003", "Faiz yükü = bakiye×faiz/FAVÖK",
         "bakiye=GIRDI!B9; faiz=GIRDI!B10; FAVÖK=GIRDI!B12",
         0.35,
         "=IFERROR(ROUND(GIRDI!B9*GIRDI!B10/GIRDI!B12,4),0)",
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=bke_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Faiz yükü"),
        ("V-004", "Erken uyarı bandı × DSCR min",
         "min=BKE-001; band=BKE-004",
         1.32,
         "=IFERROR(ROUND(INDEX(tblKurallar[deger_2026],MATCH(\"BKE-001\",tblKurallar[kural_id],0))"
         "*(1+INDEX(tblKurallar[deger_2026],MATCH(\"BKE-004\",tblKurallar[kural_id],0))),4),0)",
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=bke_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Uyarı eşiği"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v,
              yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = CATI

    formuller = {
        "fark": "=IF(tblVakalar[[#This Row],[vaka_id]]=\"\",\"\",IFERROR(tblVakalar[[#This Row],[beklenen_sonuc]]-tblVakalar[[#This Row],[hesaplanan]],0))",
        "durum": '=IF(tblVakalar[[#This Row],[vaka_id]]="","",IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=bke_tolerans,"TUTARLI","KIRIK"),"KIRIK"))',
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
    for i, (b, he) in enumerate([(1.20, 1.20), (1.875, 1.875), (0.35, 0.35), (1.32, 1.32)], 6):
        h(ws, i, 10, b, sayi=CATI, yazi=GRİ)
        h(ws, i, 11, he, sayi=CATI, yazi=GRİ)
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
    h(ws, 3, 1, "KANIT RAPORU — Banka Kredi + Covenant Erken Uyarı", kalin=True, boyut=14, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:F3")
    h(ws, 4, 1, "Rapor tarihi", yazi="333333")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 5, 1, "Şirket", yazi="333333")
    h(ws, 5, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Dönem / yıl", yazi="333333")
    h(ws, 6, 2, '=CONCATENATE(GIRDI!B7," / ",hesapYili)', yazi=DIS_REF_YEŞIL)

    h(ws, 8, 1, "Girdi özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, (et, form, sayi) in enumerate([
        ("Kredi bakiyesi", "=GIRDI!B9", TL),
        ("Faiz oranı", "=GIRDI!B10", YÜZDE),
        ("FAVÖK", "=GIRDI!B12", TL),
        ("Net borç", "=GIRDI!B13", TL),
        ("Cari oran", "=GIRDI!B14", CATI),
        ("DSCR gerçekleşen", "=GIRDI!B15", CATI),
    ], 9):
        h(ws, i, 1, et, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)

    h(ws, 16, 1, "Hesap zinciri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 17, 1, "Borç/FAVÖK", yazi="333333")
    h(ws, 17, 2, "=MOTOR!G21", sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 1, "Faiz yükü", yazi="333333")
    h(ws, 18, 2, "=MOTOR!G22", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Erken uyarı skoru", yazi="333333")
    h(ws, 19, 2, "=bke_erkenUyariSkor", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "DSCR", yazi="333333")
    h(ws, 20, 2, "=MOTOR!G19", sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 21, 1, "KARAR", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 2, "=bke_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 22, 1, "Gerekçe", yazi="333333")
    h(ws, 22, 2, "=bke_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B22:F22")

    h(ws, 24, 1, "Madde atıfları", yazi="333333")
    h(ws, 24, 2, "=bke_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B24:F24")
    h(ws, 25, 1, "Süreç özeti", yazi="333333")
    h(ws, 25, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"Skor=",TEXT(bke_erkenUyariSkor,"0"),'
      '"Açık kuyruk=",TEXT(COUNTIFS(tblAkis[Durum],"IZLEME")+COUNTIFS(tblAkis[Durum],"UYARI")+COUNTIFS(tblAkis[Durum],"IHLAL")+COUNTIFS(tblAkis[Durum],"IYILESTIRME"),"0")),"-")',
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
    h(ws, 30, 1, "Uyarı: Bu çıktı karar destek niteliğindedir; banka sözleşmesinin yerine geçmez.",
      yazi=KRITIK, boyut=9, kaydir=True)
    ws.merge_cells("A30:F30")
    h(ws, 31, 1, f"Sürüm {SURUM} | Parmak izi: ", yazi=GRİ, boyut=9)
    h(ws, 31, 2, "=bke_parmakIzi", yazi=GRİ, boyut=9)
    baski_hazirla(ws, "A1:F32", f"{URUN_AD} · KANIT")
    genislik(ws, {"A": 28, "B": 55})

def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim paneli — Banka Kredi + Covenant Erken Uyarı", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpiler = [
        (2, "DSCR", "=IFERROR(MOTOR!G19,0)", CATI),
        (3, "BorcFAVOK", "=IFERROR(MOTOR!G21,0)", CATI),
        (4, "Cari", "=IFERROR(MOTOR!G18,0)", CATI),
        (5, "FaizYuku", "=IFERROR(MOTOR!G22,0)", YÜZDE),
        (6, "Güven", "=IFERROR(MOTOR!G39,0)", CATI),
        (7, "Risk", "=IFERROR(MOTOR!G40,0)", CATI),
        (8, "DSCR.Min", "=IFERROR(MOTOR!G20,0)", CATI),
        (9, "UyariSkor", "=IFERROR(bke_erkenUyariSkor,0)", CATI),
        (10, "DSCR/Esik", "=IFERROR(MOTOR!G36,0)", YÜZDE),
        (11, "M-SEN", "=IFERROR(bke_kuralDegisimSenaryo,0)", CATI),
        (12, "Tahmin", "=IFERROR(SENARYO!B32,0)", CATI),
        (13, "Vaka", "=IFERROR(bke_vakaDurum,\"-\")", None),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=9, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=bke_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=bke_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 8, 1, "Açık süreç adedi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 8, 2,
      '=COUNTIFS(tblAkis[Durum],"IZLEME")+COUNTIFS(tblAkis[Durum],"UYARI")+COUNTIFS(tblAkis[Durum],"IHLAL")+COUNTIFS(tblAkis[Durum],"IYILESTIRME")',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 8, 3, "Açık tutar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 8, 4,
      '=SUMIF(tblAkis[Durum],"IZLEME",tblAkis[Tutar])+SUMIF(tblAkis[Durum],"UYARI",tblAkis[Tutar])+SUMIF(tblAkis[Durum],"IHLAL",tblAkis[Tutar])+SUMIF(tblAkis[Durum],"IYILESTIRME",tblAkis[Tutar])',
      sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 10, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "DSCR / eşik oranı",
         "=IFERROR(MOTOR!G36,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"DSCR/eşik ",TEXT(C11,"0.0%"))'),
        ("T2", "Borç-FAVÖK / max",
         "=IFERROR(MOTOR!G37,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kaldıraç yoğunluğu ",TEXT(C12,"0.0%"))'),
        ("O1", "Senaryo skor sapması",
         "=IFERROR(IF(COUNT(SENARYO!C6:C9)<2,0,STDEV.P(SENARYO!C6:C9)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo sapması ",TEXT(C13,"0.0")," puan")'),
        ("O2", "Tornado max etki",
         "=IFERROR(bke_duyarlilikSira,\"-\")",
         "=IFERROR(bke_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo bant",
         "=IFERROR(bke_senaryoKarsilastirma,0)",
         "=IFERROR(bke_yorumSenaryo,\"-\")"),
        ("O6", "Faiz yükü / max",
         "=IFERROR(MOTOR!G38,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Faiz yoğunluğu ",TEXT(C16,"0.0%"))'),
        ("O8", "Kalite skoru",
         "=IFERROR(KONTROLLER!B12,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalite ",TEXT(C17,"0")," puan")'),
        ("M", "M-SEN eşik farkı",
         "=IFERROR(bke_kuralDegisimSenaryo,0)",
         "=IFERROR(bke_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin aralık",
         "=IFERROR(bke_tahminAralik,0)",
         "=IFERROR(bke_yorumTahmin,\"-\")"),
        ("I2", "Senaryo skor P90",
         "=IFERROR(IF(COUNT(SENARYO!C6:C9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!C6:C9,bke_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"P90 erken uyarı ",TEXT(C20,"0")," puan")'),
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
            ws.cell(r, 3).number_format = CATI

    h(ws, 23, 1, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("DSCR", 1.45),
        ("Borç/FAVÖK", 1.875),
        ("Cari oran", 1.25),
        ("Faiz yükü", 0.35),
        ("Uyarı skoru", 20),
    ], 24):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=CATI, yazi=GRİ)

    h(ws, 23, 4, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 5, "Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("DSCR min", 1.20),
        ("Borç/FAVÖK max", 3.50),
        ("İyimser skor", 10),
        ("Kötümser skor", 45),
        ("M-SEN fark", 0),
    ], 24):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=CATI, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Covenant Metrikleri"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=28), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=28))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G9")

    g2 = BarChart()
    g2.type = "col"
    g2.title = "Eşik / Senaryo"
    g2.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "G23")

    g3 = LineChart()
    g3.title = "Skor Çizgisi"
    g3.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g3.height, g3.width = 7, 12
    ws.add_chart(g3, "P9")

    h(ws, 30, 1, "Metrik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 2, "Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("DSCR", 1.45),
        ("Uyarı skoru", 20),
        ("Borç/FAVÖK", 1.875),
    ], 31):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=CATI, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "KPI DSCR / Skor / Kaldıraç"
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
        ("Negatif kredi bakiyesi?",
         '=IF(N(GIRDI!B9)<0,"NEGATİF","NORMAL")'),
        ("FAVÖK ≤ 0?",
         '=IF(N(GIRDI!B12)<=0,"AŞIYOR","NORMAL")'),
        ("Faiz oranı > 1?",
         '=IF(N(GIRDI!B10)>1,"AŞIYOR","NORMAL")'),
        ("Vaka tutarlı mı?",
         '=IFERROR(IF(bke_vakaDurum="TUTARLI","TUTARLI","KIRIK"),"KIRIK")'),
        ("Yetim kayıt var mı?",
         '=IF(COUNTIF(tblAkis[YetimBayrak],"YETİM")>0,"HATALI","TEMİZ")'),
        ("Geçersiz geçiş var mı?",
         '=IF(COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")>0,"HATALI","TEMİZ")'),
        ("Zincir kırık var mı?",
         '=IF(COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")>0,"HATALI","TEMİZ")'),
        ("Karar üretildi mi?",
         '=IF(OR(bke_kararMetni="",bke_kararMetni="VERİ YOK"),"EKSİK","TAM")'),
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
    h(ws, 15, 1, "Erken uyarı skoru", yazi="333333")
    h(ws, 15, 2, "=IFERROR(bke_erkenUyariSkor,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "DSCR", yazi="333333")
    h(ws, 16, 2, "=IFERROR(MOTOR!G19,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Skor (ayna)", yazi="333333")
    h(ws, 17, 2, "=IFERROR(bke_erkenUyariSkor,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "DSCR (ayna)", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G19,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 36, "B": 18})

def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "548235", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Örnek senaryo satırları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    basliklar = ["Senaryo", "KrediBakiye", "FaizOran", "VadeAy", "Favok",
                 "NetBorc", "CariOran", "Dscr", "Yil", "Kontrol"]
    baslik_satiri(ws, 5, basliklar)
    satirlar = [
        ("Demo ana", 10_000_000, 0.28, 36, 8_000_000, 15_000_000, 1.25, 1.45, 2026),
        ("Uyari sinir", 10_000_000, 0.28, 36, 8_000_000, 15_000_000, 1.10, 1.25, 2026),
        ("Ihlal DSCR", 10_000_000, 0.28, 36, 8_000_000, 15_000_000, 1.25, 0.95, 2026),
        ("Yuksek kaldirac", 12_000_000, 0.30, 48, 4_000_000, 18_000_000, 1.05, 1.50, 2026),
        ("2025 oran", 10_000_000, 0.28, 36, 8_000_000, 15_000_000, 1.25, 1.45, 2025),
    ]
    for i, row in enumerate(satirlar, 6):
        for k, v in enumerate(row, 1):
            sayi = None
            if k in (2, 5, 6):
                sayi = TL
            elif k == 3:
                sayi = YÜZDE
            elif k in (4, 7, 8, 9):
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
            'IF(OR(tblOrnek[[#This Row],[KrediBakiye]]="",tblOrnek[[#This Row],[Yil]]=""),"EKSİK","TAM"))'
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

    h(ws, 3, 7, "Kredi Türü", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 7, "Tur", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate(["ISLETME", "YATIRIM"], 6):
        h(ws, i, 7, v, zemin=GIRIS_SARI)
        ws.cell(i, 7).protection = Protection(locked=False)
        _kim(ws, f"G{i}", "Kredi Türü", "ISLETME veya YATIRIM")

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
    alt_bant(ws, 3, "Kredi / covenant kartları — her kaydın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, '=CONCATENATE("BKE-",TEXT(COUNTA(tblKartlar[KartId])+1,"00000"))', kalin=True)
    _kim(ws, "B4", "Sonraki Kimlik", "Yeni kartta kullanın")

    sutunlar = ["KartId", "Unvan", "Donem", "KrediTur", "LimitTl",
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
        ("BKE-00001", "Örnek Sanayi A.Ş.", "2026-Q2", "ISLETME"),
        ("BKE-00002", "Demo Metal A.Ş.", "2026-Q1", "YATIRIM"),
        ("BKE-00003", "Anadolu Tekstil A.Ş.", "2025-Q4", "ISLETME"),
    ]
    for i, (kid, unvan, donem, tur) in enumerate(demo_kartlar):
        r = KART_ILK + i
        _sari(ws, r, 1, kid, baslik="Kart Kimliği", mesaj="BKE-#####")
        _sari(ws, r, 2, unvan, baslik="Unvan", mesaj="Firma unvanı")
        _sari(ws, r, 3, donem, baslik="Dönem", mesaj="Covenant dönemi")
        _sari(ws, r, 4, tur, baslik="Kredi Türü", mesaj="ISLETME/YATIRIM")
        _sari(ws, r, 5, 10_000_000 + i * 2_000_000, sayi=TL, baslik="Limit", mesaj="Limit TL")
        _sari(ws, r, 6, "Normal", baslik="Risk", mesaj="Risk notu")
        _sari(ws, r, 7, "EVET", baslik="Aktif", mesaj="EVET/HAYIR")
    for r in range(KART_ILK + 3, KART_SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None, sayi=TL if c == 5 else None,
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblKartlar", f"A{KART_HDR}:I{KART_SON}", sutunlar, formuller=form)
    dogrulama(ws, "list", "ListeKrediTur", f"D{KART_ILK}:D{KART_SON}",
              baslik="Kredi Türü", mesaj="ISLETME veya YATIRIM",
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
             "Süreç satırları — İZLEME→UYARI→İHLAL + yetim + geçersiz geçiş + zincir",
             son_kolon=18)
    sutunlar = [
        "IslemId", "KaynakKartId", "Tarih", "OncekiDurum", "Durum",
        "OncekiSkor", "YeniSkor", "LimitTl", "Tutar", "Aciklama",
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
        "tblAkis[[#This Row],[OncekiSkor]]",
        "tblAkis[[#This Row],[YeniSkor]]",
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
        {"id": "COV-00001", "kart": "BKE-00001", "tarih": date(2026, 4, 10),
         "onceki": "", "durum": "IZLEME", "osk": 20, "ysk": 20,
         "lim": 10_000_000, "tutar": 10_000_000, "acik": "Q2 izleme"},
        {"id": "COV-00002", "kart": "BKE-00001", "tarih": date(2026, 5, 1),
         "onceki": "IZLEME", "durum": "UYARI", "osk": 20, "ysk": 55,
         "lim": 10_000_000, "tutar": 10_000_000, "acik": "Erken uyarı"},
        {"id": "COV-00003", "kart": "BKE-00002", "tarih": date(2026, 3, 15),
         "onceki": "UYARI", "durum": "IHLAL", "osk": 55, "ysk": 90,
         "lim": 12_000_000, "tutar": 12_000_000, "acik": "DSCR ihlali"},
        {"id": "COV-00004", "kart": "BKE-00003", "tarih": date(2026, 2, 20),
         "onceki": "IHLAL", "durum": "IYILESTIRME", "osk": 90, "ysk": 40,
         "lim": 8_000_000, "tutar": 8_000_000, "acik": "İyileştirme planı"},
        {"id": "COV-00005", "kart": "", "tarih": date(2026, 6, 1),
         "onceki": "IZLEME", "durum": "UYARI", "osk": 20, "ysk": 50,
         "lim": 5_000_000, "tutar": 5_000_000, "acik": "Yetim örnek"},
        {"id": "COV-00006", "kart": "BKE-00002", "tarih": date(2026, 6, 5),
         "onceki": "IZLEME", "durum": "IHLAL", "osk": 20, "ysk": 90,
         "lim": 8_000_000, "tutar": 8_000_000, "acik": "Geçersiz geçiş örnek"},
        {"id": "COV-00007", "kart": "BKE-00001", "tarih": date(2026, 5, 20),
         "onceki": "UYARI", "durum": "IYILESTIRME", "osk": 55, "ysk": 5,
         "lim": 10_000_000, "tutar": 10_000_000, "acik": "Zincir sapma örnek"},
    ]
    for i, s in enumerate(ornek):
        r = AKIS_ILK + i
        _sari(ws, r, 1, s["id"], baslik="İşlem", mesaj="COV-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Kart", mesaj="KartId")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Olay tarihi")
        _sari(ws, r, 4, s["onceki"] or None, baslik="Önceki", mesaj="Önceki durum")
        _sari(ws, r, 5, s["durum"], baslik="Durum", mesaj="Durum haritası")
        _sari(ws, r, 6, s["osk"], sayi=CATI, baslik="Önceki Skor", mesaj="Önceki skor")
        _sari(ws, r, 7, s["ysk"], sayi=CATI, baslik="Yeni Skor", mesaj="Yeni skor")
        _sari(ws, r, 8, s["lim"], sayi=TL, baslik="Limit", mesaj="Limit TL")
        _sari(ws, r, 9, s["tutar"], sayi=TL, baslik="Tutar", mesaj="Kuyruk tutarı")
        _sari(ws, r, 10, s["acik"], baslik="Açıklama", mesaj="Kısa açıklama")
    for r in range(AKIS_ILK + len(ornek), AKIS_SON + 1):
        for c in range(1, 11):
            fmt = TARİH if c == 3 else (TL if c in (8, 9) else (CATI if c in (6, 7) else None))
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblAkis", f"A{AKIS_HDR}:P{AKIS_SON}", sutunlar, formuller=form)
    dogrulama(ws, "list", "ListeDurumlar", f"D{AKIS_ILK}:D{AKIS_SON}",
              baslik="Önceki Durum", mesaj="Durum seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"E{AKIS_ILK}:E{AKIS_SON}",
              baslik="Durum", mesaj="Durum seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    for c, ad in [(6, "Önceki Skor"), (7, "Yeni Skor"), (8, "Limit"), (9, "Tutar")]:
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
        ("İzleme", "IZLEME"),
        ("Uyarı", "UYARI"),
        ("İhlal", "IHLAL"),
        ("İyileştirme", "IYILESTIRME"),
    ], 6):
        h(ws, r, 1, ad)
        h(ws, r, 2, f'=COUNTIFS(tblAkis[Durum],"{kod}",tblAkis[Tutar],">0")', sayi=CATI)
        h(ws, r, 3, f'=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"{kod}")', sayi=TL)
    h(ws, 11, 1, "Kapalı — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 1, "Kapalı")
    h(ws, 12, 2, '=COUNTIFS(tblAkis[Durum],"KAPALI")', sayi=CATI)
    h(ws, 12, 3, '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"KAPALI")', sayi=TL)
    h(ws, 14, 1, "Yetim adet")
    h(ws, 14, 2, '=COUNTIF(tblAkis[YetimBayrak],"YETİM")', sayi=CATI)
    h(ws, 15, 1, "Geçersiz geçiş")
    h(ws, 15, 2, '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")', sayi=CATI)
    h(ws, 16, 1, "Zincir kırık")
    h(ws, 16, 2, '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")', sayi=CATI)
    h(ws, 18, 1, "Açık toplam adet", kalin=True)
    h(ws, 18, 2, "=B6+B7+B8+B9", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Açık toplam tutar", kalin=True)
    h(ws, 19, 2, "=C6+C7+C8+C9", sayi=TL, yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 40, "B": 12, "C": 16})

def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Parametreler (tek kaynak) + durum haritası", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = ["anahtar", "deger", "birim", "aciklama", "kaynak", "yururluk_tarihi", "dogrulama_tarihi", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    params = [
        ("hesapYili", DEMO["hesapYili"], "yıl", "Aktif hesap yılı", "Kullanıcı / GIRDI", "01.01.2025", "11.08.2026"),
        ("raporTarihi", RAPOR_TARIH, "tarih", "Rapor tarihi (sabit)", "Üretim", "01.01.2025", "11.08.2026"),
        ("bke_tolerans", 0.01, "oran", "Vaka tutarlılık toleransı", "Uygulama notu", "01.01.2025", "11.08.2026"),
        ("bke_kacirilanEsik", 50, "puan", "PANO uyarı skor eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("bke_senaryoIyi", 10, "puan", "İyimser skor iyileştirme", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("bke_senaryoKotu", 25, "puan", "Kötümser skor bozulma", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("bke_esikArtis", 0.05, "oran", "M-SEN / Yorum B DSCR +0,05", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("bke_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("bke_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("bke_tornadoDscr", 1.05, "çarpan", "Tornado DSCR şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("bke_tornadoFavok", 1.05, "çarpan", "Tornado FAVÖK şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("bke_tornadoFaiz", 1.10, "çarpan", "Tornado faiz şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("bke_dscrIyi", 1.10, "çarpan", "İyimser DSCR çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("bke_dscrKotu", 0.85, "çarpan", "Kötümser DSCR çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("bke_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "11.08.2026"),
        ("bke_parmakIzi", "BKE-PRO-1.0.0", "metin", "Dosya parmak izi etiketi", "Üretim", "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv", "01.01.2025", "11.08.2026"),
        ("bke_girisBeklenen", 11, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("bke_riskDusuk", 15, "puan", "UYGUN risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("bke_riskOrta", 55, "puan", "UYARI risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("bke_riskYuksek", 85, "puan", "İHLAL risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("bke_riskEsikOran", 0.5, "oran", "Yedek risk oranı", "İç politika", "01.01.2025", "11.08.2026"),
        ("bke_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("bke_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("bke_dayanakCarpan", 10, "adet", "Yedek çarpan", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("bke_bazCarpan", 1, "çarpan", "Baz senaryo çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("bke_sifirCarpan", 0, "çarpan", "Nötr çarpan", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("bke_siraBir", 1, "adet", "Tornado sıra 1", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("bke_siraIki", 2, "adet", "Tornado sıra 2", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("bke_siraUc", 3, "adet", "Tornado sıra 3", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("bke_tahminX", 5, "adet", "FORECAST X noktası", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("bke_skorTavan", 100, "puan", "Erken uyarı skor tavanı", "İç politika", "01.01.2025", "11.08.2026"),
        ("bke_skorIhlal", 90, "puan", "İhlal skor seviyesi", "İç politika", "01.01.2025", "11.08.2026"),
        ("bke_skorUyari", 60, "puan", "Uyarı skor seviyesi", "İç politika", "01.01.2025", "11.08.2026"),
        ("bke_skorTaban", 20, "puan", "Uygun taban skor", "İç politika", "01.01.2025", "11.08.2026"),
        ("bke_paydaMin", 0.01, "oran", "Bölme paydası alt sınırı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("bke_agirlikDscr", 30, "puan", "Skor DSCR ağırlığı", "İç politika", "01.01.2025", "11.08.2026"),
        ("bke_agirlikKaldirac", 25, "puan", "Skor kaldıraç ağırlığı", "İç politika", "01.01.2025", "11.08.2026"),
        ("bke_agirlikCari", 20, "puan", "Skor cari ağırlığı", "İç politika", "01.01.2025", "11.08.2026"),
        ("zincirTolerans", 40, "puan", "ZİNCİR KIRIK eşiği (skor)", "A04", "01.01.2025", "11.08.2026"),
        ("bke_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi", "01.01.2025", "11.08.2026"),
        ("bke_kapaliDusum", "kapali kategori kuyruktan düşer", "metin", "A06 kapali", "01.01.2025", "11.08.2026"),
        ("bke_olcekHedef", 20000, "adet", "Ölçek hedefi", "Manda A3", "01.01.2025", "11.08.2026"),
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
    h(ws, 48, 2, f"Sürüm {SURUM} | Şifre koruması: {SIFRE} (formül alanları)", yazi=GRİ, boyut=9, kaydir=True)

def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar kuralları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "VERİ YOK — giriş tablosu boşsa hesap yapılmaz.",
        "UYGUN — DSCR, borç/FAVÖK, cari ve faiz yükü tüm eşikleri sağlar.",
        "UYARI — eşikler içinde ama erken uyarı bandında; izlemeyi sıkılaştırın.",
        "İHLAL — en az bir covenant eşiği aşıldı; banka bildirimi gerekir.",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 10, 1, "Belirsizlik beyanları (M05)", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari(["covenant_esik_yorumu"]), 11):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 13, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2,
      "DSCR eşiği KURALLAR'daki aktif yıl değeriyle aynen kullanılır. "
      "Erken uyarı bandı eşiğin ± yüzdesidir.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B13:J13")
    h(ws, 14, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 2,
      "Yorum B, DSCR minimum eşiğini +0,05 sıkılaştırır (daha muhafazakâr banka yorumu). "
      "Dosya her iki yorumun sonucunu MOTOR'da üretir. Tartışmalı / belirsiz eşik yorumu açıkça beyan edilir.",
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
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("İHLAL",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("UYARI",B6))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B6", FormulaRule(formula=['ISNUMBER(SEARCH("UYGUN",B6))'], font=yesil, fill=y_fill))
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
        ws.conditional_formatting.add(f"{col}9:{col}13", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
        ws.conditional_formatting.add(f"{col}9:{col}13", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("F25:F35", FormulaRule(formula=['F25="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F25:F35", FormulaRule(formula=['F25="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F6:F9", CellIsRule(operator="notEqual", formula=["0"], font=amber))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("C6:C9", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("C6:C9", CellIsRule(operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("İHLAL",E6))'], font=kirmizi))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("UYARI",E6))'], font=amber))
    ws.conditional_formatting.add("E6:E9", FormulaRule(formula=['ISNUMBER(SEARCH("UYGUN",E6))'], font=yesil))
    ws.conditional_formatting.add("B16:B18", CellIsRule(operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add("B5:B11", FormulaRule(formula=['OR(B5="KIRIK",B5="HATALI",B5="NEGATİF",B5="BOŞ",B5="EKSİK",B5="AŞIYOR")'],
                                                        font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B5:B11", FormulaRule(formula=['OR(B5="TEMİZ",B5="DOLU",B5="NORMAL",B5="TUTARLI",B5="TAM")'],
                                                        font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B12", CellIsRule(operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("B12", CellIsRule(operator="lessThan", formula=["50"], font=kirmizi))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("G7:G56", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("F7:F56", FormulaRule(formula=['F7="BKE-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F56", FormulaRule(formula=['F7="BKE-004"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "E", "F"):
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
    for durum, fill in [("IZLEME", a_fill), ("UYARI", a_fill), ("IHLAL", k_fill),
                        ("IYILESTIRME", a_fill), ("KAPALI", y_fill)]:
        ws.conditional_formatting.add(
            f"E{AKIS_ILK}:E{AKIS_SON}",
            CellIsRule(operator="equal", formula=[f'"{durum}"'], fill=fill))

    ws = wb["KUYRUKLAR"]
    ws.conditional_formatting.add("B6:B9", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("C6:C9", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("B14:B16", CellIsRule(operator="greaterThan", formula=["0"], font=kirmizi))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B21", FormulaRule(formula=['ISNUMBER(SEARCH("İHLAL",B21))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B21", FormulaRule(formula=['ISNUMBER(SEARCH("UYARI",B21))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B21", FormulaRule(formula=['ISNUMBER(SEARCH("UYGUN",B21))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B19", CellIsRule(operator="greaterThan", formula=["0"], font=yesil))

def ad_tanimlari(wb):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "zincirTolerans", "AYARLAR!$B$45")
    ad_ekle(wb, "bke_erkenUyariSkor", "MOTOR!$G$55")
    ad_ekle(wb, "bke_kararMetni", "MOTOR!$G$34")
    ad_ekle(wb, "bke_kararGerekce", "MOTOR!$G$35")
    ad_ekle(wb, "bke_maddeAtifMetni", "MOTOR!$G$53")
    ad_ekle(wb, "bke_kanitRaporu", "MOTOR!$G$54")
    ad_ekle(wb, "bke_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "bke_parmakIzi", "AYARLAR!$B$21")

    ad_ekle(wb, "bke_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "bke_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "bke_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "bke_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "bke_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "bke_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "bke_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "bke_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "bke_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "bke_vakaFark", "VAKALAR!$B$13")
    ad_ekle(wb, "bke_sonrakiKimlik", "KARTLAR!$B$4")

    ayar_map = {
        "bke_tolerans": 8,
        "bke_kacirilanEsik": 9,
        "bke_senaryoIyi": 10,
        "bke_senaryoKotu": 11,
        "bke_esikArtis": 12,
        "bke_tahminAlt": 13,
        "bke_tahminUst": 14,
        "bke_tornadoDscr": 15,
        "bke_tornadoFavok": 16,
        "bke_tornadoFaiz": 17,
        "bke_dscrIyi": 18,
        "bke_dscrKotu": 19,
        "bke_yuzdelikOran": 20,
        "bke_girisBeklenen": 23,
        "bke_riskDusuk": 24,
        "bke_riskOrta": 25,
        "bke_riskYuksek": 26,
        "bke_riskEsikOran": 27,
        "bke_yilTaban": 28,
        "bke_yuvarMax": 29,
        "bke_dayanakCarpan": 30,
        "bke_bazCarpan": 31,
        "bke_sifirCarpan": 32,
        "bke_siraBir": 33,
        "bke_siraIki": 34,
        "bke_siraUc": 35,
        "bke_tahminX": 36,
        "bke_skorTavan": 37,
        "bke_skorIhlal": 38,
        "bke_skorUyari": 39,
        "bke_skorTaban": 40,
        "bke_paydaMin": 41,
        "bke_agirlikDscr": 42,
        "bke_agirlikKaldirac": 43,
        "bke_agirlikCari": 44,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "bke_modulT1": (11, "C"), "bke_modulT1Yorum": (11, "D"),
        "bke_modulT2": (12, "C"), "bke_modulT2Yorum": (12, "D"),
        "bke_modulO1": (13, "C"), "bke_modulO1Yorum": (13, "D"),
        "bke_modulO6": (16, "C"), "bke_modulO6Yorum": (16, "D"),
        "bke_modulO8": (17, "C"), "bke_modulO8Yorum": (17, "D"),
        "bke_modulI2": (20, "C"), "bke_modulI2Yorum": (20, "D"),
    }
    for ad, (satir, kol) in modul_map.items():
        ad_ekle(wb, ad, f"PANO!${kol}${satir}")

    ad_ekle(wb, "ListeYillar", "LISTELER!$A$6:$A$8")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$C$6:$C$7")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$6:$E$7")
    ad_ekle(wb, "ListeKrediTur", "LISTELER!$G$6:$G$7")
    ad_ekle(wb, "ListeDurumlar", "LISTELER!$I$6:$I$10")
    # IZLEME|UYARI, IZLEME|KAPALI, UYARI|IHLAL, UYARI|IYILESTIRME, UYARI|IZLEME,
    # IHLAL|IYILESTIRME, IHLAL|KAPALI, IYILESTIRME|IZLEME, IYILESTIRME|UYARI, IYILESTIRME|KAPALI = 10
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$K$6:$K$15")

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
    varsayilan = os.path.join(kok, "BankaKrediCovenantErkenUyari.xlsx")
    hedef = cikti_yolu or varsayilan
    os.makedirs(os.path.dirname(os.path.abspath(hedef)) or ".", exist_ok=True)
    wb.save(hedef)

    hsh = hashlib.sha256()
    with open(hedef, "rb") as f:
        for b in iter(lambda: f.read(65536), b""):
            hsh.update(b)
    dig = hsh.hexdigest()

    wb2 = __import__("openpyxl").load_workbook(hedef)
    wb2["AYARLAR"]["B21"] = dig[:16]
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
    cikti = os.path.join(repo, "cikti", "BankaKrediCovenantErkenUyari.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
