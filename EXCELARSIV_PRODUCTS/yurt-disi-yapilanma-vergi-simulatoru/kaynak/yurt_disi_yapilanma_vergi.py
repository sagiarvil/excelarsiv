#!/usr/bin/env python3
"""
Yurt Dışı Yapılanma Vergi Simülatörü — üretim betiği (A1 / manda v6).
Şube / iştirak / limited / ofis senaryolarında efektif vergi ve yapı kararı.
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

URUN_AD = "Yurt Dışı Yapılanma Vergi Simülatörü"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 11)

# Demo: Almanya iştirak — 500k EUR × 36 = 18M TL; yerel %30; stopaj %5; istisna Evet; TF 40
DEMO = {
    "sirket_adi": "Örnek Holding A.Ş.",
    "donem": "2026",
    "hesapYili": 2026,
    "ulke": "Almanya",
    "yapilanma_tipi": "İŞTİRAK",
    "yurt_disi_kar": 500_000,
    "yerel_vergi_orani": 0.30,
    "stopaj_orani": 0.05,
    "doviz_kuru": 36.0,
    "istirak_istisnasi": "Evet",
    "tf_risk_puani": 40,
    "yorum_modu": "A",
}


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: Yurt dışı yapılanma vergi kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Efektif vergi ve karar bozulur | Kaynak: KVK md.5 / md.33",
    )


def kural_cek(kural_id: str) -> str:
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-ydv_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(ydv_yuvarMax,IFERROR(N(G12),0)))"


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
      "Ülke, yapılanma tipi ve kâr girdilerini girin; efektif vergi, çift vergilendirme "
      "yükü ve şube / iştirak / ofis kararını senaryolu üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Şube / iştirak / limited / ofis senaryolarında efektif vergi karşılaştırması",
        "Çift vergilendirme yükü, TR ek matrah riski ve TF puanı ile risk rozeti",
        "Çok yıllı kural tablosu (2025/2026/2027) + M-SEN kural değişimi",
        "Altın vakalar ve KANIT_RAPORU (A4 imza)",
        "Belirsizlik beyanı: PE atıf yorumu A / B",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "SMMM / YMM — yapı önerisini imzalı kanıtla sunmak",
        "CFO — efektif vergi ve çift vergi yükünü bütçelemek",
        "Ortak / yönetim — şube vs iştirak kararını riskle görmek",
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
        ("Adım 1 — Girdiler",
         "GIRDI'de ülke, yapılanma tipi, YD kâr, yerel oran, stopaj, kur, istisna ve TF puanını girin."),
        ("Adım 2 — Karar",
         "PANO'da efektif vergi, çift vergi yükü, yapı önerisi ve UYGUN/DİKKAT/RİSKLİ rozetini izleyin."),
        ("Adım 3 — Kanıt",
         "KANIT_RAPORU'nu yazdırıp imzalayın; VAKALAR tutarlılığını kontrol edin."),
    ]
    for i, (b, m) in enumerate(adimlar, 5):
        h(ws, i, 1, b, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        _kim(ws, f"A{i}", b, "Sırayla izleyin")
    h(ws, 9, 1, "Sık yapılan hatalar", kalin=True, yazi=KRITIK)
    for i, m in enumerate([
        "YD kârı TL yerine döviz biriminde girip kuru 1 bırakmak",
        "İştirak istisnasını yanlış işaretlemek",
        "TF risk puanını boş bırakmak",
    ], 10):
        h(ws, i, 1, "• " + m, yazi=KRITIK)
    h(ws, 14, 1,
      "Bu dosya karar destek aracıdır; kesin vergi sonucu veya hukuki görüş yerine geçmez.",
      yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 72, "B": 70})
    for r in (10, 11, 12):
        h(ws, r, 1, ws.cell(r, 1).value, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yapılanma girdileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. TL dönüşümü ve vergi formülle üretilir.",
      yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Şirket Adı", DEMO["sirket_adi"], None, "sirket_adi",
         "Kurumun resmi unvanını yazın."),
        (7, "Dönem", DEMO["donem"], None, "donem",
         "Vergi dönemi etiketini yazın (ör. 2026)."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili",
         "2025, 2026 veya 2027 seçin."),
        (9, "Ülke", DEMO["ulke"], None, "ulke",
         "Yurt dışı yapılanmanın bulunduğu ülkeyi yazın."),
        (10, "Yapılanma Tipi", DEMO["yapilanma_tipi"], None, "yapilanma_tipi",
         "ŞUBE, İŞTİRAK, LİMİTED veya OFİS seçin."),
        (11, "Yurt Dışı Kâr [PB]", DEMO["yurt_disi_kar"], TL, "yurt_disi_kar",
         "Yerel para biriminde dönem kârını girin."),
        (12, "Yerel Vergi Oranı [%]", DEMO["yerel_vergi_orani"], YÜZDE, "yerel_vergi_orani",
         "Ülke kurumlar vergisi oranını girin."),
        (13, "Stopaj Oranı [%]", DEMO["stopaj_orani"], YÜZDE, "stopaj_orani",
         "Kâr dağıtımı / PE stopaj oranını girin."),
        (14, "Döviz Kuru [TL/PB]", DEMO["doviz_kuru"], CATI, "doviz_kuru",
         "1 yabancı para birimi = kaç TL."),
        (15, "İştirak İstisnası", DEMO["istirak_istisnasi"], None, "istirak_istisnasi",
         "KVK md.5/1-a koşulları sağlanıyorsa Evet."),
        (16, "TF Risk Puanı", DEMO["tf_risk_puani"], CATI, "tf_risk_puani",
         "Transfer fiyatı risk puanı 0-100."),
        (17, "PE Atıf Yorumu (A/B)", DEMO["yorum_modu"], None, "yorum_modu",
         "Belirsizlikte yorum A veya B seçin."),
    ]
    h(ws, 5, 1, "Alan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 2, "Değer", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    h(ws, 5, 3, "Birim", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, hiza="center")
    birim_map = {
        "hesapYili": "yıl", "yurt_disi_kar": "PB", "yerel_vergi_orani": "%",
        "stopaj_orani": "%", "doviz_kuru": "TL/PB", "tf_risk_puani": "puan",
    }
    for satir, etiket, deger, sayi, _ad, mesaj in alanlar:
        h(ws, satir, 1, etiket, yazi="333333")
        h(ws, satir, 2, deger, zemin=GIRIS_SARI, sayi=sayi, yazi="1F4E79")
        h(ws, satir, 3, birim_map.get(_ad, "metin"), yazi=GRİ)
        _kim(ws, f"B{satir}", etiket, mesaj)

    h(ws, 19, 1, "YD Kâr TL [₺]", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2, "=IFERROR(B11*B14,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B17)=0,ydv_girisBeklenen=0),0,'
      'ROUND(COUNTA(B6:B17)/ydv_girisBeklenen*100,0)),0)',
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
        ("ulke", DEMO["ulke"], "metin", "Evet", "GIRDI"),
        ("yapilanma_tipi", DEMO["yapilanma_tipi"], "liste", "Evet", "GIRDI"),
        ("yurt_disi_kar", DEMO["yurt_disi_kar"], "PB", "Evet", "GIRDI"),
        ("yerel_vergi_orani", DEMO["yerel_vergi_orani"], "%", "Evet", "GIRDI"),
        ("stopaj_orani", DEMO["stopaj_orani"], "%", "Evet", "GIRDI"),
        ("doviz_kuru", DEMO["doviz_kuru"], "TL/PB", "Evet", "GIRDI"),
        ("istirak_istisnasi", DEMO["istirak_istisnasi"], "E/H", "Evet", "GIRDI"),
        ("tf_risk_puani", DEMO["tf_risk_puani"], "puan", "Evet", "GIRDI"),
        ("yorum_modu", DEMO["yorum_modu"], "A/B", "Evet", "GIRDI"),
    ]
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 24):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in ("yurt_disi_kar",):
            ws.cell(i, 2).number_format = TL
        elif a in ("yerel_vergi_orani", "stopaj_orani"):
            ws.cell(i, 2).number_format = YÜZDE
        elif a in ("hesapYili", "doviz_kuru", "tf_risk_puani"):
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
    dogrulama(ws, "list", "ListeYapi", "B10",
              baslik="Yapılanma", mesaj="ŞUBE, İŞTİRAK, LİMİTED veya OFİS.",
              hata_baslik="Geçersiz", hata_mesaj="Listeden seçin.", bos=False)
    dogrulama(ws, "list", "ListeEvetHayir", "B15",
              baslik="İstisna", mesaj="Evet veya Hayır.",
              hata_baslik="Geçersiz", hata_mesaj="Evet/Hayır.", bos=False)
    dogrulama(ws, "list", "ListeYorumAB", "B17",
              baslik="Yorum", mesaj="A veya B seçin.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.", bos=False)
    dogrulama(ws, "list", "ListeYillar", "B26",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeYorumAB", "B35",
              baslik="Yorum", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.")
    for aralik, baslik, mesaj in [
        ("B11", "YD Kâr", "0 ile 1e12 arasında."),
        ("B14", "Kur", "0 ile 1e6 arasında."),
        ("B16", "TF Puan", "0 ile 100 arasında."),
        ("B29", "YD Kâr", "Tablo: PB tutar."),
        ("B32", "Kur", "Tablo: kur."),
        ("B34", "TF", "Tablo: puan."),
    ]:
        f2 = "100" if "TF" in baslik or "TF" in mesaj else (
            "1000000" if "Kur" in baslik else "1000000000000")
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj=mesaj,
                  hata_baslik="Geçersiz", hata_mesaj="Sınır dışı.",
                  isaret="between", f2=f2)
    for aralik, baslik in [("B12", "Yerel oran"), ("B13", "Stopaj"),
                           ("B30", "Yerel oran"), ("B31", "Stopaj")]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj="0-1 arası oran.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ile 1 arasında.",
                  isaret="between", f2="1")

    giris_hucreleri(ws, 6, 17, [2])
    giris_hucreleri(ws, 24, 35, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 36, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1024, "Sarı alanlar giriş; TL dönüşümü ve kalite formülleri kilitlidir.")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Mevzuat kural tablosu (yıl yan yana)", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.",
      yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["YDV-001", "KVK md.32", "tr_kv_orani", "her zaman", "tr_kv_orani",
         0.25, 0.25, 0.25, "01.01.2025", "KVK md.32"],
        ["YDV-002", "KVK md.5/1-a", "istirak_istisna_orani", "istirak+koşul", "istirak_istisna_orani",
         1.0, 1.0, 1.0, "01.01.2025", "KVK md.5/1-a"],
        ["YDV-003", "KVK md.33", "mahsup_tavan_carpan", "yabanci_vergi>0", "mahsup_tavan_carpan",
         1.0, 1.0, 1.0, "01.01.2025", "KVK md.33"],
        ["YDV-004", "VUK/TF", "tf_risk_esik", "her zaman", "tf_risk_esik",
         70, 70, 70, "01.01.2025", "TF mevzuatı"],
        ["YDV-005", "Genel", "yuvarlama", "her hesap", "yuvarlama_ondalik",
         2, 2, 2, "01.01.2025", "Uygulama notu"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = YÜZDE if k in (6, 7, 8) and isinstance(v, float) and v <= 1 else (
                CATI if k in (6, 7, 8) else None)
            if k in (6, 7, 8) and s[0] in ("YDV-004", "YDV-005"):
                sayi = CATI
            if k in (6, 7, 8) and s[0] in ("YDV-002", "YDV-003") and isinstance(v, float):
                sayi = YÜZDE
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-ydv_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "TR KV oranı (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("YDV-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "İştirak istisna oranı", yazi="333333")
    h(ws, 16, 2, kural_cek("YDV-002"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=_xlfn.TEXTJOIN(" | ",TRUE,"YDV-001=",TEXT(B15,"0.0%"),"YDV-002=",TEXT(B16,"0.0%"),'
      '"yıl=",hesapYili)',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    for col in ("F", "G", "H"):
        dogrulama(ws, "decimal", "0", f"{col}7:{col}11",
                  baslik="Kural değeri", mesaj="Yıl kolonuna sayısal parametre girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ile 100 arasında.",
                  isaret="between", f2="100")
    genislik(ws, {get_column_letter(i): w for i, w in enumerate(
        [14, 14, 22, 22, 20, 12, 12, 12, 12, 18, 14], 1)})
    ws.merge_cells("A3:K3")
    ws.merge_cells("A4:K4")
    sabitle(ws, "A7")


def motor(ws):
    sayfa_hazirla(ws, "MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Hesap zinciri — her adımda kural_id", kalin=True, boyut=13,
      yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar tblKurallar'dan çekilir; sabit oran yoktur.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")

    basliklar = [*MOTOR_BASLIK, "deger"]
    baslik_satiri(ws, 6, basliklar)

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif mevzuat yılı", "YDV-005")  # G7
    ekle("TR KV oranı", km("YDV-001"), "oran", "KVK md.32 oranı", "YDV-001")  # G8
    ekle("İştirak istisna oranı", km("YDV-002"), "oran", "KVK md.5/1-a", "YDV-002")  # G9
    ekle("Mahsup tavan çarpanı", km("YDV-003"), "çarpan", "KVK md.33 tavan", "YDV-003")  # G10
    ekle("TF risk eşiği", km("YDV-004"), "puan", "TF eşik puanı", "YDV-004")  # G11
    ekle("Yuvarlama ondalık", km("YDV-005"), "adet", "Yuvarlama basamağı", "YDV-005")  # G12
    ekle("YD kâr PB", "=GIRDI!B11", "PB", "Yerel para kârı", "YDV-001")  # G13
    ekle("Yerel vergi oranı", "=GIRDI!B12", "oran", "Ülke KV oranı", "YDV-001")  # G14
    ekle("Stopaj oranı", "=GIRDI!B13", "oran", "Stopaj", "YDV-001")  # G15
    ekle("Döviz kuru", "=GIRDI!B14", "kur", "TL / PB", "YDV-005")  # G16
    ekle("TF risk puanı", "=GIRDI!B16", "puan", "Transfer fiyat riski", "YDV-004")  # G17
    ekle("İstisna bayrağı",
         '=IF(GIRDI!B15="Evet",1,0)',
         "bayrak", "İştirak istisnası", "YDV-002")  # G18
    ekle("Yorum çarpanı",
         '=IF(GIRDI!B17="B",1+ydv_peArtisOran,1)',
         "çarpan", "PE atıf A/B", "YDV-005")  # G19
    ekle("YD kâr TL", f"=ROUND(G13*G16,{_YV})", "TL", "Kâr × kur", "YDV-005")  # G20
    ekle("Yerel vergi TL", f"=ROUND(G20*G14,{_YV})", "TL", "YD kâr × yerel oran", "YDV-001")  # G21
    ekle("Stopaj TL", f"=ROUND(G20*G15,{_YV})", "TL", "YD kâr × stopaj", "YDV-001")  # G22
    # TR matrah: ŞUBE tam; OFİS düşük pay×yorum; İŞTİRAK/LİMİTED istisna varsa dağıtım payı
    ekle("TR matrah",
         f'=ROUND(IF(GIRDI!B10="ŞUBE",G20,'
         f'IF(GIRDI!B10="OFİS",G20*ydv_ofisMatrahOran*G19,'
         f'IF(AND(OR(GIRDI!B10="İŞTİRAK",GIRDI!B10="LİMİTED"),G18=1),'
         f'G20*(1-G9)*ydv_dagitimOran,G20))),{_YV})',
         "TL", "Yapıya göre TR matrah", "YDV-002")  # G23
    ekle("TR brüt KV", f"=ROUND(G23*G8,{_YV})", "TL", "Matrah × TR KV", "YDV-001")  # G24
    ekle("Mahsup tavan", f"=ROUND(G24*G10,{_YV})", "TL", "Mahsup tavanı", "YDV-003")  # G25
    ekle("Mahsup tutarı", "=MIN(G21+G22,G25)", "TL", "Yabancı vergi mahsubu", "YDV-003")  # G26
    ekle("TR net KV", "=MAX(0,G24-G26)", "TL", "TR'de kalan KV", "YDV-003")  # G27
    ekle("Toplam vergi", "=G21+G22+G27", "TL", "Yerel+stopaj+TR net", "YDV-001")  # G28
    ekle("Efektif vergi", "=IF(G20=0,0,G28/G20)", "oran", "Toplam/kâr", "YDV-001")  # G29
    ekle("Çift vergi yükü", "=MAX(0,G21+G22+G27-G24)", "TL", "Çakışan yük", "YDV-003")  # G30
    ekle("TR ek matrah riski",
         f"=ROUND(G20*G17/ydv_tfBolum*ydv_ekMatrahCarpan*G19,{_YV})",
         "TL", "TF kaynaklı ek matrah", "YDV-004")  # G31
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(OR(G17>=G11,G30>G20*ydv_ciftVergiEsik),"RISKLI",'
         'IF(OR(G17>=ydv_dikkatEsik,G30>G20*ydv_dikkatCiftEsik),"DIKKAT","UYGUN")))',
         "kod", "Risk karar kodu", "YDV-004")  # G32
    ekle("Karar metni",
         '=IF(G32="VERI_YOK","VERİ YOK",'
         'IF(G32="RISKLI","RİSKLİ",'
         'IF(G32="DIKKAT","DİKKAT","UYGUN")))',
         "metin", "Kullanıcıya karar", "YDV-004")  # G33
    ekle("Yapı önerisi",
         '=IF(G32="VERI_YOK","VERİ YOK",'
         'IF(AND(G18=1,OR(GIRDI!B10="İŞTİRAK",GIRDI!B10="LİMİTED")),"İŞTİRAK",'
         'IF(GIRDI!B10="OFİS","OFİS",'
         'IF(G29<=ydv_subeEfektifEsik,"ŞUBE","İŞTİRAK"))))',
         "metin", "Önerilen yapı", "YDV-002")  # G34
    ekle("Gerekçe",
         '=IF(G32="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G32="RISKLI","TF eşiği aşıldı veya çift vergi yükü yüksek; yapı ve TF dosyasını gözden geçirin.",'
         'IF(G32="DIKKAT","Orta risk: mahsup, stopaj ve PE atıfını doğrulayın.",'
         '"Efektif vergi ve TF profili kabul edilebilir aralıkta.")))',
         "metin", "Gerekçe cümlesi", "YDV-004")  # G35
    ekle("Yerel / kâr payı", "=IF(G20=0,0,G21/G20)", "oran", "T1 pay", "YDV-001")  # G36
    ekle("TR net / toplam", "=IF(G28=0,0,G27/G28)", "oran", "T2 pay", "YDV-001")  # G37
    ekle("Mahsup / brüt", "=IF(G24=0,0,G26/G24)", "oran", "O6 yoğunluk", "YDV-003")  # G38
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B17)=0,ydv_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B17)/ydv_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "YDV-005")  # G39
    ekle("Risk skoru",
         '=IF(G32="VERI_YOK",0,IF(G32="UYGUN",ydv_riskDusuk,'
         'IF(G32="DIKKAT",ydv_riskOrta,ydv_riskYuksek)))',
         "puan", "Risk skoru", "YDV-004")  # G40
    ekle("Senaryo iyimser kâr", "=G20*ydv_senaryoIyi", "TL", "İyimser kâr TL", "YDV-001")  # G41
    ekle("Senaryo kötümser kâr", "=G20*ydv_senaryoKotu", "TL", "Kötümser kâr TL", "YDV-001")  # G42
    ekle("İyimser toplam vergi", f"=ROUND(G41*G29,{_YV})", "TL", "İyimser yük", "YDV-001")  # G43
    ekle("Kötümser toplam vergi", f"=ROUND(G42*G29,{_YV})", "TL", "Kötümser yük", "YDV-001")  # G44
    ekle("M-SEN TR oran+", "=MIN(1,G8+ydv_oranArtisPuan)", "oran", "TR KV +1 puan", "YDV-001")  # G45
    ekle("M-SEN TR brüt", f"=ROUND(G23*G45,{_YV})", "TL", "M-SEN brüt", "YDV-001")  # G46
    ekle("M-SEN TR net", "=MAX(0,G46-G26)", "TL", "M-SEN net", "YDV-001")  # G47
    ekle("M-SEN toplam", "=G21+G22+G47", "TL", "M-SEN toplam vergi", "YDV-001")  # G48
    ekle("Tahmin üst", "=G28*ydv_tahminUst", "TL", "Tahmin üst bant", "YDV-005")  # G49
    ekle("Tahmin alt", "=G28*ydv_tahminAlt", "TL", "Tahmin alt bant", "YDV-005")  # G50
    ekle("Tornado: kur etkisi",
         f"=ABS(ROUND(G13*G16*ydv_tornadoKur*G14,{_YV})-G21)",
         "TL", "Kur ± etki", "YDV-005")  # G51
    ekle("Tornado: yerel oran",
         f"=ABS(ROUND(G20*(G14+ydv_oranArtisPuan),{_YV})-G21)",
         "TL", "Yerel oran ±", "YDV-001")  # G52
    ekle("Tornado: TF etkisi",
         f"=ABS(ROUND(G20*(G17+ydv_tornadoTf)/ydv_tfBolum*ydv_ekMatrahCarpan,{_YV})-G31)",
         "TL", "TF ± etki", "YDV-004")  # G53
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"YDV-001→KVK md.32","YDV-002→KVK md.5/1-a",'
         '"YDV-003→KVK md.33","YDV-004→TF")',
         "metin", "Kanıt atıfları", "YDV-001")  # G54
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"KârTL=",TEXT(G20,"₺ #,##0"),'
         '"Efektif=",TEXT(G29,"0.0%"),"Çift=",TEXT(G30,"₺ #,##0"),'
         '"Karar=",G33,"Yapı=",G34)',
         "metin", "Rapor özeti", "YDV-005")  # G55

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
        elif birim in ("puan", "yıl", "adet", "bayrak", "kur", "PB"):
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
    h(ws, 5, 2, "Kâr Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Toplam Vergi", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Efektif", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(ydv_senaryoIyi+N(GIRDI!B11)*ydv_sifirCarpan,0)",
         "=IFERROR(ROUND(IFERROR(ydv_karTl,0)*B6*IFERROR(ydv_efektifVergi,0),2),0)",
         "=IFERROR(ydv_efektifVergi,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(ydv_kararMetni,"-"))'),
        ("Baz", "=IFERROR(ydv_bazCarpan+N(GIRDI!B11)*ydv_sifirCarpan,ydv_bazCarpan)",
         "=IFERROR(ydv_toplamVergi,0)",
         "=IFERROR(ydv_efektifVergi,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(ydv_kararMetni,"-"))'),
        ("Kötümser", "=IFERROR(ydv_senaryoKotu+N(GIRDI!B11)*ydv_sifirCarpan,0)",
         "=IFERROR(ROUND(IFERROR(ydv_karTl,0)*B8*IFERROR(ydv_efektifVergi,0),2),0)",
         "=IFERROR(ydv_efektifVergi,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(ydv_kararMetni,"-"))'),
        ("M-SEN TR+1", "=IFERROR(ydv_bazCarpan+N(GIRDI!B11)*ydv_sifirCarpan,ydv_bazCarpan)",
         "=IFERROR(MOTOR!G48,0)",
         "=IFERROR(IF(ydv_karTl=0,0,MOTOR!G48/ydv_karTl),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK","DİKKAT")'),
    ]
    for i, (ad, carp, top, ef, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, top, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, ef, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant genişliği (iyimser−kötümser vergi)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(C6-C8,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser ",TEXT(C6,"₺ #,##0")," · Baz ",TEXT(C7,"₺ #,##0"),'
      '" · Kötümser ",TEXT(C8,"₺ #,##0")," · Bant ",TEXT(B11,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:H12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Tornado kur etkisi", yazi="333333")
    h(ws, 16, 2, "=IFERROR(MOTOR!G51,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado yerel oran", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G52,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado TF etkisi", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G53,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 3,
      '=IFERROR(IF(B16>=MAX(B16:B18),ydv_siraBir,IF(B16>=MIN(B16:B18)+ABS(B16-B17),ydv_siraIki,ydv_siraUc)),ydv_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 3,
      '=IFERROR(IF(B17>=MAX(B16:B18),ydv_siraBir,IF(B17=MEDIAN(B16:B18),ydv_siraIki,ydv_siraUc)),ydv_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 3,
      '=IFERROR(IF(B18>=MAX(B16:B18),ydv_siraBir,IF(B18=MEDIAN(B16:B18),ydv_siraIki,ydv_siraUc)),ydv_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(CONCATENATE("1)",INDEX(A16:A18,MATCH(ydv_siraBir,C16:C18,0)),'
      '" 2)",INDEX(A16:A18,MATCH(ydv_siraIki,C16:C18,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B16:B18),"₺ #,##0"),'
      '" TL — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN kural değişimi toplam", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=C9", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"TR KV +1 puan senaryosunda toplam vergi ",TEXT(C9,"₺ #,##0")," TL")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı (toplam vergi)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=MOTOR!G50", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 3, "=ydv_toplamVergi", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 4, "=MOTOR!G49", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Tahmin (FORECAST)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1, "Seri nokta", yazi=GRİ)
    for i, v in enumerate([1, 2, 3, 4], 28):
        h(ws, i, 1, v, sayi=CATI)
        h(ws, i, 2, f"=C{5+v}", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 32, 1, "Tahmin sonraki", yazi="333333")
    h(ws, 32, 2, "=IFERROR(FORECAST.LINEAR(ydv_tahminX,B28:B31,A28:A31),0)", sayi=TL,
      yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Tahmin yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 33, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Tahmin ",TEXT(B32,"₺ #,##0")," · Alt ",TEXT(B25,"₺ #,##0"),'
      '" · Üst ",TEXT(D25,"₺ #,##0"))',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 5, 7, "VergiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([5_940_000, 6_300_000, 7_200_000, 6_750_000], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([270000, 180000, 360000], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)
    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([5_940_000, 6_300_000, 7_200_000, 6_750_000], 28):
        h(ws, i, 4, v, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo Toplam Vergi"
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
    g3.title = "Senaryo Vergi Trendi"
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

    # V-001: 18M × %30 yerel = 5.4M
    # V-002: 18M × %25 TR = 4.5M (şube brüt, mahsup öncesi)
    # V-003: istisna %100 → TR matrah dağıtım payı düşük
    # V-004: TF 80 ≥ 70 → riskli eşik
    vakalar_data = [
        ("V-001", "Yerel vergi — Almanya %30",
         "kar_tl=vaka_kar_1; oran=vaka_yerel_oran",
         5_400_000,
         "=IFERROR(ROUND(vaka_kar_1*vaka_yerel_oran+N(GIRDI!B11)*ydv_sifirCarpan,2),0)",
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=ydv_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Yerel KV örnek"),
        ("V-002", "TR brüt KV — şube matrah",
         "kar_tl=vaka_kar_1; TR=%25",
         4_500_000,
         "=IFERROR(ROUND(vaka_kar_1*INDEX(tblKurallar[deger_2025],MATCH(\"YDV-001\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=ydv_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "KVK md.32"),
        ("V-003", "İştirak istisna sonrası TR matrah",
         "kar=vaka_kar_1; istisna=%100; dağıtım=ydv_dagitimOran",
         0,
         "=IFERROR(ROUND(vaka_kar_1*(1-INDEX(tblKurallar[deger_2025],MATCH(\"YDV-002\",tblKurallar[kural_id],0)))*ydv_dagitimOran,2),0)",
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=ydv_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "KVK md.5/1-a"),
        ("V-004", "TF eşik kontrolü",
         "esik=YDV-004; puan=vaka_tf_puan",
         70,
         "=IFERROR(INDEX(tblKurallar[deger_2025],MATCH(\"YDV-004\",tblKurallar[kural_id],0)),0)",
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=ydv_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "TF mevzuatı"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            sayi = TL if k in (4, 5, 6) and not (isinstance(v, str) and v.startswith("=")) else None
            if k == 4 and row[0] == "V-004":
                sayi = CATI
            h(ws, i, k, v, sayi=sayi if not (isinstance(v, str) and v.startswith("=")) else None,
              yazi=DIS_REF_YEŞIL if isinstance(v, str) and v.startswith("=") else "333333")
            if isinstance(v, str) and v.startswith("=") and k in (4, 5, 6):
                ws.cell(i, k).number_format = TL if row[0] != "V-004" else CATI
            if k == 4 and row[0] != "V-004":
                ws.cell(i, k).number_format = TL
            if k == 4 and row[0] == "V-004":
                ws.cell(i, k).number_format = CATI

    formuller = {
        "durum": (
            '=IF(tblVakalar[[#This Row],[vaka_id]]="","",'
            'IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=ydv_tolerans,"TUTARLI","KIRIK"),"KIRIK"))'
        ),
    }
    tablo_ekle(ws, "tblVakalar", "A5:H9", VAKA_BASLIK, formuller)

    h(ws, 11, 1, "Vaka özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 1, "Durum", yazi="333333")
    h(ws, 12, 2,
      '=IFERROR(IF(COUNTIF(G6:G9,"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "Toplam fark", yazi="333333")
    h(ws, 13, 2, "=IFERROR(SUM(F6:F9),0)", sayi=TL, yazi=DIS_REF_YEŞIL)

    h(ws, 5, 10, "Beklenen", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 11, "Hesaplanan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, (be, he) in enumerate([
        (5_400_000, 5_400_000), (4_500_000, 4_500_000), (0, 0), (70, 70),
    ], 6):
        h(ws, i, 10, be, sayi=TL if i < 9 else CATI, yazi=GRİ)
        h(ws, i, 11, he, sayi=TL if i < 9 else CATI, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Vaka Beklenen vs Hesaplanan"
    g.add_data(Reference(ws, min_col=10, min_row=5, max_col=11, max_row=9), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=6, max_row=9))
    g.height, g.width = 8, 14
    ws.add_chart(g, "A15")

    genislik(ws, {"A": 12, "B": 36, "C": 40, "D": 14, "E": 14, "F": 12, "G": 12, "H": 18})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", "0F2742", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "KANIT RAPORU — Yurt Dışı Yapılanma Vergi", kalin=True, boyut=14,
      yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:F3")
    h(ws, 4, 1, "Rapor tarihi", yazi="333333")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 5, 1, "Şirket", yazi="333333")
    h(ws, 5, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Dönem / yıl / ülke", yazi="333333")
    h(ws, 6, 2, '=CONCATENATE(GIRDI!B7," / ",hesapYili," / ",GIRDI!B9)', yazi=DIS_REF_YEŞIL)

    h(ws, 8, 1, "Girdi özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, (et, form, sayi) in enumerate([
        ("Yapılanma tipi", "=GIRDI!B10", None),
        ("YD kâr (PB)", "=GIRDI!B11", TL),
        ("Yerel vergi oranı", "=GIRDI!B12", YÜZDE),
        ("Stopaj oranı", "=GIRDI!B13", YÜZDE),
        ("Döviz kuru", "=GIRDI!B14", CATI),
        ("İştirak istisnası", "=GIRDI!B15", None),
        ("TF risk puanı", "=GIRDI!B16", CATI),
    ], 9):
        h(ws, i, 1, et, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)

    h(ws, 17, 1, "Hesap zinciri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 18, 1, "YD kâr TL", yazi="333333")
    h(ws, 18, 2, "=ydv_karTl", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Efektif vergi", yazi="333333")
    h(ws, 19, 2, "=ydv_efektifVergi", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Çift vergi yükü", yazi="333333")
    h(ws, 20, 2, "=ydv_ciftVergi", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 21, 1, "TR ek matrah riski", yazi="333333")
    h(ws, 21, 2, "=ydv_trEkMatrah", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 22, 1, "KARAR", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=ydv_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "Yapı önerisi", yazi="333333")
    h(ws, 23, 2, "=ydv_yapiOneri", kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 24, 1, "Gerekçe", yazi="333333")
    h(ws, 24, 2, "=ydv_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B24:F24")

    h(ws, 26, 1, "Madde atıfları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 2, "=ydv_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B26:F26")

    h(ws, 28, 1, "Kanıt gövdesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 28, 2, "=ydv_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B28:F28")

    h(ws, 30, 1, "Parmak izi (SHA-256 kısaltma)", yazi="333333")
    h(ws, 30, 2, "=ydv_parmakIzi", yazi=GRİ)

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
    h(ws, 3, 1, "Yönetim paneli — Yurt Dışı Yapılanma", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpiler = [
        (2, "KârTL", "=IFERROR(ydv_karTl,0)", TL),
        (3, "YerelV", "=IFERROR(MOTOR!G21,0)", TL),
        (4, "TRNet", "=IFERROR(MOTOR!G27,0)", TL),
        (5, "Toplam", "=IFERROR(ydv_toplamVergi,0)", TL),
        (6, "Efektif", "=IFERROR(ydv_efektifVergi,0)", YÜZDE),
        (7, "ÇiftV", "=IFERROR(ydv_ciftVergi,0)", TL),
        (8, "EkMatrah", "=IFERROR(ydv_trEkMatrah,0)", TL),
        (9, "Güven", "=IFERROR(MOTOR!G39,0)", CATI),
        (10, "Risk", "=IFERROR(MOTOR!G40,0)", CATI),
        (11, "M-SEN", "=IFERROR(ydv_kuralDegisimSenaryo,0)", TL),
        (12, "Tahmin", "=IFERROR(SENARYO!B32,0)", TL),
        (13, "Vaka", "=IFERROR(ydv_vakaDurum,\"-\")", None),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=9, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=ydv_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Yapı önerisi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=ydv_yapiOneri", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 8, 2, "=ydv_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B8:H8")

    h(ws, 10, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 11, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Yerel vergi / YD kâr TL",
         "=IFERROR(MOTOR!G36,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Yerel pay ",TEXT(C12,"0.0%"))'),
        ("T2", "TR net / toplam vergi",
         "=IFERROR(MOTOR!G37,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"TR net pay ",TEXT(C13,"0.0%"))'),
        ("O1", "Senaryo vergi sapması",
         "=IFERROR(IF(COUNT(SENARYO!C6:C9)<2,0,STDEV.P(SENARYO!C6:C9)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo sapması ",TEXT(C14,"₺ #,##0")," TL")'),
        ("O2", "Tornado max etki",
         "=IFERROR(ydv_duyarlilikSira,\"-\")",
         "=IFERROR(ydv_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo bant",
         "=IFERROR(ydv_senaryoKarsilastirma,0)",
         "=IFERROR(ydv_yorumSenaryo,\"-\")"),
        ("O6", "Mahsup / TR brüt",
         "=IFERROR(MOTOR!G38,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Mahsup yoğunluğu ",TEXT(C17,"0.0%"))'),
        ("O8", "Kalite skoru",
         "=IFERROR(KONTROLLER!B12,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalite ",TEXT(C18,"0")," puan")'),
        ("M", "M-SEN toplam vergi",
         "=IFERROR(ydv_kuralDegisimSenaryo,0)",
         "=IFERROR(ydv_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin aralık",
         "=IFERROR(ydv_tahminAralik,0)",
         "=IFERROR(ydv_yorumTahmin,\"-\")"),
        ("I2", "Senaryo vergi P90",
         "=IFERROR(IF(COUNT(SENARYO!C6:C9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!C6:C9,ydv_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"P90 vergi ",TEXT(C21,"₺ #,##0")," TL")'),
    ]
    for i, (kod, ad, form, yorum) in enumerate(moduller):
        r = 12 + i
        h(ws, r, 1, kod, kalin=True, hiza="center", yazi="B08948")
        h(ws, r, 2, ad, kaydir=True)
        h(ws, r, 3, form, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[r].height = 28
        if kod in ("T1", "T2", "O6"):
            ws.cell(r, 3).number_format = YÜZDE
        elif kod != "O2":
            ws.cell(r, 3).number_format = TL

    h(ws, 24, 1, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("YD kâr TL", 18_000_000),
        ("Yerel vergi", 5_400_000),
        ("Stopaj", 900_000),
        ("TR net", 0),
        ("Toplam vergi", 6_300_000),
    ], 25):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)

    h(ws, 24, 4, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 5, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("İyimser vergi", 5_940_000),
        ("Baz vergi", 6_300_000),
        ("Kötümser vergi", 7_200_000),
        ("M-SEN vergi", 6_750_000),
        ("Çift vergi", 1_800_000),
    ], 25):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Vergi Bileşenleri"
    g1.add_data(Reference(ws, min_col=2, min_row=24, max_row=29), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=25, max_row=29))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G10")

    g2 = BarChart()
    g2.type = "col"
    g2.title = "Senaryo Vergi"
    g2.add_data(Reference(ws, min_col=5, min_row=24, max_row=29), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=25, max_row=29))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "G24")

    g3 = LineChart()
    g3.title = "Vergi Çizgisi"
    g3.add_data(Reference(ws, min_col=5, min_row=24, max_row=29), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=4, min_row=25, max_row=29))
    g3.height, g3.width = 7, 12
    ws.add_chart(g3, "P10")

    h(ws, 31, 1, "Metrik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 31, 2, "Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("YD kâr TL", 18_000_000),
        ("Toplam vergi", 6_300_000),
        ("Çift vergi", 1_800_000),
    ], 32):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "KPI Kâr / Vergi / Çift"
    g4.add_data(Reference(ws, min_col=2, min_row=31, max_row=34), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=1, min_row=32, max_row=34))
    g4.height, g4.width = 7, 10
    ws.add_chart(g4, "P24")

    baski_hazirla(ws, "A1:F36", f"{URUN_AD} · PANO")
    genislik(ws, {"A": 28, "B": 16, "C": 18, "D": 55})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Canlı kontrol paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    kontroller_list = [
        ("Giriş tablosu boş mu?",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"BOŞ","DOLU")'),
        ("Negatif YD kâr?",
         '=IF(GIRDI!B11<0,"NEGATİF","TEMİZ")'),
        ("Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(ydv_yilTaban),N(hesapYili)<=N(ydv_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        ("Kâr TL sıfır mı?",
         '=IF(IFERROR(MOTOR!G20,0)=0,"SIFIR","NORMAL")'),
        ("Vaka kırık mı?",
         '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")'),
        ("Çift vergi eşiği aşıyor mu?",
         '=IF(IFERROR(MOTOR!G30,0)>ydv_kacirilanEsik+N(GIRDI!B11)*ydv_sifirCarpan,"AŞIYOR","NORMAL")'),
        ("Yorum A/B seçili mi?",
         '=IF(OR(GIRDI!B17="A",GIRDI!B17="B"),"TEMİZ","EKSİK")'),
        ("Motor adım sayısı",
         "=COUNTA(MOTOR!B7:B55)"),
    ]
    for i, (ad, form) in enumerate(kontroller_list, 5):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Canlı formül | Değiştirmeyin")

    h(ws, 14, 1, "Toplam giriş alanı", yazi="333333")
    h(ws, 14, 2, "=IFERROR(ydv_girisBeklenen+COUNTA(GIRDI!B6:B17)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Dolu giriş", yazi="333333")
    h(ws, 15, 2, "=COUNTA(GIRDI!B6:B17)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Kalite skoru (modül)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=IFERROR(ROUND(IF(OR(B14=0,B14=\"\"),0,B15/B14*100),0),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)

    h(ws, 17, 1, "Kâr TL kontrol", yazi="333333")
    h(ws, 17, 2, "=IFERROR(ydv_karTl,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Efektif kontrol", yazi="333333")
    h(ws, 18, 2, "=IFERROR(ydv_efektifVergi,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Karar kontrol", yazi="333333")
    h(ws, 19, 2, "=IFERROR(ydv_kararMetni,\"-\")", yazi=DIS_REF_YEŞIL)

    genislik(ws, {"A": 40, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek senaryo kütüphanesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:L3")
    h(ws, 4, 1, "Bu sayfa demo değerleri saklar; GIRDI'ye kopyalanabilir.", yazi=GRİ)
    sutunlar = ["Senaryo", "Ülke", "Yapı", "YDKâr", "YerelOran", "Stopaj",
                "Kur", "İstisna", "TF", "Yıl", "Yorum", "KârTL"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "KârTL": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(tblOrnek[[#This Row],[YDKâr]]*tblOrnek[[#This Row],[Kur]],0))"
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:L1004", sutunlar, formuller)
    ornekler = [
        ("Demo ana", "Almanya", "İŞTİRAK", 500_000, 0.30, 0.05, 36, "Evet", 40, 2026, "A"),
        ("Şube yüksek vergi", "Fransa", "ŞUBE", 400_000, 0.28, 0.00, 36, "Hayır", 55, 2026, "A"),
        ("Ofis düşük risk", "BAE", "OFİS", 100_000, 0.09, 0.00, 10, "Hayır", 20, 2026, "A"),
        ("TF riskli limited", "Hollanda", "LİMİTED", 600_000, 0.25, 0.05, 36, "Evet", 85, 2026, "B"),
        ("2025 geçiş", "Almanya", "İŞTİRAK", 500_000, 0.30, 0.05, 34, "Evet", 40, 2025, "A"),
    ]
    for i, row in enumerate(ornekler, 6):
        for k, v in enumerate(row, 1):
            ws.cell(i, k).value = v
            if k == 4:
                ws.cell(i, k).number_format = TL
            elif k in (5, 6):
                ws.cell(i, k).number_format = YÜZDE
            elif k in (7, 9, 10):
                ws.cell(i, k).number_format = CATI
            if k <= 11:
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", sutunlar[k - 1], "Örnek senaryo değeri")

    for col, ad in [("D", "YDKâr"), ("G", "Kur"), ("I", "TF")]:
        dogrulama(ws, "decimal", "0", f"{col}6:{col}1004",
                  baslik=ad, mesaj=f"{ad} değerini girin.",
                  hata_baslik="Geçersiz", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000000")
    for col, ad in [("E", "Yerel"), ("F", "Stopaj")]:
        dogrulama(ws, "decimal", "0", f"{col}6:{col}1004",
                  baslik=ad, mesaj="0-1 oran.",
                  hata_baslik="Geçersiz", hata_mesaj="0-1.",
                  isaret="between", f2="1")

    giris_hucreleri(ws, 6, 1004, list(range(1, 12)))
    h(ws, 5, 14, "KarDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([18_000_000, 14_400_000, 1_000_000, 21_600_000, 17_000_000], 6):
        h(ws, i, 14, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Örnek Senaryo Kâr TL"
    g.add_data(Reference(ws, min_col=14, min_row=5, max_row=10), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=6, max_row=10))
    g.height, g.width = 8, 12
    ws.add_chart(g, "N5")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 15)})
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
        _kim(ws, f"C{i}", "Yorum", "PE atıf yorumu")
    h(ws, 5, 5, "Evet/Hayır", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate(["Evet", "Hayır"], 6):
        h(ws, i, 5, v, zemin=GIRIS_SARI)
        ws.cell(i, 5).protection = Protection(locked=False)
        _kim(ws, f"E{i}", "Seçim", "Evet veya Hayır")
    h(ws, 5, 7, "Yapılanma", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate(["ŞUBE", "İŞTİRAK", "LİMİTED", "OFİS"], 6):
        h(ws, i, 7, v, zemin=GIRIS_SARI)
        ws.cell(i, 7).protection = Protection(locked=False)
        _kim(ws, f"G{i}", "Yapı", "Yapılanma tipi")
    genislik(ws, {"A": 12, "C": 12, "E": 12, "G": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Parametreler (tek kaynak)", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    sutunlar = ["anahtar", "deger", "birim", "aciklama", "kaynak", "yururluk_tarihi",
                "dogrulama_tarihi", "kontrol"]
    baslik_satiri(ws, 5, sutunlar)
    params = [
        ("hesapYili", DEMO["hesapYili"], "yıl", "Aktif hesap yılı", "Kullanıcı / GIRDI",
         "01.01.2025", "11.08.2026"),
        ("raporTarihi", RAPOR_TARIH, "tarih", "Rapor tarihi (sabit)", "Üretim",
         "01.01.2025", "11.08.2026"),
        ("ydv_tolerans", 0.01, "TL", "Vaka tutarlılık toleransı", "Uygulama notu",
         "01.01.2025", "11.08.2026"),
        ("ydv_kacirilanEsik", 500_000, "TL", "Çift vergi uyarı eşiği", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("ydv_senaryoIyi", 0.90, "çarpan", "İyimser kâr çarpanı", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("ydv_senaryoKotu", 1.15, "çarpan", "Kötümser kâr çarpanı", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("ydv_oranArtisPuan", 0.01, "oran", "M-SEN sabit +1 puan", "Senaryo motoru",
         "01.01.2025", "11.08.2026"),
        ("ydv_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("ydv_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("ydv_tornadoKur", 1.05, "çarpan", "Tornado kur şoku", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("ydv_tornadoTf", 10, "puan", "Tornado TF şoku", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("ydv_peArtisOran", 0.10, "oran", "Yorum B PE +%10", "PE atıf modeli",
         "01.01.2025", "11.08.2026"),
        ("ydv_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik",
         "01.01.2025", "11.08.2026"),
        ("ydv_parmakIzi", "YDV-PRO-1.0.0", "metin", "Dosya parmak izi etiketi", "Üretim",
         "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv",
         "01.01.2025", "11.08.2026"),
        ("ydv_girisBeklenen", 12, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("ydv_riskDusuk", 15, "puan", "UYGUN risk skoru", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("ydv_riskOrta", 55, "puan", "DİKKAT risk skoru", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("ydv_riskYuksek", 85, "puan", "RİSKLİ risk skoru", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("ydv_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("ydv_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("ydv_bazCarpan", 1, "çarpan", "Baz senaryo çarpanı", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("ydv_sifirCarpan", 0, "çarpan", "Nötr çarpan (sıfır)", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("ydv_siraBir", 1, "adet", "Tornado sıra 1", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("ydv_siraIki", 2, "adet", "Tornado sıra 2", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("ydv_siraUc", 3, "adet", "Tornado sıra 3", "Ürün mimarisi",
         "01.01.2025", "11.08.2026"),
        ("ydv_tahminX", 5, "adet", "FORECAST X noktası", "Model varsayımı",
         "01.01.2025", "11.08.2026"),
        ("ydv_ofisMatrahOran", 0.15, "oran", "Ofis TR matrah payı", "PE modeli",
         "01.01.2025", "11.08.2026"),
        ("ydv_dagitimOran", 0.0, "oran", "İstisna sonrası dağıtım payı", "KVK md.5",
         "01.01.2025", "11.08.2026"),
        ("ydv_tfBolum", 100, "adet", "TF puan böleni", "Model",
         "01.01.2025", "11.08.2026"),
        ("ydv_ekMatrahCarpan", 0.25, "oran", "TF ek matrah çarpanı", "Model",
         "01.01.2025", "11.08.2026"),
        ("ydv_ciftVergiEsik", 0.20, "oran", "RİSKLİ çift vergi / kâr", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("ydv_dikkatEsik", 50, "puan", "DİKKAT TF eşiği", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("ydv_dikkatCiftEsik", 0.10, "oran", "DİKKAT çift vergi / kâr", "İç politika",
         "01.01.2025", "11.08.2026"),
        ("ydv_subeEfektifEsik", 0.28, "oran", "Şube önerisi efektif tavan", "Model",
         "01.01.2025", "11.08.2026"),
        ("vaka_kar_1", 18_000_000, "TL", "Vaka1/2/3 kâr TL", "Tebliğ örneği",
         "01.01.2025", "11.08.2026"),
        ("vaka_yerel_oran", 0.30, "oran", "Vaka1 yerel oran", "Tebliğ örneği",
         "01.01.2025", "11.08.2026"),
        ("vaka_tf_puan", 70, "puan", "Vaka4 TF eşik beklenen", "Tebliğ örneği",
         "01.01.2025", "11.08.2026"),
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
              baslik="Eşik", mesaj="Uyarı eşiği TL.",
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

    genislik(ws, {"A": 28, "B": 24, "C": 10, "D": 36, "E": 22, "F": 14, "G": 14, "H": 12})
    h(ws, son_satir + 4, 2,
      f"Sürüm {SURUM} | Şifre koruması: {SIFRE} (formül alanları)",
      yazi=GRİ, boyut=9, kaydir=True)


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "1F4E79", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar kuralları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "VERİ YOK — giriş tablosu boşsa hesap yapılmaz.",
        "UYGUN — TF ve çift vergi yükü düşük; mevcut yapı sürdürülebilir.",
        "DİKKAT — orta risk; mahsup, stopaj ve PE atıfını doğrulayın.",
        "RİSKLİ — TF eşiği aşıldı veya çift vergi yüksek; yapıyı yeniden değerlendirin.",
        "Yapı önerisi — İŞTİRAK / ŞUBE / OFİS / VERİ YOK (risk + efektif vergiye göre).",
    ], 5):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 11, 1, "Belirsizlik beyanları (M05)", kalin=True, boyut=12, yazi=KRITIK)
    for i, m in enumerate(belirsizlik_beyanlari(["pe_atif_yorumu"]), 12):
        h(ws, i, 1, m, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=20)

    h(ws, 14, 1, "Yorum A", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 14, 2,
      "PE atıfı ve ofis matrah payı girdideki oranlarla aynen uygulanır. "
      "İştirak istisnası işaretliyse TR matrahı dağıtım payına iner.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B14:J14")
    h(ws, 15, 1, "Yorum B", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 2,
      "Yorum B, PE / ofis atıfını +%10 sıkılaştırır (daha yüksek TR matrah ve ek risk). "
      "Dosya her iki yorumun etkisini MOTOR'da üretir.",
      kaydir=True, yazi="333333")
    ws.merge_cells("B15:J15")

    h(ws, 17, 1, "Kullanım sırası", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 1,
      "GIRDI → KURALLAR (gerekirse) → MOTOR → SENARYO → VAKALAR → PANO → KANIT_RAPORU → AYARLAR.",
      kaydir=True, yazi="333333")
    ws.merge_cells("A18:J18")
    h(ws, 20, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 18, "B": 70})


def kosullu_bicimlendirme(wb):
    yesil = Font(color=NORMAL, bold=True)
    kirmizi = Font(color=KRITIK, bold=True)
    amber = Font(color="B7791F", bold=True)
    y_fill = PatternFill("solid", fgColor="E2EFDA")
    k_fill = PatternFill("solid", fgColor="FDE9E9")
    a_fill = PatternFill("solid", fgColor="FFF2CC")

    ws = wb["PANO"]
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("VERİ YOK",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("RİSKLİ",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("UYGUN",B6))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("DİKKAT",B6))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("I4", CellIsRule(
        operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("I4", CellIsRule(
        operator="between", formula=["50", "79"], font=amber))
    ws.conditional_formatting.add("I4", CellIsRule(
        operator="lessThan", formula=["50"], font=kirmizi))
    ws.conditional_formatting.add("J4", CellIsRule(
        operator="greaterThanOrEqual", formula=["70"], font=kirmizi))
    ws.conditional_formatting.add("J4", CellIsRule(
        operator="between", formula=["40", "69"], font=amber))
    ws.conditional_formatting.add("J4", CellIsRule(
        operator="lessThan", formula=["40"], font=yesil))
    for col in range(2, 14):
        harf = get_column_letter(col)
        ws.conditional_formatting.add(
            f"{harf}4", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))

    ws = wb["GIRDI"]
    ws.conditional_formatting.add("B11:B14", CellIsRule(
        operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("B11", CellIsRule(
        operator="equal", formula=["0"], font=amber))
    ws.conditional_formatting.add("F24:F35", FormulaRule(
        formula=['F24="EKSİK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F24:F35", FormulaRule(
        formula=['F24="TAM"'], font=yesil, fill=y_fill))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("G6:G9", FormulaRule(
        formula=['G6="TUTARLI"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("G6:G9", FormulaRule(
        formula=['G6="KIRIK"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("F6:F9", CellIsRule(
        operator="notEqual", formula=["0"], font=amber))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("C6:C9", CellIsRule(
        operator="greaterThan", formula=["0"], font=yesil))
    ws.conditional_formatting.add("E6:E9", FormulaRule(
        formula=['ISNUMBER(SEARCH("RİSKLİ",E6))'], font=kirmizi))
    ws.conditional_formatting.add("E6:E9", FormulaRule(
        formula=['ISNUMBER(SEARCH("UYGUN",E6))'], font=yesil))
    ws.conditional_formatting.add("E6:E9", FormulaRule(
        formula=['ISNUMBER(SEARCH("DİKKAT",E6))'], font=amber))
    ws.conditional_formatting.add("B16:B18", CellIsRule(
        operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add("B5:B11", FormulaRule(
        formula=['OR(B5="KIRIK",B5="HATALI",B5="NEGATİF",B5="BOŞ",B5="EKSİK",B5="AŞIYOR")'],
        font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B5:B11", FormulaRule(
        formula=['OR(B5="TEMİZ",B5="DOLU",B5="NORMAL",B5="TUTARLI")'],
        font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B12", CellIsRule(
        operator="greaterThanOrEqual", formula=["80"], font=yesil))
    ws.conditional_formatting.add("B12", CellIsRule(
        operator="lessThan", formula=["50"], font=kirmizi))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("G7:G55", CellIsRule(
        operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("F7:F55", FormulaRule(
        formula=['F7="YDV-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F55", FormulaRule(
        formula=['F7="YDV-004"'], font=Font(color=KRITIK, bold=True)))

    ws = wb["ORNEK_VERI"]
    for harf in ("D", "E", "F", "G", "I", "L"):
        ws.conditional_formatting.add(
            f"{harf}6:{harf}20", CellIsRule(operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("L6:L20", CellIsRule(
        operator="greaterThan", formula=["0"], font=amber))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B22", FormulaRule(
        formula=['ISNUMBER(SEARCH("RİSKLİ",B22))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B22", FormulaRule(
        formula=['ISNUMBER(SEARCH("UYGUN",B22))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B22", FormulaRule(
        formula=['ISNUMBER(SEARCH("DİKKAT",B22))'], font=amber, fill=a_fill))
    ws.conditional_formatting.add("B20", CellIsRule(
        operator="greaterThan", formula=["0"], font=amber))


def ad_tanimlari(wb):
    ad_ekle(wb, "hesapYili", "GIRDI!$B$8")
    ad_ekle(wb, "raporTarihi", "AYARLAR!$B$7")
    ad_ekle(wb, "ydv_karTl", "MOTOR!$G$20")
    ad_ekle(wb, "ydv_toplamVergi", "MOTOR!$G$28")
    ad_ekle(wb, "ydv_efektifVergi", "MOTOR!$G$29")
    ad_ekle(wb, "ydv_ciftVergi", "MOTOR!$G$30")
    ad_ekle(wb, "ydv_trEkMatrah", "MOTOR!$G$31")
    ad_ekle(wb, "ydv_kararMetni", "MOTOR!$G$33")
    ad_ekle(wb, "ydv_yapiOneri", "MOTOR!$G$34")
    ad_ekle(wb, "ydv_kararGerekce", "MOTOR!$G$35")
    ad_ekle(wb, "ydv_maddeAtifMetni", "MOTOR!$G$54")
    ad_ekle(wb, "ydv_kanitRaporu", "MOTOR!$G$55")
    ad_ekle(wb, "ydv_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "ydv_parmakIzi", "AYARLAR!$B$19")

    ad_ekle(wb, "ydv_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "ydv_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "ydv_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "ydv_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "ydv_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "ydv_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "ydv_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "ydv_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "ydv_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "ydv_vakaFark", "VAKALAR!$B$13")

    ayar_map = {
        "ydv_tolerans": 8,
        "ydv_kacirilanEsik": 9,
        "ydv_senaryoIyi": 10,
        "ydv_senaryoKotu": 11,
        "ydv_oranArtisPuan": 12,
        "ydv_tahminAlt": 13,
        "ydv_tahminUst": 14,
        "ydv_tornadoKur": 15,
        "ydv_tornadoTf": 16,
        "ydv_peArtisOran": 17,
        "ydv_yuzdelikOran": 18,
        "ydv_girisBeklenen": 21,
        "ydv_riskDusuk": 22,
        "ydv_riskOrta": 23,
        "ydv_riskYuksek": 24,
        "ydv_yilTaban": 25,
        "ydv_yuvarMax": 26,
        "ydv_bazCarpan": 27,
        "ydv_sifirCarpan": 28,
        "ydv_siraBir": 29,
        "ydv_siraIki": 30,
        "ydv_siraUc": 31,
        "ydv_tahminX": 32,
        "ydv_ofisMatrahOran": 33,
        "ydv_dagitimOran": 34,
        "ydv_tfBolum": 35,
        "ydv_ekMatrahCarpan": 36,
        "ydv_ciftVergiEsik": 37,
        "ydv_dikkatEsik": 38,
        "ydv_dikkatCiftEsik": 39,
        "ydv_subeEfektifEsik": 40,
        "vaka_kar_1": 41,
        "vaka_yerel_oran": 42,
        "vaka_tf_puan": 43,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "ydv_modulT1": (12, "C"), "ydv_modulT1Yorum": (12, "D"),
        "ydv_modulT2": (13, "C"), "ydv_modulT2Yorum": (13, "D"),
        "ydv_modulO1": (14, "C"), "ydv_modulO1Yorum": (14, "D"),
        "ydv_modulO6": (17, "C"), "ydv_modulO6Yorum": (17, "D"),
        "ydv_modulO8": (18, "C"), "ydv_modulO8Yorum": (18, "D"),
        "ydv_modulI2": (21, "C"), "ydv_modulI2Yorum": (21, "D"),
    }
    for ad, (satir, kol) in modul_map.items():
        ad_ekle(wb, ad, f"PANO!${kol}${satir}")

    ad_ekle(wb, "ListeYillar", "LISTELER!$A$6:$A$8")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$C$6:$C$7")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$6:$E$7")
    ad_ekle(wb, "ListeYapi", "LISTELER!$G$6:$G$9")


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
    varsayilan = os.path.join(kok, "YurtDisiYapilanmaVergiSimulatoru.xlsx")
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
    cikti = os.path.join(repo, "cikti", "YurtDisiYapilanmaVergiSimulatoru.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
