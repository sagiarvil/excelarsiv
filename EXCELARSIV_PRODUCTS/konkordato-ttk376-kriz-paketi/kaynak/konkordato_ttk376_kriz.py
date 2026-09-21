#!/usr/bin/env python3
"""
Konkordato + TTK 376 Kriz Paketi — üretim betiği (A1 / manda v6).
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

URUN_AD = "Konkordato + TTK 376 Kriz Paketi"
SURUM = "1.0.0"
RENK = "0F2742"
RAPOR_TARIH = date(2026, 8, 11)

# Demo: özkaynak 4.5M / sermaye tabanı 11M ≈ %40.9 → DİKKAT (1/2 altı, 1/3 üstü)
# borç/aktif 14/25 = %56
DEMO = {
    "sirket_adi": "Örnek Sanayi A.Ş.",
    "donem": "2025-ARA-06",
    "hesapYili": 2025,
    "odenmis_sermaye": 10_000_000,
    "sermaye_yedekleri": 1_000_000,
    "gecmis_yil_zararlari": 4_000_000,
    "donem_net_kar_zarar": -2_500_000,
    "toplam_varlik": 25_000_000,
    "kisa_vadeli_borc": 8_000_000,
    "uzun_vadeli_borc": 6_000_000,
    "yorum_modu": "A",
}


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: TTK 376 / konkordato kriz kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: Özkaynak oranı ve karar bozulur | Kaynak: TTK 376",
    )


def kural_cek(kural_id: str) -> str:
    """Oranları tblKurallar'dan çeker; yıl seçimi ktk_yilTaban ile (G07: adında rakam yok)."""
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-ktk_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


_YV = "MAX(0,MIN(ktk_yuvarMax,IFERROR(N(G12),0)))"


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
      "Ödenmiş sermaye, yedek, zarar ve borç girdilerini girin; TTK 376 özkaynak/sermaye "
      "oranını hesaplayıp NORMAL / DİKKAT / KRİZ kararını üretin.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Çok yıllı kural tablosu (2025/2026/2027) ile 1/2 ve 2/3 eşik güncellemesine hazır motor",
        "Özkaynak, sermaye tabanı, kayıp oranı ve eksik sermaye tutarının yan yana hesabı",
        "Baz / iyimser / kötümser + kural değişimi (eşik −1 puan) senaryoları",
        "Tebliğ tarzı altın vakalar ve KANIT_RAPORU (A4 imza)",
        "Belirsizlik beyanı: sermaye tamamlama süresi (yorum A / B)",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "CFO — özkaynak / sermaye tamamlamayı görmek",
        "Ortaklar kurulu — NORMAL / DİKKAT / KRİZ kararını almak",
        "SMMM / hukuk — imzalı kanıt raporunu dosyalamak",
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
        ("Adım 1 — Girdiler", "GIRDI sayfasında yıl, sermaye, zarar, varlık ve borçları sarı hücrelere girin."),
        ("Adım 2 — Karar", "PANO'da özkaynak oranını ve NORMAL / DİKKAT / KRİZ rozetini izleyin."),
        ("Adım 3 — Kanıt", "KANIT_RAPORU'nu yazdırıp imzalayın; VAKALAR tutarlılığını kontrol edin."),
    ]
    for i, (b, m) in enumerate(adimlar, 5):
        h(ws, i, 1, b, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        _kim(ws, f"A{i}", b, "Sırayla izleyin")
    h(ws, 9, 1, "Sık yapılan hatalar", kalin=True, yazi=KRITIK)
    for i, m in enumerate([
        "Geçmiş yıl zararını pozitif yerine net özkaynak olarak yazmak",
        "Sermaye yedeklerini ödenmiş sermayeye eklememek",
        "Hesap yılını kural tablosuyla uyumsuz seçmek",
    ], 10):
        h(ws, i, 1, "• " + m, yazi=KRITIK)
    h(ws, 14, 1,
      "Bu dosya karar destek aracıdır; konkordato başvurusu veya hukuki görüş yerine geçmez.",
      yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 72, "B": 70})
    for r in (10, 11, 12):
        h(ws, r, 1, ws.cell(r, 1).value, yazi=KRITIK, kaydir=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", "B08948", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Şirket ve bilanço girdileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücreler manuel giriştir. Özkaynak ve karar formülle üretilir.", yazi=GRİ, kaydir=True)

    alanlar = [
        (6, "Şirket Adı", DEMO["sirket_adi"], None, "sirket_adi",
         "Şirket unvanını yazın."),
        (7, "Bilanço Dönemi", DEMO["donem"], None, "donem",
         "Ara dönem veya yıl sonu bilanço referansını yazın."),
        (8, "Hesap Yılı", DEMO["hesapYili"], CATI, "hesapYili",
         "2025, 2026 veya 2027 seçin."),
        (9, "Ödenmiş Sermaye [₺]", DEMO["odenmis_sermaye"], TL, "odenmis_sermaye",
         "Ödenmiş sermaye tutarını TL girin."),
        (10, "Sermaye Yedekleri [₺]", DEMO["sermaye_yedekleri"], TL, "sermaye_yedekleri",
         "Sermaye yedekleri toplamını TL girin."),
        (11, "Geçmiş Yıl Zararları [₺]", DEMO["gecmis_yil_zararlari"], TL, "gecmis_yil_zararlari",
         "Birikmiş geçmiş yıl zararını pozitif tutar olarak girin."),
        (12, "Dönem Net Kâr/Zarar [₺]", DEMO["donem_net_kar_zarar"], TL, "donem_net_kar_zarar",
         "Dönem net kârı pozitif, zararı negatif girin."),
        (13, "Toplam Varlık [₺]", DEMO["toplam_varlik"], TL, "toplam_varlik",
         "Bilanço toplam varlık tutarını TL girin."),
        (14, "Kısa Vadeli Borç [₺]", DEMO["kisa_vadeli_borc"], TL, "kisa_vadeli_borc",
         "Kısa vadeli yabancı kaynakları TL girin."),
        (15, "Uzun Vadeli Borç [₺]", DEMO["uzun_vadeli_borc"], TL, "uzun_vadeli_borc",
         "Uzun vadeli yabancı kaynakları TL girin."),
        (16, "Yorum Modu (A/B)", DEMO["yorum_modu"], None, "yorum_modu",
         "Sermaye tamamlama süresi yorumu A (standart) veya B (sıkı) seçin."),
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

    h(ws, 18, 1, "Özkaynak (ayna) [₺]", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 2, "=IFERROR(ktk_ozKaynak,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Giriş doluluk (kalite)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(IF(OR(COUNTA(B6:B16)=0,ktk_girisBeklenen=0),0,ROUND(COUNTA(B6:B16)/ktk_girisBeklenen*100,0)),0)',
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
        ("odenmis_sermaye", DEMO["odenmis_sermaye"], "TL", "Evet", "GIRDI"),
        ("sermaye_yedekleri", DEMO["sermaye_yedekleri"], "TL", "Evet", "GIRDI"),
        ("gecmis_yil_zararlari", DEMO["gecmis_yil_zararlari"], "TL", "Evet", "GIRDI"),
        ("donem_net_kar_zarar", DEMO["donem_net_kar_zarar"], "TL", "Evet", "GIRDI"),
        ("toplam_varlik", DEMO["toplam_varlik"], "TL", "Evet", "GIRDI"),
        ("kisa_vadeli_borc", DEMO["kisa_vadeli_borc"], "TL", "Evet", "GIRDI"),
        ("uzun_vadeli_borc", DEMO["uzun_vadeli_borc"], "TL", "Evet", "GIRDI"),
        ("yorum_modu", DEMO["yorum_modu"], "A/B", "Evet", "GIRDI"),
    ]
    for i, (a, d, b, z, k) in enumerate(ornek_satir, 24):
        ws.cell(i, 1).value = a
        ws.cell(i, 2).value = d
        if a in ("odenmis_sermaye", "sermaye_yedekleri", "gecmis_yil_zararlari",
                 "donem_net_kar_zarar", "toplam_varlik", "kisa_vadeli_borc", "uzun_vadeli_borc"):
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
              baslik="Yorum", mesaj="Yorum modunda A veya B seçin.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B seçin.", bos=False)
    dogrulama(ws, "list", "ListeYillar", "B26",
              baslik="Hesap Yılı", mesaj="Tablo satırında yıl seçin.",
              hata_baslik="Geçersiz yıl", hata_mesaj="Listeden seçin.")
    dogrulama(ws, "list", "ListeYorumAB", "B34",
              baslik="Yorum", mesaj="A veya B.",
              hata_baslik="Geçersiz", hata_mesaj="A veya B.")
    for aralik, baslik, mesaj, f2 in [
        ("B9", "Ödenmiş Sermaye", "0 ile 1e12 arasında TL.", "1000000000000"),
        ("B10", "Sermaye Yedekleri", "0 ile 1e12 arasında TL.", "1000000000000"),
        ("B11", "Geçmiş Yıl Zararları", "0 ile 1e12 arasında TL.", "1000000000000"),
        ("B13", "Toplam Varlık", "0 ile 1e12 arasında TL.", "1000000000000"),
        ("B14", "Kısa Vadeli Borç", "0 ile 1e12 arasında TL.", "1000000000000"),
        ("B15", "Uzun Vadeli Borç", "0 ile 1e12 arasında TL.", "1000000000000"),
        ("B27", "Ödenmiş Sermaye", "Tablo: TL tutar.", "1000000000000"),
        ("B28", "Sermaye Yedekleri", "Tablo: TL tutar.", "1000000000000"),
        ("B29", "Geçmiş Zarar", "Tablo: TL tutar.", "1000000000000"),
        ("B31", "Toplam Varlık", "Tablo: TL tutar.", "1000000000000"),
        ("B32", "Kısa Borç", "Tablo: TL tutar.", "1000000000000"),
        ("B33", "Uzun Borç", "Tablo: TL tutar.", "1000000000000"),
    ]:
        dogrulama(ws, "decimal", "0", aralik,
                  baslik=baslik, mesaj=mesaj,
                  hata_baslik="Geçersiz tutar", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2=f2)
    dogrulama(ws, "decimal", "-1000000000000", "B12",
              baslik="Dönem Net", mesaj="-1e12 ile 1e12 arasında TL.",
              hata_baslik="Geçersiz", hata_mesaj="Sınır dışı.",
              isaret="between", f2="1000000000000")
    dogrulama(ws, "decimal", "-1000000000000", "B30",
              baslik="Dönem Net", mesaj="Tablo: dönem net.",
              hata_baslik="Geçersiz", hata_mesaj="Sınır dışı.",
              isaret="between", f2="1000000000000")

    giris_hucreleri(ws, 6, 16, [2])
    giris_hucreleri(ws, 24, 34, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 40, "B": 28, "C": 12, "D": 12, "E": 12, "F": 12})
    alt_bant(ws, 1024, "Sarı alanlar giriş; özkaynak ve kalite formülleri kilitlidir.")


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Mevzuat kural tablosu (yıl yan yana)", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 4, 1, "Oranlar MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["KTK-001", "TTK 376/1", "esik_yari", "oran < yarı", "esik_yari",
         0.50, 0.50, 0.50, "01.01.2025", "Sermaye ve yedeklerin yarısı karşılıksız"],
        ["KTK-002", "TTK 376/2", "esik_ucte_bir", "oran < 1/3", "esik_ucte_bir",
         0.333333, 0.333333, 0.333333, "01.01.2025", "2/3 kayıp sonrası kalan 1/3 eşiği"],
        ["KTK-003", "TTK 376/3", "borca_batik", "borç >= varlık", "borca_batik",
         1.0, 1.0, 1.0, "01.01.2025", "Borca batıklık: borç/aktif oranı tavanı"],
        ["KTK-004", "İç politika", "dikkat_borc_aktif", "borç/aktif uyarı", "dikkat_borc_aktif",
         0.75, 0.75, 0.75, "01.01.2025", "Borç/aktif dikkat eşiği"],
        ["KTK-005", "Genel", "yuvarlama", "her hesap", "yuvarlama_ondalik",
         2, 2, 2, "01.01.2025", "Uygulama notu"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = YÜZDE if k in (6, 7, 8) and isinstance(v, float) and v <= 1 else (
                CATI if k in (6, 7, 8) else None)
            if k in (6, 7, 8) and s[0] in ("KTK-005",):
                sayi = CATI
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                _kim(ws, f"{get_column_letter(k)}{i}", s[2], "Yıl bazlı parametre; kaynak kolonuna bakın")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-ktk_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)

    h(ws, 14, 1, "Kural-yıl matrisi özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Yarı eşik (aktif yıl)", yazi="333333")
    h(ws, 15, 2, kural_cek("KTK-001"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "1/3 eşik (aktif yıl)", yazi="333333")
    h(ws, 16, 2, kural_cek("KTK-002"), sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Matris metni", yazi="333333")
    h(ws, 17, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" | ",TRUE,"KTK-001=",TEXT(B15,"0.00"),"KTK-002=",TEXT(B16,"0.0%"),'
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
    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif mevzuat yılı", "KTK-005")  # G7
    ekle("Yarı eşik", km("KTK-001"), "oran", "TTK 376/1 1/2 eşiği", "KTK-001")  # G8
    ekle("1/3 eşik", km("KTK-002"), "oran", "TTK 376/2 kalan 1/3", "KTK-002")  # G9
    ekle("Borca batık tavan", km("KTK-003"), "oran", "Borç/aktif tavan", "KTK-003")  # G10
    ekle("Dikkat borç/aktif", km("KTK-004"), "oran", "Borç/aktif uyarı", "KTK-004")  # G11
    ekle("Yuvarlama ondalık", km("KTK-005"), "adet", "Yuvarlama basamağı", "KTK-005")  # G12
    ekle("Ödenmiş sermaye", "=GIRDI!B9", "TL", "Ödenmiş sermaye", "KTK-001")  # G13
    ekle("Sermaye yedekleri", "=GIRDI!B10", "TL", "Sermaye yedekleri", "KTK-001")  # G14
    ekle("Geçmiş yıl zararları", "=GIRDI!B11", "TL", "Birikmiş zarar", "KTK-001")  # G15
    ekle("Dönem net kâr/zarar", "=GIRDI!B12", "TL", "Dönem sonucu", "KTK-001")  # G16
    ekle("Toplam varlık", "=GIRDI!B13", "TL", "Aktif toplamı", "KTK-003")  # G17
    ekle("Kısa vadeli borç", "=GIRDI!B14", "TL", "KV yabancı kaynak", "KTK-003")  # G18
    ekle("Uzun vadeli borç", "=GIRDI!B15", "TL", "UV yabancı kaynak", "KTK-003")  # G19
    ekle("Yarı eşik efektif",
         '=IF(GIRDI!B16="B",MAX(0,G8-ktk_oranArtisPuan),G8)',
         "oran", "Yorum A/B yarı eşik", "KTK-001")  # G20
    ekle("Sermaye tabanı", "=G13+G14", "TL", "Sermaye + yedek", "KTK-001")  # G21
    ekle("Özkaynak", "=G13+G14-G15+G16", "TL", "Özkaynak bileşeni", "KTK-001")  # G22
    ekle("Özkaynak / sermaye", "=IF(G21=0,0,G22/G21)", "oran", "TTK 376 oran", "KTK-001")  # G23
    ekle("Kayıp oranı", "=MAX(0,1-G23)", "oran", "1 − oran", "KTK-002")  # G24
    ekle("Eksik sermaye (yarıya)",
         f"=ROUND(MAX(0,G21*G20-G22),{_YV})",
         "TL", "Yarı eşiğe tamamlamak", "KTK-001")  # G25
    ekle("Eksik sermaye (1/3'e)",
         f"=ROUND(MAX(0,G21*G9-G22),{_YV})",
         "TL", "1/3 eşiğe tamamlamak", "KTK-002")  # G26
    ekle("Toplam borç", "=G18+G19", "TL", "KV+UV borç", "KTK-003")  # G27
    ekle("Özkaynak (pano)", "=G22", "TL", "Özkaynak ayna", "KTK-001")  # G28
    ekle("Eksik (karar)", "=G25", "TL", "Eksik sermaye ayna", "KTK-001")  # G29
    ekle("Kritik açık", "=MAX(G25,G26)", "TL", "En yüksek eksik", "KTK-002")  # G30
    ekle("Borç / aktif", "=IF(G17=0,0,G27/G17)", "oran", "Borç yoğunluğu", "KTK-003")  # G31
    ekle("Varlık − borç", "=G17-G27", "TL", "Net aktif", "KTK-003")  # G32
    ekle("Karar kodu",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERI_YOK",'
         'IF(OR(G22<0,G31>=G10,G23<G9),"KRIZ",'
         'IF(OR(G23<G20,G31>=G11),"DIKKAT","NORMAL")))',
         "kod", "Karar kodu", "KTK-005")  # G33
    ekle("Karar metni",
         '=IF(G33="VERI_YOK","VERİ YOK",'
         'IF(G33="KRIZ","KRİZ",'
         'IF(G33="DIKKAT","DİKKAT","NORMAL")))',
         "metin", "Kullanıcıya karar", "KTK-005")  # G34
    ekle("Gerekçe",
         '=IF(G33="VERI_YOK","Giriş tablosu boş — hesap yapılamaz.",'
         'IF(G33="KRIZ","Özkaynak 1/3 altına düşmüş, borca batık veya negatif; genel kurul / konkordato yolu değerlendirin.",'
         'IF(G33="DIKKAT","Özkaynak yarı eşiğin altında veya borç/aktif yüksek; sermaye tamamlama planı hazırlayın.",'
         '"Özkaynak/sermaye oranı eşiklerin üstünde; TTK 376 zorunlu önlem gerekmiyor.")))',
         "metin", "Gerekçe cümlesi", "KTK-005")  # G35
    ekle("Özkaynak / varlık", "=IF(G17=0,0,G22/G17)", "oran", "Özkaynak yoğunluğu", "KTK-001")  # G36
    ekle("Eksik / sermaye", "=IF(G21=0,0,G30/G21)", "oran", "Eksik oranı", "KTK-001")  # G37
    ekle("Güven skoru",
         "=IFERROR(ROUND(IF(OR(COUNTA(GIRDI!B6:B16)=0,ktk_girisBeklenen=0),0,"
         "COUNTA(GIRDI!B6:B16)/ktk_girisBeklenen*100),0),0)",
         "puan", "Giriş bütünlüğü", "KTK-005")  # G38
    ekle("Risk skoru",
         '=IF(G33="VERI_YOK",0,IF(G33="NORMAL",ktk_riskDusuk,'
         'IF(G33="DIKKAT",ktk_riskOrta,ktk_riskYuksek)))',
         "puan", "Kriz riski", "KTK-005")  # G39
    ekle("Senaryo iyimser özkaynak", "=G22*ktk_senaryoIyi", "TL", "İyimser özkaynak", "KTK-001")  # G40
    ekle("Senaryo kötümser özkaynak", "=G22*ktk_senaryoKotu", "TL", "Kötümser özkaynak", "KTK-001")  # G41
    ekle("İyimser oran", "=IF(G21=0,0,G40/G21)", "oran", "İyimser oran", "KTK-001")  # G42
    ekle("Kötümser oran", "=IF(G21=0,0,G41/G21)", "oran", "Kötümser oran", "KTK-001")  # G43
    ekle("M-SEN yarı−", "=MAX(0,G8-ktk_oranArtisPuan)", "oran", "Yarı − delta", "KTK-001")  # G44
    ekle("M-SEN eksik",
         f"=ROUND(MAX(0,G21*G44-G22),{_YV})",
         "TL", "M-SEN eksik sermaye", "KTK-001")  # G45
    ekle("M-SEN açık", "=G45", "TL", "M-SEN farkı", "KTK-001")  # G46
    ekle("Tahmin üst", "=G30*ktk_tahminUst", "TL", "Tahmin üst bant", "KTK-005")  # G47
    ekle("Tahmin alt", "=G30*ktk_tahminAlt", "TL", "Tahmin alt bant", "KTK-005")  # G48
    ekle("Tornado: sermaye etkisi",
         f"=ABS(ROUND(MAX(0,(G21*ktk_tornadoMatrah)*G20-G22),{_YV})-G25)",
         "TL", "Sermaye ± etki", "KTK-001")  # G49
    ekle("Tornado: zarar etkisi",
         f"=ABS(ROUND(MAX(0,G21*G20-(G13+G14-G15*ktk_tornadoMevcut+G16)),{_YV})-G25)",
         "TL", "Zarar ± etki", "KTK-001")  # G50
    ekle("Tornado: eşik etkisi",
         f"=ABS(ROUND(MAX(0,G21*(G8+ktk_oranArtisPuan)-G22),{_YV})-G25)",
         "TL", "Eşik ± etki", "KTK-001")  # G51
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"KTK-001→1/2 eşik","KTK-002→1/3 eşik","KTK-003→borca batık","KTK-004→borç/aktif")',
         "metin", "Kanıt atıfları", "KTK-001")  # G52
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Oran=",TEXT(G23,"0.0%"),"Özkaynak=",TEXT(G22,"₺ #,##0"),'
         '"Eksik=",TEXT(G30,"₺ #,##0"),"Borç/Aktif=",TEXT(G31,"0.0%"))',
         "metin", "Rapor özeti", "KTK-005")  # G53

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
    h(ws, 5, 2, "Özkaynak Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Özkaynak", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "Eksik Sermaye", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    senaryolar = [
        ("İyimser", "=IFERROR(ktk_senaryoIyi+N(GIRDI!B9)*ktk_sifirCarpan,0)",
         "=IFERROR(MOTOR!G40,0)",
         "=IFERROR(ROUND(MAX(0,MOTOR!G21*MOTOR!G20-MOTOR!G40),2),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(D6=0,"NORMAL","DİKKAT"))'),
        ("Baz", "=IFERROR(ktk_bazCarpan+N(GIRDI!B9)*ktk_sifirCarpan,ktk_bazCarpan)",
         "=IFERROR(ktk_ozKaynak,0)",
         "=IFERROR(ktk_eksikSermaye,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IFERROR(ktk_kararMetni,"-"))'),
        ("Kötümser", "=IFERROR(ktk_senaryoKotu+N(GIRDI!B9)*ktk_sifirCarpan,0)",
         "=IFERROR(MOTOR!G41,0)",
         "=IFERROR(ROUND(MAX(0,MOTOR!G21*MOTOR!G20-MOTOR!G41),2),0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(C8<MOTOR!G21*MOTOR!G9,"KRİZ","DİKKAT"))'),
        ("M-SEN −1 puan", "=IFERROR(ktk_bazCarpan+N(GIRDI!B9)*ktk_sifirCarpan,ktk_bazCarpan)",
         "=IFERROR(ktk_ozKaynak,0)",
         "=IFERROR(MOTOR!G46,0)",
         '=IF(COUNTA(tblGirdi[Alan Anahtarı])=0,"VERİ YOK",IF(D9=0,"NORMAL","DİKKAT"))'),
    ]
    for i, (ad, carp, asg, fark, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, asg, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, fark, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant genişliği (iyimser−kötümser eksik)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(D8-D6,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser eksik ",TEXT(D6,"₺ #,##0")," · Baz ",TEXT(D7,"₺ #,##0"),'
      '" · Kötümser ",TEXT(D8,"₺ #,##0")," · Bant ",TEXT(B11,"₺ #,##0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:H12")

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 3, "Sıra", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Tornado sermaye etkisi", yazi="333333")
    h(ws, 16, 2, "=IFERROR(MOTOR!G49,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Tornado zarar etkisi", yazi="333333")
    h(ws, 17, 2, "=IFERROR(MOTOR!G50,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Tornado eşik etkisi", yazi="333333")
    h(ws, 18, 2, "=IFERROR(MOTOR!G51,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 3,
      '=IFERROR(IF(B16>=MAX(B16:B18),ktk_siraBir,IF(B16>=MIN(B16:B18)+ABS(B16-B17),ktk_siraIki,ktk_siraUc)),ktk_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 3,
      '=IFERROR(IF(B17>=MAX(B16:B18),ktk_siraBir,IF(B17=MEDIAN(B16:B18),ktk_siraIki,ktk_siraUc)),ktk_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 3,
      '=IFERROR(IF(B18>=MAX(B16:B18),ktk_siraBir,IF(B18=MEDIAN(B16:B18),ktk_siraIki,ktk_siraUc)),ktk_siraUc)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Duyarlılık sıra özeti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 2,
      '=IFERROR(CONCATENATE("1)",INDEX(A16:A18,MATCH(ktk_siraBir,C16:C18,0)),'
      '" 2)",INDEX(A16:A18,MATCH(ktk_siraIki,C16:C18,0))),"-")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Duyarlılık yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek etki ",TEXT(MAX(B16:B18),"₺ #,##0")," TL — öncelik bu değişkende")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 22, 1, "M-SEN kural değişimi eksik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=D9", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "M-SEN yorumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2,
      '=_xlfn.TEXTJOIN("; ",TRUE,"Yarı eşik −1 puan senaryosunda eksik sermaye ",TEXT(D9,"₺ #,##0")," TL")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 25, 1, "Tahmin aralığı (eksik)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=MOTOR!G48", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 3, "=IFERROR(MOTOR!G30,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
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

    h(ws, 5, 7, "EksikDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([550000, 1000000, 1900000, 890000], 6):
        h(ws, i, 7, v, sayi=TL, yazi=GRİ)
    h(ws, 15, 5, "EtkiDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([550000, 400000, 110000], 16):
        h(ws, i, 5, v, sayi=TL, yazi=GRİ)

    h(ws, 27, 4, "TrendDemo", kalin=True, yazi=KOYU_LACIVERT)
    for i, v in enumerate([550000, 1000000, 1900000, 890000], 28):
        h(ws, i, 4, v, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Senaryo Eksik Sermaye Karşılaştırması"
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
    g3.title = "Senaryo Eksik Sermaye Trendi"
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

    # V-001: sermaye 100000 * yarı 0.50 = 50000
    # V-002: sermaye 200000 * 1/3 ≈ 66666.6
    # V-003: V-001 + V-002
    # V-004: varlık 150000 * borç/aktif dikkat 0.75 = 112500
    vakalar_data = [
        ("V-001", "Yarı eşik × sermaye",
         "sermaye=vaka_matrah_1; esik=KTK-001",
         50000,
         "=IFERROR(ROUND(vaka_matrah_1*INDEX(tblKurallar[deger_2025],MATCH(\"KTK-001\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=ktk_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "TTK 376/1 yarı eşik"),
        ("V-002", "1/3 eşik × sermaye",
         "sermaye=vaka_matrah_2; esik=KTK-002",
         66666.6,
         "=IFERROR(ROUND(vaka_matrah_2*INDEX(tblKurallar[deger_2025],MATCH(\"KTK-002\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=ktk_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "TTK 376/2 1/3 eşik"),
        ("V-003", "Birikimli yarı + 1/3",
         "V-001 + V-002",
         116666.6,
         "=IFERROR(ROUND(vaka_matrah_1*INDEX(tblKurallar[deger_2025],MATCH(\"KTK-001\",tblKurallar[kural_id],0))"
         "+vaka_matrah_2*INDEX(tblKurallar[deger_2025],MATCH(\"KTK-002\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=ktk_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "Birikim yorum A"),
        ("V-004", "Varlık × borç/aktif dikkat",
         "varlik=vaka_tavan; oran=KTK-004",
         112500,
         "=IFERROR(ROUND(vaka_tavan*INDEX(tblKurallar[deger_2025],MATCH(\"KTK-004\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=ktk_tolerans,"TUTARLI","KIRIK"),"KIRIK")',
         "KTK-004 borç/aktif"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v,
              yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if k in (4, 5, 6):
                ws.cell(i, k).number_format = TL

    formuller = {
        "fark": "=IF(tblVakalar[[#This Row],[vaka_id]]=\"\",\"\",IFERROR(tblVakalar[[#This Row],[beklenen_sonuc]]-tblVakalar[[#This Row],[hesaplanan]],0))",
        "durum": '=IF(tblVakalar[[#This Row],[vaka_id]]="","",IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=ktk_tolerans,"TUTARLI","KIRIK"),"KIRIK"))',
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
    for i, (b, he) in enumerate([(50000, 50000), (66666.6, 66666.6), (116666.6, 116666.6), (112500, 112500)], 6):
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
    h(ws, 3, 1, "KANIT RAPORU — Konkordato + TTK 376 Kriz", kalin=True, boyut=14, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:F3")
    h(ws, 4, 1, "Rapor tarihi", yazi="333333")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH, yazi=DIS_REF_YEŞIL)
    h(ws, 5, 1, "Şirket", yazi="333333")
    h(ws, 5, 2, "=GIRDI!B6", yazi=DIS_REF_YEŞIL)
    h(ws, 6, 1, "Dönem / yıl", yazi="333333")
    h(ws, 6, 2, '=CONCATENATE(GIRDI!B7," / ",hesapYili)', yazi=DIS_REF_YEŞIL)

    h(ws, 8, 1, "Girdi özeti", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, (et, form, sayi) in enumerate([
        ("Ödenmiş sermaye", "=GIRDI!B9", TL),
        ("Sermaye yedekleri", "=GIRDI!B10", TL),
        ("Geçmiş yıl zararları", "=GIRDI!B11", TL),
        ("Dönem net kâr/zarar", "=GIRDI!B12", TL),
        ("Toplam varlık", "=GIRDI!B13", TL),
        ("Kısa vadeli borç", "=GIRDI!B14", TL),
        ("Uzun vadeli borç", "=GIRDI!B15", TL),
    ], 9):
        h(ws, i, 1, et, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)

    h(ws, 17, 1, "Hesap zinciri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 18, 1, "Özkaynak", yazi="333333")
    h(ws, 18, 2, "=ktk_ozKaynak", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 19, 1, "Eksik sermaye (karar)", yazi="333333")
    h(ws, 19, 2, "=ktk_eksikSermaye", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "KARAR", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=ktk_kararMetni", kalin=True, boyut=12, yazi=DIS_REF_YEŞIL)
    h(ws, 21, 1, "Gerekçe", yazi="333333")
    h(ws, 21, 2, "=ktk_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B21:F21")

    h(ws, 23, 1, "Madde atıfları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "=ktk_maddeAtifMetni", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B23:F23")

    h(ws, 25, 1, "Kanıt gövdesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=ktk_kanitRaporu", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B25:F25")

    h(ws, 27, 1, "Parmak izi (SHA-256 kısaltma)", yazi="333333")
    h(ws, 27, 2, "=ktk_parmakIzi", yazi=GRİ)

    h(ws, 29, 1, "Hazırlayan (imza)", yazi="333333")
    h(ws, 29, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B29", "İmza", "Raporu hazırlayan kişinin adını yazın.")
    h(ws, 30, 1, "Onaylayan (imza)", yazi="333333")
    h(ws, 30, 2, "", zemin=GIRIS_SARI)
    _kim(ws, "B30", "Onay", "Onaylayan kişinin adını yazın.")

    h(ws, 32, 1,
      "Uyarı: Bu çıktı karar destektir; kesin konkordato / TTK 376 sonucu veya hukuki görüş değildir. "
      f"Sürüm {SURUM}.",
      yazi=GRİ, boyut=9, kaydir=True)
    ws.merge_cells("A32:F32")

    baski_hazirla(ws, "A1:F33", f"{URUN_AD} · {SURUM}")
    ws.page_setup.orientation = "portrait"
    genislik(ws, {"A": 28, "B": 22, "C": 14, "D": 14, "E": 14, "F": 14})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim paneli — TTK 376 / Konkordato", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpiler = [
        (2, "Oran", "=IFERROR(MOTOR!G23,0)", YÜZDE),
        (3, "Özkaynak", "=IFERROR(ktk_ozKaynak,0)", TL),
        (4, "Sermaye", "=IFERROR(MOTOR!G21,0)", TL),
        (5, "Eksik", "=IFERROR(ktk_eksikSermaye,0)", TL),
        (6, "Güven", "=IFERROR(MOTOR!G38,0)", CATI),
        (7, "Risk", "=IFERROR(MOTOR!G39,0)", CATI),
        (8, "Yarı eşik", "=IFERROR(MOTOR!G8,0)", YÜZDE),
        (9, "1/3 eşik", "=IFERROR(MOTOR!G9,0)", YÜZDE),
        (10, "Borç/Aktif", "=IFERROR(MOTOR!G31,0)", YÜZDE),
        (11, "M-SEN", "=IFERROR(ktk_kuralDegisimSenaryo,0)", TL),
        (12, "Tahmin", "=IFERROR(SENARYO!B32,0)", TL),
        (13, "Vaka", "=IFERROR(ktk_vakaDurum,\"-\")", None),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=9, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center", boyut=11)

    h(ws, 6, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=ktk_kararMetni", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 7, 2, "=ktk_kararGerekce", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B7:H7")

    h(ws, 9, 1, "Analitik modüller", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 10, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Özkaynak / sermaye oranı",
         "=IFERROR(MOTOR!G23,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Oran ",TEXT(C11,"0.0%")," — yarı ve 1/3 eşiklerle karşılaştırın")'),
        ("T2", "Eksik sermaye tutarı",
         "=IFERROR(ktk_eksikSermaye,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Eksik sermaye ",TEXT(C12,"₺ #,##0")," TL")'),
        ("O1", "Senaryo eksik sapması",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,STDEV.P(SENARYO!D6:D9)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo eksik sapması ",TEXT(C13,"₺ #,##0")," TL")'),
        ("O2", "Tornado max etki",
         "=IFERROR(ktk_duyarlilikSira,\"-\")",
         "=IFERROR(ktk_yorumDuyarlilik,\"-\")"),
        ("O3", "Senaryo bant",
         "=IFERROR(ktk_senaryoKarsilastirma,0)",
         "=IFERROR(ktk_yorumSenaryo,\"-\")"),
        ("O6", "Borç / aktif",
         "=IFERROR(MOTOR!G31,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Borç/aktif ",TEXT(C16,"0.0%"))'),
        ("O8", "Kalite skoru",
         "=IFERROR(KONTROLLER!B12,0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalite ",TEXT(C17,"0")," puan")'),
        ("M", "M-SEN eksik",
         "=IFERROR(ktk_kuralDegisimSenaryo,0)",
         "=IFERROR(ktk_yorumKuralDegisim,\"-\")"),
        ("I1", "Tahmin aralık",
         "=IFERROR(ktk_tahminAralik,0)",
         "=IFERROR(ktk_yorumTahmin,\"-\")"),
        ("I2", "Senaryo eksik P90",
         "=IFERROR(IF(COUNT(SENARYO!D6:D9)<2,0,_xlfn.PERCENTILE.INC(SENARYO!D6:D9,ktk_yuzdelikOran)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"P90 eksik ",TEXT(C20,"₺ #,##0")," TL")'),
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
        ("Ödenmiş sermaye", 10_000_000),
        ("Sermaye yedekleri", 1_000_000),
        ("Geçmiş yıl zararları", 4_000_000),
        ("Özkaynak", 4_500_000),
        ("Eksik (yarıya)", 1_000_000),
    ], 24):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)

    h(ws, 23, 4, "Kalem", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 5, "Tutar", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Özkaynak", 4_500_000),
        ("Sermaye tabanı", 11_000_000),
        ("İyimser özkaynak", 4_950_000),
        ("Kötümser özkaynak", 3_825_000),
        ("M-SEN eksik", 890_000),
    ], 24):
        h(ws, i, 4, ad, yazi=GRİ)
        h(ws, i, 5, sabit, sayi=TL, yazi=GRİ)

    g1 = BarChart()
    g1.type = "col"
    g1.title = "Sermaye Bileşenleri"
    g1.add_data(Reference(ws, min_col=2, min_row=23, max_row=28), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=24, max_row=28))
    g1.height, g1.width = 8, 12
    ws.add_chart(g1, "G9")

    g2 = BarChart()
    g2.type = "col"
    g2.title = "Özkaynak / Sermaye / Senaryo"
    g2.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g2.height, g2.width = 8, 12
    ws.add_chart(g2, "G23")

    g3 = LineChart()
    g3.title = "Özkaynak Çizgisi"
    g3.add_data(Reference(ws, min_col=5, min_row=23, max_row=28), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=4, min_row=24, max_row=28))
    g3.height, g3.width = 7, 12
    ws.add_chart(g3, "P9")

    h(ws, 30, 1, "Metrik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 2, "Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, sabit) in enumerate([
        ("Özkaynak", 4_500_000),
        ("Eksik", 1_000_000),
        ("Sermaye", 11_000_000),
    ], 31):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, sabit, sayi=TL, yazi=GRİ)
    g4 = BarChart()
    g4.type = "col"
    g4.title = "KPI Özkaynak / Eksik / Sermaye"
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
        ("Negatif ödenmiş sermaye?",
         '=IF(GIRDI!B9<0,"NEGATİF","TEMİZ")'),
        ("Hesap yılı geçerli mi?",
         '=IFERROR(IF(AND(N(hesapYili)>N(ktk_yilTaban),N(hesapYili)<=N(ktk_yilTaban)+3),"TEMİZ","HATALI"),"HATALI")'),
        ("Sermaye tabanı sıfır mı?",
         '=IF(IFERROR(MOTOR!G21,0)=0,"SIFIR","NORMAL")'),
        ("Vaka kırık mı?",
         '=IFERROR(IF(COUNTIF(tblVakalar[durum],"KIRIK")>0,"KIRIK","TUTARLI"),"KIRIK")'),
        ("Eksik eşiği aşıyor mu?",
         '=IF(IFERROR(MOTOR!G30,0)>ktk_kacirilanEsik+N(GIRDI!B9)*ktk_sifirCarpan,"AŞIYOR","NORMAL")'),
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
    h(ws, 14, 2, "=IFERROR(ktk_girisBeklenen+COUNTA(GIRDI!B6:B16)*0,0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Dolu giriş", yazi="333333")
    h(ws, 15, 2, "=COUNTA(GIRDI!B6:B16)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Kalite skoru (modül)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=IFERROR(ROUND(IF(OR(B14=0,B14=\"\"),0,B15/B14*100),0),0)", sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)

    h(ws, 17, 1, "Özkaynak kontrol", yazi="333333")
    h(ws, 17, 2, "=IFERROR(ktk_ozKaynak,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Eksik (karar) kontrol", yazi="333333")
    h(ws, 18, 2, "=IFERROR(ktk_eksikSermaye,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Karar kontrol", yazi="333333")
    h(ws, 19, 2, "=IFERROR(ktk_kararMetni,\"-\")", yazi=DIS_REF_YEŞIL)

    genislik(ws, {"A": 40, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek senaryo kütüphanesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.merge_cells("A3:I3")
    h(ws, 4, 1, "Bu sayfa demo değerleri saklar; GIRDI'ye kopyalanabilir.", yazi=GRİ)
    sutunlar = ["Senaryo", "Ödenmiş", "Yedek", "Geçmiş Zarar", "Dönem Net", "Varlık",
                "KV Borç", "UV Borç", "Özkaynak"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Özkaynak": (
            "=IF(tblOrnek[[#This Row],[Senaryo]]=\"\",\"\","
            "IFERROR(tblOrnek[[#This Row],[Ödenmiş]]+tblOrnek[[#This Row],[Yedek]]"
            "-tblOrnek[[#This Row],[Geçmiş Zarar]]+tblOrnek[[#This Row],[Dönem Net]],0))"
        ),
    }
    tablo_ekle(ws, "tblOrnek", "A5:I1004", sutunlar, formuller)
    ornekler = [
        ("Demo ana", 10_000_000, 1_000_000, 4_000_000, -2_500_000, 25_000_000, 8_000_000, 6_000_000),
        ("Kriz", 5_000_000, 200_000, 3_500_000, -1_200_000, 12_000_000, 7_000_000, 5_000_000),
        ("Normal", 8_000_000, 2_000_000, 500_000, 800_000, 30_000_000, 4_000_000, 3_000_000),
        ("Yarı eşiğe yakın", 6_000_000, 500_000, 2_800_000, -200_000, 18_000_000, 6_000_000, 4_000_000),
        ("Borca batık", 3_000_000, 100_000, 2_000_000, -800_000, 8_000_000, 5_000_000, 4_000_000),
    ]
    for i, row in enumerate(ornekler, 6):
        for k, v in enumerate(row, 1):
            ws.cell(i, k).value = v
            if k >= 2:
                ws.cell(i, k).number_format = TL
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
        ws.cell(i, 1).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
        ws.cell(i, 1).protection = Protection(locked=False)
        for kolon in range(1, 9):
            _kim(ws, f"{get_column_letter(kolon)}{i}", sutunlar[kolon - 1], "Örnek senaryo değeri")

    for col, ad in [("B", "Ödenmiş"), ("C", "Yedek"), ("D", "Geçmiş Zarar"),
                    ("E", "Dönem Net"), ("F", "Varlık"), ("G", "KV Borç"), ("H", "UV Borç")]:
        lo = "-1000000000000" if col == "E" else "0"
        dogrulama(ws, "decimal", lo, f"{col}6:{col}1004",
                  baslik=ad, mesaj=f"{ad} değerini girin.",
                  hata_baslik="Geçersiz", hata_mesaj="Sınır dışı değer.",
                  isaret="between", f2="1000000000000")

    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6, 7, 8])
    h(ws, 5, 12, "OzDemo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, v in enumerate([4500000, 500000, 10300000, 3500000, 300000], 6):
        h(ws, i, 12, v, sayi=TL, yazi=GRİ)
    g = BarChart()
    g.type = "col"
    g.title = "Örnek Senaryo Özkaynak"
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
        ("ktk_tolerans", 0.01, "TL", "Vaka tutarlılık toleransı", "Uygulama notu", "01.01.2025", "11.08.2026"),
        ("ktk_kacirilanEsik", 50_000, "TL", "PANO fark uyarı eşiği", "İç politika", "01.01.2025", "11.08.2026"),
        ("ktk_senaryoIyi", 0.90, "çarpan", "İyimser endeks çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ktk_senaryoKotu", 1.10, "çarpan", "Kötümser endeks çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ktk_oranArtisPuan", 0.01, "oran", "M-SEN sabit +1 puan", "Senaryo motoru", "01.01.2025", "11.08.2026"),
        ("ktk_tahminAlt", 0.85, "çarpan", "Tahmin alt bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ktk_tahminUst", 1.15, "çarpan", "Tahmin üst bant", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ktk_tornadoMatrah", 1.05, "çarpan", "Tornado endeks şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ktk_tornadoMevcut", 0.95, "çarpan", "Tornado hakediş şoku", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ktk_isverenOran", 0.10, "oran", "Yorum B yedek sabit artışı", "Fiyat farkı model", "01.01.2025", "11.08.2026"),
        ("ktk_yuzdelikOran", 0.90, "oran", "Yüzdelik P90", "İstatistik", "01.01.2025", "11.08.2026"),
        ("ktk_parmakIzi", "KTK-PRO-1.0.0", "metin", "Dosya parmak izi etiketi", "Üretim", "01.01.2025", "11.08.2026"),
        ("dosya_surumu", SURUM, "metin", "Ürün sürümü", "ExcelArşiv", "01.01.2025", "11.08.2026"),
        ("ktk_girisBeklenen", 11, "adet", "Zorunlu giriş alanı sayısı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ktk_riskDusuk", 15, "puan", "TALEP ET risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("ktk_riskOrta", 55, "puan", "BEKLE risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("ktk_riskYuksek", 85, "puan", "VAZGEÇ risk skoru", "İç politika", "01.01.2025", "11.08.2026"),
        ("ktk_riskEsikOran", 0.5, "oran", "Yedek risk oranı", "İç politika", "01.01.2025", "11.08.2026"),
        ("ktk_yilTaban", 2024, "yıl", "CHOOSE yıl tabanı (hesapYili−taban)", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ktk_yuvarMax", 10, "adet", "ROUND basamak tavanı", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ktk_dayanakCarpan", 10, "adet", "Yedek çarpan", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ktk_bazCarpan", 1, "çarpan", "Baz senaryo çarpanı", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ktk_sifirCarpan", 0, "çarpan", "Nötr çarpan (sıfır)", "Model varsayımı", "01.01.2025", "11.08.2026"),
        ("ktk_siraBir", 1, "adet", "Tornado sıra 1", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ktk_siraIki", 2, "adet", "Tornado sıra 2", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ktk_siraUc", 3, "adet", "Tornado sıra 3", "Ürün mimarisi", "01.01.2025", "11.08.2026"),
        ("ktk_tahminX", 5, "adet", "FORECAST X noktası", "Model varsayımı", "01.01.2025", "11.08.2026"),
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
    ws.conditional_formatting.add("F7:F53", FormulaRule(formula=['F7="KTK-001"'], font=Font(color="B08948", bold=True)))
    ws.conditional_formatting.add("F7:F53", FormulaRule(formula=['F7="KTK-005"'], font=Font(color=KRITIK, bold=True)))

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
    ad_ekle(wb, "ktk_ozKaynak", "MOTOR!$G$28")
    ad_ekle(wb, "ktk_eksikSermaye", "MOTOR!$G$30")
    ad_ekle(wb, "ktk_kritikAcik", "MOTOR!$G$30")
    ad_ekle(wb, "ktk_kacirilan", "MOTOR!$G$30")
    ad_ekle(wb, "ktk_kararMetni", "MOTOR!$G$34")
    ad_ekle(wb, "ktk_kararGerekce", "MOTOR!$G$35")
    ad_ekle(wb, "ktk_maddeAtifMetni", "MOTOR!$G$52")
    ad_ekle(wb, "ktk_kanitRaporu", "MOTOR!$G$53")
    ad_ekle(wb, "ktk_kuralYilMatris", "KURALLAR!$B$17")
    ad_ekle(wb, "ktk_parmakIzi", "AYARLAR!$B$19")

    ad_ekle(wb, "ktk_senaryoKarsilastirma", "SENARYO!$B$11")
    ad_ekle(wb, "ktk_yorumSenaryo", "SENARYO!$B$12")
    ad_ekle(wb, "ktk_duyarlilikSira", "SENARYO!$B$19")
    ad_ekle(wb, "ktk_yorumDuyarlilik", "SENARYO!$B$20")
    ad_ekle(wb, "ktk_kuralDegisimSenaryo", "SENARYO!$B$22")
    ad_ekle(wb, "ktk_yorumKuralDegisim", "SENARYO!$B$23")
    ad_ekle(wb, "ktk_tahminAralik", "SENARYO!$B$32")
    ad_ekle(wb, "ktk_yorumTahmin", "SENARYO!$B$33")

    ad_ekle(wb, "ktk_vakaDurum", "VAKALAR!$B$12")
    ad_ekle(wb, "ktk_vakaFark", "VAKALAR!$B$13")

    ayar_map = {
        "ktk_tolerans": 8,
        "ktk_kacirilanEsik": 9,
        "ktk_senaryoIyi": 10,
        "ktk_senaryoKotu": 11,
        "ktk_oranArtisPuan": 12,
        "ktk_tahminAlt": 13,
        "ktk_tahminUst": 14,
        "ktk_tornadoMatrah": 15,
        "ktk_tornadoMevcut": 16,
        "ktk_isverenOran": 17,
        "ktk_yuzdelikOran": 18,
        "ktk_girisBeklenen": 21,
        "ktk_riskDusuk": 22,
        "ktk_riskOrta": 23,
        "ktk_riskYuksek": 24,
        "ktk_riskEsikOran": 25,
        "ktk_yilTaban": 26,
        "ktk_yuvarMax": 27,
        "ktk_dayanakCarpan": 28,
        "ktk_bazCarpan": 29,
        "ktk_sifirCarpan": 30,
        "ktk_siraBir": 31,
        "ktk_siraIki": 32,
        "ktk_siraUc": 33,
        "ktk_tahminX": 34,
        "vaka_matrah_1": 35,
        "vaka_matrah_2": 36,
        "vaka_tavan": 37,
        "vaka_endeks_orani": 38,
    }
    for ad, satir in ayar_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")

    modul_map = {
        "ktk_modulT1": (11, "C"), "ktk_modulT1Yorum": (11, "D"),
        "ktk_modulT2": (12, "C"), "ktk_modulT2Yorum": (12, "D"),
        "ktk_modulO1": (13, "C"), "ktk_modulO1Yorum": (13, "D"),
        "ktk_modulO6": (16, "C"), "ktk_modulO6Yorum": (16, "D"),
        "ktk_modulO8": (17, "C"), "ktk_modulO8Yorum": (17, "D"),
        "ktk_modulI2": (20, "C"), "ktk_modulI2Yorum": (20, "D"),
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
    varsayilan = os.path.join(kok, "KonkordatoTtk376KrizPaketi.xlsx")
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
    cikti = os.path.join(repo, "cikti", "KonkordatoTtk376KrizPaketi.xlsx")
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(uretilen) != os.path.abspath(cikti):
        shutil.copy2(uretilen, cikti)
        print(f"Kopya: {cikti}")
